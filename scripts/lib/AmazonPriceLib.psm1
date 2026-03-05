Set-StrictMode -Version Latest

function ConvertTo-PlainText {
    param([SecureString]$Secure)
    $ptr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($Secure)
    try {
        [Runtime.InteropServices.Marshal]::PtrToStringBSTR($ptr)
    }
    finally {
        [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($ptr)
    }
}

function Write-Log {
    param(
        [string]$Message,
        [string]$LogPath,
        [string]$Level = 'INFO'
    )

    $line = "$(Get-Date -Format o) [$Level] $Message"
    Add-Content -Path $LogPath -Value $line
    # suppress WARN messages on console to reduce noise
    if ($Level -ne 'WARN') {
        Write-Host $line
    }
}


function Initialize-RunStats {
    $script:RunStats = [ordered]@{
        TotalApiCalls      = 0
        PricingCalls       = 0
        CatalogBatchCalls  = 0
        RetryCount         = 0
        Http429Count       = 0
        TotalWaitSec       = 0.0
        WaitEvents         = 0
        NextPricingAllowedAt = Get-Date
        PricingCooldownSec = 0.0
        ResponseHeaderWarningShown = $false
    }
}

function Add-WaitMetric {
    param([double]$Seconds)
    if ($Seconds -le 0) { return }
    $script:RunStats.TotalWaitSec += $Seconds
    $script:RunStats.WaitEvents += 1
}

function Get-HeaderValue {
    param([object]$Headers, [string]$Name)
    if (-not $Headers) { return $null }
    $v = $Headers[$Name]
    if (-not $v) { $v = $Headers[$Name.ToLowerInvariant()] }
    if (-not $v) { $v = $Headers[$Name.ToUpperInvariant()] }
    if ($v -is [array]) { return [string]$v[0] }
    return [string]$v
}

function Get-PropertyValue {
    param(
        [object]$Object,
        [string]$Name
    )

    if ($null -eq $Object -or [string]::IsNullOrWhiteSpace($Name)) {
        return $null
    }

    $property = $Object.PSObject.Properties[$Name]
    if ($null -eq $property) {
        return $null
    }

    return $property.Value
}

function Mask-SensitiveText {
    param([string]$Text)

    if ([string]::IsNullOrEmpty($Text)) {
        return $Text
    }

    $masked = $Text
    $patterns = @(
        '(?i)(x-amz-access-token\s*[=:]\s*)([^\s"'',;]+)',
        '(?i)(Authorization\s*[=:]\s*Bearer\s+)([^\s"'',;]+)',
        '(?i)("x-amz-access-token"\s*:\s*")(.*?)(")',
        '(?i)("Authorization"\s*:\s*")(.*?)(")'
    )

    foreach ($pattern in $patterns) {
        $masked = [regex]::Replace($masked, $pattern, {
            param($m)
            if ($m.Groups.Count -ge 4) {
                return "$($m.Groups[1].Value)***MASKED***$($m.Groups[3].Value)"
            }

            return "$($m.Groups[1].Value)***MASKED***"
        })
    }

    return $masked
}

function Write-SpApiResponseDebugLog {
    param(
        [string]$Endpoint,
        [object]$Response,
        [hashtable]$Config,
        [string]$LogPath
    )

    if (-not $Config.DebugSpApiResponse) {
        return
    }

    $maxChars = if ($Config.DebugSpApiResponseMaxChars) { [int]$Config.DebugSpApiResponseMaxChars } else { 4000 }
    $maxChars = [Math]::Max(200, $maxChars)

    $shouldLogFull = $false

    $responses = Get-PropertyValue -Object $Response -Name 'responses'
    if ($responses) {
        $responseItems = @($responses)
        $responseCount = $responseItems.Count
        $index = 0
        foreach ($item in $responseItems) {
            $index++
            $status = Get-PropertyValue -Object $item -Name 'status'
            $request = Get-PropertyValue -Object $item -Name 'request'
            $requestUri = Get-PropertyValue -Object $request -Name 'uri'
            $body = Get-PropertyValue -Object $item -Name 'body'
            $errors = Get-PropertyValue -Object $body -Name 'errors'
            $payload = Get-PropertyValue -Object $body -Name 'payload'
            $asin = Get-PropertyValue -Object $payload -Name 'ASIN'
            $offers = Get-PropertyValue -Object $payload -Name 'Offers'
            $offersCount = if ($offers) { @($offers).Count } else { 0 }
            $errorCount = if ($errors) { @($errors).Count } else { 0 }

            Write-Log -Message "$Endpoint debug[$index/$responseCount]: status=$status request.uri=$requestUri payload.ASIN=$asin offers.count=$offersCount errors.count=$errorCount" -LogPath $LogPath
            if (($status -as [int]) -ge 400 -or $errorCount -gt 0) { $shouldLogFull = $true }
        }
    }
    else {
        $status = Get-PropertyValue -Object $Response -Name 'status'
        $payload = Get-PropertyValue -Object $Response -Name 'payload'
        $errors = Get-PropertyValue -Object $Response -Name 'errors'
        $requestUri = Get-PropertyValue -Object $Response -Name 'uri'
        $asin = Get-PropertyValue -Object $payload -Name 'ASIN'
        $offers = Get-PropertyValue -Object $payload -Name 'Offers'
        $offersCount = if ($offers) { @($offers).Count } else { 0 }
        $errorCount = if ($errors) { @($errors).Count } else { 0 }

        Write-Log -Message "$Endpoint debug: status=$status request.uri=$requestUri payload.ASIN=$asin offers.count=$offersCount errors.count=$errorCount" -LogPath $LogPath
        if (($status -as [int]) -ge 400 -or $errorCount -gt 0) { $shouldLogFull = $true }
    }

    if (-not $shouldLogFull) {
        return
    }

    $responseJson = Mask-SensitiveText -Text (($Response | ConvertTo-Json -Depth 20 -Compress) 2>$null)
    if (-not [string]::IsNullOrWhiteSpace($responseJson)) {
        $snippet = if ($responseJson.Length -gt $maxChars) {
            "$($responseJson.Substring(0, $maxChars))...(truncated)"
        }
        else {
            $responseJson
        }
        Write-Log -Message "$Endpoint debug.full(max=$maxChars): $snippet" -LogPath $LogPath
    }
}

function Wait-ForPricingSlot {
    param([hashtable]$Config)
    if (-not $script:RunStats) { return }

    $now = Get-Date
    $baseInterval = if ($Config.PricingMinIntervalSec) { [double]$Config.PricingMinIntervalSec } else { 2.2 }
    $target = if ($script:RunStats.NextPricingAllowedAt -gt $now) { $script:RunStats.NextPricingAllowedAt } else { $now }
    $waitSec = ($target - $now).TotalSeconds
    if ($waitSec -gt 0) {
        Start-Sleep -Milliseconds ([int]([Math]::Ceiling($waitSec * 1000)))
        Add-WaitMetric -Seconds $waitSec
    }

    $script:RunStats.NextPricingAllowedAt = (Get-Date).AddSeconds($baseInterval + [double]$script:RunStats.PricingCooldownSec)
}

function Update-PricingThrottleFromLimit {
    param(
        [string]$RateLimitLimit,
        [hashtable]$Config,
        [switch]$Had429
    )

    if (-not $script:RunStats) { return }

    if ($Had429) {
        $script:RunStats.PricingCooldownSec = [Math]::Min(15.0, [double]$script:RunStats.PricingCooldownSec + 1.0)
    }
    elseif ($script:RunStats.PricingCooldownSec -gt 0) {
        $script:RunStats.PricingCooldownSec = [Math]::Max(0.0, [double]$script:RunStats.PricingCooldownSec - 0.2)
    }

    $baseInterval = if ($Config.PricingMinIntervalSec) { [double]$Config.PricingMinIntervalSec } else { 2.2 }
    if ($RateLimitLimit) {
        $limit = $RateLimitLimit -as [double]
        if ($limit -and $limit -gt 0) {
            if ($limit -le 0.5) { $baseInterval = Get-Random -Minimum 2.2 -Maximum 2.5 }
            elseif ($limit -le 1.0) { $baseInterval = Get-Random -Minimum 1.1 -Maximum 1.3 }
            else { $baseInterval = [Math]::Max(0.4, (1.0 / $limit) * 1.2) }
        }
    }

    $script:RunStats.NextPricingAllowedAt = (Get-Date).AddSeconds($baseInterval + [double]$script:RunStats.PricingCooldownSec)
}

function Invoke-SpApiRequest {
    param(
        [string]$Endpoint,
        [string]$Method,
        [string]$Uri,
        [hashtable]$Headers,
        [object]$Body,
        [hashtable]$Config,
        [string]$LogPath
    )

    $maxAttempts = if ($Config.RetryMaxAttempts) { [int]$Config.RetryMaxAttempts } else { 6 }
    $maxWaitSec = if ($Config.RetryMaxWaitSec) { [int]$Config.RetryMaxWaitSec } else { 120 }

    for ($attempt = 1; $attempt -le $maxAttempts; $attempt++) {
        $responseHeaders = $null
        try {
            if ($script:RunStats) { $script:RunStats.TotalApiCalls++ }
            if ($Endpoint -match '^Pricing') { $script:RunStats.PricingCalls++ }
            if ($Endpoint -match '^CatalogBatch') { $script:RunStats.CatalogBatchCalls++ }

            $params = @{ Method = $Method; Uri = $Uri; Headers = $Headers }
            if ($null -ne $Body -and "$Body" -ne '') { $params.Body = $Body }

            $irmCommand = Get-Command -Name 'Invoke-RestMethod' -ErrorAction Stop
            if ($irmCommand.Parameters.ContainsKey('ResponseHeadersVariable')) {
                $res = Invoke-RestMethod @params -ResponseHeadersVariable responseHeaders
            }
            else {
                $res = Invoke-RestMethod @params
                $responseHeaders = $null
                if (-not $script:RunStats -or -not $script:RunStats.ResponseHeaderWarningShown) {
                    Write-Log -Message "response headers unavailable on this PowerShell runtime (ResponseHeadersVariable unsupported)" -LogPath $LogPath -Level 'WARN'
                    if ($script:RunStats) { $script:RunStats.ResponseHeaderWarningShown = $true }
                }
            }

            $limit = Get-HeaderValue -Headers $responseHeaders -Name 'x-amzn-RateLimit-Limit'
            $requestId = Get-HeaderValue -Headers $responseHeaders -Name 'x-amzn-RequestId'
            if ($limit) {
                Write-Log -Message "$Endpoint success: limit=$limit, requestId=$requestId" -LogPath $LogPath
            }
            if ($Endpoint -match '^Pricing') {
                Update-PricingThrottleFromLimit -RateLimitLimit $limit -Config $Config
            }
            Write-SpApiResponseDebugLog -Endpoint $Endpoint -Response $res -Config $Config -LogPath $LogPath
            return $res
        }
        catch {
            $detail = Get-ErrorDetail -ErrorRecord $_
            $status = if ($detail.StatusCode) { [int]$detail.StatusCode } else { 0 }
            if ($status -eq 0) {
                # log the raw exception for troubleshooting
                Write-Log -Message "Invoke-SpApiRequest exception: $($_.Exception.GetType().FullName) - $($_.Exception.Message)" -LogPath $LogPath -Level 'WARN'
            }
            if ($status -eq 403) {
                Write-Log -Message "HTTP 403 exception: $($_.Exception.GetType().FullName) - $($_.Exception.Message)" -LogPath $LogPath -Level 'WARN'
            }
            if ($status -eq 429 -and $script:RunStats) { $script:RunStats.Http429Count++ }

            $errorCode = ''
            $requestId = ''
            if ($detail.BodyText) {
                if ($detail.BodyText -match '"code"\s*:\s*"([^"]+)"') { $errorCode = $matches[1] }
                if ($detail.BodyText -match '"requestId"\s*:\s*"([^"]+)"') { $requestId = $matches[1] }
            }

            $retryable = ($status -eq 429 -or $status -eq 500 -or $status -eq 503 -or $detail.Class -eq 'RateLimit/Server' -or $detail.Class -eq 'Auth')
            if (-not $retryable -or $attempt -ge $maxAttempts) {
                $bodyMsg = if ($detail.BodyText) { " body='$($detail.BodyText)'" } else { '' }
                Write-Log -Message "$Endpoint failed: status=$status class=$($detail.Class) code=$errorCode requestId=$requestId attempt=$attempt/$maxAttempts$bodyMsg" -LogPath $LogPath -Level 'WARN'
                throw
            }

            $sleepSec = 0.0
            if ($detail.RetryAfterSec -and $detail.RetryAfterSec -gt 0) {
                $sleepSec = [double]$detail.RetryAfterSec
            }
            else {
                $base = [Math]::Pow(2, $attempt)
                $jitter = (Get-Random -Minimum 0 -Maximum 1000) / 1000.0
                $sleepSec = [Math]::Min($maxWaitSec, $base + $jitter)
            }

            if ($script:RunStats) { $script:RunStats.RetryCount++ }
            Add-WaitMetric -Seconds $sleepSec
            if ($Endpoint -match '^Pricing' -and $status -eq 429) {
                Update-PricingThrottleFromLimit -RateLimitLimit $detail.RateLimitLimit -Config $Config -Had429
            }

            $bodyMsg = if ($detail.BodyText) { " body='$($detail.BodyText)'" } else { '' }
            Write-Log -Message "$Endpoint retry: status=$status class=$($detail.Class) code=$errorCode requestId=$requestId wait=$([Math]::Round($sleepSec,2))s limit=$($detail.RateLimitLimit) attempt=$attempt/$maxAttempts$bodyMsg" -LogPath $LogPath -Level 'WARN'
            Start-Sleep -Milliseconds ([int]([Math]::Ceiling($sleepSec * 1000)))
        }
    }
}

function Get-StatusClassification {
    param(
        [Nullable[int]]$StatusCode,
        [string]$BodyText
    )

    $text = if ($BodyText) { $BodyText.ToLowerInvariant() } else { '' }

    if ($StatusCode -eq 429 -or ($StatusCode -ge 500 -and $StatusCode -lt 600)) {
        return [PSCustomObject]@{ Class = 'RateLimit/Server'; IsTransient = $true; IsPermanentNotFound = $false }
    }

    if ($StatusCode -eq 401 -or $StatusCode -eq 403) {
        return [PSCustomObject]@{ Class = 'Auth'; IsTransient = $false; IsPermanentNotFound = $false }
    }

    if ($StatusCode -eq 404 -or $StatusCode -eq 400 -or $StatusCode -eq 422) {
        return [PSCustomObject]@{ Class = 'NotFound/Validation'; IsTransient = $false; IsPermanentNotFound = $true }
    }

    if ($text -match 'not\s*found|invalid|validation|no\s*matching|notfound') {
        return [PSCustomObject]@{ Class = 'NotFound/Validation'; IsTransient = $false; IsPermanentNotFound = $true }
    }

    if ($text -match 'throttl|rate\s*exceed|too\s*many\s*requests|temporar|timeout|service\s*unavailable') {
        return [PSCustomObject]@{ Class = 'RateLimit/Server'; IsTransient = $true; IsPermanentNotFound = $false }
    }

    return [PSCustomObject]@{ Class = 'Other'; IsTransient = $false; IsPermanentNotFound = $false }
}

function Get-ErrorDetail {
    param([object]$ErrorRecord)

    $statusCode = $null
    $bodyText = $null
    $retryAfterSec = $null
    $rateLimitLimit = $null

    if ($ErrorRecord -and $ErrorRecord.Exception) {
        try {
            $hasResponse = $ErrorRecord.Exception | Get-Member -Name 'Response' -MemberType 'Property' -ErrorAction SilentlyContinue
            if ($hasResponse -and $ErrorRecord.Exception.Response) {
                if ($ErrorRecord.Exception.Response.StatusCode) {
                    $statusCode = [int]$ErrorRecord.Exception.Response.StatusCode
                }
            }
        }
        catch {}

        try {
            $hasResponse = $ErrorRecord.Exception | Get-Member -Name 'Response' -MemberType 'Property' -ErrorAction SilentlyContinue
            if ($hasResponse) {
                $headers = $ErrorRecord.Exception.Response.Headers
                if ($headers) {
                    $retryAfterRaw = $headers['Retry-After']
                    if (-not $retryAfterRaw) { $retryAfterRaw = $headers['retry-after'] }
                    if ($retryAfterRaw) {
                        $retryAfterInt = $retryAfterRaw -as [int]
                        if ($retryAfterInt -and $retryAfterInt -gt 0) {
                            $retryAfterSec = [Math]::Max(1, $retryAfterInt)
                        }
                        else {
                            try {
                                $retryAfterDate = [DateTime]::ParseExact($retryAfterRaw, 'r', [System.Globalization.CultureInfo]::InvariantCulture)
                                $delta = [int][Math]::Ceiling(($retryAfterDate - (Get-Date)).TotalSeconds)
                                if ($delta -gt 0) { $retryAfterSec = $delta }
                            }
                            catch { }
                        }
                    }

                    $rateLimitLimit = $headers['x-amzn-RateLimit-Limit']
                    if (-not $rateLimitLimit) { $rateLimitLimit = $headers['X-Amzn-RateLimit-Limit'] }
                }
            }
        }
        catch {}

        try {
            $hasResponse = $ErrorRecord.Exception | Get-Member -Name 'Response' -MemberType 'Property' -ErrorAction SilentlyContinue
            if ($hasResponse) {
                $stream = $ErrorRecord.Exception.Response.GetResponseStream()
                if ($stream) {
                    $reader = New-Object System.IO.StreamReader($stream)
                    $bodyText = $reader.ReadToEnd()
                    $reader.Close()
                }
            }
        }
        catch {}
    }

    $classification = Get-StatusClassification -StatusCode $statusCode -BodyText $bodyText
    return [PSCustomObject]@{
        StatusCode          = $statusCode
        BodyText            = $bodyText
        RetryAfterSec       = $retryAfterSec
        RateLimitLimit      = $rateLimitLimit
        Class               = $classification.Class
        IsTransient         = $classification.IsTransient
        IsPermanentNotFound = $classification.IsPermanentNotFound
    }
}

function Invoke-WithRetry {
    param(
        [scriptblock]$Action,
        [string]$Label,
        [int]$MaxRetries,
        [string]$LogPath
    )

    $attempt = 0
    while ($true) {
        $attempt++
        try {
            return & $Action
        }
        catch {
            $detail = Get-ErrorDetail -ErrorRecord $_
            $rateLimitInfo = if ($detail.RateLimitLimit) { ", limit=$($detail.RateLimitLimit)" } else { '' }

            if ($detail.IsTransient -and $attempt -lt $MaxRetries) {
                if ($detail.RetryAfterSec -and $detail.RetryAfterSec -gt 0) {
                    $sleepSec = [int]$detail.RetryAfterSec
                    Write-Log -Message "$Label 失敗 (分類=$($detail.Class), HTTP $($detail.StatusCode)$rateLimitInfo)。Retry-After=$sleepSec 秒を優先してリトライします (試行 $attempt/$MaxRetries)。" -LogPath $LogPath -Level 'WARN'
                }
                else {
                    $baseSec = [Math]::Pow(2, $attempt)
                    $jitterMs = Get-Random -Minimum 0 -Maximum 1000
                    $sleepSec = [double]$baseSec + ([double]$jitterMs / 1000.0)
                    $sleepSecText = [Math]::Round($sleepSec, 2)
                    Write-Log -Message "$Label 失敗 (分類=$($detail.Class), HTTP $($detail.StatusCode)$rateLimitInfo)。指数バックオフ+$jitterMs ms ジッターで $sleepSecText 秒後にリトライします (試行 $attempt/$MaxRetries)。" -LogPath $LogPath -Level 'WARN'
                }

                Start-Sleep -Milliseconds ([int]([Math]::Ceiling($sleepSec * 1000)))
                continue
            }

            Write-Log -Message "$Label 失敗 (分類=$($detail.Class), HTTP $($detail.StatusCode)$rateLimitInfo)。再試行を終了します。" -LogPath $LogPath -Level 'WARN'
            throw
        }
    }
}


function Get-AmzDateHeaderValue {
    return (Get-Date).ToUniversalTime().ToString('yyyyMMddTHHmmssZ')
}

function New-SpApiHeaders {
    param(
        [string]$AccessToken,
        [hashtable]$Config,
        [string]$ContentType
    )

    $headers = @{
        'x-amz-access-token' = $AccessToken
        'x-amz-date'         = Get-AmzDateHeaderValue
        'User-Agent'         = $Config.UserAgent
        'Accept'             = 'application/json'
    }

    if ($ContentType) {
        $headers['Content-Type'] = $ContentType
    }

    return $headers
}

function Split-IntoChunks {
    param(
        [array]$Items,
        [int]$ChunkSize
    )

    $chunks = @()
    if (-not $Items -or $Items.Count -eq 0) {
        return $chunks
    }

    for ($i = 0; $i -lt $Items.Count; $i += $ChunkSize) {
        $end = [Math]::Min($i + $ChunkSize - 1, $Items.Count - 1)
        $chunk = @($Items[$i..$end])
        $chunks += ,$chunk
    }

    return $chunks
}


function Read-AccessTokenCache {
    param([string]$Path)

    if (-not (Test-Path $Path)) { return $null }

    try {
        $raw = Get-Content -Path $Path -Raw -Encoding UTF8
        if ([string]::IsNullOrWhiteSpace($raw)) { return $null }
        $parsed = ConvertFrom-Json -InputObject $raw
        if (-not $parsed -or -not $parsed.token -or -not $parsed.expires_at) { return $null }

        try {
            $expiresAt = [DateTime]::Parse($parsed.expires_at)
        }
        catch {
            return $null
        }

        [PSCustomObject]@{ token = [string]$parsed.token; expires_at = $expiresAt }
    }
    catch {
        return $null
    }
}

function Save-AccessTokenCache {
    param(
        [string]$Path,
        [string]$Token,
        [int]$ExpiresInSeconds
    )

    $parentDir = Split-Path -Path $Path -Parent
    if ($parentDir -and -not (Test-Path $parentDir)) {
        New-Item -ItemType Directory -Path $parentDir -Force | Out-Null
    }

    $safeExpires = if ($ExpiresInSeconds -and $ExpiresInSeconds -gt 90) { $ExpiresInSeconds - 60 } else { 3300 }
    $expiresAt = (Get-Date).AddSeconds($safeExpires).ToString('o')

    [PSCustomObject]@{
        token      = $Token
        expires_at = $expiresAt
    } | ConvertTo-Json -Depth 3 | Set-Content -Path $Path -Encoding UTF8
}

function Get-LwaAccessTokenCached {
    param(
        [string]$ClientId,
        [string]$ClientSecret,
        [string]$RefreshToken,
        [hashtable]$Config,
        [string]$LogPath,
        [string]$TokenCachePath,
        [switch]$ForceRefresh
    )

    if (-not $ForceRefresh) {
        $cached = Read-AccessTokenCache -Path $TokenCachePath
        if ($cached -and $cached.expires_at -gt (Get-Date).AddSeconds(30)) {
            Write-Log -Message 'LWAアクセストークンをキャッシュから再利用します。' -LogPath $LogPath
            return $cached.token
        }
    }

    $body = @{
        grant_type    = 'refresh_token'
        refresh_token = $RefreshToken
        client_id     = $ClientId
        client_secret = $ClientSecret
    }

    $res = Invoke-WithRetry -Label 'LWAトークン取得' -MaxRetries $Config.MaxRetries -LogPath $LogPath -Action {
        Invoke-RestMethod -Method Post -Uri 'https://api.amazon.com/auth/o2/token' -ContentType 'application/x-www-form-urlencoded' -Body $body -Headers @{
            'User-Agent' = $Config.UserAgent
        }
    }

    if (-not $res.access_token) {
        throw 'LWAアクセストークンの取得に失敗しました。'
    }

    $expiresIn = 3600
    if ($res.expires_in) { $expiresIn = [int]$res.expires_in }
    Save-AccessTokenCache -Path $TokenCachePath -Token ([string]$res.access_token) -ExpiresInSeconds $expiresIn
    Write-Log -Message 'LWAアクセストークンを新規取得してキャッシュへ保存しました。' -LogPath $LogPath
    return [string]$res.access_token
}

function Get-LwaAccessToken {
    param(
        [string]$ClientId,
        [string]$ClientSecret,
        [string]$RefreshToken,
        [hashtable]$Config,
        [string]$LogPath,
        [string]$TokenCachePath
    )

    return Get-LwaAccessTokenCached -ClientId $ClientId -ClientSecret $ClientSecret -RefreshToken $RefreshToken -Config $Config -LogPath $LogPath -TokenCachePath $TokenCachePath
}

function Get-LowestNewPriceFromOffers {
    param([array]$Offers)

    if (-not $Offers) {
        return $null
    }

    $landedMin = $null
    $fallbackMin = $null

    foreach ($offer in $Offers) {
        if ($offer.LandedPrice -and $null -ne $offer.LandedPrice.Amount) {
            $landed = [decimal]$offer.LandedPrice.Amount
            if ($null -eq $landedMin -or $landed -lt $landedMin) {
                $landedMin = $landed
            }
            continue
        }

        if ($offer.ListingPrice -and $offer.Shipping -and $null -ne $offer.ListingPrice.Amount -and $null -ne $offer.Shipping.Amount) {
            $listingPlusShip = [decimal]$offer.ListingPrice.Amount + [decimal]$offer.Shipping.Amount
            if ($null -eq $fallbackMin -or $listingPlusShip -lt $fallbackMin) {
                $fallbackMin = $listingPlusShip
            }
        }
    }

    if ($null -ne $landedMin) {
        return $landedMin
    }

    return $fallbackMin
}

function Get-AsinMapByJanBatch {
    param(
        [array]$Jans,
        [string]$AccessToken,
        [hashtable]$Config,
        [string]$LogPath,
        [hashtable]$AuthContext
    )

    $resultMap = @{}
    $errorClassMap = @{}
    $minBatchSize = 5
    $batchSize = [Math]::Max($minBatchSize, [int]$Config.CatalogBatchSize)
    $index = 0

    $applyCatalogItems = {
        param(
            [object]$Items,
            [hashtable]$TargetMap,
            [hashtable]$TargetErrorClassMap
        )

        if (-not $Items) { return }

        foreach ($item in $Items) {
            $itemIdentifiers = Get-PropertyValue -Object $item -Name 'identifiers'
            if (-not $itemIdentifiers) { continue }

            $identifierGroups = @()
            # Catalog Items API の identifiers は配列で返る実装があるため、
            # 直接配列として扱える場合はそのまま利用する。
            if ($itemIdentifiers -is [System.Collections.IEnumerable] -and -not ($itemIdentifiers -is [string])) {
                $identifierGroups = @($itemIdentifiers)
            }
            else {
                $nestedGroups = Get-PropertyValue -Object $itemIdentifiers -Name 'identifiers'
                if ($nestedGroups) {
                    $identifierGroups = @($nestedGroups)
                }
            }
            if ($identifierGroups.Count -eq 0) { continue }

            $matchedIdentifier = $null
            foreach ($idGroup in $identifierGroups) {
                # identifiers 配下に identifiers 配列がネストされるケースと、
                # 直接 identifierType/identifier を持つケースの両方に対応する。
                $leafIdentifiers = @()
                $nestedLeafIdentifiers = Get-PropertyValue -Object $idGroup -Name 'identifiers'
                if ($nestedLeafIdentifiers) {
                    $leafIdentifiers = @($nestedLeafIdentifiers)
                }
                else {
                    $leafIdentifiers = @($idGroup)
                }

                foreach ($leaf in $leafIdentifiers) {
                    $identifierType = [string](Get-PropertyValue -Object $leaf -Name 'identifierType')
                    $identifierValue = [string](Get-PropertyValue -Object $leaf -Name 'identifier')
                    if (($identifierType -in @('JAN', 'EAN')) -and -not [string]::IsNullOrWhiteSpace($identifierValue)) {
                        $normalizedIdentifier = $identifierValue.Trim()
                        if ($TargetMap.ContainsKey($normalizedIdentifier)) {
                            $matchedIdentifier = $normalizedIdentifier
                            break
                        }
                    }
                }
                if ($matchedIdentifier) { break }
            }

            $asin = [string](Get-PropertyValue -Object $item -Name 'asin')
            if ($matchedIdentifier -and -not [string]::IsNullOrWhiteSpace($asin)) {
                $TargetMap[$matchedIdentifier] = $asin
                $TargetErrorClassMap.Remove($matchedIdentifier) | Out-Null
            }
        }

        return
    }

    $resolveCatalogItems = {
        param([object]$Response)

        $items = Get-PropertyValue -Object $Response -Name 'items'
        if ($items) { return @($items) }

        $payload = Get-PropertyValue -Object $Response -Name 'payload'
        if ($payload) {
            $payloadItems = Get-PropertyValue -Object $payload -Name 'items'
            if ($payloadItems) { return @($payloadItems) }
        }

        return @()
    }

    $resolveAsinBySingleIdentifier = {
        param(
            [string]$Identifier,
            [string]$IdentifiersType
        )

        if ([string]::IsNullOrWhiteSpace($Identifier)) { return $null }

        $singleUri = "$($Config.SpApiBaseUrl)/catalog/2022-04-01/items?identifiers=$([Uri]::EscapeDataString($Identifier.Trim()))&identifiersType=$IdentifiersType&includedData=identifiers&marketplaceIds=$($Config.MarketplaceId)"
        try {
            $singleRes = Invoke-SpApiRequest -Endpoint "CatalogSingleFallback(type=$IdentifiersType)" -Method 'Get' -Uri $singleUri -Headers (New-SpApiHeaders -AccessToken $AccessToken -Config $Config) -Config $Config -LogPath $LogPath
            $singleItems = & $resolveCatalogItems -Response $singleRes
            if ($singleItems.Count -gt 0) {
                $firstAsin = [string](Get-PropertyValue -Object $singleItems[0] -Name 'asin')
                if (-not [string]::IsNullOrWhiteSpace($firstAsin)) {
                    return $firstAsin
                }
            }
        }
        catch {
            # 呼び出し元で最終分類するためここでは握りつぶし
        }

        return $null
    }

    while ($index -lt $Jans.Count) {
        $end = [Math]::Min($index + $batchSize - 1, $Jans.Count - 1)
        $chunk = @($Jans[$index..$end])
        foreach ($jan in $chunk) { $resultMap[$jan] = $null }

        $identifiers = ($chunk | ForEach-Object { $_.Trim() }) -join ','
        $uri = "$($Config.SpApiBaseUrl)/catalog/2022-04-01/items?identifiers=$([Uri]::EscapeDataString($identifiers))&identifiersType=JAN&includedData=identifiers&marketplaceIds=$($Config.MarketplaceId)"
        Write-Log -Message "JAN検索: $($chunk.Count)件 (index=$index,size=$batchSize)" -LogPath $LogPath

        $res = $null
        $attemptDetail = $null
        try {
            $res = Invoke-SpApiRequest -Endpoint "CatalogBatch(index=$index,size=$batchSize)" -Method 'Get' -Uri $uri -Headers (New-SpApiHeaders -AccessToken $AccessToken -Config $Config) -Config $Config -LogPath $LogPath
        }
        catch {
            $attemptDetail = Get-ErrorDetail -ErrorRecord $_
            if ($attemptDetail.Class -eq 'Auth' -and $AuthContext) {
                try {
                    $AccessToken = Get-LwaAccessTokenCached -ClientId $AuthContext.ClientId -ClientSecret $AuthContext.ClientSecret -RefreshToken $AuthContext.RefreshToken -Config $Config -LogPath $LogPath -TokenCachePath $AuthContext.TokenCachePath -ForceRefresh
                    $res = Invoke-SpApiRequest -Endpoint "CatalogBatchAuthRetry(index=$index,size=$batchSize)" -Method 'Get' -Uri $uri -Headers (New-SpApiHeaders -AccessToken $AccessToken -Config $Config) -Config $Config -LogPath $LogPath
                }
                catch {
                    $attemptDetail = Get-ErrorDetail -ErrorRecord $_
                }
            }
        }

        if (-not $res) {
            if ($attemptDetail -and $attemptDetail.Class -eq 'RateLimit/Server' -and $batchSize -gt $minBatchSize) {
                $nextBatchSize = [Math]::Max($minBatchSize, [int][Math]::Floor($batchSize / 2))
                Write-Log -Message "Catalogバッチを縮小します: $batchSize -> $nextBatchSize (index=$index, HTTP=$($attemptDetail.StatusCode), limit=$($attemptDetail.RateLimitLimit))" -LogPath $LogPath -Level 'WARN'
                $batchSize = $nextBatchSize
                continue
            }

            foreach ($jan in $chunk) {
                $errorClassMap[$jan] = if ($attemptDetail) { $attemptDetail.Class } else { 'Other' }
            }
            $index = $end + 1
            continue
        }

        $catalogItems = & $resolveCatalogItems -Response $res
        Write-Log -Message "Catalog応答items件数: $($catalogItems.Count) (index=$index)" -LogPath $LogPath
        & $applyCatalogItems -Items $catalogItems -TargetMap $resultMap -TargetErrorClassMap $errorClassMap | Out-Null

        $unresolvedJans = @($chunk | Where-Object { -not $resultMap[$_] })
        Write-Log -Message "EANフォールバック件数: $($unresolvedJans.Count)件 (index=$index)" -LogPath $LogPath
        if ($unresolvedJans.Count -gt 0) {
            $eanIdentifiers = ($unresolvedJans | ForEach-Object { $_.Trim() }) -join ','
            $eanUri = "$($Config.SpApiBaseUrl)/catalog/2022-04-01/items?identifiers=$([Uri]::EscapeDataString($eanIdentifiers))&identifiersType=EAN&includedData=identifiers&marketplaceIds=$($Config.MarketplaceId)"

            $eanRes = $null
            $eanAttemptDetail = $null
            try {
                $eanRes = Invoke-SpApiRequest -Endpoint "CatalogBatchEanFallback(index=$index,size=$($unresolvedJans.Count))" -Method 'Get' -Uri $eanUri -Headers (New-SpApiHeaders -AccessToken $AccessToken -Config $Config) -Config $Config -LogPath $LogPath
            }
            catch {
                $eanAttemptDetail = Get-ErrorDetail -ErrorRecord $_
                if ($eanAttemptDetail.Class -eq 'Auth' -and $AuthContext) {
                    try {
                        $AccessToken = Get-LwaAccessTokenCached -ClientId $AuthContext.ClientId -ClientSecret $AuthContext.ClientSecret -RefreshToken $AuthContext.RefreshToken -Config $Config -LogPath $LogPath -TokenCachePath $AuthContext.TokenCachePath -ForceRefresh
                        $eanRes = Invoke-SpApiRequest -Endpoint "CatalogBatchEanFallbackAuthRetry(index=$index,size=$($unresolvedJans.Count))" -Method 'Get' -Uri $eanUri -Headers (New-SpApiHeaders -AccessToken $AccessToken -Config $Config) -Config $Config -LogPath $LogPath
                    }
                    catch {
                        $eanAttemptDetail = Get-ErrorDetail -ErrorRecord $_
                    }
                }
            }

            if ($eanRes) {
                $eanItems = & $resolveCatalogItems -Response $eanRes
                Write-Log -Message "Catalog EAN応答items件数: $($eanItems.Count) (index=$index)" -LogPath $LogPath
                & $applyCatalogItems -Items $eanItems -TargetMap $resultMap -TargetErrorClassMap $errorClassMap | Out-Null
            }
            elseif ($eanAttemptDetail) {
                foreach ($jan in $unresolvedJans) {
                    $errorClassMap[$jan] = $eanAttemptDetail.Class
                }
            }
        }

        $finalUnresolvedJans = @($chunk | Where-Object { -not $resultMap[$_] })
        if ($finalUnresolvedJans.Count -gt 0) {
            Write-Log -Message "Catalog単発フォールバック件数: $($finalUnresolvedJans.Count)件 (index=$index)" -LogPath $LogPath -Level 'WARN'
            foreach ($jan in $finalUnresolvedJans) {
                $asinFromJan = & $resolveAsinBySingleIdentifier -Identifier $jan -IdentifiersType 'JAN'
                if ($asinFromJan) {
                    $resultMap[$jan] = $asinFromJan
                    $errorClassMap.Remove($jan) | Out-Null
                    continue
                }

                $asinFromEan = & $resolveAsinBySingleIdentifier -Identifier $jan -IdentifiersType 'EAN'
                if ($asinFromEan) {
                    $resultMap[$jan] = $asinFromEan
                    $errorClassMap.Remove($jan) | Out-Null
                    continue
                }
            }
        }

        foreach ($jan in $chunk) {
            if (-not $resultMap[$jan] -and -not $errorClassMap.ContainsKey($jan)) {
                $errorClassMap[$jan] = 'NotFound/Validation'
            }
        }

        $index = $end + 1
    }

    [PSCustomObject]@{ AsinMap = $resultMap; ErrorClassMap = $errorClassMap }
}


function Get-PriceBySingleAsin {
    param(
        [string]$Asin,
        [string]$AccessToken,
        [hashtable]$Config,
        [string]$LogPath,
        [hashtable]$AuthContext
    )

    $uri = "$($Config.SpApiBaseUrl)/products/pricing/v0/items/$([Uri]::EscapeDataString($Asin))/offers?MarketplaceId=$($Config.MarketplaceId)&ItemCondition=New"

    $res = $null
    try {
        Wait-ForPricingSlot -Config $Config
        $res = Invoke-SpApiRequest -Endpoint "PricingSingle(ASIN=$Asin)" -Method 'Get' -Uri $uri -Headers (New-SpApiHeaders -AccessToken $AccessToken -Config $Config) -Config $Config -LogPath $LogPath
    }
    catch {
        $detail = Get-ErrorDetail -ErrorRecord $_
        if ($detail.Class -eq 'Auth' -and $AuthContext) {
            $AccessToken = Get-LwaAccessTokenCached -ClientId $AuthContext.ClientId -ClientSecret $AuthContext.ClientSecret -RefreshToken $AuthContext.RefreshToken -Config $Config -LogPath $LogPath -TokenCachePath $AuthContext.TokenCachePath -ForceRefresh
            Wait-ForPricingSlot -Config $Config
            $res = Invoke-SpApiRequest -Endpoint "PricingSingleAuthRetry(ASIN=$Asin)" -Method 'Get' -Uri $uri -Headers (New-SpApiHeaders -AccessToken $AccessToken -Config $Config) -Config $Config -LogPath $LogPath
        }
        else {
            throw
        }
    }

    if (-not $res) {
        throw "ASIN=$Asin の単発価格取得結果が空です。"
    }

    $statusCode = if ($res.status) { [int]$res.status } else { $null }
    if ($statusCode -and $statusCode -ge 400) {
        $bodyText = $res | ConvertTo-Json -Depth 8
        $detail = Classify-StatusAndBody -StatusCode $statusCode -BodyText $bodyText
        return [PSCustomObject]@{ Price = $null; ErrorClass = $detail.Class }
    }

    $offers = if ($res.payload -and $res.payload.Offers) { $res.payload.Offers } else { $res.Offers }
    $price = Get-LowestNewPriceFromOffers -Offers $offers
    if ($null -eq $price) {
        return [PSCustomObject]@{ Price = $null; ErrorClass = 'NotFound/Validation' }
    }

    return [PSCustomObject]@{ Price = $price; ErrorClass = $null }
}

function Get-PriceMapByAsinSequential {
    param(
        [array]$Asins,
        [string]$AccessToken,
        [hashtable]$Config,
        [string]$LogPath,
        [hashtable]$AuthContext
    )

    $priceMap = @{}
    $errorClassMap = @{}

    Write-Log -Message "Pricingは単発APIで順次取得します。対象ASIN件数=$($Asins.Count)" -LogPath $LogPath
    foreach ($asin in $Asins) {
        $priceMap[$asin] = $null
        try {
            $single = Get-PriceBySingleAsin -Asin $asin -AccessToken $AccessToken -Config $Config -LogPath $LogPath -AuthContext $AuthContext
            $priceMap[$asin] = $single.Price
            if ($single.ErrorClass) { $errorClassMap[$asin] = $single.ErrorClass }
            else { $errorClassMap.Remove($asin) | Out-Null }
        }
        catch {
            $detail = Get-ErrorDetail -ErrorRecord $_
            $errorClassMap[$asin] = if ($detail.Class) { $detail.Class } else { 'Other' }
        }
    }

    [PSCustomObject]@{ PriceMap = $priceMap; ErrorClassMap = $errorClassMap }
}

# Backward-compatibility wrapper
function Get-PriceMapByAsinBatch {
    param(
        [array]$Asins,
        [string]$AccessToken,
        [hashtable]$Config,
        [string]$LogPath,
        [hashtable]$AuthContext
    )

    Get-PriceMapByAsinSequential -Asins $Asins -AccessToken $AccessToken -Config $Config -LogPath $LogPath -AuthContext $AuthContext
}

function Get-FunctionDependencyMap {
    [ordered]@{
        'ConvertTo-PlainText' = @()
        'Write-Log' = @()
        'Initialize-RunStats' = @()
        'Add-WaitMetric' = @()
        'Get-HeaderValue' = @()
        'Get-PropertyValue' = @()
        'Mask-SensitiveText' = @()
        'Write-SpApiResponseDebugLog' = @('Get-PropertyValue','Write-Log','Mask-SensitiveText')
        'Wait-ForPricingSlot' = @('Add-WaitMetric')
        'Update-PricingThrottleFromLimit' = @()
        'Invoke-SpApiRequest' = @('Get-ErrorDetail','Get-HeaderValue','Write-SpApiResponseDebugLog','Update-PricingThrottleFromLimit','Write-Log')
        'Get-StatusClassification' = @()
        'Get-ErrorDetail' = @('Get-StatusClassification')
        'Invoke-WithRetry' = @('Write-Log')
        'Get-AmzDateHeaderValue' = @()
        'New-SpApiHeaders' = @('Get-AmzDateHeaderValue')
        'Split-IntoChunks' = @()
        'Read-AccessTokenCache' = @()
        'Save-AccessTokenCache' = @()
        'Get-LwaAccessTokenCached' = @('Read-AccessTokenCache','Save-AccessTokenCache','Get-LwaAccessToken')
        'Get-LwaAccessToken' = @('Invoke-WithRetry')
        'Get-LowestNewPriceFromOffers' = @()
        'Get-AsinMapByJanBatch' = @('Invoke-SpApiRequest','New-SpApiHeaders','Split-IntoChunks','Get-ErrorDetail','Get-StatusClassification','Get-LwaAccessTokenCached')
        'Get-PriceBySingleAsin' = @('Wait-ForPricingSlot','Invoke-SpApiRequest','New-SpApiHeaders','Get-LowestNewPriceFromOffers','Get-ErrorDetail','Get-LwaAccessTokenCached')
        'Get-PriceMapByAsinSequential' = @('Get-PriceBySingleAsin','Get-ErrorDetail')
        'Get-PriceMapByAsinBatch' = @('Get-PriceMapByAsinSequential')
        'Import-PersistentCache' = @('Write-Log')
        'Save-PersistentCache' = @('Write-Log')
        'Add-DailyPriceHistory' = @('Write-Log')
        'Get-JanCacheKey' = @()
        'Get-OfferCacheKey' = @()
        'Get-CacheTtlHoursByStatus' = @()
        'Test-CacheFreshByStatus' = @('Get-CacheTtlHoursByStatus','Test-CacheFresh')
        'Test-CacheFresh' = @()
        'Save-SecretsInteractive' = @()
        'Invoke-AmazonPriceUpdate' = @('Import-PersistentCache','Get-LwaAccessTokenCached','Get-AsinMapByJanBatch','Get-PriceMapByAsinSequential','Save-PersistentCache','Add-DailyPriceHistory','Get-JanCacheKey','Get-OfferCacheKey','Test-CacheFreshByStatus','ConvertTo-PlainText','Initialize-RunStats','Get-ErrorDetail','Write-Log')
    }
}

function Import-PersistentCache {
    param([string]$Path, [string]$LogPath)

    if (-not (Test-Path $Path)) { return @{} }

    try {
        $raw = Get-Content -Path $Path -Raw -Encoding UTF8
        if ([string]::IsNullOrWhiteSpace($raw)) { return @{} }

        $parsed = ConvertFrom-Json -InputObject $raw
        $map = @{}
        if ($parsed) {
            foreach ($item in $parsed) {
                if (-not $item.jan) { continue }
                $map[[string]$item.jan] = [PSCustomObject]@{
                    asin         = $item.asin
                    price        = $item.price
                    fetched_at   = $item.fetched_at
                    cache_status = if ($item.cache_status) { $item.cache_status } else { 'ok' }
                }
            }
        }

        return $map
    }
    catch {
        Write-Log -Message "キャッシュ読込に失敗したため空キャッシュで継続します: $($_.Exception.Message)" -LogPath $LogPath -Level 'WARN'
        return @{}
    }
}

function Save-PersistentCache {
    param([hashtable]$CacheMap, [string]$Path)

    $rows = foreach ($key in $CacheMap.Keys) {
        [PSCustomObject]@{
            jan          = $key
            asin         = $CacheMap[$key].asin
            price        = $CacheMap[$key].price
            fetched_at   = $CacheMap[$key].fetched_at
            cache_status = $CacheMap[$key].cache_status
        }
    }

    $rows | Sort-Object jan | ConvertTo-Json -Depth 5 | Set-Content -Path $Path -Encoding UTF8
}

function Add-DailyPriceHistory {
    param([hashtable]$RunCache, [string]$DirPath)

    if (-not (Test-Path $DirPath)) { New-Item -ItemType Directory -Path $DirPath -Force | Out-Null }

    $historyPath = Join-Path $DirPath "prices_$(Get-Date -Format 'yyyy-MM-dd').jsonl"
    $rows = New-Object System.Collections.Generic.List[string]

    foreach ($jan in $RunCache.Keys) {
        $entry = $RunCache[$jan]
        if (-not $entry -or $entry.cache_status -ne 'ok') { continue }
        if ($null -eq $entry.price -or "$($entry.price)" -eq '') { continue }

        $record = [PSCustomObject]@{
            logged_at    = (Get-Date).ToString('o')
            fetched_at   = $entry.fetched_at
            jan          = $jan
            asin         = $entry.asin
            price        = $entry.price
            cache_status = $entry.cache_status
        }

        $rows.Add(($record | ConvertTo-Json -Compress -Depth 5)) | Out-Null
    }

    if ($rows.Count -eq 0) { return 0 }

    Add-Content -Path $historyPath -Value $rows -Encoding UTF8
    return $rows.Count
}


function Get-JanCacheKey {
    param([string]$MarketplaceId, [string]$Jan)
    return "jan|$MarketplaceId|$Jan"
}

function Get-OfferCacheKey {
    param([string]$MarketplaceId, [string]$Condition, [string]$Asin)
    return "offer|$MarketplaceId|$Condition|$Asin"
}

function Get-CacheTtlHoursByStatus {
    param([string]$Status, [hashtable]$Config, [string]$CacheKind)

    if ($Status -eq 'not_found') {
        if ($Config.NegativeCacheTtlHours) { return [int]$Config.NegativeCacheTtlHours }
        return 12
    }

    if ($CacheKind -eq 'jan') {
        if ($Config.JanAsinCacheTtlHours) { return [int]$Config.JanAsinCacheTtlHours }
        return 168
    }

    if ($Config.OfferCacheTtlHours) { return [int]$Config.OfferCacheTtlHours }
    return 24
}

function Test-CacheFreshByStatus {
    param([object]$Entry, [hashtable]$Config, [string]$CacheKind)
    if (-not $Entry) { return $false }
    $ttl = Get-CacheTtlHoursByStatus -Status $Entry.cache_status -Config $Config -CacheKind $CacheKind
    return Test-CacheFresh -Entry $Entry -TtlHours $ttl
}

function Test-CacheFresh {
    param([object]$Entry, [int]$TtlHours)

    if (-not $Entry -or -not $Entry.fetched_at) { return $false }

    try {
        $fetchedAt = [DateTime]::Parse($Entry.fetched_at)
    }
    catch {
        return $false
    }

    return ((Get-Date) - $fetchedAt).TotalHours -lt $TtlHours
}

function Save-SecretsInteractive {
    param([string]$SecretFile)

    $secretsDir = Split-Path -Path $SecretFile -Parent
    if (-not (Test-Path $secretsDir)) { New-Item -ItemType Directory -Path $secretsDir -Force | Out-Null }

    Write-Host 'Amazon SP-API の認証情報を入力してください。'
    $clientId = Read-Host 'client_id'
    $clientSecret = Read-Host 'client_secret' -AsSecureString
    $refreshToken = Read-Host 'refresh_token' -AsSecureString

    [PSCustomObject]@{
        client_id     = $clientId
        client_secret = $clientSecret
        refresh_token = $refreshToken
        created_at    = (Get-Date).ToString('o')
    } | Export-Clixml -Path $SecretFile

    Write-Host "保存完了: $SecretFile"
    Write-Host 'このファイルはDPAPIで暗号化され、同じWindowsユーザーのみ復号できます。'
}

function Invoke-AmazonPriceUpdate {
    param(
        [string]$RepoRoot,
        [hashtable]$Config
    )

    $paths = $Config.Paths
    $secretFile = Join-Path $RepoRoot $paths.SecretsFile
    $inputPath = Join-Path $RepoRoot $paths.InputFile
    $outputPath = Join-Path $RepoRoot $paths.OutputFile
    $logDir = Join-Path $RepoRoot $paths.LogDir
    $logPath = Join-Path $RepoRoot $paths.LogFile
    $cacheDir = Join-Path $RepoRoot $paths.CacheDir
    $cachePath = Join-Path $RepoRoot $paths.CacheFile
    $historyDir = Join-Path $RepoRoot $paths.HistoryDir
    $accessTokenCachePath = Join-Path $RepoRoot $paths.AccessTokenCacheFile

    foreach ($dir in @($logDir, $cacheDir, $historyDir)) {
        if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
    }

    if (-not (Test-Path $secretFile)) {
        Write-Host 'secrets/lwa_secrets.xml が見つかりません。run_init.bat を先に実行してください。'
        exit 1
    }

    if (-not (Test-Path $inputPath)) {
        Write-Host "入力ファイルが見つかりません: $inputPath"
        exit 1
    }

    $secret = Import-Clixml -Path $secretFile
    $clientId = $secret.client_id
    $clientSecret = ConvertTo-PlainText -Secure $secret.client_secret
    $refreshToken = ConvertTo-PlainText -Secure $secret.refresh_token

    Write-Log -Message '更新処理を開始します。' -LogPath $logPath
    Initialize-RunStats
    $authContext = @{
        ClientId = $clientId
        ClientSecret = $clientSecret
        RefreshToken = $refreshToken
        TokenCachePath = $accessTokenCachePath
    }
    $accessToken = Get-LwaAccessTokenCached -ClientId $clientId -ClientSecret $clientSecret -RefreshToken $refreshToken -Config $Config -LogPath $logPath -TokenCachePath $accessTokenCachePath

    $excel = $null
    $workbook = $null
    $sheet = $null

    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false

        $workbook = $excel.Workbooks.Open($inputPath)
        $sheet = $workbook.Worksheets.Item(1)
        
        # Set headers for G, H, I columns
        $sheet.Cells.Item(1, 7).Value2 = 'ASIN'
        $sheet.Cells.Item(1, 8).Value2 = '価格'
        $sheet.Cells.Item(1, 9).Value2 = '取得日時'
        
        $lastRow = $sheet.Cells($sheet.Rows.Count, 2).End(-4162).Row
        $totalDataRows = [Math]::Max(0, $lastRow - 1)

        $persistentCache = Import-PersistentCache -Path $cachePath -LogPath $logPath
        $runCache = @{}
        $processed = 0
        $transientErrorCount = 0
        $cacheHitCount = 0
        $cacheMissCount = 0
        $catalogApiCalls = 0
        $pricingApiCalls = 0
        $notFoundValidationCount = 0
        $rateLimitServerCount = 0
        $otherErrorCount = 0

        $janByRow = @{}
        $targetJans = New-Object System.Collections.Generic.HashSet[string]

        for ($row = 2; $row -le $lastRow; $row++) {
            $jan = ([string]$sheet.Cells.Item($row, 2).Text).Trim()
            $janByRow[$row] = $jan
            if (-not [string]::IsNullOrWhiteSpace($jan)) { [void]$targetJans.Add($jan) }
        }

        $janList = @($targetJans)
        $needApiJans = @()
        Write-Log -Message "依存チェック: 入力JAN収集完了 (rows=$totalDataRows, unique_jan=$($janList.Count))" -LogPath $logPath

        foreach ($jan in $janList) {
            $janCacheKey = Get-JanCacheKey -MarketplaceId $Config.MarketplaceId -Jan $jan
            if ($persistentCache.ContainsKey($janCacheKey) -and (Test-CacheFreshByStatus -Entry $persistentCache[$janCacheKey] -Config $Config -CacheKind 'jan')) {
                $runCache[$jan] = $persistentCache[$janCacheKey]
                $cacheHitCount++
            }
            else {
                $needApiJans += $jan
                $cacheMissCount++
            }
        }

        $uniqueAsinCount = 0
        Write-Log -Message "依存チェック: JANキャッシュ判定完了 (api_required_jan=$($needApiJans.Count))" -LogPath $logPath
        if ($needApiJans.Count -gt 0) {
            $catalogResult = Get-AsinMapByJanBatch -Jans $needApiJans -AccessToken $accessToken -Config $Config -LogPath $logPath -AuthContext $authContext
            $asinMap = $catalogResult.AsinMap
            $catalogErrorMap = $catalogResult.ErrorClassMap
            $catalogApiCalls = $script:RunStats.CatalogBatchCalls

            $allAsins = @($asinMap.Values | Where-Object { $_ } | Sort-Object -Unique)
            $uniqueAsinCount = $allAsins.Count

            $needPriceAsins = @()
            $priceMap = @{}
            $priceErrorMap = @{}

            foreach ($jan in $needApiJans) {
                $asin = $asinMap[$jan]
                if (-not $asin) { continue }

                $offerKey = Get-OfferCacheKey -MarketplaceId $Config.MarketplaceId -Condition 'New' -Asin $asin
                if ($persistentCache.ContainsKey($offerKey) -and (Test-CacheFreshByStatus -Entry $persistentCache[$offerKey] -Config $Config -CacheKind 'offer')) {
                    $cachedOffer = $persistentCache[$offerKey]
                    $priceMap[$asin] = $cachedOffer.price
                    if ($cachedOffer.cache_status -ne 'ok') {
                        $priceErrorMap[$asin] = if ($cachedOffer.cache_status -eq 'not_found') { 'NotFound/Validation' } else { 'Other' }
                    }
                    $cacheHitCount++
                }
                else {
                    $needPriceAsins += $asin
                }
            }

            if ($needPriceAsins.Count -gt 0) {
                $distinctAsins = @($needPriceAsins | Sort-Object -Unique)
                $pricingResult = Get-PriceMapByAsinSequential -Asins $distinctAsins -AccessToken $accessToken -Config $Config -LogPath $logPath -AuthContext $authContext
                foreach ($k in $pricingResult.PriceMap.Keys) { $priceMap[$k] = $pricingResult.PriceMap[$k] }
                foreach ($k in $pricingResult.ErrorClassMap.Keys) { $priceErrorMap[$k] = $pricingResult.ErrorClassMap[$k] }
                $pricingApiCalls = $script:RunStats.PricingCalls
            }

            Write-Log -Message "依存チェック: Pricing確定 (need_price_asin=$($needPriceAsins.Count), resolved_asin=$uniqueAsinCount)" -LogPath $logPath
            $fetchedAt = (Get-Date).ToString('o')
            foreach ($jan in $needApiJans) {
                $cacheStatus = 'ok'
                $asin = $asinMap[$jan]
                $price = $null

                if (-not $asin) {
                    if ($catalogErrorMap.ContainsKey($jan) -and $catalogErrorMap[$jan] -eq 'NotFound/Validation') {
                        $cacheStatus = 'not_found'; $notFoundValidationCount++
                    }
                    elseif ($catalogErrorMap.ContainsKey($jan) -and $catalogErrorMap[$jan] -eq 'RateLimit/Server') {
                        $cacheStatus = 'transient_error'; $rateLimitServerCount++; $transientErrorCount++
                    }
                    elseif ($catalogErrorMap.ContainsKey($jan)) {
                        $cacheStatus = 'transient_error'; $otherErrorCount++; $transientErrorCount++
                    }
                    else {
                        $cacheStatus = 'not_found'; $notFoundValidationCount++
                    }
                }
                else {
                    if ($priceMap.ContainsKey($asin)) { $price = $priceMap[$asin] }
                    if ($priceErrorMap.ContainsKey($asin)) {
                        $errClass = $priceErrorMap[$asin]
                        if ($errClass -eq 'NotFound/Validation') {
                            $cacheStatus = 'not_found'; $notFoundValidationCount++
                        }
                        elseif ($errClass -eq 'RateLimit/Server') {
                            $cacheStatus = 'transient_error'; $rateLimitServerCount++; $transientErrorCount++
                        }
                        else {
                            $cacheStatus = 'transient_error'; $otherErrorCount++; $transientErrorCount++
                        }
                    }
                }

                $entry = [PSCustomObject]@{ asin = $asin; price = $price; fetched_at = $fetchedAt; cache_status = $cacheStatus }
                $runCache[$jan] = $entry

                $janCacheKey = Get-JanCacheKey -MarketplaceId $Config.MarketplaceId -Jan $jan
                if ($cacheStatus -eq 'not_found' -or $cacheStatus -eq 'ok') { $persistentCache[$janCacheKey] = $entry }
                elseif ($persistentCache.ContainsKey($janCacheKey)) { $persistentCache.Remove($janCacheKey) }

                if ($asin) {
                    $offerKey = Get-OfferCacheKey -MarketplaceId $Config.MarketplaceId -Condition 'New' -Asin $asin
                    if ($cacheStatus -eq 'not_found' -or $cacheStatus -eq 'ok') { $persistentCache[$offerKey] = $entry }
                    elseif ($persistentCache.ContainsKey($offerKey)) { $persistentCache.Remove($offerKey) }
                }
            }
        }

        for ($row = 2; $row -le $lastRow; $row++) {
            $jan = $janByRow[$row]
            $currentIndex = $row - 1

            if ($totalDataRows -gt 0) {
                $percent = [int](($currentIndex * 100) / $totalDataRows)
                # 50件ごとに進捗表示
                $shouldReport = (($currentIndex % 50) -eq 0) -or ($currentIndex -eq 1) -or ($currentIndex -eq $totalDataRows)
                if ($shouldReport) {
                    # avoid blue progress bar; just log text
                    Write-Host "進捗: $currentIndex / $totalDataRows"
                }
            }

            if ([string]::IsNullOrWhiteSpace($jan)) {
                $sheet.Cells.Item($row, 7).Value2 = ''
                $sheet.Cells.Item($row, 8).Value2 = ''
                $sheet.Cells.Item($row, 9).Value2 = ''
                continue
            }

            try {
                $result = $runCache[$jan]
                if ($result -and ($result.cache_status -eq 'not_found' -or $result.cache_status -eq 'transient_error')) {
                    $sheet.Cells.Item($row, 7).Value2 = ''
                    $sheet.Cells.Item($row, 8).Value2 = ''
                    $sheet.Cells.Item($row, 9).Value2 = ''
                    if ($result.cache_status -eq 'transient_error') {
                        Write-Log -Message "行$row JAN=$jan は一時エラーのため空欄出力します。" -LogPath $logPath -Level 'WARN'
                    }
                    continue
                }

                $sheet.Cells.Item($row, 7).Value2 = if ($result -and $result.asin) { $result.asin } else { '' }
                if ($result -and $null -ne $result.price -and "$($result.price)" -ne '') {
                    $sheet.Cells.Item($row, 8).Value2 = [double]$result.price
                    $sheet.Cells.Item($row, 9).Value2 = if ($result.fetched_at) { [string]$result.fetched_at } else { '' }
                }
                else {
                    $sheet.Cells.Item($row, 8).Value2 = ''
                    $sheet.Cells.Item($row, 9).Value2 = ''
                }
            }
            catch {
                $detail = Get-ErrorDetail -ErrorRecord $_
                if ($detail.Class -eq 'NotFound/Validation') { $notFoundValidationCount++ }
                elseif ($detail.Class -eq 'RateLimit/Server') { $rateLimitServerCount++; $transientErrorCount++ }
                else { $otherErrorCount++; $transientErrorCount++ }

                $sheet.Cells.Item($row, 7).Value2 = ''
                $sheet.Cells.Item($row, 8).Value2 = ''
                $sheet.Cells.Item($row, 9).Value2 = ''
                Write-Log -Message "行$row JAN=$jan の処理でエラー: 分類=$($detail.Class), HTTP=$($detail.StatusCode), msg=$($_.Exception.Message)" -LogPath $logPath -Level 'ERROR'
            }

            $processed++
        }

        Write-Log -Message "依存チェック: runCache確定・Excel反映完了" -LogPath $logPath
        Save-PersistentCache -CacheMap $persistentCache -Path $cachePath
        $historySavedCount = Add-DailyPriceHistory -RunCache $runCache -DirPath $historyDir
        Write-Log -Message "価格履歴の追記件数: $historySavedCount" -LogPath $logPath

        try {
            $workbook.SaveAs($outputPath)
        }
        catch {
            Write-Host 'output.xlsx を保存できませんでした。Excelを閉じてから再実行してください。'
            throw
        }

        $avgWait = if ($script:RunStats.WaitEvents -gt 0) { [Math]::Round($script:RunStats.TotalWaitSec / $script:RunStats.WaitEvents, 2) } else { 0 }
        $apiReducedBase = [Math]::Max(1, $uniqueAsinCount)
        $pricingReductionPct = [Math]::Round((1 - ([double]$pricingApiCalls / [double]$apiReducedBase)) * 100, 2)
        Write-Log -Message "呼び出し統計: input_rows=$totalDataRows, unique_jan=$($janList.Count), unique_asin=$uniqueAsinCount, cache_hit=$cacheHitCount, cache_miss=$cacheMissCount, catalog_calls=$catalogApiCalls, pricing_calls=$pricingApiCalls, pricing_reduction_pct=$pricingReductionPct" -LogPath $logPath
        Write-Log -Message "再試行統計: api_total_calls=$($script:RunStats.TotalApiCalls), retry_count=$($script:RunStats.RetryCount), http429_count=$($script:RunStats.Http429Count), total_wait_sec=$([Math]::Round($script:RunStats.TotalWaitSec,2)), avg_wait_sec=$avgWait" -LogPath $logPath
        $metricsPath = Join-Path $logDir 'metrics.jsonl'
        $metricsRecord = [PSCustomObject]@{ ts=(Get-Date).ToString('o'); input_rows=$totalDataRows; unique_jan=$($janList.Count); unique_asin=$uniqueAsinCount; pricing_calls=$pricingApiCalls; pricing_reduction_pct=$pricingReductionPct; api_total_calls=$($script:RunStats.TotalApiCalls); retry_count=$($script:RunStats.RetryCount); http429_count=$($script:RunStats.Http429Count); total_wait_sec=[Math]::Round($script:RunStats.TotalWaitSec,2); avg_wait_sec=$avgWait }
        Add-Content -Path $metricsPath -Value ($metricsRecord | ConvertTo-Json -Compress -Depth 5) -Encoding UTF8
        $totalUnresolvedCount = $notFoundValidationCount + $rateLimitServerCount + $otherErrorCount
        Write-Log -Message "エラー分類統計: NotFound/Validation=$notFoundValidationCount, RateLimit/Server=$rateLimitServerCount, Other=$otherErrorCount" -LogPath $logPath
        Write-Log -Message "更新完了: 処理件数=$processed, 一時エラー件数=$transientErrorCount, 未解決件数=$totalUnresolvedCount, 出力=$outputPath" -LogPath $logPath
    }
    finally {
        Write-Progress -Activity 'Excel出力処理' -Completed
        if ($workbook) { $workbook.Close($false) }
        if ($excel) { $excel.Quit() }
        if ($sheet) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($sheet) }
        if ($workbook) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) }
        if ($excel) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel) }
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}

Export-ModuleMember -Function ConvertTo-PlainText,Write-Log,Get-StatusClassification,Get-ErrorDetail,Invoke-WithRetry,Get-AmzDateHeaderValue,New-SpApiHeaders,Split-IntoChunks,Read-AccessTokenCache,Save-AccessTokenCache,Get-LwaAccessTokenCached,Get-LwaAccessToken,Get-LowestNewPriceFromOffers,Get-PriceBySingleAsin,Get-AsinMapByJanBatch,Get-PriceMapByAsinSequential,Get-PriceMapByAsinBatch,Get-FunctionDependencyMap,Import-PersistentCache,Save-PersistentCache,Add-DailyPriceHistory,Test-CacheFresh,Save-SecretsInteractive,Invoke-AmazonPriceUpdate
