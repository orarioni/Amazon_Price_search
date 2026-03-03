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
    Write-Host $line
}

function Classify-StatusAndBody {
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

    if ($ErrorRecord -and $ErrorRecord.Exception -and $ErrorRecord.Exception.Response) {
        try {
            if ($ErrorRecord.Exception.Response.StatusCode) {
                $statusCode = [int]$ErrorRecord.Exception.Response.StatusCode
            }
        }
        catch {}

        try {
            $headers = $ErrorRecord.Exception.Response.Headers
            if ($headers) {
                $retryAfterRaw = $headers['Retry-After']
                if (-not $retryAfterRaw) { $retryAfterRaw = $headers['retry-after'] }
                if ($retryAfterRaw) {
                    $retryAfterInt = 0
                    if ([int]::TryParse([string]$retryAfterRaw, [ref]$retryAfterInt)) {
                        $retryAfterSec = [Math]::Max(1, $retryAfterInt)
                    }
                    else {
                        $retryAfterDate = $null
                        if ([DateTime]::TryParse([string]$retryAfterRaw, [ref]$retryAfterDate)) {
                            $delta = [int][Math]::Ceiling(($retryAfterDate - (Get-Date)).TotalSeconds)
                            if ($delta -gt 0) { $retryAfterSec = $delta }
                        }
                    }
                }

                $rateLimitLimit = $headers['x-amzn-RateLimit-Limit']
                if (-not $rateLimitLimit) { $rateLimitLimit = $headers['X-Amzn-RateLimit-Limit'] }
            }
        }
        catch {}

        try {
            $stream = $ErrorRecord.Exception.Response.GetResponseStream()
            if ($stream) {
                $reader = New-Object System.IO.StreamReader($stream)
                $bodyText = $reader.ReadToEnd()
                $reader.Close()
            }
        }
        catch {}
    }

    $classification = Classify-StatusAndBody -StatusCode $statusCode -BodyText $bodyText
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

        $expiresAt = $null
        if (-not [DateTime]::TryParse($parsed.expires_at, [ref]$expiresAt)) { return $null }

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

    while ($index -lt $Jans.Count) {
        $end = [Math]::Min($index + $batchSize - 1, $Jans.Count - 1)
        $chunk = @($Jans[$index..$end])
        foreach ($jan in $chunk) { $resultMap[$jan] = $null }

        $identifiers = ($chunk | ForEach-Object { $_.Trim() }) -join ','
        $uri = "$($Config.SpApiBaseUrl)/catalog/2022-04-01/items?identifiers=$([Uri]::EscapeDataString($identifiers))&identifiersType=EAN&marketplaceIds=$($Config.MarketplaceId)"

        $res = $null
        $attemptDetail = $null
        try {
            $res = Invoke-WithRetry -Label "Catalogバッチ取得(index=$index,size=$batchSize)" -MaxRetries $Config.MaxRetries -LogPath $LogPath -Action {
                Invoke-RestMethod -Method Get -Uri $uri -Headers (New-SpApiHeaders -AccessToken $AccessToken -Config $Config)
            }
        }
        catch {
            $attemptDetail = Get-ErrorDetail -ErrorRecord $_
            if ($attemptDetail.Class -eq 'Auth' -and $AuthContext) {
                try {
                    $AccessToken = Get-LwaAccessTokenCached -ClientId $AuthContext.ClientId -ClientSecret $AuthContext.ClientSecret -RefreshToken $AuthContext.RefreshToken -Config $Config -LogPath $LogPath -TokenCachePath $AuthContext.TokenCachePath -ForceRefresh
                    $res = Invoke-WithRetry -Label "Catalogバッチ再取得(認証更新,index=$index,size=$batchSize)" -MaxRetries $Config.MaxRetries -LogPath $LogPath -Action {
                        Invoke-RestMethod -Method Get -Uri $uri -Headers (New-SpApiHeaders -AccessToken $AccessToken -Config $Config)
                    }
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

        if ($res.items) {
            foreach ($item in $res.items) {
                if (-not $item.identifiers -or -not $item.identifiers.identifiers) { continue }

                $matchedJan = $null
                foreach ($idGroup in $item.identifiers.identifiers) {
                    if ($idGroup.identifierType -eq 'EAN' -and $idGroup.identifier) {
                        $matchedJan = [string]$idGroup.identifier
                        break
                    }
                }

                if ($matchedJan -and $resultMap.ContainsKey($matchedJan) -and $item.asin) {
                    $resultMap[$matchedJan] = [string]$item.asin
                    $errorClassMap.Remove($matchedJan) | Out-Null
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
        $res = Invoke-WithRetry -Label "Pricing単発取得(ASIN=$Asin)" -MaxRetries $Config.MaxRetries -LogPath $LogPath -Action {
            Invoke-RestMethod -Method Get -Uri $uri -Headers (New-SpApiHeaders -AccessToken $AccessToken -Config $Config)
        }
    }
    catch {
        $detail = Get-ErrorDetail -ErrorRecord $_
        if ($detail.Class -eq 'Auth' -and $AuthContext) {
            $AccessToken = Get-LwaAccessTokenCached -ClientId $AuthContext.ClientId -ClientSecret $AuthContext.ClientSecret -RefreshToken $AuthContext.RefreshToken -Config $Config -LogPath $LogPath -TokenCachePath $AuthContext.TokenCachePath -ForceRefresh
            $res = Invoke-WithRetry -Label "Pricing単発再取得(認証更新,ASIN=$Asin)" -MaxRetries $Config.MaxRetries -LogPath $LogPath -Action {
                Invoke-RestMethod -Method Get -Uri $uri -Headers (New-SpApiHeaders -AccessToken $AccessToken -Config $Config)
            }
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

function Get-PriceMapByAsinBatch {
    param(
        [array]$Asins,
        [string]$AccessToken,
        [hashtable]$Config,
        [string]$LogPath,
        [hashtable]$AuthContext
    )

    $priceMap = @{}
    $errorClassMap = @{}

    $singleThreshold = if ($Config.PricingSingleFallbackThreshold) { [int]$Config.PricingSingleFallbackThreshold } else { 3 }
    if ($Asins.Count -le $singleThreshold) {
        Write-Log -Message "ASIN件数=$($Asins.Count) のため Pricing単発APIへフォールバックします。" -LogPath $LogPath
        foreach ($asin in $Asins) {
            try {
                $single = Get-PriceBySingleAsin -Asin $asin -AccessToken $AccessToken -Config $Config -LogPath $LogPath -AuthContext $AuthContext
                $priceMap[$asin] = $single.Price
                if ($single.ErrorClass) { $errorClassMap[$asin] = $single.ErrorClass }
                else { $errorClassMap.Remove($asin) | Out-Null }
            }
            catch {
                $detail = Get-ErrorDetail -ErrorRecord $_
                $priceMap[$asin] = $null
                $errorClassMap[$asin] = if ($detail.Class) { $detail.Class } else { 'Other' }
            }
        }

        return [PSCustomObject]@{ PriceMap = $priceMap; ErrorClassMap = $errorClassMap }
    }

    $minBatchSize = 5
    $batchSize = [Math]::Max($minBatchSize, [int]$Config.PricingBatchSize)
    $index = 0

    while ($index -lt $Asins.Count) {
        $end = [Math]::Min($index + $batchSize - 1, $Asins.Count - 1)
        $chunk = @($Asins[$index..$end])
        foreach ($asin in $chunk) { $priceMap[$asin] = $null }

        $requests = @()
        foreach ($asin in $chunk) {
            $requests += @{ uri = "/products/pricing/v0/items/$([Uri]::EscapeDataString($asin))/offers?MarketplaceId=$($Config.MarketplaceId)&ItemCondition=New"; method = 'GET' }
        }

        $body = @{ requests = $requests } | ConvertTo-Json -Depth 5

        $res = $null
        $attemptDetail = $null
        try {
            $res = Invoke-WithRetry -Label "Pricingバッチ取得(index=$index,size=$batchSize)" -MaxRetries $Config.MaxRetries -LogPath $LogPath -Action {
                Invoke-RestMethod -Method Post -Uri "$($Config.SpApiBaseUrl)/batches/products/pricing/v0/itemOffers" -Headers (New-SpApiHeaders -AccessToken $AccessToken -Config $Config -ContentType 'application/json') -Body $body
            }
        }
        catch {
            $attemptDetail = Get-ErrorDetail -ErrorRecord $_
            if ($attemptDetail.Class -eq 'Auth' -and $AuthContext) {
                try {
                    $AccessToken = Get-LwaAccessTokenCached -ClientId $AuthContext.ClientId -ClientSecret $AuthContext.ClientSecret -RefreshToken $AuthContext.RefreshToken -Config $Config -LogPath $LogPath -TokenCachePath $AuthContext.TokenCachePath -ForceRefresh
                    $res = Invoke-WithRetry -Label "Pricingバッチ再取得(認証更新,index=$index,size=$batchSize)" -MaxRetries $Config.MaxRetries -LogPath $LogPath -Action {
                        Invoke-RestMethod -Method Post -Uri "$($Config.SpApiBaseUrl)/batches/products/pricing/v0/itemOffers" -Headers (New-SpApiHeaders -AccessToken $AccessToken -Config $Config -ContentType 'application/json') -Body $body
                    }
                }
                catch {
                    $attemptDetail = Get-ErrorDetail -ErrorRecord $_
                }
            }
        }

        if (-not $res) {
            if ($attemptDetail -and $attemptDetail.Class -eq 'RateLimit/Server' -and $batchSize -gt $minBatchSize) {
                $nextBatchSize = [Math]::Max($minBatchSize, [int][Math]::Floor($batchSize / 2))
                Write-Log -Message "Pricingバッチを縮小します: $batchSize -> $nextBatchSize (index=$index, HTTP=$($attemptDetail.StatusCode), limit=$($attemptDetail.RateLimitLimit))" -LogPath $LogPath -Level 'WARN'
                $batchSize = $nextBatchSize
                continue
            }

            if ($chunk.Count -le $singleThreshold) {
                Write-Log -Message "Pricingバッチ失敗のため単発APIへフォールバックします (index=$index,size=$($chunk.Count))。" -LogPath $LogPath -Level 'WARN'
                foreach ($asin in $chunk) {
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
                $index = $end + 1
                continue
            }

            foreach ($asin in $chunk) {
                $errorClassMap[$asin] = if ($attemptDetail) { $attemptDetail.Class } else { 'Other' }
            }
            $index = $end + 1
            continue
        }

        if (-not $res.responses) {
            foreach ($asin in $chunk) { $errorClassMap[$asin] = 'Other' }
            $index = $end + 1
            continue
        }

        foreach ($response in $res.responses) {
            $statusCode = if ($response.status) { [int]$response.status } else { $null }

            $asin = $null
            if ($response.body -and $response.body.payload -and $response.body.payload.ASIN) {
                $asin = [string]$response.body.payload.ASIN
            }
            elseif ($response.request -and $response.request.uri) {
                if ($response.request.uri -match '/items/([^/]+)/offers') {
                    $asin = [Uri]::UnescapeDataString($matches[1])
                }
            }

            if (-not $asin -or -not $priceMap.ContainsKey($asin)) { continue }

            if ($statusCode -ge 400) {
                $bodyText = if ($response.body) { ($response.body | ConvertTo-Json -Depth 8) } else { '' }
                $detail = Classify-StatusAndBody -StatusCode $statusCode -BodyText $bodyText
                $errorClassMap[$asin] = $detail.Class
                continue
            }

            $offers = $response.body.payload.Offers
            $priceMap[$asin] = Get-LowestNewPriceFromOffers -Offers $offers
            if ($null -eq $priceMap[$asin]) {
                $errorClassMap[$asin] = 'NotFound/Validation'
            }
            else {
                $errorClassMap.Remove($asin) | Out-Null
            }
        }

        $index = $end + 1
    }

    [PSCustomObject]@{ PriceMap = $priceMap; ErrorClassMap = $errorClassMap }
}

function Load-PersistentCache {
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

function Append-DailyPriceHistory {
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

function Is-CacheFresh {
    param([object]$Entry, [int]$TtlHours)

    if (-not $Entry -or -not $Entry.fetched_at) { return $false }

    $fetchedAt = $null
    if (-not [DateTime]::TryParse($Entry.fetched_at, [ref]$fetchedAt)) { return $false }

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
        $lastRow = $sheet.Cells($sheet.Rows.Count, 2).End(-4162).Row
        $totalDataRows = [Math]::Max(0, $lastRow - 1)

        $persistentCache = Load-PersistentCache -Path $cachePath -LogPath $logPath
        $runCache = @{}
        $processed = 0
        $errorCount = 0
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

        foreach ($jan in $janList) {
            if ($persistentCache.ContainsKey($jan) -and (Is-CacheFresh -Entry $persistentCache[$jan] -TtlHours $Config.CacheTtlHours)) {
                $runCache[$jan] = $persistentCache[$jan]
                $cacheHitCount++
            }
            else {
                $needApiJans += $jan
                $cacheMissCount++
            }
        }

        if ($needApiJans.Count -gt 0) {
            $catalogResult = Get-AsinMapByJanBatch -Jans $needApiJans -AccessToken $accessToken -Config $Config -LogPath $logPath -AuthContext $authContext
            $asinMap = $catalogResult.AsinMap
            $catalogErrorMap = $catalogResult.ErrorClassMap
            $catalogApiCalls = (Split-IntoChunks -Items $needApiJans -ChunkSize $Config.CatalogBatchSize).Count

            $needPriceAsins = @()
            foreach ($jan in $needApiJans) {
                $asin = $asinMap[$jan]
                if ($asin) { $needPriceAsins += $asin }
            }

            $priceMap = @{}
            $priceErrorMap = @{}
            if ($needPriceAsins.Count -gt 0) {
                $distinctAsins = @($needPriceAsins | Sort-Object -Unique)
                $pricingResult = Get-PriceMapByAsinBatch -Asins $distinctAsins -AccessToken $accessToken -Config $Config -LogPath $logPath -AuthContext $authContext
                $priceMap = $pricingResult.PriceMap
                $priceErrorMap = $pricingResult.ErrorClassMap
                $singleThreshold = if ($Config.PricingSingleFallbackThreshold) { [int]$Config.PricingSingleFallbackThreshold } else { 3 }
                if ($distinctAsins.Count -le $singleThreshold) {
                    $pricingApiCalls = $distinctAsins.Count
                }
                else {
                    $pricingApiCalls = (Split-IntoChunks -Items $distinctAsins -ChunkSize $Config.PricingBatchSize).Count
                }
            }

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
                        $cacheStatus = 'transient_error'; $rateLimitServerCount++; $errorCount++
                    }
                    elseif ($catalogErrorMap.ContainsKey($jan)) {
                        $cacheStatus = 'transient_error'; $otherErrorCount++; $errorCount++
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
                            $cacheStatus = 'transient_error'; $rateLimitServerCount++; $errorCount++
                        }
                        else {
                            $cacheStatus = 'transient_error'; $otherErrorCount++; $errorCount++
                        }
                    }
                }

                $entry = [PSCustomObject]@{ asin = $asin; price = $price; fetched_at = $fetchedAt; cache_status = $cacheStatus }
                $runCache[$jan] = $entry
                if ($cacheStatus -eq 'not_found' -or $cacheStatus -eq 'ok') { $persistentCache[$jan] = $entry }
                elseif ($persistentCache.ContainsKey($jan)) { $persistentCache.Remove($jan) }
            }
        }

        for ($row = 2; $row -le $lastRow; $row++) {
            $jan = $janByRow[$row]
            $currentIndex = $row - 1

            if ($totalDataRows -gt 0) {
                $percent = [int](($currentIndex * 100) / $totalDataRows)
                $shouldReport = (($currentIndex % 10) -eq 0) -or ($currentIndex -eq 1) -or ($currentIndex -eq $totalDataRows)
                if ($shouldReport) {
                    Write-Progress -Activity 'Excel出力処理' -Status "$currentIndex / $totalDataRows 行を処理中" -PercentComplete $percent
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
                $errorCount++
                if ($detail.Class -eq 'NotFound/Validation') { $notFoundValidationCount++ }
                elseif ($detail.Class -eq 'RateLimit/Server') { $rateLimitServerCount++ }
                else { $otherErrorCount++ }

                $sheet.Cells.Item($row, 7).Value2 = ''
                $sheet.Cells.Item($row, 8).Value2 = ''
                $sheet.Cells.Item($row, 9).Value2 = ''
                Write-Log -Message "行$row JAN=$jan の処理でエラー: 分類=$($detail.Class), HTTP=$($detail.StatusCode), msg=$($_.Exception.Message)" -LogPath $logPath -Level 'ERROR'
            }

            $processed++
        }

        Save-PersistentCache -CacheMap $persistentCache -Path $cachePath
        $historySavedCount = Append-DailyPriceHistory -RunCache $runCache -DirPath $historyDir
        Write-Log -Message "価格履歴の追記件数: $historySavedCount" -LogPath $logPath

        try {
            $workbook.SaveAs($outputPath)
        }
        catch {
            Write-Host 'output.xlsx を保存できませんでした。Excelを閉じてから再実行してください。'
            throw
        }

        Write-Log -Message "呼び出し統計: JAN総数=$($janList.Count), cache_hit=$cacheHitCount, cache_miss=$cacheMissCount, catalog_calls=$catalogApiCalls, pricing_calls=$pricingApiCalls" -LogPath $logPath
        Write-Log -Message "エラー分類統計: NotFound/Validation=$notFoundValidationCount, RateLimit/Server=$rateLimitServerCount, Other=$otherErrorCount" -LogPath $logPath
        Write-Log -Message "更新完了: 処理件数=$processed, エラー件数=$errorCount, 出力=$outputPath" -LogPath $logPath
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

Export-ModuleMember -Function ConvertTo-PlainText,Write-Log,Classify-StatusAndBody,Get-ErrorDetail,Invoke-WithRetry,Get-AmzDateHeaderValue,New-SpApiHeaders,Split-IntoChunks,Read-AccessTokenCache,Save-AccessTokenCache,Get-LwaAccessTokenCached,Get-LwaAccessToken,Get-LowestNewPriceFromOffers,Get-PriceBySingleAsin,Get-AsinMapByJanBatch,Get-PriceMapByAsinBatch,Load-PersistentCache,Save-PersistentCache,Append-DailyPriceHistory,Is-CacheFresh,Save-SecretsInteractive,Invoke-AmazonPriceUpdate
