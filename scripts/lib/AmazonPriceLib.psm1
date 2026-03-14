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
        PricingBatchCalls  = 0
        CatalogBatchCalls  = 0
        RetryCount         = 0
        Http429Count       = 0
        TotalWaitSec       = 0.0
        WaitEvents         = 0
        NextPricingAllowedAt = Get-Date
        NextCatalogAllowedAt = Get-Date
        PricingCooldownSec = 0.0
        PricingIntervalSec = 12.0
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

function Get-StatusCodeValue {
    param([object]$Status)

    if ($null -eq $Status) { return $null }

    if ($Status -is [int] -or $Status -is [long] -or $Status -is [double] -or $Status -is [decimal]) {
        return [int]$Status
    }

    if ($Status -is [string]) {
        $statusText = $Status.Trim()
        if ($statusText -match '^\d+$') {
            return [int]$statusText
        }
        return $null
    }

    $statusCode = Get-PropertyValue -Object $Status -Name 'statusCode'
    if ($null -eq $statusCode) {
        $statusCode = Get-PropertyValue -Object $Status -Name 'StatusCode'
    }
    $fromProp = Get-StatusCodeValue -Status $statusCode
    if ($null -ne $fromProp) { return $fromProp }

    if ($Status -is [System.Collections.IDictionary]) {
        foreach ($k in @('statusCode', 'StatusCode')) {
            if ($Status.Contains($k)) {
                $fromDict = Get-StatusCodeValue -Status $Status[$k]
                if ($null -ne $fromDict) { return $fromDict }
            }
        }
    }

    return $null
}


function ConvertTo-ObjectArray {
    param([object]$Value)

    if ($null -eq $Value) { return @() }
    if ($Value -is [string]) { return @($Value) }
    if ($Value -is [System.Array]) { return @($Value) }
    if ($Value -is [System.Collections.IList]) { return @($Value) }
    if ($Value -is [System.Collections.IEnumerable]) {
        return @($Value)
    }

    return @($Value)
}

function Get-ObjectTypeName {
    param([object]$Value)

    if ($null -eq $Value) { return '<null>' }

    try {
        return $Value.GetType().FullName
    }
    catch {
        return '<unknown>'
    }
}

function Get-IdentifierMatchKeys {
    param([string]$Identifier)

    $keys = New-Object System.Collections.Generic.List[string]
    if ([string]::IsNullOrWhiteSpace($Identifier)) { return @() }

    $trimmed = $Identifier.Trim()
    if (-not [string]::IsNullOrWhiteSpace($trimmed)) {
        $keys.Add($trimmed)
    }

    $digits = [regex]::Replace($trimmed, '[^0-9]', '')
    if (-not [string]::IsNullOrWhiteSpace($digits)) {
        $keys.Add($digits)
        $trimmedZero = $digits.TrimStart('0')
        if (-not [string]::IsNullOrWhiteSpace($trimmedZero)) {
            $keys.Add($trimmedZero)
        }
    }

    return @($keys | Select-Object -Unique)
}

function Find-TargetJanByIdentifier {
    param(
        [string]$Identifier,
        [hashtable]$JanLookupMap
    )

    if (-not $JanLookupMap) { return $null }

    foreach ($key in (Get-IdentifierMatchKeys -Identifier $Identifier)) {
        if ($JanLookupMap.ContainsKey($key)) {
            return [string]$JanLookupMap[$key]
        }
    }

    return $null
}

function Get-CatalogIdentifierSampleText {
    param([object]$Items)

    if (-not $Items) { return '<no-items>' }

    $samples = New-Object System.Collections.Generic.List[string]
    foreach ($item in (ConvertTo-ObjectArray -Value $Items)) {
        $itemIdentifiers = Get-PropertyValue -Object $item -Name 'identifiers'
        if (-not $itemIdentifiers) { continue }

        $groups = Get-PropertyValue -Object $itemIdentifiers -Name 'identifiers'
        if ($groups) { $groups = ConvertTo-ObjectArray -Value $groups }
        else { $groups = ConvertTo-ObjectArray -Value $itemIdentifiers }

        foreach ($group in $groups) {
            $leafs = Get-PropertyValue -Object $group -Name 'identifiers'
            if ($leafs) { $leafs = ConvertTo-ObjectArray -Value $leafs }
            else { $leafs = ConvertTo-ObjectArray -Value $group }

            foreach ($leaf in $leafs) {
                $identifierType = [string](Get-PropertyValue -Object $leaf -Name 'identifierType')
                $identifierValue = [string](Get-PropertyValue -Object $leaf -Name 'identifier')
                if ([string]::IsNullOrWhiteSpace($identifierValue)) {
                    $identifierValue = [string](Get-PropertyValue -Object $leaf -Name 'value')
                }
                if ([string]::IsNullOrWhiteSpace($identifierValue)) { continue }

                $samples.Add(("{0}:{1}" -f $identifierType, $identifierValue))
                if ($samples.Count -ge 5) {
                    return ($samples -join ';')
                }
            }
        }
    }

    if ($samples.Count -eq 0) { return '<no-identifiers>' }
    return ($samples -join ';')
}

function Get-CatalogItemPropertySample {
    param([object]$Items)

    if (-not $Items) { return '<no-items>' }

    $first = (Expand-CatalogItems -Items $Items | Select-Object -First 1)
    if (-not $first) { return '<no-expanded-items>' }

    $props = @($first.PSObject.Properties.Name)
    if (@($props).Count -eq 0) { return '<no-properties>' }

    return ($props -join ',')
}

function Format-PreviewList {
    param(
        [array]$Items,
        [int]$MaxCount = 20
    )

    if (-not $Items) { return '<none>' }

    $arr = @($Items)
    if ($arr.Count -eq 0) { return '<none>' }

    $preview = @($arr | Select-Object -First $MaxCount)
    $text = ($preview -join ',')
    if ($arr.Count -gt $MaxCount) {
        return "$text ...(+$(($arr.Count - $MaxCount)) more)"
    }

    return $text
}

function Expand-CatalogItems {
    param([object]$Items)

    if ($null -eq $Items) { return @() }

    $arr = ConvertTo-ObjectArray -Value $Items
    if (@($arr).Count -eq 1) {
        $single = $arr[0]
        $singleAsin = Get-PropertyValue -Object $single -Name 'asin'
        if (-not $singleAsin) {
            $nestedItems = Get-PropertyValue -Object $single -Name 'items'
            if ($nestedItems) {
                return ConvertTo-ObjectArray -Value $nestedItems
            }

            $nestedItem = Get-PropertyValue -Object $single -Name 'item'
            if ($nestedItem) {
                return ConvertTo-ObjectArray -Value $nestedItem
            }

            $propertyNames = @($single.PSObject.Properties.Name)
            if ($propertyNames.Count -gt 0 -and @($propertyNames | Where-Object { $_ -match '^\d+$' }).Count -gt 0) {
                $values = @()
                foreach ($name in $propertyNames | Sort-Object {[int]$_}) {
                    $values += $single.PSObject.Properties[$name].Value
                }
                if (@($values).Count -gt 0) {
                    return @($values)
                }
            }
        }
    }

    return @($arr)
}

function ConvertFrom-JsonIfNeeded {
    param([object]$Value)

    if ($null -eq $Value) {
        return $null
    }

    if ($Value -is [string]) {
        $text = [string]$Value
        $trimmed = $text.TrimStart()
        if ($trimmed.StartsWith('{') -or $trimmed.StartsWith('[')) {
            try {
                return ($text | ConvertFrom-Json -Depth 100)
            }
            catch {
                return $Value
            }
        }
    }

    return $Value
}

function Write-SpApiResponseShapeLog {
    param(
        [string]$Endpoint,
        [object]$Response,
        [string]$LogPath
    )

    if ($null -eq $Response) {
        Write-Log -Message "$Endpoint response-shape: <null>" -LogPath $LogPath -Level 'WARN'
        return
    }

    $responseType = $Response.GetType().FullName
    $topProperties = @($Response.PSObject.Properties.Name)
    $topPropertyText = if ($topProperties.Count -gt 0) { ($topProperties -join ',') } else { '<none>' }

    $items = Get-PropertyValue -Object $Response -Name 'items'
    $responses = Get-PropertyValue -Object $Response -Name 'responses'
    $payload = Get-PropertyValue -Object $Response -Name 'payload'
    $errors = Get-PropertyValue -Object $Response -Name 'errors'

    $itemsCount = if ($items) { @($items).Count } else { 0 }
    $responsesCount = if ($responses) { @($responses).Count } else { 0 }
    $errorCount = if ($errors) { @($errors).Count } else { 0 }
    $hasPayload = $null -ne $payload

    Write-Log -Message "$Endpoint response-shape: type=$responseType props=[$topPropertyText] items.count=$itemsCount responses.count=$responsesCount errors.count=$errorCount hasPayload=$hasPayload" -LogPath $LogPath -Level 'WARN'
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

    $batchResponses = Get-PropertyValue -Object $Response -Name 'responses'
    if ($batchResponses) {
        $index = 0
        $batchResponsesArray = @($batchResponses)
        foreach ($item in $batchResponsesArray) {
            $index++
            $statusRaw = Get-PropertyValue -Object $item -Name 'status'
            $statusCode = Get-StatusCodeValue -Status $statusRaw
            # Guard known crash: status can be PSCustomObject/hashtable and must never be cast directly to [int].
            $status = if ($null -ne $statusCode) { $statusCode } else { '' }
            $request = Get-PropertyValue -Object $item -Name 'request'
            $requestUri = Get-PropertyValue -Object $request -Name 'uri'
            $body = Get-PropertyValue -Object $item -Name 'body'
            $errors = Get-PropertyValue -Object $body -Name 'errors'
            $payload = Get-PropertyValue -Object $body -Name 'payload'
            $asin = Get-PropertyValue -Object $payload -Name 'ASIN'
            $offers = Get-PropertyValue -Object $payload -Name 'Offers'
            $offersCount = if ($offers) { @($offers).Count } else { 0 }
            $errorCount = if ($errors) { @($errors).Count } else { 0 }

            Write-Log -Message "$Endpoint debug[$index/$($batchResponsesArray.Count)]: status=$status request.uri=$requestUri payload.ASIN=$asin offers.count=$offersCount errors.count=$errorCount" -LogPath $LogPath
            if ((($statusCode -as [int]) -ge 400) -or $errorCount -gt 0) { $shouldLogFull = $true }
        }
    }
    else {
        $statusRaw = Get-PropertyValue -Object $Response -Name 'status'
        $statusCode = Get-StatusCodeValue -Status $statusRaw
        # Guard known crash: status can be PSCustomObject/hashtable and must never be cast directly to [int].
        $status = if ($null -ne $statusCode) { $statusCode } else { '' }
        $payload = Get-PropertyValue -Object $Response -Name 'payload'
        $errors = Get-PropertyValue -Object $Response -Name 'errors'
        $requestUri = Get-PropertyValue -Object $Response -Name 'uri'
        $asin = Get-PropertyValue -Object $payload -Name 'ASIN'
        $offers = Get-PropertyValue -Object $payload -Name 'Offers'
        $offersCount = if ($offers) { @($offers).Count } else { 0 }
        $errorCount = if ($errors) { @($errors).Count } else { 0 }

        Write-Log -Message "$Endpoint debug: status=$status request.uri=$requestUri payload.ASIN=$asin offers.count=$offersCount errors.count=$errorCount" -LogPath $LogPath
        if ((($statusCode -as [int]) -ge 400) -or $errorCount -gt 0) { $shouldLogFull = $true }
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

    $defaultInterval = if ($Config.PricingDefaultIntervalSec) { [double]$Config.PricingDefaultIntervalSec } else { 12.0 }
    if (-not $script:RunStats.Contains('PricingIntervalSec')) {
        $script:RunStats.PricingIntervalSec = $defaultInterval
    }

    $now = Get-Date
    $baseInterval = [double]$script:RunStats.PricingIntervalSec
    $target = if ($script:RunStats.NextPricingAllowedAt -gt $now) { $script:RunStats.NextPricingAllowedAt } else { $now }
    $waitSec = ($target - $now).TotalSeconds
    if ($waitSec -gt 0) {
        Start-Sleep -Milliseconds ([int]([Math]::Ceiling($waitSec * 1000)))
        Add-WaitMetric -Seconds $waitSec
    }

    $candidateNext = (Get-Date).AddSeconds($baseInterval + [double]$script:RunStats.PricingCooldownSec)
    if ($script:RunStats.NextPricingAllowedAt -lt $candidateNext) {
        $script:RunStats.NextPricingAllowedAt = $candidateNext
    }
}

function Wait-ForCatalogSlot {
    param([hashtable]$Config)
    if (-not $script:RunStats) { return }

    $now = Get-Date
    $baseInterval = if ($Config.CatalogMinIntervalSec) { [double]$Config.CatalogMinIntervalSec } else { 1.2 }
    $target = if ($script:RunStats.NextCatalogAllowedAt -gt $now) { $script:RunStats.NextCatalogAllowedAt } else { $now }
    $waitSec = ($target - $now).TotalSeconds
    if ($waitSec -gt 0) {
        Start-Sleep -Milliseconds ([int]([Math]::Ceiling($waitSec * 1000)))
        Add-WaitMetric -Seconds $waitSec
    }

    $script:RunStats.NextCatalogAllowedAt = (Get-Date).AddSeconds($baseInterval)
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

    $defaultInterval = if ($Config.PricingDefaultIntervalSec) { [double]$Config.PricingDefaultIntervalSec } else { 12.0 }
    $minFloor = if ($Config.PricingMinIntervalSec) { [double]$Config.PricingMinIntervalSec } else { 0.4 }
    $safetyFactor = if ($Config.PricingSafetyFactor) { [double]$Config.PricingSafetyFactor } else { 1.2 }
    $jitterMaxMs = if ($Config.PricingJitterMaxMs) { [int]$Config.PricingJitterMaxMs } else { 500 }

    $baseInterval = [Math]::Max($minFloor, $defaultInterval)
    $limit = $RateLimitLimit -as [double]
    if ($limit -and $limit -gt 0) {
        $jitterSec = 0.0
        if ($jitterMaxMs -gt 0) {
            $jitterSec = (Get-Random -Minimum 0 -Maximum ($jitterMaxMs + 1)) / 1000.0
        }
        $baseInterval = [Math]::Max($minFloor, ((1.0 / $limit) * $safetyFactor) + $jitterSec)
    }

    $script:RunStats.PricingIntervalSec = $baseInterval

    $candidateNext = (Get-Date).AddSeconds($baseInterval + [double]$script:RunStats.PricingCooldownSec)
    if ($script:RunStats.NextPricingAllowedAt -lt $candidateNext) {
        $script:RunStats.NextPricingAllowedAt = $candidateNext
    }
}


function Write-SpApiRequestResponseTraceLog {
    param(
        [string]$Endpoint,
        [string]$Method,
        [string]$Uri,
        [object]$RequestHeaders,
        [object]$RequestBody,
        [object]$ResponseHeaders,
        [object]$ResponseBody,
        [hashtable]$Config,
        [string]$LogPath,
        [string]$TraceKind
    )

    if (-not $Config.DebugSpApiResponse) {
        return
    }

    $maxChars = if ($Config.DebugSpApiResponseMaxChars) { [int]$Config.DebugSpApiResponseMaxChars } else { 4000 }
    $maxChars = [Math]::Max(200, $maxChars)

    $convertToTraceText = {
        param([object]$Value)

        if ($null -eq $Value) { return '<null>' }
        if ($Value -is [string]) { return [string]$Value }

        if ($Value -is [System.Collections.Specialized.NameValueCollection]) {
            $headerMap = @{}
            foreach ($key in @($Value.AllKeys)) {
                if ($null -eq $key) { continue }
                $headerMap[[string]$key] = @($Value.GetValues($key)) -join ','
            }
            return (($headerMap | ConvertTo-Json -Depth 5 -Compress) 2>$null)
        }

        if ($Value -is [System.Collections.IDictionary]) {
            return (($Value | ConvertTo-Json -Depth 20 -Compress) 2>$null)
        }

        try {
            return (($Value | ConvertTo-Json -Depth 20 -Compress) 2>$null)
        }
        catch {
            try {
                return [string]$Value
            }
            catch {
                return '<unserializable>'
            }
        }
    }

    $truncateAndMask = {
        param([string]$Text)

        $masked = Mask-SensitiveText -Text $Text
        if ([string]::IsNullOrWhiteSpace($masked)) { return '<empty>' }
        if ($masked.Length -gt $maxChars) {
            return "$($masked.Substring(0, $maxChars))...(truncated)"
        }

        return $masked
    }

    $requestHeadersText = & $truncateAndMask (& $convertToTraceText $RequestHeaders)
    $requestBodyText = & $truncateAndMask (& $convertToTraceText $RequestBody)
    $responseHeadersText = & $truncateAndMask (& $convertToTraceText $ResponseHeaders)
    $responseBodyText = & $truncateAndMask (& $convertToTraceText $ResponseBody)

    Write-Log -Message "$Endpoint trace[$TraceKind](max=$maxChars): method=$Method uri=$Uri request.headers=$requestHeadersText request.body=$requestBodyText response.headers=$responseHeadersText response.body=$responseBodyText" -LogPath $LogPath
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
            if ($Endpoint -match '^PricingBatch') { $script:RunStats.PricingBatchCalls++ }
            if ($Endpoint -match '^CatalogBatch') { $script:RunStats.CatalogBatchCalls++ }

            $params = @{ Method = $Method; Uri = $Uri; Headers = $Headers }
            if ($null -ne $Body -and "$Body" -ne '') { $params.Body = $Body }
            if (($Method -eq 'Post' -or $Method -eq 'Put') -and $params.ContainsKey('Body')) {
                $params.ContentType = 'application/json'
            }

            $irmCommand = Get-Command -Name 'Invoke-RestMethod' -ErrorAction Stop
            if ($irmCommand.Parameters.ContainsKey('ResponseHeadersVariable')) {
                $res = Invoke-RestMethod @params -ResponseHeadersVariable responseHeaders
            }
            else {
                $iwrParams = @{ Method = $Method; Uri = $Uri; Headers = $Headers; UseBasicParsing = $true }
                if ($params.ContainsKey('Body')) { $iwrParams.Body = $params.Body }
                if ($params.ContainsKey('ContentType')) { $iwrParams.ContentType = $params.ContentType }

                $web = Invoke-WebRequest @iwrParams
                $responseHeaders = $web.Headers

                $rawContent = if ($null -ne $web.Content) { [string]$web.Content } else { '' }
                if (-not [string]::IsNullOrWhiteSpace($rawContent)) {
                    try {
                        $res = $rawContent | ConvertFrom-Json -Depth 20
                    }
                    catch {
                        $res = $rawContent
                    }
                }
                else {
                    $res = $null
                }
            }

            $res = ConvertFrom-JsonIfNeeded -Value $res

            $limit = Get-HeaderValue -Headers $responseHeaders -Name 'x-amzn-RateLimit-Limit'
            $requestId = Get-HeaderValue -Headers $responseHeaders -Name 'x-amzn-RequestId'
            if ($limit) {
                Write-Log -Message "$Endpoint success: limit=$limit, requestId=$requestId" -LogPath $LogPath
            }
            if ($Endpoint -match '^Pricing') {
                Update-PricingThrottleFromLimit -RateLimitLimit $limit -Config $Config
            }
            try {
                Write-SpApiRequestResponseTraceLog -Endpoint $Endpoint -Method $Method -Uri $Uri -RequestHeaders $Headers -RequestBody $Body -ResponseHeaders $responseHeaders -ResponseBody $res -Config $Config -LogPath $LogPath -TraceKind 'success'
                Write-SpApiResponseDebugLog -Endpoint $Endpoint -Response $res -Config $Config -LogPath $LogPath
            }
            catch {
                Write-Log -Message "$Endpoint debug-hook failed: $($_.Exception.GetType().FullName) - $($_.Exception.Message)" -LogPath $LogPath -Level 'WARN'
                Write-SpApiResponseShapeLog -Endpoint $Endpoint -Response $res -LogPath $LogPath
                if ($_.ScriptStackTrace) {
                    Write-Log -Message "$Endpoint debug-hook stack: $($_.ScriptStackTrace)" -LogPath $LogPath -Level 'WARN'
                }
            }
            return $res
        }
        catch {
            $detail = Get-ErrorDetail -ErrorRecord $_
            $errorResponseHeaders = $null
            try {
                $hasResponse = $_.Exception | Get-Member -Name 'Response' -MemberType 'Property' -ErrorAction SilentlyContinue
                if ($hasResponse -and $_.Exception.Response) {
                    $errorResponseHeaders = $_.Exception.Response.Headers
                }
            }
            catch {}
            try {
                Write-SpApiRequestResponseTraceLog -Endpoint $Endpoint -Method $Method -Uri $Uri -RequestHeaders $Headers -RequestBody $Body -ResponseHeaders $errorResponseHeaders -ResponseBody $detail.BodyText -Config $Config -LogPath $LogPath -TraceKind 'failure'
            }
            catch {
                Write-Log -Message "$Endpoint trace-hook failed: $($_.Exception.GetType().FullName) - $($_.Exception.Message)" -LogPath $LogPath -Level 'WARN'
            }
            $status = Get-StatusCodeValue -Status $detail.StatusCode
            if ($null -eq $status) { $status = 0 }
            if ($status -eq 0) {
                # log the raw exception for troubleshooting
                Write-Log -Message "Invoke-SpApiRequest exception: $($_.Exception.GetType().FullName) - $($_.Exception.Message)" -LogPath $LogPath -Level 'WARN'
                Write-Log -Message "Invoke-SpApiRequest context: endpoint=$Endpoint method=$Method uri=$Uri attempt=$attempt/$maxAttempts" -LogPath $LogPath -Level 'WARN'
                if ($_.InvocationInfo -and $_.InvocationInfo.PositionMessage) {
                    Write-Log -Message "Invoke-SpApiRequest location: $($_.InvocationInfo.PositionMessage)" -LogPath $LogPath -Level 'WARN'
                }
                if ($_.ScriptStackTrace) {
                    Write-Log -Message "Invoke-SpApiRequest stack: $($_.ScriptStackTrace)" -LogPath $LogPath -Level 'WARN'
                }
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
            $endpointKind = if ($Endpoint -match '^Catalog') { 'Catalog' } elseif ($Endpoint -match '^Pricing') { 'Pricing' } else { 'Other' }
            Write-Log -Message "$Endpoint retry: kind=$endpointKind status=$status class=$($detail.Class) code=$errorCode requestId=$requestId wait=$([Math]::Round($sleepSec,2))s limit=$($detail.RateLimitLimit) attempt=$attempt/$maxAttempts$bodyMsg" -LogPath $LogPath -Level 'WARN'
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
            if ($ErrorRecord.ErrorDetails -and $ErrorRecord.ErrorDetails.Message) {
                $bodyText = [string]$ErrorRecord.ErrorDetails.Message
            }
        }
        catch {}

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

        if ([string]::IsNullOrWhiteSpace($bodyText)) {
            try {
                $hasResponse = $ErrorRecord.Exception | Get-Member -Name 'Response' -MemberType 'Property' -ErrorAction SilentlyContinue
                if ($hasResponse) {
                    $stream = $ErrorRecord.Exception.Response.GetResponseStream()
                    if ($stream -and $stream.CanRead) {
                        $reader = New-Object System.IO.StreamReader($stream)
                        try {
                            $bodyText = $reader.ReadToEnd()
                        }
                        finally {
                            $reader.Dispose()
                        }
                    }
                }
            }
            catch {}
        }
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

    foreach ($offer in @($Offers)) {
        $landedPrice = Get-PropertyValue -Object $offer -Name 'LandedPrice'
        $landedAmount = Get-PropertyValue -Object $landedPrice -Name 'Amount'
        if ($null -ne $landedAmount) {
            $landed = [decimal]$landedAmount
            if ($null -eq $landedMin -or $landed -lt $landedMin) {
                $landedMin = $landed
            }
            continue
        }

        $listingPrice = Get-PropertyValue -Object $offer -Name 'ListingPrice'
        $shipping = Get-PropertyValue -Object $offer -Name 'Shipping'
        $listingAmount = Get-PropertyValue -Object $listingPrice -Name 'Amount'
        $shippingAmount = Get-PropertyValue -Object $shipping -Name 'Amount'
        if ($null -ne $listingAmount -and $null -ne $shippingAmount) {
            $listingPlusShip = [decimal]$listingAmount + [decimal]$shippingAmount
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

function Get-CandidateTitleByItem {
    param([object]$Item)

    if (-not $Item) { return $null }

    $summaries = Get-PropertyValue -Object $Item -Name 'summaries'
    foreach ($summary in @(ConvertTo-ObjectArray -Value $summaries)) {
        $title = [string](Get-PropertyValue -Object $summary -Name 'itemName')
        if (-not [string]::IsNullOrWhiteSpace($title)) { return $title }
    }

    $titleDirect = [string](Get-PropertyValue -Object $Item -Name 'title')
    if (-not [string]::IsNullOrWhiteSpace($titleDirect)) { return $titleDirect }

    return $null
}

function Test-MultipackTitleCandidate {
    param([string]$Title)

    if ([string]::IsNullOrWhiteSpace($Title)) { return $false }
    # 寸法(例: 10×20cm)の誤検知を避けるため、x/×数量は pack 指標付きのみ multipack 扱いにする。
    return ($Title -match '(\d+)(個|入り|入|パック|本|枚|セット)' -or $Title -match '(?:×|x|X)\s*\d{1,3}\s*(個|入|入り|セット|pack|pcs|本|枚|袋)')
}

function Select-BestAsinForJan {
    param(
        [array]$CandidateAsins,
        [hashtable]$PriceByAsin,
        [hashtable]$TitleByAsin,
        [hashtable]$Config
    )

    if (-not $CandidateAsins -or -not $PriceByAsin) {
        return [PSCustomObject]@{ Asin = $null; Price = $null; EffectivePrice = $null }
    }

    $avoidMultipack = $false
    if ($null -ne $Config -and $Config.ContainsKey('AvoidMultipackByTitle')) {
        $avoidMultipack = [bool]$Config.AvoidMultipackByTitle
    }

    $multipackPenalty = 999999
    if ($null -ne $Config -and $Config.ContainsKey('MultipackTitlePenalty')) {
        $multipackPenalty = [decimal]$Config.MultipackTitlePenalty
    }

    $bestAsin = $null
    $bestPrice = $null
    $bestEffectivePrice = $null

    foreach ($asin in @($CandidateAsins)) {
        if ([string]::IsNullOrWhiteSpace([string]$asin)) { continue }
        if (-not $PriceByAsin.ContainsKey($asin)) { continue }

        $basePrice = $PriceByAsin[$asin]
        if ($null -eq $basePrice -or "$basePrice" -eq '') { continue }

        $effectivePrice = [decimal]$basePrice
        if ($avoidMultipack -and $TitleByAsin -and $TitleByAsin.ContainsKey($asin)) {
            if (Test-MultipackTitleCandidate -Title ([string]$TitleByAsin[$asin])) {
                $effectivePrice = $effectivePrice + $multipackPenalty
            }
        }

        if ($null -eq $bestEffectivePrice -or $effectivePrice -lt $bestEffectivePrice -or ($effectivePrice -eq $bestEffectivePrice -and [string]$asin -lt [string]$bestAsin)) {
            $bestAsin = [string]$asin
            $bestPrice = [decimal]$basePrice
            $bestEffectivePrice = $effectivePrice
        }
    }

    return [PSCustomObject]@{ Asin = $bestAsin; Price = $bestPrice; EffectivePrice = $bestEffectivePrice }
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
    $errorReasonMap = @{}
    $candidateAsinsMap = @{}
    $candidateTitleMap = @{}
    $candidateMaxAsinsPerJan = if ($Config.CandidateMaxAsinsPerJan) { [int]$Config.CandidateMaxAsinsPerJan } else { 5 }
    if ($candidateMaxAsinsPerJan -lt 1) { $candidateMaxAsinsPerJan = 1 }
    $minBatchSize = 5
    $batchSize = [Math]::Max($minBatchSize, [int]$Config.CatalogBatchSize)
    $index = 0

    $applyCatalogItems = {
        param(
            [object]$Items,
            [hashtable]$TargetMap,
            [hashtable]$TargetErrorClassMap,
            [hashtable]$TargetJanLookupMap,
            [hashtable]$ParseStats,
            [hashtable]$CandidateAsinsMap,
            [hashtable]$CandidateTitleMap,
            [int]$CandidateMaxAsinsPerJan
        )

        if (-not $Items) { return }

        if ($ParseStats) {
            if (-not $ParseStats.ContainsKey('ExpandedItems')) { $ParseStats.ExpandedItems = 0 }
            if (-not $ParseStats.ContainsKey('WithoutIdentifiers')) { $ParseStats.WithoutIdentifiers = 0 }
            if (-not $ParseStats.ContainsKey('IdentifierCandidates')) { $ParseStats.IdentifierCandidates = 0 }
            if (-not $ParseStats.ContainsKey('MatchedIdentifiers')) { $ParseStats.MatchedIdentifiers = 0 }
            if (-not $ParseStats.ContainsKey('ItemsWithAsin')) { $ParseStats.ItemsWithAsin = 0 }
        }

        foreach ($item in (Expand-CatalogItems -Items $Items)) {
            if ($ParseStats) { $ParseStats.ExpandedItems++ }
            $itemIdentifiers = Get-PropertyValue -Object $item -Name 'identifiers'
            if (-not $itemIdentifiers) {
                if ($ParseStats) { $ParseStats.WithoutIdentifiers++ }
                continue
            }

            $identifierGroups = @()
            $nestedGroups = Get-PropertyValue -Object $itemIdentifiers -Name 'identifiers'
            if ($nestedGroups) {
                $identifierGroups = ConvertTo-ObjectArray -Value $nestedGroups
            }
            else {
                $identifierGroups = ConvertTo-ObjectArray -Value $itemIdentifiers
            }

            if (@($identifierGroups).Count -eq 0) {
                Write-Log -Message "Catalog identifier parse: empty identifierGroups (item.identifiers.type=$(Get-ObjectTypeName -Value $itemIdentifiers), nested.type=$(Get-ObjectTypeName -Value $nestedGroups))" -LogPath $LogPath -Level 'WARN'
                continue
            }

            $matchedIdentifier = $null
            foreach ($idGroup in @($identifierGroups)) {
                # identifiers 配下に identifiers 配列がネストされるケースと、
                # 直接 identifierType/identifier を持つケースの両方に対応する。
                $leafIdentifiers = @()
                $nestedLeafIdentifiers = Get-PropertyValue -Object $idGroup -Name 'identifiers'
                if ($nestedLeafIdentifiers) {
                    $leafIdentifiers = ConvertTo-ObjectArray -Value $nestedLeafIdentifiers
                }
                else {
                    $leafIdentifiers = ConvertTo-ObjectArray -Value $idGroup
                }

                foreach ($leaf in @($leafIdentifiers)) {
                    $identifierType = [string](Get-PropertyValue -Object $leaf -Name 'identifierType')
                    $identifierValue = [string](Get-PropertyValue -Object $leaf -Name 'identifier')
                    if ([string]::IsNullOrWhiteSpace($identifierValue)) {
                        $identifierValue = [string](Get-PropertyValue -Object $leaf -Name 'value')
                    }
                    if (($identifierType -in @('JAN', 'EAN', 'UPC', 'GTIN')) -and -not [string]::IsNullOrWhiteSpace($identifierValue)) {
                        if ($ParseStats) { $ParseStats.IdentifierCandidates++ }
                        $matchedOriginalJan = Find-TargetJanByIdentifier -Identifier $identifierValue -JanLookupMap $TargetJanLookupMap
                        if ($matchedOriginalJan) {
                            if ($ParseStats) { $ParseStats.MatchedIdentifiers++ }
                            $matchedIdentifier = $matchedOriginalJan
                            break
                        }
                    }
                }
                if ($matchedIdentifier) { break }
            }

            $asin = [string](Get-PropertyValue -Object $item -Name 'asin')
            if ($matchedIdentifier -and -not [string]::IsNullOrWhiteSpace($asin)) {
                if ($ParseStats) { $ParseStats.ItemsWithAsin++ }
                if (-not $CandidateAsinsMap.ContainsKey($matchedIdentifier)) {
                    $CandidateAsinsMap[$matchedIdentifier] = New-Object System.Collections.Generic.HashSet[string]
                }
                $candidateSet = $CandidateAsinsMap[$matchedIdentifier]
                if ($candidateSet.Count -lt $CandidateMaxAsinsPerJan) {
                    [void]$candidateSet.Add($asin)
                }

                if (-not $CandidateTitleMap.ContainsKey($matchedIdentifier)) {
                    $CandidateTitleMap[$matchedIdentifier] = @{}
                }
                $titleByAsin = $CandidateTitleMap[$matchedIdentifier]
                if (-not $titleByAsin.ContainsKey($asin)) {
                    $titleByAsin[$asin] = Get-CandidateTitleByItem -Item $item
                }

                if (-not $TargetMap[$matchedIdentifier]) {
                    $TargetMap[$matchedIdentifier] = $asin
                }
                $TargetErrorClassMap.Remove($matchedIdentifier) | Out-Null
                $errorReasonMap.Remove($matchedIdentifier) | Out-Null
            }
        }

        return
    }

    while ($index -lt $Jans.Count) {
        $end = [Math]::Min($index + $batchSize - 1, $Jans.Count - 1)
        $chunk = @($Jans[$index..$end])
        $chunkJanLookupMap = @{}
        $chunkParseStats = @{
            ExpandedItems = 0
            WithoutIdentifiers = 0
            IdentifierCandidates = 0
            MatchedIdentifiers = 0
            ItemsWithAsin = 0
            Unresolved = 0
        }
        foreach ($jan in $chunk) {
            $resultMap[$jan] = $null
            $candidateAsinsMap[$jan] = New-Object System.Collections.Generic.HashSet[string]
            if (-not $candidateTitleMap.ContainsKey($jan)) { $candidateTitleMap[$jan] = @{} }
            if (-not $errorClassMap.ContainsKey($jan)) { $errorClassMap[$jan] = $null }
            if (-not $errorReasonMap.ContainsKey($jan)) { $errorReasonMap[$jan] = $null }
            foreach ($lookupKey in (Get-IdentifierMatchKeys -Identifier ([string]$jan))) {
                if (-not $chunkJanLookupMap.ContainsKey($lookupKey)) {
                    $chunkJanLookupMap[$lookupKey] = [string]$jan
                }
            }
        }

        $identifiers = ($chunk | ForEach-Object { $_.Trim() }) -join ','
        $catalogPageSize = [Math]::Min($batchSize, 20)
        $baseUri = "$($Config.SpApiBaseUrl)/catalog/2022-04-01/items?identifiers=$([Uri]::EscapeDataString($identifiers))&identifiersType=JAN&marketplaceIds=$($Config.MarketplaceId)&includedData=identifiers,summaries&pageSize=$catalogPageSize"
        Write-Log -Message "JAN検索: $($chunk.Count)件 (index=$index,size=$batchSize,pageSize=$catalogPageSize)" -LogPath $LogPath

        $res = $null
        $attemptDetail = $null
        $pageResponses = @()
        $pageNumber = 1
        $maxPages = 10
        $seenNextTokens = @{}
        $nextToken = $null

        while ($true) {
            $requestUri = $baseUri
            if (-not [string]::IsNullOrWhiteSpace($nextToken)) {
                $requestUri = "$baseUri&pageToken=$([Uri]::EscapeDataString($nextToken))"
            }

            $res = $null
            $attemptDetail = $null
            try {
                Wait-ForCatalogSlot -Config $Config
                $res = Invoke-SpApiRequest -Endpoint "CatalogBatch(index=$index,size=$batchSize,page=$pageNumber)" -Method 'Get' -Uri $requestUri -Headers (New-SpApiHeaders -AccessToken $AccessToken -Config $Config) -Config $Config -LogPath $LogPath
            }
            catch {
                $attemptDetail = Get-ErrorDetail -ErrorRecord $_
                if ($attemptDetail.Class -eq 'Auth' -and $AuthContext) {
                    try {
                        $AccessToken = Get-LwaAccessTokenCached -ClientId $AuthContext.ClientId -ClientSecret $AuthContext.ClientSecret -RefreshToken $AuthContext.RefreshToken -Config $Config -LogPath $LogPath -TokenCachePath $AuthContext.TokenCachePath -ForceRefresh
                        Wait-ForCatalogSlot -Config $Config
                        $res = Invoke-SpApiRequest -Endpoint "CatalogBatchAuthRetry(index=$index,size=$batchSize,page=$pageNumber)" -Method 'Get' -Uri $requestUri -Headers (New-SpApiHeaders -AccessToken $AccessToken -Config $Config) -Config $Config -LogPath $LogPath
                    }
                    catch {
                        $attemptDetail = Get-ErrorDetail -ErrorRecord $_
                    }
                }
            }

            if (-not $res) {
                break
            }

            $pageResponses += @($res)
            $catalogItemsPage = Get-PropertyValue -Object $res -Name 'items'
            $expandedItemsPageCount = @(Expand-CatalogItems -Items $catalogItemsPage).Count
            $pagination = Get-PropertyValue -Object $res -Name 'pagination'
            $nextTokenCandidate = [string](Get-PropertyValue -Object $pagination -Name 'nextToken')
            if ([string]::IsNullOrWhiteSpace($nextTokenCandidate)) {
                $nextTokenCandidate = [string](Get-PropertyValue -Object $pagination -Name 'NextToken')
            }
            $nextTokenPreview = if ([string]::IsNullOrWhiteSpace($nextTokenCandidate)) { '<none>' } else { $nextTokenCandidate.Substring(0, [Math]::Min(16, $nextTokenCandidate.Length)) }
            Write-Log -Message "Catalog page fetch: index=$index page=$pageNumber items=$expandedItemsPageCount nextToken=$nextTokenPreview" -LogPath $LogPath

            if ([string]::IsNullOrWhiteSpace($nextTokenCandidate)) {
                break
            }
            if ($pageNumber -ge $maxPages) {
                Write-Log -Message "Catalog page fetch reached maxPages=$maxPages (index=$index)" -LogPath $LogPath -Level 'WARN'
                break
            }
            if ($seenNextTokens.ContainsKey($nextTokenCandidate)) {
                Write-Log -Message "Catalog page fetch detected duplicated nextToken (index=$index,page=$pageNumber)" -LogPath $LogPath -Level 'WARN'
                break
            }

            $seenNextTokens[$nextTokenCandidate] = $true
            $nextToken = $nextTokenCandidate
            $pageNumber++
        }

        if (@($pageResponses).Count -eq 0) {
            if ($attemptDetail -and $attemptDetail.Class -eq 'RateLimit/Server' -and $batchSize -gt $minBatchSize) {
                $nextBatchSize = [Math]::Max($minBatchSize, [int][Math]::Floor($batchSize / 2))
                Write-Log -Message "Catalogバッチを縮小します: $batchSize -> $nextBatchSize (index=$index, HTTP=$($attemptDetail.StatusCode), limit=$($attemptDetail.RateLimitLimit))" -LogPath $LogPath -Level 'WARN'
                $batchSize = $nextBatchSize
                continue
            }

            foreach ($jan in $chunk) {
                $errorClassMap[$jan] = if ($attemptDetail) { $attemptDetail.Class } else { 'Other' }
                $errorReasonMap[$jan] = 'ApiError'
            }
            Write-Log -Message "Catalog unresolved reason stats: index=$index unresolvedByReason=ApiError=$($chunk.Count)" -LogPath $LogPath -Level 'WARN'
            $index = $end + 1
            continue
        }

        $catalogItems = @()
        foreach ($pageRes in @($pageResponses)) {
            $catalogItemsPage = Get-PropertyValue -Object $pageRes -Name 'items'
            $catalogItems += @(Expand-CatalogItems -Items $catalogItemsPage)
            & $applyCatalogItems -Items $catalogItemsPage -TargetMap $resultMap -TargetErrorClassMap $errorClassMap -TargetJanLookupMap $chunkJanLookupMap -ParseStats $chunkParseStats -CandidateAsinsMap $candidateAsinsMap -CandidateTitleMap $candidateTitleMap -CandidateMaxAsinsPerJan $candidateMaxAsinsPerJan | Out-Null
        }

        $unresolvedJans = @($chunk | Where-Object { -not $resultMap[$_] })
        $chunkParseStats.Unresolved = @($unresolvedJans).Count
        if (@($unresolvedJans).Count -gt 0) {
            $resolvedJans = @($chunk | Where-Object { $resultMap[$_] })
            Write-Log -Message "Catalog unresolved JAN detail: index=$index unresolved.count=$($unresolvedJans.Count) unresolved.list=$(Format-PreviewList -Items $unresolvedJans -MaxCount 30)" -LogPath $LogPath -Level 'WARN'
            if (@($resolvedJans).Count -gt 0) {
                Write-Log -Message "Catalog resolved JAN detail: index=$index resolved.count=$($resolvedJans.Count) resolved.list=$(Format-PreviewList -Items $resolvedJans -MaxCount 20)" -LogPath $LogPath
            }
        }
        if ($chunkParseStats.Count -gt 0) {
            Write-Log -Message "Catalog parse stats: index=$index size=$($chunk.Count) expanded=$($chunkParseStats.ExpandedItems) withoutIdentifiers=$($chunkParseStats.WithoutIdentifiers) identifierCandidates=$($chunkParseStats.IdentifierCandidates) matchedIdentifiers=$($chunkParseStats.MatchedIdentifiers) itemsWithAsin=$($chunkParseStats.ItemsWithAsin) unresolved=$($chunkParseStats.Unresolved)" -LogPath $LogPath
            if (($chunkParseStats.ExpandedItems -gt 0) -and ($chunkParseStats.WithoutIdentifiers -eq $chunkParseStats.ExpandedItems)) {
                $propertySample = Get-CatalogItemPropertySample -Items $catalogItems
                Write-Log -Message "Catalog parse diagnostic: expanded items have no identifiers (index=$index, item.props=$propertySample)" -LogPath $LogPath -Level 'WARN'
            }
        }
        if ($unresolvedJans.Count -eq $chunk.Count) {
            $sampleText = Get-CatalogIdentifierSampleText -Items $catalogItems
            $numberOfResults = Get-PropertyValue -Object $res -Name 'numberOfResults'
            $expandedCount = @(Expand-CatalogItems -Items $catalogItems).Count
            Write-Log -Message "Catalog parse diagnostic: all unresolved in chunk (index=$index,size=$($chunk.Count), response.items.type=$(Get-ObjectTypeName -Value $catalogItems), expanded.items.count=$expandedCount, numberOfResults=$numberOfResults, response.props=$((@($res.PSObject.Properties.Name) -join ',')), sample.identifiers=$sampleText)" -LogPath $LogPath -Level 'WARN'
        }
        Write-Log -Message "EANフォールバックは無効化されています: unresolved=$($unresolvedJans.Count)件 (index=$index)" -LogPath $LogPath

        $chunkUnresolvedReason = 'IdentifierMismatch'
        if ($chunkParseStats.ExpandedItems -eq 0) {
            $chunkUnresolvedReason = 'CatalogNoItems'
        }
        elseif (($chunkParseStats.ExpandedItems -gt 0) -and ($chunkParseStats.WithoutIdentifiers -eq $chunkParseStats.ExpandedItems)) {
            $chunkUnresolvedReason = 'CatalogNoIdentifiers'
        }
        elseif (($chunkParseStats.IdentifierCandidates -eq 0) -and ($chunkParseStats.ExpandedItems -gt 0)) {
            $chunkUnresolvedReason = 'IdentifierMismatch'
        }

        foreach ($jan in $unresolvedJans) {
            if (-not $errorReasonMap.ContainsKey($jan)) {
                $errorReasonMap[$jan] = $chunkUnresolvedReason
            }
        }

        if (@($unresolvedJans).Count -gt 0) {
            $reasonCountMap = @{}
            foreach ($jan in $unresolvedJans) {
                $reason = [string]$errorReasonMap[$jan]
                if ([string]::IsNullOrWhiteSpace($reason)) { $reason = 'Unknown' }
                if (-not $reasonCountMap.ContainsKey($reason)) { $reasonCountMap[$reason] = 0 }
                $reasonCountMap[$reason]++
            }
            $reasonText = (@($reasonCountMap.Keys | Sort-Object | ForEach-Object { "$_=$($reasonCountMap[$_])" })) -join ','
            Write-Log -Message "Catalog unresolved reason stats: index=$index unresolvedByReason=$reasonText" -LogPath $LogPath -Level 'WARN'
        }

        foreach ($jan in $chunk) {
            if (-not $resultMap[$jan] -and -not $errorClassMap.ContainsKey($jan)) {
                $errorClassMap[$jan] = 'NotFound/Validation'
            }
        }

        $index = $end + 1
    }

    $candidateAsinsMapResult = @{}
    foreach ($jan in $resultMap.Keys) {
        $candidateList = @($candidateAsinsMap[$jan])
        $candidateAsinsMapResult[$jan] = @($candidateList | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) } | Sort-Object -Unique)
    }

    [PSCustomObject]@{ AsinMap = $resultMap; CandidateAsinsMap = $candidateAsinsMapResult; CandidateTitleMap = $candidateTitleMap; ErrorClassMap = $errorClassMap; ErrorReasonMap = $errorReasonMap }
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

    $statusCode = Get-StatusCodeValue -Status (Get-PropertyValue -Object $res -Name 'status')
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
        $skippedBlankAsin = 0
        foreach ($asin in $chunk) {
            if ([string]::IsNullOrWhiteSpace([string]$asin)) {
                $skippedBlankAsin++
                continue
            }

            $normalizedAsin = ([string]$asin).Trim()
            $requests += @{
                Asin          = $normalizedAsin
                MarketplaceId = [string]$Config.MarketplaceId
                ItemCondition = 'New'
                method        = 'GET'
                uri           = "/products/pricing/v0/items/$([Uri]::EscapeDataString($normalizedAsin))/offers"
            }
        }

        if ($skippedBlankAsin -gt 0) {
            Write-Log -Message "Pricing: skip blank ASIN count=$skippedBlankAsin (index=$index,size=$batchSize)" -LogPath $LogPath -Level 'WARN'
        }

        if (@($requests).Count -eq 0) {
            foreach ($asin in $chunk) { $errorClassMap[$asin] = 'NotFound/Validation' }
            $index = $end + 1
            continue
        }

        $body = @{ requests = $requests } | ConvertTo-Json -Depth 10 -Compress

        $res = $null
        $attemptDetail = $null
        try {
            Wait-ForPricingSlot -Config $Config
            $res = Invoke-SpApiRequest -Endpoint "PricingBatch(index=$index,size=$batchSize)" -Method 'Post' -Uri "$($Config.SpApiBaseUrl)/batches/products/pricing/v0/itemOffers" -Headers (New-SpApiHeaders -AccessToken $AccessToken -Config $Config -ContentType 'application/json') -Body $body -Config $Config -LogPath $LogPath
        }
        catch {
            $attemptDetail = Get-ErrorDetail -ErrorRecord $_
            if ($attemptDetail.Class -eq 'Auth' -and $AuthContext) {
                try {
                    $AccessToken = Get-LwaAccessTokenCached -ClientId $AuthContext.ClientId -ClientSecret $AuthContext.ClientSecret -RefreshToken $AuthContext.RefreshToken -Config $Config -LogPath $LogPath -TokenCachePath $AuthContext.TokenCachePath -ForceRefresh
                    Wait-ForPricingSlot -Config $Config
                    $res = Invoke-SpApiRequest -Endpoint "PricingBatchAuthRetry(index=$index,size=$batchSize)" -Method 'Post' -Uri "$($Config.SpApiBaseUrl)/batches/products/pricing/v0/itemOffers" -Headers (New-SpApiHeaders -AccessToken $AccessToken -Config $Config -ContentType 'application/json') -Body $body -Config $Config -LogPath $LogPath
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

        $responseItems = Get-PropertyValue -Object $res -Name 'responses'
        if (-not $responseItems) {
            foreach ($asin in $chunk) { $errorClassMap[$asin] = 'Other' }
            $index = $end + 1
            continue
        }

        $retryableResponseAsins = New-Object System.Collections.Generic.List[string]
        foreach ($response in @($responseItems)) {
            $statusRaw = Get-PropertyValue -Object $response -Name 'status'
            $statusCode = Get-StatusCodeValue -Status $statusRaw

            $asin = $null
            if ($response.body -and $response.body.payload -and $response.body.payload.ASIN) {
                $asin = [string]$response.body.payload.ASIN
            }
            elseif ($response.request -and $response.request.uri) {
                if ($response.request.uri -match '/items/([^/]+)/offers') {
                    $asin = [Uri]::UnescapeDataString($matches[1])
                }
            }
            elseif ($response.request -and $response.request.Asin) {
                $asin = [string]$response.request.Asin
            }

            if (-not $asin -or -not $priceMap.ContainsKey($asin)) { continue }

            if ($statusCode -ge 400) {
                $bodyText = if ($response.body) { ($response.body | ConvertTo-Json -Depth 8) } else { '' }
                $detail = Get-StatusClassification -StatusCode $statusCode -BodyText $bodyText
                $errorClassMap[$asin] = $detail.Class
                if ($statusCode -eq 429 -or $statusCode -eq 500 -or $statusCode -eq 503) {
                    $retryableResponseAsins.Add($asin) | Out-Null
                }
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

        if ($retryableResponseAsins.Count -gt 0) {
            $retryAsins = @($retryableResponseAsins | Sort-Object -Unique)
            Write-Log -Message "Pricing部分失敗ASINを単発再試行します: count=$($retryAsins.Count)" -LogPath $LogPath -Level 'WARN'
            foreach ($asin in $retryAsins) {
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
        }

        $index = $end + 1
    }

    [PSCustomObject]@{ PriceMap = $priceMap; ErrorClassMap = $errorClassMap }
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
        
        # Set headers for G, H columns (I列は空欄維持)
        $sheet.Cells.Item(1, 7).Value2 = 'ASIN'
        $sheet.Cells.Item(1, 8).Value2 = '価格'
        $sheet.Cells.Item(1, 9).Value2 = ''
        
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
        if ($needApiJans.Count -gt 0) {
            $catalogResult = Get-AsinMapByJanBatch -Jans $needApiJans -AccessToken $accessToken -Config $Config -LogPath $logPath -AuthContext $authContext
            $asinMap = if ($catalogResult.AsinMap -is [hashtable]) { $catalogResult.AsinMap } else { @{} }
            $candidateAsinsMap = if ($catalogResult.CandidateAsinsMap -is [hashtable]) { $catalogResult.CandidateAsinsMap } else { @{} }
            $candidateTitleMap = if ($catalogResult.CandidateTitleMap -is [hashtable]) { $catalogResult.CandidateTitleMap } else { @{} }
            $catalogErrorMap = if ($catalogResult.ErrorClassMap -is [hashtable]) { $catalogResult.ErrorClassMap } else { @{} }
            $catalogApiCalls = $script:RunStats.CatalogBatchCalls

            $priceMap = @{}
            $priceErrorMap = @{}
            $allCandidateAsins = @()
            foreach ($jan in $needApiJans) {
                if ($candidateAsinsMap.ContainsKey($jan)) {
                    $candidates = @($candidateAsinsMap[$jan])
                    $allCandidateAsins += $candidates
                }
                elseif ($asinMap[$jan]) {
                    $allCandidateAsins += @([string]$asinMap[$jan])
                }
            }
            $allAsins = @($allCandidateAsins | Where-Object { $_ } | Sort-Object -Unique)
            $uniqueAsinCount = $allAsins.Count

            $needPriceAsins = @()
            foreach ($asin in $allAsins) {
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

            if ($allAsins.Count -eq 0) {
                # Guard for the known crash path when all Catalog responses are empty.
                Write-Log -Message 'No candidate ASINs found; skipping pricing.' -LogPath $logPath
            }
            if ($needPriceAsins.Count -gt 0) {
                $distinctAsins = @($needPriceAsins | Sort-Object -Unique)
                $pricingResult = Get-PriceMapByAsinBatch -Asins $distinctAsins -AccessToken $accessToken -Config $Config -LogPath $logPath -AuthContext $authContext
                foreach ($k in $pricingResult.PriceMap.Keys) { $priceMap[$k] = $pricingResult.PriceMap[$k] }
                foreach ($k in $pricingResult.ErrorClassMap.Keys) { $priceErrorMap[$k] = $pricingResult.ErrorClassMap[$k] }
                $pricingApiCalls = $script:RunStats.PricingBatchCalls
            }

            $fetchedAt = (Get-Date).ToString('o')
            foreach ($jan in $needApiJans) {
                $cacheStatus = 'ok'
                $asin = $null
                $price = $null
                $candidateAsins = if ($candidateAsinsMap.ContainsKey($jan)) { @($candidateAsinsMap[$jan]) } elseif ($asinMap[$jan]) { @([string]$asinMap[$jan]) } else { @() }
                $candidateAsins = @($candidateAsins)
                $titleByAsin = if ($candidateTitleMap.ContainsKey($jan)) { $candidateTitleMap[$jan] } else { @{} }

                if ($candidateAsins.Count -eq 0) {
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
                    $selection = Select-BestAsinForJan -CandidateAsins $candidateAsins -PriceByAsin $priceMap -TitleByAsin $titleByAsin -Config $Config
                    $asin = $selection.Asin
                    $price = $selection.Price
                    if ($Config.DebugSpApiResponse) { Write-Log -Message "JAN選定: jan=$jan candidate_count=$($candidateAsins.Count) chosen_asin=$asin chosen_price=$price" -LogPath $logPath }
                    if (-not $asin) {
                        $cacheStatus = 'not_found'; $notFoundValidationCount++
                    }
                    elseif ($priceErrorMap.ContainsKey($asin)) {
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

            $sheet.Cells.Item($row, 9).Value2 = ''

            if ([string]::IsNullOrWhiteSpace($jan)) {
                $sheet.Cells.Item($row, 7).Value2 = ''
                $sheet.Cells.Item($row, 8).Value2 = ''
                continue
            }

            try {
                $result = $runCache[$jan]
                if ($result -and ($result.cache_status -eq 'not_found' -or $result.cache_status -eq 'transient_error')) {
                    $sheet.Cells.Item($row, 7).Value2 = ''
                    $sheet.Cells.Item($row, 8).Value2 = ''
                    if ($result.cache_status -eq 'transient_error') {
                        Write-Log -Message "行$row JAN=$jan は一時エラーのため空欄出力します。" -LogPath $logPath -Level 'WARN'
                    }
                    continue
                }

                $sheet.Cells.Item($row, 7).Value2 = if ($result -and $result.asin) { $result.asin } else { '' }
                if ($result -and $null -ne $result.price -and "$($result.price)" -ne '') {
                    $sheet.Cells.Item($row, 8).Value2 = [double]$result.price
                }
                else {
                    $sheet.Cells.Item($row, 8).Value2 = ''
                }
            }
            catch {
                $detail = Get-ErrorDetail -ErrorRecord $_
                if ($detail.Class -eq 'NotFound/Validation') { $notFoundValidationCount++ }
                elseif ($detail.Class -eq 'RateLimit/Server') { $rateLimitServerCount++; $transientErrorCount++ }
                else { $otherErrorCount++; $transientErrorCount++ }

                $sheet.Cells.Item($row, 7).Value2 = ''
                $sheet.Cells.Item($row, 8).Value2 = ''
                Write-Log -Message "行$row JAN=$jan の処理でエラー: 分類=$($detail.Class), HTTP=$($detail.StatusCode), msg=$($_.Exception.Message)" -LogPath $logPath -Level 'ERROR'
            }

            $processed++
        }

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

Export-ModuleMember -Function ConvertTo-PlainText,Write-Log,Get-StatusClassification,Get-ErrorDetail,Invoke-WithRetry,Get-AmzDateHeaderValue,New-SpApiHeaders,Split-IntoChunks,Read-AccessTokenCache,Save-AccessTokenCache,Get-LwaAccessTokenCached,Get-LwaAccessToken,Get-LowestNewPriceFromOffers,Get-PriceBySingleAsin,Get-AsinMapByJanBatch,Get-PriceMapByAsinBatch,Import-PersistentCache,Save-PersistentCache,Add-DailyPriceHistory,Test-CacheFresh,Save-SecretsInteractive,Invoke-AmazonPriceUpdate
