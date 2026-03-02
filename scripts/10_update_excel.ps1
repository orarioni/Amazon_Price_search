$ErrorActionPreference = 'Stop'

$repoRoot = Resolve-Path (Join-Path $PSScriptRoot '..')
$secretFile = Join-Path $repoRoot 'secrets/lwa_secrets.xml'
$dataDir = Join-Path $repoRoot 'data'
$inputPath = Join-Path $dataDir 'input.xlsx'
$outputPath = Join-Path $dataDir 'output.xlsx'
$logDir = Join-Path $repoRoot 'logs'
$logPath = Join-Path $logDir 'run.log'
$cacheDir = Join-Path $repoRoot 'cache'
$cachePath = Join-Path $cacheDir 'price_cache.json'

$marketplaceId = 'A1VC38T7YXB528'
$spBase = 'https://sellingpartnerapi-fe.amazon.com'
$userAgent = 'AmazonPriceTool/0.3'
$maxRetries = 4
$catalogBatchSize = 20
$pricingBatchSize = 20
$cacheTtlHours = 24

if (-not (Test-Path $logDir)) {
    New-Item -ItemType Directory -Path $logDir -Force | Out-Null
}

if (-not (Test-Path $cacheDir)) {
    New-Item -ItemType Directory -Path $cacheDir -Force | Out-Null
}

function Write-Log {
    param(
        [string]$Message,
        [string]$Level = 'INFO'
    )
    $line = "$(Get-Date -Format o) [$Level] $Message"
    Add-Content -Path $logPath -Value $line
    Write-Host $line
}

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

function Classify-StatusAndBody {
    param(
        [Nullable[int]]$StatusCode,
        [string]$BodyText
    )

    $text = if ($BodyText) { $BodyText.ToLowerInvariant() } else { '' }

    if ($StatusCode -eq 429 -or ($StatusCode -ge 500 -and $StatusCode -lt 600)) {
        return [PSCustomObject]@{ Class = 'RateLimit/Server'; IsTransient = $true; IsPermanentNotFound = $false }
    }

    if ($StatusCode -eq 404) {
        return [PSCustomObject]@{ Class = 'NotFound/Validation'; IsTransient = $false; IsPermanentNotFound = $true }
    }

    if ($StatusCode -eq 400 -or $StatusCode -eq 422) {
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

    if ($ErrorRecord -and $ErrorRecord.Exception -and $ErrorRecord.Exception.Response) {
        try {
            if ($ErrorRecord.Exception.Response.StatusCode) {
                $statusCode = [int]$ErrorRecord.Exception.Response.StatusCode
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
        Class               = $classification.Class
        IsTransient         = $classification.IsTransient
        IsPermanentNotFound = $classification.IsPermanentNotFound
    }
}

function Invoke-WithRetry {
    param(
        [scriptblock]$Action,
        [string]$Label
    )

    $attempt = 0
    while ($true) {
        $attempt++
        try {
            return & $Action
        }
        catch {
            $detail = Get-ErrorDetail -ErrorRecord $_

            if ($detail.IsTransient -and $attempt -lt $maxRetries) {
                $sleepSec = [Math]::Pow(2, $attempt)
                Write-Log "$Label 失敗 (分類=$($detail.Class), HTTP $($detail.StatusCode))。$sleepSec 秒後にリトライします (試行 $attempt/$maxRetries)。" 'WARN'
                Start-Sleep -Seconds $sleepSec
                continue
            }

            Write-Log "$Label 失敗 (分類=$($detail.Class), HTTP $($detail.StatusCode))。再試行を終了します。" 'WARN'
            throw
        }
    }
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

function Get-LwaAccessToken {
    param(
        [string]$ClientId,
        [string]$ClientSecret,
        [string]$RefreshToken
    )

    $body = @{
        grant_type    = 'refresh_token'
        refresh_token = $RefreshToken
        client_id     = $ClientId
        client_secret = $ClientSecret
    }

    $res = Invoke-WithRetry -Label 'LWAトークン取得' -Action {
        Invoke-RestMethod -Method Post -Uri 'https://api.amazon.com/auth/o2/token' -ContentType 'application/x-www-form-urlencoded' -Body $body -Headers @{
            'User-Agent' = $userAgent
        }
    }

    if (-not $res.access_token) {
        throw 'LWAアクセストークンの取得に失敗しました。'
    }
    return $res.access_token
}

function Get-LowestNewPriceFromOffers {
    param(
        [array]$Offers
    )

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
        [string]$AccessToken
    )

    $resultMap = @{}
    $errorClassMap = @{}
    $chunks = Split-IntoChunks -Items $Jans -ChunkSize $catalogBatchSize

    for ($i = 0; $i -lt $chunks.Count; $i++) {
        $chunk = $chunks[$i]
        foreach ($jan in $chunk) {
            $resultMap[$jan] = $null
        }

        $identifiers = ($chunk | ForEach-Object { $_.Trim() }) -join ','
        $uri = "$spBase/catalog/2022-04-01/items?identifiers=$([Uri]::EscapeDataString($identifiers))&identifiersType=EAN&marketplaceIds=$marketplaceId"

        try {
            $res = Invoke-WithRetry -Label "Catalogバッチ取得 ($($i + 1)/$($chunks.Count))" -Action {
                Invoke-RestMethod -Method Get -Uri $uri -Headers @{
                    'Authorization'      = "Bearer $AccessToken"
                    'x-amz-access-token' = $AccessToken
                    'User-Agent'         = $userAgent
                    'Accept'             = 'application/json'
                }
            }
        }
        catch {
            $detail = Get-ErrorDetail -ErrorRecord $_
            foreach ($jan in $chunk) {
                $errorClassMap[$jan] = $detail.Class
            }
            continue
        }

        if ($res.items) {
            foreach ($item in $res.items) {
                if (-not $item.identifiers -or -not $item.identifiers.identifiers) {
                    continue
                }

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
    }

    return [PSCustomObject]@{
        AsinMap = $resultMap
        ErrorClassMap = $errorClassMap
    }
}

function Get-PriceMapByAsinBatch {
    param(
        [array]$Asins,
        [string]$AccessToken
    )

    $priceMap = @{}
    $errorClassMap = @{}
    $chunks = Split-IntoChunks -Items $Asins -ChunkSize $pricingBatchSize

    for ($i = 0; $i -lt $chunks.Count; $i++) {
        $chunk = $chunks[$i]

        foreach ($asin in $chunk) {
            $priceMap[$asin] = $null
        }

        $requests = @()
        foreach ($asin in $chunk) {
            $requests += @{
                uri    = "/products/pricing/v0/items/$([Uri]::EscapeDataString($asin))/offers?MarketplaceId=$marketplaceId&ItemCondition=New"
                method = 'GET'
            }
        }

        $body = @{ requests = $requests } | ConvertTo-Json -Depth 5

        try {
            $res = Invoke-WithRetry -Label "Pricingバッチ取得 ($($i + 1)/$($chunks.Count))" -Action {
                Invoke-RestMethod -Method Post -Uri "$spBase/batches/products/pricing/v0/itemOffers" -Headers @{
                    'Authorization'      = "Bearer $AccessToken"
                    'x-amz-access-token' = $AccessToken
                    'User-Agent'         = $userAgent
                    'Accept'             = 'application/json'
                    'Content-Type'       = 'application/json'
                } -Body $body
            }
        }
        catch {
            $detail = Get-ErrorDetail -ErrorRecord $_
            foreach ($asin in $chunk) {
                $errorClassMap[$asin] = $detail.Class
            }
            continue
        }

        if (-not $res.responses) {
            foreach ($asin in $chunk) {
                $errorClassMap[$asin] = 'Other'
            }
            continue
        }

        foreach ($response in $res.responses) {
            $statusCode = $null
            if ($response.status) {
                $statusCode = [int]$response.status
            }

            $asin = $null
            if ($response.body -and $response.body.payload -and $response.body.payload.ASIN) {
                $asin = [string]$response.body.payload.ASIN
            }
            elseif ($response.request -and $response.request.uri) {
                if ($response.request.uri -match '/items/([^/]+)/offers') {
                    $asin = [Uri]::UnescapeDataString($matches[1])
                }
            }

            if (-not $asin -or -not $priceMap.ContainsKey($asin)) {
                continue
            }

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
    }

    return [PSCustomObject]@{
        PriceMap = $priceMap
        ErrorClassMap = $errorClassMap
    }
}

function Load-PersistentCache {
    param([string]$Path)

    if (-not (Test-Path $Path)) {
        return @{}
    }

    try {
        $raw = Get-Content -Path $Path -Raw -Encoding UTF8
        if ([string]::IsNullOrWhiteSpace($raw)) {
            return @{}
        }

        $parsed = ConvertFrom-Json -InputObject $raw
        $map = @{}
        if ($parsed) {
            foreach ($item in $parsed) {
                if (-not $item.jan) {
                    continue
                }
                $map[[string]$item.jan] = [PSCustomObject]@{
                    asin        = $item.asin
                    price       = $item.price
                    fetched_at  = $item.fetched_at
                    cache_status = if ($item.cache_status) { $item.cache_status } else { 'ok' }
                }
            }
        }

        return $map
    }
    catch {
        Write-Log "キャッシュ読込に失敗したため空キャッシュで継続します: $($_.Exception.Message)" 'WARN'
        return @{}
    }
}

function Save-PersistentCache {
    param(
        [hashtable]$CacheMap,
        [string]$Path
    )

    $rows = @()
    foreach ($key in $CacheMap.Keys) {
        $rows += [PSCustomObject]@{
            jan         = $key
            asin        = $CacheMap[$key].asin
            price       = $CacheMap[$key].price
            fetched_at  = $CacheMap[$key].fetched_at
            cache_status = $CacheMap[$key].cache_status
        }
    }

    $rows | Sort-Object jan | ConvertTo-Json -Depth 5 | Set-Content -Path $Path -Encoding UTF8
}

function Is-CacheFresh {
    param(
        [object]$Entry,
        [int]$TtlHours
    )

    if (-not $Entry -or -not $Entry.fetched_at) {
        return $false
    }

    $fetchedAt = $null
    if (-not [DateTime]::TryParse($Entry.fetched_at, [ref]$fetchedAt)) {
        return $false
    }

    return ((Get-Date) - $fetchedAt).TotalHours -lt $TtlHours
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

Write-Log '更新処理を開始します。'
$accessToken = Get-LwaAccessToken -ClientId $clientId -ClientSecret $clientSecret -RefreshToken $refreshToken

$excel = $null
$workbook = $null

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    $workbook = $excel.Workbooks.Open($inputPath)
    $sheet = $workbook.Worksheets.Item(1)
    $lastRow = $sheet.Cells($sheet.Rows.Count, 2).End(-4162).Row

    $persistentCache = Load-PersistentCache -Path $cachePath
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

        if (-not [string]::IsNullOrWhiteSpace($jan)) {
            [void]$targetJans.Add($jan)
        }
    }

    $janList = @($targetJans)
    $needApiJans = @()

    foreach ($jan in $janList) {
        if ($persistentCache.ContainsKey($jan) -and (Is-CacheFresh -Entry $persistentCache[$jan] -TtlHours $cacheTtlHours)) {
            $runCache[$jan] = $persistentCache[$jan]
            $cacheHitCount++
        }
        else {
            $needApiJans += $jan
            $cacheMissCount++
        }
    }

    if ($needApiJans.Count -gt 0) {
        $catalogResult = Get-AsinMapByJanBatch -Jans $needApiJans -AccessToken $accessToken
        $asinMap = $catalogResult.AsinMap
        $catalogErrorMap = $catalogResult.ErrorClassMap
        $catalogApiCalls = (Split-IntoChunks -Items $needApiJans -ChunkSize $catalogBatchSize).Count

        $needPriceAsins = @()
        foreach ($jan in $needApiJans) {
            $asin = $asinMap[$jan]
            if ($asin) {
                $needPriceAsins += $asin
            }
        }

        $priceMap = @{}
        $priceErrorMap = @{}
        if ($needPriceAsins.Count -gt 0) {
            $distinctAsins = @($needPriceAsins | Sort-Object -Unique)
            $pricingResult = Get-PriceMapByAsinBatch -Asins $distinctAsins -AccessToken $accessToken
            $priceMap = $pricingResult.PriceMap
            $priceErrorMap = $pricingResult.ErrorClassMap
            $pricingApiCalls = (Split-IntoChunks -Items $distinctAsins -ChunkSize $pricingBatchSize).Count
        }

        $fetchedAt = (Get-Date).ToString('o')
        foreach ($jan in $needApiJans) {
            $cacheStatus = 'ok'
            $asin = $asinMap[$jan]
            $price = $null

            if (-not $asin) {
                if ($catalogErrorMap.ContainsKey($jan) -and $catalogErrorMap[$jan] -eq 'NotFound/Validation') {
                    $cacheStatus = 'not_found'
                    $notFoundValidationCount++
                }
                elseif ($catalogErrorMap.ContainsKey($jan) -and $catalogErrorMap[$jan] -eq 'RateLimit/Server') {
                    $rateLimitServerCount++
                    $errorCount++
                }
                elseif ($catalogErrorMap.ContainsKey($jan)) {
                    $otherErrorCount++
                    $errorCount++
                }
                else {
                    $cacheStatus = 'not_found'
                    $notFoundValidationCount++
                }
            }
            else {
                if ($priceMap.ContainsKey($asin)) {
                    $price = $priceMap[$asin]
                }

                if ($priceErrorMap.ContainsKey($asin)) {
                    $errClass = $priceErrorMap[$asin]
                    if ($errClass -eq 'NotFound/Validation') {
                        $cacheStatus = 'not_found'
                        $notFoundValidationCount++
                    }
                    elseif ($errClass -eq 'RateLimit/Server') {
                        $rateLimitServerCount++
                        $errorCount++
                    }
                    else {
                        $otherErrorCount++
                        $errorCount++
                    }
                }
            }

            $entry = [PSCustomObject]@{
                asin         = $asin
                price        = $price
                fetched_at   = $fetchedAt
                cache_status = $cacheStatus
            }

            $runCache[$jan] = $entry
            if ($cacheStatus -eq 'not_found' -or $cacheStatus -eq 'ok') {
                $persistentCache[$jan] = $entry
            }
        }
    }

    for ($row = 2; $row -le $lastRow; $row++) {
        $timestamp = (Get-Date).ToString('o')
        $jan = $janByRow[$row]
        $sheet.Cells.Item($row, 5).Value2 = $timestamp

        if ([string]::IsNullOrWhiteSpace($jan)) {
            $sheet.Cells.Item($row, 3).Value2 = ''
            $sheet.Cells.Item($row, 4).Value2 = ''
            continue
        }

        try {
            $result = $runCache[$jan]

            if ($result -and $result.cache_status -eq 'not_found') {
                $sheet.Cells.Item($row, 3).Value2 = ''
                $sheet.Cells.Item($row, 4).Value2 = ''
                continue
            }

            if ($result -and $result.asin) {
                $sheet.Cells.Item($row, 3).Value2 = $result.asin
            }
            else {
                $sheet.Cells.Item($row, 3).Value2 = ''
            }

            if ($result -and $null -ne $result.price -and "$($result.price)" -ne '') {
                $sheet.Cells.Item($row, 4).Value2 = [double]$result.price
            }
            else {
                $sheet.Cells.Item($row, 4).Value2 = ''
            }
        }
        catch {
            $detail = Get-ErrorDetail -ErrorRecord $_
            $errorCount++
            if ($detail.Class -eq 'NotFound/Validation') {
                $notFoundValidationCount++
            }
            elseif ($detail.Class -eq 'RateLimit/Server') {
                $rateLimitServerCount++
            }
            else {
                $otherErrorCount++
            }

            $sheet.Cells.Item($row, 3).Value2 = ''
            $sheet.Cells.Item($row, 4).Value2 = ''
            Write-Log "行$row JAN=$jan の処理でエラー: 分類=$($detail.Class), HTTP=$($detail.StatusCode), msg=$($_.Exception.Message)" 'ERROR'
        }

        $processed++
    }

    Save-PersistentCache -CacheMap $persistentCache -Path $cachePath

    try {
        $workbook.SaveAs($outputPath)
    }
    catch {
        Write-Host 'output.xlsx を保存できませんでした。Excelを閉じてから再実行してください。'
        throw
    }

    Write-Log "呼び出し統計: JAN総数=$($janList.Count), cache_hit=$cacheHitCount, cache_miss=$cacheMissCount, catalog_calls=$catalogApiCalls, pricing_calls=$pricingApiCalls"
    Write-Log "エラー分類統計: NotFound/Validation=$notFoundValidationCount, RateLimit/Server=$rateLimitServerCount, Other=$otherErrorCount"
    Write-Log "更新完了: 処理件数=$processed, エラー件数=$errorCount, 出力=$outputPath"
}
finally {
    if ($workbook) { $workbook.Close($false) }
    if ($excel) { $excel.Quit() }
    if ($sheet) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($sheet) }
    if ($workbook) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) }
    if ($excel) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel) }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
