$ErrorActionPreference = 'Stop'

$repoRoot = Resolve-Path (Join-Path $PSScriptRoot '..')
$secretFile = Join-Path $repoRoot 'secrets/lwa_secrets.xml'
$dataDir = Join-Path $repoRoot 'data'
$inputPath = Join-Path $dataDir 'input.xlsx'
$outputPath = Join-Path $dataDir 'output.xlsx'
$logDir = Join-Path $repoRoot 'logs'
$logPath = Join-Path $logDir 'run.log'

$marketplaceId = 'A1VC38T7YXB528'
$spBase = 'https://sellingpartnerapi-fe.amazon.com'
$userAgent = 'AmazonPriceTool/0.1'
$maxRetries = 4

$asinCol = 7
$priceCol = 8
$timestampCol = 9

if (-not (Test-Path $logDir)) {
    New-Item -ItemType Directory -Path $logDir -Force | Out-Null
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

function Invoke-WithRetry {
    param(
        [scriptblock]$Action,
        [string]$Label,
        [bool]$RetryOnTransient = $true
    )

    $attempt = 0
    while ($true) {
        $attempt++
        try {
            return & $Action
        }
        catch {
            $statusCode = $null
            if ($_.Exception.Response -and $_.Exception.Response.StatusCode) {
                $statusCode = [int]$_.Exception.Response.StatusCode
            }

            $retryable = $RetryOnTransient -and (($statusCode -eq 429) -or ($statusCode -ge 500 -and $statusCode -lt 600))
            if ($retryable -and $attempt -lt $maxRetries) {
                $sleepSec = [Math]::Pow(2, $attempt)
                Write-Log "$Label 失敗 (HTTP $statusCode)。$sleepSec 秒後にリトライします (試行 $attempt/$maxRetries)。" 'WARN'
                Start-Sleep -Seconds $sleepSec
                continue
            }

            throw
        }
    }
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

function New-SpApiHeaders {
    param(
        [string]$Uri,
        [string]$AccessToken
    )

    $requestUri = [uri]$Uri
    $amzDate = (Get-Date).ToUniversalTime().ToString('yyyyMMddTHHmmssZ')

    return @{
        'host'               = $requestUri.Host
        'x-amz-access-token' = $AccessToken
        'x-amz-date'         = $amzDate
        'User-Agent'         = $userAgent
        'Accept'             = 'application/json'
    }
}

function Get-AsinMapByJanBatch {
    param(
        [string]$Jan,
        [string]$AccessToken
    )

    $asinMap = @{}
    if (-not $Jans -or $Jans.Count -eq 0) {
        return $asinMap
    }

    $escapedJans = $Jans | ForEach-Object { [Uri]::EscapeDataString($_) }
    $identifiers = ($escapedJans -join ',')
    $uri = "$spBase/catalog/2022-04-01/items?identifiers=$identifiers&identifiersType=EAN&marketplaceIds=$marketplaceId"

    $res = Invoke-WithRetry -Label "Catalog取得 JAN=$Jan" -Action {
        $headers = New-SpApiHeaders -Uri $uri -AccessToken $AccessToken
        Invoke-RestMethod -Method Get -Uri $uri -Headers $headers
    } -RetryOnTransient $false

    foreach ($item in ($response.items | Where-Object { $_ -and $_.asin })) {
        $itemJans = @()

        if ($item.identifiers) {
            foreach ($idGroup in $item.identifiers) {
                if ($idGroup.identifiers) {
                    foreach ($idObj in $idGroup.identifiers) {
                        if ($idObj.identifier) {
                            $itemJans += [string]$idObj.identifier
                        }
                    }
                }
            }
        }

        if ($itemJans.Count -eq 0 -and $item.externalIds -and $item.externalIds.eans) {
            $itemJans += $item.externalIds.eans
        }

        foreach ($itemJan in $itemJans) {
            $normalizedJan = ([string]$itemJan).Trim()
            if ([string]::IsNullOrWhiteSpace($normalizedJan)) {
                continue
            }

            if ($Jans -contains $normalizedJan -and -not $asinMap.ContainsKey($normalizedJan)) {
                $asinMap[$normalizedJan] = [string]$item.asin
            }
        }
    }

    return $asinMap
}

function Get-LowestNewPrice {
    param(
        [string]$Asin,
        [string]$AccessToken
    )

    $uri = "$spBase/products/pricing/v0/items/$([Uri]::EscapeDataString($Asin))/offers?MarketplaceId=$marketplaceId&ItemCondition=New"

    $res = Invoke-WithRetry -Label "Pricing取得 ASIN=$Asin" -Action {
        $headers = New-SpApiHeaders -Uri $uri -AccessToken $AccessToken
        Invoke-RestMethod -Method Get -Uri $uri -Headers $headers
    } -RetryOnTransient $false

    if (-not $res.payload -or -not $res.payload.Offers) {
        return $null
    }

    # 要件: LandedPrice の最小を優先。LandedPrice が1件も無い場合のみ ListingPrice+Shipping の最小を使う。
    $landedMin = $null
    $fallbackMin = $null

    foreach ($offer in $res.payload.Offers) {
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

    $cache = @{}
    $batchSize = 20
    $processed = 0
    $errorCount = 0

    $janSet = [System.Collections.Generic.HashSet[string]]::new()
    for ($row = 2; $row -le $lastRow; $row++) {
        $jan = ([string]$sheet.Cells.Item($row, 2).Text).Trim()
        if (-not [string]::IsNullOrWhiteSpace($jan)) {
            [void]$janSet.Add($jan)
        }
    }

    $uniqueJans = @($janSet)
    Write-Log "JAN事前収集: 全$($uniqueJans.Count)件"

    if ($uniqueJans.Count -gt 0) {
        $asinByJan = @{}
        for ($offset = 0; $offset -lt $uniqueJans.Count; $offset += $batchSize) {
            $end = [Math]::Min($offset + $batchSize - 1, $uniqueJans.Count - 1)
            $janBatch = @($uniqueJans[$offset..$end])
            $batchMap = Get-AsinMapByJanBatch -Jans $janBatch -AccessToken $accessToken -AwsAccessKeyId $awsAccessKeyId -AwsSecretAccessKey $awsSecretAccessKey -AwsSessionToken $awsSessionToken

            foreach ($entry in $batchMap.GetEnumerator()) {
                $asinByJan[$entry.Key] = $entry.Value
            }

            $resolved = $batchMap.Count
            $unresolved = $janBatch.Count - $resolved
            Write-Log "JANバッチ取得: offset=$offset 件数=$($janBatch.Count) 成功=$resolved 未解決=$unresolved"
        }

        foreach ($jan in $uniqueJans) {
            try {
                $asin = if ($asinByJan.ContainsKey($jan)) { $asinByJan[$jan] } else { $null }
                if ($asin) {
                    $price = Get-LowestNewPrice -Asin $asin -AccessToken $accessToken -AwsAccessKeyId $awsAccessKeyId -AwsSecretAccessKey $awsSecretAccessKey -AwsSessionToken $awsSessionToken
                    $cache[$jan] = [PSCustomObject]@{ asin = $asin; price = $price }
                }
                else {
                    $cache[$jan] = [PSCustomObject]@{ asin = $null; price = $null }
                }
            }
            catch {
                $errorCount++
                $cache[$jan] = [PSCustomObject]@{ asin = $null; price = $null }
                Write-Log "JAN=$jan の事前解決でエラー: $($_.Exception.Message)" 'ERROR'
            }
        }
    }

    for ($row = 2; $row -le $lastRow; $row++) {
        $timestamp = (Get-Date).ToString('o')
        $jan = [string]$sheet.Cells.Item($row, 2).Text
        $jan = $jan.Trim()

        $sheet.Cells.Item($row, $timestampCol).Value2 = $timestamp

        if ([string]::IsNullOrWhiteSpace($jan)) {
            $sheet.Cells.Item($row, $asinCol).Value2 = ''
            $sheet.Cells.Item($row, $priceCol).Value2 = ''
            continue
        }

        try {
            if ($cache.ContainsKey($jan)) {
                $result = $cache[$jan]
            }
            else {
                $asin = Get-AsinByJan -Jan $jan -AccessToken $accessToken
                if ($asin) {
                    $price = Get-LowestNewPrice -Asin $asin -AccessToken $accessToken
                    $result = [PSCustomObject]@{ asin = $asin; price = $price }
                }
                else {
                    $result = [PSCustomObject]@{ asin = $null; price = $null }
                }
                $cache[$jan] = $result
            }

            if ($result.asin) {
                $sheet.Cells.Item($row, $asinCol).Value2 = $result.asin
            }
            else {
                $sheet.Cells.Item($row, $asinCol).Value2 = ''
            }

            if ($null -ne $result.price) {
                $sheet.Cells.Item($row, $priceCol).Value2 = [double]$result.price
            }
            else {
                $sheet.Cells.Item($row, $priceCol).Value2 = ''
            }
        }
        catch {
            $errorCount++
            $sheet.Cells.Item($row, $asinCol).Value2 = ''
            $sheet.Cells.Item($row, $priceCol).Value2 = ''

            $statusCode = $null
            if ($_.Exception.Response -and $_.Exception.Response.StatusCode) {
                $statusCode = [int]$_.Exception.Response.StatusCode
            }

            if ($statusCode) {
                Write-Log "行$row JAN=$jan の処理でエラー (HTTP $statusCode): $($_.Exception.Message)" 'ERROR'
            }
            else {
                Write-Log "行$row JAN=$jan の処理でエラー: $($_.Exception.Message)" 'ERROR'
            }
        }

        $processed++
    }

    try {
        $workbook.SaveAs($outputPath)
    }
    catch {
        Write-Host 'output.xlsx を保存できませんでした。Excelを閉じてから再実行してください。'
        throw
    }

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
