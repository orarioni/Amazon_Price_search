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
$awsRegion = 'us-west-2'
$awsService = 'execute-api'
$userAgent = 'AmazonPriceTool/0.1'
$maxRetries = 4

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
        [string]$Label
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

            $retryable = ($statusCode -eq 429) -or ($statusCode -ge 500 -and $statusCode -lt 600)
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

function Get-Sha256Hex {
    param([string]$Text)
    $bytes = [Text.Encoding]::UTF8.GetBytes($Text)
    $hash = [Security.Cryptography.SHA256]::Create().ComputeHash($bytes)
    return ([BitConverter]::ToString($hash)).Replace('-', '').ToLowerInvariant()
}

function Get-HmacSha256 {
    param(
        [byte[]]$Key,
        [string]$Data
    )

    $hmac = [Security.Cryptography.HMACSHA256]::new($Key)
    try {
        $bytes = [Text.Encoding]::UTF8.GetBytes($Data)
        return $hmac.ComputeHash($bytes)
    }
    finally {
        $hmac.Dispose()
    }
}

function ConvertTo-SigV4Encoded {
    param([string]$Value)

    if ($null -eq $Value) {
        return ''
    }

    return [Uri]::EscapeDataString($Value).Replace('+', '%20').Replace('*', '%2A').Replace('%7E', '~')
}

function Get-CanonicalQueryString {
    param([uri]$Uri)

    $query = $Uri.Query.TrimStart('?')
    if ([string]::IsNullOrEmpty($query)) {
        return ''
    }

    $pairs = @()
    foreach ($part in ($query -split '&')) {
        if ($part -eq '') {
            continue
        }

        $kv = $part -split '=', 2
        $keyRaw = [Uri]::UnescapeDataString($kv[0])
        $valueRaw = if ($kv.Count -gt 1) { [Uri]::UnescapeDataString($kv[1]) } else { '' }
        $pairs += [PSCustomObject]@{
            Key = ConvertTo-SigV4Encoded -Value $keyRaw
            Value = ConvertTo-SigV4Encoded -Value $valueRaw
        }
    }

    return (($pairs | Sort-Object Key, Value | ForEach-Object { "{0}={1}" -f $_.Key, $_.Value }) -join '&')
}

function Get-CanonicalUri {
    param([uri]$Uri)

    if ([string]::IsNullOrEmpty($Uri.AbsolutePath) -or $Uri.AbsolutePath -eq '/') {
        return '/'
    }

    $segments = $Uri.AbsolutePath -split '/'
    $encoded = foreach ($segment in $segments) {
        ConvertTo-SigV4Encoded -Value ([Uri]::UnescapeDataString($segment))
    }

    return ($encoded -join '/')
}

function New-SpApiAuthHeaders {
    param(
        [string]$Method,
        [string]$Uri,
        [string]$AccessToken,
        [string]$AwsAccessKeyId,
        [string]$AwsSecretAccessKey,
        [string]$AwsSessionToken
    )

    $requestUri = [uri]$Uri
    $amzDate = (Get-Date).ToUniversalTime().ToString('yyyyMMddTHHmmssZ')
    $dateStamp = (Get-Date).ToUniversalTime().ToString('yyyyMMdd')
    $payloadHash = Get-Sha256Hex -Text ''

    $canonicalUri = Get-CanonicalUri -Uri $requestUri
    $canonicalQueryString = Get-CanonicalQueryString -Uri $requestUri

    $headers = [ordered]@{
        host                 = $requestUri.Host
        'x-amz-access-token' = $AccessToken
        'x-amz-content-sha256' = $payloadHash
        'x-amz-date'         = $amzDate
    }
    if (-not [string]::IsNullOrWhiteSpace($AwsSessionToken)) {
        $headers['x-amz-security-token'] = $AwsSessionToken
    }

    $canonicalHeaders = (($headers.GetEnumerator() | Sort-Object Name | ForEach-Object { "{0}:{1}" -f $_.Name.ToLowerInvariant(), $_.Value.Trim() }) -join "`n") + "`n"
    $signedHeaders = ($headers.Keys | Sort-Object | ForEach-Object { $_.ToLowerInvariant() }) -join ';'

    $canonicalRequest = @(
        $Method.ToUpperInvariant()
        $canonicalUri
        $canonicalQueryString
        $canonicalHeaders
        $signedHeaders
        $payloadHash
    ) -join "`n"

    $credentialScope = "$dateStamp/$awsRegion/$awsService/aws4_request"
    $stringToSign = @(
        'AWS4-HMAC-SHA256'
        $amzDate
        $credentialScope
        (Get-Sha256Hex -Text $canonicalRequest)
    ) -join "`n"

    $kDate = Get-HmacSha256 -Key ([Text.Encoding]::UTF8.GetBytes("AWS4$AwsSecretAccessKey")) -Data $dateStamp
    $kRegion = Get-HmacSha256 -Key $kDate -Data $awsRegion
    $kService = Get-HmacSha256 -Key $kRegion -Data $awsService
    $kSigning = Get-HmacSha256 -Key $kService -Data 'aws4_request'
    $signatureBytes = Get-HmacSha256 -Key $kSigning -Data $stringToSign
    $signature = ([BitConverter]::ToString($signatureBytes)).Replace('-', '').ToLowerInvariant()

    $authorizationHeader = "AWS4-HMAC-SHA256 Credential=$AwsAccessKeyId/$credentialScope, SignedHeaders=$signedHeaders, Signature=$signature"

    $requestHeaders = @{
        'Authorization'      = $authorizationHeader
        'x-amz-access-token' = $AccessToken
        'x-amz-content-sha256' = $payloadHash
        'x-amz-date'         = $amzDate
        'User-Agent'         = $userAgent
        'Accept'             = 'application/json'
    }

    if (-not [string]::IsNullOrWhiteSpace($AwsSessionToken)) {
        $requestHeaders['x-amz-security-token'] = $AwsSessionToken
    }

    return $requestHeaders
}

function Get-AsinByJan {
    param(
        [string]$Jan,
        [string]$AccessToken,
        [string]$AwsAccessKeyId,
        [string]$AwsSecretAccessKey,
        [string]$AwsSessionToken
    )

    $uri = "$spBase/catalog/2022-04-01/items?identifiers=$([Uri]::EscapeDataString($Jan))&identifiersType=EAN&marketplaceIds=$marketplaceId"

    $res = Invoke-WithRetry -Label "Catalog取得 JAN=$Jan" -Action {
        $headers = New-SpApiAuthHeaders -Method 'GET' -Uri $uri -AccessToken $AccessToken -AwsAccessKeyId $AwsAccessKeyId -AwsSecretAccessKey $AwsSecretAccessKey -AwsSessionToken $AwsSessionToken
        Invoke-RestMethod -Method Get -Uri $uri -Headers $headers
    }

    if ($res.items -and $res.items.Count -gt 0) {
        return $res.items[0].asin
    }

    return $null
}

function Get-LowestNewPrice {
    param(
        [string]$Asin,
        [string]$AccessToken,
        [string]$AwsAccessKeyId,
        [string]$AwsSecretAccessKey,
        [string]$AwsSessionToken
    )

    $uri = "$spBase/products/pricing/v0/items/$([Uri]::EscapeDataString($Asin))/offers?MarketplaceId=$marketplaceId&ItemCondition=New"

    $res = Invoke-WithRetry -Label "Pricing取得 ASIN=$Asin" -Action {
        $headers = New-SpApiAuthHeaders -Method 'GET' -Uri $uri -AccessToken $AccessToken -AwsAccessKeyId $AwsAccessKeyId -AwsSecretAccessKey $AwsSecretAccessKey -AwsSessionToken $AwsSessionToken
        Invoke-RestMethod -Method Get -Uri $uri -Headers $headers
    }

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
$awsAccessKeyId = $secret.aws_access_key_id
$awsSecretAccessKey = ConvertTo-PlainText -Secure $secret.aws_secret_access_key
$awsSessionToken = $secret.aws_session_token

if ([string]::IsNullOrWhiteSpace($awsAccessKeyId) -or [string]::IsNullOrWhiteSpace($awsSecretAccessKey)) {
    throw 'AWS認証情報が secrets/lwa_secrets.xml に設定されていません。run_init.bat を再実行して aws_access_key_id / aws_secret_access_key を登録してください。'
}

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
    $processed = 0
    $errorCount = 0

    for ($row = 2; $row -le $lastRow; $row++) {
        $timestamp = (Get-Date).ToString('o')
        $jan = [string]$sheet.Cells.Item($row, 2).Text
        $jan = $jan.Trim()

        $sheet.Cells.Item($row, 5).Value2 = $timestamp

        if ([string]::IsNullOrWhiteSpace($jan)) {
            $sheet.Cells.Item($row, 3).Value2 = ''
            $sheet.Cells.Item($row, 4).Value2 = ''
            continue
        }

        try {
            if ($cache.ContainsKey($jan)) {
                $result = $cache[$jan]
            }
            else {
                $asin = Get-AsinByJan -Jan $jan -AccessToken $accessToken -AwsAccessKeyId $awsAccessKeyId -AwsSecretAccessKey $awsSecretAccessKey -AwsSessionToken $awsSessionToken
                if ($asin) {
                    $price = Get-LowestNewPrice -Asin $asin -AccessToken $accessToken -AwsAccessKeyId $awsAccessKeyId -AwsSecretAccessKey $awsSecretAccessKey -AwsSessionToken $awsSessionToken
                    $result = [PSCustomObject]@{ asin = $asin; price = $price }
                }
                else {
                    $result = [PSCustomObject]@{ asin = $null; price = $null }
                }
                $cache[$jan] = $result
            }

            if ($result.asin) {
                $sheet.Cells.Item($row, 3).Value2 = $result.asin
            }
            else {
                $sheet.Cells.Item($row, 3).Value2 = ''
            }

            if ($null -ne $result.price) {
                $sheet.Cells.Item($row, 4).Value2 = [double]$result.price
            }
            else {
                $sheet.Cells.Item($row, 4).Value2 = ''
            }
        }
        catch {
            $errorCount++
            $sheet.Cells.Item($row, 3).Value2 = ''
            $sheet.Cells.Item($row, 4).Value2 = ''
            Write-Log "行$row JAN=$jan の処理でエラー: $($_.Exception.Message)" 'ERROR'
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
