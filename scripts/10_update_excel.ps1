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

$awsRegion = 'us-east-1'
$awsService = 'execute-api'
$spApiHost = 'sellingpartnerapi-fe.amazon.com'

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
    if ($null -eq $Secure) { return $null }

    $ptr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($Secure)
    try {
        [Runtime.InteropServices.Marshal]::PtrToStringBSTR($ptr)
    }
    finally {
        [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($ptr)
    }
}

function Get-HexSha256 {
    param([string]$Text)
    $sha = [System.Security.Cryptography.SHA256]::Create()
    try {
        $bytes = [System.Text.Encoding]::UTF8.GetBytes($Text)
        $hash = $sha.ComputeHash($bytes)
        return ([System.BitConverter]::ToString($hash) -replace '-', '').ToLowerInvariant()
    }
    finally {
        $sha.Dispose()
    }
}

function Get-HmacSha256 {
    param(
        [byte[]]$Key,
        [string]$Data
    )

    $hmac = [System.Security.Cryptography.HMACSHA256]::new($Key)
    try {
        $bytes = [System.Text.Encoding]::UTF8.GetBytes($Data)
        return $hmac.ComputeHash($bytes)
    }
    finally {
        $hmac.Dispose()
    }
}

function Get-SignatureKey {
    param(
        [string]$SecretKey,
        [string]$DateStamp,
        [string]$Region,
        [string]$Service
    )

    $kDate = Get-HmacSha256 -Key ([System.Text.Encoding]::UTF8.GetBytes("AWS4$SecretKey")) -Data $DateStamp
    $kRegion = Get-HmacSha256 -Key $kDate -Data $Region
    $kService = Get-HmacSha256 -Key $kRegion -Data $Service
    $kSigning = Get-HmacSha256 -Key $kService -Data 'aws4_request'
    return $kSigning
}

function Get-Rfc3986Encoded {
    param([string]$Value)
    return [Uri]::EscapeDataString([string]$Value).Replace('+', '%20').Replace('*', '%2A').Replace('%7E', '~')
}

function ConvertTo-CanonicalQueryString {
    param([string]$Query)
    if ([string]::IsNullOrEmpty($Query)) { return '' }

    $pairs = @()
    foreach ($part in $Query -split '&') {
        if ($part -eq '') { continue }
        $kv = $part -split '=', 2
        $k = if ($kv.Count -gt 0) { $kv[0] } else { '' }
        $v = if ($kv.Count -gt 1) { $kv[1] } else { '' }
        $pairs += [PSCustomObject]@{ k = (Get-Rfc3986Encoded $k); v = (Get-Rfc3986Encoded $v) }
    }

    $sorted = $pairs | Sort-Object k, v
    return (($sorted | ForEach-Object { "{0}={1}" -f $_.k, $_.v }) -join '&')
}

function New-SpApiHeaders {
    param(
        [string]$Method,
        [string]$Uri,
        [string]$AccessToken,
        [string]$AwsAccessKeyId,
        [string]$AwsSecretAccessKey,
        [string]$AwsSessionToken
    )

    $parsed = [Uri]$Uri
    $headers = [ordered]@{
        'x-amz-access-token' = $AccessToken
        'User-Agent'         = $userAgent
        'Accept'             = 'application/json'
    }

    if ([string]::IsNullOrWhiteSpace($AwsAccessKeyId) -or [string]::IsNullOrWhiteSpace($AwsSecretAccessKey)) {
        $headers['Authorization'] = "Bearer $AccessToken"
        return $headers
    }

    $amzDate = (Get-Date).ToUniversalTime().ToString('yyyyMMddTHHmmssZ')
    $dateStamp = (Get-Date).ToUniversalTime().ToString('yyyyMMdd')
    $canonicalUri = if ([string]::IsNullOrWhiteSpace($parsed.AbsolutePath)) { '/' } else { $parsed.AbsolutePath }
    $canonicalQuery = ConvertTo-CanonicalQueryString -Query $parsed.Query.TrimStart('?')

    $headers['host'] = $parsed.Host
    $headers['x-amz-date'] = $amzDate
    if (-not [string]::IsNullOrWhiteSpace($AwsSessionToken)) {
        $headers['x-amz-security-token'] = $AwsSessionToken
    }

    $signedHeaders = ($headers.Keys | ForEach-Object { $_.ToLowerInvariant() } | Sort-Object) -join ';'
    $canonicalHeaders = ($headers.Keys | ForEach-Object {
        $key = $_.ToLowerInvariant()
        $value = [string]$headers[$_]
        "{0}:{1}`n" -f $key, (($value -replace '\s+', ' ').Trim())
    } | Sort-Object) -join ''

    $payloadHash = Get-HexSha256 -Text ''
    $canonicalRequest = @(
        $Method.ToUpperInvariant(),
        $canonicalUri,
        $canonicalQuery,
        $canonicalHeaders,
        $signedHeaders,
        $payloadHash
    ) -join "`n"

    $algorithm = 'AWS4-HMAC-SHA256'
    $credentialScope = "$dateStamp/$awsRegion/$awsService/aws4_request"
    $stringToSign = @(
        $algorithm,
        $amzDate,
        $credentialScope,
        (Get-HexSha256 -Text $canonicalRequest)
    ) -join "`n"

    $signingKey = Get-SignatureKey -SecretKey $AwsSecretAccessKey -DateStamp $dateStamp -Region $awsRegion -Service $awsService
    $signatureBytes = Get-HmacSha256 -Key $signingKey -Data $stringToSign
    $signature = ([System.BitConverter]::ToString($signatureBytes) -replace '-', '').ToLowerInvariant()

    $headers['Authorization'] = "$algorithm Credential=$AwsAccessKeyId/$credentialScope, SignedHeaders=$signedHeaders, Signature=$signature"

    return $headers
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
        Invoke-RestMethod -Method Post -Uri 'https://api.amazon.com/auth/o2/token' -ContentType 'application/x-www-form-urlencoded' -Body $body -Headers @{ 'User-Agent' = $userAgent }
    }

    if (-not $res.access_token) {
        throw 'LWAアクセストークンの取得に失敗しました。'
    }
    return $res.access_token
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
        $headers = New-SpApiHeaders -Method 'GET' -Uri $uri -AccessToken $AccessToken -AwsAccessKeyId $AwsAccessKeyId -AwsSecretAccessKey $AwsSecretAccessKey -AwsSessionToken $AwsSessionToken
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
        $headers = New-SpApiHeaders -Method 'GET' -Uri $uri -AccessToken $AccessToken -AwsAccessKeyId $AwsAccessKeyId -AwsSecretAccessKey $AwsSecretAccessKey -AwsSessionToken $AwsSessionToken
        Invoke-RestMethod -Method Get -Uri $uri -Headers $headers
    }

    if (-not $res.payload -or -not $res.payload.Offers) {
        return $null
    }

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
    Write-Log 'AWS認証情報が未登録のため、BearerモードでSP-APIを呼び出します。401/403が出る場合は run_init.bat でAWS認証情報を登録してください。' 'WARN'
}
else {
    Write-Log 'AWS SigV4 署名付きでSP-APIを呼び出します。'
}

Write-Log '更新処理を開始します。'
$accessToken = Get-LwaAccessToken -ClientId $clientId -ClientSecret $clientSecret -RefreshToken $refreshToken

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
