# load JAN list from data\input.xlsx (B column)
$inputPath = Join-Path $PSScriptRoot 'data\input.xlsx'
if (-not (Test-Path $inputPath)) {
    throw "input.xlsx が見つかりません: $inputPath"
}

$excel = $null
$wb = $null
$sh = $null
$sample = New-Object System.Collections.Generic.List[string]
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    # ReadOnly=true で開き、input.xlsx を変更・上書きしない
    $wb = $excel.Workbooks.Open($inputPath, 0, $true)
    $sh = $wb.Worksheets.Item(1)

    $lastRow = $sh.Cells($sh.Rows.Count, 2).End(-4162).Row # xlUp
    for ($row = 2; $row -le $lastRow; $row++) {
        $jan = [string]$sh.Cells.Item($row, 2).Value2
        if (-not [string]::IsNullOrWhiteSpace($jan)) {
            [void]$sample.Add($jan.Trim())
        }
    }
} finally {
    if ($wb) { $wb.Close($false) }
    if ($excel) { $excel.Quit() }
    if ($sh) { [Runtime.InteropServices.Marshal]::ReleaseComObject($sh) | Out-Null }
    if ($wb) { [Runtime.InteropServices.Marshal]::ReleaseComObject($wb) | Out-Null }
    if ($excel) { [Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null }
}

if ($sample.Count -eq 0) {
    throw 'input.xlsx のB列(JAN)に有効な値がありません。'
}

$sample = $sample | Select-Object -Unique
Write-Host "Loaded JAN count: $($sample.Count)"

# start actual debug operations
$Config = Import-PowerShellDataFile -Path 'config.psd1'
$libPath = Join-Path $PSScriptRoot 'scripts/lib/AmazonPriceLib.psm1'
Import-Module $libPath -Force
$secret = Import-Clixml -Path $Config.Paths.SecretsFile
$clientId = $secret.client_id
$clientSecret = ConvertTo-PlainText -Secure $secret.client_secret
$refresh = ConvertTo-PlainText -Secure $secret.refresh_token
$token = Get-LwaAccessTokenCached -ClientId $clientId -ClientSecret $clientSecret -RefreshToken $refresh -Config $Config -LogPath 'logs/run.log' -TokenCachePath (Join-Path $Config.Paths.CacheDir 'access_token.json')
Write-Host "Token length: $($token.Length)"
$authContext = @{ClientId=$clientId; ClientSecret=$clientSecret; RefreshToken=$refresh; TokenCachePath=(Join-Path $Config.Paths.CacheDir 'access_token.json')}
try {
    $r = Get-AsinMapByJanBatch -Jans $sample -AccessToken $token -Config $Config -LogPath 'logs/run.log' -AuthContext $authContext
    Write-Host 'Result:'
    $r
} catch {
    Write-Host 'Manual call failed:' $_.Exception.Message
    if ($_.Exception.Response) { Write-Host 'Response status:' $_.Exception.Response.StatusCode }
    Write-Host 'Exception dump:' $_.Exception | Format-List * -Force
}
