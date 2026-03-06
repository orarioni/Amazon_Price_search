param(
    [string]$InputPath = (Join-Path $PSScriptRoot 'data\input.xlsx'),
    [int]$MaxJanCount = 20
)

$Config = Import-PowerShellDataFile -Path 'config.psd1'
$libPath = Join-Path $PSScriptRoot 'scripts/lib/AmazonPriceLib.psm1'
Import-Module $libPath -Force

if (-not (Test-Path -LiteralPath $InputPath)) {
    throw "input.xlsx not found: $InputPath"
}

$excel = $null
$wb = $null
$sh = $null
$sample = @()

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $wb = $excel.Workbooks.Open($InputPath, 0, $true)
    $sh = $wb.Worksheets.Item(1)

    $lastRow = $sh.Cells($sh.Rows.Count, 2).End(-4162).Row
    $janSet = New-Object System.Collections.Generic.HashSet[string]

    for ($row = 2; $row -le $lastRow; $row++) {
        $jan = ([string]$sh.Cells.Item($row, 2).Text).Trim()
        if ([string]::IsNullOrWhiteSpace($jan)) { continue }
        [void]$janSet.Add($jan)
        if ($janSet.Count -ge $MaxJanCount) { break }
    }

    $sample = @($janSet)
}
finally {
    if ($null -ne $wb) { $wb.Close($false) }
    if ($null -ne $excel) { $excel.Quit() }
    if ($null -ne $sh) { [Runtime.InteropServices.Marshal]::ReleaseComObject($sh) | Out-Null }
    if ($null -ne $wb) { [Runtime.InteropServices.Marshal]::ReleaseComObject($wb) | Out-Null }
    if ($null -ne $excel) { [Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null }
}

if (@($sample).Count -eq 0) {
    throw "No valid JAN found in column B of input.xlsx: $InputPath"
}

Write-Host "Loaded JAN count: $(@($sample).Count) (max=$MaxJanCount, readOnly=true)"
Write-Host "JAN preview: $((@($sample) | Select-Object -First 10) -join ',')"

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
}
catch {
    Write-Host 'Manual call failed:' $_.Exception.Message
    if ($_.Exception.Response) { Write-Host 'Response status:' $_.Exception.Response.StatusCode }
    Write-Host 'Exception dump:' $_.Exception | Format-List * -Force
}
