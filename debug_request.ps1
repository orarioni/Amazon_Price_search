# prepare small input file
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$wb = $excel.Workbooks.Add()
$sh = $wb.Worksheets.Item(1)
$sh.Cells.Item(1,2).Value2 = 'JAN'
$sample = @('4901234567890','4547274043587','4547274044041','4547274044065','4547274044072')
$i=2
foreach($j in $sample){ $sh.Cells.Item($i,2).Value2 = $j; $i++ }
$wb.SaveAs((Join-Path $PSScriptRoot 'data\input.xlsx'))
$wb.Close($false)
$excel.Quit()
[Runtime.InteropServices.Marshal]::ReleaseComObject($sh)|Out-Null
[Runtime.InteropServices.Marshal]::ReleaseComObject($wb)|Out-Null
[Runtime.InteropServices.Marshal]::ReleaseComObject($excel)|Out-Null

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
