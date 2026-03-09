function Import-ConfigDataFile {
    param([Parameter(Mandatory=$true)][string]$Path)

    if (Get-Command -Name Import-PowerShellDataFile -ErrorAction SilentlyContinue) {
        return Import-PowerShellDataFile -Path $Path
    }

    $raw = Get-Content -LiteralPath $Path -Raw -ErrorAction Stop
    $parsed = Invoke-Expression $raw
    if ($parsed -isnot [hashtable]) {
        throw "Failed to parse config data file as hashtable: $Path"
    }
    return $parsed
}

$ErrorActionPreference = 'Stop'

try {
    $repoRoot = Resolve-Path (Join-Path $PSScriptRoot '..')
    $configPath = Join-Path $repoRoot 'config.psd1'
    $libPath = Join-Path $PSScriptRoot 'lib/AmazonPriceLib.psm1'

    Import-Module $libPath -Force
    $config = Import-ConfigDataFile -Path $configPath
    $secretFile = Join-Path $repoRoot $config.Paths.SecretsFile

    Save-SecretsInteractive -SecretFile $secretFile
    exit 0
}
catch {
    Write-Host ''
    Write-Host '初期設定に失敗しました。' -ForegroundColor Red
    Write-Host "エラー: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host ''
    Write-Host '復旧手順:' -ForegroundColor Yellow
    Write-Host '1) config.psd1 が存在し、破損していないか確認してください。'
    Write-Host '2) client_id / client_secret / refresh_token を再入力して run_init.bat を再実行してください。'
    Write-Host '3) 会社PCの権限で失敗する場合は、PowerShellを通常ユーザーで開き scripts\00_init_secrets.ps1 を実行して原因を確認してください。'
    Write-Host ''
    Write-Host '上記で解決しない場合は、表示されたエラー全文を添えて管理者に連絡してください。'
    exit 1
}
