$ErrorActionPreference = 'Stop'

try {
    $repoRoot = Resolve-Path (Join-Path $PSScriptRoot '..')
    $configPath = Join-Path $repoRoot 'config.psd1'
    $libPath = Join-Path $PSScriptRoot 'lib/AmazonPriceLib.psm1'

    Import-Module $libPath -Force
    $config = Import-PowerShellDataFile -Path $configPath

    Invoke-AmazonPriceUpdate -RepoRoot $repoRoot -Config $config
    exit 0
}
catch {
    Write-Host ''
    Write-Host '更新処理に失敗しました。' -ForegroundColor Red
    Write-Host "エラー: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host ''
    Write-Host '復旧手順:' -ForegroundColor Yellow

    $message = "$($_.Exception.Message)"
    if ($message -match 'invalid_grant') {
        Write-Host '- refresh_token が失効/不正の可能性があります。run_init.bat を再実行して認証情報を再登録してください。'
    }
    elseif ($message -match 'output\.xlsx|保存できません') {
        Write-Host '- Excelで data\output.xlsx が開いている可能性があります。Excelをすべて閉じて再実行してください。'
    }
    elseif ($message -match 'input\.xlsx|入力ファイルが見つかりません') {
        Write-Host '- data\input.xlsx が存在するか確認し、B列にJANが入っていることを確認してください。'
    }
    else {
        Write-Host '- logs\run.log の末尾を確認し、失敗したJANと分類（NotFound/Validation, RateLimit/Server, Other）を確認してください。'
        Write-Host '- ネットワーク不調の可能性があるため、時間を置いて run_update.bat を再実行してください。'
    }

    Write-Host ''
    Write-Host '上記で解決しない場合は、エラー全文と logs\run.log を添えて管理者に連絡してください。'
    exit 1
}
