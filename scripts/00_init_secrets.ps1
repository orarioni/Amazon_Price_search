$ErrorActionPreference = 'Stop'

$repoRoot = Resolve-Path (Join-Path $PSScriptRoot '..')
$secretsDir = Join-Path $repoRoot 'secrets'
$secretFile = Join-Path $secretsDir 'lwa_secrets.xml'

if (-not (Test-Path $secretsDir)) {
    New-Item -ItemType Directory -Path $secretsDir -Force | Out-Null
}

Write-Host 'Amazon SP-API の認証情報を入力してください。'
$clientId = Read-Host 'client_id'
$clientSecret = Read-Host 'client_secret' -AsSecureString
$refreshToken = Read-Host 'refresh_token' -AsSecureString

$payload = [PSCustomObject]@{
    client_id     = $clientId
    client_secret = $clientSecret
    refresh_token = $refreshToken
    created_at    = (Get-Date).ToString('o')
}

$payload | Export-Clixml -Path $secretFile

Write-Host "保存完了: $secretFile"
Write-Host 'このファイルはDPAPIで暗号化され、同じWindowsユーザーのみ復号できます。'
