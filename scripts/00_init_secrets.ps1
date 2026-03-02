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

Write-Host ''
Write-Host 'SP-API呼び出しにはAWS SigV4認証情報が必須です。'
$awsAccessKeyId = Read-Host 'aws_access_key_id'
$awsSecretAccessKey = Read-Host 'aws_secret_access_key' -AsSecureString
$awsSessionToken = Read-Host 'aws_session_token (不要なら空欄でEnter)'

$payload = [PSCustomObject]@{
    client_id             = $clientId
    client_secret         = $clientSecret
    refresh_token         = $refreshToken
    aws_access_key_id     = $awsAccessKeyId
    aws_secret_access_key = $awsSecretAccessKey
    aws_session_token     = $awsSessionToken
    created_at            = (Get-Date).ToString('o')
}

$payload | Export-Clixml -Path $secretFile

Write-Host "保存完了: $secretFile"
Write-Host 'このファイルはDPAPIで暗号化され、同じWindowsユーザーのみ復号できます。'
