Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

param(
    [ValidateSet('x64', 'x86', 'arm64')]
    [string]$Architecture = 'x64',
    [string]$OutputDirectory = '.\\installers'
)

function Resolve-LatestPwshMsiAsset {
    param([string]$Arch)

    $apiUrl = 'https://api.github.com/repos/PowerShell/PowerShell/releases/latest'
    $release = Invoke-RestMethod -Uri $apiUrl -Headers @{ 'User-Agent' = 'AmazonPriceSearch-PS7Installer' }

    $pattern = "PowerShell-.*-win-$Arch\.msi$"
    $asset = $release.assets | Where-Object { $_.name -match $pattern } | Select-Object -First 1
    if (-not $asset) {
        throw "PowerShell 7 の MSI アセットが見つかりませんでした。arch=$Arch"
    }

    return [pscustomobject]@{
        Name = $asset.name
        Url  = $asset.browser_download_url
        Tag  = $release.tag_name
    }
}

$repoRoot = Split-Path -Path $PSScriptRoot -Parent
$targetDir = Join-Path -Path $repoRoot -ChildPath $OutputDirectory
if (-not (Test-Path -Path $targetDir)) {
    New-Item -Path $targetDir -ItemType Directory | Out-Null
}

$assetInfo = Resolve-LatestPwshMsiAsset -Arch $Architecture
$outFile = Join-Path -Path $targetDir -ChildPath $assetInfo.Name

Write-Host "PowerShell 7 最新版 ($($assetInfo.Tag)) を取得します: $($assetInfo.Name)"
Invoke-WebRequest -Uri $assetInfo.Url -OutFile $outFile

Write-Host "保存先: $outFile"
Write-Host '上司PCへこの MSI をコピーして実行してください。'
