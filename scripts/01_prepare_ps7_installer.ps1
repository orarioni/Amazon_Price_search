param(
    [ValidateSet('x64', 'x86', 'arm64')]
    [string]$Architecture = 'x64',
    [string]$OutputDirectory = '.\\installers'
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'


function Resolve-LatestPwshMsiAsset {
    param([string]$Arch)

    $apiUrl = 'https://api.github.com/repos/PowerShell/PowerShell/releases/latest'
    $release = Invoke-RestMethod -Uri $apiUrl -Headers @{ 'User-Agent' = 'AmazonPriceSearch-PS7Installer' }

    $pattern = "PowerShell-.*-win-$Arch\.msi$"
    $asset = $release.assets | Where-Object { $_.name -match $pattern } | Select-Object -First 1
    if (-not $asset) {
        throw "PowerShell 7 MSI asset was not found. arch=$Arch"
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

Write-Host "Downloading latest PowerShell 7 ($($assetInfo.Tag)): $($assetInfo.Name)"
Invoke-WebRequest -Uri $assetInfo.Url -OutFile $outFile

Write-Host "Saved to: $outFile"
Write-Host 'Copy this MSI to the manager PC and run it.'
