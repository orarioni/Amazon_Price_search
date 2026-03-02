$ErrorActionPreference = 'Stop'

$repoRoot = Resolve-Path (Join-Path $PSScriptRoot '..')
$configPath = Join-Path $repoRoot 'config.psd1'
$libPath = Join-Path $PSScriptRoot 'lib/AmazonPriceLib.psm1'

Import-Module $libPath -Force
$config = Import-PowerShellDataFile -Path $configPath
$secretFile = Join-Path $repoRoot $config.Paths.SecretsFile

Save-SecretsInteractive -SecretFile $secretFile
