# Builds dist\WinVectorPptAddon-<version>.zip: the whole folder including the bundled Poppler,
# ready to attach to a GitHub release (get.ps1 and the README's "Zip" install look for it there).
param([string]$Version = '1.0.0')
$ErrorActionPreference = 'Stop'
$root = $PSScriptRoot

if (-not (Test-Path (Join-Path $root 'tools\poppler\bin\pdftocairo.exe'))) {
    throw 'tools\poppler is missing. Run install.ps1 first so it gets downloaded.'
}

$dist = Join-Path $root 'dist'
New-Item -ItemType Directory -Force -Path $dist | Out-Null
$stage = Join-Path $env:TEMP "WinVectorPptAddon-$Version"
Remove-Item $stage -Recurse -Force -ErrorAction SilentlyContinue
New-Item -ItemType Directory -Path $stage | Out-Null

$exclude = @('.git', 'dist', '.gitignore')
Get-ChildItem $root -Force |
    Where-Object { $exclude -notcontains $_.Name -and $_.Extension -ne '.pdf' } |
    Copy-Item -Destination $stage -Recurse -Force

$zip = Join-Path $dist "WinVectorPptAddon-$Version.zip"
Remove-Item $zip -Force -ErrorAction SilentlyContinue
Compress-Archive -Path $stage -DestinationPath $zip -CompressionLevel Optimal
Remove-Item $stage -Recurse -Force
Write-Host "Created $zip ($([math]::Round((Get-Item $zip).Length / 1MB, 1)) MB)" -ForegroundColor Green
