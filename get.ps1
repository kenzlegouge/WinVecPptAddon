# One-line installer for WinVectorPptAddon. Run in PowerShell:
#
#   irm https://raw.githubusercontent.com/KLG/WinVectorPptAddon/main/get.ps1 | iex
#
# Downloads the latest release (add-in + bundled Poppler) into %LOCALAPPDATA%\WinVectorPptAddon
# and registers it for the current user. No admin rights needed. Re-running updates in place.
#
# Optional, set before running:  $WinVectorPptAddonDir = 'D:\somewhere'   (install location)
#                                $WinVectorPptAddonNoInstall = $true      (download/extract only)

$ErrorActionPreference = 'Stop'
$repo = 'KLG/WinVectorPptAddon'
$dir = if ($WinVectorPptAddonDir) { $WinVectorPptAddonDir } else { Join-Path $env:LOCALAPPDATA 'WinVectorPptAddon' }

if (Get-Process POWERPNT -ErrorAction SilentlyContinue) {
    Write-Host 'PowerPoint is running. Please close it, then run this installer again.' -ForegroundColor Yellow
    return
}

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
$headers = @{ 'User-Agent' = 'WinVectorPptAddon-get' }

# Prefer the latest release zip (includes Poppler); fall back to a snapshot of main (install.ps1 then downloads Poppler).
$url = $null
try {
    $release = Invoke-RestMethod "https://api.github.com/repos/$repo/releases/latest" -Headers $headers
    $asset = $release.assets | Where-Object { $_.name -like '*.zip' } | Select-Object -First 1
    if ($asset) { $url = $asset.browser_download_url; Write-Host "Downloading WinVectorPptAddon $($release.tag_name)..." }
}
catch { }
if (-not $url) {
    $url = "https://github.com/$repo/archive/refs/heads/main.zip"
    Write-Host 'Downloading WinVectorPptAddon (latest source snapshot)...'
}

$tmp = Join-Path $env:TEMP ('WinVectorPptAddon-' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $tmp | Out-Null
try {
    $zip = Join-Path $tmp 'WinVectorPptAddon.zip'
    Invoke-WebRequest $url -OutFile $zip -UseBasicParsing -Headers $headers
    Expand-Archive $zip -DestinationPath (Join-Path $tmp 'x') -Force
    $installer = Get-ChildItem (Join-Path $tmp 'x') -Recurse -Filter 'install.ps1' | Select-Object -First 1
    if (-not $installer) { throw 'install.ps1 not found in the downloaded archive' }
    $srcDir = $installer.DirectoryName

    New-Item -ItemType Directory -Force -Path $dir | Out-Null
    foreach ($stale in 'bin', 'src') { Remove-Item (Join-Path $dir $stale) -Recurse -Force -ErrorAction SilentlyContinue }
    Copy-Item (Join-Path $srcDir '*') $dir -Recurse -Force
    Get-ChildItem $dir -Recurse -File | Unblock-File -ErrorAction SilentlyContinue
    Write-Host "Files installed to $dir"
}
finally {
    Remove-Item $tmp -Recurse -Force -ErrorAction SilentlyContinue
}

if ($WinVectorPptAddonNoInstall) { Write-Host 'Skipping registration (NoInstall).'; return }
& (Join-Path $dir 'install.ps1')
