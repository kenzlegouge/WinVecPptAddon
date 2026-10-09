# Builds and registers the PdfDrop add-in for the current user (no admin rights needed).
# After this, restart PowerPoint: dropping a PDF on a slide inserts it as a picture.
$ErrorActionPreference = 'Stop'

$root = $PSScriptRoot
$dll = Join-Path $root 'bin\PdfDropAddin.dll'
$src = Get-ChildItem (Join-Path $root 'src') -Filter *.cs | Sort-Object LastWriteTime -Descending | Select-Object -First 1

# Files downloaded from the internet carry a "blocked" mark that stops PowerPoint from loading the add-in.
Get-ChildItem $root -Recurse -File | Unblock-File -ErrorAction SilentlyContinue

# Rebuild only when the source is clearly newer than the prebuilt DLL (a fresh checkout or unzip gives both
# about the same timestamp). Building needs the Windows SDK; if it is not there, the prebuilt DLL is used.
$needBuild = -not (Test-Path $dll) -or ($src.LastWriteTime - (Get-Item $dll).LastWriteTime).TotalSeconds -gt 10
if ($needBuild) {
    try { & (Join-Path $root 'build.ps1') }
    catch {
        if (-not (Test-Path $dll)) { throw }
        Write-Warning "Not rebuilding ($($_.Exception.Message)). Using the prebuilt $dll"
    }
}
else {
    Write-Host "Using prebuilt $dll"
}

# --- PDF -> SVG converter (Poppler's pdftocairo). Bundled in tools\poppler; downloaded if missing. ---
$tools = Join-Path $root 'tools\poppler'
if (-not (Test-Path (Join-Path $tools 'bin\pdftocairo.exe'))) {
    Write-Host 'Downloading Poppler (pdftocairo) for the PDF -> vector conversion...'
    try {
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        $release = Invoke-RestMethod 'https://api.github.com/repos/oschwartz10612/poppler-windows/releases/latest' -Headers @{ 'User-Agent' = 'PdfDrop-install' }
        $asset = $release.assets | Where-Object { $_.name -like '*.zip' } | Select-Object -First 1
        $tmp = Join-Path $env:TEMP ('PdfDrop-poppler-' + [guid]::NewGuid().ToString('N'))
        New-Item -ItemType Directory -Path $tmp | Out-Null
        $zip = Join-Path $tmp $asset.name
        Invoke-WebRequest $asset.browser_download_url -OutFile $zip -UseBasicParsing
        Expand-Archive $zip -DestinationPath (Join-Path $tmp 'x') -Force
        $exe = Get-ChildItem (Join-Path $tmp 'x') -Recurse -Filter 'pdftocairo.exe' | Select-Object -First 1
        if (-not $exe) { throw 'pdftocairo.exe not found in the downloaded archive' }
        New-Item -ItemType Directory -Force -Path (Join-Path $tools 'bin'), (Join-Path $tools 'share') | Out-Null
        Copy-Item $exe.FullName (Join-Path $tools 'bin') -Force
        Copy-Item (Join-Path $exe.DirectoryName '*.dll') (Join-Path $tools 'bin') -Force
        $data = Get-ChildItem (Join-Path $tmp 'x') -Recurse -Directory -Filter 'cMap' | Select-Object -First 1
        if ($data) { Copy-Item $data.Parent.FullName (Join-Path $tools 'share\poppler') -Recurse -Force }
        Remove-Item $tmp -Recurse -Force -ErrorAction SilentlyContinue
        Write-Host "Poppler $($release.tag_name) installed to $tools"
    }
    catch {
        Write-Warning "Could not download Poppler ($($_.Exception.Message)). Without it, PDFs are inserted as bitmaps unless pdftocairo or Inkscape is installed."
    }
}
else {
    Write-Host "Using bundled pdftocairo in $tools"
}

$clsid = '{7D3F0C8E-5B1A-4E7B-9C2D-3A8F1E6B4D21}'
$progId = 'PdfDrop.Addin'
$className = 'PdfDrop.Addin'
$assembly = 'PdfDropAddin, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null'
$codeBase = 'file:///' + ($dll -replace '\\', '/')

if (Get-Process POWERPNT -ErrorAction SilentlyContinue) {
    Write-Warning 'PowerPoint is running. Close and reopen it after installation for the add-in to load.'
}

function Set-RegDefault([string]$Path, [string]$Value) {
    New-Item -Path $Path -Force | Out-Null
    Set-ItemProperty -Path $Path -Name '(default)' -Value $Value
}

# --- COM class registration (manual equivalent of RegAsm /codebase, but per-user) ---
# 32-bit Office reads HKCU\Software\Classes\Wow6432Node\CLSID, 64-bit Office reads HKCU\Software\Classes\CLSID.
$clsidRoots = @('HKCU:\Software\Classes\CLSID', 'HKCU:\Software\Classes\Wow6432Node\CLSID')
foreach ($rootKey in $clsidRoots) {
    $key = "$rootKey\$clsid"
    Set-RegDefault $key $className
    Set-RegDefault "$key\ProgId" $progId
    Set-RegDefault "$key\Implemented Categories\{62C8FE65-4EBB-45e7-B440-6E39B2CDBF29}" ''
    $inproc = "$key\InprocServer32"
    Set-RegDefault $inproc 'mscoree.dll'
    Set-ItemProperty $inproc -Name 'ThreadingModel' -Value 'Both'
    Set-ItemProperty $inproc -Name 'Class' -Value $className
    Set-ItemProperty $inproc -Name 'Assembly' -Value $assembly
    Set-ItemProperty $inproc -Name 'RuntimeVersion' -Value 'v4.0.30319'
    Set-ItemProperty $inproc -Name 'CodeBase' -Value $codeBase
    $ver = "$inproc\1.0.0.0"
    New-Item -Path $ver -Force | Out-Null
    Set-ItemProperty $ver -Name 'Class' -Value $className
    Set-ItemProperty $ver -Name 'Assembly' -Value $assembly
    Set-ItemProperty $ver -Name 'RuntimeVersion' -Value 'v4.0.30319'
    Set-ItemProperty $ver -Name 'CodeBase' -Value $codeBase
}
Set-RegDefault "HKCU:\Software\Classes\$progId" $className
Set-RegDefault "HKCU:\Software\Classes\$progId\CLSID" $clsid

# --- Tell PowerPoint to load it ---
$addinKey = "HKCU:\Software\Microsoft\Office\PowerPoint\Addins\$progId"
New-Item -Path $addinKey -Force | Out-Null
Set-ItemProperty $addinKey -Name 'FriendlyName' -Value 'PDF drop (insert PDFs as vector graphics)'
Set-ItemProperty $addinKey -Name 'Description' -Value 'Drag a PDF onto a slide to insert it as a vector graphic, like on macOS.'
Set-ItemProperty $addinKey -Name 'LoadBehavior' -Value 3 -Type DWord

# If PowerPoint had previously disabled the add-in, clear that.
Get-ChildItem 'HKCU:\Software\Microsoft\Office\16.0\PowerPoint\Resiliency\DisabledItems' -ErrorAction SilentlyContinue |
    ForEach-Object { Remove-Item $_.PSPath -Force -ErrorAction SilentlyContinue }

Write-Host ''
Write-Host 'Installed. Start (or restart) PowerPoint and drag a PDF onto a slide.' -ForegroundColor Green
