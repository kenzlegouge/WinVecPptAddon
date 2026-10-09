# Compiles the PdfDrop add-in with the C# compiler that ships with Windows (.NET Framework 4.x).
# No Visual Studio or SDK install required. Output: bin\PdfDropAddin.dll
$ErrorActionPreference = 'Stop'

$root = $PSScriptRoot
$fw = Join-Path $env:WINDIR 'Microsoft.NET\Framework64\v4.0.30319'
if (-not (Test-Path (Join-Path $fw 'csc.exe'))) { $fw = Join-Path $env:WINDIR 'Microsoft.NET\Framework\v4.0.30319' }
$csc = Join-Path $fw 'csc.exe'
if (-not (Test-Path $csc)) { throw "C# compiler not found under $env:WINDIR\Microsoft.NET" }

# WinRT metadata for Windows.Data.Pdf. System.Runtime.WindowsRuntime.dll requires the unified
# "Windows.winmd" from the Windows SDK (UnionMetadata\<version>\), not the per-namespace files in System32.
$union = Get-ChildItem "${env:ProgramFiles(x86)}\Windows Kits\10\UnionMetadata" -Directory -ErrorAction SilentlyContinue |
    Where-Object { $_.Name -match '^\d+(\.\d+)+$' -and (Test-Path (Join-Path $_.FullName 'Windows.winmd')) } |
    Sort-Object { [version]$_.Name } -Descending | Select-Object -First 1
if (-not $union) {
    throw "Windows SDK not found (needs ...\Windows Kits\10\UnionMetadata\<version>\Windows.winmd). Install the Windows 10/11 SDK, or use the prebuilt bin\PdfDropAddin.dll."
}
$winmdRefs = @((Join-Path $union.FullName 'Windows.winmd'))

$outDir = Join-Path $root 'bin'
New-Item -ItemType Directory -Force -Path $outDir | Out-Null
$out = Join-Path $outDir 'PdfDropAddin.dll'

$refs = @(
    (Join-Path $fw 'System.Runtime.WindowsRuntime.dll'),
    (Join-Path $fw 'System.Runtime.dll'),
    (Join-Path $fw 'System.Core.dll'),
    (Join-Path $fw 'System.Windows.Forms.dll'),
    (Join-Path $fw 'System.Drawing.dll'),
    (Join-Path $fw 'Microsoft.CSharp.dll')
) + $winmdRefs

$args = @('/nologo', '/target:library', '/platform:anycpu', '/optimize+', "/out:$out")
foreach ($r in $refs) { $args += "/r:$r" }
$args += (Get-ChildItem (Join-Path $root 'src') -Filter *.cs | Sort-Object Name | ForEach-Object { $_.FullName })

& $csc @args
if ($LASTEXITCODE -ne 0) { throw "Compilation failed." }
Write-Host "Built $out"
