# Removes the PdfDrop add-in registration for the current user.
$ErrorActionPreference = 'SilentlyContinue'

$clsid = '{7D3F0C8E-5B1A-4E7B-9C2D-3A8F1E6B4D21}'
$progId = 'PdfDrop.Addin'

Remove-Item "HKCU:\Software\Microsoft\Office\PowerPoint\Addins\$progId" -Recurse -Force
Remove-Item "HKCU:\Software\Classes\CLSID\$clsid" -Recurse -Force
Remove-Item "HKCU:\Software\Classes\Wow6432Node\CLSID\$clsid" -Recurse -Force
Remove-Item "HKCU:\Software\Classes\$progId" -Recurse -Force

Write-Host 'PdfDrop add-in unregistered. Restart PowerPoint to finish.' -ForegroundColor Green
