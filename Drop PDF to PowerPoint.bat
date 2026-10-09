@echo off
rem Double-click: opens the drop window.
rem Or drag PDF files directly onto this .bat to insert them right away.
powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "%~dp0DropPdfToPowerPoint.ps1" %*
