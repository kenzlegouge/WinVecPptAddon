@echo off
rem Double-click to install the add-in for the current user (no admin rights needed).
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0install.ps1"
echo.
pause
