@echo off
rem Double-click to remove the add-in registration for the current user.
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0uninstall.ps1"
echo.
pause
