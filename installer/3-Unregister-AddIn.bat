@echo off
rem Double-click launcher — unregisters the add-in for the current user.
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0Unregister-AddIn.ps1"
echo.
pause
