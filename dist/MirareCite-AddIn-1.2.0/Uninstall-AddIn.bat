@echo off
rem Removes the Mirare Cite Word add-in for the current user.
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0Uninstall-AddIn.ps1"
echo.
pause
