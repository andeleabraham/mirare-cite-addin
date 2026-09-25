@echo off
rem Installs the Mirare Cite Word add-in for the current user (no admin).
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0Install-AddIn.ps1"
echo.
pause
