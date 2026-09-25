@echo off
rem Double-click launcher — keeps the window open so you can read the output.
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0verify-loadbehavior.ps1"
echo.
pause
