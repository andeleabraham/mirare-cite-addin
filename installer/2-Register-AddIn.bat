@echo off
rem Double-click launcher — builds nothing; registers the already-built add-in
rem per-user (no admin needed) and keeps the window open afterwards.
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0Register-AddIn.ps1"
echo.
pause
