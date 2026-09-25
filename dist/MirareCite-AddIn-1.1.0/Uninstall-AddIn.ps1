<#
.SYNOPSIS
    Per-user uninstaller for the Mirare Cite Word add-in (reverses
    Install-AddIn.ps1: registry keys + install folder).

.EXAMPLE
    .\Uninstall-AddIn.ps1
#>

[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'
$ProgId  = 'MirareCite.AddIn'
$ClassId = '{B3F5D2A8-7E1A-4C2B-9D3F-8A1B2C3D4E5F}'
$InstallDir = Join-Path $env:LOCALAPPDATA 'Mirare\WordAddIn'

Write-Host "==> Uninstalling Mirare Cite Word add-in" -ForegroundColor Cyan

# 1. Word Addins key
$key = "HKCU:\Software\Microsoft\Office\Word\Addins\$ProgId"
if (Test-Path $key) { Remove-Item $key -Recurse -Force; Write-Host "[1/3] Removed $key" }
else { Write-Host "[1/3] Key did not exist: $key" }

# 2. Per-user COM registration
$removed = @()
foreach ($k in "HKCU:\Software\Classes\$ProgId",
               "HKCU:\Software\Classes\CLSID\$ClassId") {
    if (Test-Path $k) { Remove-Item $k -Recurse -Force; $removed += $k }
}
if ($removed) { $removed | ForEach-Object { Write-Host "[2/3] Removed $_" } }
else { Write-Host "[2/3] No per-user COM keys found" }

# 3. Install folder
if (Test-Path $InstallDir) {
    Remove-Item $InstallDir -Recurse -Force
    Write-Host "[3/3] Removed $InstallDir"
} else {
    Write-Host "[3/3] Install folder not present: $InstallDir"
}

Write-Host ""
Write-Host "==> Done. Restart Word — the 'Mirare Cite' tab is gone." -ForegroundColor Green
