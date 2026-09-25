<#
.SYNOPSIS
    Unregisters the Mirare Cite Word COM add-in.

.DESCRIPTION
    Reverses everything Register-AddIn.ps1 did:
      1. Removes the HKCU Word Addins key for MirareCite.AddIn.
      2. Removes the per-user COM registration (HKCU\Software\Classes) that
         Register-AddIn.ps1 created — the ProgID key and the CLSID key
         from Connect.cs.
      3. Does NOT delete the .dll itself (in case you want to re-register
         a fresh build).

.EXAMPLE
    .\Unregister-AddIn.ps1
#>

[CmdletBinding()]
param(
    # Default resolved in the body — $PSScriptRoot is empty inside param()
    # default expressions under Windows PowerShell 5.1.
    [string]$DllPath = '',
    [string]$ProgId = 'MirareCite.AddIn'
)

$ErrorActionPreference = 'Stop'
if (-not $DllPath) {
    $DllPath = Join-Path $PSScriptRoot '..\src\MirareCiteAddIn\bin\x64\Release\MirareCiteAddIn.dll'
}
$ClassId = '{B3F5D2A8-7E1A-4C2B-9D3F-8A1B2C3D4E5F}'   # keep in sync with Connect.cs [Guid]

Write-Host "==> Unregistering Mirare Cite Word add-in" -ForegroundColor Cyan

# 1. Remove the Word Addins key
$key = "HKCU:\Software\Microsoft\Office\Word\Addins\$ProgId"
if (Test-Path $key) {
    Remove-Item $key -Recurse -Force
    Write-Host "[1/2] Removed $key"
} else {
    Write-Host "[1/2] Key did not exist: $key"
}

# 2. Remove the per-user COM registration (HKCU\Software\Classes).
#    Note: regasm /regfile cannot be combined with /unregister, and plain
#    "regasm /unregister" writes to HKCR (admin-only) — but we registered
#    by importing a regfile redirected to HKCU, so we simply delete those
#    keys. The ProgID/CLSID are stable by design (see Connect.cs).
$removed = @()
foreach ($k in "HKCU:\Software\Classes\$ProgId",
               "HKCU:\Software\Classes\CLSID\$ClassId") {
    if (Test-Path $k) {
        Remove-Item $k -Recurse -Force
        $removed += $k
    }
}
if ($removed) { $removed | ForEach-Object { Write-Host "[2/2] Removed $_" } }
else { Write-Host "[2/2] No per-user COM keys found (already unregistered)" }

Write-Host ""
Write-Host "==> Done.  Restart Word — the 'Mirare Cite' tab should be gone." -ForegroundColor Green
