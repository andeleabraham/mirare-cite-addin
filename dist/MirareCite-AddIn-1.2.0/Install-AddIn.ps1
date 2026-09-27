<#
.SYNOPSIS
    Self-contained per-user installer for the Mirare Cite Word add-in.
    Designed to run from the release zip (or be invoked silently by a
    parent installer, e.g. the Mirare desktop app's Inno Setup / NSIS).

.DESCRIPTION
    1. Copies the add-in payload (DLL + interop dependencies) to
       %LOCALAPPDATA%\Mirare\WordAddIn (skipped with -NoCopy when the
       parent installer already placed the files).
    2. Registers the COM class per-user (HKCU\Software\Classes) via
       regasm /regfile + redirect — no admin rights needed.
    3. Writes HKCU\Software\Microsoft\Office\Word\Addins\MirareCite.AddIn
       with LoadBehavior=3.
    .NET Framework 4.8 is assumed present (inbox on Windows 10 1809+ /
    Windows 11 — nothing to ship).

.PARAMETER PayloadDir
    Folder containing MirareCiteAddIn.dll and its dependency DLLs.
    Defaults to the folder this script lives in (zip layout).

.PARAMETER NoCopy
    Skip step 1 — files are already in place at PayloadDir.

.EXAMPLE
    .\Install-AddIn.ps1                     # double-click layout / zip
    .\Install-AddIn.ps1 -PayloadDir C:\app\addin -NoCopy   # parent installer
#>

[CmdletBinding()]
param(
    [string]$PayloadDir = '',
    [switch]$NoCopy
)

$ErrorActionPreference = 'Stop'
if (-not $PayloadDir) { $PayloadDir = $PSScriptRoot }
$PayloadDir = (Resolve-Path $PayloadDir).Path

$DllName   = 'MirareCiteAddIn.dll'
$ProgId    = 'MirareCite.AddIn'
$ClassId   = '{B3F5D2A8-7E1A-4C2B-9D3F-8A1B2C3D4E5F}'   # keep in sync with Connect.cs [Guid]
$InstallDir = Join-Path $env:LOCALAPPDATA 'Mirare\WordAddIn'

$DllPath = Join-Path $PayloadDir $DllName
if (-not (Test-Path $DllPath)) { throw "MirareCiteAddIn.dll not found in: $PayloadDir" }

Write-Host "==> Installing Mirare Cite Word add-in (per-user, no admin)" -ForegroundColor Cyan

# ── 1. Copy payload to a stable install location ──────────────────────────
if (-not $NoCopy) {
    if (-not (Test-Path $InstallDir)) { New-Item -Path $InstallDir -ItemType Directory -Force | Out-Null }
    Copy-Item (Join-Path $PayloadDir '*.dll') $InstallDir -Force
    Copy-Item (Join-Path $PayloadDir '*.xml') $InstallDir -Force -ErrorAction SilentlyContinue
    Write-Host "[1/3] Copied payload -> $InstallDir"
    $DllPath = Join-Path $InstallDir $DllName
} else {
    Write-Host "[1/3] Skipping copy (-NoCopy) — using $PayloadDir"
}

# ── 2. Per-user COM registration ──────────────────────────────────────────
$regasm = 'C:\Windows\Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe'
if (-not (Test-Path $regasm)) { throw "64-bit regasm not found — is .NET Framework 4.x installed?" }

$tmpReg = Join-Path $env:TEMP "MirareCiteAddIn-install.reg"
if (Test-Path $tmpReg) { Remove-Item $tmpReg -Force }
& $regasm /nologo /codebase /regfile:"$tmpReg" $DllPath | Out-Null
if (-not (Test-Path $tmpReg)) { throw "regasm /regfile produced no output" }
$payload = Get-Content $tmpReg -Raw
$payload = $payload -replace '\[HKEY_CLASSES_ROOT', '[HKEY_CURRENT_USER\Software\Classes'
Set-Content -Path $tmpReg -Value $payload -Encoding Unicode
$importOut = cmd /c "reg import `"$tmpReg`" 2>&1"
if ($LASTEXITCODE -ne 0) { throw "reg import failed: $importOut" }
Remove-Item $tmpReg -Force -ErrorAction SilentlyContinue
Write-Host "[2/3] COM class registered per-user ($ProgId)"

# ── 3. Word Addins key ────────────────────────────────────────────────────
$key = "HKCU:\Software\Microsoft\Office\Word\Addins\$ProgId"
if (-not (Test-Path $key)) { New-Item -Path $key -Force | Out-Null }
Set-ItemProperty -Path $key -Name 'FriendlyName' -Value 'Mirare Cite' -Type String
Set-ItemProperty -Path $key -Name 'Description' -Value 'Word COM add-in for Mirare Cite.' -Type String
Set-ItemProperty -Path $key -Name 'LoadBehavior' -Value 3 -Type DWord
Set-ItemProperty -Path $key -Name 'CommandLineSafe' -Value 0 -Type DWord

# Clear Word's disabled-items list + enable ribbon-extensibility logging
$disabled = "HKCU:\Software\Microsoft\Office\16.0\Word\Resiliency\DisabledItems"
if (Test-Path $disabled) {
    Get-ChildItem $disabled | ForEach-Object { Remove-Item $_.PSPath -Force -ErrorAction SilentlyContinue }
}
$debugKey = "HKCU:\Software\Microsoft\Office\16.0\Common\Debug"
if (-not (Test-Path $debugKey)) { New-Item -Path $debugKey -Force | Out-Null }
Set-ItemProperty -Path $debugKey -Name 'EnableRibbonExtensibilityLogging' -Value 1 -Type DWord

Write-Host "[3/3] Word Addins key written (LoadBehavior=3)"
Write-Host ""
Write-Host "==> Done. Start (or restart) Microsoft Word — look for the 'Mirare Cite' tab." -ForegroundColor Green
