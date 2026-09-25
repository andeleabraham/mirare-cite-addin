<#
.SYNOPSIS
    Registers the Mirare Cite Word COM add-in for the current user.

.DESCRIPTION
    This script performs the FOUR steps required for a Word COM add-in to
    show up on the ribbon:

      1. Locates the 64-bit regasm and runs it against MirareCiteAddIn.dll.
         64-bit regasm is mandatory because you're running 64-bit Word.
      2. Writes the registry keys under HKCU\Software\Microsoft\Office\Word\
         Addins\MirareCite.AddIn with LoadBehavior = 3 (auto-load).
      3. Clears Word's disabled-items list so a previous crash doesn't
         keep the add-in blocked.
      4. Prints a verification: the registry keys, the existing log file
         if any, and instructions for the next Word launch.

    No admin rights are required (everything is HKCU + per-user regasm).

.PARAMETER DllPath
    Full path to MirareCiteAddIn.dll. Defaults to the Release x64 output
    folder of the .csproj.

.PARAMETER ProgId
    The ProgID exposed by the Connect class.  Must match the [ProgId(...)]
    attribute in Connect.cs.

.EXAMPLE
    .\Register-AddIn.ps1
    .\Register-AddIn.ps1 -DllPath "C:\src\MirareCiteAddIn\bin\x64\Release\MirareCiteAddIn.dll"
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
$DllPath = (Resolve-Path $DllPath).Path

if (-not (Test-Path $DllPath)) {
    throw "DLL not found at: $DllPath`nBuild the project first (see README.md → Build)."
}

Write-Host "==> Registering $DllPath" -ForegroundColor Cyan
Write-Host "    ProgID: $ProgId"
Write-Host "    Target: 64-bit Word 2016+"
Write-Host ""

# ── 1. regasm via Framework64 ──────────────────────────────────────────────
# Plain "regasm /codebase" writes to HKEY_CLASSES_ROOT (= HKLM\Software\Classes)
# and FAILS without admin. Instead: generate the exact .reg payload regasm
# would write, redirect it to HKCU\Software\Classes (per-user COM — honored
# by CoCreateInstance with no elevation), and import it.
$regasm = 'C:\Windows\Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe'
if (-not (Test-Path $regasm)) {
    throw "64-bit regasm not found at: $regasm`nIs .NET Framework 4.x installed?"
}

Write-Host "[1/4] Registering COM class per-user (HKCU\Software\Classes)…"
$tmpReg = Join-Path $env:TEMP "MirareCiteAddIn-per-user.reg"
if (Test-Path $tmpReg) { Remove-Item $tmpReg -Force }
& $regasm /nologo /codebase /regfile:"$tmpReg" $DllPath | Out-Null
if (-not (Test-Path $tmpReg)) {
    throw "regasm /regfile did not produce $tmpReg"
}
# Redirect machine-wide HKCR keys to the per-user classes hive.
$payload = Get-Content $tmpReg -Raw
$payload = $payload -replace '\[HKEY_CLASSES_ROOT', '[HKEY_CURRENT_USER\Software\Classes'
Set-Content -Path $tmpReg -Value $payload -Encoding Unicode
# reg.exe writes its success line to stderr; merging streams inside cmd
# (not in PowerShell) keeps PS 5.1 from raising a terminating
# NativeCommandError under $ErrorActionPreference='Stop'.
$importOut = cmd /c "reg import `"$tmpReg`" 2>&1"
Write-Host $importOut
if ($LASTEXITCODE -ne 0) { throw "reg import failed: $importOut" }
Remove-Item $tmpReg -Force -ErrorAction SilentlyContinue

# ── 2. Write the Word Addins registry key ─────────────────────────────────
Write-Host "[2/4] Writing HKCU\Software\Microsoft\Office\Word\Addins\$ProgId…"
$key = "HKCU:\Software\Microsoft\Office\Word\Addins\$ProgId"
if (-not (Test-Path $key)) { New-Item -Path $key -Force | Out-Null }
Set-ItemProperty -Path $key -Name 'FriendlyName' -Value 'Mirare Cite' -Type String
Set-ItemProperty -Path $key -Name 'Description'  -Value 'Word COM add-in for Mirare Cite.' -Type String
Set-ItemProperty -Path $key -Name 'LoadBehavior' -Value 3 -Type DWord
Set-ItemProperty -Path $key -Name 'CommandLineSafe' -Value 0 -Type DWord

# ── 3. Clear Word's disabled-items list ────────────────────────────────────
Write-Host "[3/4] Clearing Word's Resiliency\DisabledItems (if any)…"
$disabled = "HKCU:\Software\Microsoft\Office\16.0\Word\Resiliency\DisabledItems"
if (Test-Path $disabled) {
    Get-ChildItem $disabled | ForEach-Object {
        Remove-Item $_.PSPath -Force -ErrorAction SilentlyContinue
    }
    Write-Host "      cleared."
} else {
    Write-Host "      (none to clear)"
}

# Also enable Office's own add-in logging for diagnosis
$debugKey = "HKCU:\Software\Microsoft\Office\16.0\Common\Debug"
if (-not (Test-Path $debugKey)) { New-Item -Path $debugKey -Force | Out-Null }
Set-ItemProperty -Path $debugKey -Name 'EnableAddInLogging' -Value 1 -Type DWord
Set-ItemProperty -Path $debugKey -Name 'EnableRibbonExtensibilityLogging' -Value 1 -Type DWord

# ── 4. Verify ──────────────────────────────────────────────────────────────
Write-Host "[4/4] Verifying…"
$lb = (Get-ItemProperty $key -Name LoadBehavior).LoadBehavior
Write-Host "    LoadBehavior = $lb  (expected: 3)"

Write-Host ""
Write-Host "==> Done.  Restart Microsoft Word." -ForegroundColor Green
Write-Host "    A new 'Mirare Cite' tab should appear on the ribbon."
Write-Host "    If it doesn't, run:  .\verify-loadbehavior.ps1"
Write-Host "    And see:  DIAGNOSIS.md  (sections 1, 4, 5 in particular)."
