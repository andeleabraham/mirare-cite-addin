<#
.SYNOPSIS
    Diagnoses whether Word can see the Mirare Cite add-in and reports any
    blockers (registry, bitness, disabled-items list, recent event-log
    errors from the add-in).

.DESCRIPTION
    Run this if Word launched but the 'Mirare Cite' tab is not on the
    ribbon.  It prints:
      - Office + Word version + Word bitness
      - LoadBehavior for MirareCite.AddIn
      - Whether the add-in is in Word's disabled-items list
      - The last 10 Office/.NET entries in the Application event log

.EXAMPLE
    .\verify-loadbehavior.ps1
#>

$ProgId = 'MirareCite.AddIn'

Write-Host ""
Write-Host "=================== Mirare Cite Add-in Diagnostic ===================" -ForegroundColor Cyan

# ── Word version + bitness ─────────────────────────────────────────────────
try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $ver = $word.Version
    $word.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
    Write-Host "Word version      : $ver"
} catch {
    Write-Host "Word version      : (could not start Word.Application — $($_.Exception.Message))" -ForegroundColor Red
}

# Covers Click-to-Run (...\root\Office16\...) and MSI (...\Office16\...) layouts.
$x64 = (Test-Path 'C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE') -or
       (Test-Path 'C:\Program Files\Microsoft Office\Office16\WINWORD.EXE')
$x86 = (Test-Path 'C:\Program Files (x86)\Microsoft Office\root\Office16\WINWORD.EXE') -or
       (Test-Path 'C:\Program Files (x86)\Microsoft Office\Office16\WINWORD.EXE')
$bitness = if ($x64) { '64-bit (x64)' } elseif ($x86) { '32-bit (x86)' } else { 'unknown — is Office installed under Program Files?' }
Write-Host "Word bitness      : $bitness"

# ── Addin registry key ────────────────────────────────────────────────────
$key = "HKCU:\Software\Microsoft\Office\Word\Addins\$ProgId"
if (Test-Path $key) {
    $props = Get-ItemProperty $key
    Write-Host "Add-in key        : PRESENT ($key)"
    Write-Host "  FriendlyName   : $($props.FriendlyName)"
    Write-Host "  LoadBehavior   : $($props.LoadBehavior)  (3 = auto-load; 9/16 = disabled)"
} else {
    Write-Host "Add-in key        : MISSING ($key)" -ForegroundColor Red
    Write-Host "  → Run Register-AddIn.ps1 to create it."
}

# ── Disabled items ─────────────────────────────────────────────────────────
$disabled = 'HKCU:\Software\Microsoft\Office\16.0\Word\Resiliency\DisabledItems'
if (Test-Path $disabled) {
    $items = Get-ChildItem $disabled
    if ($items) {
        Write-Host ""
        Write-Host "DISABLED ITEMS    : $($items.Count) entry/entries found" -ForegroundColor Red
        Write-Host "  → Word has hard-disabled one or more add-ins. Clear with:"
        Write-Host "    Remove-Item '$disabled\*' -Force"
        Write-Host "    New-ItemProperty '$key' -Name LoadBehavior -Value 3 -Type DWord -Force"
    } else {
        Write-Host "Disabled items    : (none)"
    }
} else {
    Write-Host "Disabled items    : (none — key does not exist)"
}

# ── Recent event-log entries ───────────────────────────────────────────────
Write-Host ""
Write-Host "Recent Office/.NET Application event log (last 10):" -ForegroundColor Cyan
try {
    Get-WinEvent -LogName Application -MaxEvents 500 `
      | Where-Object {
          $_.ProviderName -match 'Office|Word|\.NET|Application' -or
          $_.Message -match 'Mirare|MirareCite|MirareCite'
        } `
      | Select-Object -First 10 `
      | Format-Table TimeCreated, Id, LevelDisplayName,
        @{N='Message';E={ if ($_.Message.Length -gt 100) { $_.Message.Substring(0, 100) + '…' } else { $_.Message } }} `
        -AutoSize
} catch {
    Write-Host "  (could not read event log)"
}

Write-Host ""
Write-Host "If LoadBehavior = 3 and the add-in is not in the disabled list but" -ForegroundColor Yellow
Write-Host "the ribbon tab still doesn't appear, see DIAGNOSIS.md §4 (silent"
Write-Host "GetCustomUI exceptions) and check %LOCALAPPDATA%\MirareCite\addin.log"
Write-Host "for trace output."
