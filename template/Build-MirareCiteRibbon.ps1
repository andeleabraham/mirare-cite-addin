<#
.SYNOPSIS
    Builds MirareCiteRibbon.dotm from customUI.xml + MirareCiteRibbon.vba.bas
    on Windows using the Word COM API.

.DESCRIPTION
    This script creates a fresh .dotm template, embeds the customUI.xml as a
    ribbon relationship, imports the VBA module, saves, and copies it to the
    Word STARTUP folder.

    Why a script and not a static .dotm?  Because generating vbaProject.bin
    (the binary CFB container that holds VBA inside a .dotm) from scratch on
    Linux is not feasible without Microsoft tooling.  Running this script on
    Windows via the Word COM API is the cleanest path.

.NOTES
    Run from a normal PowerShell (no admin needed unless the STARTUP folder
    is in Program Files, which it isn't for user installs).

    Requires:
      - Microsoft Word 2016+ installed and macro security set to at least
        "Notifications for all macros".
      - The two source files (customUI.xml, MirareCiteRibbon.vba.bas) next
        to this script.

.EXAMPLE
    .\Build-MirareCiteRibbon.ps1
#>

[CmdletBinding()]
param(
    [string]$OutputPath = "$env:APPDATA\Microsoft\Word\STARTUP\MirareCiteRibbon.dotm",
    [string]$CustomUiPath = "$PSScriptRoot\customUI.xml",
    [string]$VbaSourcePath = "$PSScriptRoot\MirareCiteRibbon.vba.bas"
)

$ErrorActionPreference = 'Stop'

# ── Sanity checks ────────────────────────────────────────────────────────────
foreach ($p in @($CustomUiPath, $VbaSourcePath)) {
    if (-not (Test-Path $p)) { throw "Source file not found: $p" }
}
$outDir = Split-Path $OutputPath -Parent
if (-not (Test-Path $outDir)) {
    New-Item -ItemType Directory -Path $outDir -Force | Out-Null
    Write-Host "Created STARTUP folder: $outDir"
}

# ── 1. Create a fresh .dotm with the customUI relationship ──────────────────
# Word's SaveAs2 will pack customUI.xml into the .dotm ZIP correctly when
# the CustomUI part is added through the Office.CustomXMLParts mechanism.
# But the simplest, most reliable path is to use the Word.Application object
# model directly: create doc → save as .dotm → unzip → inject customUI.xml
# + relationship → re-zip.  Doing the unzip/re-zip ourselves avoids relying
# on undocumented Word APIs.

$tempDir = Join-Path $env:TEMP ("mrc_build_" + [System.Guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $tempDir -Force | Out-Null

try {
    Write-Host "Step 1/4: creating fresh .dotm via Word COM API…"
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0  # wdAlertsNone
    $doc = $word.Documents.Add()

    $tempDotm = Join-Path $tempDir "MirareCiteRibbon.dotm"
    $doc.SaveAs2($tempDotm, 16)   # 16 = wdFormatXMLTemplateMacroEnabled = .dotm

    # Step 2 — import the VBA module
    Write-Host "Step 2/4: importing VBA module…"
    $vbaProject = $doc.VBProject
    $module = $vbaProject.VBComponents.Import($VbaSourcePath)
    # Ensure module name matches the customUI.xml callbacks.  The .bas file
    # already has Attribute VB_Name = "MirareCiteRibbon", but if not, set it.
    if ($module.Name -ne 'MirareCiteRibbon') {
        $module.Name = 'MirareCiteRibbon'
    }
    $doc.Save()

    $doc.Close()
    $word.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null

    # Step 3 — unzip the .dotm, inject customUI.xml + relationship
    Write-Host "Step 3/4: injecting customUI.xml…"
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    [System.IO.Compression.ZipFile]::ExtractToDirectory($tempDotm, (Join-Path $tempDir 'pkg'))

    $pkgDir = Join-Path $tempDir 'pkg'
    $customUiDir = Join-Path $pkgDir 'customUI'
    New-Item -ItemType Directory -Path $customUiDir -Force | Out-Null
    Copy-Item $CustomUiPath (Join-Path $customUiDir 'customUI.xml')

    # Update _rels/.rels to add a relationship to customUI\customUI.xml
    $relsPath = Join-Path $pkgDir '_rels\.rels'
    [xml]$rels = Get-Content $relsPath
    $ns = 'http://schemas.openxmlformats.org/package/2006/relationships'
    $newRel = $rels.Relationships.AppendChild($rels.CreateElement('Relationship', $ns))
    $newRel.SetAttribute('Id', "rIdCustomUI1")
    $newRel.SetAttribute('Type',
        'http://schemas.microsoft.com/office/2007/relationships/encodedTarget')
    $newRel.SetAttribute('Target', 'customUI/customUI.xml')
    $rels.Save($relsPath)

    # Re-zip
    if (Test-Path $tempDotm) { Remove-Item $tempDotm -Force }
    [System.IO.Compression.ZipFile]::CreateFromDirectory(
        $pkgDir, $tempDotm,
        [System.IO.Compression.CompressionLevel]::Optimal, $false)

    # Step 4 — copy to STARTUP
    Write-Host "Step 4/4: copying to $OutputPath"
    Copy-Item $tempDotm $OutputPath -Force

    Write-Host ""
    Write-Host "Done." -ForegroundColor Green
    Write-Host "  Built: $OutputPath"
    Write-Host "  Restart Word — the 'Mirare Cite' tab should appear on the ribbon."
    Write-Host "  If it doesn't, run the diagnostic steps in DIAGNOSIS.md §3."
}
finally {
    if (Test-Path $tempDir) { Remove-Item $tempDir -Recurse -Force }
}
