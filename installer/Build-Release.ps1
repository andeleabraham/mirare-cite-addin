<#
.SYNOPSIS
    Builds the add-in (Release|x64) and packages a distributable zip:
        dist\MirareCite-AddIn-<Version>.zip

    The zip layout is what end users get from the website:
        MirareCite-AddIn-<Version>\
            MirareCiteAddIn.dll + dependency DLLs
            Install-AddIn.ps1 / Uninstall-AddIn.ps1
            Install-AddIn.bat / Uninstall-AddIn.bat   (double-click wrappers)
            README-install.txt

.DESCRIPTION
    .NET Framework 4.8 is inbox on Windows 10 1809+/11, so the zip needs
    no runtime download — users extract and double-click Install-AddIn.bat.

.PARAMETER Version
    Version string baked into the zip filename.

.EXAMPLE
    .\Build-Release.ps1 -Version 1.0.0
#>

[CmdletBinding()]
param(
    [string]$Version = '1.0.0'
)

$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path $PSScriptRoot -Parent
$csproj   = Join-Path $repoRoot 'src\MirareCiteAddIn\MirareCiteAddIn.csproj'
$distDir  = Join-Path $repoRoot "dist\MirareCite-AddIn-$Version"

# ── 1. Locate MSBuild (VS 2022) and build ─────────────────────────────────
$vswhere = "${env:ProgramFiles(x86)}\Microsoft Visual Studio\Installer\vswhere.exe"
$msbuild = & $vswhere -latest -requires Microsoft.Component.MSBuild -find "MSBuild\**\Bin\MSBuild.exe" | Select-Object -First 1
if (-not $msbuild) { throw "MSBuild not found — is Visual Studio 2022 installed?" }

Write-Host "==> Building Release|x64…"
& $msbuild $csproj -t:Restore,Build -p:Configuration=Release -p:Platform=x64 -v:m -nologo
if ($LASTEXITCODE -ne 0) { throw "Build failed with exit code $LASTEXITCODE" }

$bin = Join-Path $repoRoot 'src\MirareCiteAddIn\bin\x64\Release'
if (-not (Test-Path (Join-Path $bin 'MirareCiteAddIn.dll'))) { throw "Build output missing: $bin" }

# ── 2. Stage the payload ──────────────────────────────────────────────────
if (Test-Path $distDir) { Remove-Item $distDir -Recurse -Force }
New-Item -Path $distDir -ItemType Directory -Force | Out-Null

Copy-Item (Join-Path $bin '*.dll') $distDir -Force
Copy-Item (Join-Path $bin '*.xml') $distDir -Force -ErrorAction SilentlyContinue
Copy-Item (Join-Path $PSScriptRoot 'Install-AddIn.ps1')   $distDir -Force
Copy-Item (Join-Path $PSScriptRoot 'Uninstall-AddIn.ps1') $distDir -Force
Copy-Item (Join-Path $PSScriptRoot 'zip\Install-AddIn.bat')   $distDir -Force
Copy-Item (Join-Path $PSScriptRoot 'zip\Uninstall-AddIn.bat') $distDir -Force
Copy-Item (Join-Path $PSScriptRoot 'zip\README-install.txt')  $distDir -Force
Write-Host "==> Staged payload -> $distDir"

# ── 3. Zip it ─────────────────────────────────────────────────────────────
$zipPath = Join-Path $repoRoot "dist\MirareCite-AddIn-$Version.zip"
if (Test-Path $zipPath) { Remove-Item $zipPath -Force }
Compress-Archive -Path $distDir -DestinationPath $zipPath
Write-Host "==> Packaged: $zipPath" -ForegroundColor Green
