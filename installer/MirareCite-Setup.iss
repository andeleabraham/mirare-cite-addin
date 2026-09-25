; ============================================================================
;  MirareCite-Setup.iss — optional single-file setup.exe for the
;  Mirare Cite Word add-in, built with Inno Setup 6 (https://jrsoftware.org).
;
;  Why this exists: a .zip download works, but a signed setup.exe gives
;  public users a familiar installer + a proper uninstall entry.
;  Everything is per-user ({localappdata} + HKCU) — no admin elevation.
;
;  Build:
;    1. Run installer\Build-Release.ps1 first (stages dist\MirareCite-AddIn-<v>\)
;    2. Point #define ReleaseDir below at that folder
;    3. Compile this script with Inno Setup Compiler (ISCC.exe)
;
;  Bundle-with-PySide6-app variant: skip Inno entirely and call
;  Install-AddIn.ps1 -NoCopy from YOUR installer's [Run] section after
;  placing the payload files — see README.md "Packaging & distribution".
; ============================================================================

#define ReleaseDir "dist\MirareCite-AddIn-1.0.0"
#define AppName "Mirare Cite Word Add-in"
#define AppVersion "1.0.0"
#define ProgId "MirareCite.AddIn"

[Setup]
AppId={{8E6D3F2A-1B4C-4E2A-9F1D-7C8E5A6B2D40}
AppName={#AppName}
AppVersion={#AppVersion}
AppPublisher=Mirare
DefaultDirName={localappdata}\Mirare\WordAddIn
DefaultGroupName=Mirare Cite
; Everything is per-user — never request elevation
PrivilegesRequired=lowest
OutputDir=dist
OutputBaseFilename=MirareCite-AddIn-Setup-{#AppVersion}
Compression=lzma2
SolidCompression=yes
ArchitecturesInstallIn64BitMode=x64compatible
UninstallDisplayIcon={app}\MirareCiteAddIn.dll

[Files]
Source: "{#ReleaseDir}\*.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#ReleaseDir}\*.xml"; DestDir: "{app}"; Flags: ignoreversion skipifsourcedoesntexist
Source: "{#ReleaseDir}\Install-AddIn.ps1"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#ReleaseDir}\README-install.txt"; DestDir: "{app}"; Flags: ignoreversion isreadme

[Run]
; Register per-user (COM + Word Addins key). -NoCopy: files are already
; placed by this installer.
Filename: "powershell.exe"; \
    Parameters: "-NoProfile -ExecutionPolicy Bypass -File ""{app}\Install-AddIn.ps1"" -PayloadDir ""{app}"" -NoCopy"; \
    Flags: runhidden

[UninstallRun]
; Reverse the registration before files are deleted
Filename: "powershell.exe"; \
    Parameters: "-NoProfile -ExecutionPolicy Bypass -Command ""Remove-Item 'HKCU:\Software\Microsoft\Office\Word\Addins\{#ProgId}' -Recurse -Force -ErrorAction SilentlyContinue; Remove-Item 'HKCU:\Software\Classes\{#ProgId}' -Recurse -Force -ErrorAction SilentlyContinue; Remove-Item 'HKCU:\Software\Classes\CLSID\{{B3F5D2A8-7E1A-4C2B-9D3F-8A1B2C3D4E5F}' -Recurse -Force -ErrorAction SilentlyContinue"""; \
    Flags: runhidden; RunOnceId: "UnregAddIn"
