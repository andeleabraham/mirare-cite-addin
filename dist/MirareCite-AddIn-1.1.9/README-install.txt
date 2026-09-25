Mirare Cite â€” Word add-in
=========================

Install (no admin rights needed):
  1. Extract this folder anywhere (it can be deleted after installing).
  2. Double-click  Install-AddIn.bat
  3. Start (or restart) Microsoft Word â€” a "Mirare Cite" tab appears
     on the ribbon.

Requirements:
  - Windows 10 (1809+) or Windows 11  (the required .NET Framework 4.8
    is already part of Windows â€” nothing else to download)
  - 64-bit Microsoft Word 2016 or newer (desktop)

Uninstall:
  Double-click  Uninstall-AddIn.bat  (or use the Mirare app's uninstall).

If the tab does not appear:
  - Fully quit Word (all windows) and start it again â€” Word only reads
    add-ins at startup.
  - Check the log:  %LOCALAPPDATA%\MirareCite\addin.log
  - Word may have hard-disabled the add-in after a previous crash:
    File > Options > Add-ins > Manage: COM Add-ins > Go... > tick the
    "Mirare Cite" checkbox.

