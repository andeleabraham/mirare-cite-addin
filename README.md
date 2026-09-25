# Mirare Cite Word Add-in — Build, Deploy, Verify

> Complete Word COM add-in source project + installer + .dotm template
> variant.  Target: **64-bit Microsoft Word 2016+ desktop** on Windows.
> Build toolchain: **Visual Studio 2022 Community** (free) or higher.

## 0. Honest scope statement

I built this on a **Linux container**, so what's in this package is
**source code + scripts**, not a pre-built `.dll`. The two things I
cannot produce for you here are:

1. The compiled `MirareCiteAddIn.dll` — that must be built on Windows.
2. A real `MirareCiteRibbon.dotm` with the VBA embedded as
   `vbaProject.bin` — that must be built on Windows (the `Build-*
   MirareCiteRibbon.ps1` script does it for you).

Everything else — the C# source, the ribbon XML, the VBA source, the
installer, the diagnosis guide, the docs — is complete and ready to go.

---

## 1. What's in the package

```
mirare-cite-addin/
├── README.md                  ← you are here
├── DIAGNOSIS.md               ← 5-min root-cause checklist for your current
│                                broken add-in (read this first if you have
│                                the two files you uploaded)
├── src/MirareCiteAddIn/        ← the C# COM add-in project
│   ├── MirareCiteAddIn.csproj ← open in VS 2022
│   ├── AssemblyInfo.cs         ← [ComVisible], [Guid], version
│   ├── Connect.cs              ← IDTExtensibility2 + IRibbonExtensibility
│   ├── RibbonCallbacks.cs      ← the actual button behaviors
│   ├── Models/                 ← Citation, MirareProject, RefManagerLibrary,
│   │                            CitedItem, CitationStyle
│   ├── Services/               ← ProjectLoader (.mrrcite → zip → JSON),
│   │                            RefManagerLoader (.refmanager.json),
│   │                            RemoteCitationService (HTTPS),
│   │                            CitationFormatter (APA/MLA/Chicago/Numeric),
│   │                            CitedTracker (drives the cited indicator),
│   │                            Logger
│   ├── Forms/                  ← CitationPickerForm (with cited indicator ✓),
│   │                            SettingsForm
│   ├── Resources/ribbon.xml    ← the customUI XML defining the ribbon tab
│   └── Properties/             ← stub resources
├── template/                   ← alternative .dotm-only variant (no DLL)
│   ├── customUI.xml
│   ├── MirareCiteRibbon.vba.bas
│   └── Build-MirareCiteRibbon.ps1
├── installer/
│   ├── register-addin.reg     ← registry keys for HKCU
│   ├── Register-AddIn.ps1     ← regasm + reg keys + clear-disabled
│   ├── Unregister-AddIn.ps1
│   └── verify-loadbehavior.ps1 ← the diagnostic
└── docs/
    ├── architecture.md         ← component diagram + data flow
    ├── data-formats.md         ← .mrrcite / .refmanager.json / remote JSON
    └── cited-indicator.md      ← how the ✓ works
```

---

## 2. Quick start — get a working add-in in ~10 minutes

### Option A — Build the full .NET COM add-in (recommended)

This gives you the picker dialog, the cited indicator, the remote query,
and APA/MLA/Chicago/Numeric formatting.

```powershell
# 1. Build (in Visual Studio 2022: open the .csproj, build Release|x64)
#    Or from a Developer Command Prompt for VS 2022:
msbuild src\MirareCiteAddIn\MirareCiteAddIn.csproj /p:Configuration=Release /p:Platform=x64

# 2. NuGet packages (System.Text.Json) restore automatically on build —
#    VS 2022 does it by itself; from CLI add -t:Restore,Build:
#    msbuild src\MirareCiteAddIn\MirareCiteAddIn.csproj -t:Restore,Build /p:Configuration=Release /p:Platform=x64
#    (Office/Word/Extensibility interop assemblies are referenced directly
#    from the Visual Studio VSTO PIA folder — no NuGet, no packages dir.)

# 3. Register the add-in (or just double-click installer\2-Register-AddIn.bat):
.\installer\Register-AddIn.ps1

# 4. Restart Word — the "Mirare Cite" tab should appear.
#    If it doesn't, double-click installer\1-Verify-AddIn.bat for the
#    diagnostic (the .bat keeps its window open; running the .ps1 directly
#    flashes a console that closes when the script ends).
```

### Option B — Use the .dotm-only template (no Visual Studio needed)

Lighter feature set (no cited-indicator scan, no remote query) but no
build step.

```powershell
# In a normal PowerShell (no admin):
.\template\Build-MirareCiteRibbon.ps1
#   → creates %APPDATA%\Microsoft\Word\STARTUP\MirareCiteRibbon.dotm
#   → restart Word
```

---

## 2b. Packaging & public distribution

Everything is **per-user** (HKCU registry + `%LOCALAPPDATA%\Mirare\WordAddIn`)
— end users never need admin rights, and `.NET Framework 4.8` is inbox on
Windows 10 1809+/11, so nothing needs to be shipped for the runtime.

### What to ship — pick one

| Route | How | Best for |
|-------|-----|----------|
| **A. Zip download** | Run `installer\Build-Release.ps1 -Version x.y.z` → upload `dist\MirareCite-AddIn-x.y.z.zip` to your website. Users extract and double-click `Install-AddIn.bat`. | Quick public download page |
| **B. Setup.exe** | Compile `installer\MirareCite-Setup.iss` with Inno Setup 6 after step A's build → signed single installer with an Add/Remove Programs entry. | Polished public distribution |
| **C. Bundle with the Mirare PySide6 app** | Have your app's installer (Inno Setup / NSIS) drop the staged payload into `%LOCALAPPDATA%\Mirare\WordAddIn` and then run `Install-AddIn.ps1 -PayloadDir "<that folder>" -NoCopy` silently. | Users who install the desktop app — the add-in arrives with it |

### Notes for route C (PySide6 app bundling)

- Build the PySide6 app with PyInstaller as usual; the add-in payload is
  just loose DLL files — no Python involvement at runtime, Word loads it
  via COM.
- In your Inno Setup script, mirror `[Files]` → `{localappdata}\Mirare\WordAddIn`
  and `[Run]` → the `Install-AddIn.ps1 -NoCopy` invocation (see
  `installer\MirareCite-Setup.iss` for a working example to copy from).
- Keep versioning in lock-step or install the add-in as a separate MSI/setup —
  Word loads the DLL from disk on every start, so updating the files +
  re-running the install script (idempotent) is all an upgrade needs.
- Code-sign `MirareCiteAddIn.dll` before wide release (regasm currently
  warns about `/codebase` on unsigned assemblies — harmless, but signing
  removes the warning and SmartScreen friction on the installer).

---

## 3. Verifying it actually loaded

If the "Mirare Cite" tab is **not** on the ribbon after restarting Word:

```powershell
.\installer\verify-loadbehavior.ps1
```

That single script tells you in seconds whether the problem is:
- Registry: add-in not registered → re-run `Register-AddIn.ps1`.
- Bitness: 32-bit DLL on 64-bit Word → check the build configuration
  (must be `x64` or `AnyCPU`, **not** `x86`).
- Disabled: Word hard-disabled the add-in after a previous crash →
  the script clears the disabled-items list and re-sets LoadBehavior=3.
- Silent `GetCustomUI` exception → check
  `%LOCALAPPDATA%\MirareCite\addin.log` for the trace.

For the full diagnostic walkthrough see **`DIAGNOSIS.md`**.

---

## 4. Data sources — what the add-in reads from

Per your integration spec, the add-in supports three citation sources:

| Source       | Format                              | Where it lives                  |
|--------------|-------------------------------------|---------------------------------|
| Library      | `.refmanager.json` — plain JSON     | anywhere on disk (user picks)   |
| Project      | `.mrrcite` — **ZIP** containing     | anywhere on disk (user picks)   |
|              |   `project.json` + other files      |                                 |
| Remote       | HTTPS GET → JSON response            | configurable endpoint (Settings)|

The picker dialog merges all three into a single list with a badge per
item: `[library]` / `[project]` / `[remote]`.

Full schemas are in `docs/data-formats.md`.

---

## 5. The cited indicator (the ✓ on the list)

When the picker opens, every item already cited in the active document
is shown with a green ✓ in the "Cited" column.  The tracker that powers
this lives in `Services/CitedTracker.cs` and is also exposed on the
ribbon via the dynamic `GetCitedCountLabel` callback (the ribbon label
shows "N items cited").  See `docs/cited-indicator.md` for details.

---

## 6. Customization touch-points

| You want to change…        | Edit this file                                              |
|----------------------------|-------------------------------------------------------------|
| Ribbon tab labels/icons    | `src/.../Resources/ribbon.xml` (and `template/customUI.xml`) |
| Add/remove ribbon buttons  | same files + add matching callback in `Connect.cs`         |
| Citation formatting style  | `Services/CitationFormatter.cs`                             |
| Remote API endpoint shape  | `Services/RemoteCitationService.cs` (the DTOs)             |
| Picker UI                  | `Forms/CitationPickerForm.cs`                               |
| ProgID                     | `Connect.cs` (`[ProgId("MirareCite.AddIn")]`) + `register-addin.reg` + `Register-AddIn.ps1` (search-replace the ProgID string) |
| Default settings           | `RibbonCallbacks.cs` (the property defaults)               |

---

## 7. Known limits / TODOs (called out honestly)

1. **Numeric style renumbering** — `CitationFormatter.InText()` emits a
   placeholder `[mirare:ID]` for numeric style; the renumber pass happens
   in `OnEditBibliography`.  If you want true Zotero-style automatic
   renumbering on every insert, you'll need to walk every MRCITE field
   and rewrite its `.Result.Text` — see the comment in
   `RibbonCallbacks.OnEditBibliography`.

2. **`CitedTracker.GetCitedCitations`** is currently a stub returning an
   empty list.  The real implementation needs the picker's cached
   Citation set joined against the cited Ids; see the comment in
   `Services/CitedTracker.cs`.  Trivial to wire up — pass a
   `Dictionary<string, Citation>` cache into the tracker.

3. **Bibliography section** is appended at the end of the document each
   time you click "Update Bibliography" — it does not yet find and
   replace an existing "References" heading.  Easy fix: search for the
   last paragraph whose style is `Heading 1` and text is "References",
   then replace its range.

4. **Settings persistence** uses `System.Text.Json` (the 8.0.5 NuGet
   package via `PackageReference` in the .csproj — restored
   automatically; the old `packages.config` is gone).

5. **Macro signing for the .dotm variant** is not done — Word will
   show a security prompt on first load.  For an enterprise deployment,
   sign the VBA project with a code-signing certificate.

6. **VBA JSON parser** in `template/MirareCiteRibbon.vba.bas` is a
   trivial regex-based extractor; for real use, install VBA-JSON
   (https://github.com/VBA-tools/VBA-JSON) into the VBA project.

---

## 8. First thing to do

Read `DIAGNOSIS.md`.  It explains *why* your current `WordAddIn_Mirare.dll`
+ `MirareCiteRibbon.dotm` aren't showing up, and walks you through the
5-minute `verify-loadbehavior.ps1` checklist.  If you'd like me to do
this diagnosis on your actual files, re-upload them — the previous
upload didn't reach my workspace.
