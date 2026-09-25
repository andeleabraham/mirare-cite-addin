# Diagnosis: why your current `WordAddIn_Mirare.dll` + `MirareCiteRibbon.dotm` don't show on the ribbon

> The two files you uploaded did not actually reach my workspace (the upload
> folder was empty), so I cannot do a byte-level inspection. The list below is
> the **same checklist I would run** if I had the files. Run it yourself in
> about 5 minutes and you will almost certainly find the root cause in one of
> sections 1–4.

## 0. How a COM Word add-in is *supposed* to show up on the ribbon

A working COM add-in must satisfy **all four** conditions at once. If any one
fails, Word silently loads the DLL (or does not load it at all) and you get a
ribbon with no Mirare Cite tab — exactly your symptom.

```
┌──────────────────────────────────────────────────────────────────────────┐
│  1. Assembly is COM-visible + signed with a stable ProgID                  │
│        → [AssemblyGuid] + [ComVisible(true)] + [ProgID("MirareCite.AddIn")]│
│                                                                            │
│  2. Assembly is registered in the registry                                 │
│        → regasm /codebase MirareCiteAddIn.dll                              │
│        → HKCU\Software\Microsoft\Office\Word\Addins\MirareCite.AddIn        │
│             LoadBehavior = 3   (auto-load at startup)                      │
│                                                                            │
│  3. The Connect class implements IDTExtensibility2 *and*                  │
│     IRibbonExtensibility. GetCustomUI("Microsoft.Word.Word") returns the  │
│     customUI XML (not null, not empty).                                   │
│                                                                            │
│  4. The customUI XML is well-formed and every callback signature in it    │
│     exists as a public method on the Connect class with the right attr.   │
└──────────────────────────────────────────────────────────────────────────┘
```

If `MirareCiteRibbon.dotm` is supposed to provide the ribbon (instead of, or
in addition to, the DLL's `IRibbonExtensibility`), there is a *parallel*
checklist in §3.

---

## 1. Registry: is the add-in even registered? (most common cause)

Open PowerShell and run:

```powershell
# Replace MirareCite.AddIn with whatever ProgID your DLL uses.
# You can find the real ProgID by inspecting the DLL (see §5).
$key = "HKCU:\Software\Microsoft\Office\Word\Addins\MirareCite.AddIn"
Get-ItemProperty $key -ErrorAction SilentlyContinue
```

You should see **at least**:

| Value name     | Type   | Expected value |
|----------------|--------|-----------------|
| `LoadBehavior` | DWORD  | `3`             |
| `FriendlyName` | SZ     | `Mirare Cite`   |
| `Description`  | SZ     | (any text)      |

Interpret `LoadBehavior`:

| Value | Meaning | What it means for you |
|-------|---------|------------------------|
| 0     | Disabled | Add-in is registered but Word will not load it. |
| 1     | Loaded on demand | Not on the ribbon until you click something — usually invisible. |
| 2     | Loaded at startup, but only if user requests | Still effectively invisible. |
| **3** | **Loaded at startup (auto)** | ✅ This is what you want. |
| 9     | Tried and disabled due to load error | Word crash-disabled it. |
| 16    | Hard-disabled | Word blocked it after a crash. |

**If the key is missing entirely** → `regasm /codebase` was never run, or it
was run with the wrong bitness (see §2). Re-run the installer in this package
(`installer\Register-AddIn.ps1`) to fix this in one shot.

**If `LoadBehavior` is 9 or 16** → Word has *disabled* the add-in. Even if you
fix the underlying bug and re-register, Word will refuse to load it. Clear the
disabled-items list:

```powershell
# Word 2016 = 16.0
Remove-Item "HKCU:\Software\Microsoft\Office\16.0\Word\Resiliency\DisabledItems\*" `
  -ErrorAction SilentlyContinue
Remove-Item "HKCU:\Software\Microsoft\Office\16.0\Word\Resiliency\CrashingAddinList\*" `
  -ErrorAction SilentlyContinue
# Then set LoadBehavior back to 3:
New-ItemProperty -Path $key -Name LoadBehavior -Value 3 -PropertyType DWord -Force
```

Then restart Word.

---

## 2. Bitness mismatch (second most common cause)

You said you're on **64-bit Word**. A COM add-in DLL must match the host
bitness. If your `WordAddIn_Mirare.dll` was compiled for x86 (32-bit), Word
64-bit will *not* even try to load it — and will not give you an error
message.

Check Word's bitness from PowerShell:

```powershell
$word = New-Object -ComObject Word.Application
"$($word.Version) - $($word.Build)"
Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Office\16.0\Word\InstallRoot" `
  -Name "InstallPath" -ErrorAction SilentlyContinue
# Word's bitness matches the Office install:
(Get-Item "C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE").Length
# vs
(Get-Item "C:\Program Files (x86)\Microsoft Office\root\Office16\WINWORD.EXE").Length
```

If Word is 64-bit:
- The DLL must be built **x64 or AnyCPU** (AnyCPU on .NET Framework 4.8 is fine).
- `regasm` must be the **64-bit** `regasm` (`C:\Windows\Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe`).
- The registry key must land in `HKCU\Software\Microsoft\Office\Word\Addins\`
  (NOT under `HKCU\Software\Wow6432Node\...`, which is where 32-bit regasm
  puts it).

This single mistake — running 32-bit `regasm` against a 32-bit DLL when
Word is 64-bit — accounts for the majority of "my add-in doesn't show up"
cases. The fix in this package's `Register-AddIn.ps1` always uses the
Framework64 regasm.

---

## 3. If the `.dotm` is supposed to provide the ribbon

A `.dotm` does **not** need to be a COM add-in at all — it's a Word template
with VBA macros and an embedded `customUI.xml`. For its ribbon tab to appear:

1. It must be loaded as a **global template** — i.e. placed in:
   ```
   %APPDATA%\Microsoft\Word\STARTUP\MirareCiteRibbon.dotm
   ```
   Or attached via `Developer → Word Add-ins → Add…`. If it just sits in
   your Documents folder, Word will never load it.

2. The `customUI.xml` inside the `.dotm` must be at the path:
   ```
   customUI\customUI.xml
   ```
   inside the `.docm`/`.dotm` ZIP package, and the relationship part
   `_rels/.rels` must reference it. If it's at the wrong path inside the
   ZIP, Word will silently ignore it.

3. The VBA project must be **trusted**. If macro security is set to
   "Disable all macros without notification" (the default in some
   enterprises), the template loads but the VBA project doesn't run, and
   no callbacks fire even though the ribbon tab shows. Set to "Disable
   all macros except digitally signed macros" and sign the VBA project,
   OR set to "Notifications for all macros" and accept the prompt.

4. Every callback name in the `customUI.xml` must exist as a `Public Sub`
   in a standard module (not a class module) of the VBA project.

You can crack open your existing `.dotm` to inspect these without Word —
rename to `.zip` and double-click:

```
MirareCiteRibbon.dotm
  ├── [Content_Types].xml
  ├── _rels\.rels                  ← must contain a Relationship to customUI.xml
  ├── customUI\customUI.xml        ← MUST EXIST
  ├── word\
  │   ├── document.xml             ← tiny stub
  │   ├── vbaProject.bin           ← the actual VBA code (CFB compound file)
  │   └── ...
  └── docProps\
```

If any of these are missing or the relationship is missing, that's your bug.
The `Build-MirareCiteRibbon.ps1` in this package produces a `.dotm` with
the correct structure.

---

## 4. The DLL is registered, but Word still doesn't show the tab

If §1, §2, §3 are all correct, the next-most-common cause is **a silent
exception in `OnConnection`** or in `GetCustomUI`. Word swallows both and
you get no error message.

Get the real error from the Windows event log:

```powershell
Get-WinEvent -LogName Application -MaxEvents 200 `
  | Where-Object { $_.ProviderName -match "Office|Word|\.NET|Application" } `
  | Format-Table TimeCreated, Id, LevelDisplayName, Message -AutoSize
```

Or enable Word's own add-in logging:

```powershell
# HKCU\Software\Microsoft\Office\16.0\Word\Options
New-ItemProperty "HKCU:\Software\Microsoft\Office\16.0\Word\Options" `
  -Name "EnableAddInLogging" -Value 1 -PropertyType DWord -Force
```

Restart Word, reproduce, then check the same Application event log.

If `GetCustomUI` is throwing, the most common reasons are:
- Returning `null` instead of the XML string.
- Returning the XML for the wrong `ribbonID` (must be the literal string
  `"Microsoft.Word.Word"`, case-sensitive).
- The XML is missing the `xmlns` namespace declaration:
  `http://schemas.microsoft.com/office/2009/07/customui` (Word 2010+).

---

## 5. Inspect your existing DLL without source

You can disassemble the existing `WordAddIn_Mirare.dll` to see what ProgID
it declared and whether it even implements `IRibbonExtensibility`. On
Windows with .NET SDK:

```powershell
# List the assembly's public types andProgIDs
& "C:\Program Files (x86)\Microsoft SDKs\Windows\v10.0A\bin\NETFX 4.8 Tools\ildasm.exe" `
  WordAddIn_Mirare.dll /TEXT /NOBAR /CLASSLIST
```

Look for:
- `Extensibility.IDTExtensibility2`
- `Microsoft.Office.Core.IRibbonExtensibility` (PIA reference)

If `IRibbonExtensibility` is **missing**, that explains everything: the
DLL is a plain COM automation server, not a ribbon add-in — it can never
show a ribbon tab on its own. You need either a new DLL (the one in this
package) or the `.dotm` path (§3).

If `IDTExtensibility2` is **missing**, the DLL isn't a Word add-in at all —
it's probably a helper library used by the `.dotm`. That's fine, but then
the ribbon comes from the `.dotm` only, and the bug is in §3.

---

## 6. Quick "does Word even see it?" check

In Word, go to **File → Options → Add-ins**. At the bottom, set *Manage:*
to **COM Add-ins** and click *Go…*. Look for `Mirare Cite` (or whatever the
FriendlyName is):

- **Not listed** → §1 or §2 (registry/bitness).
- **Listed but unchecked** → check the box; if it stays unchecked after
  restart, §1 has `LoadBehavior = 9` or `16` (disabled).
- **Checked but still no ribbon tab** → §4 or §5 (GetCustomUI throws / not
  implemented).

---

## 7. What to send me if you want me to do this analysis for you

If you re-upload the files, I can produce the exact diagnosis (not just the
checklist). I'd need:

1. `WordAddIn_Mirare.dll`
2. `MirareCiteRibbon.dotm`
3. The output of this PowerShell snippet (one line, paste it back):

```powershell
@{
  WordVersion = (New-Object -ComObject Word.Application).Version
  WordExe     = if (Test-Path "C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE") {"x64"}
                elseif (Test-Path "C:\Program Files (x86)\Microsoft Office\root\Office16\WINWORD.EXE") {"x86"}
                else {"unknown"}
  AddinKeys   = (Get-ChildItem "HKCU:\Software\Microsoft\Office\Word\Addins" -ErrorAction SilentlyContinue).Name
  Disabled    = (Get-ChildItem "HKCU:\Software\Microsoft\Office\16.0\Word\Resiliency\DisabledItems" -ErrorAction SilentlyContinue).Name
} | ConvertTo-Json
```

That single JSON object will tell me in seconds whether the problem is
registration, bitness, Word-disabled, or GetCustomUI.
