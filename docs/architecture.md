# Architecture

## Component diagram

```
┌─────────────────────────────────────────────────────────────────────────┐
│                         Microsoft Word 2016 (x64)                       │
│                                                                         │
│   ┌──────────────────────────────────────────────────────────────┐    │
│   │   Ribbon ("Mirare Cite" tab)                                  │    │
│   │                                                               │    │
│   │   Office's ribbon engine reads customUI XML from the DLL     │    │
│   │   (IRibbonExtensibility.GetCustomUI("Microsoft.Word.Word"))   │    │
│   │   and late-binds callbacks by name to Connect.cs methods.    │    │
│   └──────────────────────────────────────────────────────────────┘    │
│                                  │                                       │
│                                  ▼                                       │
│   ┌──────────────────────────────────────────────────────────────┐    │
│   │   Connect : IDTExtensibility2, IRibbonExtensibility           │    │
│   │                                                               │    │
│   │   OnConnection  → capture Word.Application                    │    │
│   │   GetCustomUI   → return embedded ribbon.xml                  │    │
│   │   OnInsertCitation / OnRefreshCitedIndicator / ...           │    │
│   │       → delegate to RibbonCallbacks                           │    │
│   └──────────────────────────────────────────────────────────────┘    │
│                                  │                                       │
│                                  ▼                                       │
│   ┌──────────────────────────────────────────────────────────────┐    │
│   │   RibbonCallbacks                                             │    │
│   │                                                               │    │
│   │   - holds CitedTracker + logger + settings                    │    │
│   │   - opens CitationPickerForm                                  │    │
│   │   - inserts MRCITE Word fields at the cursor                  │    │
│   └──────────────────────────────────────────────────────────────┘    │
│            │                  │                  │                   │
│            ▼                  ▼                  ▼                   │
│   ┌──────────────┐   ┌────────────────┐   ┌───────────────────┐     │
│   │ Citation-    │   │ CitedTracker   │   │ CitationFormatter │     │
│   │ PickerForm   │   │                │   │ (APA/MLA/Chi/Num) │     │
│   │              │   │ Scans every    │   │                   │     │
│   │ Shows list   │   │ wdFieldAddin   │   │ InText(c)         │     │
│   │ with cited ✓ │   │ in every Story │   │ BuildBibliography │     │
│   └──────────────┘   └────────────────┘   └───────────────────┘     │
│            │                                                         │
│            ▼                                                         │
│   ┌──────────────────────────────────────────────────────────────┐    │
│   │   Loaders (read-only — never write to source files)          │    │
│   │                                                               │    │
│   │   ProjectLoader   (.mrrcite = ZIP → JSON)                     │    │
│   │   RefManagerLoader (.refmanager.json)                        │    │
│   │   RemoteCitationService (HTTPS → JSON)                        │    │
│   │                                                               │    │
│   │   All three produce List<Citation> with Origin set, so the   │    │
│   │   picker and formatter treat them uniformly.                 │    │
│   └──────────────────────────────────────────────────────────────┘    │
└─────────────────────────────────────────────────────────────────────────┘
                                  │
                                  ▼
   ┌──────────────────────────────────────────────────────────────┐
   │   File system / network                                      │
   │                                                              │
   │   .mrrcite  → ZIP → project.json + refs.refmanager.json     │
   │   .refmanager.json  → plain JSON                            │
   │   Remote HTTPS endpoint → JSON                              │
   │                                                              │
   │   Settings:    %APPDATA%\MirareCite\settings.json            │
   │   Log:         %LOCALAPPDATA%\MirareCite\addin.log           │
   └──────────────────────────────────────────────────────────────┘
```

## Insertion pipeline (what happens when you click Insert Citation)

1. **User clicks** "Insert Citation" on the ribbon.
2. Office's ribbon engine calls `Connect.OnInsertCitation(IRibbonControl)`.
3. `Connect` delegates to `RibbonCallbacks.OnInsertCitation`.
4. `RibbonCallbacks` calls `CitedTracker.ScanDocument(ActiveDocument)` to
   build the set of already-cited IDs.
5. `RibbonCallbacks` opens `CitationPickerForm` with that set.
6. The form loads citations via the appropriate Loader(s) depending on
   the selected Source scope.
7. The user picks an item, clicks Insert (or double-clicks the row).
8. `RibbonCallbacks` calls
   `doc.Fields.Add(selection, wdFieldAddin, "MRCITE id=… style=…")`
   and sets `field.Result.Text = formatter.InText(citation)`.
9. `CitedTracker.AddCited(id)` updates the in-memory cited set.
10. The picker dialog closes.

## Ribbon refresh mechanism

The dynamic label "N items cited" updates when Office invalidates the
control.  Currently the invalidation happens implicitly on the next
DocumentChange event; if you want a live update after each insertion,
hold the `IRibbonUI` pointer (already in the VBA variant — see
`template/MirareCiteRibbon.vba.bas` `OnRibbonLoad`) and call
`ribbon.InvalidateControl("mrcCitedCount")` from `RibbonCallbacks`.

To wire this up in the .NET add-in: store the `IRibbonUI` from an
`onLoad="OnRibbonLoad"` callback declared in `ribbon.xml`, then call
`InvalidateControl` whenever the tracker changes.  This is a 3-line
change — see `Connect.cs` comment.
