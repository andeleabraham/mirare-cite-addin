# The cited indicator

The cited indicator is the green ✓ that appears next to any item in the
picker list when that item is **already cited in the active document**.

You asked for this specifically ("Additionally, cited indicator on the
list") so this doc explains how it works end-to-end.

## How it works

```
┌─────────────────────────────────────────────────────────────────┐
│  Active Word document                                            │
│  ┌──────────────────────────────────────────────────────────┐  │
│  │  …  { MRCITE id=abc123 style=apa }(Smith et al., 2022) … │  │
│  │  …  { MRCITE id=def456 style=apa }(Frankel et al., 2021) │  │
│  │  …                                                       │  │
│  └──────────────────────────────────────────────────────────┘  │
│              │                                                   │
└──────────────┼───────────────────────────────────────────────────┘
               │ CitedTracker.ScanDocument(doc)
               ▼
       ┌─────────────────────┐
       │  HashSet<string>    │
       │  { "abc123",         │
       │    "def456" }        │
       └─────────────────────┘
               │ passed into CitationPickerForm constructor
               ▼
   ┌───────────────────────────────────────────────────────────┐
   │  CitationPickerForm                                       │
   │  ┌─────────────────────────────────────────────────────┐  │
   │  │ Cited │ Origin │ Year │ Title / Authors / miRNA      │  │
   │  ├───────┼─────────┼──────┼─────────────────────────────┤  │
   │  │  ✓    │ library │ 2022 │ miR-21 PTEN …               │  │
   │  │       │ project │ 2021 │ Compre- hensive analysis…    │  │
   │  │  ✓    │ library │ 2021 │ Frankel et al. — PDCD4…     │  │
   │  └─────────────────────────────────────────────────────┘  │
   └───────────────────────────────────────────────────────────┘
```

## Key code locations

| What                                | Where                                                   |
|-------------------------------------|---------------------------------------------------------|
| Scan all fields in all stories      | `Services/CitedTracker.cs` → `ScanDocument`             |
| Parse `MRCITE id=…` field codes     | `Services/CitedTracker.cs` → `ParseFieldCode`           |
| Build the cited-id set              | `Services/CitedTracker.cs` → `_citedIds`                |
| Render the ✓ in the list view       | `Forms/CitationPickerForm.cs` → `RenderList`            |
| Trigger re-scan on doc change       | `Connect.cs` → `OnDocumentChange` / `OnDocumentOpen`    |
| Update the dynamic ribbon label     | `RibbonCallbacks.cs` → `GetCitedCountLabel`             |

## Why use `wdFieldAddin` fields

We store the citation marker as a Word field, **not** as plain text,
for two reasons:

1. **Stable across saves.**  Word preserves `wdFieldAddin` fields
   through `.docx` save cycles without trying to re-evaluate them.
2. **Recoverable on re-scan.**  The field code (`MRCITE id=… style=…`)
   survives round-trips, so `CitedTracker` can re-derive the cited set
   even if the user pastes content from another document, deletes
   citations, etc.

## Limitations (called out honestly)

1. **No invalidation callback wired yet in the .NET variant.**  The
   dynamic ribbon label `mrcCitedCount` is currently only refreshed
   when Office invalidates it (e.g. on tab switch).  To get a live
   update after each insertion, store the `IRibbonUI` pointer (via an
   `onLoad="OnRibbonLoad"` callback in `ribbon.xml`) and call
   `ribbon.InvalidateControl("mrcCitedCount")` from
   `RibbonCallbacks` after each `AddCited`.

2. **`CitedTracker.GetCitedCitations` is currently a stub** returning
   an empty list, so `OnEditBibliography` produces a stub bibliography.
   To wire it up, hold a `Dictionary<string, Citation>` cache of every
   citation the user has ever picked this session and join it against
   `_citedIds` inside `GetCitedCitations`.  ~10 lines of code.

3. **Field update on style change.**  If the user changes the citation
   style via Settings after citations are already inserted, the
   visible text of those fields will NOT auto-update.  You'd need a
   "refresh all fields" button that walks every MRCITE field, parses
   the cached Id+style, looks up the original Citation, and rewrites
   `field.Result.Text = formatter.InText(citation)`.  Not implemented
   in this skeleton.

## How to test it manually

1. Open a blank Word document.
2. Click **Insert Citation** on the Mirare Cite tab.
3. Pick two items; they're now in the document with `(Author, Year)` in-text.
4. Click **Insert Citation** again — the two items you just inserted
   now show a green ✓ in the picker list.
5. Delete one of the citations in the document (select it, press Del).
6. Click **Refresh Cited Indicator** on the ribbon.  The picker should
   now show only one ✓.

If step 6 doesn't work, check `%LOCALAPPDATA%\MirareCite\addin.log`
for a `ScanDocument` line — if it shows `cited=0` after a deletion,
the deletion didn't actually remove the field (Word sometimes leaves
orphaned fields; select-and-delete is safer than pressing Backspace).
