# Data formats

The add-in reads three kinds of citation sources.  All three are JSON
(or JSON-in-ZIP) and all three deserialise into the same
`List<Citation>` shape so the picker treats them uniformly.

## 1. `.mrrcite` project file

A `.mrrcite` is a **ZIP archive**.  You can rename it to `.zip` and
double-click it to see the contents.

### Required structure

```
my_project.mrrcite  (ZIP)
├── project.json                    ← REQUIRED
├── refs.refmanager.json            ← optional embedded library copy
├── notes/
│   └── *.md
├── attachments/
│   └── *.pdf
└── project.json.meta               ← optional checksum / signature
```

The add-in only reads `project.json` and (if present)
`refs.refmanager.json`.  The other files are user content the add-in
ignores.

### `project.json` schema

```json
{
  "projectId": "uuid-string",
  "name": "PTEN regulation study",
  "createdAt": "2025-09-01T10:00:00Z",
  "lastModified": "2025-09-20T14:30:00Z",
  "metadata": {
    "owner": "Jane Doe",
    "description": "miR-21 / PTEN validation cohort"
  },
  "citations": [
    {
      "id": "mrc_001",
      "title": "miR-21 targets PTEN in breast cancer cell lines",
      "authors": ["Doe, Jane", "Smith, John"],
      "year": 2022,
      "journal": "RNA Biology",
      "doi": "10.1080/15476286.2022.123456",
      "url": "https://doi.org/10.1080/...",
      "miRNA": "hsa-miR-21-5p",
      "targetGene": "PTEN",
      "evidenceType": "validated",
      "sourceDb": "miRTarBase",
      "userNote": "Looking at the MCF7 cohort specifically",
      "tags": ["breast", "PTEN"]
    }
  ]
}
```

### `refs.refmanager.json` schema

Same as the standalone library format in §2 below.  When embedded
inside a `.mrrcite`, the add-in merges it into the picker list tagged
with `Origin = Project` (because the user opened the project, the
library is now contextual to that project).

## 2. `.refmanager.json` library file

A plain JSON file (not zipped) — the primary "library" the user cites
from across all projects.

```json
{
  "libraryId": "uuid-string",
  "name": "My Mirare library",
  "entries": [
    {
      "id": "lib_001",
      "title": "Comprehensive analysis of miR-21 targets",
      "authors": ["Frankel, LB", "Lund, AH"],
      "year": 2021,
      "journal": "Nature Reviews Molecular Cell Biology",
      "doi": "10.1038/s41580-021-00360-x",
      "url": "https://doi.org/10.1038/...",
      "miRNA": "hsa-miR-21-5p",
      "targetGene": "PDCD4",
      "evidenceType": "validated",
      "sourceDb": "TarBase"
    }
  ]
}
```

## 3. Remote HTTPS API

The remote query is a single GET with a `q=…` query parameter:

```
GET  {endpoint}?q=miR-21+PTEN
```

Expected response:

```json
{
  "results": [
    {
      "id": "rem_001",
      "title": "Remote miR-21 PTEN record",
      "authors": ["…"],
      "year": 2024,
      "journal": "…",
      "doi": "…",
      "url": "…",
      "miRNA": "hsa-miR-21-5p",
      "targetGene": "PTEN",
      "evidenceType": "predicted",
      "sourceDb": "TargetScan"
    }
  ]
}
```

If your real endpoint uses a different shape (e.g. POST, auth header,
pagination), edit `Services/RemoteCitationService.cs` — the only thing
that matters is that the method returns `List<Citation>`.

## 4. The MRCITE Word field (what gets written into the document)

When the user inserts a citation, the add-in writes a Word field:

- **Type:** `wdFieldAddin` (literal enum value `81`)
- **Code (hidden):** `MRCITE id=abc123 style=apa`
- **Result (visible):** `(Smith et al., 2022)` (or whatever the formatter produces)

The `MRCITE` prefix is what `CitedTracker` greps for when re-scanning
the document.  Don't change it without updating the tracker's
`ParseFieldCode` too.

## 5. Settings file

Location: `%APPDATA%\MirareCite\settings.json`

```json
{
  "Style": "Apa",
  "RemoteEndpoint": "https://api.mirare.example.org/cite",
  "LastLibraryPath": "C:\\Users\\jane\\Documents\\library.refmanager.json",
  "LastProjectPath": "C:\\Users\\jane\\Documents\\study.mrrcite"
}
```

## 6. Log file

Location: `%LOCALAPPDATA%\MirareCite\addin.log`

Plain text, append-only.  One line per entry:

```
2026-09-25 09:00:00.123 [INFO ] OnConnection: mode=ext_cm_Startup
2026-09-25 09:00:00.150 [INFO ] GetCustomUI ribbonID=Microsoft.Word.Word
2026-09-25 09:00:00.170 [INFO ] OnStartupComplete — add-in fully loaded
```

If something goes wrong, this file is the first place to look —
exceptions in `OnConnection` and `GetCustomUI` are silently swallowed
by Word otherwise.
