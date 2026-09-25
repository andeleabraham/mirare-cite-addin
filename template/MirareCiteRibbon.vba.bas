'============================================================================
'  MirareCiteRibbon.vba.bas
'  ============================================================================
'  VBA module that backs the .dotm-only variant of the Mirare Cite add-in.
'
'  If you do not want to compile a .NET COM add-in (the .dll variant in
'  src/), this module + customUI.xml provide a ribbon tab using pure VBA.
'
'  To embed this in a .dotm on Windows:
'      1. Run template\Build-MirareCiteRibbon.ps1  (recommended)
'         OR
'      2. Manually:
'         - Open Word, create a new document, save as .dotm.
'         - Press Alt+F11 → VBA editor → Insert → Module.
'         - Paste this entire file content.
'         - Rename the module to "MirareCiteRibbon" (Properties window).
'         - Save. Place the .dotm in:
'               %APPDATA%\Microsoft\Word\STARTUP\
'         - Restart Word — the "Mirare Cite" tab appears.
'
'  This VBA variant does NOT include the cited-indicator logic or the
'  remote HTTPS query — those require the .NET COM add-in.  See README.md
'  for the full feature comparison.
'============================================================================

Option Explicit

' ---- Ribbon handle (for invalidating dynamic labels) ----------------------
Public gRibbon As IRibbonUI

' ---- Settings (persisted to %APPDATA%\MirareCite\settings.json) ------------
Public gDefaultStyle As String
Public gRemoteEndpoint As String

' ---- Cited id cache (scanned from the active document) --------------------
Public gCitedIds As Collection

'============================================================================
'  OnRibbonLoad — Office's ribbon engine hands us the IRibbonUI pointer.
'  Stash it globally so we can call Invalidate() later to refresh dynamic
'  controls (the cited-count labelControl).
'============================================================================
Public Sub OnRibbonLoad(ribbon As IRibbonUI)
    Set gRibbon = ribbon
    LoadSettings
    Set gCitedIds = New Collection
    RefreshCitedIndicator
End Sub

'============================================================================
'  onAction callbacks — one per button in customUI.xml.
'============================================================================
Public Sub OnInsertCitation(control As IRibbonControl)
    MsgBox "Insert Citation (All sources) — picker dialog not implemented in the VBA-only variant." & vbCrLf & _
           "For full functionality (picker, remote query, cited indicator), use the .NET COM add-in (see src\).", _
           vbInformation, "Mirare Cite"
End Sub

Public Sub OnInsertFromLibrary(control As IRibbonControl)
    Dim path As String
    path = PickFileOpen("RefManager library (*.refmanager.json)|*.refmanager.json")
    If Len(path) = 0 Then Exit Sub
    InsertFromJsonLibrary path
End Sub

Public Sub OnInsertFromProject(control As IRibbonControl)
    Dim path As String
    path = PickFileOpen("Mirare Cite project (*.mrrcite)|*.mrrcite")
    If Len(path) = 0 Then Exit Sub
    MsgBox "Project (.mrrcite) parsing requires ZIP+JSON support not available in pure VBA." & vbCrLf & _
           "Use the .NET COM add-in.", vbInformation, "Mirare Cite"
End Sub

Public Sub OnInsertFromRemote(control As IRibbonControl)
    MsgBox "Remote HTTPS query requires the .NET COM add-in (XMLHTTP can be used, but parsing + UI is cleaner in .NET).", _
           vbInformation, "Mirare Cite"
End Sub

Public Sub OnRefreshCitedIndicator(control As IRibbonControl)
    RefreshCitedIndicator
    If Not gRibbon Is Nothing Then gRibbon.InvalidateControl "mrcCitedCount"
End Sub

Public Sub OnEditBibliography(control As IRibbonControl)
    ' Append a "References" section at the end of the active document with
    ' every cited item (basic — VBA variant only).
    AppendReferencesSection
End Sub

Public Sub OnSettings(control As IRibbonControl)
    Dim s As String
    s = InputBox("Remote endpoint URL:", "Mirare Cite Settings", gRemoteEndpoint)
    If StrPtr(s) <> 0 Then
        gRemoteEndpoint = s
        SaveSettings
    End If
End Sub

'============================================================================
'  getLabel callback for the dynamic cited-count labelControl.
'============================================================================
Public Sub GetCitedCountLabel(control As IRibbonControl, ByRef returnedVal)
    If gCitedIds Is Nothing Then
        returnedVal = "0 items cited"
    ElseIf gCitedIds.Count = 1 Then
        returnedVal = "1 item cited"
    Else
        returnedVal = gCitedIds.Count & " items cited"
    End If
End Sub

'============================================================================
'  CitedTracker — scans the active document for MRCITE fields and rebuilds
'  gCitedIds.  Word's wdFieldAddin (type 81) is what the .NET add-in inserts.
'============================================================================
Private Sub RefreshCitedIndicator()
    On Error GoTo done
    If ActiveDocument Is Nothing Then Exit Sub
    Set gCitedIds = New Collection

    Dim story As Range
    For Each story In ActiveDocument.StoryRanges
        WalkStory story
    Next story
done:
End Sub

Private Sub WalkStory(storyRange As Range)
    Dim r As Range
    Set r = storyRange
    Do While Not r Is Nothing
        If r.Fields.Count > 0 Then
            Dim f As Field
            For Each f In r.Fields
                If f.Type = 81 Then  ' wdFieldAddin
                    Dim code As String
                    code = f.Code.Text
                    If Left(code, 7) = "MRCITE " Then
                        Dim id As String
                        id = ParseIdFromCode(code)
                        If Len(id) > 0 Then
                            On Error Resume Next
                            gCitedIds.Add id, id
                            On Error GoTo 0
                        End If
                    End If
                End If
            Next f
        End If
        Set r = r.NextStoryRange
    Loop
End Sub

Private Function ParseIdFromCode(code As String) As String
    Dim tokens() As String, i As Long
    tokens = Split(code, " ")
    For i = LBound(tokens) To UBound(tokens)
        If Left(tokens(i), 3) = "id=" Then
            ParseIdFromCode = Mid(tokens(i), 4)
            Exit Function
        End If
    Next i
    ParseIdFromCode = ""
End Function

'============================================================================
'  Library loader (.refmanager.json) — minimal: parses entries with the
'  Scripting.FileSystemObject + a tiny JSON tokenizer.  For real use, use
'  VBA-JSON (https://github.com/VBA-tools/VBA-JSON) by adding it to your
'  VBA project references.
'============================================================================
Private Sub InsertFromJsonLibrary(path As String)
    ' STUB: prompt the user to type the citation ID they want.
    ' (Full JSON parsing in VBA requires a 3rd-party module — see README.)
    Dim id As String
    id = InputBox("Type the citation Id to insert from " & path & ":", "Mirare Cite")
    If Len(id) = 0 Then Exit Sub
    InsertMrciteField id, gDefaultStyle
End Sub

Private Sub InsertMrciteField(id As String, style As String)
    Dim sel As Selection
    Set sel = Application.Selection
    Dim f As Field
    Set f = ActiveDocument.Fields.Add(sel.Range, 81, "MRCITE id=" & id & " style=" & style, False)
    f.Result.Text = "(Mirare " & id & ")"
    RefreshCitedIndicator
    If Not gRibbon Is Nothing Then gRibbon.InvalidateControl "mrcCitedCount"
End Sub

Private Sub AppendReferencesSection()
    Dim rng As Range
    Set rng = ActiveDocument.Content
    rng.Collapse Direction:=wdCollapseEnd
    rng.InsertParagraphAfter
    rng.InsertAfter "References" & vbCrLf
    Dim id As Variant
    For Each id In gCitedIds
        rng.InsertAfter "- [" & id & "]" & vbCrLf
    Next id
End Sub

'============================================================================
'  Helpers — file open dialog, settings persistence
'============================================================================
Private Function PickFileOpen(filter As String) As String
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogOpen)
    With fd
        .Filters.Clear
        Dim parts() As String, p As Variant
        parts = Split(filter, "|")
        Dim i As Long
        For i = LBound(parts) To UBound(parts) - 1 Step 2
            .Filters.Add parts(i), parts(i + 1)
        Next i
        .AllowMultiSelect = False
        If .Show = -1 Then PickFileOpen = .SelectedItems(1) Else PickFileOpen = ""
    End With
End Function

Private Function SettingsPath() As String
    SettingsPath = Environ$("APPDATA") & "\MirareCite\settings.json"
End Function

Private Sub LoadSettings()
    On Error Resume Next
    Dim s As String
    s = ReadFile(SettingsPath)
    If Len(s) > 0 Then
        gDefaultStyle = ExtractJsonField(s, "Style")
        gRemoteEndpoint = ExtractJsonField(s, "RemoteEndpoint")
    End If
    If Len(gDefaultStyle) = 0 Then gDefaultStyle = "Apa"
End Sub

Private Sub SaveSettings()
    On Error Resume Next
    Dim s As String
    s = "{""Style"":""" & gDefaultStyle & """,""RemoteEndpoint"":""" & gRemoteEndpoint & """}"
    MkDir parentDir(SettingsPath)
    WriteFile SettingsPath, s
End Sub

Private Function parentDir(path As String) As String
    Dim i As Long
    i = InStrRev(path, "\")
    parentDir = Left(path, i - 1)
End Function

Private Function ReadFile(path As String) As String
    On Error GoTo done
    Dim h As Integer: h = FreeFile
    Open path For Input As h
    Dim s As String: s = Input$(LOF(h), h)
    Close h
    ReadFile = s
done:
End Function

Private Sub WriteFile(path As String, contents As String)
    On Error Resume Next
    Dim h As Integer: h = FreeFile
    Open path For Output As h
    Print #h, contents;
    Close h
End Sub

'  Trivial JSON field extractor — returns the value for a top-level key.
'  Replace with VBA-JSON for production use.
Private Function ExtractJsonField(json As String, key As String) As String
    Dim needle As String: needle = """" & key & """:"""
    Dim i As Long: i = InStr(json, needle)
    If i = 0 Then Exit Function
    i = i + Len(needle)
    Dim j As Long: j = InStr(i, json, """")
    If j > i Then ExtractJsonField = Mid(json, i, j - i)
End Function
