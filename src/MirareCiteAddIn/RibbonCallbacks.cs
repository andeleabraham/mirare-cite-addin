// ============================================================================
//  RibbonCallbacks.cs — the actual behavior for each ribbon button.
//
//  Connect.cs is just the COM plumbing; this file is where the add-in
//  actually does things.  Every method here is invoked late-bound by
//  Office's ribbon engine when the user clicks a button.  The signatures
//  MUST match the ribbon.xml callback contracts (single IRibbonControl
//  parameter for onAction, returning string for getLabel, etc.).
// ============================================================================

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;
using Newtonsoft.Json;
// Alias for the qualified spellings below (Word.Application); the plain
// using above stays so Document/Range/Field resolve unqualified.
using Word = Microsoft.Office.Interop.Word;
using MirareCiteAddIn.Forms;
using MirareCiteAddIn.Models;
using MirareCiteAddIn.Services;

namespace MirareCiteAddIn
{
    public class RibbonCallbacks
    {
        private readonly Word.Application _word;
        private readonly Logger _log;
        private readonly CitedTracker _tracker;

        // Every citation picked this session, keyed by Id — Update
        // Bibliography joins this against the document's cited ids.
        // (MRCITE fields only store the id, so the full record has to come
        // from somewhere; surviving Word restarts is a future improvement.)
        private readonly Dictionary<string, Citation> _citationCache =
            new Dictionary<string, Citation>();

        // Defaults — overridable via Settings dialog (stored in
        // %APPDATA%\MirareCite\settings.json).
        public CitationStyle Style { get; set; } = CitationStyle.Apa;
        public string RemoteEndpoint { get; set; } = "https://api.mirare.example.org/cite";
        public string LastLibraryPath { get; set; } = "";
        public string LastProjectPath { get; set; } = "";

        /// <summary>.csl file used when Style == Csl — same style files the
        /// Mirare desktop app keeps in its styles folder.</summary>
        public string CslStylePath { get; set; } = "";

        /// <summary>The active formatter: CSL-file driven when a style file
        /// is configured, otherwise one of the built-in styles.</summary>
        private CitationFormatter CreateFormatter()
            => Style == CitationStyle.Csl && !string.IsNullOrEmpty(CslStylePath)
                ? new CitationFormatter(CslStylePath)
                : new CitationFormatter(Style);

        /// <summary>Full display text for a citation field: the joined
        /// in-text citations, wrapped in parentheses for the built-in
        /// author-date styles. CSL layouts carry their own brackets/parens
        /// (from the style's layout prefix/suffix), numeric styles use
        /// [brackets] — so no extra wrapping there. The whole text lives
        /// INSIDE the field, Zotero/Mendeley-style, so Word treats the
        /// entire "(Author, Year)" as one field.</summary>
        private string RenderCitationDisplay(IList<Citation> citations, CitationStyle style,
                                             CitationFormatter formatter)
            => formatter.RenderGroup(citations);

        /// <summary>Injected by Connect.OnRibbonLoad — lets us re-activate
        /// the tab after a modal dialog collapses it.</summary>
        public IRibbonUI RibbonUi { get; set; }

        public RibbonCallbacks(Word.Application word, Logger log)
        {
            _word = word;
            _log = log;
            _tracker = new CitedTracker(log);
            LoadSettings();
        }

        // ─────────────────────────────────────────────────────────────────
        //  Word collapses a temporarily-expanded ribbon tab as soon as a
        //  modal dialog steals focus, so after every dialog we re-activate
        //  our tab — otherwise the user has to click the tab again.
        // ─────────────────────────────────────────────────────────────────
        private void RestoreRibbonTab()
        {
            try { RibbonUi?.ActivateTab("mrcTab"); }
            catch (Exception ex) { _log.Warn("ActivateTab failed: " + ex.Message); }
        }

        // ─────────────────────────────────────────────────────────────────
        //  The four "Insert…" buttons.  They all funnel through the same
        //  picker dialog; the SourceScope enum just pre-filters it.
        //  The main button doubles as "Edit Citation": when the cursor sits
        //  inside an existing MRCITE field, picking replaces that field's
        //  citations instead of inserting a new one (Zotero-style).
        // ─────────────────────────────────────────────────────────────────
        public void OnInsertCitation(IRibbonControl control)
        {
            Field field = FindMirareCitationFieldAtSelection();
            OpenPicker(CitationSourceScope.All, field);
        }

        public void OnInsertFromLibrary(IRibbonControl control)
            => OpenPicker(CitationSourceScope.Library, null);

        public void OnInsertFromProject(IRibbonControl control)
            => OpenPicker(CitationSourceScope.Project, null);

        public void OnInsertFromRemote(IRibbonControl control)
            => OpenPicker(CitationSourceScope.Remote, null);

        /// <summary>Button label: "Edit Citation" when the cursor is inside
        /// one of our fields. Refreshed via SelectionChange → Invalidate.</summary>
        public string GetInsertButtonLabel()
            => FindMirareCitationFieldAtSelection() != null ? "Edit Citation" : "Insert Citation";

        /// <summary>The MRCITE citation field containing the cursor, if any.
        /// The zone spans the surrounding plain-text parentheses too — from
        /// the opening "(" to the closing ")" — so the button reads
        /// "Edit Citation" anywhere inside the whole citation.</summary>
        private Field FindMirareCitationFieldAtSelection()
        {
            try
            {
                Document doc = _word.ActiveDocument;
                if (doc == null) return null;
                Range sel = _word.Selection.Range;
                foreach (Field f in doc.Fields)
                {
                    if (!CitedTracker.IsMirareCitationCode(f.Code.Text)) continue;
                    // Field chars sit one position out from the code range;
                    // the parens are one further (field span:
                    // "(" at Code.Start-2 … ")" at Result.End+1).
                    if (sel.Start >= f.Code.Start - 2 && sel.Start <= f.Result.End + 2 &&
                        sel.End <= f.Result.End + 2)
                        return f;
                }
            }
            catch (Exception ex) { _log.Warn("FindMirareCitationFieldAtSelection: " + ex.Message); }
            return null;
        }

        /// <summary>An existing citation field the collapsed cursor touches —
        /// directly after its ")" or directly before its "(". Inserting there
        /// merges with that field instead of creating "(a) (b)" pairs.</summary>
        private Field FindAdjacentMirareField(Document doc)
        {
            try
            {
                Range sel = _word.Selection.Range;
                if (sel.Start != sel.End) return null;   // collapsed cursor only
                foreach (Field f in doc.Fields)
                {
                    if (!CitedTracker.IsMirareCitationCode(f.Code.Text)) continue;
                    if (sel.Start == f.Result.End + 1 || sel.End == f.Code.Start - 1)
                        return f;
                }
            }
            catch (Exception ex) { _log.Warn("FindAdjacentMirareField: " + ex.Message); }
            return null;
        }

        private void OpenPicker(CitationSourceScope scope, Field fieldToEdit)
        {
            try
            {
                Document doc = _word.ActiveDocument;
                if (doc == null)
                {
                    MessageBox.Show(
                        "Open a Word document first — Mirare Cite has nowhere to insert the citation.",
                        "Mirare Cite",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);
                    return;
                }

                // Pre-load cited IDs from the active document so the picker
                // can show the cited indicator immediately.
                var citedIds = _tracker.ScanDocument(doc);
                _log.Info($"OpenPicker: scope={scope} edit={fieldToEdit != null} cited={citedIds.Count}");

                // Editing: the dialog's preview starts with the field's
                // current citations (from the session cache; anything missing
                // is matched against the picker's loaded data by id).
                IEnumerable<Citation> preloaded = null;
                HashSet<string> preselectedIds = null;
                if (fieldToEdit != null)
                {
                    preselectedIds = new HashSet<string>(
                        CitedTracker.ParseFieldIds(fieldToEdit.Code.Text));
                    preloaded = preselectedIds
                        .Select(id => _citationCache.TryGetValue(id, out var c) ? c : null)
                        .Where(c => c != null)
                        .ToList();
                }

                using (var form = new CitationPickerForm(scope, Style, RemoteEndpoint,
                            LastLibraryPath, LastProjectPath, citedIds, _log,
                            preloaded, preselectedIds, CreateFormatter()))
                {
                    if (form.ShowDialog() == DialogResult.OK && form.SelectedCitation != null)
                    {
                        var picked = form.SelectedCitations ?? new List<Citation> { form.SelectedCitation };

                        // Inserting right next to an existing citation → merge
                        // into that field so we never grow "(a) (b)" pairs.
                        Field mergeTarget = fieldToEdit ?? FindAdjacentMirareField(doc);
                        if (mergeTarget != null)
                        {
                            var lookup = BuildRecordLookup();
                            var combined = new List<Citation>();
                            var seen = new HashSet<string>();
                            foreach (var id in CitedTracker.ParseFieldIds(mergeTarget.Code.Text))
                                if (lookup.TryGetValue(id, out var ec) && seen.Add(id))
                                    combined.Add(ec);
                            foreach (var c in picked)
                                if (seen.Add(c.Id)) combined.Add(c);
                            ReplaceCitationField(doc, mergeTarget, combined);
                            foreach (var c in combined) _citationCache[c.Id] = c;
                        }
                        else
                        {
                            InsertCitationsAtSelection(doc, picked, Style);
                            foreach (var c in picked) _citationCache[c.Id] = c;
                        }

                        // After insertion, update the tracker so the dynamic
                        // label on the ribbon reflects the new count.
                        foreach (var c in picked) _tracker.AddCited(c.Id);
                    }
                }

                // Numeric styles: give the new field (and any others) their
                // current numbers — Update Bibliography finalizes them.
                RenumberAllCitations(_word.ActiveDocument ?? doc);
                RestoreRibbonTab();
            }
            catch (Exception ex)
            {
                _log.Error("OpenPicker failed", ex);
                MessageBox.Show("Insert citation failed: " + ex.Message,
                    "Mirare Cite", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // ─────────────────────────────────────────────────────────────────
        //  Refresh — re-scan the active document for cited items and update
        //  the indicator.  Useful if the user has been editing citations
        //  manually or pasted text from another document.
        // ─────────────────────────────────────────────────────────────────
        public void OnRefreshCitedIndicator(IRibbonControl control)
        {
            try
            {
                Document doc = _word.ActiveDocument;
                if (doc == null) return;
                _tracker.ScanDocument(doc);
                _log.Info($"Refresh: cited={_tracker.CitedIds.Count}");
                // The ribbon label is only re-queried when its control is
                // invalidated — clicking Refresh must refresh it too, or it
                // keeps showing the stale count.
                try { RibbonUi?.InvalidateControl("mrcCitedCount"); } catch { }
                MessageBox.Show(
                    $"Refreshed. {_tracker.CitedIds.Count} item(s) currently cited in this document.",
                    "Mirare Cite", MessageBoxButtons.OK, MessageBoxIcon.Information);
                RestoreRibbonTab();
            }
            catch (Exception ex)
            {
                _log.Error("OnRefreshCitedIndicator failed", ex);
            }
        }

        public void OnEditBibliography(IRibbonControl control)
        {
            try
            {
                Document doc = _word.ActiveDocument;
                if (doc == null) return;

                var cited = ResolveCitedRecords(doc);
                if (cited.Count == 0)
                {
                    MessageBox.Show(
                        "No cited items to list yet — insert at least one citation first.\n" +
                        "(Cited articles are matched against the session cache and the\n" +
                        "library / project files configured in Settings.)",
                        "Mirare Cite", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                var bibText = CreateFormatter().BuildBibliography(cited).Replace("\r\n", "\r");

                // Zotero-style: the bibliography lives in ONE DOCVARIABLE
                // field that is updated in place on every click — Word owns
                // the result section, so updates never corrupt the field.
                string bibVar = $"MRCITE BIBLIOGRAPHY style={Style}";
                Field bibField = FindBibliographyField(doc);
                if (bibField != null)
                {
                    RenameFieldVariable(doc, bibField, bibVar, bibText);
                    _log.Info($"Bibliography field updated with {cited.Count} entries");
                }
                else
                {
                    // First run: heading paragraph + bibliography field at end.
                    Range tail = doc.Content;
                    tail.Collapse(WdCollapseDirection.wdCollapseEnd);
                    tail.InsertParagraphAfter();
                    tail = doc.Content;
                    tail.Collapse(WdCollapseDirection.wdCollapseEnd);
                    tail.InsertAfter("References\r");
                    Range atEnd = doc.Range(doc.Content.End - 1, doc.Content.End - 1);
                    InsertDocVariableField(doc, atEnd, bibVar, bibText);
                    _log.Info($"Bibliography field created with {cited.Count} entries");
                }
                _word.ActiveWindow.View.ShowFieldCodes = false;

                // "Update" is a full refresh: re-render every citation field
                // with the current style (fixes multi-bracket groups and any
                // stale displays), then renumber numeric citations.
                RestyleDocument(doc);
            }
            catch (Exception ex)
            {
                _log.Error("OnEditBibliography failed", ex);
                MessageBox.Show("Bibliography update failed: " + ex.Message,
                    "Mirare Cite", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // ─────────────────────────────────────────────────────────────────
        //  Record lookup — cited ids carry no article data, so resolve them
        //  against the session cache first, then the library and project
        //  files from Settings. This makes restyling and bibliography
        //  generation survive Word restarts.
        // ─────────────────────────────────────────────────────────────────
        private Dictionary<string, Citation> BuildRecordLookup()
        {
            var lookup = new Dictionary<string, Citation>();
            foreach (var kv in _citationCache)
                lookup[kv.Key] = kv.Value;

            try
            {
                if (!string.IsNullOrEmpty(LastLibraryPath) && File.Exists(LastLibraryPath))
                    foreach (var c in new RefManagerLoader(_log).Load(LastLibraryPath))
                        lookup[c.Id] = c;
            }
            catch (Exception ex) { _log.Warn("RecordLookup library: " + ex.Message); }

            try
            {
                if (!string.IsNullOrEmpty(LastProjectPath) && File.Exists(LastProjectPath))
                    foreach (var c in new ProjectLoader(_log).Load(LastProjectPath))
                        lookup[c.Id] = c;
            }
            catch (Exception ex) { _log.Warn("RecordLookup project: " + ex.Message); }

            return lookup;
        }

        private List<Citation> ResolveCitedRecords(Document doc)
        {
            var lookup = BuildRecordLookup();
            var cited = new List<Citation>();
            foreach (var id in _tracker.CitedIds)
                if (lookup.TryGetValue(id, out var c))
                    cited.Add(c);
            return cited;
        }

        /// <summary>
        /// Re-renders every Mirare citation field and the bibliography field
        /// in the document with the current style — called when the user
        /// changes the citation style (or CSL file) in Settings.
        /// </summary>
        private void RestyleDocument(Document doc)
        {
            try
            {
                var lookup = BuildRecordLookup();
                var formatter = CreateFormatter();

                // Snapshot the fields first — rewrites during iteration
                // would invalidate the collection.
                var fields = new List<Field>();
                foreach (Field f in doc.Fields) fields.Add(f);

                int restyled = 0;
                foreach (var f in fields)
                {
                    string norm = CitedTracker.NormalizeFieldCode(f.Code.Text);
                    if (!CitedTracker.IsMirareCitationCode(norm)) continue;

                    var ids = CitedTracker.ParseFieldIds(norm);
                    var citations = new List<Citation>();
                    bool complete = ids.Count > 0;
                    foreach (var id in ids)
                    {
                        if (lookup.TryGetValue(id, out var c)) citations.Add(c);
                        else { complete = false; break; }
                    }
                    if (!complete)
                    {
                        _log.Warn("Restyle skipped field (records missing): " + norm);
                        continue;
                    }

                    string display = RenderCitationDisplay(citations, Style, formatter);
                    string newVarName = BuildMirareFieldCode(citations, Style);
                    RenameFieldVariable(doc, f, newVarName, display);
                    restyled++;
                }

                // Bibliography field follows the new style too.
                var bib = FindBibliographyField(doc);
                if (bib != null)
                {
                    _tracker.ScanDocument(doc);
                    var cited = ResolveCitedRecords(doc);
                    if (cited.Count > 0)
                    {
                        var bibText = formatter.BuildBibliography(cited).Replace("\r\n", "\r");
                        RenameFieldVariable(doc, bib, $"MRCITE BIBLIOGRAPHY style={Style}", bibText);
                    }
                }

                _word.ActiveWindow.View.ShowFieldCodes = false;

                // Switching to a numeric style: placeholders → real numbers.
                RenumberAllCitations(doc);
                _log.Info($"Restyled {restyled} citation field(s) to {Style}");
            }
            catch (Exception ex)
            {
                _log.Error("RestyleDocument failed", ex);
                MessageBox.Show("Restyling citations failed: " + ex.Message,
                    "Mirare Cite", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private static Field FindBibliographyField(Document doc)
        {
            foreach (Field f in doc.Fields)
            {
                string code = CitedTracker.NormalizeFieldCode(f.Code.Text);
                if (code.StartsWith("MRCITE BIBLIOGRAPHY", StringComparison.OrdinalIgnoreCase))
                    return f;
            }
            return null;
        }

        // ─────────────────────────────────────────────────────────────────
        //  DOCVARIABLE field machinery — the reliable way to keep display
        //  text INSIDE a field. Lab-verified: ADDIN fields created via
        //  Fields.Add have NO result section and Word's Result range points
        //  PAST the field end, so Result.Text writes land OUTSIDE the field
        //  (plain text next to it) — the "broken structure". DOCVARIABLE
        //  fields have real results that Word itself maintains via Update().
        // ─────────────────────────────────────────────────────────────────

        /// <summary>
        /// Points the field at a (possibly new) variable and value, then
        /// lets Word re-render the result. If the name changed, the old
        /// variable is removed once nothing references it anymore.
        /// </summary>
        private void RenameFieldVariable(Document doc, Field field, string newVarName, string value)
        {
            string oldVarName = CitedTracker.NormalizeFieldCode(field.Code.Text);
            bool nameChanged = !string.Equals(oldVarName, newVarName, StringComparison.Ordinal);

            if (nameChanged)
                field.Code.Text = $" DOCVARIABLE \"{newVarName}\" ";

            Variable v = FindVariable(doc, newVarName);
            if (v == null) doc.Variables.Add(newVarName, value);
            else v.Value = value;

            field.Update();
            ApplyRichFormatting(doc, field);

            if (nameChanged && !VariableStillReferenced(doc, oldVarName))
                FindVariable(doc, oldVarName)?.Delete();
        }

        private Variable FindVariable(Document doc, string name)
        {
            foreach (Variable v in doc.Variables)
                if (string.Equals(v.Name, name, StringComparison.Ordinal))
                    return v;
            return null;
        }

        private void InsertDocVariableField(Document doc, Range at, string varName, string value)
        {
            Variable v = FindVariable(doc, varName);
            if (v == null)
                doc.Variables.Add(varName, value);
            else
                v.Value = value;

            // NOTE: Fields.Add prepends the field keyword itself — pass only
            // the quoted variable name or the code doubles ("DOCVARIABLE DOCVARIABLE").
            Field fld = doc.Fields.Add(at, WdFieldType.wdFieldDocVariable,
                $"\"{varName}\"", PreserveFormatting: false);
            fld.ShowCodes = false;
            fld.Update();
            ApplyRichFormatting(doc, fld);
        }

        /// <summary>
        /// The CSL renderer emits font-style/font-weight as marker characters
        /// (italic = \u0001..\u0002, bold = \u0003..\u0004). After the field
        /// updates, convert each marker pair into real Word character
        /// formatting on the enclosed text and delete the markers.
        /// </summary>
        private void ApplyRichFormatting(Document doc, Field fld)
        {
            const char ItalicOn = '\u0001', ItalicOff = '\u0002';
            const char BoldOn = '\u0003', BoldOff = '\u0004';
            try
            {
                for (int guard = 0; guard < 1000; guard++)
                {
                    Range res = fld.Result;
                    string text = res?.Text;
                    if (string.IsNullOrEmpty(text)) break;

                    int open = text.IndexOfAny(new[] { ItalicOn, BoldOn });
                    if (open < 0) break;
                    bool bold = text[open] == BoldOn;
                    char closeChar = bold ? BoldOff : ItalicOff;
                    int close = text.IndexOf(closeChar, open + 1);
                    if (close < 0) break;   // unpaired marker — leave as-is

                    int basePos = res.Start;
                    if (close > open + 1)
                    {
                        Range inner = doc.Range(basePos + open + 1, basePos + close);
                        inner.Font.Italic = bold ? 1 : 0;
                        inner.Font.Bold = bold ? 1 : 0;
                    }
                    // Delete close first so the open position stays valid.
                    doc.Range(basePos + close, basePos + close + 1).Delete();
                    doc.Range(basePos + open, basePos + open + 1).Delete();
                }
            }
            catch (Exception ex)
            {
                _log?.Warn("Rich formatting failed: " + ex.Message);
            }
        }

        /// <summary>
        /// Numeric styles (built-in Numeric, IEEE/Vancouver CSL) cite by
        /// bibliography position. After the document has been scanned (so
        /// CitedIds reflects document order) this rewrites every numeric
        /// citation field's display from its [mirare:id] placeholder to the
        /// real [n]. Called after insert, edit, Update Bibliography, and
        /// style changes.
        /// </summary>
        private void RenumberAllCitations(Document doc)
        {
            try
            {
                var formatter = CreateFormatter();
                if (!formatter.IsNumericStyle) return;

                _tracker.ScanDocument(doc);
                var lookup = BuildRecordLookup();
                var numbers = new Dictionary<string, int>();
                int n = 1;
                foreach (var id in _tracker.CitedIds)
                    if (lookup.ContainsKey(id) && !numbers.ContainsKey(id))
                        numbers[id] = n++;

                if (numbers.Count == 0) return;

                var fields = new List<Field>();
                foreach (Field f in doc.Fields) fields.Add(f);

                int renumbered = 0;
                foreach (var f in fields)
                {
                    string norm = CitedTracker.NormalizeFieldCode(f.Code.Text);
                    if (!CitedTracker.IsMirareCitationCode(norm)) continue;

                    var ids = CitedTracker.ParseFieldIds(norm);
                    if (ids.Count == 0 || ids.Any(id => !numbers.ContainsKey(id))) continue;

                    var citations = new List<Citation>();
                    var nums = new List<int>();
                    foreach (var id in ids)
                    {
                        citations.Add(lookup.TryGetValue(id, out var cc)
                            ? cc : new Citation { Id = id });
                        nums.Add(numbers[id]);
                    }
                    string display = formatter.RenderGroup(citations, nums);
                    RenameFieldVariable(doc, f, norm, display);
                    renumbered++;
                }
                if (renumbered > 0)
                    _log.Info($"Renumbered {renumbered} numeric citation field(s)");
            }
            catch (Exception ex)
            {
                _log.Error("RenumberAllCitations failed", ex);
            }
        }

        public void OnSettings(IRibbonControl control)
        {
            var styleBefore = (Style, CslStylePath);
            using (var f = new SettingsForm(Style, RemoteEndpoint,
                       LastLibraryPath, LastProjectPath, CslStylePath, _log))
            {
                if (f.ShowDialog() == DialogResult.OK)
                {
                    Style = f.Style;
                    RemoteEndpoint = f.RemoteEndpoint;
                    LastLibraryPath = f.LibraryPath;
                    LastProjectPath = f.ProjectPath;
                    CslStylePath = f.CslStylePath;
                    SaveSettings();
                }
            }

            // Style (or CSL file) changed → re-render every Mirare field in
            // the document, Zotero-style: in-text citations and the
            // bibliography all switch to the new style at once.
            Document doc = _word.ActiveDocument;
            if (doc != null && styleBefore != (Style, CslStylePath))
                RestyleDocument(doc);

            RestoreRibbonTab();
        }

        // ─────────────────────────────────────────────────────────────────
        //  Dynamic label on the ribbon — "N cited" — invalidated when the
        //  tracker's count changes.  Returns the string Word renders.
        // ─────────────────────────────────────────────────────────────────
        public string GetCitedCountLabel(IRibbonControl control)
        {
            int n = _tracker?.CitedIds?.Count ?? 0;
            if (n == 0)
            {
                // Nothing cached — rescan (covers Word-start document restore
                // and manually added/deleted Mirare fields).
                try
                {
                    Document doc = _word.ActiveDocument;
                    if (doc != null) n = _tracker.ScanDocument(doc).Count;
                }
                catch { /* no document — keep 0 */ }
            }
            return n == 0 ? "No items cited"
                 : n == 1 ? "1 item cited"
                 : n + " items cited";
        }

        // ─────────────────────────────────────────────────────────────────
        //  Called from Connect.cs when the user switches documents.
        // ─────────────────────────────────────────────────────────────────
        public void OnActiveDocumentChanged(Document doc)
        {
            if (doc == null) return;
            _tracker.ScanDocument(doc);
            _log.Info($"Active doc changed: {doc.Name} cited={_tracker.CitedIds.Count}");
        }

        // ─────────────────────────────────────────────────────────────────
        //  Insert citation — writes a Word ADDIN field at the cursor whose
        //  hidden code carries the citation ids + style, and whose display
        //  text is the rendered citation. The parentheses are PLAIN TEXT
        //  around the field (Zotero-style), so the user can style them:
        //      ( { ADDIN MRCITE id=… style=Apa }Kaur & Newell, 2024 )
        //
        //  Multiple picks share ONE field with the results joined by "; ",
        //  exactly like Zotero/Mendeley multi-citations.
        //
        //  Lab-verified Word behavior this code relies on:
        //   * ADDIN fields created via Fields.Add have NO result section and
        //     Word's Result range points PAST the field end — Result.Text
        //     writes land OUTSIDE the field as plain text ("broken structure").
        //   * DOCVARIABLE fields have REAL result sections that Word itself
        //     maintains: Variables[name].Value = text; field.Update() renders
        //     the text inside the field, survives ShowCodes, copies cleanly,
        //     and updates in place. So that's what we use.
        // ─────────────────────────────────────────────────────────────────
        private static string BuildMirareFieldCode(IList<Citation> citations, CitationStyle style)
            => "MRCITE" + string.Join("", citations.Select(c => " id=" + c.Id)) + " style=" + style;

        private void InsertCitationsAtSelection(Document doc, IList<Citation> citations, CitationStyle style)
        {
            Range sel = _word.Selection.Range;
            var formatter = CreateFormatter();
            string display = RenderCitationDisplay(citations, style, formatter);
            string varName = BuildMirareFieldCode(citations, style);

            // The whole "(Author, Year)" — parens included — lives INSIDE the
            // field, Zotero/Mendeley-style, so Word itself treats the full
            // segment as one field and cursor detection is native.
            InsertDocVariableField(doc, sel, varName, display);

            // If field-codes view got toggled on (Alt+F9 or the Word option),
            // every field renders as its raw code — force result view.
            _word.ActiveWindow.View.ShowFieldCodes = false;

            _log.Info($"Inserted {citations.Count} citation(s) text=\"{display}\"");
        }

        /// <summary>
        /// Edit mode: swap the citations in an existing field. The field's
        /// document variable carries the new value and Word re-renders the
        /// result in place — no deletion needed, so nothing is left behind.
        /// If the citation ids changed, the field is rebuilt (the variable
        /// name encodes them); the old variable is removed if now orphaned.
        /// </summary>
        private void ReplaceCitationField(Document doc, Field oldField, IList<Citation> citations)
        {
            string oldVarName = CitedTracker.NormalizeFieldCode(oldField.Code.Text);
            bool wasDocVariable = oldField.Type == WdFieldType.wdFieldDocVariable;

            var formatter = CreateFormatter();
            string display = RenderCitationDisplay(citations, Style, formatter);
            string varName = BuildMirareFieldCode(citations, Style);

            if (wasDocVariable && string.Equals(oldVarName, varName, StringComparison.Ordinal))
            {
                RenameFieldVariable(doc, oldField, varName, display);
            }
            else
            {
                // Ids (or field kind) changed — rebuild the field in place.
                // Legacy ADDIN fields have unreliable Result ranges, so find
                // the true field-end character by scanning for chr(21).
                bool legacyAddin = oldField.Type == WdFieldType.wdFieldAddin;
                int start = oldField.Code.Start - 1;
                int end = legacyAddin
                    ? FindFieldTrueEnd(doc, oldField.Code.End) + 1
                    : oldField.Result.End + 1;
                doc.Range(start, end).Delete();
                InsertDocVariableField(doc, doc.Range(start, start), varName, display);

                // Remove the old variable if no other field still uses it.
                if (!string.Equals(oldVarName, varName, StringComparison.Ordinal)
                    && !VariableStillReferenced(doc, oldVarName))
                {
                    Variable v = FindVariable(doc, oldVarName);
                    v?.Delete();
                }
            }
            _word.ActiveWindow.View.ShowFieldCodes = false;
            _log.Info($"Replaced citation field with {citations.Count} citation(s) text=\"{display}\"");
        }

        /// <summary>True end of an ADDIN field = position of its field-end
        /// character. ADDIN Result ranges are unreliable, so scan for chr21.</summary>
        private static int FindFieldTrueEnd(Document doc, int fromPos)
        {
            try
            {
                int last = Math.Min(fromPos + 1000, doc.Content.End - 1);
                Range probe = doc.Range(fromPos, fromPos + 1);
                for (int pos = fromPos; pos < last; pos++)
                {
                    if (probe.Text == "\u0015") return pos;
                    probe = doc.Range(pos + 1, pos + 2);
                }
            }
            catch { }
            return fromPos;
        }

        private bool VariableStillReferenced(Document doc, string varName)
        {
            foreach (Field f in doc.Fields)
            {
                if (string.Equals(CitedTracker.NormalizeFieldCode(f.Code.Text), varName,
                                  StringComparison.Ordinal))
                    return true;
            }
            return false;
        }

        // ─────────────────────────────────────────────────────────────────
        //  Settings persistence — %APPDATA%\MirareCite\settings.json
        // ─────────────────────────────────────────────────────────────────
        private string SettingsPath =>
            Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                "MirareCite", "settings.json");

        private void LoadSettings()
        {
            try
            {
                if (!File.Exists(SettingsPath)) return;
                var s = JsonConvert.DeserializeObject<SettingsDto>(File.ReadAllText(SettingsPath));
                if (s == null) return;
                Style = s.Style;
                RemoteEndpoint = s.RemoteEndpoint;
                LastLibraryPath = s.LastLibraryPath;
                LastProjectPath = s.LastProjectPath;
                CslStylePath = s.CslStylePath ?? "";
            }
            catch (Exception ex) { _log.Warn("LoadSettings failed: " + ex.Message); }
        }

        private void SaveSettings()
        {
            try
            {
                Directory.CreateDirectory(Path.GetDirectoryName(SettingsPath));
                var dto = new SettingsDto
                {
                    Style = Style,
                    RemoteEndpoint = RemoteEndpoint,
                    LastLibraryPath = LastLibraryPath,
                    LastProjectPath = LastProjectPath,
                    CslStylePath = CslStylePath
                };
                File.WriteAllText(SettingsPath, JsonConvert.SerializeObject(dto, Formatting.Indented));
            }
            catch (Exception ex) { _log.Warn("SaveSettings failed: " + ex.Message); }
        }

        private class SettingsDto
        {
            public CitationStyle Style { get; set; }
            public string RemoteEndpoint { get; set; }
            public string LastLibraryPath { get; set; }
            public string LastProjectPath { get; set; }
            public string CslStylePath { get; set; }
        }
    }
}
