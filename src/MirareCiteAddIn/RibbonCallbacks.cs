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

        /// <summary>The MRCITE citation field containing the cursor, if any.</summary>
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
                    // Field span: one char before the code (field-begin) up to
                    // the end of the result (field-end). Use a small margin so
                    // a cursor right at either edge counts as "inside".
                    if (sel.Start >= f.Code.Start - 1 && sel.Start <= f.Result.End + 1 &&
                        sel.End <= f.Result.End + 1)
                        return f;
                }
            }
            catch (Exception ex) { _log.Warn("FindMirareCitationFieldAtSelection: " + ex.Message); }
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

                using (var form = new CitationPickerForm(scope, Style, RemoteEndpoint,
                            LastLibraryPath, LastProjectPath, citedIds, _log))
                {
                    if (form.ShowDialog() == DialogResult.OK && form.SelectedCitation != null)
                    {
                        var picked = form.SelectedCitations ?? new List<Citation> { form.SelectedCitation };

                        if (fieldToEdit != null)
                            ReplaceCitationField(doc, fieldToEdit, picked);
                        else
                            InsertCitationsAtSelection(doc, picked, Style);

                        // Remember the full records so Update Bibliography can
                        // render them later (the field only stores the ids).
                        foreach (var c in picked) _citationCache[c.Id] = c;

                        // After insertion, update the tracker so the dynamic
                        // label on the ribbon reflects the new count.
                        foreach (var c in picked) _tracker.AddCited(c.Id);
                    }
                }
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
                // The ribbon label is invalidated by Word automatically when
                // the user clicks the button (since onAction has fired); to
                // force a refresh from code we'd need the IRibbonUI pointer.
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

                // Join the document's cited ids (parsed from the MRCITE
                // fields) against the session cache of full Citation records.
                var cited = new List<Citation>();
                foreach (var id in _tracker.CitedIds)
                    if (_citationCache.TryGetValue(id, out var c))
                        cited.Add(c);

                if (cited.Count == 0)
                {
                    MessageBox.Show(
                        "No cited items to list yet — insert at least one citation first.\n" +
                        "(The bibliography is built from citations inserted in this Word session.)",
                        "Mirare Cite", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                var formatter = new CitationFormatter(Style);
                var bibText = formatter.BuildBibliography(cited).Replace("\r\n", "\r");

                // Zotero-style: the bibliography lives in ONE field that is
                // updated in place on every click — never appended twice.
                // Repeated Result.Text writes corrupt ADDIN fields, so the
                // field is deleted and recreated at the same position.
                Field bibField = FindBibliographyField(doc);
                if (bibField != null)
                {
                    int at = bibField.Code.Start - 1;
                    bibField.Delete();
                    CreateBibliographyField(doc, doc.Range(at, at), bibText);
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
                    CreateBibliographyField(doc, atEnd, bibText);
                    _log.Info($"Bibliography field created with {cited.Count} entries");
                }
                _word.ActiveWindow.View.ShowFieldCodes = false;
            }
            catch (Exception ex)
            {
                _log.Error("OnEditBibliography failed", ex);
                MessageBox.Show("Bibliography update failed: " + ex.Message,
                    "Mirare Cite", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private static Field FindBibliographyField(Document doc)
        {
            foreach (Field f in doc.Fields)
            {
                if (f.Type != WdFieldType.wdFieldAddin) continue;
                string code = CitedTracker.NormalizeFieldCode(f.Code.Text);
                if (code.StartsWith("MRCITE BIBLIOGRAPHY", StringComparison.OrdinalIgnoreCase))
                    return f;
            }
            return null;
        }

        private void CreateBibliographyField(Document doc, Range at, string bibText)
        {
            Field fld = doc.Fields.Add(at, WdFieldType.wdFieldAddin,
                $"MRCITE BIBLIOGRAPHY style={Style}", PreserveFormatting: false);
            fld.ShowCodes = false;
            fld.Result.Text = bibText;
        }

        public void OnSettings(IRibbonControl control)
        {
            using (var f = new SettingsForm(Style, RemoteEndpoint,
                       LastLibraryPath, LastProjectPath, _log))
            {
                if (f.ShowDialog() == DialogResult.OK)
                {
                    Style = f.Style;
                    RemoteEndpoint = f.RemoteEndpoint;
                    LastLibraryPath = f.LibraryPath;
                    LastProjectPath = f.ProjectPath;
                    SaveSettings();
                }
            }
            RestoreRibbonTab();
        }

        // ─────────────────────────────────────────────────────────────────
        //  Dynamic label on the ribbon — "N cited" — invalidated when the
        //  tracker's count changes.  Returns the string Word renders.
        // ─────────────────────────────────────────────────────────────────
        public string GetCitedCountLabel(IRibbonControl control)
        {
            int n = _tracker?.CitedIds?.Count ?? 0;
            return n == 1 ? "1 item cited" : n + " items cited";
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
        //   * Fields.Add + ShowCodes(false) + Result.Text renders only the
        //     result — but Result.Text always READS back empty for ADDIN
        //     fields, and writing it a SECOND time corrupts (prepends), so
        //     updates must delete + recreate the field (see Replace…).
        // ─────────────────────────────────────────────────────────────────
        private static string BuildMirareFieldCode(IList<Citation> citations, CitationStyle style)
            => "MRCITE" + string.Join("", citations.Select(c => " id=" + c.Id)) + " style=" + style;

        private void InsertCitationsAtSelection(Document doc, IList<Citation> citations, CitationStyle style)
        {
            Range sel = _word.Selection.Range;
            var formatter = new CitationFormatter(style);
            string core = string.Join("; ", citations.Select(c => formatter.InTextCore(c)));
            string fieldCode = BuildMirareFieldCode(citations, style);
            bool wrapParens = style != CitationStyle.Numeric;   // "[1]" brackets its own

            if (wrapParens)
            {
                // Plain-text parens first, then the field goes between them.
                Range r = sel;
                r.Text = "()";
                Range inner = doc.Range(r.Start + 1, r.Start + 1);
                Field fld = doc.Fields.Add(inner, WdFieldType.wdFieldAddin,
                    fieldCode, PreserveFormatting: false);
                fld.ShowCodes = false;
                fld.Result.Text = core;
            }
            else
            {
                Field fld = doc.Fields.Add(sel, WdFieldType.wdFieldAddin,
                    fieldCode, PreserveFormatting: false);
                fld.ShowCodes = false;
                fld.Result.Text = core;
            }

            // If field-codes view got toggled on (Alt+F9 or the Word option),
            // every field renders as its raw "{ ADDIN ... }" code — force the
            // window back to result view.
            _word.ActiveWindow.View.ShowFieldCodes = false;

            _log.Info($"Inserted {citations.Count} citation(s) text=\"{core}\"");
        }

        /// <summary>
        /// Edit mode: swap the citations in an existing field. Repeated
        /// Result.Text writes corrupt ADDIN fields, so the field is deleted
        /// and recreated at the same position (the plain-text parens around
        /// it survive untouched).
        /// </summary>
        private void ReplaceCitationField(Document doc, Field oldField, IList<Citation> citations)
        {
            int at = oldField.Code.Start - 1;           // the field-begin char
            oldField.Delete();
            Range r = doc.Range(at, at);

            var formatter = new CitationFormatter(Style);
            string core = string.Join("; ", citations.Select(c => formatter.InTextCore(c)));
            Field fld = doc.Fields.Add(r, WdFieldType.wdFieldAddin,
                BuildMirareFieldCode(citations, Style), PreserveFormatting: false);
            fld.ShowCodes = false;
            fld.Result.Text = core;
            _word.ActiveWindow.View.ShowFieldCodes = false;
            _log.Info($"Replaced citation field with {citations.Count} citation(s) text=\"{core}\"");
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
                    LastProjectPath = LastProjectPath
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
        }
    }
}
