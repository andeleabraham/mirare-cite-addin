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
        {
            string core = string.Join(formatter.MultiCitationDelimiter,
                citations.Select(c => formatter.InTextCore(c)));
            bool wrap = style == CitationStyle.Apa
                     || style == CitationStyle.Mla
                     || style == CitationStyle.Chicago;
            return wrap ? "(" + core + ")" : core;
        }

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
                            preloaded, preselectedIds))
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

                var formatter = CreateFormatter();
                var bibText = formatter.BuildBibliography(cited).Replace("\r\n", "\r");

                // Zotero-style: the bibliography lives in ONE DOCVARIABLE
                // field that is updated in place on every click — Word owns
                // the result section, so updates never corrupt the field.
                string bibVar = $"MRCITE BIBLIOGRAPHY style={Style}";
                Field bibField = FindBibliographyField(doc);
                if (bibField != null)
                {
                    SetDocVariable(doc, bibField, bibText);
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
        private void SetDocVariable(Document doc, Field field, string value)
        {
            string varName = CitedTracker.NormalizeFieldCode(field.Code.Text);
            Variable v = FindVariable(doc, varName);
            if (v != null) v.Value = value;
            field.Update();
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
        }

        public void OnSettings(IRibbonControl control)
        {
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
                SetDocVariable(doc, oldField, display);
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
