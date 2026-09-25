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
        // ─────────────────────────────────────────────────────────────────
        public void OnInsertCitation(IRibbonControl control)
            => OpenPicker(CitationSourceScope.All);

        public void OnInsertFromLibrary(IRibbonControl control)
            => OpenPicker(CitationSourceScope.Library);

        public void OnInsertFromProject(IRibbonControl control)
            => OpenPicker(CitationSourceScope.Project);

        public void OnInsertFromRemote(IRibbonControl control)
            => OpenPicker(CitationSourceScope.Remote);

        private void OpenPicker(CitationSourceScope scope)
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
                _log.Info($"OpenPicker: scope={scope} cited={citedIds.Count}");

                using (var form = new CitationPickerForm(scope, Style, RemoteEndpoint,
                            LastLibraryPath, LastProjectPath, citedIds, _log))
                {
                    if (form.ShowDialog() == DialogResult.OK && form.SelectedCitation != null)
                    {
                        InsertCitationAtSelection(doc, form.SelectedCitation, Style);

                        // Remember the full record so Update Bibliography can
                        // render it later (the field only stores the id).
                        _citationCache[form.SelectedCitation.Id] = form.SelectedCitation;

                        // After insertion, update the tracker so the dynamic
                        // label on the ribbon reflects the new count.
                        _tracker.AddCited(form.SelectedCitation.Id);
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
                var bibText = formatter.BuildBibliography(cited);

                // Append a References section at the end of the document.
                Range tail = doc.Content;
                tail.Collapse(WdCollapseDirection.wdCollapseEnd);
                tail.InsertParagraphAfter();
                tail = doc.Content;
                tail.Collapse(WdCollapseDirection.wdCollapseEnd);
                tail.InsertAfter("References\r" + bibText);
                _log.Info($"Bibliography updated with {cited.Count} entries");
            }
            catch (Exception ex)
            {
                _log.Error("OnEditBibliography failed", ex);
                MessageBox.Show("Bibliography update failed: " + ex.Message,
                    "Mirare Cite", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
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
        //  Insert citation — writes a Word field at the cursor.  Using a
        //  hidden bookmarked field (not just plain text) is what lets the
        //  CitedTracker reliably re-scan the document later and what lets
        //  OnEditBibliography rebuild the references list without dupes.
        //
        //  The field's text is the rendered in-text citation (e.g.
        //  "(Smith et al., 2022)"); the field's code is a hidden
        //  MRCITE instruction carrying the citation ID + style, so we
        //  can round-trip on refresh:
        //      { MRCITE id=abc123 style=apa }
        // ─────────────────────────────────────────────────────────────────
        private void InsertCitationAtSelection(Document doc, Citation c, CitationStyle style)
        {
            Range sel = _word.Selection.Range;
            var formatter = new CitationFormatter(style);
            string inText = formatter.InText(c);
            string fieldCode = $"MRCITE id={c.Id} style={style}";

            // Verified against a live Word instance: Fields.Add + ShowCodes(false)
            // + Result.Text renders ONLY the citation text — the ADDIN code stays
            // hidden. NOTE: Field.Result.Text always READS back empty for ADDIN
            // fields (Word quirk); the text is nevertheless stored and displayed.
            Field fld = doc.Fields.Add(sel,
                WdFieldType.wdFieldAddin,
                fieldCode,
                PreserveFormatting: false);
            fld.ShowCodes = false;
            fld.Result.Text = inText;

            // If field-codes view got toggled on (Alt+F9 or the Word option),
            // every field renders as its raw "{ ADDIN ... }" code — force the
            // window back to result view.
            _word.ActiveWindow.View.ShowFieldCodes = false;

            _log.Info($"Inserted citation id={c.Id} style={style} text=\"{inText}\"");
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
