// ============================================================================
//  CitedTracker.cs — scans the active Word document for MRCITE fields
//  (Word's wdFieldAddin) and builds the set of citation IDs that are
//  currently cited.
//
//  This set drives the "cited indicator" badge shown next to each item in
//  the picker dialog: if a citation Id is in this set, the picker shows a
//  ✓ next to it.  The set is recomputed:
//    - On add-in load
//    - On DocumentChange (user switches docs)
//    - On DocumentOpen (user opens a doc)
//    - On "Refresh Cited Indicator" button click
//    - After every successful insertion
//
//  For numeric-style citations, we also keep the order in which they were
//  first cited so the bibliography can renumber consistently.
// ============================================================================

using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Office.Interop.Word;
using MirareCiteAddIn.Models;

namespace MirareCiteAddIn.Services
{
    public class CitedTracker
    {
        private readonly Logger _log;
        // The order of insertion is preserved — important for numeric styles.
        private readonly List<string> _citedIds = new List<string>();
        private readonly Dictionary<string, CitationStyle> _styles = new Dictionary<string, CitationStyle>();

        public IReadOnlyList<string> CitedIds => _citedIds;

        public CitedTracker(Logger log) { _log = log; }

        /// <summary>
        /// Walk every field in every story of the active document; for each
        /// wdFieldAddin field whose code starts with "MRCITE ", parse the Id
        /// and style out of the code and record it.  Clears prior state.
        /// </summary>
        public HashSet<string> ScanDocument(Document doc)
        {
            _citedIds.Clear();
            _styles.Clear();
            if (doc == null) return new HashSet<string>();

            // Word fields live in stories (main, headers, footers, footnotes,
            // endnotes).  Iterate them all — citations could legitimately be
            // in footnotes (numeric style puts them there).
            foreach (Range storyRange in doc.StoryRanges)
            {
                WalkStory(storyRange);
            }
            _log.Info($"ScanDocument: {_citedIds.Count} cited ids in {doc.Name}");
            return new HashSet<string>(_citedIds);
        }

        private void WalkStory(Range r)
        {
            Range cur = r;
            while (cur != null)
            {
                if (cur.Fields != null && cur.Fields.Count > 0)
                {
                    foreach (Field fld in cur.Fields)
                    {
                        // Our citations live in DOCVARIABLE fields (current,
                        // wdFieldDocVariable = 38) or legacy ADDIN fields
                        // (wdFieldAddin = 81).
                        if (fld.Type != WdFieldType.wdFieldDocVariable &&
                            fld.Type != WdFieldType.wdFieldAddin) continue;

                        string code = fld.Code.Text ?? "";
                        if (!IsMirareCitationCode(code)) continue;

                        // One field may carry several citations
                        // ("MRCITE id=a id=b style=Apa" — multi-pick inserts).
                        var ids = ParseFieldIds(code);
                        var style = ParseFieldStyle(code);
                        foreach (var id in ids)
                        {
                            if (!_citedIds.Contains(id))
                            {
                                _citedIds.Add(id);
                                _styles[id] = style;
                            }
                        }
                    }
                }
                // Walk linked story ranges (e.g. footnotes → endnotes → comments).
                cur = cur.NextStoryRange;
            }
        }

        /// <summary>
        /// Field.Code returns the instruction WITH its type keyword —
        /// " ADDIN MRCITE id=abc123 style=apa " — so normalize before
        /// matching. Public because the ribbon callbacks use the same
        /// normalization to find citable fields.
        /// </summary>
        public static string NormalizeFieldCode(string code)
        {
            if (string.IsNullOrEmpty(code)) return "";
            code = code.Trim();
            // DOCVARIABLE fields: DOCVARIABLE  "MRCITE id=… style=…"
            if (code.StartsWith("DOCVARIABLE", StringComparison.OrdinalIgnoreCase))
            {
                int q1 = code.IndexOf('"');
                int q2 = code.LastIndexOf('"');
                code = (q1 >= 0 && q2 > q1) ? code.Substring(q1 + 1, q2 - q1 - 1) : code.Substring(11);
            }
            // Legacy ADDIN fields: ADDIN MRCITE …
            if (code.StartsWith("ADDIN ", StringComparison.OrdinalIgnoreCase))
                code = code.Substring(6).TrimStart();
            return code.Trim();
        }

        /// <summary>True if the field is an MRCITE citation (not the
        /// bibliography field — that has no id= tokens).</summary>
        public static bool IsMirareCitationCode(string code)
        {
            code = NormalizeFieldCode(code);
            if (!code.StartsWith("MRCITE ", StringComparison.Ordinal)) return false;
            if (code.IndexOf("BIBLIOGRAPHY", StringComparison.OrdinalIgnoreCase) >= 0) return false;
            return code.Contains("id=");
        }

        /// <summary>All id= tokens — a multi-citation field carries several.</summary>
        public static List<string> ParseFieldIds(string code)
        {
            var ids = new List<string>();
            foreach (var t in NormalizeFieldCode(code).Split(
                         new[] { ' ', '\t', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries))
            {
                if (t.StartsWith("id=", StringComparison.Ordinal) && t.Length > 3)
                    ids.Add(t.Substring(3));
            }
            return ids;
        }

        public static CitationStyle ParseFieldStyle(string code)
        {
            foreach (var t in NormalizeFieldCode(code).Split(
                         new[] { ' ', '\t', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries))
            {
                if (t.StartsWith("style=", StringComparison.Ordinal))
                {
                    if (Enum.TryParse(t.Substring(6), true, out CitationStyle st))
                        return st;
                    break;
                }
            }
            return CitationStyle.Apa;
        }

        /// <summary>
        /// Add a citation id after a successful insertion — keeps the tracker
        /// current without a full re-scan.
        /// </summary>
        public void AddCited(string id)
        {
            if (string.IsNullOrEmpty(id)) return;
            if (!_citedIds.Contains(id)) _citedIds.Add(id);
        }

        /// <summary>
        /// Returns the full Citation objects for every id cited in the doc,
        /// in citation order.  Used by OnEditBibliography to rebuild the
        /// references section.  The actual Citation objects are fetched
        /// from the picker's loaded set — we don't re-query them from disk
        /// here, because the user may have inserted a remote citation
        /// that is no longer cached anywhere.
        ///
        /// TODO (you): pass the picker's cached Citation set into here.
        /// For now, we return empty so OnEditBibliography produces a stub.
        /// </summary>
        public List<Citation> GetCitedCitations(Document doc)
        {
            // Stub implementation — see method comment above.
            // To wire this up fully, hold a Dictionary<string, Citation> of
            // every citation the user has ever picked in this session and
            // join it against _citedIds here.
            return new List<Citation>();
        }
    }
}
