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
                        // wdFieldAddin = 81.  We use this generic field type
                        // because it preserves our code+result across saves
                        // without Word trying to re-evaluate it.
                        if ((WdFieldType)fld.Type != WdFieldType.wdFieldAddin) continue;

                        string code = fld.Code.Text ?? "";
                        if (!IsMirareFieldCode(code)) continue;

                        var parsed = ParseFieldCode(code);
                        if (parsed == null) continue;

                        if (!_citedIds.Contains(parsed.Id))
                        {
                            _citedIds.Add(parsed.Id);
                            _styles[parsed.Id] = parsed.Style;
                        }
                    }
                }
                // Walk linked story ranges (e.g. footnotes → endnotes → comments).
                cur = cur.NextStoryRange;
            }
        }

        /// <summary>
        /// Field.Code returns the instruction WITH its type keyword —
        /// " ADDIN MRCITE id=abc123 style=apa " (leading/trailing space and
        /// the ADDIN keyword included) — so normalize before matching.
        /// </summary>
        private static bool IsMirareFieldCode(string code)
        {
            if (string.IsNullOrEmpty(code)) return false;
            code = code.Trim();
            if (code.StartsWith("ADDIN ", StringComparison.OrdinalIgnoreCase))
                code = code.Substring(6).TrimStart();
            return code.StartsWith("MRCITE ", StringComparison.Ordinal);
        }

        /// <summary>
        /// Parse a field code like "MRCITE id=abc123 style=apa" into (id, style).
        /// </summary>
        private static CitedItem ParseFieldCode(string code)
        {
            code = code.Trim();
            if (code.StartsWith("ADDIN ", StringComparison.OrdinalIgnoreCase))
                code = code.Substring(6).TrimStart();
            // Tokenize by whitespace.
            var tokens = code.Split(new[] { ' ', '\t', '\r', '\n' },
                StringSplitOptions.RemoveEmptyEntries);
            string id = null;
            string styleStr = null;
            foreach (var t in tokens)
            {
                if (t.StartsWith("id=", StringComparison.Ordinal)) id = t.Substring(3);
                else if (t.StartsWith("style=", StringComparison.Ordinal)) styleStr = t.Substring(6);
            }
            if (string.IsNullOrEmpty(id)) return null;
            CitationStyle st = CitationStyle.Apa;
            if (!string.IsNullOrEmpty(styleStr))
                Enum.TryParse(styleStr, true, out st);
            return new CitedItem { Id = id, Style = st };
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
