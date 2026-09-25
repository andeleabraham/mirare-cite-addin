// ============================================================================
//  CitedItem.cs — the tracker's notion of a cited item in a document.
//
//  A MRCITE Word field (see RibbonCallbacks.cs → InsertCitationAtSelection)
//  carries the citation Id and the style at insertion time.  CitedTracker
//  scans all MRCITE fields in the active document and produces this set.
// ============================================================================

namespace MirareCiteAddIn.Models
{
    public class CitedItem
    {
        public string Id { get; set; }
        public CitationStyle Style { get; set; }
        public int FieldIndex { get; set; }   // for highlighting in the doc
    }
}
