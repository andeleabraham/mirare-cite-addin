// ============================================================================
//  RefManagerLibrary.cs — in-memory model for the primary RefManager
//  library file (.refmanager.json).
//
//  Unlike .mrrcite (a ZIP), a .refmanager.json file is plain JSON — the
//  primary "library" the user cites from across all projects.
//  See docs/data-formats.md for the schema.
// ============================================================================

using System.Collections.Generic;

namespace MirareCiteAddIn.Models
{
    public class RefManagerLibrary
    {
        public string LibraryId { get; set; }
        public string Name { get; set; }
        public List<RefManagerEntry> Entries { get; set; } = new List<RefManagerEntry>();
    }

    public class RefManagerEntry
    {
        public string Id { get; set; }
        public string Title { get; set; }
        public List<string> Authors { get; set; } = new List<string>();
        public int? Year { get; set; }
        public string Journal { get; set; }
        public string Doi { get; set; }
        public string Url { get; set; }
        public string MiRNA { get; set; }
        public string TargetGene { get; set; }
        public string EvidenceType { get; set; }
        public string SourceDb { get; set; }
    }
}
