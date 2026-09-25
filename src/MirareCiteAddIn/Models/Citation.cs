// ============================================================================
//  Citation.cs — the common in-memory representation of a single Mirare
//  citation, regardless of where it came from (.mrrcite project,
//  .refmanager.json library, or remote HTTPS query).
//
//  All three loaders produce instances of this class so the picker dialog
//  and the formatter can treat them uniformly.
// ============================================================================

using System.Collections.Generic;

namespace MirareCiteAddIn.Models
{
    public class Citation
    {
        /// <summary>Stable unique id used for cited-indicator tracking.</summary>
        public string Id { get; set; }

        /// <summary>Display title (paper title, miRNA record, etc.).</summary>
        public string Title { get; set; }

        /// <summary>List of authors in "Surname, Firstname" order.</summary>
        public List<string> Authors { get; set; } = new List<string>();

        public int? Year { get; set; }
        public string Journal { get; set; }
        public string Doi { get; set; }
        public string Url { get; set; }

        // miRNA-specific fields (ignored by general formatters but preserved
        // for round-tripping back to .mrrcite projects).
        public string MiRNA { get; set; }            // e.g. "hsa-miR-21-5p"
        public string TargetGene { get; set; }       // e.g. "PTEN"
        public string EvidenceType { get; set; }     // e.g. "validated" / "predicted"
        public string SourceDb { get; set; }         // e.g. "miRTarBase", "TarBase"

        /// <summary>Where this instance came from — drives the badge in the picker.</summary>
        public CitationOrigin Origin { get; set; }

        /// <summary>Human-readable label for the picker list, e.g.
        /// "Smith J. (2022) PTEN regulation by miR-21 [miRTarBase]".</summary>
        public string DisplayLabel()
        {
            string auth = Authors.Count == 0 ? "Anon"
                        : Authors.Count == 1 ? FirstSurname(Authors[0])
                        : FirstSurname(Authors[0]) + " et al.";
            string year = Year.HasValue ? $" ({Year})" : "";
            string badge = Origin == CitationOrigin.Remote ? "[remote]"
                         : Origin == CitationOrigin.Project ? "[project]"
                         : "[library]";
            string miRNA = string.IsNullOrEmpty(MiRNA) ? "" : $" {MiRNA}";
            return $"{auth}{year}{miRNA} {Title} {badge}".Trim();
        }

        private static string FirstSurname(string full)
        {
            if (string.IsNullOrEmpty(full)) return "?";
            int i = full.IndexOf(',');
            return i >= 0 ? full.Substring(0, i).Trim() : full.Split(' ')[0];
        }
    }

    public enum CitationOrigin
    {
        Library,   // .refmanager.json
        Project,   // .mrrcite (zip → JSON)
        Remote     // HTTPS query
    }
}
