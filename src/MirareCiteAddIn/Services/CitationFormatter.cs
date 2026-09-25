// ============================================================================
//  CitationFormatter.cs — renders in-text citations and bibliography entries
//  in APA, MLA, Chicago, or numeric style.
//
//  This is intentionally a simple formatter (no CSL engine).  If you need
//  CSL-style validation or more journals, swap this class out — the rest
//  of the add-in depends only on the two public methods below.
// ============================================================================

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using MirareCiteAddIn.Models;

namespace MirareCiteAddIn.Services
{
    public class CitationFormatter
    {
        private readonly CitationStyle _style;

        public CitationFormatter(CitationStyle style) { _style = style; }

        // ─────────────────────────────────────────────────────────────────
        //  In-text citation — e.g. "(Smith et al., 2022)".
        // ─────────────────────────────────────────────────────────────────
        public string InText(Citation c)
        {
            switch (_style)
            {
                case CitationStyle.Apa:
                    return $"({AuthorInText(c)}, {YearOrNd(c)})";

                case CitationStyle.Mla:
                    return $"({FirstAuthorSurname(c)} {PagePlaceholder(c)})";

                case CitationStyle.Chicago:
                    return $"({FirstAuthorSurname(c)} {YearOrNd(c)})";

                case CitationStyle.Numeric:
                    // Numeric style uses the bibliography index, but at insert
                    // time we don't yet know the index — emit a placeholder
                    // that OnEditBibliography will renumber.
                    return $"[mirare:{c.Id}]";

                default:
                    return $"({AuthorInText(c)}, {YearOrNd(c)})";
            }
        }

        // ─────────────────────────────────────────────────────────────────
        //  Full bibliography entry, one per line.
        // ─────────────────────────────────────────────────────────────────
        public string BuildBibliography(IEnumerable<Citation> cited)
        {
            var sb = new StringBuilder();
            int n = 1;
            foreach (var c in cited ?? Enumerable.Empty<Citation>())
            {
                string entry = _style == CitationStyle.Numeric
                    ? $"[{n}]  {BibEntry(c)}"
                    : BibEntry(c);
                sb.AppendLine(entry);
                n++;
            }
            return sb.ToString();
        }

        // ─────────────────────────────────────────────────────────────────
        //  Full bib entry — APA / MLA / Chicago / Numeric share the same
        //  author+title+journal+year skeleton, just punctuated differently.
        //  miRNA + targetGene are appended in brackets so they round-trip
        //  for Mirare-style citations without breaking general formatting.
        // ─────────────────────────────────────────────────────────────────
        private string BibEntry(Citation c)
        {
            string authors = FormatAuthors(c.Authors);
            string year = YearOrNd(c);
            string title = c.Title ?? "";
            string journal = string.IsNullOrEmpty(c.Journal) ? "" : $". {c.Journal}";
            string doi = string.IsNullOrEmpty(c.Doi) ? "" : $". doi:{c.Doi}";
            string url = string.IsNullOrEmpty(c.Url) ? "" : $". {c.Url}";

            string miRNA = string.IsNullOrEmpty(c.MiRNA) ? "" :
                $" [miRNA: {c.MiRNA}";
            string target = string.IsNullOrEmpty(c.TargetGene) ? "" :
                $" → {c.TargetGene}]";
            if (!string.IsNullOrEmpty(miRNA) && string.IsNullOrEmpty(target))
                miRNA += "]";

            switch (_style)
            {
                case CitationStyle.Apa:
                case CitationStyle.Numeric:
                    return $"{authors} ({year}). {title}{journal}{doi}{url}{miRNA}{target}";

                case CitationStyle.Mla:
                    return $"{authors}. \"{title}.{journal}.\" {year}{doi}{url}{miRNA}{target}";

                case CitationStyle.Chicago:
                    return $"{authors}. \"{title}.{journal}\" {year}{doi}{url}{miRNA}{target}";

                default:
                    return $"{authors} ({year}). {title}{journal}{doi}{url}{miRNA}{target}";
            }
        }

        // ─────────────────────────────────────────────────────────────────
        //  Author helpers — comma-separated list with proper et al.
        //  truncation. APA cuts to "et al." after 3, MLA after 4, etc.
        // ─────────────────────────────────────────────────────────────────
        private string AuthorInText(Citation c)
        {
            if (c.Authors == null || c.Authors.Count == 0) return "Anonymous";
            if (c.Authors.Count == 1) return FirstAuthorSurname(c);
            if (c.Authors.Count == 2)
                return $"{FirstAuthorSurname(c)} & {SurnameOf(c.Authors[1])}";
            return $"{FirstAuthorSurname(c)} et al.";
        }

        private string FormatAuthors(List<string> authors)
        {
            if (authors == null || authors.Count == 0) return "Anonymous";
            if (authors.Count == 1) return authors[0];
            if (authors.Count <= 3)
                return string.Join(", ", authors.Take(authors.Count - 1))
                       + ", & " + authors[authors.Count - 1];
            // 4+ — list first three + "et al." (APA-style truncation; tweak
            // per-style if you need stricter CSL rules).
            return string.Join(", ", authors.Take(3)) + ", et al.";
        }

        private static string FirstAuthorSurname(Citation c)
            => c.Authors == null || c.Authors.Count == 0 ? "Anonymous"
             : SurnameOf(c.Authors[0]);

        private static string SurnameOf(string full)
        {
            if (string.IsNullOrEmpty(full)) return "?";
            int i = full.IndexOf(',');
            return i >= 0 ? full.Substring(0, i).Trim() : full.Split(' ')[0];
        }

        private static string YearOrNd(Citation c)
            => c.Year.HasValue ? c.Year.Value.ToString() : "n.d.";

        private static string PagePlaceholder(Citation c)
            => c.Year.HasValue ? c.Year.Value.ToString() : "n.pag.";
    }
}
