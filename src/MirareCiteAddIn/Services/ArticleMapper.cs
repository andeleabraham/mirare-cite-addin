// ============================================================================
//  ArticleMapper.cs — maps one article object from the real Mirare data
//  files into a Citation.
//
//  Both file types share the same article shape (Crossref-style, as written
//  by the Mirare PySide6 app):
//
//    refmanager.json  (library):  { "articles": [ {...}, ... ] }
//    .mrrcite         (project):  zip → project.json =
//                                 { "<ProjectName>": { "articles": [...] } }
//
//  Article fields observed in the wild (all optional, various types):
//    DOI, URL, title (string OR array), author (array of
//    {family, given, sequence}), authors (array of plain strings — the
//    app's flattened copy), container-title (array), journal (string),
//    published-print / published-online / created (object with
//    date-parts: [[y,m,d]]), year (number), unique_id, page/pages,
//    venue, publicationVenue, notes, collections.
//
//  Parsing is deliberately defensive: the library JSON comes from several
//  fetch platforms, so field presence and types vary per entry.
//
//  NOTE: JSON is Newtonsoft.Json, not System.Text.Json — the latter's
//  .NET Framework dependency chain (System.Runtime.CompilerServices.Unsafe
//  binding redirects) cannot be satisfied inside a COM host like WINWORD.
// ============================================================================

using System;
using System.Collections.Generic;
using Newtonsoft.Json.Linq;
using MirareCiteAddIn.Models;

namespace MirareCiteAddIn.Services
{
    internal static class ArticleMapper
    {
        /// <summary>Maps an array of article objects to Citations.</summary>
        public static List<Citation> FromArticles(JArray articles, CitationOrigin origin)
        {
            var list = new List<Citation>();
            foreach (var item in articles)
            {
                try
                {
                    if (item is JObject obj)
                    {
                        var c = Map(obj, origin);
                        if (c != null) list.Add(c);
                    }
                }
                catch
                {
                    // A malformed entry must never break loading the whole
                    // library — skip it.
                }
            }
            return list;
        }

        /// <summary>Maps a single article object to a Citation (null if unusable).</summary>
        public static Citation Map(JObject e, CitationOrigin origin)
        {
            string title = GetTitle(e);
            string id = GetStr(e, "unique_id");
            string doi = GetStr(e, "DOI") ?? GetStr(e, "doi");
            if (string.IsNullOrEmpty(id)) id = doi;
            if (string.IsNullOrEmpty(id)) id = Guid.NewGuid().ToString("N");
            if (string.IsNullOrEmpty(title) && string.IsNullOrEmpty(doi)) return null;

            return new Citation
            {
                Id = id,
                Title = title,
                Authors = GetAuthors(e),
                Year = GetYear(e),
                Journal = GetJournal(e),
                Doi = doi,
                Url = GetStr(e, "URL") ?? GetStr(e, "url"),
                Pages = GetStr(e, "page") ?? GetStr(e, "pages"),
                Volume = GetStr(e, "volume"),
                Issue = GetStr(e, "issue"),
                // Default to journal article — most Mirare records are, and
                // CSL styles condition volume/issue/pages on this type.
                Type = GetStr(e, "type") ?? "journal-article",
                Origin = origin
            };
        }

        // ── field extractors ──────────────────────────────────────────────

        /// <summary>Case-insensitive property lookup — the files mix
        /// "DOI"/"URL" casing with lowercase app fields.</summary>
        private static JToken Prop(JObject e, string name)
            => e.GetValue(name, StringComparison.OrdinalIgnoreCase);

        private static string GetStr(JObject e, string name)
            => Prop(e, name) is JValue v && v.Type == JTokenType.String ? (string)v : null;

        private static string GetTitle(JObject e)
        {
            var t = Prop(e, "title");
            if (t == null) return null;
            if (t.Type == JTokenType.String) return (string)t;
            if (t is JArray arr)
            {
                foreach (var part in arr)
                    if (part.Type == JTokenType.String && !string.IsNullOrWhiteSpace((string)part))
                        return (string)part;
            }
            return null;
        }

        private static string GetJournal(JObject e)
        {
            string j = GetStr(e, "journal");
            if (!string.IsNullOrEmpty(j)) return j;
            if (Prop(e, "container-title") is JArray ct)
            {
                foreach (var part in ct)
                    if (part.Type == JTokenType.String && !string.IsNullOrWhiteSpace((string)part))
                        return (string)part;
            }
            return GetStr(e, "venue") ?? GetStr(e, "publicationVenue");
        }

        /// <summary>Year: explicit "year" number first, then the date-parts
        /// of published-print → published-online → created → deposited.</summary>
        private static int? GetYear(JObject e)
        {
            var y = Prop(e, "year");
            if (y != null && y.Type == JTokenType.Integer) return (int)y;

            foreach (var container in new[] { "published-print", "published-online", "created", "deposited" })
            {
                if (!(Prop(e, container) is JObject c)) continue;
                if (!(c.GetValue("date-parts", StringComparison.OrdinalIgnoreCase) is JArray dp)) continue;

                // date-parts: [[2024, 6, 5]] — first number anywhere is the year.
                foreach (var outer in dp)
                {
                    if (!(outer is JArray inner)) continue;
                    foreach (var part in inner)
                    {
                        if (part.Type == JTokenType.Integer)
                        {
                            int v = (int)part;
                            if (v >= 1400 && v <= 2200) return v;
                            break;  // first component wasn't a sane year — stop
                        }
                        break;
                    }
                }
            }

            // e.g. "publicationDate": "2020-06-02"
            var s = GetStr(e, "publicationDate");
            if (s != null && s.Length >= 4 && int.TryParse(s.Substring(0, 4), out int py))
                return py;

            return null;
        }

        /// <summary>Authors: prefer structured "author" objects
        /// ({family, given}) rendered as "Family, Given"; fall back to the
        /// app's flattened "authors" string list.</summary>
        private static List<string> GetAuthors(JObject e)
        {
            var authors = new List<string>();
            if (Prop(e, "author") is JArray a)
            {
                foreach (var item in a)
                {
                    if (!(item is JObject p)) continue;
                    string family = GetStr(p, "family") ?? GetStr(p, "name") ?? "";
                    string given = GetStr(p, "given") ?? "";
                    string full = string.IsNullOrEmpty(given) ? family : $"{family}, {given}";
                    if (!string.IsNullOrWhiteSpace(full)) authors.Add(full);
                }
            }
            if (authors.Count == 0 && Prop(e, "authors") is JArray flat)
            {
                foreach (var item in flat)
                    if (item.Type == JTokenType.String && !string.IsNullOrWhiteSpace((string)item))
                        authors.Add((string)item);
            }
            return authors;
        }
    }
}
