// ============================================================================
//  CslStyle.cs — a compact CSL 1.0.1 renderer.
//
//  Reads the same .csl files the Mirare desktop app uses (its styles folder)
//  and renders in-text citations + bibliography entries from the style's
//  <citation>/<bibliography> layouts, resolving macros the way citeproc does.
//
//  This is deliberately a SUBSET of full CSL — the parts that matter for
//  author-date and numeric styles:
//    macro / text (variable, value, macro) / name (et-al-min, et-al-use-first,
//    initialize-with, and=, delimiter) / date (year) / group (delimiter,
//    prefix, suffix) / choose-if-else (variable presence only) /
//    citation-number.
//  Locale dictionaries, plural labels, collapse/cite-group disambiguation
//  and page-range formatting are NOT implemented — those fall back to sane
//  defaults. Good enough to match the app's styles for everyday articles.
// ============================================================================

using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml;
using MirareCiteAddIn.Models;

namespace MirareCiteAddIn.Services
{
    public class CslStyle
    {
        private const string CslNs = "http://purl.org/net/xbiblio/csl";

        private readonly XmlDocument _doc;
        private readonly XmlNamespaceManager _ns;
        private readonly XmlElement _root;
        private readonly string _filePath;

        public string CitationDelimiter { get; private set; } = "; ";
        public string DisplayName { get; private set; } = "CSL style";

        private CslStyle(string path)
        {
            _filePath = path;
            _doc = new XmlDocument();
            using (var s = File.OpenRead(path))
                _doc.Load(s);
            _root = _doc.DocumentElement;
            _ns = new XmlNamespaceManager(_doc.NameTable);
            _ns.AddNamespace("c", CslNs);

            var info = _root.SelectSingleNode("c:info/c:title", _ns) as XmlElement;
            if (info != null) DisplayName = info.InnerText.Trim();

            var citation = _root.SelectSingleNode("c:citation", _ns) as XmlElement;
            if (citation != null && !string.IsNullOrEmpty(citation.GetAttribute("delimiter")))
                CitationDelimiter = citation.GetAttribute("delimiter");
        }

        public static CslStyle Load(string path)
        {
            var style = new CslStyle(path);
            // Fail fast if there is nothing we can render.
            if (style._root.SelectSingleNode("c:citation/c:layout", style._ns) == null &&
                style._root.SelectSingleNode("c:bibliography/c:layout", style._ns) == null)
                throw new InvalidDataException(
                    $"'{Path.GetFileName(path)}' has no citation/bibliography layout — not a renderable CSL style.");
            return style;
        }

        /// <summary>True for numeric styles (IEEE, Vancouver…) whose in-text
        /// citation is just the reference number — we can't know that number
        /// at insert time, so the caller renders a placeholder instead.</summary>
        public bool IsNumeric
        {
            get
            {
                var layout = _root.SelectSingleNode("c:citation/c:layout", _ns) as XmlElement;
                return layout != null &&
                       layout.InnerXml.IndexOf("citation-number", StringComparison.OrdinalIgnoreCase) >= 0;
            }
        }

        // ── public rendering API ──────────────────────────────────────────

        /// <summary>In-text citation WITHOUT surrounding parentheses (the
        /// caller places the parens inside the field result).</summary>
        public string RenderInText(Citation c)
        {
            var layout = _root.SelectSingleNode("c:citation/c:layout", _ns) as XmlElement;
            if (layout == null) return FallbackInText(c);
            // APA-style layouts carry the parentheses themselves as layout
            // prefix/suffix — apply them around the rendered content.
            return (Attr(layout, "prefix")
                  + RenderChildren(layout, c, 0, Attr(layout, "delimiter"), 0).Trim()
                  + Attr(layout, "suffix")).Trim();
        }

        /// <summary>One formatted bibliography entry. 'number' feeds the
        /// citation-number variable used by numeric styles.</summary>
        public string RenderBibliographyEntry(Citation c, int number)
        {
            var layout = _root.SelectSingleNode("c:bibliography/c:layout", _ns) as XmlElement;
            if (layout == null) return FallbackBibEntry(c);
            string entry = RenderChildren(layout, c, number, Attr(layout, "delimiter"), 0).Trim();
            // Numeric layouts render "[1]Title" — give the number breathing room.
            return System.Text.RegularExpressions.Regex.Replace(entry, @"^\[\d+\](?=\S)", "$0 ");
        }

        // ── fallbacks when a style can't be rendered ─────────────────────

        private static string FallbackInText(Citation c)
        {
            string auth = c.Authors.Count == 0 ? "Anonymous"
                        : c.Authors.Count == 1 ? Surname(c.Authors[0])
                        : Surname(c.Authors[0]) + " et al.";
            return $"{auth}, {c.Year?.ToString(CultureInfo.InvariantCulture) ?? "n.d."}";
        }

        private static string FallbackBibEntry(Citation c)
        {
            string authors = c.Authors.Count == 0 ? "Anonymous" : string.Join(", ", c.Authors);
            return $"{authors} ({c.Year}). {c.Title}. {c.Journal}. {c.Doi}";
        }

        // ── renderer core ─────────────────────────────────────────────────

        private string RenderChildren(XmlElement parent, Citation c, int number,
                                      string joinDelimiter, int depth)
        {
            var parts = new List<string>();
            foreach (XmlNode node in parent.ChildNodes)
            {
                if (node.NodeType != XmlNodeType.Element) continue;
                var el = (XmlElement)node;
                string rendered = RenderElement(el, c, number, depth);
                if (!string.IsNullOrEmpty(rendered))
                    parts.Add(rendered);
            }
            return string.Join(joinDelimiter ?? "", parts);
        }

        private string RenderElement(XmlElement el, Citation c, int number, int depth)
        {
            switch (el.LocalName)
            {
                case "text": return RenderText(el, c, number, depth);
                case "name": return RenderName(el, c);
                case "names": return RenderNamesElement(el, c, depth);
                case "date": return RenderDate(el, c);
                case "group": return RenderGroup(el, c, number, depth);
                case "choose": return RenderChoose(el, c, number, depth);
                case "number":
                    if (Attr(el, "variable") == "citation-number")
                        return number > 0 ? number.ToString(CultureInfo.InvariantCulture) : "";
                    return "";
                case "label": return "";     // plural labels not implemented
                case "key": return "";       // sort keys ignored
                default: return "";
            }
        }

        /// <summary>CSL wraps <name> in <names variable="author editor …"> —
        /// render the name config inside for the first variable we supply.
        /// When the variable isn't one we carry, CSL's <substitute> chain
        /// decides the fallback (APA's author-intext tries "composer" first
        /// and falls back to "author"), so always attempt substitutes.
        /// fallbackNameEl: a <name> config inherited into substitute names
        /// that carry none of their own.</summary>
        private string RenderNamesElement(XmlElement el, Citation c, int depth,
                                          XmlElement fallbackNameEl = null)
        {
            string var = Attr(el, "variable");
            if (string.IsNullOrEmpty(var)) var = "author";
            // We only carry author data — editor/translator names have no
            // model and must render empty (APA bib has editor-only groups
            // that must disappear).
            bool supplies = var.Split(' ').Any(v => v == "author");

            var nameEl = el.SelectSingleNode("c:name", _ns) as XmlElement ?? fallbackNameEl;
            string body = supplies ? RenderName(nameEl, c) : "";
            if (!string.IsNullOrEmpty(body)) return Wrap(el, body);

            // <substitute> covers entries without the primary variable —
            // CSL renders the FIRST substitute that yields content, not all.
            var sub = el.SelectSingleNode("c:substitute", _ns) as XmlElement;
            if (sub == null) return "";
            foreach (XmlNode node in sub.ChildNodes)
            {
                if (node.NodeType != XmlNodeType.Element) continue;
                var se = (XmlElement)node;
                string rendered = se.LocalName == "names"
                    ? RenderNamesElement(se, c, depth + 1, nameEl)
                    : RenderElement(se, c, 0, depth + 1);
                if (!string.IsNullOrEmpty(rendered)) return rendered;
            }
            return "";
        }

        private string RenderGroup(XmlElement el, Citation c, int number, int depth)
        {
            string inner = RenderChildren(el, c, number, Attr(el, "delimiter"), depth);
            if (string.IsNullOrEmpty(inner)) return "";
            return Wrap(el, inner);
        }

        private string RenderChoose(XmlElement el, Citation c, int number, int depth)
        {
            foreach (XmlNode branch in el.ChildNodes)
            {
                if (branch.NodeType != XmlNodeType.Element) continue;
                var be = (XmlElement)branch;
                if (be.LocalName == "else" || ConditionsMatch(be, c))
                    return RenderChildren(be, c, number, null, depth + 1);
            }
            return "";
        }

        private bool ConditionsMatch(XmlElement branch, Citation c)
        {
            if (branch.LocalName != "if" && branch.LocalName != "else-if") return false;

            // match="all" (default) / "any" / "none" over the listed variables.
            string match = Attr(branch, "match");
            if (string.IsNullOrEmpty(match)) match = "all";

            string vars = Attr(branch, "variable");
            var present = (vars ?? "").Split(' ')
                .Where(v => v.Length > 0)
                .Select(v => VariableValue(v, c, 0) != "")
                .ToList();

            // @type/@version conditions are not modelled — a branch with only
            // those never matches (we fall through to else-if/else).
            if (present.Count == 0) return false;

            switch (match)
            {
                case "any": return present.Any(p => p);
                case "none": return present.All(p => !p);
                default: return present.All(p => p);
            }
        }

        private string RenderText(XmlElement el, Citation c, int number, int depth)
        {
            string prefix = Attr(el, "prefix");
            string suffix = Attr(el, "suffix");
            string inner;

            string macro = Attr(el, "macro");
            if (!string.IsNullOrEmpty(macro))
                inner = RenderMacro(macro, c, number, depth);
            else if (!string.IsNullOrEmpty(Attr(el, "variable")))
                inner = VariableValue(Attr(el, "variable"), c, number);
            else if (!string.IsNullOrEmpty(Attr(el, "value")))
                inner = Attr(el, "value");
            else
                inner = "";

            if (string.IsNullOrEmpty(inner)) return "";
            return prefix + inner + suffix;
        }

        private string RenderMacro(string name, Citation c, int number, int depth)
        {
            if (depth > 10) return "";                       // macro recursion guard
            var macro = _root.SelectSingleNode($"c:macro[@name='{name}']", _ns) as XmlElement;
            if (macro == null) return "";
            return RenderChildren(macro, c, number, Attr(macro, "delimiter"), depth + 1);
        }

        private string VariableValue(string variable, Citation c, int number)
        {
            switch ((variable ?? "").ToLowerInvariant())
            {
                case "citation-number": return number > 0 ? number.ToString(CultureInfo.InvariantCulture) : "";
                case "title": return c.Title ?? "";
                case "container-title": return c.Journal ?? "";
                case "doi": return c.Doi ?? "";
                case "url": return c.Url ?? "";
                case "page": return c.Pages ?? "";
                case "volume": return "";                        // not modelled yet
                case "issue": return "";
                case "publisher": return "";
                case "author": case "editor": return RenderName(null, c);   // bare variable form
                case "issued": return c.Year?.ToString(CultureInfo.InvariantCulture) ?? "";
                default: return "";
            }
        }

        // ── name rendering ────────────────────────────────────────────────

        private string RenderName(XmlElement el, Citation c)
        {
            if (c.Authors == null || c.Authors.Count == 0) return "";

            int etAlMin = GetInt(el, "et-al-min", int.MaxValue);
            int etAlUseFirst = GetInt(el, "et-al-use-first", 1);
            string initializeWith = Attr(el, "initialize-with");
            string delimiter = el != null && !string.IsNullOrEmpty(Attr(el, "delimiter"))
                ? Attr(el, "delimiter") : ", ";

            var rendered = new List<string>();
            int limit = c.Authors.Count >= etAlMin ? etAlUseFirst : c.Authors.Count;
            for (int i = 0; i < Math.Min(limit, c.Authors.Count); i++)
                rendered.Add(FormatPerson(c.Authors[i], initializeWith));

            string body;
            if (c.Authors.Count >= etAlMin && rendered.Count < c.Authors.Count)
            {
                body = rendered.Count > 0 ? string.Join(delimiter, rendered) + " et al." : "et al.";
            }
            else
            {
                string and = el != null ? Attr(el, "and") : "";
                if (rendered.Count <= 2)
                {
                    body = rendered.Count == 2
                        ? rendered[0] + JoinAnd(and) + rendered[1]
                        : rendered.Count == 1 ? rendered[0] : "";
                }
                else
                {
                    string head = string.Join(delimiter, rendered.Take(rendered.Count - 1));
                    body = head + JoinAnd(and) + rendered[rendered.Count - 1];
                }
            }
            return el != null ? Wrap(el, body) : body;
        }

        private static string JoinAnd(string andAttr)
        {
            if (andAttr == "symbol") return " & ";
            if (andAttr == "none" || string.IsNullOrEmpty(andAttr)) return ", ";
            return " and ";
        }

        /// <summary>Authors are stored as "Family, Given".</summary>
        private static string FormatPerson(string author, string initializeWith)
        {
            if (string.IsNullOrWhiteSpace(author)) return "";
            int comma = author.IndexOf(',');
            string family = comma >= 0 ? author.Substring(0, comma).Trim() : author.Trim();
            string given = comma >= 0 ? author.Substring(comma + 1).Trim() : "";

            if (given.Length > 0 && initializeWith != null)
            {
                // initialize-with=". " → "T."; "" → bare initial
                string init = initializeWith.TrimEnd();
                given = string.Join(" ", given
                    .Split(new[] { ' ', '-', '.' }, StringSplitOptions.RemoveEmptyEntries)
                    .Select(p => p[0] + init));
            }
            return string.IsNullOrEmpty(given) ? family : $"{family}, {given}";
        }

        // ── date rendering ────────────────────────────────────────────────

        private string RenderDate(XmlElement el, Citation c)
        {
            // Only the primary issue date maps to our model — CSL styles
            // also render original-date/accessed, which we don't carry
            // (APA's date-intext renders "original/issued" and would emit
            // the year twice without this guard).
            string var = Attr(el, "variable");
            if (var != "" && var != "issued") return "";
            string year = c.Year?.ToString(CultureInfo.InvariantCulture) ?? "";
            if (string.IsNullOrEmpty(year)) return "";
            return Wrap(el, year);       // year-only — full date forms not modelled
        }

        // ── helpers ───────────────────────────────────────────────────────

        private static string Wrap(XmlElement el, string inner)
        {
            string prefix = Attr(el, "prefix");
            string suffix = Attr(el, "suffix");
            return prefix + inner + suffix;
        }

        private static string Attr(XmlElement el, string name)
            => el?.GetAttribute(name) ?? "";

        private static int GetInt(XmlElement el, string name, int fallback)
        {
            string v = Attr(el, name);
            return int.TryParse(v, out int n) ? n : fallback;
        }

        private static string Surname(string author)
        {
            int comma = (author ?? "").IndexOf(',');
            return comma >= 0 ? author.Substring(0, comma).Trim() : (author ?? "").Trim();
        }
    }
}
