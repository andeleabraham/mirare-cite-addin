// ============================================================================
//  CslStyle.cs — a compact CSL 1.0.1 renderer.
//
//  Reads the same .csl files the Mirare desktop app uses (its styles folder)
//  and renders in-text citations + bibliography entries from the style's
//  <citation>/<bibliography> layouts, resolving macros the way citeproc does.
//
//  This is deliberately a SUBSET of full CSL — the parts that matter for
//  author-date and numeric styles:
//    macro / text (variable, value, macro) / name (form=short, et-al-min,
//    et-al-use-first, initialize-with, and=, delimiter,
//    delimiter-precedes-last) / date (year) / group (delimiter, prefix,
//    suffix, font-style, font-weight) / choose (type + variable conditions
//    with match=any|all|none) / citation-number / bibliography <sort>.
//
//  RICH TEXT: elements with font-style="italic" / font-weight="bold"
//  (journal names, volume — per the loaded .csl) are emitted wrapped in
//  marker characters; the add-in converts the markers to real Word
//  character formatting inside the field result.
//
//  Not implemented: locale dictionaries, plural labels, page-range
//  collapsing, name disambiguation — those fall back to sane defaults.
// ============================================================================

using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;
using MirareCiteAddIn.Models;

namespace MirareCiteAddIn.Services
{
    public class CslStyle
    {
        private const string CslNs = "http://purl.org/net/xbiblio/csl";
        private const char ItalicOn = '\u0001', ItalicOff = '\u0002';
        private const char BoldOn = '\u0003', BoldOff = '\u0004';

        private readonly struct Fmt
        {
            public readonly bool Italic;
            public readonly bool Bold;
            public Fmt(bool i, bool b) { Italic = i; Bold = b; }
        }

        private static readonly Fmt Plain = new Fmt(false, false);

        private readonly XmlDocument _doc;
        private readonly XmlNamespaceManager _ns;
        private readonly XmlElement _root;

        public string CitationDelimiter { get; private set; } = "; ";
        public string DisplayName { get; private set; } = "CSL style";

        // et-al context of the layout currently being rendered — macros are
        // defined at style level, so the ancestor walk from <name> never
        // reaches <citation et-al-min="3">; we thread it explicitly instead.
        private int _ctxEtAlMin = int.MaxValue;
        private int _ctxEtAlUseFirst = 1;

        private CslStyle(string path)
        {
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

        public string RenderInText(Citation c)
        {
            var layout = _root.SelectSingleNode("c:citation/c:layout", _ns) as XmlElement;
            if (layout == null) return FallbackInText(c);
            var citationEl = _root.SelectSingleNode("c:citation", _ns) as XmlElement;
            _ctxEtAlMin = GetInt(citationEl, "et-al-min", int.MaxValue);
            _ctxEtAlUseFirst = GetInt(citationEl, "et-al-use-first", 1);
            // APA-style layouts carry the parentheses themselves as layout
            // prefix/suffix — apply them around the rendered content.
            return (Attr(layout, "prefix")
                  + RenderChildren(layout, c, 0, Attr(layout, "delimiter"), 0, Plain).Trim()
                  + Attr(layout, "suffix")).Trim();
        }

        /// <summary>Renders a multi-citation group the way CSL intends: the
        /// layout's parentheses (prefix/suffix) appear ONCE around the whole
        /// group, individual citations joined by the layout delimiter —
        /// "(A; B; C)", never "(A) (B) (C)".</summary>
        public string RenderInTextGroup(IList<Citation> list, IList<int> numbers = null)
        {
            var layout = _root.SelectSingleNode("c:citation/c:layout", _ns) as XmlElement;
            if (layout == null || list == null || list.Count == 0)
                return FallbackInText(list != null && list.Count > 0 ? list[0] : new Citation());
            var citationEl = _root.SelectSingleNode("c:citation", _ns) as XmlElement;
            _ctxEtAlMin = GetInt(citationEl, "et-al-min", int.MaxValue);
            _ctxEtAlUseFirst = GetInt(citationEl, "et-al-use-first", 1);

            string delim = Attr(layout, "delimiter");
            if (string.IsNullOrEmpty(delim)) delim = "; ";

            var parts = new List<string>();
            for (int i = 0; i < list.Count; i++)
            {
                int num = numbers != null && i < numbers.Count ? numbers[i] : 0;
                string p = RenderChildren(layout, list[i], num, null, 0, Plain).Trim();
                if (!string.IsNullOrEmpty(p)) parts.Add(p);
            }
            return Attr(layout, "prefix") + string.Join(delim, parts) + Attr(layout, "suffix");
        }

        /// <summary>In-text rendering with a known citation number — used by
        /// the renumber pass for numeric styles ([2], [3] …).</summary>
        public string RenderInText(Citation c, int number)
        {
            var layout = _root.SelectSingleNode("c:citation/c:layout", _ns) as XmlElement;
            if (layout == null) return $"[{number}]";
            var citationEl = _root.SelectSingleNode("c:citation", _ns) as XmlElement;
            _ctxEtAlMin = GetInt(citationEl, "et-al-min", int.MaxValue);
            _ctxEtAlUseFirst = GetInt(citationEl, "et-al-use-first", 1);
            return (Attr(layout, "prefix")
                  + RenderChildren(layout, c, number, Attr(layout, "delimiter"), 0, Plain).Trim()
                  + Attr(layout, "suffix")).Trim();
        }

        /// <summary>One formatted bibliography entry (with rich-text markers).
        /// 'number' feeds the citation-number variable of numeric styles.</summary>
        public string RenderBibliographyEntry(Citation c, int number)
        {
            var layout = _root.SelectSingleNode("c:bibliography/c:layout", _ns) as XmlElement;
            if (layout == null) return FallbackBibEntry(c);
            var bibEl = _root.SelectSingleNode("c:bibliography", _ns) as XmlElement;
            _ctxEtAlMin = GetInt(bibEl, "et-al-min", int.MaxValue);
            _ctxEtAlUseFirst = GetInt(bibEl, "et-al-use-first", 1);
            string entry = RenderChildren(layout, c, number, Attr(layout, "delimiter"), 0, Plain).Trim();
            // Numeric layouts render "[1]Title" — give the number breathing room.
            entry = Regex.Replace(entry, @"^\[\d+\](?=\S)", "$0 ");
            // Collapse stray double periods ("S.." where a name-initial period
            // meets the layout's sentence period).
            while (entry.Contains("..")) entry = entry.Replace("..", ".");
            return entry;
        }

        /// <summary>
        /// Orders bibliography entries by the style's &lt;bibliography&gt;&lt;sort&gt;
        /// keys (APA: author, then date, then title). Styles without a &lt;sort&gt;
        /// element (IEEE) keep citation order — numbers stay aligned.
        /// </summary>
        public List<Citation> SortBibliography(List<Citation> citations)
        {
            var sortNode = _root.SelectSingleNode("c:bibliography/c:sort", _ns) as XmlElement;
            if (sortNode == null || citations.Count < 2) return citations;

            var keyEls = new List<XmlElement>();
            foreach (XmlNode node in sortNode.ChildNodes)
                if (node.NodeType == XmlNodeType.Element && ((XmlElement)node).LocalName == "key")
                    keyEls.Add((XmlElement)node);
            if (keyEls.Count == 0) return citations;

            // Precompute key values: one string per <key> per citation.
            var keyed = new List<KeyValuePair<Citation, List<string>>>(citations.Count);
            foreach (var c in citations)
            {
                var values = new List<string>(keyEls.Count);
                foreach (var k in keyEls)
                {
                    string macro = Attr(k, "macro");
                    string v = !string.IsNullOrEmpty(macro)
                        ? RenderMacro(macro, c, 0, 0, Plain)
                        : VariableValue(Attr(k, "variable"), c, 0);
                    values.Add(StripMarkers(v ?? ""));
                }
                keyed.Add(new KeyValuePair<Citation, List<string>>(c, values));
            }

            var directions = keyEls.Select(k => Attr(k, "sort") == "descending" ? -1 : 1).ToList();

            int Compare(KeyValuePair<Citation, List<string>> a, KeyValuePair<Citation, List<string>> b)
            {
                for (int i = 0; i < keyEls.Count; i++)
                {
                    string va = a.Value[i], vb = b.Value[i];
                    if (string.Equals(va, vb, StringComparison.OrdinalIgnoreCase)) continue;
                    int cmp;
                    if (double.TryParse(va, NumberStyles.Any, CultureInfo.InvariantCulture, out double na) &&
                        double.TryParse(vb, NumberStyles.Any, CultureInfo.InvariantCulture, out double nb))
                        cmp = na.CompareTo(nb);
                    else
                        cmp = string.Compare(va, vb, StringComparison.OrdinalIgnoreCase);
                    return cmp * directions[i];
                }
                return 0;
            }

            keyed.Sort((a, b) => Compare(a, b));
            return keyed.Select(k => k.Key).ToList();
        }

        // ── renderer core ─────────────────────────────────────────────────

        private string RenderChildren(XmlElement parent, Citation c, int number,
                                      string joinDelimiter, int depth, Fmt fmt)
        {
            var parts = new List<string>();
            foreach (XmlNode node in parent.ChildNodes)
            {
                if (node.NodeType != XmlNodeType.Element) continue;
                string rendered = RenderElement((XmlElement)node, c, number, depth, fmt);
                if (!string.IsNullOrEmpty(rendered))
                    parts.Add(rendered);
            }
            return string.Join(joinDelimiter ?? "", parts);
        }

        private string RenderElement(XmlElement el, Citation c, int number, int depth, Fmt fmt)
        {
            // Formatting attributes inherit down the tree, per element.
            var eff = new Fmt(
                fmt.Italic || Attr(el, "font-style") == "italic",
                fmt.Bold || Attr(el, "font-weight") == "bold");

            switch (el.LocalName)
            {
                case "text": return RenderText(el, c, number, depth, eff);
                case "name": return RenderName(el, c, eff);
                case "names": return RenderNamesElement(el, c, depth, null, eff);
                case "date": return RenderDate(el, c, eff);
                case "group": return RenderGroup(el, c, number, depth, eff);
                case "choose": return RenderChoose(el, c, number, depth, eff);
                case "number":
                    if (Attr(el, "variable") == "citation-number" && number > 0)
                        return Mark(number.ToString(CultureInfo.InvariantCulture), eff);
                    return "";
                case "label": return "";     // plural labels not implemented
                case "key": return "";       // sort keys ignored
                default: return "";
            }
        }

        private string RenderGroup(XmlElement el, Citation c, int number, int depth, Fmt fmt)
        {
            string inner = RenderChildren(el, c, number, Attr(el, "delimiter"), depth, fmt);
            if (string.IsNullOrEmpty(inner)) return "";
            return Wrap(el, inner);
        }

        private string RenderChoose(XmlElement el, Citation c, int number, int depth, Fmt fmt)
        {
            foreach (XmlNode branch in el.ChildNodes)
            {
                if (branch.NodeType != XmlNodeType.Element) continue;
                var be = (XmlElement)branch;
                if (be.LocalName == "else" || ConditionsMatch(be, c))
                    return RenderChildren(be, c, number, null, depth + 1, fmt);
            }
            return "";
        }

        /// <summary>CSL type for the citation — styles condition volume/
        /// issue/pages groups on @type="article-journal" etc. Our data uses
        /// Crossref type names, so map them.</summary>
        private static string MapType(Citation c)
        {
            string raw = (c.Type ?? "").Trim().ToLowerInvariant();
            if (raw.Length == 0) raw = "article-journal";
            switch (raw)
            {
                case "journal-article": return "article-journal";
                case "proceedings-article": return "paper-conference";
                case "book-chapter": return "chapter";
                case "posted-content": return "article";
                case "report-component": return "report";
                default: return raw;
            }
        }

        private bool ConditionsMatch(XmlElement branch, Citation c)
        {
            if (branch.LocalName != "if" && branch.LocalName != "else-if") return false;

            var conditions = new List<bool>();

            // @type — compare against the mapped CSL type of the citation.
            string types = Attr(branch, "type");
            if (!string.IsNullOrEmpty(types))
            {
                string ct = MapType(c);
                foreach (var t in types.Split(' '))
                    if (t.Length > 0) conditions.Add(ct == t.ToLowerInvariant());
            }

            // @variable — true when the citation carries the variable.
            string vars = Attr(branch, "variable");
            if (!string.IsNullOrEmpty(vars))
                foreach (var v in vars.Split(' '))
                    if (v.Length > 0) conditions.Add(VariableValue(v, c, 0) != "");

            if (conditions.Count == 0) return false;

            string match = Attr(branch, "match");
            if (string.IsNullOrEmpty(match)) match = "all";

            switch (match)
            {
                case "any": return conditions.Any(p => p);
                case "none": return conditions.All(p => !p);
                default: return conditions.All(p => p);
            }
        }

        private string RenderText(XmlElement el, Citation c, int number, int depth, Fmt fmt)
        {
            string prefix = Attr(el, "prefix");
            string suffix = Attr(el, "suffix");
            string inner;

            string macro = Attr(el, "macro");
            if (!string.IsNullOrEmpty(macro))
                inner = RenderMacro(macro, c, number, depth, fmt);
            else if (!string.IsNullOrEmpty(Attr(el, "variable")))
                inner = VariableValue(Attr(el, "variable"), c, number);
            else if (!string.IsNullOrEmpty(Attr(el, "value")))
                inner = Attr(el, "value");
            else
                inner = "";

            if (string.IsNullOrEmpty(inner)) return "";
            return prefix + Mark(inner, fmt) + suffix;
        }

        private string RenderMacro(string name, Citation c, int number, int depth, Fmt fmt)
        {
            if (depth > 10) return "";                       // macro recursion guard
            var macro = _root.SelectSingleNode($"c:macro[@name='{name}']", _ns) as XmlElement;
            if (macro == null) return "";
            return RenderChildren(macro, c, number, Attr(macro, "delimiter"), depth + 1, fmt);
        }

        /// <summary>CSL wraps <name> in <names variable="author editor …"> —
        /// render the name config inside for the first variable we supply.
        /// When the variable isn't one we carry, CSL's <substitute> chain
        /// decides the fallback (APA's author-intext tries "composer" first
        /// and falls back to "author"), so always attempt substitutes.
        /// fallbackNameEl: a <name> config inherited into substitute names
        /// that carry none of their own.</summary>
        private string RenderNamesElement(XmlElement el, Citation c, int depth,
                                          XmlElement fallbackNameEl, Fmt fmt)
        {
            string var = Attr(el, "variable");
            if (string.IsNullOrEmpty(var)) var = "author";
            // We only carry author data — editor/translator names have no
            // model and must render empty (APA bib has editor-only groups
            // that must disappear).
            bool supplies = var.Split(' ').Any(v => v == "author");

            var nameEl = el.SelectSingleNode("c:name", _ns) as XmlElement ?? fallbackNameEl;
            string body = supplies ? RenderName(nameEl, c, fmt) : "";
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
                    ? RenderNamesElement(se, c, depth + 1, nameEl, fmt)
                    : RenderElement(se, c, 0, depth + 1, fmt);
                if (!string.IsNullOrEmpty(rendered)) return rendered;
            }
            return "";
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
                case "volume": return c.Volume ?? "";
                case "issue": return c.Issue ?? "";
                case "publisher": return "";
                case "author": case "editor": return RenderName(null, c, Plain);   // bare variable form
                case "issued": return c.Year?.ToString(CultureInfo.InvariantCulture) ?? "";
                default: return "";
            }
        }

        // ── name rendering ────────────────────────────────────────────────

        private string RenderName(XmlElement el, Citation c, Fmt fmt)
        {
            if (c.Authors == null || c.Authors.Count == 0) return "";

            // et-al rules may sit on <name> itself or be inherited from
            // <names>/<citation>/<bibliography> (APA declares them there —
            // in-text et-al-min=3, bibliography et-al-min=21). Macros live
            // outside those elements, so _ctx* carries the layout context.
            int etAlMin = GetInheritedInt(el, "et-al-min", _ctxEtAlMin);
            int etAlUseFirst = GetInheritedInt(el, "et-al-use-first", _ctxEtAlUseFirst);
            string initializeWith = el != null && el.HasAttribute("initialize-with")
                ? el.GetAttribute("initialize-with") : null;
            string delimiter = el != null && !string.IsNullOrEmpty(Attr(el, "delimiter"))
                ? Attr(el, "delimiter") : ", ";
            // form="short" → family name only: APA in-text is "(Doe, 2022)",
            // NOT "(Doe, J., 2022)".
            bool shortForm = el != null && Attr(el, "form") == "short";
            string and = el != null ? Attr(el, "and") : "";
            string dpl = el != null ? Attr(el, "delimiter-precedes-last") : "";

            var rendered = new List<string>();
            int limit = c.Authors.Count >= etAlMin ? etAlUseFirst : c.Authors.Count;
            for (int i = 0; i < Math.Min(limit, c.Authors.Count); i++)
            {
                string person = c.Authors[i];
                if (shortForm) person = Surname(person);
                else person = FormatPerson(person, initializeWith);
                if (!string.IsNullOrWhiteSpace(person)) rendered.Add(person);
            }

            string body;
            if (c.Authors.Count >= etAlMin && rendered.Count < c.Authors.Count)
            {
                // "Kaur et al." / "Kaur, T., et al."
                body = rendered.Count > 0
                    ? string.Join(delimiter, rendered) + " et al."
                    : "et al.";
            }
            else
            {
                string andWord = and == "symbol" ? "&"
                               : and == "text" ? "and"
                               : "";
                bool precedesLast = dpl == "always" || (dpl != "never" && rendered.Count > 2);

                if (rendered.Count == 0) body = "";
                else if (rendered.Count == 1) body = rendered[0];
                else if (string.IsNullOrEmpty(andWord))
                    body = string.Join(delimiter, rendered);          // last delimiter covers all
                else if (rendered.Count == 2)
                    body = rendered[0] + (precedesLast ? delimiter : " ")
                         + andWord + " " + rendered[1];
                else
                    body = string.Join(delimiter, rendered.Take(rendered.Count - 1))
                         + (precedesLast ? delimiter : " ") + andWord + " "
                         + rendered[rendered.Count - 1];
            }
            return el != null ? Wrap(el, Mark(body, fmt)) : Mark(body, fmt);
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

        private string RenderDate(XmlElement el, Citation c, Fmt fmt)
        {
            // Only the primary issue date maps to our model — CSL styles
            // also render original-date/accessed, which we don't carry
            // (APA's date-intext renders "original/issued" and would emit
            // the year twice without this guard).
            string var = Attr(el, "variable");
            if (var != "" && var != "issued") return "";
            string year = c.Year?.ToString(CultureInfo.InvariantCulture) ?? "";
            if (string.IsNullOrEmpty(year)) return "";
            return Wrap(el, Mark(year, fmt));       // year-only — full date forms not modelled
        }

        // ── helpers ───────────────────────────────────────────────────────

        /// <summary>Wraps leaf text in rich-text markers when the effective
        /// formatting (from the enclosing CSL elements) is non-plain.</summary>
        private static string Mark(string s, Fmt fmt)
        {
            if (string.IsNullOrEmpty(s)) return s;
            if (fmt.Italic) s = ItalicOn + s + ItalicOff;
            if (fmt.Bold) s = BoldOn + s + BoldOff;
            return s;
        }

        private static string StripMarkers(string s)
            => s?.Replace(ItalicOn.ToString(), "")
                 .Replace(ItalicOff.ToString(), "")
                 .Replace(BoldOn.ToString(), "")
                 .Replace(BoldOff.ToString(), "");

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

        /// <summary>Integer attribute looked up on the element and then its
        /// ancestors — CSL inherits et-al-min / et-al-use-first from the
        /// enclosing citation/bibliography element.</summary>
        private static int GetInheritedInt(XmlElement el, string name, int fallback)
        {
            for (XmlElement e = el; e != null; e = e.ParentNode as XmlElement)
            {
                string v = e.GetAttribute(name);
                if (int.TryParse(v, out int n)) return n;
            }
            return fallback;
        }

        private static string Surname(string author)
        {
            int comma = (author ?? "").IndexOf(',');
            return comma >= 0 ? author.Substring(0, comma).Trim() : (author ?? "").Trim();
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
    }
}
