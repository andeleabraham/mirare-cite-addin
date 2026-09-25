// ============================================================================
//  CitationStyle.cs — common enums used across the picker, formatter, and
//  tracker. Declared ONCE here so we never get a duplicate-definition
//  compile error.
// ============================================================================

namespace MirareCiteAddIn.Models
{
    /// <summary>Pre-filters the picker dialog.</summary>
    public enum CitationSourceScope
    {
        All,        // show library + project + remote merged
        Library,    // primary .refmanager.json library
        Project,    // .mrrcite (zip → JSON)
        Remote      // remote HTTPS query
    }

    /// <summary>Supported citation styles for in-text + bibliography rendering.</summary>
    public enum CitationStyle
    {
        Apa,
        Mla,
        Chicago,
        Numeric,
        Csl        // rendered from a .csl file (same files the Mirare app uses)
    }
}
