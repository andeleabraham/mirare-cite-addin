using System.Reflection;
using System.Runtime.InteropServices;

// ============================================================================
//  AssemblyInfo for MirareCiteAddIn.dll  — Mirare Cite Word COM add-in
//
//  Two attributes below are NON-NEGOTIABLE for a working Word COM add-in:
//
//    [assembly: ComVisible(true)]
//        Makes every public type COM-visible. Without this, regasm will
//        register the assembly but Word cannot instantiate the Connect class.
//
//    [assembly: Guid("...")]
//        A stable LibID. regasm writes this into HKCR\TypeLib\... . Do NOT
//        regenerate it on every build — keep this literal.
//
//  The ProgID and Class-GUID are declared on the Connect class itself
//  (see Connect.cs) because that's where Office looks them up.
// ============================================================================

[assembly: AssemblyTitle("Mirare Cite Word Add-in")]
[assembly: AssemblyDescription("Word COM add-in for Mirare Cite — insert citations from local .mrrcite projects, .refmanager.json libraries, or the remote Mirare API.")]
[assembly: AssemblyConfiguration("")]
[assembly: AssemblyCompany("Mirare Cite")]
[assembly: AssemblyProduct("Mirare Cite Word Add-in")]
[assembly: AssemblyCopyright("Copyright (c) Mirare Cite contributors")]
[assembly: AssemblyTrademark("")]
[assembly: AssemblyCulture("")]

[assembly: ComVisible(true)]

[assembly: Guid("9F4A2C7E-3D1B-4E8A-A9F0-1C2B3D4E5F60")]

[assembly: AssemblyVersion("1.0.0.0")]
[assembly: AssemblyFileVersion("1.0.0.0")]

// Let the type-lib version track the assembly version so regasm /tlb stays stable.
[assembly: TypeLibVersion(1, 0)]
