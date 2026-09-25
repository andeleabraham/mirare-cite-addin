// ============================================================================
//  Connect.cs — the COM entry point for the Mirare Cite Word add-in.
//
//  Implements:
//    - Extensibility.IDTExtensibility2      → so Word can load/unload us.
//    - Microsoft.Office.Core.IRibbonExtensibility  → so Word asks us for the
//      customUI XML that defines the "Mirare Cite" ribbon tab.
//
//  If your existing WordAddIn_Mirare.dll does not implement BOTH of these
//  interfaces, that is *exactly* why no ribbon tab appears — see
//  DIAGNOSIS.md §5.
//
//  Ribbon callbacks live in RibbonCallbacks.cs so this file stays focused on
//  the COM plumbing.
// ============================================================================

using System;
using System.Reflection;
using System.Runtime.InteropServices;
using Extensibility;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;
// Alias for the few qualified spellings below (Word.Application etc.);
// the plain using above stays so Document/Range/Field resolve unqualified.
using Word = Microsoft.Office.Interop.Word;
using System.Windows.Forms;
using MirareCiteAddIn.Services;

namespace MirareCiteAddIn
{
    // ─────────────────────────────────────────────────────────────────────
    //  ComVisible + Guid + ProgID  — these THREE attributes are what Word's
    //  add-in loader looks up.  If the ProgID does not match the registry key
    //  HKCU\Software\Microsoft\Office\Word\Addins\<ProgID>, Word will never
    //  find you.
    // ─────────────────────────────────────────────────────────────────────
    [ComVisible(true)]
    [Guid("B3F5D2A8-7E1A-4C2B-9D3F-8A1B2C3D4E5F")]   // ← stable, do not regenerate
    [ProgId("MirareCite.AddIn")]
    // AutoDispatch (NOT None): Word's ribbon engine invokes onAction/onLoad
    // callbacks late-bound through IDispatch. ClassInterfaceType.None
    // exposes no IDispatch, so every button click silently fails — the
    // ribbon renders but nothing happens.
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    public class Connect : IDTExtensibility2, IRibbonExtensibility
    {
        // The ProgID is repeated here as a const so the installer .reg file
        // and the Connect class can never drift out of sync.
        public const string ProgId = "MirareCite.AddIn";

        private Word.Application _wordApp;
        private Document _activeDocument;
        private RibbonCallbacks _callbacks;
        private Logger _log;

        // ─────────────────────────────────────────────────────────────────
        //  IDTExtensibility2: Word calls this once when the add-in loads.
        //  This is where we capture the Word.Application pointer — everything
        //  else (ribbon, dialogs, citation insertion) hangs off this pointer.
        // ─────────────────────────────────────────────────────────────────
        public void OnConnection(object application, ext_ConnectMode connectMode,
            object addInInst, ref Array custom)
        {
            try
            {
                _wordApp = (Word.Application)application;
                _log = new Logger();
                _log.Info("OnConnection: mode=" + connectMode);

                _callbacks = new RibbonCallbacks(_wordApp, _log);

                // Wire document-open / document-change events so the
                // CitedTracker can refresh the cited-indicator list
                // whenever the user switches documents.
                _wordApp.DocumentChange += OnDocumentChange;
                ((Word.ApplicationEvents4_Event)_wordApp).DocumentOpen += OnDocumentOpen;
            }
            catch (Exception ex)
            {
                // NEVER let an exception escape OnConnection — Word will
                // hard-disable the add-in (LoadBehavior becomes 16) and you
                // won't see a single error message.  Log and swallow.
                _log?.Error("OnConnection failed", ex);
            }
        }

        public void OnDisconnection(ext_DisconnectMode removeMode, ref Array custom)
        {
            try
            {
                if (_wordApp != null)
                {
                    _wordApp.DocumentChange -= OnDocumentChange;
                    // Word.ApplicationEvents4_Event.DocumentOpen -= ...
                }
                _wordApp = null;
                _activeDocument = null;
                _callbacks = null;
            }
            catch (Exception ex)
            {
                _log?.Error("OnDisconnection failed", ex);
            }
        }

        public void OnAddInsUpdate(ref Array custom) { }
        public void OnStartupComplete(ref Array custom)
        {
            _log?.Info("OnStartupComplete — add-in fully loaded");
        }
        public void OnBeginShutdown(ref Array custom) { }

        // ─────────────────────────────────────────────────────────────────
        //  IRibbonExtensibility: Word calls this to ask for the customUI XML
        //  that defines our ribbon tab.  The ribbonID parameter is ALWAYS
        //  the literal string "Microsoft.Word.Word" (case-sensitive) —
        //  returning the wrong case or returning null is the #2 cause of
        //  "no ribbon tab".
        // ─────────────────────────────────────────────────────────────────
        public string GetCustomUI(string ribbonID)
        {
            try
            {
                _log.Info("GetCustomUI ribbonID=" + ribbonID);
                // Word 2016/365 passes "Microsoft.Word.Document" for the main
                // document window; some hosts/versions use "Microsoft.Word.Word".
                // Accept every Microsoft.Word* ribbon ID — returning null for
                // the document window is why the tab silently never appeared.
                if (!string.IsNullOrEmpty(ribbonID) &&
                    ribbonID.StartsWith("Microsoft.Word", StringComparison.OrdinalIgnoreCase))
                {
                    return ResourceLoader.LoadRibbonXml();
                }
                // Any other host's ribbon — we are a Word add-in only.
                return null;
            }
            catch (Exception ex)
            {
                _log.Error("GetCustomUI failed — ribbon will NOT appear", ex);
                // Returning null here means Word silently drops the ribbon.
                // Better to surface the error than to swallow it.
                return null;
            }
        }

        // ─────────────────────────────────────────────────────────────────
        //  Ribbon callbacks — exposed to COM because Office's ribbon engine
        //  late-binds them by name.  Every callback named in ribbon.xml
        //  MUST be a public method on this class with the right signature.
        //  We delegate to RibbonCallbacks to keep this file focused.
        //
        //  NOTE on attributes:  IRibbonControl is the Office PIA interface.
        //  Using a string for control.Id would also work but loses IntelliSense.
        // ─────────────────────────────────────────────────────────────────
        public void OnInsertCitation(IRibbonControl control)
            => _callbacks?.OnInsertCitation(control);

        public void OnInsertFromLibrary(IRibbonControl control)
            => _callbacks?.OnInsertFromLibrary(control);

        public void OnInsertFromProject(IRibbonControl control)
            => _callbacks?.OnInsertFromProject(control);

        public void OnInsertFromRemote(IRibbonControl control)
            => _callbacks?.OnInsertFromRemote(control);

        public void OnRefreshCitedIndicator(IRibbonControl control)
            => _callbacks?.OnRefreshCitedIndicator(control);

        public void OnEditBibliography(IRibbonControl control)
            => _callbacks?.OnEditBibliography(control);

        public void OnSettings(IRibbonControl control)
            => _callbacks?.OnSettings(control);

        public string GetCitedCountLabel(IRibbonControl control)
            => _callbacks?.GetCitedCountLabel(control) ?? "0 cited";

        // ─────────────────────────────────────────────────────────────────
        //  Ribbon lifecycle + image callbacks — these MUST exist on the
        //  Connect class because ribbon.xml declares them.  Returning null
        //  for image callbacks is safe; Office just renders the button
        //  without a custom icon.
        // ─────────────────────────────────────────────────────────────────
        /// <summary>
        /// Called by Office once when the ribbon loads — we stash the
        /// IRibbonUI pointer so RibbonCallbacks can later call
        /// InvalidateControl("mrcCitedCount") to refresh the dynamic
        /// cited-count label after each insertion.
        /// </summary>
        public void OnRibbonLoad(IRibbonUI ribbon)
        {
            try
            {
                _ribbonUi = ribbon;
                // RibbonCallbacks re-activates the tab after modal dialogs
                // close — without this Word collapses the tab the moment a
                // picker steals focus.
                if (_callbacks != null) _callbacks.RibbonUi = ribbon;
                _log?.Info("OnRibbonLoad: IRibbonUI captured");
            }
            catch (Exception ex) { _log?.Error("OnRibbonLoad failed", ex); }
        }
        private IRibbonUI _ribbonUi;

        // ─────────────────────────────────────────────────────────────────
        //  Document events — keep the CitedTracker up to date.
        // ─────────────────────────────────────────────────────────────────
        private void OnDocumentChange()
        {
            try
            {
                // Word fires DocumentChange with no document open (e.g. right
                // after closing the last one) — get_ActiveDocument would throw.
                if (_wordApp.Documents.Count == 0) return;
                _activeDocument = _wordApp.ActiveDocument;
                _callbacks?.OnActiveDocumentChanged(_activeDocument);
            }
            catch (Exception ex)
            {
                _log?.Error("OnDocumentChange failed", ex);
            }
        }

        private void OnDocumentOpen(Document doc)
        {
            try
            {
                _activeDocument = doc;
                _callbacks?.OnActiveDocumentChanged(doc);
            }
            catch (Exception ex)
            {
                _log?.Error("OnDocumentOpen failed", ex);
            }
        }
    }

    // ─────────────────────────────────────────────────────────────────────
    //  ResourceLoader — pulls ribbon.xml out of the embedded resource.
    // ─────────────────────────────────────────────────────────────────────
    internal static class ResourceLoader
    {
        public static string LoadRibbonXml()
        {
            var asm = Assembly.GetExecutingAssembly();
            // The default resource manifest name is:
            //   <DefaultNamespace>.<Folder>.<FileName>
            // → MirareCiteAddIn.Resources.ribbon.xml
            const string resName = "MirareCiteAddIn.Resources.ribbon.xml";
            using (var s = asm.GetManifestResourceStream(resName))
            {
                if (s == null)
                    throw new InvalidOperationException(
                        "Embedded resource not found: " + resName +
                        " — did the .csproj include Resources\\ribbon.xml as <EmbeddedResource>?");
                using (var r = new System.IO.StreamReader(s))
                    return r.ReadToEnd();
            }
        }
    }
}
