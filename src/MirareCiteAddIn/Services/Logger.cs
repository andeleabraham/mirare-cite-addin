// ============================================================================
//  Logger.cs — a tiny file logger that writes to
//  %LOCALAPPDATA%\MirareCite\addin.log.  Use this everywhere — exceptions in
//  COM add-ins are SWALLOWED by Word unless you log them yourself.
// ============================================================================

using System;
using System.IO;

namespace MirareCiteAddIn.Services
{
    // Public: the class is a constructor parameter of public types
    // (RibbonCallbacks, the loaders, the forms) — internal would trip CS0051.
    public class Logger
    {
        private readonly string _path;

        public Logger()
        {
            _path = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "MirareCite", "addin.log");
            try { Directory.CreateDirectory(Path.GetDirectoryName(_path)); }
            catch { /* if we can't even create the folder, give up silently */ }
        }

        public void Info(string msg)  => Write("INFO ", msg, null);
        public void Warn(string msg)  => Write("WARN ", msg, null);
        public void Error(string msg, Exception ex) => Write("ERROR", msg + " :: " + ex?.Message + "\n" + ex?.StackTrace, null);

        private void Write(string lvl, string msg, object _)
        {
            try
            {
                string line = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff} [{lvl}] {msg}\n";
                File.AppendAllText(_path, line);
            }
            catch { /* swallow — logging must never throw */ }
        }
    }
}
