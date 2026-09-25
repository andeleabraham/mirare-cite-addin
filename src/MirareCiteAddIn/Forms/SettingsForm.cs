// ============================================================================
//  SettingsForm.cs  —  minimal WinForms dialog for the Mirare Cite settings
//  (citation style, remote endpoint URL, default library / project paths).
//
//  On OK, exposes the values back to RibbonCallbacks, which persists them
//  to %APPDATA%\MirareCite\settings.json.
// ============================================================================

using System;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using MirareCiteAddIn.Models;
using MirareCiteAddIn.Services;

namespace MirareCiteAddIn.Forms
{
    public class SettingsForm : Form
    {
        private readonly Logger _log;
        private ComboBox _styleCombo;
        private TextBox _remoteBox;
        private TextBox _libraryBox;
        private TextBox _projectBox;
        private TextBox _cslBox;
        private Button _btnBrowseLibrary;
        private Button _btnBrowseProject;
        private Button _btnBrowseCsl;
        private Button _btnOk;
        private Button _btnCancel;

        public CitationStyle Style { get; private set; }
        public string RemoteEndpoint { get; private set; }
        public string LibraryPath { get; private set; }
        public string ProjectPath { get; private set; }
        public string CslStylePath { get; private set; }

        public SettingsForm(CitationStyle style, string remote, string library, string project,
                            string cslStylePath, Logger log)
        {
            _log = log;
            Style = style;
            RemoteEndpoint = remote;
            LibraryPath = library;
            ProjectPath = project;
            CslStylePath = cslStylePath ?? "";
            Build();
        }

        private void Build()
        {
            Text = "Mirare Cite — Settings";
            Width = 540;
            Height = 380;
            StartPosition = FormStartPosition.CenterScreen;
            Font = new Font("Segoe UI", 9F);
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;

            var t = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(16),
                ColumnCount = 3,
                RowCount = 6,
                AutoSize = false,
            };
            t.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 110));
            t.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            t.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 80));

            // Style row
            t.Controls.Add(new Label { Text = "Citation style:", AutoSize = true, Anchor = AnchorStyles.Left, Margin = new Padding(0, 8, 0, 8) }, 0, 0);
            _styleCombo = new ComboBox
            {
                DropDownStyle = ComboBoxStyle.DropDownList,
                Dock = DockStyle.Top,
                Anchor = AnchorStyles.Left | AnchorStyles.Right,
            };
            _styleCombo.Items.AddRange(new object[] {
                CitationStyle.Apa, CitationStyle.Mla, CitationStyle.Chicago,
                CitationStyle.Numeric, CitationStyle.Csl
            });
            _styleCombo.SelectedItem = Style;
            t.Controls.Add(_styleCombo, 1, 0);
            t.Controls.Add(new Panel(), 2, 0);

            // Remote endpoint
            t.Controls.Add(new Label { Text = "Remote endpoint:", AutoSize = true, Anchor = AnchorStyles.Left }, 0, 1);
            _remoteBox = new TextBox { Text = RemoteEndpoint, Dock = DockStyle.Top };
            t.Controls.Add(_remoteBox, 1, 1);
            t.Controls.Add(new Panel(), 2, 1);

            // Library path
            t.Controls.Add(new Label { Text = "Library file:", AutoSize = true, Anchor = AnchorStyles.Left }, 0, 2);
            _libraryBox = new TextBox { Text = LibraryPath, Dock = DockStyle.Top, ReadOnly = true };
            t.Controls.Add(_libraryBox, 1, 2);
            _btnBrowseLibrary = new Button { Text = "Browse…", Dock = DockStyle.Top };
            _btnBrowseLibrary.Click += (s, e) => BrowseFile(_libraryBox,
                "RefManager library (*.refmanager.json)|*.refmanager.json|All JSON|*.json");
            t.Controls.Add(_btnBrowseLibrary, 2, 2);

            // Project path
            t.Controls.Add(new Label { Text = "Project file:", AutoSize = true, Anchor = AnchorStyles.Left }, 0, 3);
            _projectBox = new TextBox { Text = ProjectPath, Dock = DockStyle.Top, ReadOnly = true };
            t.Controls.Add(_projectBox, 1, 3);
            _btnBrowseProject = new Button { Text = "Browse…", Dock = DockStyle.Top };
            _btnBrowseProject.Click += (s, e) => BrowseFile(_projectBox,
                "Mirare Cite project (*.mrrcite)|*.mrrcite");
            t.Controls.Add(_btnBrowseProject, 2, 3);

            // CSL style file (used when Citation style = Csl). The browse
            // dialog opens in the Mirare app's styles folder when it can be
            // discovered from %APPDATA%\MirareCite\user_settings.json.
            t.Controls.Add(new Label
            {
                Text = "CSL style file:",
                AutoSize = true,
                Anchor = AnchorStyles.Left,
                Enabled = Style == CitationStyle.Csl
            }, 0, 4);
            _cslBox = new TextBox { Text = CslStylePath, Dock = DockStyle.Top, ReadOnly = true };
            t.Controls.Add(_cslBox, 1, 4);
            _btnBrowseCsl = new Button { Text = "Browse…", Dock = DockStyle.Top };
            _btnBrowseCsl.Click += (s, e) =>
            {
                BrowseFile(_cslBox, "CSL style (*.csl)|*.csl");
                if (!string.IsNullOrEmpty(_cslBox.Text))
                    _styleCombo.SelectedItem = CitationStyle.Csl;   // picking a file implies CSL
            };
            t.Controls.Add(_btnBrowseCsl, 2, 4);
            _styleCombo.SelectedIndexChanged += (s, e) =>
            {
                bool csl = (CitationStyle)_styleCombo.SelectedItem == CitationStyle.Csl;
                _cslBox.Enabled = csl;
                _btnBrowseCsl.Enabled = csl;
            };
            _cslBox.Enabled = Style == CitationStyle.Csl;
            _btnBrowseCsl.Enabled = Style == CitationStyle.Csl;

            // Buttons row
            var btnBar = new FlowLayoutPanel
            {
                Dock = DockStyle.Top,
                FlowDirection = FlowDirection.RightToLeft,
                Height = 36,
                BackColor = Color.Transparent
            };
            _btnOk = new Button { Text = "OK", Width = 80, Height = 30, DialogResult = DialogResult.OK };
            _btnCancel = new Button { Text = "Cancel", Width = 80, Height = 30, DialogResult = DialogResult.Cancel };
            _btnOk.Click += (s, e) => Commit();
            btnBar.Controls.Add(_btnCancel);
            btnBar.Controls.Add(_btnOk);

            Controls.Add(t);
            Controls.Add(btnBar);
            AcceptButton = _btnOk;
            CancelButton = _btnCancel;
        }

        private void BrowseFile(TextBox box, string filter)
        {
            using (var ofd = new OpenFileDialog { Filter = filter, CheckFileExists = true })
            {
                string dir = DiscoverStylesFolder();
                if (dir != null && Directory.Exists(dir)) ofd.InitialDirectory = dir;
                if (ofd.ShowDialog() == DialogResult.OK)
                    box.Text = ofd.FileName;
            }
        }

        /// <summary>The Mirare app's styles folder, discovered from the app's
        /// own settings (%APPDATA%\MirareCite\user_settings.json →
        /// working_directory + "\styles"). Null if not discoverable.</summary>
        private static string DiscoverStylesFolder()
        {
            try
            {
                string appSettings = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                    "MirareCite", "user_settings.json");
                if (!File.Exists(appSettings)) return null;
                var json = Newtonsoft.Json.Linq.JObject.Parse(File.ReadAllText(appSettings));
                string wd = (string)json["working_directory"];
                if (string.IsNullOrEmpty(wd)) return null;
                return Path.Combine(wd, "styles");
            }
            catch
            {
                return null;
            }
        }

        private void Commit()
        {
            Style = (CitationStyle)_styleCombo.SelectedItem;
            RemoteEndpoint = _remoteBox.Text.Trim();
            LibraryPath = _libraryBox.Text.Trim();
            ProjectPath = _projectBox.Text.Trim();
            CslStylePath = _cslBox.Text.Trim();
            if (Style == CitationStyle.Csl && !File.Exists(CslStylePath))
            {
                MessageBox.Show("Citation style is set to Csl — pick a .csl style file first.",
                    "Mirare Cite", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;     // keep the dialog open
            }
            DialogResult = DialogResult.OK;
        }
    }
}
