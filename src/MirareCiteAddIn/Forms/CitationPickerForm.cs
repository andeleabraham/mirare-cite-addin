// ============================================================================
//  CitationPickerForm.cs — WinForms dialog that lets the user pick a
//  citation from library, project, remote, or all three.
//
//  The cited indicator (✓) appears next to any item already cited in the
//  active document.  The tracker's CitedIds set is passed in by
//  RibbonCallbacks.OpenPicker.
//
//  This is intentionally a hand-coded WinForms form (no Designer.cs
//  magic) — easier to read, easier to wire, and trivial to extend.
//  A stub Designer.cs exists next to this file so the .csproj compiles.
// ============================================================================

using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using MirareCiteAddIn.Models;
using MirareCiteAddIn.Services;

namespace MirareCiteAddIn.Forms
{
    public partial class CitationPickerForm : Form
    {
        private readonly CitationSourceScope _scope;
        private readonly CitationStyle _style;
        private readonly string _remoteEndpoint;
        // Not readonly — the Library…/Project… buttons replace the path when
        // the user picks a different file, then the list reloads from it.
        private string _lastLibraryPath;
        private string _lastProjectPath;
        private readonly HashSet<string> _citedIds;
        private readonly Logger _log;

        private List<Citation> _all = new List<Citation>();
        private List<Citation> _filtered = new List<Citation>();

        // UI controls — declared here, hand-built in InitializeComponent.
        private ComboBox _scopeCombo;
        private TextBox _searchBox;
        private ListView _list;
        private Button _btnLibraryPath;
        private Button _btnProjectPath;
        private Button _btnRemoteSearch;
        private Button _btnInsert;
        private Button _btnCancel;
        private Label _lblSource;

        public Citation SelectedCitation { get; private set; }

        /// <summary>All picked citations (Ctrl/Shift multi-select) — inserted
        /// into ONE citation field, Zotero-style. Single pick = list of 1.</summary>
        public List<Citation> SelectedCitations { get; private set; }

        public CitationPickerForm(CitationSourceScope scope, CitationStyle style,
            string remoteEndpoint, string lastLibraryPath, string lastProjectPath,
            HashSet<string> citedIds, Logger log)
        {
            _scope = scope;
            _style = style;
            _remoteEndpoint = remoteEndpoint;
            _lastLibraryPath = lastLibraryPath;
            _lastProjectPath = lastProjectPath;
            _citedIds = citedIds ?? new HashSet<string>();
            _log = log;
            InitializeComponent();
            LoadInitialData();
        }

        private void InitializeComponent()
        {
            // ── Form shell ─────────────────────────────────────────────────
            Text = "Mirare Cite — Insert Citation";
            Width = 800;
            Height = 600;
            MinimumSize = new Size(640, 480);
            StartPosition = FormStartPosition.CenterScreen;
            Font = new Font("Segoe UI", 9F);

            // ── Top toolbar ───────────────────────────────────────────────
            var toolbar = new FlowLayoutPanel
            {
                Dock = DockStyle.Top,
                Height = 40,
                Padding = new Padding(8, 6, 8, 0),
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = false,
            };

            var lblScope = new Label { Text = "Source:", AutoSize = true, Margin = new Padding(0, 6, 6, 0) };
            _scopeCombo = new ComboBox
            {
                DropDownStyle = ComboBoxStyle.DropDownList,
                Width = 140,
            };
            _scopeCombo.Items.AddRange(new object[] {
                CitationSourceScope.All,
                CitationSourceScope.Library,
                CitationSourceScope.Project,
                CitationSourceScope.Remote
            });
            _scopeCombo.SelectedItem = _scope;
            _scopeCombo.SelectedIndexChanged += (s, e) => LoadInitialData();

            _searchBox = new TextBox
            {
                Width = 280,
                Margin = new Padding(8, 0, 4, 0)
                // NOTE: TextBox.PlaceholderText is .NET Core+ only — not
                // available in .NET Framework 4.8 WinForms.
            };
            _searchBox.TextChanged += (s, e) => ApplyFilter();

            _btnRemoteSearch = new Button
            {
                Text = "Search remote",
                Width = 110,
                Margin = new Padding(4, 0, 0, 0)
            };
            _btnRemoteSearch.Click += OnRemoteSearch;

            toolbar.Controls.AddRange(new Control[] { lblScope, _scopeCombo, _searchBox, _btnRemoteSearch });
            Controls.Add(toolbar);

            // ── Source-path bar ──────────────────────────────────────────
            var srcBar = new FlowLayoutPanel
            {
                Dock = DockStyle.Top,
                Height = 32,
                Padding = new Padding(8, 2, 8, 0),
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = false,
            };
            _btnLibraryPath = new Button { Text = "Library…", Width = 90 };
            _btnProjectPath = new Button { Text = "Project…", Width = 90, Margin = new Padding(6, 0, 0, 0) };
            _lblSource = new Label
            {
                AutoSize = true,
                Margin = new Padding(12, 6, 0, 0),
                ForeColor = Color.Gray
            };
            _btnLibraryPath.Click += (s, e) => PickLibraryPath();
            _btnProjectPath.Click += (s, e) => PickProjectPath();
            srcBar.Controls.AddRange(new Control[] { _btnLibraryPath, _btnProjectPath, _lblSource });
            Controls.Add(srcBar);

            // ── Main list ─────────────────────────────────────────────────
            _list = new ListView
            {
                Dock = DockStyle.Fill,
                View = View.Details,
                FullRowSelect = true,
                HideSelection = false,
                MultiSelect = true,     // pick several → one combined citation field
                Font = new Font("Segoe UI", 9F),
            };
            _list.Columns.Add("Cited", 60);
            _list.Columns.Add("Origin", 70);
            _list.Columns.Add("Year", 60);
            _list.Columns.Add("Title / Authors / miRNA", 540);
            _list.DoubleClick += (s, e) => OnInsert();
            _list.SelectedIndexChanged += (s, e) =>
                _btnInsert.Enabled = _list.SelectedItems.Count > 0;

            Controls.Add(_list);
            Controls.Add(toolbar);    // re-add at top z-order
            Controls.Add(srcBar);

            // ── Bottom button bar ────────────────────────────────────────
            var btnBar = new Panel { Dock = DockStyle.Bottom, Height = 44 };
            _btnInsert = new Button
            {
                Text = "Insert",
                Width = 100,
                Height = 30,
                Enabled = false,
                Anchor = AnchorStyles.Right
            };
            _btnInsert.Click += (s, e) => OnInsert();
            _btnCancel = new Button
            {
                Text = "Cancel",
                Width = 100,
                Height = 30,
                Anchor = AnchorStyles.Right,
                Left = _btnInsert.Right + 8
            };
            _btnCancel.Click += (s, e) => { DialogResult = DialogResult.Cancel; Close(); };
            btnBar.Controls.Add(_btnInsert);
            btnBar.Controls.Add(_btnCancel);
            // Position the buttons at right edge of the panel
            btnBar.Resize += (s, e) =>
            {
                _btnCancel.Left = btnBar.Width - _btnCancel.Width - 12;
                _btnInsert.Left = _btnCancel.Left - _btnInsert.Width - 8;
                _btnInsert.Top = (btnBar.Height - _btnInsert.Height) / 2;
                _btnCancel.Top = (btnBar.Height - _btnCancel.Height) / 2;
            };
            Controls.Add(btnBar);

            AcceptButton = _btnInsert;
            CancelButton = _btnCancel;
        }

        // ─────────────────────────────────────────────────────────────────
        //  Data loading — depending on the chosen scope, load from disk
        //  (library / project) or wait for an explicit remote search.
        // ─────────────────────────────────────────────────────────────────
        private void LoadInitialData()   // sync — remote search stays behind its own button
        {
            _all.Clear();
            _list.Items.Clear();

            try
            {
                var scope = (CitationSourceScope)_scopeCombo.SelectedItem;

                if (scope == CitationSourceScope.All || scope == CitationSourceScope.Library)
                {
                    if (File.Exists(_lastLibraryPath))
                    {
                        var lib = new RefManagerLoader(_log).Load(_lastLibraryPath);
                        _all.AddRange(lib);
                        _lblSource.Text = "Library: " + _lastLibraryPath;
                    }
                }
                if (scope == CitationSourceScope.All || scope == CitationSourceScope.Project)
                {
                    if (File.Exists(_lastProjectPath))
                    {
                        var proj = new ProjectLoader(_log).Load(_lastProjectPath);
                        _all.AddRange(proj);
                        if (!string.IsNullOrEmpty(_lblSource.Text))
                            _lblSource.Text += "  |  ";
                        _lblSource.Text += "Project: " + _lastProjectPath;
                    }
                }
                if (scope == CitationSourceScope.Remote && !string.IsNullOrEmpty(_remoteEndpoint))
                {
                    // Pre-seed with an empty remote search to surface the
                    // service — the user types a term and clicks "Search remote".
                    _lblSource.Text = "Remote: " + _remoteEndpoint;
                }
                ApplyFilter();
            }
            catch (Exception ex)
            {
                _log.Error("LoadInitialData failed", ex);
                MessageBox.Show("Failed to load citations: " + ex.Message,
                    "Mirare Cite", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private async void OnRemoteSearch(object sender, EventArgs e)
        {
            string term = _searchBox.Text.Trim();
            if (string.IsNullOrEmpty(_remoteEndpoint))
            {
                MessageBox.Show("Remote endpoint is not configured — set it in Settings.",
                    "Mirare Cite", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            try
            {
                UseWaitCursor = true;
                var svc = new RemoteCitationService(_remoteEndpoint, _log);
                var results = await svc.QueryAsync(term);
                // Merge remote results, dedup by Id.
                var seen = new HashSet<string>(_all.Select(c => c.Id));
                foreach (var r in results)
                    if (seen.Add(r.Id)) _all.Add(r);
                ApplyFilter();
            }
            catch (Exception ex)
            {
                _log.Error("Remote search failed", ex);
                MessageBox.Show("Remote search failed: " + ex.Message,
                    "Mirare Cite", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally { UseWaitCursor = false; }
        }

        // ─────────────────────────────────────────────────────────────────
        //  File-open pickers for library / project paths.  The chosen path
        //  is remembered via RibbonCallbacks' settings persistence.
        // ─────────────────────────────────────────────────────────────────
        private void PickLibraryPath()
        {
            using (var ofd = new OpenFileDialog
            {
                Title = "Select RefManager library (.refmanager.json)",
                Filter = "RefManager library (*.refmanager.json)|*.refmanager.json|All JSON|*.json",
                CheckFileExists = true,
            })
            {
                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    // Remember the path so LoadInitialData picks it up;
                    // the loader itself runs from there.
                    _lastLibraryPath = ofd.FileName;
                    if ((CitationSourceScope)_scopeCombo.SelectedItem == CitationSourceScope.All)
                        LoadInitialData();
                    else
                        _scopeCombo.SelectedItem = CitationSourceScope.All;
                }
            }
        }

        private void PickProjectPath()
        {
            using (var ofd = new OpenFileDialog
            {
                Title = "Select Mirare Cite project (.mrrcite)",
                Filter = "Mirare Cite project (*.mrrcite)|*.mrrcite",
                CheckFileExists = true,
            })
            {
                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    _lastProjectPath = ofd.FileName;
                    if ((CitationSourceScope)_scopeCombo.SelectedItem == CitationSourceScope.All)
                        LoadInitialData();
                    else
                        _scopeCombo.SelectedItem = CitationSourceScope.All;
                }
            }
        }

        // ─────────────────────────────────────────────────────────────────
        //  Filter + render — the cited indicator is the ✓ in the first
        //  column.  This is the visual the user explicitly asked for.
        // ─────────────────────────────────────────────────────────────────
        private void ApplyFilter()
        {
            string term = _searchBox.Text.Trim().ToLowerInvariant();
            _filtered = _all.Where(c =>
            {
                if (string.IsNullOrEmpty(term)) return true;
                string hay = (c.Title + " " + c.MiRNA + " " + c.TargetGene
                            + " " + (c.Authors == null ? "" : string.Join(" ", c.Authors)))
                            .ToLowerInvariant();
                return hay.Contains(term);
            }).ToList();
            RenderList();
        }

        private void RenderList()
        {
            _list.BeginUpdate();
            _list.Items.Clear();
            foreach (var c in _filtered)
            {
                bool isCited = _citedIds.Contains(c.Id);
                var item = new ListViewItem(isCited ? "✓" : "");
                item.UseItemStyleForSubItems = true;
                if (isCited)
                {
                    item.ForeColor = Color.Green;
                    item.Font = new Font(_list.Font, FontStyle.Bold);
                }
                item.SubItems.Add(OriginBadge(c.Origin));
                item.SubItems.Add(c.Year.HasValue ? c.Year.Value.ToString() : "—");
                item.SubItems.Add(c.DisplayLabel());
                item.Tag = c;
                _list.Items.Add(item);
            }
            _list.EndUpdate();
        }

        private static string OriginBadge(CitationOrigin o)
            => o == CitationOrigin.Remote ? "remote"
             : o == CitationOrigin.Project ? "project" : "library";

        // ─────────────────────────────────────────────────────────────────
        //  Insert button — collects ALL selected citations (Ctrl/Shift for
        //  multiple) and closes the dialog with OK.  RibbonCallbacks does
        //  the actual Word insertion — several picks share one field.
        // ─────────────────────────────────────────────────────────────────
        private void OnInsert()
        {
            if (_list.SelectedItems.Count == 0) return;
            var picked = new List<Citation>();
            foreach (ListViewItem item in _list.SelectedItems)
                if (item.Tag is Citation c) picked.Add(c);
            if (picked.Count == 0) return;

            SelectedCitations = picked;
            SelectedCitation = picked[0];
            DialogResult = DialogResult.OK;
            Close();
        }
    }
}
