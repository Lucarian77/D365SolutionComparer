using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using D365SolutionComparer.Models.Membership;
using D365SolutionComparer.Models.Presentation;

namespace D365SolutionComparer
{
    /// <summary>Local navigation over a completed comparison; owns no service or capture callback.</summary>
    internal sealed class SimplifiedSolutionExplorerPanel : UserControl
    {
        private readonly ToolStripTextBox search;
        private readonly ToolStripComboBox status, scope;
        private readonly ToolStripButton differences;
        private readonly Label summary, preparation;
        private readonly TableLayoutPanel header;
        private readonly FlowLayoutPanel environmentPanel, metricPanel;
        private readonly Label[] metricLabels;
        private readonly List<Label> headerLabels = new List<Label>();
        private bool arrangingHeader;
        private readonly MembershipComparisonPresentation owner;
        private readonly Dictionary<ObjectsWorkspaceScope, ScopeBookmark> bookmarks = new Dictionary<ObjectsWorkspaceScope, ScopeBookmark>();
        private long intentRevision;
        private sealed class ScopeBookmark
        {
            internal SolutionObjectsExplorerControl.NavigationBookmark Navigation;
            internal long Revision;
        }
        private bool updating;
        internal SolutionObjectsExplorerControl Explorer { get; }
        internal SolutionObjectsExplorerProjection Projection { get; private set; }

        internal SimplifiedSolutionExplorerPanel(MembershipComparisonPresentation presentation, Func<MembershipGridRow, string> describe)
        {
            owner = presentation ?? throw new ArgumentNullException(nameof(presentation));
            Dock = DockStyle.Fill; BackColor = Color.White; Font = PresentationFont(9F); ForeColor = Color.FromArgb(15, 23, 42);
            Name = "SimplifiedSolutionExplorer";
            AccessibleName = "Solution component comparison";
            header = new TableLayoutPanel { Dock = DockStyle.Top, AutoSize = true, ColumnCount = 1,
                Padding = new Padding(10, 6, 10, 6), BackColor = Color.FromArgb(241, 245, 249) };
            header.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            var title = new Label { AutoSize = true, Text = "Solution: " + presentation.SolutionUniqueName,
                Font = PresentationFont(11F, FontStyle.Bold), Margin = new Padding(0, 0, 0, 3) };
            header.Controls.Add(title); headerLabels.Add(title);
            environmentPanel = new FlowLayoutPanel { Dock = DockStyle.Top, AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink,
                WrapContents = true, Margin = Padding.Empty, Padding = Padding.Empty };
            var source = new Label { AutoSize = true, Text = "Source: " + presentation.Source.Diagnostics.EnvironmentName + " (" + presentation.Source.State + ")",
                Margin = new Padding(0, 0, 16, 3), ForeColor = Color.FromArgb(71, 85, 105) };
            var target = new Label { AutoSize = true, Text = "Target: " + presentation.Target.Diagnostics.EnvironmentName + " (" + presentation.Target.State + ")",
                Margin = new Padding(0, 0, 0, 3), ForeColor = Color.FromArgb(71, 85, 105) };
            headerLabels.Add(source); headerLabels.Add(target);
            environmentPanel.Controls.Add(source); environmentPanel.Controls.Add(target); header.Controls.Add(environmentPanel);
            metricPanel = new FlowLayoutPanel { Dock = DockStyle.Top, AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink,
                WrapContents = true, Margin = new Padding(0, 2, 0, 3), Padding = Padding.Empty };
            metricLabels = new Label[5];
            for (int i = 0; i < metricLabels.Length; i++)
            {
                metricLabels[i] = new Label { AutoSize = true, Font = PresentationFont(9F, FontStyle.Bold),
                    MinimumSize = new Size(MembershipResultsForm.ScaleColumnWidth(110, DeviceDpi), 28),
                    Padding = new Padding(8, 5, 8, 5), Margin = new Padding(0, 0, 10, 0), BackColor = Color.White };
                metricPanel.Controls.Add(metricLabels[i]);
            }
            header.Controls.Add(metricPanel);
            summary = new Label { Name = "SolutionExplorerSummary", AutoSize = true, Text = "Preparing local comparison navigation…",
                Font = PresentationFont(8.5F), ForeColor = Color.FromArgb(71, 85, 105), Margin = Padding.Empty };
            header.Controls.Add(summary); headerLabels.Add(summary);
            summary.TextChanged += (sender, args) => RefreshProfessionalHeaderCounts();
            SizeChanged += (sender, args) => LayoutProfessionalHeader();
            var bar = new ToolStrip { Name = "SolutionExplorerFilters", Dock = DockStyle.Top,
                GripStyle = ToolStripGripStyle.Hidden, LayoutStyle = ToolStripLayoutStyle.Flow,
                AccessibleName = "Search and filter solution components", AutoSize = true, Stretch = true,
                Font = Font, Padding = new Padding(8, 4, 8, 4), ForeColor = Color.FromArgb(51, 65, 85),
                Renderer = new ToolStripProfessionalRenderer(new ExplorerToolStripColors()) };
            ((FlowLayoutSettings)bar.LayoutSettings).WrapContents = true;
            scope = new ToolStripComboBox { Name = "SolutionExplorerScope", DropDownStyle = ComboBoxStyle.DropDownList,
                AutoSize = false, Width = 110, AccessibleName = "Component scope" };
            scope.Items.AddRange(new object[] { "Compare", "Source", "Target" }); scope.SelectedIndex = 0;
            search = new ToolStripTextBox { Name = "SolutionExplorerSearch", AutoSize = false, Width = 180,
                AccessibleName = "Search component names and types", ToolTipText = "Search retained names and types locally. Ctrl+F focuses search." };
            status = new ToolStripComboBox { Name = "SolutionExplorerStatus", DropDownStyle = ComboBoxStyle.DropDownList,
                AutoSize = false, Width = 125, AccessibleName = "Difference status" };
            status.Items.AddRange(new object[] { "All", "Match", "Different", "Source Only", "Target Only", "Needs Review" }); status.SelectedIndex = 0;
            differences = new ToolStripButton("Show differences only") { Name = "SolutionExplorerDifferences", CheckOnClick = true };
            var reset = new ToolStripButton("Reset filters") { Name = "SolutionExplorerReset" };
            bar.Items.AddRange(new ToolStripItem[] { new ToolStripLabel("Scope:"), scope, new ToolStripLabel("Search:"), search,
                new ToolStripLabel("Status:"), status, differences, reset });
            foreach (ToolStripItem item in bar.Items) item.Margin = new Padding(0, 2, 8, 2);
            Explorer = new SolutionObjectsExplorerControl(describe,
                LocateOriginalRow, row =>
                {
                    if (row?.AuthoritativeRow.DefinitionDetail == null) return;
                    using (var form = new ComponentDefinitionDetailsForm(row.AuthoritativeRow.DefinitionDetail)) form.ShowDialog(FindForm());
                });
            Explorer.UseSimplifiedPresentation();
            Explorer.NavigationIntentChanged = () => intentRevision++;
            preparation = new Label { Name = "SolutionExplorerPreparation", Dock = DockStyle.Fill,
                Text = "Preparing local comparison navigation…", TextAlign = ContentAlignment.MiddleCenter, BackColor = Color.White };
            Controls.Add(Explorer); Controls.Add(preparation); Controls.Add(bar); Controls.Add(header);
            RefreshProfessionalHeaderCounts(); LayoutProfessionalHeader();
            search.TextChanged += (sender, args) => { if (!updating) Explorer.SetSearch(search.Text); };
            scope.SelectedIndexChanged += (sender, args) => ApplyFilters();
            status.SelectedIndexChanged += (sender, args) => ApplyFilters();
            differences.CheckedChanged += (sender, args) => ApplyFilters();
            reset.Click += (sender, args) =>
            {
                updating = true;
                try { search.Clear(); status.SelectedIndex = 0; differences.Checked = false; }
                finally { updating = false; }
                bookmarks.Clear(); intentRevision++;
                Explorer.SetSearch(string.Empty); ApplyFilters(); Explorer.ResetNavigation();
            };
        }

        private Font PresentationFont(float size, FontStyle style = FontStyle.Regular)
        {
            var font = new Font("Segoe UI", size, style, GraphicsUnit.Point);
            Disposed += (sender, args) => font.Dispose(); return font;
        }
        private void RefreshProfessionalHeaderCounts()
        {
            if (metricLabels == null) return;
            var counts = Explorer?.ProfessionalHeaderCounts;
            string[] titles = { "Total", "Match", "Different", "Needs Review", "Target Only" };
            for (int i = 0; i < titles.Length; i++) metricLabels[i].Text = titles[i] + "   " + (counts == null ? "\u2014" : counts[i].ToString());
            LayoutProfessionalHeader();
        }
        private void LayoutProfessionalHeader()
        {
            if (header == null || arrangingHeader) return;
            arrangingHeader = true;
            try
            {
                int width = ProfessionalHeaderContentWidth(ClientSize.Width, header.Padding.Horizontal);
                foreach (Label label in headerLabels) label.MaximumSize = new Size(width, 0);
                metricPanel.MaximumSize = new Size(width, 0); metricPanel.Width = width;
                environmentPanel.MaximumSize = new Size(width, 0); environmentPanel.Width = width;
            }
            finally { arrangingHeader = false; }
        }
        internal static int ProfessionalHeaderContentWidth(int clientWidth, int horizontalPadding) => Math.Max(180, clientWidth - horizontalPadding - 8);
        private sealed class ExplorerToolStripColors : ProfessionalColorTable
        {
            public ExplorerToolStripColors() { UseSystemColors = false; }
            public override Color ToolStripGradientBegin => Color.FromArgb(248, 250, 252);
            public override Color ToolStripGradientMiddle => ToolStripGradientBegin;
            public override Color ToolStripGradientEnd => ToolStripGradientBegin;
            public override Color ToolStripBorder => Color.FromArgb(226, 232, 240);
            public override Color ButtonSelectedGradientBegin => Color.FromArgb(239, 246, 255);
            public override Color ButtonSelectedGradientMiddle => ButtonSelectedGradientBegin;
            public override Color ButtonSelectedGradientEnd => ButtonSelectedGradientBegin;
            public override Color ButtonSelectedBorder => Color.FromArgb(29, 78, 216);
            public override Color ButtonPressedGradientBegin => Color.FromArgb(219, 234, 254);
            public override Color ButtonPressedGradientMiddle => ButtonPressedGradientBegin;
            public override Color ButtonPressedGradientEnd => ButtonPressedGradientBegin;
            public override Color ButtonCheckedGradientBegin => ButtonPressedGradientBegin;
            public override Color ButtonCheckedGradientMiddle => ButtonPressedGradientBegin;
            public override Color ButtonCheckedGradientEnd => ButtonPressedGradientBegin;
        }

        internal void UseObjectNavigation()
        {
            Explorer.UseObjectNavigation();
            Explorer.ObjectViewChanged += () => summary.Text = Explorer.ObjectSummary;
            Explorer.ObjectViewChanged += RefreshProfessionalHeaderCounts;
        }

        internal void Bind(SolutionObjectsExplorerProjection projection)
        {
            if (projection == null || !ReferenceEquals(projection.Owner, owner))
                throw new InvalidOperationException("The projection must belong to this completed comparison.");
            if (Projection != null && !ReferenceEquals(Projection, projection))
                throw new InvalidOperationException("A result window cannot replace its completed comparison owner.");
            Projection = projection ?? throw new ArgumentNullException(nameof(projection));
            Explorer.Bind(projection, Enumerable.Range(0, projection.Rows.Count), true);
            Explorer.ResetNavigation(); ApplyFilters(); preparation.Visible = false;
            summary.Text = Explorer.UsesObjectNavigation ? Explorer.ObjectSummary : projection.Rows.Count + " comparison rows · " + string.Join(" · ",
                projection.Rows.GroupBy(SolutionObjectsExplorerControl.SimplifiedStatus)
                    .OrderBy(g => g.Key, StringComparer.Ordinal).Select(g => g.Key + ": " + g.Count()));
        }

        internal void ShowFailure()
        {
            preparation.Text = "Component navigation could not be prepared. The completed comparison is retained. Close this window and reopen the comparison.";
            preparation.Visible = true; preparation.BringToFront();
        }

        private void ApplyFilters()
        {
            if (updating || Projection == null) return;
            var bookmark = Explorer.CaptureNavigation();
            if (bookmark != null) bookmarks[bookmark.Scope] = new ScopeBookmark { Navigation = bookmark, Revision = intentRevision };
            var currentScope = scope.SelectedIndex == 1 ? ObjectsWorkspaceScope.Source :
                scope.SelectedIndex == 2 ? ObjectsWorkspaceScope.Target : ObjectsWorkspaceScope.Compare;
            Explorer.ConfigureWorkspace(true, currentScope);
            Explorer.SetObjectFilters((string)status.SelectedItem, differences.Checked);
            var rows = Explorer.UsesObjectNavigation ? Enumerable.Range(0, Projection.Rows.Count).ToArray() : Projection.Rows.Where(row => (currentScope == ObjectsWorkspaceScope.Compare ||
                (currentScope == ObjectsWorkspaceScope.Source ? row.Original.AuthoritativeRow.SourceIdentity : row.Original.AuthoritativeRow.TargetIdentity) != null) &&
                (status.SelectedIndex == 0 || SolutionObjectsExplorerControl.SimplifiedStatus(row) == (string)status.SelectedItem) &&
                (!differences.Checked || SolutionObjectsExplorerControl.SimplifiedStatus(row) != "Match")).Select(row => row.OriginalRowIndex).ToArray();
            Explorer.ApplyFilter(rows, rows.Length == Projection.Rows.Count);
            if (bookmark != null)
            {
                // An older scope bookmark cannot override a newer explicit row, side, category or clear.
                // Scope changes migrate only this view context; row/occurrence owners and sides stay exact.
                var restore = bookmarks.TryGetValue(currentScope, out var saved) && saved.Revision == intentRevision ?
                    saved.Navigation : InScope(bookmark, currentScope);
                Explorer.RestoreNavigation(restore, restore.Row, restore.Side, restore.Occurrence);
            }
        }

        private static SolutionObjectsExplorerControl.NavigationBookmark InScope(
            SolutionObjectsExplorerControl.NavigationBookmark bookmark, ObjectsWorkspaceScope scope) =>
            new SolutionObjectsExplorerControl.NavigationBookmark { Owner = bookmark.Owner, Primary = bookmark.Primary,
                Scope = scope, Support = false, ObjectKey = bookmark.ObjectKey, SelectedPath = bookmark.SelectedPath, TopPath = bookmark.TopPath,
                Focus = bookmark.Focus, ExpandedPaths = bookmark.ExpandedPaths, Row = bookmark.Row, Side = bookmark.Side,
                Occurrence = bookmark.Occurrence, TopRow = bookmark.TopRow, HorizontalOffset = bookmark.HorizontalOffset };

        private void LocateOriginalRow(MembershipGridRow row)
        {
            // Original means the same authoritative row, never a new match or a service lookup.
            var bookmark = Explorer.CaptureNavigation();
            if (bookmark == null) return;
            Explorer.RestoreNavigation(bookmark, row.OriginalRowIndex, bookmark.Side, bookmark.Occurrence);
        }

        internal void FocusSearch() { search.Focus(); search.TextBox.SelectAll(); }
        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            if (keyData == (Keys.Control | Keys.F)) { FocusSearch(); return true; }
            return base.ProcessCmdKey(ref msg, keyData);
        }
    }
}
