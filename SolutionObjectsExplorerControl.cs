using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using D365SolutionComparer.Models.Membership;
using D365SolutionComparer.Models.ComponentDetails;
using D365SolutionComparer.Models.Presentation;

namespace D365SolutionComparer
{
    /// <summary>Local read-only object navigation. No service, cache or resolver dependency.</summary>
    internal sealed partial class SolutionObjectsExplorerControl : UserControl
    {
        private bool simplifiedPresentation;
        private TextBox simplifiedIdentity;
        private const int PageSize = 200;
        private readonly TreeView navigation;
        private readonly DataGridView grid;
        private readonly TextBox details;
        private readonly Label counts;
        private readonly LinkLabel openRow, definition;
        private readonly LinkLabel expandDetails;
        private readonly Label detailsScrollCue;
        private readonly Panel detailsLinks;
        private readonly TableLayoutPanel detailsOptions;
        private readonly ComboBox detailsView;
        private readonly DataGridView comparedProperties;
        private readonly Panel comparedPropertiesPanel;
        private readonly Label comparedPropertiesMessage;
        private SolutionObjectsExplorerRow comparedPropertiesOwner;
        private IReadOnlyList<ComponentPropertyPresentation> comparedPropertyRows = new ComponentPropertyPresentation[0];
        private bool updatingDetailsCue;
        private int appliedRowMetricHeight = -1;
        internal int PrimaryRowHeight => Math.Max(26, grid.Font.Height + 10);
        private readonly Func<MembershipGridRow, string> describe;
        private readonly Action<MembershipGridRow> openOriginal, openDefinition;
        private readonly Dictionary<int, List<int>> parents = new Dictionary<int, List<int>>();
        private HashSet<int> allowedRows, includedOccurrences, matchingOccurrences;
        private IReadOnlyList<int> baseAllowedRows;
        private string searchText = string.Empty, supportEvidence = string.Empty;
        private bool filterAll;
        private long technicalContextRevision, cachedTechnicalRevision = -1;
        private TechnicalContext technicalContext;
        private enum TechnicalRoute { Row, RowOnly, Occurrence, Description }
        private sealed class TechnicalContext
        {
            internal SolutionObjectsExplorerProjection Projection;
            internal MembershipGridRow Row;
            internal DisplayEnvironmentSide? Side;
            internal int? Occurrence;
            internal TechnicalRoute Route;
            internal string Description;
            internal object Descriptor;
            internal long Revision;
        }
        private readonly CheckBox supportToggle;
        internal int? SelectedOccurrenceIndex { get; private set; }
        private bool detailsContentDirty;
        internal string SimpleDetailsText
        {
            get { if (detailsContentDirty) RenderDetailsContent(); return details.Text; }
        }
        private IReadOnlyList<int> visibleRows = new int[0];
        private bool preparing;
        private string accountingText;
        private bool primaryWorkspace, metadataRequested, detailsRequested;
        private ObjectsWorkspaceScope workspaceScope;
        private readonly ToolTip countToolTip = new ToolTip();
        private readonly SplitContainer evidenceSplit;
        internal Action TechnicalMembershipRequested { get; set; }
        internal bool IsPrimaryWorkspace => primaryWorkspace;
        internal ObjectsWorkspaceScope WorkspaceScope => workspaceScope;
        private TreeNode inspectionNode;
        private bool restoringNavigation, navigationContextChanged;
        private string lastNavigationFocus = "tree";
        private NavigationBookmark restoredNavigation;
        internal Action NavigationIntentChanged { get; set; }
        internal sealed class NavigationBookmark
        {
            internal SolutionObjectsExplorerProjection Owner;
            internal bool Primary, Support;
            internal ObjectsWorkspaceScope Scope;
            internal string SelectedPath, TopPath, Focus;
            internal string ObjectKey;
            internal string[] ExpandedPaths;
            internal int? Row, Occurrence, TopRow;
            internal DisplayEnvironmentSide? Side;
            internal int HorizontalOffset;
        }
        private MembershipGridRow selected;
        internal SolutionObjectsExplorerProjection Projection { get; private set; }
        internal TreeView Navigation => navigation;
        internal DataGridView Grid => grid;
        internal string EvidenceText => TechnicalEvidence();
        internal IReadOnlyList<int> VisibleRows => visibleRows;
        internal MembershipGridRow SelectedComparison => selected;
        internal DisplayEnvironmentSide? SelectedSide { get; private set; }
        internal Action<MembershipGridRow> SelectedComparisonChanged { get; set; }
        internal Action<Exception> RenderingFailed { get; set; }

        internal SolutionObjectsExplorerControl(Func<MembershipGridRow, string> describe,
            Action<MembershipGridRow> openOriginal, Action<MembershipGridRow> openDefinition)
        {
            this.describe = describe; this.openOriginal = openOriginal; this.openDefinition = openDefinition;
            Dock = DockStyle.Fill; Name = "SolutionObjectsExplorer"; AccessibleName = "Solution Objects Explorer";
            var outer = new SplitContainer { Dock = DockStyle.Fill, Size = new Size(1000, 550), SplitterDistance = 255,
                Panel1MinSize = 180, Panel2MinSize = 450, FixedPanel = FixedPanel.Panel1, SplitterWidth = 5, Name = "ExplorerNavigationSplit" };
            navigation = new TreeView { Dock = DockStyle.Fill, HideSelection = false, ShowNodeToolTips = true,
                BorderStyle = BorderStyle.FixedSingle, Name = "ExplorerObjectsNavigation", AccessibleName = "Objects, Source and Target category navigation" };
            var right = new SplitContainer { Dock = DockStyle.Fill, Size = new Size(735, 550), Orientation = Orientation.Horizontal,
                SplitterDistance = 325, Panel1MinSize = 170, Panel2MinSize = 100, Name = "ExplorerEvidenceSplit" };
            evidenceSplit = right;
            grid = new DataGridView { Dock = DockStyle.Fill, Name = "ExplorerObjectGrid", AccessibleName = "Object membership comparison",
                ReadOnly = true, VirtualMode = true, AutoGenerateColumns = false, AllowUserToAddRows = false, AllowUserToDeleteRows = false,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None, AllowUserToResizeColumns = true, ScrollBars = ScrollBars.Both,
                AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None, MultiSelect = false, SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                RowHeadersVisible = false, BackgroundColor = Color.White, RowTemplate = new ExplorerAccessibleRow { Tag = this } };
            grid.RowTemplate.Height = Math.Max(46, Font.Height * 3 + 4);
            foreach (var column in new[] { new[] { "DisplayName", "Display Name", "180" }, new[] { "Name", "Name", "210" },
                new[] { "Type", "Type", "155" }, new[] { "Source", "Source", "105" }, new[] { "Target", "Target", "105" },
                new[] { "ComparisonStatus", "Comparison Status", "250" }, new[] { "Managed", "Managed", "110" },
                new[] { "Customized", "Customized", "110" }, new[] { "LastModified", "Last Modified", "185" },
                new[] { "Owner", "Owner", "185" }, new[] { "Status", "Status", "170" } })
                grid.Columns.Add(new DataGridViewTextBoxColumn { Name = column[0], HeaderText = column[1],
                    Width = MembershipResultsForm.ScaleColumnWidth(int.Parse(column[2]), DeviceDpi),
                    MinimumWidth = 65, Resizable = DataGridViewTriState.True, AutoSizeMode = DataGridViewAutoSizeColumnMode.None,
                    SortMode = DataGridViewColumnSortMode.NotSortable, Visible = grid.Columns.Count < 6,
                    DefaultCellStyle = new DataGridViewCellStyle { WrapMode = DataGridViewTriState.True } });
            grid.FontChanged += (sender, args) => UpdateGridMetrics();
            grid.RowHeightInfoNeeded += (sender, args) => { args.Height = grid.RowTemplate.Height; args.MinimumHeight = grid.RowTemplate.Height; };
            grid.CellPainting += PaintPrimaryCell;
            grid.CellValueNeeded += (s, e) =>
            {
                if (makerNavigation)
                {
                    if (e.RowIndex >= 0 && e.RowIndex < visibleObjects.Count)
                        e.Value = ObjectValue(visibleObjects[e.RowIndex], grid.Columns[e.ColumnIndex].Name);
                    return;
                }
                if (e.RowIndex < visibleRows.Count) e.Value = DisplayValue(Projection.Rows[visibleRows[e.RowIndex]], grid.Columns[e.ColumnIndex].Name);
            };
            grid.CellClick += (s, e) => Guard(() => { if (e.RowIndex >= 0) { SelectGridRow(e.RowIndex); NavigationIntentChanged?.Invoke(); } });
            grid.KeyUp += (s, e) => Guard(() => { if (grid.CurrentRow != null) { SelectGridRow(grid.CurrentRow.Index); NavigationIntentChanged?.Invoke(); } });
            grid.CellDoubleClick += (s, e) => Guard(() => { if (e.RowIndex >= 0) { SelectGridRow(e.RowIndex); if (simplifiedPresentation) ShowDetails(true); else InspectOccurrences(visibleRows[e.RowIndex]); } });
            grid.CellFormatting += (s, e) =>
            {
                if (e.RowIndex < 0 || e.RowIndex >= visibleRows.Count || grid.Columns[e.ColumnIndex].Name != "ComparisonStatus" || grid.Rows[e.RowIndex].Selected) return;
                if (makerNavigation)
                {
                    string objectResult = visibleObjects[e.RowIndex].Status(Projection);
                    e.CellStyle.BackColor = objectResult == "Match" ? Color.FromArgb(243, 247, 244) :
                        objectResult == "Different" || objectResult == "Needs Review" ? Color.FromArgb(255, 247, 232) : Color.FromArgb(251, 239, 239);
                    return;
                }
                var presence = Projection.Rows[visibleRows[e.RowIndex]].Original.AuthoritativeRow.Presence;
                e.CellStyle.BackColor = presence == MembershipPresence.PresentInBoth ? Color.FromArgb(243, 247, 244) :
                    presence == MembershipPresence.Indeterminate ? Color.FromArgb(255, 247, 232) : Color.FromArgb(251, 239, 239);
                if (simplifiedPresentation)
                {
                    string status = SimplifiedStatus(Projection.Rows[visibleRows[e.RowIndex]]);
                    e.CellStyle.BackColor = status == "Match" ? Color.FromArgb(243, 247, 244) :
                        status == "Source Only" || status == "Target Only" ? Color.FromArgb(251, 239, 239) : Color.FromArgb(255, 247, 232);
                }
                e.CellStyle.ForeColor = Color.FromArgb(55, 60, 66);
            };
            grid.CellToolTipTextNeeded += (s, e) =>
            {
                if (makerNavigation && e.RowIndex >= 0 && e.RowIndex < visibleObjects.Count)
                { e.ToolTipText = ObjectValue(visibleObjects[e.RowIndex], grid.Columns[e.ColumnIndex].Name) + "; comparison is limited to retained properties."; return; }
                if (e.RowIndex >= 0 && e.RowIndex < visibleRows.Count)
                    e.ToolTipText = (primaryWorkspace ? FullCellToolTip(visibleRows[e.RowIndex], grid.Columns[e.ColumnIndex].Name) : "Original comparison row " + visibleRows[e.RowIndex] + ". ") +
                        MembershipGridRow.DefinitionMeaning + " Missing means selected solution membership, not missing environment metadata.";
            };
            counts = new Label { Dock = DockStyle.Top, Height = 48, Name = "ExplorerAccounting", AccessibleName = "Explorer count units", AutoEllipsis = true };
            details = new TextBox { Dock = DockStyle.Fill, Multiline = true, ReadOnly = true, ScrollBars = ScrollBars.Both,
                WordWrap = false, BackColor = Color.White, Name = "ExplorerEvidence", AccessibleName = "Selected Source and Target object evidence" };
            openRow = new LinkLabel { Dock = DockStyle.Left, Width = 235, Text = "Open original row in Component Details", Enabled = false };
            definition = new LinkLabel { Dock = DockStyle.Fill, Text = "Open existing definition details", Enabled = false };
            openRow.LinkClicked += (s, e) => { if (selected != null) openOriginal(selected); };
            definition.LinkClicked += (s, e) => { if (selected?.AuthoritativeRow.DefinitionDetail != null) openDefinition(selected); };
            expandDetails = new LinkLabel { Name = "ExplorerExpandDetails", Dock = DockStyle.Right, Width = 140, Text = "&Expand details",
                AccessibleName = "Expand Details viewport; restore grid with the same link", TabIndex = 4 };
            expandDetails.LinkClicked += (sender, args) =>
            {
                evidenceSplit.Panel1Collapsed = !evidenceSplit.Panel1Collapsed;
                expandDetails.Text = evidenceSplit.Panel1Collapsed ? "&Restore grid" : "&Expand details";
                details.Focus(); EnsureChangedPropertyViewport(); UpdateDetailsCue();
            };
            detailsScrollCue = new Label { Name = "ExplorerDetailsScrollCue", Dock = DockStyle.Fill, AutoSize = true, Visible = false,
                Margin = new Padding(8, 0, 0, 0),
                Text = "More below — scroll or expand Details.", AccessibleName = "More details below; focus the read-only text and scroll or expand", ForeColor = Color.DimGray };
            details.SizeChanged += (sender, args) => UpdateDetailsCue();
            details.FontChanged += (sender, args) => UpdateDetailsCue();
            var links = new Panel { Dock = DockStyle.Bottom, Height = 28 }; detailsLinks = links; links.Controls.Add(definition); links.Controls.Add(openRow); links.Controls.Add(expandDetails);
            supportToggle = new CheckBox { Name = "ExplorerSupportEvidenceToggle", Text = "Support &details", Anchor = AnchorStyles.Left, Margin = Padding.Empty,
                AutoSize = true, AccessibleName = "Expand raw occurrence and comparison support evidence", TabIndex = 1 };
            supportToggle.CheckedChanged += (sender, args) => Guard(RefreshDetails);
            detailsView = new ComboBox { Name = "ExplorerDetailsView", DropDownStyle = ComboBoxStyle.DropDownList,
                Width = 160, Margin = new Padding(0, 0, 8, 0), AccessibleName = "Selected component view: Overview or Compared Properties", TabIndex = 0 };
            detailsView.Items.AddRange(new object[] { "Overview", "Compared Properties" }); detailsView.SelectedIndex = 0;
            comparedProperties = new DataGridView { Name = "ExplorerComparedProperties", Dock = DockStyle.Fill,
                ReadOnly = true, VirtualMode = true, AutoGenerateColumns = false, AllowUserToAddRows = false,
                AllowUserToDeleteRows = false, AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None,
                RowHeadersVisible = false, MultiSelect = false, SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                BackgroundColor = Color.White, ScrollBars = ScrollBars.Both,
                ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize,
                AccessibleName = "Retained compared properties: Property, Source value, Target value, Result" };
            foreach (var column in new[] { new[] { "Property", "Property", "150" }, new[] { "SourceValue", "Source value", "200" },
                new[] { "TargetValue", "Target value", "200" }, new[] { "Result", "Result", "160" } })
                comparedProperties.Columns.Add(new DataGridViewTextBoxColumn { Name = column[0], HeaderText = column[1],
                    Width = MembershipResultsForm.ScaleColumnWidth(int.Parse(column[2]), DeviceDpi), SortMode = DataGridViewColumnSortMode.NotSortable,
                    DefaultCellStyle = new DataGridViewCellStyle { WrapMode = DataGridViewTriState.False } });
            comparedProperties.ColumnHeadersDefaultCellStyle.WrapMode = DataGridViewTriState.False;
            comparedProperties.RowTemplate.Height = Math.Max(26, comparedProperties.Font.Height + 10);
            comparedProperties.RowHeightInfoNeeded += (sender, args) => { args.Height = comparedProperties.RowTemplate.Height; args.MinimumHeight = comparedProperties.RowTemplate.Height; };
            comparedProperties.FontChanged += (sender, args) =>
            {
                comparedProperties.RowTemplate.Height = Math.Max(26, comparedProperties.Font.Height + 10);
                if (comparedProperties.RowCount > 0) comparedProperties.UpdateRowHeightInfo(0, true);
            };
            comparedProperties.CellValueNeeded += (sender, args) =>
            {
                if (args.RowIndex >= 0 && args.RowIndex < comparedPropertyRows.Count)
                    args.Value = ComparedPropertyCell(comparedPropertyRows[args.RowIndex], args.ColumnIndex);
            };
            comparedProperties.CellToolTipTextNeeded += (sender, args) =>
            {
                if (args.RowIndex >= 0 && args.RowIndex < comparedPropertyRows.Count && args.ColumnIndex >= 0)
                    args.ToolTipText = ComparedPropertyCell(comparedPropertyRows[args.RowIndex], args.ColumnIndex) +
                        Environment.NewLine + "Exact retained value; Definition Details retains all compared properties and evidence.";
            };
            comparedProperties.CellFormatting += (sender, args) =>
            {
                if (args.RowIndex >= 0 && args.RowIndex < comparedPropertyRows.Count &&
                    comparedPropertyRows[args.RowIndex].Changed && comparedPropertyRows[args.RowIndex].Comparison == "Different" &&
                    !comparedProperties.Rows[args.RowIndex].Selected) args.CellStyle.BackColor = Color.FromArgb(255, 247, 232);
            };
            comparedPropertiesMessage = new Label { Name = "ExplorerComparedPropertiesContext", Dock = DockStyle.Top,
                Height = Math.Max(32, Font.Height * 2), AutoEllipsis = true, Padding = new Padding(0, 2, 0, 2), AccessibleName = "Compared properties availability and contract limits" };
            comparedPropertiesMessage.FontChanged += (sender, args) => comparedPropertiesMessage.Height = Math.Max(32, comparedPropertiesMessage.Font.Height * 2);
            comparedPropertiesPanel = new Panel { Dock = DockStyle.Fill, Visible = false };
            comparedPropertiesPanel.Controls.Add(comparedProperties); comparedPropertiesPanel.Controls.Add(comparedPropertiesMessage);
            detailsView.SelectedIndexChanged += (sender, args) => { RefreshDetailsView(); EnsureChangedPropertyViewport(); UpdateDetailsCue(); };


            // The cue shares the options row instead of consuming a separate textbox-height row.
            // Its remaining-width cell wraps at large fonts without covering Support details.
            detailsOptions = new TableLayoutPanel { Name = "ExplorerDetailsOptions", Dock = DockStyle.Top,
                AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, ColumnCount = 3, RowCount = 1,
                Margin = Padding.Empty, Padding = Padding.Empty };
            detailsOptions.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            detailsOptions.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            detailsOptions.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            detailsOptions.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            detailsOptions.Controls.Add(detailsView, 0, 0); detailsOptions.Controls.Add(supportToggle, 1, 0); detailsOptions.Controls.Add(detailsScrollCue, 2, 0);
            right.Panel1.Controls.Add(grid); right.Panel1.Controls.Add(counts); right.Panel2.Controls.Add(comparedPropertiesPanel); right.Panel2.Controls.Add(details); right.Panel2.Controls.Add(links); right.Panel2.Controls.Add(detailsOptions);
            outer.Panel1.Controls.Add(navigation); outer.Panel2.Controls.Add(right); Controls.Add(outer);
            navigation.BeforeExpand += (s, e) => Guard(() => Populate(e.Node));
            navigation.AfterSelect += (s, e) => Guard(() =>
            {
                if (restoringNavigation) return;
                SelectNavigation(e.Node);
                if (!restoringNavigation && !preparing) NavigationIntentChanged?.Invoke();
            });
            navigation.Enter += (s, e) => lastNavigationFocus = "tree";
            grid.Enter += (s, e) => lastNavigationFocus = "grid";
            details.Enter += (s, e) => lastNavigationFocus = "details";
            navigation.KeyDown += (s, e) =>
            {
                if (e.KeyCode != Keys.Enter || navigation.SelectedNode == null) return;
                var node = navigation.SelectedNode;
                if (node.Tag is Page || node.Tag is Selection entry && !entry.Occurrence.HasValue)
                { if (node.IsExpanded) node.Collapse(); else node.Expand(); e.Handled = e.SuppressKeyPress = true; }
            };
        }

        internal void UseSimplifiedPresentation()
        {
            if (Projection != null) throw new InvalidOperationException("Choose presentation before binding a comparison.");
            simplifiedPresentation = true;
            simplifiedIdentity = new TextBox { Name = "SolutionExplorerSelectedIdentity", Dock = DockStyle.Top,
                Multiline = true, ReadOnly = true, WordWrap = false, ScrollBars = ScrollBars.Both,
                Height = Font.Height * 6 + 10, BackColor = Color.White, AccessibleName = "Selected component identity and comparison result" };
            simplifiedIdentity.FontChanged += (sender, args) => simplifiedIdentity.Height = simplifiedIdentity.Font.Height * 6 + 10;
            comparedPropertiesPanel.Controls.Add(simplifiedIdentity); comparedProperties.BringToFront();
            ConfigureWorkspace(true, ObjectsWorkspaceScope.Compare); ShowDetails(true);
            grid.Columns["Type"].HeaderText = "Component Type";
            counts.Height = Math.Max(28, Font.Height * 2);
            comparedProperties.Columns["SourceValue"].HeaderText = "Source";
            comparedProperties.Columns["TargetValue"].HeaderText = "Target";
            comparedProperties.Columns["Result"].HeaderText = "Status";
            ConfigureProfessionalExplorerAppearance();
        }

        internal static string SimplifiedStatus(SolutionObjectsExplorerRow row)
        {
            switch (row.Original.AuthoritativeRow.Presence)
            {
                case MembershipPresence.OnlyInSource: return "Source Only";
                case MembershipPresence.OnlyInTarget: return "Target Only";
                case MembershipPresence.PresentInBoth:
                    return row.Original.DefinitionStatus == "Match" ? "Match" : row.Original.DefinitionStatus == "Different" ? "Different" : "Needs Review";
                default: return "Needs Review";
            }
        }

        private IReadOnlyList<SolutionObjectsExplorerCategory> SimplifiedCategories()
        {
            return Projection.Categories.Where(c => c.Key != "objects").Select(c =>
                new SolutionObjectsExplorerCategory(c.Key == "all" ? "workspace-all" : "workspace-" + c.Key,
                    c.Key == "review" ? "Needs Review / relationship not established" : c.Key == "unmapped" ? "Unmapped comparison records" :
                    c.Name == "Model-driven Apps" ? "Apps" : c.Name == "Other / unclassified components" ? "Other Components" :
                    c.Name == "Choices (supported subset / review)" ? "Choices" : c.Name,
                    c.RowIndices, c.SourceOccurrences.Where(i => includedOccurrences.Contains(i) && (!ScopeSide().HasValue || Projection.Occurrences[i].Original.Side == ScopeSide())),
                    c.TargetOccurrences.Where(i => includedOccurrences.Contains(i) && (!ScopeSide().HasValue || Projection.Occurrences[i].Original.Side == ScopeSide()))))
                .ToList().AsReadOnly();
        }

        internal void Bind(SolutionObjectsExplorerProjection projection, IEnumerable<int> filteredRows, bool all)
        {
            ClearTechnicalContext();
            if (!ReferenceEquals(Projection, projection)) { selected = null; SelectedOccurrenceIndex = null; navigationContextChanged = true; }
            if (makerNavigation && !ReferenceEquals(Projection, projection)) ClearObjectOwner();
            Projection = projection ?? throw new ArgumentNullException(nameof(projection)); parents.Clear();
            foreach (var parent in projection.Occurrences)
                foreach (var link in parent.Children)
                {
                    if (!parents.TryGetValue(link.ChildOccurrenceIndex, out var list)) parents[link.ChildOccurrenceIndex] = list = new List<int>();
                    list.Add(parent.Original.OccurrenceIndex);
                }
            if (makerNavigation) makerModel = new SolutionObjectNavigation(projection);
            ApplyFilter(filteredRows, all);
        }
        internal void ApplyFilter(IEnumerable<int> filteredRows, bool all)
        {
            if (makerNavigation) { ApplyObjectFilter(filteredRows, all); return; }
            var bookmark = navigationContextChanged ? null : CaptureNavigation(); navigationContextChanged = false;
            baseAllowedRows = filteredRows.ToList().AsReadOnly(); filterAll = all;
            allowedRows = new HashSet<int>(baseAllowedRows.Where(i => string.IsNullOrWhiteSpace(searchText) ||
                i >= 0 && i < Projection.Rows.Count && new[] { PrimaryValue(Projection.Rows[i], "DisplayName"), PrimaryValue(Projection.Rows[i], "Name"), Projection.Rows[i].Type }
                    .Any(v => v != null && v.IndexOf(searchText, StringComparison.OrdinalIgnoreCase) >= 0)));
            if (baseAllowedRows.Any(i => i < 0 || i >= Projection.Rows.Count)) throw new ArgumentOutOfRangeException(nameof(filteredRows));
            matchingOccurrences = new HashSet<int>(Projection.Occurrences.Where(o => all && string.IsNullOrWhiteSpace(searchText) || o.Original.ComparisonRowIndices.Any(allowedRows.Contains))
                .Select(o => o.Original.OccurrenceIndex));
            includedOccurrences = new HashSet<int>(matchingOccurrences);
            var queue = new Queue<int>(includedOccurrences);
            while (queue.Count > 0)
                if (parents.TryGetValue(queue.Dequeue(), out var parentIds))
                    foreach (int id in parentIds) if (includedOccurrences.Add(id)) queue.Enqueue(id);
            var prior = selected?.OriginalRowIndex; var side = SelectedSide;
            var categories = primaryWorkspace ? (simplifiedPresentation ? SimplifiedCategories() : WorkspaceNavigationCategories(all)) : null;
            navigation.BeginUpdate(); preparing = true;
            try
            {
                navigation.Nodes.Clear();
                inspectionNode = null;
                if (primaryWorkspace)
                {
                    foreach (var category in categories.Where(c => c.Key != "workspace-technical"))
                    {
                        int raw = workspaceScope == ObjectsWorkspaceScope.Target ? category.TargetOccurrences.Count : category.SourceOccurrences.Count;
                        string unit = simplifiedPresentation ? category.RowIndices.Count(allowedRows.Contains) + " comparison rows" : workspaceScope == ObjectsWorkspaceScope.Compare ? category.RowIndices.Count + " candidate rows, partial" :
                            ScopeState() == MembershipSnapshotState.Unavailable ? "Not established" : raw + " candidates, partial";
                        navigation.Nodes.Add(new TreeNode(category.Name + " (" + unit + ")")
                        { Name = category.Key, Tag = new Selection { Rows = category.RowIndices, Side = ScopeSide(), Technical = category.Key == "workspace-technical",
                            Entries = category.Key == "workspace-all" || category.Key == "workspace-technical" ? null :
                                category.SourceOccurrences.Concat(category.TargetOccurrences).Select(i => new Selection { Occurrence = i, Side = Projection.Occurrences[i].Original.Side }).ToList().AsReadOnly(),
                            Description = simplifiedPresentation ? "Retained comparison rows; expanded entries are side-specific raw occurrences, not unique maker objects." : "Observed candidates — partial, reconciliation incomplete. " + unit + "; not a maker top-level object count." },
                            ToolTipText = category.CountText + "; " + (workspaceScope == ObjectsWorkspaceScope.Compare ? "distinct candidate original row indices" : "side-local candidate raw occurrences, not distinct objects") +
                                "; same-snapshot navigation only; partial, top-level equivalence not established" });
                        if (category.Key != "workspace-all" && category.Key != "workspace-technical") navigation.Nodes[navigation.Nodes.Count - 1].Nodes.Add(new TreeNode("Expand to inspect..."));
                    }
                }
                else foreach (var key in new[] { "objects", "all", "review", "unmapped" })
                {
                    var category = Projection.Category(key);
                    navigation.Nodes.Add(new TreeNode(category.Name + " (" + category.RowIndices.Count(allowedRows.Contains) + " rows)")
                        { Name = category.Key, Tag = new Selection { Rows = category.RowIndices, Description = category.CountText + "; categories may overlap. Not a Power Apps top-level count." } });
                }
                foreach (var role in primaryWorkspace ? new DisplayEnvironmentSide[0] : new[] { DisplayEnvironmentSide.Source, DisplayEnvironmentSide.Target })
                {
                    var sideNode = new TreeNode(role + ": " + (role == DisplayEnvironmentSide.Source ? Projection.Owner.Source.Diagnostics.EnvironmentName : Projection.Owner.Target.Diagnostics.EnvironmentName)) { Name = "side:" + role };
                    foreach (var category in Projection.Categories.Where(c => c.Key != "objects" && c.Key != "all" && c.Key != "review" && c.Key != "unmapped"))
                    {
                        var ids = (role == DisplayEnvironmentSide.Source ? category.SourceOccurrences : category.TargetOccurrences).Where(includedOccurrences.Contains).ToList();
                        if (ids.Count == 0) continue;
                        var entries = ids.Select(i => new Selection { Occurrence = i, Side = role }).ToList();
                        int contextCount = ids.Count(i => !matchingOccurrences.Contains(i));
                        var node = new TreeNode(category.Name + " (" + ids.Count + " raw" + (contextCount == 0 ? "" : "; " + contextCount + " context") + ")")
                        { Name = category.Key, Tag = new Selection { Rows = ids.SelectMany(i => Projection.Occurrences[i].Original.ComparisonRowIndices).Distinct().ToArray(), Side = role,
                            Entries = entries.AsReadOnly(), Description = "Side-specific raw occurrences, not object or comparison totals." },
                            ToolTipText = category.CountText + "; visible " + ids.Count + " raw occurrences, including " + contextCount + " navigation context occurrences" };
                        node.Nodes.Add(new TreeNode("Expand to inspect...")); sideNode.Nodes.Add(node);
                    }
                    navigation.Nodes.Add(sideNode); sideNode.Expand();
                }
            }
            finally { preparing = false; navigation.EndUpdate(); }
            if (primaryWorkspace)
            {
                var allCategory = categories.First();
                SelectedSide = ScopeSide(); SetRows(allCategory.RowIndices, simplifiedPresentation ? "All retained comparison rows; friendly names and category mapping may be partial." : "Observed candidates — partial, reconciliation incomplete. Additional components remain in Technical Membership.");
            }
            else SetRows(Projection.Category("objects").RowIndices, "All Objects (candidates), not verified maker top-level objects");
            if (bookmark != null) RestoreNavigation(bookmark, prior.HasValue && allowedRows.Contains(prior.Value) ? prior : null, side, bookmark.Occurrence);
            else if (prior.HasValue && allowedRows.Contains(prior.Value)) SelectComparison(prior.Value, side);
            if (bookmark != null && allowedRows.Count == 0 && selected == null &&
                !SelectedOccurrenceIndex.HasValue && technicalContext?.Route == TechnicalRoute.Description)
                ClearTechnicalContext();
        }
        internal bool ConfigureWorkspace(bool primary, ObjectsWorkspaceScope scope)
        {
            bool changed = primaryWorkspace != primary || workspaceScope != scope;
            if (changed) { ClearTechnicalContext(); navigationContextChanged = true; }
            if (primaryWorkspace != primary) ApplyColumnLayout(primary);
            primaryWorkspace = primary; workspaceScope = scope;
            if ((!primary || !detailsRequested) && evidenceSplit.Height >= 170 + evidenceSplit.Panel2MinSize + evidenceSplit.SplitterWidth)
                evidenceSplit.Panel1MinSize = 170;
            counts.Height = primary ? Math.Max(40, Font.Height * 3) : 48;
            if (!primary) evidenceSplit.Panel1Collapsed = false;
            if (primary && !detailsRequested) evidenceSplit.Panel1Collapsed = false;
            evidenceSplit.Panel2Collapsed = primary && !detailsRequested;
            expandDetails.Visible = primary;
            openRow.Text = primary ? "Original row / support diagnostics" : "Open original row in Component Details";
            supportToggle.Visible = primary && !simplifiedPresentation; detailsView.Visible = primary && !simplifiedPresentation;
            if (primary) supportToggle.Checked = false;
            if (simplifiedPresentation) { openRow.Text = "Locate original comparison row"; definition.Visible = false; }
            UpdateColumns(); RefreshDetails();

            return changed;
        }
        internal void ShowDetails(bool show)
        {
            detailsRequested = show;
            if (!show) evidenceSplit.Panel1Collapsed = false;
            evidenceSplit.Panel2Collapsed = primaryWorkspace && !show;
            expandDetails.Text = evidenceSplit.Panel1Collapsed ? "&Restore grid" : "&Expand details";
            if (show && detailsContentDirty) RefreshDetails();
            else { RefreshDetailsView(); EnsureChangedPropertyViewport(); UpdateDetailsCue(); }
        }
        internal void ShowMetadata(bool show)
        {
            metadataRequested = show;
            UpdateColumns();
        }
        private DisplayEnvironmentSide? ScopeSide() => workspaceScope == ObjectsWorkspaceScope.Compare ? (DisplayEnvironmentSide?)null :
            workspaceScope == ObjectsWorkspaceScope.Target ? DisplayEnvironmentSide.Target : DisplayEnvironmentSide.Source;
        private MembershipSnapshotState ScopeState() => workspaceScope == ObjectsWorkspaceScope.Source ? Projection.Owner.Source.State : Projection.Owner.Target.State;
        private bool InScope(int row) => workspaceScope == ObjectsWorkspaceScope.Compare ||
            (simplifiedPresentation ? (workspaceScope == ObjectsWorkspaceScope.Source ? Projection.Rows[row].Original.AuthoritativeRow.SourceIdentity : Projection.Rows[row].Original.AuthoritativeRow.TargetIdentity) != null :
            (workspaceScope == ObjectsWorkspaceScope.Target ? Projection.Rows[row].TargetOccurrences : Projection.Rows[row].SourceOccurrences).Count > 0);
        internal void SetSearch(string text)
        {
            text = primaryWorkspace ? (text ?? string.Empty).Trim() : string.Empty;
            if (searchText == text) return;
            searchText = text;
            if (Projection != null && baseAllowedRows != null) ApplyFilter(baseAllowedRows, filterAll);
        }
        private void UpdateColumns()
        {
            grid.Columns["Source"].Visible = grid.Columns["Target"].Visible = true;
            foreach (DataGridViewColumn c in grid.Columns)
                if (c.Index >= 6) c.Visible = metadataRequested && !simplifiedPresentation;
            grid.Columns["ComparisonStatus"].HeaderText = simplifiedPresentation ? "Difference Status" : primaryWorkspace ? "Comparison" : "Comparison Status";
            UpdateGridMetrics();
        }
        private void ApplyColumnLayout(bool primary)
        {
            // Change widths only when entering a workspace, not on each virtual-row refresh.
            var widths = primary ? new[] { 260, 220, 160, 110, 110, 200 } : new[] { 180, 210, 155, 105, 105, 250 };
            var minimums = new[] { 100, 100, 90, 70, 70, 120 };
            grid.SuspendLayout();
            try
            {
                for (int i = 0; i < widths.Length; i++)
                {
                    var column = grid.Columns[i];
                    column.AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
                    column.MinimumWidth = MembershipResultsForm.ScaleColumnWidth(primary ? minimums[i] : 65, DeviceDpi);
                    column.Width = MembershipResultsForm.ScaleColumnWidth(widths[i], DeviceDpi);
                }
            }
            finally { grid.ResumeLayout(); }
        }
        private string DisplayValue(SolutionObjectsExplorerRow row, string column) => primaryWorkspace ? SingleLine(PrimaryValue(row, column)) : Value(row, column);
        private static string SingleLine(string value) => (value ?? string.Empty).Replace("\r\n", " · ").Replace("\r", " · ").Replace("\n", " · ").Replace("\t", " ");
        private void UpdateGridMetrics()
        {
            if (grid == null) return;
            foreach (DataGridViewColumn column in grid.Columns) column.DefaultCellStyle.WrapMode = primaryWorkspace ? DataGridViewTriState.False : DataGridViewTriState.True;
            grid.RowTemplate.Height = primaryWorkspace ? PrimaryRowHeight : Math.Max(46, grid.Font.Height * 3 + 4);
            grid.ColumnHeadersDefaultCellStyle.WrapMode = primaryWorkspace ? DataGridViewTriState.False : DataGridViewTriState.True;
            grid.ColumnHeadersHeightSizeMode = primaryWorkspace ? DataGridViewColumnHeadersHeightSizeMode.AutoSize : DataGridViewColumnHeadersHeightSizeMode.EnableResizing;
            if (detailsLinks != null) detailsLinks.Height = Math.Max(28, grid.Font.Height + 8);
            if (appliedRowMetricHeight != grid.RowTemplate.Height)
            {
                appliedRowMetricHeight = grid.RowTemplate.Height;
                if (grid.RowCount > 0)
                {
                    // Synchronize WinForms cached virtual heights/scroll range, using one constant metric.
                    // This does not measure text, materialize row controls or change comparison selection.
                    grid.SuspendLayout();
                    try { grid.UpdateRowHeightInfo(0, true); }
                    finally { grid.ResumeLayout(); }
                }
            }
            grid.Invalidate(); // Only visible cells are painted; no inventory-wide text measurement.
        }
        private void PaintPrimaryCell(object sender, DataGridViewCellPaintingEventArgs e)
        {
            if (!primaryWorkspace || e.RowIndex < 0 || e.ColumnIndex < 0) return;
            e.Paint(e.ClipBounds, e.PaintParts & ~DataGridViewPaintParts.ContentForeground);
            var bounds = Rectangle.Inflate(e.CellBounds, -4, -2);
            if (bounds.Width > 0 && bounds.Height > 0)
                TextRenderer.DrawText(e.Graphics, Convert.ToString(e.FormattedValue), e.CellStyle.Font ?? grid.Font, bounds,
                    (e.State & DataGridViewElementStates.Selected) != 0 ? e.CellStyle.SelectionForeColor : e.CellStyle.ForeColor,
                    TextFormatFlags.SingleLine | TextFormatFlags.EndEllipsis | TextFormatFlags.VerticalCenter | TextFormatFlags.NoPrefix | TextFormatFlags.PreserveGraphicsClipping);
            e.Handled = true;
        }
        internal string FullCellToolTip(int originalRow, string column)
        {
            var row = Projection.Rows[originalRow];
            return "Full displayed text: " + DisplayValue(row, column) + Environment.NewLine +
                "Retained side values and availability: " + CellEvidence(row, column);
        }
        private string PrimaryValue(SolutionObjectsExplorerRow row, string column)
        {
            if (column == "DisplayName" || column == "Name")
                return AcrossSides(row, side => row.PrimaryName(side, column == "DisplayName"));
            if (column == "ComparisonStatus") return simplifiedPresentation ? SimplifiedStatus(row) : "Membership: " + row.MembershipPresentation + Environment.NewLine + "Definition: " + row.DefinitionPresentation;
            if (column == "Type") return row.PrimaryType;
            if (new[] { "Managed", "Customized", "LastModified", "Owner", "Status" }.Contains(column))
                return AcrossSides(row, side => MetadataText(row, side, column));
            return Value(row, column);
        }
        private string AcrossSides(SolutionObjectsExplorerRow row, Func<DisplayEnvironmentSide, string> read)
        {
            if (ScopeSide().HasValue) return read(ScopeSide().Value) ?? "—";
            string left = read(DisplayEnvironmentSide.Source), right = read(DisplayEnvironmentSide.Target);
            if (row.Original.AuthoritativeRow.SourceIdentity != null && row.Original.AuthoritativeRow.TargetIdentity != null &&
                StringComparer.Ordinal.Equals(left, right)) return left ?? "—";
            return "Source: " + (left ?? "—") + Environment.NewLine + "Target: " + (right ?? "—");
        }
        private static string MetadataText(SolutionObjectsExplorerRow row, DisplayEnvironmentSide side, string column)
        {
            var m = (side == DisplayEnvironmentSide.Source ? row.SourceDisplay : row.TargetDisplay)?.ObjectMetadata;
            switch (column)
            {
                case "Managed": return MetadataValue(m?.Managed); case "Customized": return MetadataValue(m?.Customized);
                case "LastModified": return MetadataValue(m?.LastModified); case "Owner": return MetadataValue(m?.Owner);
                default: return MetadataValue(m?.ObjectStatus);
            }
        }
        private static string MetadataValue<T>(ComponentObjectMetadataValue<T> value)
        {
            if (value?.Availability != DisplayAvailability.Available || (object)value.Value == null) return null;
            string text = MembershipGridRow.FormatObjectValue(value);
            return text != null && !text.StartsWith("[", StringComparison.Ordinal) ? text : null;
        }
        private static string DefinitionLabel(MembershipGridRow row) => row.DefinitionStatus == "Match" ? "Match (compared properties)" :
            string.IsNullOrWhiteSpace(row.DefinitionStatusText) ? "Not compared — result not supplied" : row.DefinitionStatusText;
        private string CellEvidence(SolutionObjectsExplorerRow row, string column)
        {
            if (column == "DisplayName" || column == "Name")
                return "Source: " + SolutionObjectsExplorerRow.Field(column == "DisplayName" ? row.SourceDisplay?.FriendlyName : row.SourceDisplay?.InternalName) +
                    "; Target: " + SolutionObjectsExplorerRow.Field(column == "DisplayName" ? row.TargetDisplay?.FriendlyName : row.TargetDisplay?.InternalName) + "; Source name basis: " + row.NameBasis(DisplayEnvironmentSide.Source) + "; Target name basis: " + row.NameBasis(DisplayEnvironmentSide.Target) + ". ";
            return Value(row, column) + "; unavailable/omitted/blank values are not negative facts. ";
        }
        private string SimpleDetails(SolutionObjectsExplorerRow row)
        {
            var text = new StringBuilder();
            text.AppendLine("Display Name: " + SingleLine(PrimaryValue(row, "DisplayName")))
                .AppendLine("Name: " + SingleLine(PrimaryValue(row, "Name")) + "; Type: " + PrimaryValue(row, "Type"))
                .AppendLine("Membership: " + row.MembershipPresentation + "; Definition: " + row.DefinitionPresentation)
                .AppendLine("Source: " + row.Source + "; Target: " + row.Target);
            var definition = row.Original.AuthoritativeRow.DefinitionDetail;
            var changed = row.Original.DefinitionStatus == "Different" ? definition?.Properties.Where(p => p.Changed && p.Comparison == "Different").ToList() : null;
            if (changed != null && changed.Count > 0)
            {
                text.AppendLine("Changed compared properties:");
                foreach (var p in changed) text.AppendLine(p.PropertyName + ": Source " + ShortValue(p.SourceValue) + "; Target " + ShortValue(p.TargetValue) + " (" + p.Comparison + ")");
            }
            else text.AppendLine(definition == null ? "Compared properties not retained / not compared." : "No changed compared properties reported; unavailable properties do not establish equality.");
            if (row.Original.DefinitionStatus == "Match") text.AppendLine("Compared properties match only within the implemented contract, not complete object equality.");
            text.AppendLine("Overview — retained Source / Target information:");
            text.AppendLine("Comparison row index: " + row.OriginalRowIndex);
            AppendSideOverview(text, row, DisplayEnvironmentSide.Source); AppendSideOverview(text, row, DisplayEnvironmentSide.Target);
            text.AppendLine("Availability / coverage:");
            foreach (var side in new[] { DisplayEnvironmentSide.Source, DisplayEnvironmentSide.Target })
            {
                var d = side == DisplayEnvironmentSide.Source ? row.SourceDisplay : row.TargetDisplay;
                text.AppendLine(side + " display name: " + Availability(d?.FriendlyName)).AppendLine(side + " internal name: " + Availability(d?.InternalName));
            }
            text.AppendLine("Missing metadata is not a negative value. Support details retain full availability and provenance.");
            return text.ToString();
        }
        private static string Availability(ComponentDisplayValue field) => field == null ? "NotRetrieved" : field.Availability != DisplayAvailability.Available ?
            field.Availability.ToString() : field.Value == null ? "Retrieved null" : string.IsNullOrWhiteSpace(field.Value) ? "Retrieved blank" : "Available";
        private static string ShortValue(string value)
        {
            if (value == null) return "[No value supplied; availability not established]";
            if (value.Length == 0) return "[Blank value supplied]";
            value = SingleLine(value);
            value = Regex.Replace(value, @"\b[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}\b", "[local identifier — see support details]");
            return value.Length <= 200 ? value : value.Substring(0, 200) + "… (full value in definition details)";
        }
        private static string ComparedPropertyCell(ComponentPropertyPresentation property, int column)
        {
            switch (column)
            {
                case 0: return property.PropertyName;
                case 1: return SolutionObjectsExplorerRow.PropertyValue(property.SourceValue);
                case 2: return SolutionObjectsExplorerRow.PropertyValue(property.TargetValue);
                default: return SolutionObjectsExplorerRow.PropertyComparison(property);
            }
        }
        private void RefreshDetailsView()
        {
            if (makerNavigation) { RenderObjectDetails(); return; }
            bool properties = primaryWorkspace && detailsRequested && !supportToggle.Checked && (simplifiedPresentation || detailsView.SelectedIndex == 1);
            details.Visible = !properties; comparedPropertiesPanel.Visible = properties;
            detailsView.Enabled = !supportToggle.Checked;
            if (!properties) return;
            var row = selected == null ? null : Projection.Rows[selected.OriginalRowIndex];
            if (ReferenceEquals(row, comparedPropertiesOwner) && row != null) return;
            comparedPropertiesOwner = row;
            comparedPropertyRows = row == null ? (IReadOnlyList<ComponentPropertyPresentation>)new ComponentPropertyPresentation[0] : row.RetainedComparedProperties
                .OrderByDescending(p => p.Changed && p.Comparison == "Different").ToList().AsReadOnly();
            comparedProperties.RowCount = comparedPropertyRows.Count;
            comparedProperties.Tag = row?.Original.AuthoritativeRow.DefinitionDetail;
            comparedPropertiesMessage.Text = row == null ? "Select a component to inspect retained compared properties." :
                "Comparison row index: " + row.OriginalRowIndex + ". " + row.ComparedPropertiesExplanation;
            if (simplifiedPresentation)
            {
                simplifiedIdentity.Text = row == null ? "Select a component to inspect its details." :
                    "Display Name: " + PrimaryValue(row, "DisplayName") + Environment.NewLine +
                    "Logical / Schema Name: " + PrimaryValue(row, "Name") + Environment.NewLine +
                    "Component Type: " + row.PrimaryType + " · Result: " + SimplifiedStatus(row) + Environment.NewLine +
                    "Source: " + row.Source + " · Target: " + row.Target + Environment.NewLine +
                    "Membership: " + row.MembershipPresentation + " · Definition: " + row.DefinitionPresentation;
                comparedPropertiesMessage.Text = row == null ? "No component selected." : row.ComparedPropertiesExplanation;
            }
            comparedProperties.AccessibleDescription = comparedPropertiesMessage.Text;
            countToolTip.SetToolTip(comparedPropertiesMessage, comparedPropertiesMessage.Text);
        }
        private static void AppendSideOverview(StringBuilder text, SolutionObjectsExplorerRow row, DisplayEnvironmentSide side)
        {
            var identity = side == DisplayEnvironmentSide.Source ? row.Original.AuthoritativeRow.SourceIdentity : row.Original.AuthoritativeRow.TargetIdentity;
            var display = side == DisplayEnvironmentSide.Source ? row.SourceDisplay : row.TargetDisplay;
            text.AppendLine(side + " information:");
            if (identity == null) { text.AppendLine("[No local observation]; environment absence is not established."); return; }
            text.AppendLine(side + " display name: " + SolutionObjectsExplorerRow.Field(display?.FriendlyName))
                .AppendLine(side + " internal name: " + SolutionObjectsExplorerRow.Field(display?.InternalName))
                .AppendLine(side + " name presentation: " + row.NameBasis(side))
                .AppendLine(side + " presence: " + (side == DisplayEnvironmentSide.Source ? row.Source : row.Target) + "; resolution: " + identity.Status);
        }


        private void RefreshDetails()
        {
            detailsContentDirty = true;
            if ((!primaryWorkspace || supportToggle.Checked) ? Visible && !evidenceSplit.Panel2Collapsed : detailsRequested) RenderDetailsContent();
            else RefreshDetailsView();
        }
        private void RenderDetailsContent()
        {
            detailsContentDirty = false;
            details.Text = !primaryWorkspace || supportToggle.Checked ? TechnicalEvidence() : selected == null ? "Select a component to inspect retained details." : SimpleDetails(Projection.Rows[selected.OriginalRowIndex]);
            details.AccessibleDescription = primaryWorkspace && !supportToggle.Checked ? "Compact retained object details; missing values preserve their availability. Full technical evidence is under Support details." : "Full read-only original-row and raw occurrence support evidence";
            details.SelectionStart = 0; details.SelectionLength = 0; details.ScrollToCaret();
            RefreshDetailsView(); EnsureChangedPropertyViewport(); UpdateDetailsCue();
        }
        private void EnsureChangedPropertyViewport()
        {
            if (!primaryWorkspace || !detailsRequested || supportToggle.Checked || !details.IsHandleCreated || !details.Visible) return;
            if (!evidenceSplit.Panel1Collapsed)
            {
                // Reserve accounting, column headings and one row, while keeping Details feasible.
                int gridMinimum = Math.Min(170, counts.Height + grid.ColumnHeadersHeight + grid.RowTemplate.Height);
                gridMinimum = Math.Min(gridMinimum, Math.Max(0, evidenceSplit.Height - evidenceSplit.Panel2MinSize - evidenceSplit.SplitterWidth));
                evidenceSplit.Panel1MinSize = gridMinimum;
                // A collapsed pane can retain its old distance after the container becomes shorter.
                // Settle that allocation before measuring the actual textbox client height.
                int maximumDistance = Math.Max(gridMinimum, evidenceSplit.Height - evidenceSplit.Panel2MinSize - evidenceSplit.SplitterWidth);
                evidenceSplit.SplitterDistance = Math.Max(gridMinimum, Math.Min(evidenceSplit.SplitterDistance, maximumDistance));
            }
            detailsOptions.PerformLayout(); evidenceSplit.Panel2.PerformLayout();
            int heading = details.Text.IndexOf("Changed compared properties:", StringComparison.Ordinal);
            if (heading < 0) return;
            int value = details.Text.IndexOf(Environment.NewLine, heading, StringComparison.Ordinal) + Environment.NewLine.Length;
            int extra = details.GetPositionFromCharIndex(value).Y + details.Font.Height + 3 - details.ClientSize.Height;
            if (extra > 0 && !evidenceSplit.Panel1Collapsed)
            {
                evidenceSplit.SplitterDistance = Math.Max(evidenceSplit.Panel1MinSize, evidenceSplit.SplitterDistance - extra);
                detailsOptions.PerformLayout(); evidenceSplit.Panel2.PerformLayout();
            }
            else if (extra > 0 && evidenceSplit.Panel1Collapsed)
            {
                // At large fonts the complete pane can still be shorter than the overview plus change.
                // Reveal the change by scrolling; all preceding and following text remains retained.
                details.Select(value, 0);
                details.ScrollToCaret();
            }
            int changedLineY = details.GetPositionFromCharIndex(value).Y;
            if (!evidenceSplit.Panel1Collapsed &&
                evidenceSplit.SplitterDistance == evidenceSplit.Panel1MinSize &&
                details.ClientSize.Height >= details.Font.Height &&
                changedLineY + details.Font.Height > details.ClientSize.Height)
            {
                details.Select(value, 0);
                details.ScrollToCaret();
            }
        }
        private void UpdateDetailsCue()
        {
            if (makerNavigation) { UpdateObjectViewport(); return; }
            if (updatingDetailsCue || details == null || detailsScrollCue == null || supportToggle == null || !details.IsHandleCreated) return;
            updatingDetailsCue = true;
            try
            {
                int heading = details.Text.IndexOf("Changed compared properties:", StringComparison.Ordinal);
                int value = heading < 0 ? -1 : details.Text.IndexOf(Environment.NewLine, heading, StringComparison.Ordinal) + Environment.NewLine.Length;
                bool constrained = primaryWorkspace && details.Visible && !supportToggle.Checked && !evidenceSplit.Panel2Collapsed && value >= 0 &&
                    details.GetPositionFromCharIndex(value).Y + details.Font.Height > details.ClientSize.Height;
                detailsScrollCue.Visible = constrained;
            }
            finally { updatingDetailsCue = false; }
        }
        internal void SelectComparison(int index, DisplayEnvironmentSide? side = null)
        {
            if (makerNavigation)
            {
                var matches = visibleObjects.Select((item, position) => new { item, position }).Where(x => x.item.Rows.Contains(index) &&
                    (!side.HasValue || x.item.Has(side.Value))).ToList();
                if (matches.Count == 1)
                { grid.CurrentCell = grid.Rows[matches[0].position].Cells[0]; SelectObjectGridRow(matches[0].position); }
                return;
            }
            if (primaryWorkspace && (!allowedRows.Contains(index) || !InScope(index) || !visibleRows.Contains(index)))
            {
                if (index < 0 || index >= Projection.Rows.Count) return;
                selected = Projection.Rows[index].Original; SelectedSide = side;
                counts.Text = accountingText + Environment.NewLine + "Selected component is outside this primary list/scope/filter. Open original row to inspect.";
                counts.AccessibleDescription = counts.Text; countToolTip.SetToolTip(counts, counts.Text);
                SetTechnicalContext(TechnicalRoute.RowOnly, selected); RefreshDetails(); openRow.Enabled = true; definition.Enabled = selected.AuthoritativeRow.DefinitionDetail != null;
                SelectedComparisonChanged?.Invoke(selected); return;
            }
            if (!allowedRows.Contains(index)) return;
            if (!visibleRows.Contains(index)) SetRows(Projection.Category("all").RowIndices, "All Components (original comparison rows)");
            if (SelectedOccurrenceIndex.HasValue && !Projection.Occurrences[SelectedOccurrenceIndex.Value].Original.ComparisonRowIndices.Contains(index)) SelectedOccurrenceIndex = null;
            SelectedSide = primaryWorkspace ? ScopeSide() ?? side : side; int position = visibleRows.ToList().IndexOf(index);
            grid.CurrentCell = grid.Rows[position].Cells[0]; SelectGridRow(position);
        }
        internal void InspectOccurrences(int rowIndex)
        {
            var row = Projection.Rows[rowIndex];
            var choices = row.SourceOccurrences.Concat(row.TargetOccurrences).Where(i => !primaryWorkspace || !ScopeSide().HasValue || Projection.Occurrences[i].Original.Side == ScopeSide()).Select(i => new Selection { Occurrence = i, Side = Projection.Occurrences[i].Original.Side }).ToList();
            var node = new TreeNode("Original row " + rowIndex + " — all " + choices.Count + " mapped raw occurrences")
                { Name = "inspection-row:" + rowIndex, Tag = new Selection { Rows = new[] { rowIndex }, Entries = choices.AsReadOnly(), Description = "Choose a side/occurrence explicitly; no GUID-based pairing." } };
            inspectionNode?.Remove(); inspectionNode = node;
            node.Nodes.Add(new TreeNode("Expand to inspect...")); navigation.Nodes.Add(node); node.Expand(); navigation.SelectedNode = node;
        }
        internal void SelectOccurrence(int index)
        {
            if (!includedOccurrences.Contains(index)) return;
            if (primaryWorkspace && ScopeSide().HasValue && Projection.Occurrences[index].Original.Side != ScopeSide()) return;
            var item = Projection.Occurrences[index]; var node = MakeNode(new Selection { Occurrence = index, Side = item.Original.Side });
            inspectionNode?.Remove(); inspectionNode = node;
            navigation.Nodes.Add(node); navigation.SelectedNode = node; node.Expand();
        }
        private sealed class Selection
        {
            internal int? Occurrence; internal DisplayEnvironmentSide? Side; internal IReadOnlyList<int> Rows;
            internal IReadOnlyList<Selection> Entries; internal string Description; internal bool Technical;
        }
        private sealed class Page { internal IReadOnlyList<Selection> Entries; internal int Start, Count; }
        private TreeNode MakeNode(Selection selection)
        {
            var item = Projection.Occurrences[selection.Occurrence.Value];
            var name = item.Original.ComparisonRowIndices.Count == 1 ? Projection.Rows[item.Original.ComparisonRowIndices[0]].PrimaryName(item.Original.Side, true) : null;
            var node = new TreeNode((simplifiedPresentation ? item.Original.Side + ": " + (name ?? item.Category + " — name not retrieved") + " [raw occurrence " + item.Original.OriginalRawIndex + "]" : primaryWorkspace ? name ?? item.Category + " — name not retrieved" : item.Original.Side + ": " + item.Label + " [raw " + item.Original.OriginalRawIndex + "]") +
                (matchingOccurrences.Contains(item.Original.OccurrenceIndex) ? "" : " (navigation context)")) { Name = "occ:" + item.Original.OccurrenceIndex, Tag = selection,
                    ToolTipText = item.Original.Placement + "; raw occurrence, not top-level proof" };
            if (item.Children.Count > 0 || item.Original.Record.ComponentType == 1 || item.Original.Record.ComponentType == 80)
                node.Nodes.Add(new TreeNode("Expand to inspect...")); return node;
        }
        private void AddEntries(TreeNode node, IReadOnlyList<Selection> entries)
        {
            if (entries.Count <= PageSize) foreach (var entry in entries) node.Nodes.Add(MakeNode(entry));
            else for (int start = 0; start < entries.Count; start += PageSize)
            {
                var page = new TreeNode("Items " + (start + 1) + "–" + Math.Min(entries.Count, start + PageSize))
                    { Name = "page:" + start, Tag = new Page { Entries = entries, Start = start, Count = Math.Min(PageSize, entries.Count - start) } };
                page.Nodes.Add(new TreeNode("Expand to inspect...")); node.Nodes.Add(page);
            }
        }
        private void Populate(TreeNode node)
        {
            if (makerNavigation) { PopulateObjects(node); return; }
            if (node.Nodes.Count != 1 || node.Nodes[0].Text != "Expand to inspect...") return;
            node.Nodes.Clear();
            if (node.Tag is Page page)
            { for (int i = page.Start; i < page.Start + page.Count; i++) node.Nodes.Add(MakeNode(page.Entries[i])); return; }
            var selection = node.Tag as Selection; if (selection == null) return;
            if (selection.Entries != null) { AddEntries(node, selection.Entries); return; }
            if (!selection.Occurrence.HasValue) return;
            var item = Projection.Occurrences[selection.Occurrence.Value];
            var presentGroups = new HashSet<string>();
            foreach (var group in item.Children.Where(l => includedOccurrences.Contains(l.ChildOccurrenceIndex) &&
                (!primaryWorkspace || ScopeSide() == null || Projection.Occurrences[l.ChildOccurrenceIndex].Original.Side == ScopeSide()) &&
                (!primaryWorkspace || Projection.Occurrences[l.ChildOccurrenceIndex].Original.Record.ComponentType != 10075 ||
                    l.OriginalAssociation.Kind == HierarchyRelationshipKind.StructuralParent && l.OriginalAssociation.Status == HierarchyAssociationStatus.Verified))
                .GroupBy(l => new { l.Group, l.OriginalAssociation.Kind }).OrderBy(g => g.Key.Group, StringComparer.Ordinal).ThenBy(g => g.Key.Kind))
            {
                var entries = group.Select(l => new Selection { Occurrence = l.ChildOccurrenceIndex, Side = item.Original.Side }).ToList();
                presentGroups.Add(group.Key.Group);
                string evidence = group.Key.Kind == HierarchyRelationshipKind.StructuralParent ? "Verified structural parent" : "Inclusion-root only — not ownership";
                int rawCount = group.Select(l => l.ChildOccurrenceIndex).Distinct().Count();
                var child = new TreeNode((primaryWorkspace && group.Key.Group == "AppSettings" ? "Settings" : group.Key.Group) + " — " + (group.Key.Kind == HierarchyRelationshipKind.StructuralParent ? "structural" : "inclusion root") +
                    " (" + rawCount + " raw" + (entries.Count == rawCount ? "" : "; " + entries.Count + " links") + ")")
                { Name = "group:" + group.Key.Kind + ":" + group.Key.Group, Tag = new Selection { Entries = entries.AsReadOnly(), Side = item.Original.Side,
                    Rows = entries.SelectMany(e => Projection.Occurrences[e.Occurrence.Value].Original.ComparisonRowIndices).Distinct().ToArray(), Description = evidence },
                    ToolTipText = evidence + "; " + rawCount + " distinct raw occurrences; " + entries.Count + " association references" };
                child.Nodes.Add(new TreeNode("Expand to inspect...")); node.Nodes.Add(child);
            }
            var expected = item.Original.Record.ComponentType == 1 ? new[] { "Columns", "Relationships", "Forms", "Views", "Other Components" } :
                item.Original.Record.ComponentType == 80 ? new[] { "AppSettings", "Other Components" } : new string[0];
            foreach (var name in expected.Where(n => !primaryWorkspace && !presentGroups.Contains(n)))
                node.Nodes.Add(new TreeNode(name + " — no retained verified links") { Name = "empty:" + name, Tag = new Selection { Rows = new int[0], Side = item.Original.Side,
                    Description = "No retained verified " + name + " links. This is not proof of no components in the environment." } });
        }
        private void SelectNavigation(TreeNode node)
        {
            if (makerNavigation) { SelectObjectNode(node); return; }
            if (preparing) return;
            if (node.Tag is Page)
            {
                selected = null; SelectedSide = null; SelectedOccurrenceIndex = null;
                openRow.Enabled = definition.Enabled = false;
                SetTechnicalContext(TechnicalRoute.Description, description: node.Text, descriptor: node.Tag);
                RefreshDetails(); SelectedComparisonChanged?.Invoke(null); return;
            }
            if (!(node.Tag is Selection selection))
            {
                selected = null; SelectedSide = null; SelectedOccurrenceIndex = null; openRow.Enabled = definition.Enabled = false;
                ClearTechnicalContext(); RefreshDetails(); SelectedComparisonChanged?.Invoke(null); return;
            }
            if (selection.Technical) { TechnicalMembershipRequested?.Invoke(); return; }
            SelectedSide = selection.Side; SelectedOccurrenceIndex = selection.Occurrence;
            if (selection.Occurrence.HasValue)
            {
                var item = Projection.Occurrences[selection.Occurrence.Value];
                SetRows(item.Original.ComparisonRowIndices, "Selected raw occurrence; all original row mappings");
                if (visibleRows.Count == 1) { SelectGridRow(0); RefreshDetails(); }
                else { SetTechnicalContext(TechnicalRoute.Occurrence, occurrence: selection.Occurrence, descriptor: selection); RefreshDetails(); }
            }
            else { SetRows(selection.Rows ?? new int[0], selection.Description); SetTechnicalContext(TechnicalRoute.Description, description: selection.Description, descriptor: selection); RefreshDetails(); }
        }
        private void SetRows(IEnumerable<int> indices, string context)
        {
            ClearTechnicalContext();
            preparing = true;
            try
            {
                visibleRows = indices.Where(allowedRows.Contains).Distinct().OrderBy(i => i).ToList().AsReadOnly();
                grid.CurrentCell = null; grid.RowCount = visibleRows.Count; selected = null; openRow.Enabled = definition.Enabled = false;
                counts.Text = simplifiedPresentation ? visibleRows.Count + " visible / " + Projection.Rows.Count + " comparison rows. Names/categories may be partial; expanded entries are raw occurrences." : primaryWorkspace ? context + Environment.NewLine +
                    (workspaceScope == ObjectsWorkspaceScope.Compare ? visibleRows.Count + " original comparison rows in Compare scope" :
                     ScopeState() == MembershipSnapshotState.Unavailable ? "Selected side inventory unavailable; counts not established" :
                     workspaceScope + ": " + (workspaceScope == ObjectsWorkspaceScope.Target ? Projection.TargetOccurrenceCount : Projection.SourceOccurrenceCount) +
                     " retained raw occurrences; " + visibleRows.Count + " mapped comparison rows displayed" +
                     (ScopeState() == MembershipSnapshotState.SolutionAbsent ? "; selected solution absent" : "")) :
                    context + Environment.NewLine + visibleRows.Count + " visible comparison rows / " + Projection.Rows.Count +
                    " total; Source " + Projection.SourceOccurrenceCount + ", Target " + Projection.TargetOccurrenceCount + " raw occurrences";
            }
            finally { preparing = false; }
            UpdateColumns(); RefreshDetails();
            accountingText = counts.Text;
            counts.AccessibleDescription = counts.Text + "; " + (primaryWorkspace ?
                workspaceScope == ObjectsWorkspaceScope.Compare ? "distinct candidate original row indices; partial" : "side-local raw candidate occurrences, not distinct maker objects; partial" :
                "complete technical comparison with original row and raw occurrence units");
            countToolTip.SetToolTip(counts, counts.AccessibleDescription);
            SelectedComparisonChanged?.Invoke(null);
        }
        protected override void Dispose(bool disposing)
        {
            if (disposing) { ClearTechnicalContext(); countToolTip.Dispose(); }
            base.Dispose(disposing);
        }
        private void SelectGridRow(int position)
        {
            if (makerNavigation) { SelectObjectGridRow(position); return; }
            if (preparing || position < 0 || position >= visibleRows.Count) return;
            var row = Projection.Rows[visibleRows[position]]; selected = row.Original;
            if (SelectedOccurrenceIndex.HasValue && !Projection.Occurrences[SelectedOccurrenceIndex.Value].Original.ComparisonRowIndices.Contains(row.OriginalRowIndex)) SelectedOccurrenceIndex = null;
            SetTechnicalContext(TechnicalRoute.Row, selected, SelectedOccurrenceIndex);
            RefreshDetails(); openRow.Enabled = true; definition.Enabled = selected.AuthoritativeRow.DefinitionDetail != null;
            SelectedComparisonChanged?.Invoke(selected);
        }
        private IReadOnlyList<SolutionObjectsExplorerCategory> WorkspaceNavigationCategories(bool all)
        {
            var filtered = Projection.WorkspaceCategories(workspaceScope, allowedRows, all && string.IsNullOrWhiteSpace(searchText)).ToDictionary(c => c.Key);
            return Projection.WorkspaceCategories(workspaceScope, Enumerable.Range(0, Projection.Rows.Count), true)
                .Select(c => filtered.TryGetValue(c.Key, out var visible) ? visible :
                    new SolutionObjectsExplorerCategory(c.Key, c.Name, new int[0], new int[0], new int[0])).ToList().AsReadOnly();
        }
        private static IEnumerable<TreeNode> MaterializedNodes(TreeNodeCollection nodes)
        {
            foreach (TreeNode node in nodes)
            {
                if (string.IsNullOrEmpty(node.Name)) continue;
                yield return node;
                foreach (var child in MaterializedNodes(node.Nodes)) yield return child;
            }
        }
        private static string NavigationPath(TreeNode node)
        {
            if (node == null || string.IsNullOrEmpty(node.Name)) return null;
            var names = new Stack<string>();
            while (node != null) { names.Push(node.Name); node = node.Parent; }
            return string.Join("\u001f", names);
        }
        private TreeNode ResolveNavigationPath(string path)
        {
            if (string.IsNullOrEmpty(path)) return null;
            TreeNode node = null; var nodes = navigation.Nodes;
            var keys = path.Split('\u001f');
            for (int i = 0; i < keys.Length; i++)
            {
                node = nodes.Cast<TreeNode>().FirstOrDefault(n => n.Name == keys[i]);
                if (node == null) return null;
                if (i < keys.Length - 1) { node.Expand(); Populate(node); nodes = node.Nodes; }
            }
            return node;
        }
        internal NavigationBookmark CaptureNavigation()
        {
            int top = grid.RowCount == 0 ? -1 : grid.FirstDisplayedScrollingRowIndex;
            return new NavigationBookmark { Owner = Projection, Primary = primaryWorkspace, Scope = workspaceScope,
                ObjectKey = makerNavigation && selectedObject != null ?
                    (selectedObject.Locals.FirstOrDefault(l => l.Side == SelectedSide) ?? selectedObject.Locals[0]).Key : null,
                SelectedPath = NavigationPath(navigation.SelectedNode), TopPath = NavigationPath(navigation.TopNode),
                ExpandedPaths = MaterializedNodes(navigation.Nodes).Where(n => n.IsExpanded).Select(NavigationPath).ToArray(),
                Row = selected?.OriginalRowIndex, Side = SelectedSide, Occurrence = SelectedOccurrenceIndex,
                TopRow = top >= 0 && top < visibleRows.Count ? (int?)visibleRows[top] : null,
                HorizontalOffset = grid.HorizontalScrollingOffset, Focus = lastNavigationFocus, Support = supportToggle.Checked };
        }
        internal void RestoreNavigation(NavigationBookmark bookmark, int? row, DisplayEnvironmentSide? side, int? occurrence)
        {
            if (makerNavigation) { RestoreObjectNavigation(bookmark, row, side); return; }
            if (bookmark == null || !ReferenceEquals(bookmark.Owner, Projection) || bookmark.Primary != primaryWorkspace ||
                primaryWorkspace && bookmark.Scope != workspaceScope || IsDisposed || Disposing) return;
            restoringNavigation = true;
            try
            {
                foreach (string path in bookmark.ExpandedPaths ?? new string[0]) ResolveNavigationPath(path)?.Expand();
                var node = ResolveNavigationPath(bookmark.SelectedPath);
                if (node != null) { navigation.SelectedNode = node; SelectNavigation(node); }
                if (row.HasValue)
                {
                    if (occurrence.HasValue && occurrence.Value >= 0 && occurrence.Value < Projection.Occurrences.Count && includedOccurrences.Contains(occurrence.Value) &&
                        Projection.Occurrences[occurrence.Value].Original.ComparisonRowIndices.Contains(row.Value) &&
                        (!primaryWorkspace || !ScopeSide().HasValue || Projection.Occurrences[occurrence.Value].Original.Side == ScopeSide())) SelectOccurrence(occurrence.Value);
                    SelectComparison(row.Value, side);
                    // Preserve exact inactive context without activating a route outside the current list.
                    bool exactOutsideRow = primaryWorkspace && Projection != null && ReferenceEquals(bookmark.Owner, Projection) &&
                        row.Value >= 0 && row.Value < Projection.Rows.Count &&
                        ReferenceEquals(selected, Projection.Rows[row.Value].Original) && SelectedSide == side &&
                        (!allowedRows.Contains(row.Value) || !InScope(row.Value) || !visibleRows.Contains(row.Value));
                    if (exactOutsideRow && side.HasValue && (side.Value == DisplayEnvironmentSide.Source || side.Value == DisplayEnvironmentSide.Target) &&
                        occurrence.HasValue && occurrence.Value >= 0 && occurrence.Value < Projection.Occurrences.Count &&
                        Projection.Occurrences[occurrence.Value].Original.Side == side.Value &&
                        Projection.Occurrences[occurrence.Value].Original.ComparisonRowIndices.Contains(row.Value))
                        SelectedOccurrenceIndex = occurrence.Value;
                }
                else { selected = null; openRow.Enabled = definition.Enabled = false; }
                var top = ResolveNavigationPath(bookmark.TopPath); if (top != null) navigation.TopNode = top;
                RestoreGridScroll(bookmark); restoredNavigation = bookmark;
                lastNavigationFocus = bookmark.Focus ?? "tree";
                if (primaryWorkspace) supportToggle.Checked = bookmark.Support;
            }
            finally { restoringNavigation = false; }
        }
        internal void ResetNavigation(bool support = false)
        {
            if (makerNavigation)
            {
                bool previous = restoringNavigation; restoringNavigation = true;
                try
                {
                    var all = navigation.Nodes.Cast<TreeNode>().FirstOrDefault(n => n.Name == "workspace-all");
                    if (all != null) { navigation.SelectedNode = all; SelectObjectNode(all); }
                }
                finally { restoringNavigation = previous; }
                return;
            }
            restoringNavigation = true;
            try
            {
                SelectedOccurrenceIndex = null; SelectedSide = ScopeSide();
                var node = navigation.Nodes.Cast<TreeNode>().FirstOrDefault(n => n.Name == (primaryWorkspace ? "workspace-all" : "all"));
                if (node != null) { navigation.SelectedNode = node; SelectNavigation(node); }
                else SetRows(new int[0], "No retained navigation entries");
                if (primaryWorkspace) supportToggle.Checked = support;
            }
            finally { restoringNavigation = false; }
        }
        internal void WithNavigationRestoration(Action action)
        {
            bool wasRestoring = restoringNavigation;
            restoringNavigation = true;
            try { action(); }
            finally { restoringNavigation = wasRestoring; }
        }
        internal void RestoreNavigationFocus()
        {
            if (!Visible || IsDisposed || Disposing) return;
            bool wasRestoring = restoringNavigation;
            restoringNavigation = true;
            try
            {
                if (lastNavigationFocus == "grid") grid.Focus();
                else if (lastNavigationFocus == "details" && details.Visible) details.Focus();
                else navigation.Focus();
                RestoreGridScroll(restoredNavigation);
            }
            finally { restoringNavigation = wasRestoring; }
        }
        private void RestoreGridScroll(NavigationBookmark bookmark)
        {
            if (bookmark == null || !ReferenceEquals(bookmark.Owner, Projection)) return;
            int topPosition = bookmark.TopRow.HasValue ? visibleRows.ToList().IndexOf(bookmark.TopRow.Value) : -1;
            if (grid.RowCount > 0) grid.FirstDisplayedScrollingRowIndex = Math.Max(0, topPosition);
            int maxOffset = Math.Max(0, grid.Columns.Cast<DataGridViewColumn>().Where(c => c.Visible).Sum(c => c.Width) - grid.ClientSize.Width);
            grid.HorizontalScrollingOffset = Math.Min(bookmark.HorizontalOffset, maxOffset);
        }

        private void ClearTechnicalContext()
        {
            technicalContext = null; technicalContextRevision++; cachedTechnicalRevision = -1;
            supportEvidence = string.Empty; details.Clear(); detailsContentDirty = true;
        }
        private void SetTechnicalContext(TechnicalRoute route, MembershipGridRow row = null, int? occurrence = null,
            string description = null, object descriptor = null)
        {
            var old = technicalContext;
            if (old != null && ReferenceEquals(old.Projection, Projection) && ReferenceEquals(old.Row, row) &&
                old.Side == SelectedSide && old.Occurrence == occurrence && old.Route == route &&
                ReferenceEquals(old.Descriptor, descriptor) && old.Description == description) return;
            ClearTechnicalContext();
            technicalContext = new TechnicalContext { Projection = Projection, Row = row, Side = SelectedSide,
                Occurrence = occurrence, Route = route, Description = description, Descriptor = descriptor, Revision = technicalContextRevision };
        }
        private string TechnicalEvidence()
        {
            var context = technicalContext;
            if (context == null) return string.Empty;
            if (cachedTechnicalRevision == context.Revision) return supportEvidence;
            try
            {
                string value;
                if (context.Route == TechnicalRoute.Description) value = context.Description ?? string.Empty;
                else if (context.Route == TechnicalRoute.Occurrence)
                    value = OccurrenceEvidence(context.Projection.Occurrences[context.Occurrence.Value]) + Environment.NewLine +
                        "Choose a comparison row explicitly; no arbitrary first mapping selected.";
                else if (context.Route == TechnicalRoute.RowOnly) value = describe(context.Row);
                else
                {
                    var row = context.Projection.Rows[context.Row.OriginalRowIndex];
                    value = "Original comparison row: " + row.OriginalRowIndex + Environment.NewLine +
                        "Source raw occurrence mappings: " + string.Join(",", row.SourceOccurrences) + Environment.NewLine +
                        "Target raw occurrence mappings: " + string.Join(",", row.TargetOccurrences) + Environment.NewLine + describe(context.Row);
                    if (context.Occurrence.HasValue) value = OccurrenceEvidence(context.Projection.Occurrences[context.Occurrence.Value]) + Environment.NewLine + value;
                    var definitionDetail = context.Row.AuthoritativeRow.DefinitionDetail;
                    if (definitionDetail != null)
                    {
                        value += Environment.NewLine + "Definition diagnostic: " + definitionDetail.Diagnostic;
                        foreach (var property in definitionDetail.Properties) value += Environment.NewLine + property.PropertyName + ": Source " + property.SourceValue + "; Target " + property.TargetValue + "; " + property.Comparison;
                    }
                }
                // A callback cannot publish text for a context replaced during that callback.
                if (!ReferenceEquals(context, technicalContext)) return string.Empty;
                supportEvidence = value; cachedTechnicalRevision = context.Revision;
                return supportEvidence;
            }
            catch { ClearTechnicalContext(); throw; }
        }
        protected override void OnVisibleChanged(EventArgs e)
        {
            base.OnVisibleChanged(e);
            if (evidenceSplit != null && Visible && detailsContentDirty && (!primaryWorkspace || detailsRequested) && !evidenceSplit.Panel2Collapsed)
                Guard(RenderDetailsContent);
        }
        private static string OccurrenceEvidence(SolutionObjectsExplorerOccurrence item)
        {
            var o = item.Original; var r = o.Record; var text = new StringBuilder();
            text.AppendLine(o.Side + ": " + o.Snapshot.Environment.DisplayName).AppendLine("Solution: " + o.Snapshot.SolutionUniqueName);
            text.AppendLine("Original raw index: " + o.OriginalRawIndex).AppendLine("H1 occurrence index: " + o.OccurrenceIndex);
            text.AppendLine("SolutionComponentId: " + r.SolutionComponentId).AppendLine("ObjectId (local audit): " + r.ObjectId);
            text.AppendLine("ComponentType: " + r.ComponentType).AppendLine("RootSolutionComponentId: " + r.RootSolutionComponentId);
            text.AppendLine("RootComponentBehavior: " + r.RootComponentBehavior).AppendLine("IsMetadata: " + r.IsMetadata);
            text.AppendLine("Identity resolution: " + o.Identity.Status).AppendLine("Portable identity: " + o.Identity.ComparisonKey);
            text.AppendLine("Placement: " + o.Placement + "; multiple parents: " + o.HasDifferentParentAssociations + "; not top-level proof");
            foreach (var a in o.Associations) text.AppendLine(a.Kind + ": " + a.Status + "; parent occurrence=" + a.ParentOccurrenceIndex + "; provenance=" + a.Provenance + "; " + a.Diagnostic);
            text.AppendLine("Original identity diagnostic: " + o.Identity.Diagnostic);
            return text.ToString();
        }
        private static string Value(SolutionObjectsExplorerRow row, string column)
        {
            switch (column)
            {
                case "DisplayName": return row.DisplayName; case "Name": return row.Name; case "Type": return row.Type;
                case "Source": return row.Source; case "Target": return row.Target; case "ComparisonStatus": return row.ComparisonStatus;
                default: return row.ObjectField(column);
            }
        }
        private void Guard(Action action)
        { try { action(); } catch (Exception error) { selected = null; ClearTechnicalContext(); RenderingFailed?.Invoke(error); } }
        private sealed class ExplorerAccessibleRow : DataGridViewRow
        {
            protected override AccessibleObject CreateAccessibilityInstance() => new ExplorerAccessibleObject(this);
            private sealed class ExplorerAccessibleObject : DataGridViewRowAccessibleObject
            {
                private readonly DataGridViewRow row;
                internal ExplorerAccessibleObject(DataGridViewRow row) : base(row) { this.row = row; }
                public override Rectangle Bounds => row.DataGridView == null || !row.Displayed || row.Index < 0 || !row.DataGridView.IsHandleCreated
                    ? Rectangle.Empty : row.DataGridView.RectangleToScreen(row.DataGridView.GetRowDisplayRectangle(row.Index, true));
                public override string Description
                {
                    get
                    {
                        var owner = row.Tag as SolutionObjectsExplorerControl;
                        return owner == null || row.Index < 0 || row.Index >= owner.visibleRows.Count ? base.Description :
                            owner.primaryWorkspace ? string.Join(Environment.NewLine, new[] { "DisplayName", "Name", "Type", "ComparisonStatus" }.Select(c => owner.FullCellToolTip(owner.visibleRows[row.Index], c))) : "Original comparison row " + owner.visibleRows[row.Index] + "; " + owner.Projection.Rows[owner.visibleRows[row.Index]].ComparisonStatus;
                    }
                }
            }
        }
    }
}
