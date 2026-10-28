using System;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Collections.Generic;
using System.Diagnostics;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using D365SolutionComparer.Models.Membership;
using D365SolutionComparer.Models.Presentation;
using D365SolutionComparer.Services.Membership;

namespace D365SolutionComparer
{
    internal sealed class MembershipResultsForm : Form
    {
        private SimplifiedSolutionExplorerPanel simplifiedExplorer;
        private readonly MembershipComparisonPresentation presentation;
        private readonly CompletedMembershipComparison completedComparison;
        internal CompletedMembershipComparison CompletedComparison => completedComparison;
        private readonly string sourceSolutionVersion;
        private readonly string targetSolutionVersion;
        private readonly DataGridView resultsGrid;
        private readonly Button displayEvidenceButton;
        private readonly Action captureAppSettingEvidence;
        private readonly Microsoft.Xrm.Sdk.IOrganizationService sourceService;
        private readonly Microsoft.Xrm.Sdk.IOrganizationService targetService;
        private readonly IReadOnlyList<MembershipGridRow> gridRows;
        private readonly ToolStripComboBox viewSelector;
        private readonly ToolStripComboBox membershipFilter;
        private readonly ToolStripLabel legacyViewLabel;
        private readonly ToolStripLabel viewStatus;
        private MembershipHierarchyControl hierarchyControl;
        private MembershipHierarchyView currentHierarchyModel;
        private readonly ToolStripButton explorerButton;
        private SolutionObjectsExplorerControl explorerControl;
        private Task<SolutionObjectsExplorerProjection> explorerTask;
        private DisplayEnvironmentSide? retainedExplorerSide;
        private int? retainedExplorerOccurrence;
        private readonly ToolStripTextBox objectsSearch;
        private bool showExplorerMetadata;
        private bool technicalWorkspace, switchingWorkspace;
        private int lastTechnicalView;
        private readonly ToolStrip primaryBar, technicalBar;
        private readonly ToolStripButton workspaceObjectsAction, objectsMetadata;
        private readonly Label objectsNotice;
        private readonly ToolStripComboBox objectsScope;
        private readonly ToolStripButton objectsAction, technicalAction, objectsDifferences, objectsDetails, originalRowAction;
        private readonly Panel headerPanel;
        private readonly Label summaryLabel, diagnosticLabel;
        private readonly ObjectsSnapshotHeader objectsHeader;
        private readonly Panel navigationPanel;
        private readonly ToolStripButton advancedDiagnosticsAction;
        private readonly ToolStripButton sharedAdvancedDiagnosticsAction;
        private bool arrangingHeader;
        private readonly ToolStripComboBox objectsMembershipFilter;
        private readonly Dictionary<string, SolutionObjectsExplorerControl.NavigationBookmark> explorerBookmarks = new Dictionary<string, SolutionObjectsExplorerControl.NavigationBookmark>();
        private MembershipHierarchyControl.NavigationBookmark hierarchyBookmark;
        private bool navigationTransaction, restoringNavigation, resetNavigationPending, resetSupportChoice, restoreFocusPending;
        private bool focusWorkspace;
        private long selectionIntentRevision, hierarchyBookmarkRevision;
        private bool retainContextWithoutRow;
        private int ActiveExplorerFilter => technicalWorkspace ? membershipFilter.SelectedIndex : objectsMembershipFilter.SelectedIndex;
        private bool ExplorerFilterAll => technicalWorkspace ? membershipFilter.SelectedIndex == 0 && viewSelector.SelectedIndex != 2 :
            objectsMembershipFilter.SelectedIndex == 0 && (!objectsDifferences.Checked || CurrentObjectsScope != ObjectsWorkspaceScope.Compare);
        internal bool IsTechnicalWorkspace => technicalWorkspace;
        private ObjectsWorkspaceScope CurrentObjectsScope => (ObjectsWorkspaceScope)objectsScope.SelectedIndex;
        private bool initialViewQueued;
        private bool initialViewStarted;
        private bool shownCallbacksCompleted;
        private readonly Label viewPreparation;
        private int boundExplorerFilter = -1;
        private bool boundExplorerDifferences;
        internal Func<ComponentHierarchyProjection, CancellationToken, SolutionObjectsExplorerProjection> ProjectExplorer { get; set; }
        internal SolutionObjectsExplorerProjection ExplorerProjection { get; private set; }
        private readonly CancellationTokenSource hierarchyLifetime = new CancellationTokenSource();
        private Task<ComponentHierarchyProjection> hierarchyTask;
        private int viewRequestVersion;
        private int boundHierarchyFilter = -1;
        private bool hierarchyLifetimeDisposed;
        private int? retainedSelectionRow;
        internal Func<CancellationToken, ComponentHierarchyProjection> ProjectHierarchy { get; set; }
        internal ComponentHierarchyProjection HierarchyProjection { get; private set; }
        internal List<string> HierarchyDiagnostics { get; } = new List<string>();
#if DEBUG
        internal Action CaptureType92Evidence { get; set; }
        internal Action CaptureType59Evidence { get; set; }
        internal Action DiscoverType59Evidence { get; set; }
        internal Action CaptureCloudFlowEvidence { get; set; }
        internal Action CaptureSavedQueryEvidence { get; set; }
        internal Action CaptureUnsupportedCoverageInventory { get; set; }
        internal Action CaptureType36Evidence { get; set; }
        internal Action CaptureType31Evidence { get; set; }
        internal Action CaptureType10072Evidence { get; set; }
        internal Action CaptureType300Evidence { get; set; }
        internal Action CaptureType74Evidence { get; set; }
        internal Action CaptureType511Evidence { get; set; }
        internal Action CaptureType10276Evidence { get; set; }
        internal Action CaptureAppActionEvidence { get; set; }
#endif

        public MembershipResultsForm(CompletedMembershipComparison completedComparison,
            string sourceSolutionVersion = null, string targetSolutionVersion = null,
            Action captureAppSettingEvidence = null,
            Microsoft.Xrm.Sdk.IOrganizationService sourceService = null,
            Microsoft.Xrm.Sdk.IOrganizationService targetService = null)
            : this((completedComparison ?? throw new ArgumentNullException(nameof(completedComparison))).Presentation,
                completedComparison, sourceSolutionVersion, targetSolutionVersion, captureAppSettingEvidence, sourceService, targetService)
        {
        }

        public MembershipResultsForm(MembershipComparisonPresentation presentation,
            string sourceSolutionVersion = null, string targetSolutionVersion = null,
            Action captureAppSettingEvidence = null,
            Microsoft.Xrm.Sdk.IOrganizationService sourceService = null,
            Microsoft.Xrm.Sdk.IOrganizationService targetService = null)
            : this(presentation, null, sourceSolutionVersion, targetSolutionVersion,
                captureAppSettingEvidence, sourceService, targetService)
        {
        }

        private MembershipResultsForm(MembershipComparisonPresentation presentation,
            CompletedMembershipComparison completedComparison, string sourceSolutionVersion,
            string targetSolutionVersion, Action captureAppSettingEvidence,
            Microsoft.Xrm.Sdk.IOrganizationService sourceService,
            Microsoft.Xrm.Sdk.IOrganizationService targetService)
        {
            this.presentation = presentation ?? throw new ArgumentNullException(nameof(presentation));
            this.completedComparison = completedComparison;
            this.sourceSolutionVersion = sourceSolutionVersion ?? string.Empty;
            this.targetSolutionVersion = targetSolutionVersion ?? string.Empty;
            this.captureAppSettingEvidence = captureAppSettingEvidence;
            this.sourceService = sourceService;
            this.targetService = targetService;
            Text = "Solution Membership Compare - " + presentation.SolutionUniqueName;
            StartPosition = FormStartPosition.CenterParent;
            MinimumSize = new Size(900, 520);
            Size = new Size(1400, 760);
            Font = new Font("Segoe UI", 9F);
            SuspendLayout();

            var header = new Panel { Dock = DockStyle.Top, Height = 174, Padding = new Padding(10), BackColor = Color.White };
            headerPanel = header;
            var title = new Label
            {
                Dock = DockStyle.Top,
                Height = 28,
                Font = new Font("Segoe UI", 11F, FontStyle.Bold),
                Text = "Solution: " + presentation.SolutionUniqueName
            };
            var states = new Label
            {
                Dock = DockStyle.Top,
                Height = 24,
                AutoEllipsis = true,
                Text = "Source (" + presentation.Source.Diagnostics.EnvironmentName + "): " +
                    DisplayState(presentation.Source) + "    |    Target (" +
                    presentation.Target.Diagnostics.EnvironmentName + "): " + DisplayState(presentation.Target)
            };
            var summary = new Label
            {
                Dock = DockStyle.Top,
                Height = 44,
                Name = "MembershipSummary",
                AccessibleName = "Membership comparison row summary",
                AccessibleDescription = "Each comparison row is counted once. Indeterminate resolution categories use Ambiguous, Unresolved, Unsupported, then Resolved precedence.",
                Font = new Font("Segoe UI", 9F, FontStyle.Bold),
                Text = FormatSummary(presentation.Summary)
            };
            var diagnostics = new Label
            {
                Dock = DockStyle.Fill,
                AutoEllipsis = true,
                ForeColor = Color.DimGray,
                Text = FormatDiagnostics("Source", presentation.Source.Diagnostics) + Environment.NewLine +
                    FormatDiagnostics("Target", presentation.Target.Diagnostics)
            };
            summaryLabel = summary; diagnosticLabel = diagnostics;
            var coverageDetails = new LinkLabel
            {
                Dock = DockStyle.Right,
                Width = 180,
                Height = 24,
                Text = "Coverage Details...",
                TextAlign = ContentAlignment.MiddleRight,
                LinkBehavior = LinkBehavior.HoverUnderline
            };
            coverageDetails.LinkClicked += CoverageDetails_LinkClicked;
            displayEvidenceButton = new Button
            {
                Text = "Display &evidence...",
                Name = "DisplayEvidenceAction",
                AccessibleName = "Display evidence...",
                AccessibleDescription = "Read-only Source and Target display values, availability and provenance for the selected comparison row.",
                Dock = DockStyle.Left,
                Width = 155,
                Enabled = false
            };
            displayEvidenceButton.Click += (sender, args) =>
            {
                var row = SelectedComparisonRow();
                if (row == null) return;
                using (var form = CreateDisplayEvidenceDetails(row)) form.ShowDialog(this);
            };
            var navigation = new Panel { Dock = DockStyle.Bottom, Height = 24 };
            navigationPanel = navigation;
            var legend = new Label
            {
                Name = "MembershipLegend",
                Text = "Green: Shared | Rose: One-sided | Yellow: Indeterminate",
                AccessibleName = "Membership comparison color legend",
                AccessibleDescription = "Comparison statuses are available in each row's accessible description, tooltip and Display evidence details.",
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleCenter,
                ForeColor = Color.DimGray,
                AutoEllipsis = true
            };
            navigation.Controls.Add(legend);
            navigation.Controls.Add(coverageDetails);
            navigation.Controls.Add(displayEvidenceButton);
            var metadataToggle = new CheckBox
            {
                Name = "OptionalMetadataToggle", Text = "&Metadata", Dock = DockStyle.Left, Width = 95,
                AccessibleName = "Show optional metadata columns",
                AccessibleDescription = "Show Managed, Customized, Last Modified, Owner and Status. Retrieval limitations remain available in Display evidence."
            };
            metadataToggle.CheckedChanged += (sender, args) =>
            {
                showExplorerMetadata = metadataToggle.Checked;
                if (objectsMetadata != null) objectsMetadata.Checked = showExplorerMetadata;
                explorerControl?.ShowMetadata(showExplorerMetadata);
                resultsGrid.SuspendLayout();
                foreach (var name in new[] { "ManagedText", "CustomizedText", "LastModifiedText", "OwnerText", "ObjectStatusText" })
                    resultsGrid.Columns[name].Visible = metadataToggle.Checked;
                resultsGrid.ResumeLayout();
            };
            navigation.Controls.Add(metadataToggle);
            header.Controls.Add(diagnostics);
            header.Controls.Add(navigation);
            header.Controls.Add(summary);
            header.Controls.Add(states);
            header.Controls.Add(title);
            objectsHeader = new ObjectsSnapshotHeader(title.Text,
                "Source (" + presentation.Source.Diagnostics.EnvironmentName + "): " + DisplayState(presentation.Source),
                "Target (" + presentation.Target.Diagnostics.EnvironmentName + "): " + DisplayState(presentation.Target),
                AdvancedDiagnosticsText,
                "Snapshot state: Source " + presentation.Source.State + " | Target " + presentation.Target.State,
                FormatCompactSummary(presentation.Summary)) { Dock = DockStyle.Top, TabIndex = 0 };
            objectsHeader.ExpandedChanged += (sender, args) =>
            {
                if (advancedDiagnosticsAction != null) advancedDiagnosticsAction.Checked = objectsHeader.Expanded;
                if (sharedAdvancedDiagnosticsAction != null) sharedAdvancedDiagnosticsAction.Checked = objectsHeader.Expanded;
                LayoutWorkspaceHeader();
            };
            header.Controls.Add(objectsHeader);

            resultsGrid = new DataGridView
            {
                Dock = DockStyle.Fill,
                ReadOnly = true,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                AutoGenerateColumns = false,
                AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None,
                ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize,
                BackgroundColor = Color.White,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                MultiSelect = false,
                RowHeadersVisible = false,
                RowTemplate = new MembershipAccessibleRow()
            };
            resultsGrid.RowTemplate.Height = Math.Max(22, Font.Height * 2 + 8);
            AddColumn("FriendlyNameText", "Component", 100, 300, binding: "ComponentText");
            AddColumn("ComponentCategoryText", "Type", 0, 130, false, "TypeText");
            AddColumn("MembershipStatus", "Membership Status", 0, 175, false);
            AddColumn("DefinitionStatus", "Definition Status", 0, 190, false, "DefinitionStatusText");
            AddColumn("ManagedText", "Managed", 0, 110, false);
            AddColumn("CustomizedText", "Customized", 0, 140, false);
            AddColumn("LastModifiedText", "Last Modified", 0, 200, false);
            AddColumn("OwnerText", "Owner", 0, 200, false);
            AddColumn("ObjectStatusText", "Status", 0, 170, false);
            foreach (DataGridViewColumn column in resultsGrid.Columns)
                if (column.Index >= 4) column.Visible = false;
            resultsGrid.Columns["MembershipStatus"].ToolTipText = "Authoritative Source/Target membership comparison result.";
            resultsGrid.Columns["DefinitionStatus"].ToolTipText = MembershipGridRow.DefinitionMeaning;
            resultsGrid.Columns["ObjectStatusText"].ToolTipText = "Component object operational status when independently verified; not Membership Status or Definition Status.";
            resultsGrid.CellFormatting += ResultsGrid_CellFormatting;
            resultsGrid.CellDoubleClick += ResultsGrid_CellDoubleClick;
            resultsGrid.CellToolTipTextNeeded += ResultsGrid_CellToolTipTextNeeded;
            resultsGrid.SelectionChanged += (sender, args) =>
                displayEvidenceButton.Enabled = SelectedComparisonRow() != null;
            gridRows = MembershipGridRow.CreateRows(presentation, completedComparison, refinedPresentation: true);
            resultsGrid.DataSource = gridRows.ToList();

            var workspaceBar = new ToolStrip { Dock = DockStyle.Top, GripStyle = ToolStripGripStyle.Hidden,
                Name = "MembershipWorkspaceBar", LayoutStyle = ToolStripLayoutStyle.Flow, AccessibleName = "Membership workspaces" };
            ((FlowLayoutSettings)workspaceBar.LayoutSettings).WrapContents = true;
            var viewBar = new ToolStrip { Dock = DockStyle.Top, GripStyle = ToolStripGripStyle.Hidden, Name = "MembershipViewBar",
                LayoutStyle = ToolStripLayoutStyle.Flow, AccessibleName = "Technical Membership views and filters" };
            ((FlowLayoutSettings)viewBar.LayoutSettings).WrapContents = true;
            technicalBar = viewBar;
            primaryBar = new ToolStrip { Dock = DockStyle.Top, GripStyle = ToolStripGripStyle.Hidden, Name = "ObjectsWorkspaceBar",
                LayoutStyle = ToolStripLayoutStyle.Flow, TabIndex = 1, AccessibleName = "Objects workspace navigation" };
            ((FlowLayoutSettings)primaryBar.LayoutSettings).WrapContents = true;
            objectsAction = new ToolStripButton("&Objects") { Name = "PrimaryObjectsAction", Available = false, Checked = true, Overflow = ToolStripItemOverflow.Never, AccessibleName = "Objects workspace" };
            technicalAction = new ToolStripButton("&Technical Membership") { Name = "TechnicalMembershipAction", Overflow = ToolStripItemOverflow.Never, AccessibleName = "Technical Membership: complete authoritative comparison and support evidence" };
            objectsScope = new ToolStripComboBox { Name = "ObjectsScopeSelector", DropDownStyle = ComboBoxStyle.DropDownList,
                AutoSize = false, Width = 225, AccessibleName = "Objects scope: Target, Source or Compare", Overflow = ToolStripItemOverflow.Never };
            objectsScope.Items.AddRange(new object[] { "Target (" + presentation.Target.Diagnostics.EnvironmentName + ")",
                "Source (" + presentation.Source.Diagnostics.EnvironmentName + ")", "Compare" }); objectsScope.SelectedIndex = 0;
            objectsDifferences = new ToolStripButton("Differences") { Name = "ObjectsDifferencesFilter", CheckOnClick = true,
                AccessibleName = "Filter original comparison differences", Overflow = ToolStripItemOverflow.Never };
            objectsDetails = new ToolStripButton("Details") { Name = "ObjectsDetailsAction", CheckOnClick = true,
                AccessibleName = "Show selected object details", Overflow = ToolStripItemOverflow.Never };
            originalRowAction = new ToolStripButton("Original row") { Name = "ObjectsOriginalRowAction", Enabled = false, Overflow = ToolStripItemOverflow.Never };
            objectsSearch = new ToolStripTextBox { Name = "ObjectsSearchFilter", AutoSize = false, Width = 170,
                AccessibleName = "Search retained object names", ToolTipText = "Local search of retained display name, name and type; no requests", Overflow = ToolStripItemOverflow.Never };
            objectsSearch.TextChanged += (sender, args) =>
            {
                if (navigationTransaction || technicalWorkspace) return;
                CaptureNavigation(); RetainViewSelection();
                explorerControl?.SetSearch(objectsSearch.Text); RestoreExplorerNavigation();
            };
            objectsMetadata = new ToolStripButton("Metadata") { Name = "ObjectsMetadataAction", CheckOnClick = true,
                AccessibleName = "Show retained optional metadata columns", Overflow = ToolStripItemOverflow.Never };
            objectsMetadata.CheckedChanged += (sender, args) => metadataToggle.Checked = objectsMetadata.Checked;
            var clearSearch = new ToolStripButton("Clear search") { Name = "ObjectsClearSearchAction",
                AccessibleName = "Clear retained object name search", Overflow = ToolStripItemOverflow.Never };
            clearSearch.Click += (sender, args) => objectsSearch.Clear();
            objectsMembershipFilter = new ToolStripComboBox { Name = "ObjectsMembershipPresenceFilter", DropDownStyle = ComboBoxStyle.DropDownList,
                AutoSize = false, Width = 165, Overflow = ToolStripItemOverflow.Never, AccessibleName = "Filter Objects membership independently from Technical Membership" };
            objectsMembershipFilter.Items.AddRange(new object[] { "All", "Present in Both", "Source Only", "Target Only", "Indeterminate" });
            objectsMembershipFilter.SelectedIndex = 0;
            objectsMembershipFilter.SelectedIndexChanged += async (sender, args) =>
            {
                if (navigationTransaction || technicalWorkspace) return;
                CaptureNavigation(); RetainViewSelection(); if (CanStartLocalView) await ShowExplorerAsync();
            };
            primaryBar.Items.AddRange(new ToolStripItem[] { objectsAction, new ToolStripLabel("Scope:"), objectsScope,
                new ToolStripLabel("Search:"), objectsSearch, clearSearch, new ToolStripLabel("Membership:"), objectsMembershipFilter,
                objectsDifferences, objectsDetails, originalRowAction, objectsMetadata });
            objectsAction.Click += async (sender, args) => { RequestNavigationFocus(false); await SelectWorkspaceAsync(false); };
            technicalAction.Click += async (sender, args) => { RequestNavigationFocus(true); await SelectWorkspaceAsync(true); };
            objectsScope.SelectedIndexChanged += async (sender, args) =>
            { if (navigationTransaction) return; CaptureNavigation(); RetainViewSelection(); objectsScope.ToolTipText = objectsScope.SelectedItem.ToString(); UpdateWorkspaceChrome(); if (CanStartLocalView) await ShowExplorerAsync(); };
            objectsDifferences.CheckedChanged += async (sender, args) =>
            { if (navigationTransaction) return; CaptureNavigation(); RetainViewSelection(); if (!technicalWorkspace && CanStartLocalView) await ShowExplorerAsync(); };
            objectsDetails.CheckedChanged += (sender, args) => explorerControl?.ShowDetails(objectsDetails.Checked);
            originalRowAction.Click += (sender, args) => { var row = SelectedComparisonRow(); if (row != null) NavigateToOriginalRow(row); };
            workspaceObjectsAction = new ToolStripButton("&Solution Objects") { Name = "ObjectsBackAction", Checked = true,
                Overflow = ToolStripItemOverflow.Never, AccessibleName = "Solution Objects: partial evidence-backed object view" };
            workspaceObjectsAction.Click += async (sender, args) => { RequestNavigationFocus(false); await SelectWorkspaceAsync(false); };
            sharedAdvancedDiagnosticsAction = new ToolStripButton("Advanced diagnostics") { Name = "SharedAdvancedDiagnosticsAction",
                CheckOnClick = true, Overflow = ToolStripItemOverflow.Never, AccessibleName = "Show complete retained Advanced diagnostics" };
            sharedAdvancedDiagnosticsAction.Click += (sender, args) => objectsHeader.Expanded = sharedAdvancedDiagnosticsAction.Checked;
            var sharedCoverageAction = new ToolStripButton("Coverage Details") { Name = "SharedCoverageDetailsAction",
                Overflow = ToolStripItemOverflow.Never, AccessibleName = "Open retained Coverage Details without capture" };
            sharedCoverageAction.Click += (sender, args) => CoverageDetails_LinkClicked(sender, null);
            var resetFiltersAction = new ToolStripButton("Reset Filters") { Name = "MembershipResetFiltersAction",
                Overflow = ToolStripItemOverflow.Never, AccessibleName = "Reset filters in the active workspace; preserve scope and snapshot" };
            resetFiltersAction.Click += async (sender, args) => await ResetNavigationFiltersAsync();
            workspaceBar.Items.AddRange(new ToolStripItem[] { workspaceObjectsAction, technicalAction, sharedAdvancedDiagnosticsAction, sharedCoverageAction, resetFiltersAction });
            objectsNotice = new Label { Name = "ObjectsPartialNotice", Dock = DockStyle.Top, Padding = new Padding(10, 3, 10, 3),
                Text = "Objects — partial view. Additional components are available in Technical Membership.",
                AccessibleName = "Objects are a partial projection; Technical Membership retains complete comparison rows and raw occurrences",
                ForeColor = Color.DimGray, BackColor = Color.White };
            viewSelector = new ToolStripComboBox { Name = "MembershipViewSelector", DropDownStyle = ComboBoxStyle.DropDownList,
                AutoSize = false, Width = 195, AccessibleName = "Membership presentation view" };
            viewSelector.Items.AddRange(new object[] { "Flat View", "Hierarchy View", "Differences Only", "Solution Objects Explorer" }); viewSelector.SelectedIndex = 3;
            membershipFilter = new ToolStripComboBox { Name = "MembershipPresenceFilter", DropDownStyle = ComboBoxStyle.DropDownList,
                AutoSize = false, Width = 165, AccessibleName = "Filter authoritative membership rows" };
            membershipFilter.Items.AddRange(new object[] { "All", "Present in Both", "Source Only", "Target Only", "Indeterminate" });
            membershipFilter.SelectedIndex = 0;
            viewStatus = new ToolStripLabel { Name = "MembershipViewStatus", Text = "Preparing local object navigation...", AutoSize = false, Width = 155, AutoToolTip = true };
            legacyViewLabel = new ToolStripLabel("View:");
            viewBar.Items.Add(legacyViewLabel); viewBar.Items.Add(viewSelector);
            viewBar.Items.Add(new ToolStripLabel("Membership:")); viewBar.Items.Add(membershipFilter); viewBar.Items.Add(viewStatus);
            explorerButton = new ToolStripButton("Solution Objects Explorer") { Name = "SolutionObjectsExplorerAction",
                CheckOnClick = false, Checked = true, AccessibleName = "Open read-only Solution Objects Explorer" };
            explorerButton.Click += (sender, args) => viewSelector.SelectedIndex = 3;
            viewBar.Items.Add(explorerButton);
            advancedDiagnosticsAction = new ToolStripButton("Advanced diagnostics") { Name = "AdvancedDiagnosticsAction",
                CheckOnClick = true, AccessibleName = "Show complete advanced diagnostics" };
            advancedDiagnosticsAction.Click += (sender, args) => objectsHeader.Expanded = advancedDiagnosticsAction.Checked;
            viewBar.Items.Add(advancedDiagnosticsAction);
            ProjectHierarchy = token => completedComparison == null ? ComponentHierarchyProjector.Project(presentation, cancellationToken: token) :
                ComponentHierarchyProjector.Project(completedComparison, completedComparison.SourceHierarchyParents,
                    completedComparison.TargetHierarchyParents, token);
            ProjectExplorer = (hierarchy, token) => SolutionObjectsExplorerProjector.Project(presentation, completedComparison, hierarchy, token);
            bool differencesOnlyFilter = false;
            resultsGrid.CellClick += (sender, args) =>
            { if (args.RowIndex >= 0) AcceptFlatSelection(resultsGrid.Rows[args.RowIndex].DataBoundItem as MembershipGridRow); };
            resultsGrid.KeyUp += (sender, args) =>
            { AcceptFlatSelection(resultsGrid.CurrentRow?.DataBoundItem as MembershipGridRow); };
            viewSelector.SelectedIndexChanged += async (sender, args) =>
            {
                if (navigationTransaction) return;
                CaptureNavigation(); RetainViewSelection();
                if (!switchingWorkspace) { technicalWorkspace = true; lastTechnicalView = viewSelector.SelectedIndex; }
                UpdateWorkspaceChrome(); explorerButton.Checked = viewSelector.SelectedIndex == 3;
                viewRequestVersion++;
                if (differencesOnlyFilter != (viewSelector.SelectedIndex == 2))
                { differencesOnlyFilter = viewSelector.SelectedIndex == 2; ApplyMembershipFilter(); }
                // A deliberate selection after Show can precede the queued initial view.
                // Before a handle exists, only update presentation; never start local work.
                if (CanStartLocalView) await ShowSelectedViewAsync();
                else if (viewSelector.SelectedIndex == 0 || viewSelector.SelectedIndex == 2) ShowFlatView();
                else ShowPreparation(viewSelector.SelectedIndex == 3 ? "Preparing local object navigation..." : "Preparing local hierarchy...");
            };
            membershipFilter.SelectedIndexChanged += async (sender, args) =>
            {
                if (navigationTransaction) return;
                CaptureNavigation(); RetainViewSelection(); ApplyMembershipFilter();
                if (CanStartLocalView) await ShowSelectedViewAsync();
            };

            viewPreparation = new Label { Name = "MembershipViewPreparation", Dock = DockStyle.Fill,
                Text = "Preparing local object navigation...", TextAlign = ContentAlignment.MiddleCenter,
                AccessibleName = "Preparing local membership presentation", BackColor = Color.White };
            resultsGrid.Visible = false;
            Controls.Add(resultsGrid);
            Controls.Add(viewPreparation);
            Controls.Add(viewBar);
            Controls.Add(primaryBar);
            Controls.Add(objectsNotice);
            Controls.Add(workspaceBar);
            Controls.Add(header);
            UpdateWorkspaceChrome();
            ResumeLayout(true);
        }

        private static bool MatchesPresence(MembershipGridRow row, int filter) => filter == 0 || row.AuthoritativeRow.Presence ==
            (filter == 1 ? MembershipPresence.PresentInBoth : filter == 2 ? MembershipPresence.OnlyInSource :
             filter == 3 ? MembershipPresence.OnlyInTarget : MembershipPresence.Indeterminate);

        private void RequestNavigationFocus(bool technical) { restoreFocusPending = true; focusWorkspace = technical; }
        private string ExplorerBookmarkKey(bool primary, ObjectsWorkspaceScope scope) => primary ? "Objects:" + scope : "Technical:Explorer";
        private void CaptureNavigation()
        {
            if (restoringNavigation || IsDisposed || Disposing) return;
            if (explorerControl?.Visible == true)
                explorerBookmarks[ExplorerBookmarkKey(explorerControl.IsPrimaryWorkspace, explorerControl.WorkspaceScope)] = explorerControl.CaptureNavigation();
            if (hierarchyControl?.Visible == true)
            { hierarchyBookmark = hierarchyControl.CaptureNavigation(); hierarchyBookmarkRevision = selectionIntentRevision; }
        }
        private void ConnectExplorerNavigation()
        {
            explorerControl.NavigationIntentChanged = () =>
            {
                if (restoringNavigation || navigationTransaction || !explorerControl.Visible) return;
                retainedSelectionRow = explorerControl.SelectedComparison?.OriginalRowIndex;
                retainedExplorerSide = explorerControl.SelectedSide;
                retainedExplorerOccurrence = explorerControl.SelectedOccurrenceIndex;
                retainContextWithoutRow = !retainedSelectionRow.HasValue; selectionIntentRevision++;
                CaptureNavigation();
            };
        }
        private void RestoreExplorerNavigation()
        {
            if (explorerControl == null || explorerControl.Projection == null || restoringNavigation) return;
            restoringNavigation = true;
            try
            {
                if (resetNavigationPending)
                {
                    explorerControl.ResetNavigation(resetSupportChoice); resetNavigationPending = false;
                    retainedSelectionRow = retainedExplorerOccurrence = null; retainedExplorerSide = null;
                    return;
                }
                if (explorerBookmarks.TryGetValue(ExplorerBookmarkKey(!technicalWorkspace, CurrentObjectsScope), out var bookmark))
                    explorerControl.RestoreNavigation(bookmark, retainedSelectionRow, retainedExplorerSide, retainedExplorerOccurrence);
            }
            finally { restoringNavigation = false; }
        }
        private void ConnectHierarchyNavigation()
        {
            hierarchyControl.NavigationIntentChanged = () =>
            {
                if (restoringNavigation || navigationTransaction || !hierarchyControl.Visible) return;
                retainedSelectionRow = hierarchyControl.SelectedComparison?.OriginalRowIndex;
                retainedExplorerSide = hierarchyControl.SelectedSide;
                retainedExplorerOccurrence = hierarchyControl.SelectedOccurrenceIndex;
                retainContextWithoutRow = !retainedSelectionRow.HasValue; selectionIntentRevision++;
                CaptureNavigation();
            };
        }
        private void RestoreHierarchyNavigation()
        {
            if (hierarchyControl == null) return;
            restoringNavigation = true;
            try
            {
                if (resetNavigationPending)
                {
                    hierarchyControl.ResetNavigation(); resetNavigationPending = false;
                    retainedSelectionRow = retainedExplorerOccurrence = null; retainedExplorerSide = null; retainContextWithoutRow = true; return;
                }
                if (hierarchyBookmarkRevision == selectionIntentRevision)
                    hierarchyControl.RestoreNavigation(hierarchyBookmark, retainedSelectionRow, retainedExplorerSide);
                else if (retainContextWithoutRow && !retainedSelectionRow.HasValue)
                    hierarchyControl.ClearNavigationSelection();
            }
            finally { restoringNavigation = false; }
        }
        private void AcceptFlatSelection(MembershipGridRow row)
        {
            if (!resultsGrid.Visible || restoringNavigation || navigationTransaction) return;
            if (retainedSelectionRow != row?.OriginalRowIndex) retainedExplorerOccurrence = null;
            retainedSelectionRow = row?.OriginalRowIndex;
            if (row == null) retainedExplorerSide = null;
            retainContextWithoutRow = row == null; selectionIntentRevision++;
        }
        private void RestoreNavigationFocus()
        {
            if (!restoreFocusPending || focusWorkspace != technicalWorkspace || !CanStartLocalView) return;
            restoreFocusPending = false;
            if (!technicalWorkspace && objectsSearch.TextBox.Focused) return;
            if (explorerControl?.Visible == true) explorerControl.RestoreNavigationFocus();
            else if (hierarchyControl?.Visible == true) hierarchyControl.RestoreNavigationFocus();
            else if (resultsGrid.Visible) resultsGrid.Focus();
        }
        private async Task ResetNavigationFiltersAsync()
        {
            CaptureNavigation(); resetSupportChoice = explorerControl?.CaptureNavigation().Support ?? false;
            navigationTransaction = true;
            try
            {
                retainedSelectionRow = retainedExplorerOccurrence = null; retainedExplorerSide = null; retainContextWithoutRow = true;
                if (technicalWorkspace)
                {
                    membershipFilter.SelectedIndex = 0; hierarchyBookmark = null; explorerBookmarks.Remove("Technical:Explorer");
                    if (viewSelector.SelectedIndex == 2) { viewSelector.SelectedIndex = 0; lastTechnicalView = 0; }
                }
                else
                {
                    objectsSearch.Clear(); objectsMembershipFilter.SelectedIndex = 0; objectsDifferences.Checked = false;
                    foreach (var key in explorerBookmarks.Keys.Where(k => k.StartsWith("Objects:", StringComparison.Ordinal)).ToArray()) explorerBookmarks.Remove(key);
                }
                resetNavigationPending = true;
            }
            finally { navigationTransaction = false; }
            ApplyMembershipFilter(); UpdateWorkspaceChrome();
            if (CanStartLocalView) await ShowSelectedViewAsync();
            if (technicalWorkspace && resultsGrid.Visible)
            {
                resultsGrid.CurrentCell = null; resetNavigationPending = false; displayEvidenceButton.Enabled = false;
                viewStatus.Text = "Filters reset; Flat View shows all " + gridRows.Count + " comparison rows";
            }
        }
        // The plugin's normal entry point uses this integrated browser. The retained technical
        // presentation remains an internal maintenance surface, with no end-user navigation route.
        internal void UseSimplifiedExplorer()
        {
            if (IsHandleCreated) throw new InvalidOperationException("Choose the result presentation before showing the window.");
            if (simplifiedExplorer != null) return;
            simplifiedExplorer = new SimplifiedSolutionExplorerPanel(presentation, FormatDisplayEvidenceDetails);
            simplifiedExplorer.UseObjectNavigation();
            foreach (Control control in Controls) control.Visible = false;
            Text = "Solution Explorer - " + presentation.SolutionUniqueName;
            Controls.Add(simplifiedExplorer); simplifiedExplorer.BringToFront();
        }

        private async Task ShowSimplifiedExplorerAsync()
        {
            var token = hierarchyLifetime.Token; int version = ++viewRequestVersion;
            try
            {
                if (hierarchyTask == null) hierarchyTask = Task.Run(() => ProjectHierarchy(token), token);
                if (explorerTask == null) explorerTask = ProjectExplorerAfterHierarchyAsync(hierarchyTask, token);
                var projection = await explorerTask.ConfigureAwait(false); token.ThrowIfCancellationRequested();
                await OnHierarchyUiAsync(() =>
                {
                    if (version != viewRequestVersion || simplifiedExplorer == null) return;
                    ExplorerProjection = projection; simplifiedExplorer.Bind(projection);
                }, token).ConfigureAwait(false);
            }
            catch (OperationCanceledException) { /* Closed windows cannot publish retained local results. */ }
            catch (Exception error)
            {
                Debug.WriteLine("Solution Explorer unavailable: " + error.GetType().Name);
                if (!token.IsCancellationRequested)
                    await OnHierarchyUiAsync(() => { if (version == viewRequestVersion) simplifiedExplorer?.ShowFailure(); }, token).ConfigureAwait(false);
            }
        }

        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            if (simplifiedExplorer != null && keyData == (Keys.Control | Keys.F))
            { simplifiedExplorer.FocusSearch(); return true; }
            if (!technicalWorkspace && primaryBar.Visible && keyData == (Keys.Control | Keys.F))
            { objectsSearch.Focus(); objectsSearch.TextBox.SelectAll(); return true; }
            if (!technicalWorkspace && keyData == Keys.Enter && explorerControl?.Grid.Focused == true && explorerControl.SelectedComparison != null)
            { objectsDetails.Checked = true; return true; }
            return base.ProcessCmdKey(ref msg, keyData);
        }

        protected override void OnShown(EventArgs e)
        {
            base.OnShown(e);
            shownCallbacksCompleted = true;
            if (simplifiedExplorer != null)
            { BeginInvoke(new Action(async () => await ShowSimplifiedExplorerAsync())); return; }
            if (initialViewQueued || IsDisposed || Disposing) return;
            initialViewQueued = true;
            // Defer until the visible handle, controls and caller's test injections are ready.
            BeginInvoke(new Action(async () =>
            {
                if (!CanStartLocalView || initialViewStarted) return;
                await ShowSelectedViewAsync();
            }));
        }

        private bool CanStartLocalView => shownCallbacksCompleted && IsHandleCreated && Visible && !IsDisposed && !Disposing;

        private async Task SelectWorkspaceAsync(bool technical)
        {
            CaptureNavigation();
            RetainViewSelection(); viewRequestVersion++;
            if (technicalWorkspace) lastTechnicalView = viewSelector.SelectedIndex;
            Action updateWorkspace = () =>
            {
                technicalWorkspace = technical; switchingWorkspace = true;
                try { viewSelector.SelectedIndex = technical ? lastTechnicalView : 3; }
                finally { switchingWorkspace = false; }
                UpdateWorkspaceChrome();
            };
            if (explorerControl == null) updateWorkspace();
            else explorerControl.WithNavigationRestoration(updateWorkspace);
            if (CanStartLocalView) await ShowSelectedViewAsync();
            RestoreNavigationFocus();
        }
        private void UpdateWorkspaceChrome()
        {
            technicalBar.Visible = technicalWorkspace; primaryBar.Visible = !technicalWorkspace;
            objectsAction.Checked = workspaceObjectsAction.Checked = !technicalWorkspace; technicalAction.Checked = technicalWorkspace;
            objectsNotice.Visible = !technicalWorkspace;
            summaryLabel.Visible = technicalWorkspace; diagnosticLabel.Visible = false;
            navigationPanel.Visible = technicalWorkspace;
            LayoutWorkspaceHeader();
            objectsDifferences.Enabled = CurrentObjectsScope == ObjectsWorkspaceScope.Compare;
        }
        protected override void OnLayout(LayoutEventArgs e)
        {
            base.OnLayout(e);
            LayoutWorkspaceHeader();
        }
        private void LayoutWorkspaceHeader()
        {
            if (simplifiedExplorer != null) return;
            if (arrangingHeader || objectsHeader == null || headerPanel == null || navigationPanel == null) return;
            arrangingHeader = true;
            try
            {
                foreach (var label in headerPanel.Controls.OfType<Label>()) label.Visible = technicalWorkspace && !ReferenceEquals(label, diagnosticLabel);
                objectsHeader.Primary = !technicalWorkspace;
                objectsHeader.Visible = !technicalWorkspace || objectsHeader.Expanded;
                objectsHeader.Dock = technicalWorkspace ? DockStyle.Bottom : DockStyle.Top;
                int width = Math.Max(1, headerPanel.ClientSize.Width - headerPanel.Padding.Horizontal);
                int measured = objectsHeader.MeasureHeight(width);
                objectsNotice.Height = TextRenderer.MeasureText(objectsNotice.Text, objectsNotice.Font,
                    new Size(Math.Max(1, ClientSize.Width - objectsNotice.Padding.Horizontal), int.MaxValue),
                    TextFormatFlags.WordBreak | TextFormatFlags.TextBoxControl).Height + objectsNotice.Padding.Vertical;
                objectsHeader.Height = measured;
                navigationPanel.Height = technicalWorkspace ? 24 : Math.Max(24, displayEvidenceButton.GetPreferredSize(Size.Empty).Height);
                int height = technicalWorkspace ? 174 + (objectsHeader.Expanded ? measured : 0) :
                    measured + headerPanel.Padding.Vertical;
                if (headerPanel.Height != height) headerPanel.Height = height;
            }
            finally { arrangingHeader = false; }
        }

        private string AdvancedDiagnosticsText() => summaryLabel.Text + Environment.NewLine + diagnosticLabel.Text + Environment.NewLine +
            "Source snapshot timestamp: " + RetainedSnapshotTimestamp(presentation.Source.Snapshot) + Environment.NewLine +
            "Target snapshot timestamp: " + RetainedSnapshotTimestamp(presentation.Target.Snapshot);
        private static string RetainedSnapshotTimestamp(MembershipSnapshot snapshot) => snapshot == null ?
            "[Unavailable — no retained snapshot]" : snapshot.CapturedAt.ToString("o", CultureInfo.InvariantCulture);

        private sealed class ObjectsSnapshotHeader : UserControl
        {
            private readonly TableLayoutPanel layout;
            private readonly Label title, source, target, snapshot, compactSummary;
            private readonly CheckBox expand;
            private readonly TextBox diagnosticText;
            private bool primary = true;
            private readonly Func<string> diagnosticsFactory;
            private string retainedDiagnostics;
            internal event EventHandler ExpandedChanged;
            internal bool Expanded { get => expand.Checked; set => expand.Checked = value; }
            internal bool Primary
            {
                set { primary = value; title.Visible = source.Visible = target.Visible = snapshot.Visible = compactSummary.Visible = value; expand.Visible = false; diagnosticText.Visible = Expanded; }
            }
            internal ObjectsSnapshotHeader(string heading, string sourceText, string targetText, Func<string> diagnostics,
                string snapshotText, string summaryText)
            {
                diagnosticsFactory = diagnostics;
                Name = "ObjectsSnapshotHeader"; AutoScaleMode = AutoScaleMode.Inherit;
                layout = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 7, Margin = Padding.Empty };
                layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
                for (int i = 0; i < 7; i++) layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
                title = HeaderLabel("ObjectsSolutionTitle", heading);
                title.Font = new Font("Segoe UI", 11F, FontStyle.Bold);
                source = HeaderLabel("ObjectsSourceState", sourceText); target = HeaderLabel("ObjectsTargetState", targetText);
                snapshot = HeaderLabel("ObjectsSnapshotContext", snapshotText);
                compactSummary = HeaderLabel("ObjectsGlobalSummary", summaryText);
                compactSummary.Font = new Font("Segoe UI", 9F, FontStyle.Bold);
                expand = new CheckBox { Name = "ObjectsAdvancedDiagnostics", Text = "Advanced &diagnostics", AutoSize = true,
                    AccessibleName = "Show complete advanced diagnostics", TabIndex = 0, Margin = new Padding(0, 2, 0, 2) };
                diagnosticText = new TextBox { Name = "CompleteMembershipDiagnostics", Multiline = true,
                    ReadOnly = true, WordWrap = false, ScrollBars = ScrollBars.Both, Dock = DockStyle.Top,
                    AccessibleName = "Complete comparison summary and Source Target diagnostics", TabIndex = 1, Visible = false };
                layout.Controls.Add(title, 0, 0); layout.Controls.Add(source, 0, 1); layout.Controls.Add(target, 0, 2);
                layout.Controls.Add(snapshot, 0, 3); layout.Controls.Add(compactSummary, 0, 4);
                layout.Controls.Add(expand, 0, 5); layout.Controls.Add(diagnosticText, 0, 6); Controls.Add(layout);
                expand.CheckedChanged += (sender, args) =>
                {
                    if (Expanded && retainedDiagnostics == null) diagnosticText.Text = retainedDiagnostics = diagnosticsFactory();
                    diagnosticText.Visible = Expanded; ExpandedChanged?.Invoke(this, EventArgs.Empty);
                };
            }
            private static Label HeaderLabel(string name, string text) => new Label { Name = name, Text = text, AutoSize = true,
                Dock = DockStyle.Top, Margin = new Padding(0, 0, 0, 3), AccessibleName = text, AccessibleDescription = text };
            internal int MeasureHeight(int width)
            {
                int height = 0;
                foreach (var label in new[] { title, source, target, snapshot, compactSummary })
                {
                    label.MaximumSize = new Size(width, 0);
                    if (primary) height += label.GetPreferredSize(new Size(width, 0)).Height + label.Margin.Vertical;
                }

                diagnosticText.Height = Math.Max(80, Font.Height * 6);
                if (Expanded) height += diagnosticText.Height + diagnosticText.Margin.Vertical;
                return height;
            }
        }
        private static string FormatCompactSummary(MembershipPresentationSummary summary) =>
            "Compared components: " + summary.TotalComparisonRows + " | Present in Both: " + summary.PresentInBoth +
            " | Source Only: " + summary.SourceOnly + " | Target Only: " + summary.TargetOnly +
            " | Needs Review: " + (summary.IndeterminateUnsupported + summary.IndeterminateUnresolved +
                summary.IndeterminateAmbiguous + summary.IndeterminateResolved + summary.IndeterminateOtherOrUnavailable);

        private IReadOnlyList<MembershipGridRow> WorkspaceRows() => technicalWorkspace ? FilteredRows() :
            gridRows.Where(r => (CurrentObjectsScope != ObjectsWorkspaceScope.Compare || !objectsDifferences.Checked ||
                r.AuthoritativeRow.Presence != MembershipPresence.PresentInBoth || r.DefinitionStatus == "Different") &&
                MatchesPresence(r, objectsMembershipFilter.SelectedIndex)).ToList().AsReadOnly();
        private void NavigateToOriginalRow(MembershipGridRow row)
        {
            CaptureNavigation(); RetainViewSelection(); retainedSelectionRow = row.OriginalRowIndex;
            retainContextWithoutRow = false; selectionIntentRevision++;
            Action navigate = () =>
            {
                technicalWorkspace = true;
                if (!FilteredRows().Any(r => r.OriginalRowIndex == row.OriginalRowIndex)) membershipFilter.SelectedIndex = 0;
                viewSelector.SelectedIndex = 0; UpdateWorkspaceChrome(); ShowFlatView();
            };
            if (explorerControl == null) navigate();
            else explorerControl.WithNavigationRestoration(navigate);
        }
        private Task ShowSelectedViewAsync()
        {
            if (!CanStartLocalView) return Task.CompletedTask;
            initialViewStarted = true;
            if (!technicalWorkspace) return ShowExplorerAsync();
            switch (viewSelector.SelectedIndex)
            {
                case 1: return ShowHierarchyAsync();
                case 3: return ShowExplorerAsync();
                default: ShowFlatView(); return Task.CompletedTask;
            }
        }

        private void ShowPreparation(string message)
        {
            resultsGrid.Visible = false;
            if (hierarchyControl != null) hierarchyControl.Visible = false;
            if (explorerControl != null) explorerControl.Visible = false;
            viewPreparation.Text = message; viewPreparation.Visible = true;
            displayEvidenceButton.Enabled = false; originalRowAction.Enabled = false; viewStatus.Text = message;
        }

        private MembershipGridRow SelectedComparisonRow() => explorerControl?.Visible == true ? explorerControl.SelectedComparison : hierarchyControl?.Visible == true ?
            hierarchyControl.SelectedComparison : resultsGrid.Visible && !retainContextWithoutRow ? resultsGrid.CurrentRow?.DataBoundItem as MembershipGridRow : null;

        internal IReadOnlyList<MembershipGridRow> FilteredRows() => gridRows.Where(r =>
            (viewSelector.SelectedIndex != 2 || r.AuthoritativeRow.Presence != MembershipPresence.PresentInBoth ||
                r.DefinitionStatus == "Different") && (membershipFilter.SelectedIndex == 0 ||
            r.AuthoritativeRow.Presence == (membershipFilter.SelectedIndex == 1 ? MembershipPresence.PresentInBoth :
            membershipFilter.SelectedIndex == 2 ? MembershipPresence.OnlyInSource : membershipFilter.SelectedIndex == 3 ?
            MembershipPresence.OnlyInTarget : MembershipPresence.Indeterminate))).ToList().AsReadOnly();

        private void ApplyMembershipFilter()
        {
            var selected = (resultsGrid.CurrentRow?.DataBoundItem as MembershipGridRow)?.OriginalRowIndex;
            resultsGrid.DataSource = FilteredRows().ToList();
            if (selected.HasValue)
                foreach (DataGridViewRow row in resultsGrid.Rows)
                    if ((row.DataBoundItem as MembershipGridRow)?.OriginalRowIndex == selected) { resultsGrid.CurrentCell = row.Cells[0]; break; }
            viewStatus.Text = FilteredRows().Count + " comparison rows";
        }

        private void ShowFlatView()
        {
            legacyViewLabel.Text = "View:";
            var selected = explorerControl?.Visible == true ? explorerControl.SelectedComparison :
                hierarchyControl?.Visible == true ? hierarchyControl.SelectedComparison : null;
            if (explorerControl != null) explorerControl.Visible = false;
            if (selected != null) retainedSelectionRow = selected.OriginalRowIndex;
            if (hierarchyControl != null) hierarchyControl.Visible = false;
            viewPreparation.Visible = false; resultsGrid.Visible = true;
            if (retainedSelectionRow.HasValue)
                foreach (DataGridViewRow row in resultsGrid.Rows)
                    if ((row.DataBoundItem as MembershipGridRow)?.OriginalRowIndex == retainedSelectionRow.Value)
                    { resultsGrid.CurrentCell = row.Cells[0]; break; }
            if (retainContextWithoutRow) resultsGrid.CurrentCell = null;
            displayEvidenceButton.Enabled = SelectedComparisonRow() != null;
            viewStatus.Text = FilteredRows().Count + " comparison rows";
        }

        internal async Task ShowHierarchyAsync()
        {
            if (!CanStartLocalView) return;
            CaptureNavigation();
            legacyViewLabel.Text = "View:";
            if (explorerControl != null) explorerControl.Visible = false;
            int version = ++viewRequestVersion;
            var token = hierarchyLifetime.Token;
            var filtered = FilteredRows(); int filter = membershipFilter.SelectedIndex; bool all = filter == 0;
            try
            {
                ShowPreparation("Preparing local hierarchy...");
                if (hierarchyTask == null) hierarchyTask = Task.Run(() => ProjectHierarchy(token), token);
                var projection = await hierarchyTask.ConfigureAwait(false);
                token.ThrowIfCancellationRequested();
                if (projection.Outcome != HierarchyProjectionOutcome.Completed) throw new InvalidOperationException("Hierarchy projection unavailable.");
                var model = ReferenceEquals(HierarchyProjection, projection) && boundHierarchyFilter == filter && currentHierarchyModel != null
                    ? currentHierarchyModel : await Task.Run(() => MembershipHierarchyView.CreateAdministratorView(projection, gridRows,
                        filtered.Select(r => r.OriginalRowIndex), all, token), token).ConfigureAwait(false);
                await OnHierarchyUiAsync(() =>
                {
                    if (viewSelector.SelectedIndex != 1 || version != viewRequestVersion) return;
                    if (hierarchyControl == null)
                    {
                        hierarchyControl = new MembershipHierarchyControl(FormatDisplayEvidenceDetails, OpenDefinitionDetails) { Visible = false };
                        hierarchyControl.SelectedComparisonChanged += row => displayEvidenceButton.Enabled = row != null;
                        hierarchyControl.RenderingFailed = HierarchyFailed;
                        ConnectHierarchyNavigation();
                        Controls.Add(hierarchyControl);
                        Controls.SetChildIndex(hierarchyControl, 0);
                    }
                    var selected = retainedSelectionRow ?? hierarchyControl.SelectedComparison?.OriginalRowIndex;
                    var side = retainedExplorerSide ?? hierarchyControl.SelectedSide;
                    if (!ReferenceEquals(HierarchyProjection, projection) || boundHierarchyFilter != filter)
                    { hierarchyControl.Bind(model, presentation); currentHierarchyModel = model; boundHierarchyFilter = filter; }
                    if (selected.HasValue && (hierarchyControl.SelectedComparison?.OriginalRowIndex != selected.Value || hierarchyControl.SelectedSide != side))
                        hierarchyControl.SelectComparison(selected.Value, side);
                    HierarchyProjection = projection;
                    retainedExplorerSide = side;
                    RestoreHierarchyNavigation();
                    if (viewSelector.SelectedIndex != 1 || version != viewRequestVersion) return;
                    viewPreparation.Visible = false; resultsGrid.Visible = false; hierarchyControl.Visible = true;
                    RestoreNavigationFocus();
                    displayEvidenceButton.Enabled = SelectedComparisonRow() != null;
                    viewStatus.Text = model.ComparisonRowCount + " comparison rows; " + model.VisibleOccurrenceCount + " occurrences (including context)";
                }, token).ConfigureAwait(false);
            }
            catch (OperationCanceledException)
            { if (!token.IsCancellationRequested) await ReportHierarchyFailureAsync(null, version, token).ConfigureAwait(false); }
            catch (Exception ex)
            { if (!token.IsCancellationRequested) await ReportHierarchyFailureAsync(ex, version, token).ConfigureAwait(false); }
        }

        private void RetainViewSelection()
        {
            var priorRow = retainedSelectionRow; var priorSide = retainedExplorerSide; var priorOccurrence = retainedExplorerOccurrence;
            bool priorContextWithoutRow = retainContextWithoutRow;
            if (explorerControl?.Visible == true && explorerControl.SelectedComparison != null)
            { retainedSelectionRow = explorerControl.SelectedComparison.OriginalRowIndex; retainedExplorerSide = explorerControl.SelectedSide; retainedExplorerOccurrence = explorerControl.SelectedOccurrenceIndex; retainContextWithoutRow = false; }
            else if (hierarchyControl?.Visible == true && hierarchyControl.SelectedComparison != null)
            { retainedSelectionRow = hierarchyControl.SelectedComparison.OriginalRowIndex; retainedExplorerSide = hierarchyControl.SelectedSide; retainedExplorerOccurrence = hierarchyControl.SelectedOccurrenceIndex; retainContextWithoutRow = false; }
            else if (resultsGrid.Visible && !retainContextWithoutRow && resultsGrid.CurrentRow?.DataBoundItem is MembershipGridRow row)
                retainedSelectionRow = row.OriginalRowIndex;
            if (priorRow != retainedSelectionRow || priorSide != retainedExplorerSide || priorOccurrence != retainedExplorerOccurrence || priorContextWithoutRow != retainContextWithoutRow)
                selectionIntentRevision++;
        }

        private async Task<SolutionObjectsExplorerProjection> ProjectExplorerAfterHierarchyAsync(
            Task<ComponentHierarchyProjection> pendingHierarchy, CancellationToken token)
        {
            var hierarchy = await pendingHierarchy.ConfigureAwait(false);
            token.ThrowIfCancellationRequested();
            if (hierarchy.Outcome != HierarchyProjectionOutcome.Completed)
                throw new InvalidOperationException("Hierarchy projection unavailable.");
            return await Task.Run(() => ProjectExplorer(hierarchy, token), token).ConfigureAwait(false);
        }
        internal async Task ShowExplorerAsync()
        {
            if (!CanStartLocalView || viewSelector.SelectedIndex != 3) return;
            int version = ++viewRequestVersion; var token = hierarchyLifetime.Token;
            CaptureNavigation();
            var selected = SelectedComparisonRow();
            if (selected != null) retainedSelectionRow = selected.OriginalRowIndex;
            try
            {
                ShowPreparation("Preparing local object navigation...");
                if (hierarchyTask == null) hierarchyTask = Task.Run(() => ProjectHierarchy(token), token);
                // Install the shared task on the UI thread before awaiting. Filter/view events
                // must reuse it even while the hierarchy stage is still running.
                if (explorerTask == null) explorerTask = ProjectExplorerAfterHierarchyAsync(hierarchyTask, token);
                var hierarchy = await hierarchyTask.ConfigureAwait(false); token.ThrowIfCancellationRequested();
                var projection = await explorerTask.ConfigureAwait(false); token.ThrowIfCancellationRequested();
                await OnHierarchyUiAsync(() =>
                {
                    if (viewSelector.SelectedIndex != 3 || version != viewRequestVersion) return;
                    if (explorerControl == null)
                    {
                        explorerControl = new SolutionObjectsExplorerControl(FormatDisplayEvidenceDetails, row =>
                        {
                            NavigateToOriginalRow(row);
                        }, OpenDefinitionDetails) { Visible = false };
                        explorerControl.SelectedComparisonChanged = row => { displayEvidenceButton.Enabled = row != null; originalRowAction.Enabled = row != null; };
                        explorerControl.TechnicalMembershipRequested = async () => await SelectWorkspaceAsync(true);
                        explorerControl.RenderingFailed = ExplorerFailed; Controls.Add(explorerControl); Controls.SetChildIndex(explorerControl, 0);
                        ConnectExplorerNavigation();
                        explorerControl.ConfigureWorkspace(!technicalWorkspace, CurrentObjectsScope);
                        explorerControl.Bind(projection, WorkspaceRows().Select(r => r.OriginalRowIndex), ExplorerFilterAll);
                        boundExplorerFilter = ActiveExplorerFilter; boundExplorerDifferences = viewSelector.SelectedIndex == 2;
                    }
                    else if (explorerControl.ConfigureWorkspace(!technicalWorkspace, CurrentObjectsScope) ||
                        boundExplorerFilter != ActiveExplorerFilter || boundExplorerDifferences != (viewSelector.SelectedIndex == 2) || !technicalWorkspace)
                    {
                        explorerControl.ApplyFilter(WorkspaceRows().Select(r => r.OriginalRowIndex), ExplorerFilterAll);
                        boundExplorerFilter = ActiveExplorerFilter; boundExplorerDifferences = viewSelector.SelectedIndex == 2;
                    }
                    explorerControl.ShowMetadata(showExplorerMetadata); explorerControl.ShowDetails(objectsDetails.Checked);
                    explorerControl.SetSearch(technicalWorkspace ? string.Empty : objectsSearch.Text);
                    if (retainedSelectionRow.HasValue) explorerControl.SelectComparison(retainedSelectionRow.Value, retainedExplorerSide ?? hierarchyControl?.SelectedSide);
                    if (retainedExplorerOccurrence.HasValue && retainedSelectionRow.HasValue &&
                        projection.Occurrences[retainedExplorerOccurrence.Value].Original.ComparisonRowIndices.Contains(retainedSelectionRow.Value) &&
                        (technicalWorkspace || CurrentObjectsScope == ObjectsWorkspaceScope.Compare ||
                         projection.Occurrences[retainedExplorerOccurrence.Value].Original.Side == (CurrentObjectsScope == ObjectsWorkspaceScope.Target ? DisplayEnvironmentSide.Target : DisplayEnvironmentSide.Source)))
                    {
                        explorerControl.SelectOccurrence(retainedExplorerOccurrence.Value);
                        explorerControl.SelectComparison(retainedSelectionRow.Value, retainedExplorerSide);
                    }
                    RestoreExplorerNavigation();
                    if (version != viewRequestVersion || viewSelector.SelectedIndex != 3) return;
                    ExplorerProjection = projection; HierarchyProjection = hierarchy;
                    legacyViewLabel.Text = "View:";
                    viewPreparation.Visible = false;
                    resultsGrid.Visible = false; if (hierarchyControl != null) hierarchyControl.Visible = false;
                    explorerControl.Visible = true; displayEvidenceButton.Enabled = explorerControl.SelectedComparison != null;
                    RestoreNavigationFocus();
                    viewStatus.Text = projection.Rows.Count + " original comparison rows; " + projection.Occurrences.Count + " raw occurrences; object candidates, not maker totals";
                }, token).ConfigureAwait(false);
            }
            catch (OperationCanceledException) { if (!token.IsCancellationRequested) await OnHierarchyUiAsync(() => { if (version == viewRequestVersion) ExplorerFailed(null); }, token).ConfigureAwait(false); }
            catch (Exception error) { if (!token.IsCancellationRequested) await OnHierarchyUiAsync(() => { if (version == viewRequestVersion) ExplorerFailed(error); }, token).ConfigureAwait(false); }
        }
        private void ExplorerFailed(Exception error)
        {
            var diagnostic = "Objects Explorer unavailable" + (error == null ? "" : " (" + error.GetType().Name + ")") + "; completed comparison preserved.";
            HierarchyDiagnostics.Add(diagnostic); Debug.WriteLine(diagnostic); explorerButton.Checked = false; viewRequestVersion++;
            technicalWorkspace = true; viewSelector.SelectedIndex = 0; UpdateWorkspaceChrome(); ShowFlatView(); viewStatus.Text = diagnostic; viewStatus.ToolTipText = diagnostic;
        }

        private async Task ReportHierarchyFailureAsync(Exception error, int version, CancellationToken token)
        {
            try { await OnHierarchyUiAsync(() => { if (version == viewRequestVersion) HierarchyFailed(error); }, token).ConfigureAwait(false); }
            catch (OperationCanceledException) { /* Form closed while the optional notification was queued. */ }
        }

        // Do not rely on an ambient SynchronizationContext: hosts and STA test runners can differ.
        private Task OnHierarchyUiAsync(Action action, CancellationToken token)
        {
            if (IsDisposed || Disposing || token.IsCancellationRequested || !IsHandleCreated) return Task.CompletedTask;
            if (!InvokeRequired) { action(); return Task.CompletedTask; }
            var completion = new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
            CancellationTokenRegistration registration;
            try { registration = token.Register(() => completion.TrySetCanceled()); }
            catch (ObjectDisposedException) { return Task.CompletedTask; }
            completion.Task.ContinueWith(_ => registration.Dispose(), CancellationToken.None, TaskContinuationOptions.None, TaskScheduler.Default);
            try
            {
                BeginInvoke(new Action(() =>
                {
                    try { if (!IsDisposed && !Disposing && !token.IsCancellationRequested) action(); completion.TrySetResult(true); }
                    catch (Exception ex) { completion.TrySetException(ex); }
                }));
            }
            catch (InvalidOperationException ex)
            { if (IsDisposed || token.IsCancellationRequested) completion.TrySetResult(true); else completion.TrySetException(ex); }
            return completion.Task;
        }

        private void HierarchyFailed(Exception error)
        {
            var diagnostic = "Hierarchy unavailable" + (error == null ? "" : " (" + error.GetType().Name + ")") + "; completed comparison preserved.";
            HierarchyDiagnostics.Add(diagnostic); Debug.WriteLine(diagnostic);
            viewSelector.SelectedIndex = 0; ShowFlatView();
            hierarchyControl?.ClearHierarchy(); boundHierarchyFilter = -1;
            currentHierarchyModel = null;
            viewStatus.Text = "Hierarchy unavailable. Flat results preserved."; viewStatus.ToolTipText = diagnostic;
        }

        private void OpenDefinitionDetails(MembershipGridRow row)
        {
            if (row?.AuthoritativeRow.DefinitionDetail == null) return;
            using (var form = new ComponentDefinitionDetailsForm(row.AuthoritativeRow.DefinitionDetail)) form.ShowDialog(this);
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing && !hierarchyLifetimeDisposed)
            { hierarchyLifetimeDisposed = true; hierarchyLifetime.Cancel(); hierarchyLifetime.Dispose(); }
            base.Dispose(disposing);
        }

        private void ResultsGrid_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0) return;
            var row = GetAuthoritativeRow(resultsGrid.Rows[e.RowIndex].DataBoundItem);
            if (row?.DefinitionDetail == null) return;
            using (var form = new ComponentDefinitionDetailsForm(row.DefinitionDetail))
                form.ShowDialog(this);
        }

        private void CoverageDetails_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            var builder = new MembershipCoverageDiagnosticsBuilder();
            var source = presentation.Source.Snapshot == null
                ? builder.BuildUnavailable() : builder.Build(presentation.Source.Snapshot);
            var target = presentation.Target.Snapshot == null
                ? builder.BuildUnavailable() : builder.Build(presentation.Target.Snapshot);
            using (var form = new MembershipCoverageDetailsForm(
                presentation.Source.Diagnostics.EnvironmentName, source,
                presentation.Target.Diagnostics.EnvironmentName, target, presentation,
                sourceSolutionVersion, targetSolutionVersion,
                sourceService, targetService, captureAppSettingEvidence
#if DEBUG
                , CaptureType92Evidence, CaptureType59Evidence, DiscoverType59Evidence,
                CaptureCloudFlowEvidence, CaptureSavedQueryEvidence, CaptureUnsupportedCoverageInventory, CaptureType36Evidence, CaptureType31Evidence, CaptureType10072Evidence, CaptureType300Evidence, CaptureType74Evidence, CaptureType511Evidence, CaptureType10276Evidence, CaptureAppActionEvidence
#endif
                ))
                form.ShowDialog(this);
        }

        internal static MembershipResultRow GetAuthoritativeRow(object boundItem) =>
            (boundItem as MembershipGridRow)?.AuthoritativeRow ?? boundItem as MembershipResultRow;

        internal static int ScaleColumnWidth(int width, int dpi) =>
            Math.Max(5, (int)Math.Round(width * Math.Max(1, dpi) / 96.0));

        private void AddColumn(string propertyName, string title, float fillWeight, int minimumWidth, bool fill = true, string binding = null)
        {
            resultsGrid.Columns.Add(new DataGridViewTextBoxColumn
            {
                DataPropertyName = binding ?? propertyName,
                HeaderText = title,
                Name = propertyName,
                AutoSizeMode = fill ? DataGridViewAutoSizeColumnMode.Fill : DataGridViewAutoSizeColumnMode.None,
                FillWeight = fill ? fillWeight : 100,
                MinimumWidth = ScaleColumnWidth(fill ? minimumWidth : Math.Min(100, minimumWidth), DeviceDpi),
                Width = ScaleColumnWidth(minimumWidth, DeviceDpi),
                ReadOnly = true,
                Resizable = DataGridViewTriState.True,
                SortMode = DataGridViewColumnSortMode.NotSortable,
                DefaultCellStyle = new DataGridViewCellStyle { WrapMode = DataGridViewTriState.True }
            });
        }

        private sealed class MembershipAccessibleRow : DataGridViewRow
        {
            protected override AccessibleObject CreateAccessibilityInstance() =>
                new MembershipRowAccessibleObject(this);

            private sealed class MembershipRowAccessibleObject : DataGridViewRowAccessibleObject
            {
                private readonly DataGridViewRow row;

                public MembershipRowAccessibleObject(DataGridViewRow row) : base(row)
                {
                    this.row = row;
                }

                public override string Description
                {
                    get
                    {
                        var item = GetAuthoritativeRow(row.DataBoundItem);
                        return item == null ? base.Description : StatusDescription(item);
                    }
                }

                public override Rectangle Bounds
                {
                    get
                    {
                        var grid = row.DataGridView;
                        if (grid == null || !grid.IsHandleCreated || row.Index < 0 || !row.Displayed)
                            return Rectangle.Empty;

                        // Framework row accessibility recursively asks preceding off-screen rows
                        // for Bounds. Large membership inventories make accessibility clients stall
                        // the UI thread. Obtain the visible bounds directly without that recursion.
                        var bounds = grid.GetRowDisplayRectangle(row.Index, true);
                        return bounds.IsEmpty ? Rectangle.Empty : grid.RectangleToScreen(bounds);
                    }
                }
            }
        }

        private void ResultsGrid_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            if (e.RowIndex < 0) return;
            if (e.ColumnIndex >= 4 && e.Value is string)
            {
                e.Value = MembershipGridRow.WithoutBothPrefix((string)e.Value);
                e.FormattingApplied = true;
            }
            var row = GetAuthoritativeRow(resultsGrid.Rows[e.RowIndex].DataBoundItem);
            if (row == null) return;
            if (row.MembershipStatus == "Present in Both")
            {
                e.CellStyle.BackColor = Color.Honeydew;
                e.CellStyle.ForeColor = Color.DarkGreen;
            }
            else if (row.MembershipStatus == "Source Only" || row.MembershipStatus == "Target Only")
            {
                e.CellStyle.BackColor = Color.MistyRose;
                e.CellStyle.ForeColor = Color.Firebrick;
            }
            else
            {
                e.CellStyle.BackColor = Color.LemonChiffon;
                e.CellStyle.ForeColor = Color.DarkGoldenrod;
            }
        }

        private void ResultsGrid_CellToolTipTextNeeded(object sender, DataGridViewCellToolTipTextNeededEventArgs e)
        {
            if (e.RowIndex < 0 || e.ColumnIndex < 0) return;
            var row = resultsGrid.Rows[e.RowIndex].DataBoundItem as MembershipGridRow;
            if (row == null) return;
            switch (resultsGrid.Columns[e.ColumnIndex].DataPropertyName)
            {
                case "ComponentText": e.ToolTipText = row.ComponentToolTip; break;
                case "TypeText": e.ToolTipText = row.ComponentCategoryToolTip; break;
                case "ComponentCategoryText": e.ToolTipText = row.ComponentCategoryToolTip; break;
                case "FriendlyNameText": e.ToolTipText = row.FriendlyNameToolTip; break;
                case "InternalNameText": e.ToolTipText = row.InternalNameToolTip; break;
                case "ManagedText": e.ToolTipText = row.ManagedToolTip; break;
                case "CustomizedText": e.ToolTipText = row.CustomizedToolTip; break;
                case "LastModifiedText": e.ToolTipText = row.LastModifiedToolTip; break;
                case "OwnerText": e.ToolTipText = row.OwnerToolTip; break;
                case "ObjectStatusText": e.ToolTipText = row.ObjectStatusToolTip; break;
                case "MembershipStatus": e.ToolTipText = row.MembershipStatus + Environment.NewLine + row.Diagnostic; break;
                case "DefinitionStatusText": e.ToolTipText = MembershipGridRow.DefinitionMeaning + Environment.NewLine + row.DefinitionStatus; break;
            }
            e.ToolTipText = StatusDescription(row.AuthoritativeRow) + Environment.NewLine + e.ToolTipText;
        }

        private static string StatusDescription(MembershipResultRow row) =>
            "Membership status: " + row.MembershipStatus + "; Definition status: " + row.DefinitionStatus;

        internal Form CreateDisplayEvidenceDetails(MembershipGridRow row)
        {
            if (row == null) throw new ArgumentNullException(nameof(row));
            var form = new Form
            {
                Text = "Display Evidence",
                StartPosition = FormStartPosition.CenterParent,
                MinimumSize = new Size(700, 460),
                Size = new Size(950, 680),
                Font = Font,
                ShowInTaskbar = false
            };
            var close = new Button { Text = "&Close", Dock = DockStyle.Bottom, Height = 30, DialogResult = DialogResult.Cancel };
            form.CancelButton = close;
            form.Controls.Add(new TextBox
            {
                Dock = DockStyle.Fill, Multiline = true, ReadOnly = true,
                ScrollBars = ScrollBars.Both, WordWrap = false, BackColor = Color.White,
                AccessibleName = "Source and Target display evidence",
                Text = FormatDisplayEvidenceDetails(row)
            });
            form.Controls.Add(close);
            return form;
        }

        internal string FormatDisplayEvidenceDetails(MembershipGridRow row)
        {
            var text = new StringBuilder();
            text.AppendLine("Original comparison row index: " + row.OriginalRowIndex);
            text.AppendLine("Component Kind (authoritative): " + row.ComponentKind);
            text.AppendLine("Source raw component type: " + row.SourceRawComponentType);
            text.AppendLine("Target raw component type: " + row.TargetRawComponentType);
            text.AppendLine("Source presence: " + row.SourcePresence + "; Target presence: " + row.TargetPresence);
            text.AppendLine("Membership status: " + row.MembershipStatus);
            text.AppendLine("Definition status: " + row.DefinitionStatus);
            text.AppendLine("Source resolution: " + row.SourceResolutionStatus + "; Target resolution: " + row.TargetResolutionStatus);
            text.AppendLine("Complete portable identity: " + row.PortableKey);
            text.AppendLine("Source complete portable identity: " + (row.AuthoritativeRow.SourceIdentity?.ComparisonKey ?? "[Not established]"));
            text.AppendLine("Target complete portable identity: " + (row.AuthoritativeRow.TargetIdentity?.ComparisonKey ?? "[Not established]"));
            text.AppendLine("Changed properties: " + row.ChangedProperties);
            text.AppendLine("Diagnostic / Reason: " + row.Diagnostic);
            text.AppendLine("Display projection outcome: " + (completedComparison?.DisplayProjectionOutcome.ToString() ?? "Not available (legacy caller)"));
            text.AppendLine("Source cache association: " + (completedComparison?.SourceCacheAssociation.ToString() ?? "Not available"));
            text.AppendLine("Target cache association: " + (completedComparison?.TargetCacheAssociation.ToString() ?? "Not available"));
            if (row.DisplayContextDiagnostic != null) text.AppendLine("Display context: " + row.DisplayContextDiagnostic);
            text.AppendLine("Grid escapes backslash, CR, LF and tab; values below are the exact retained strings between BEGIN/END markers.");
            AppendSide(text, "Source", presentation.Source.Diagnostics.EnvironmentName, row.SourceDisplayEvidence);
            AppendSide(text, "Target", presentation.Target.Diagnostics.EnvironmentName, row.TargetDisplayEvidence);
            return text.ToString();
        }

        private static void AppendSide(StringBuilder text, string sideName, string environment,
            ComponentSideDisplayEvidence side)
        {
            text.AppendLine().AppendLine(sideName + " (" + environment + ")");
            if (side == null) { text.AppendLine("[Display evidence unavailable]"); return; }
            text.AppendLine("Has local observation: " + side.HasLocalObservation);
            text.AppendLine("Raw component type: " + side.RawComponentType);
            text.AppendLine("SolutionComponentId (local audit): " + side.SolutionComponentId);
            text.AppendLine("ObjectId (local audit): " + side.LocalObjectId);
            text.AppendLine("Backing correlation status: " + side.BackingCorrelationStatus);
            AppendField(text, "Component category", side.ComponentCategory);
            AppendField(text, "Friendly name", side.FriendlyName);
            AppendField(text, "Internal name", side.InternalName);
            AppendField(text, "Parent context", side.ParentDisplayContext);
            text.AppendLine("Detached object metadata side: " + side.ObjectMetadata.EnvironmentSide);
            if (!side.HasLocalObservation) text.AppendLine("Object metadata: [No local observation]");
            AppendObjectField(text, "Managed", side.ObjectMetadata.Managed);
            AppendObjectField(text, "Customized", side.ObjectMetadata.Customized);
            AppendObjectField(text, "Last Modified", side.ObjectMetadata.LastModified);
            AppendObjectField(text, "Owner", side.ObjectMetadata.Owner);
            AppendObjectField(text, "Object Status", side.ObjectMetadata.ObjectStatus);
        }

        private static void AppendObjectField<T>(StringBuilder text, string name, ComponentObjectMetadataValue<T> field)
        {
            text.AppendLine(name + " availability: " + field.Availability);
            text.AppendLine(name + " requested: " + field.WasRequested);
            text.AppendLine(name + " provenance: " + (field.Provenance ?? "[No provenance retained]"));
            if (field.Availability != DisplayAvailability.Available) return;
            text.AppendLine("BEGIN " + name + " value").AppendLine(MembershipGridRow.FormatObjectValue(field) ?? "[Retrieved null]")
                .AppendLine("END " + name + " value");
            var timestamp = (object)field.Value as ComponentTimestampEvidence;
            if (timestamp != null)
            {
                text.AppendLine("Raw timestamp: " + timestamp.RawTimestamp.ToString("yyyy-MM-dd HH:mm:ss.fffffff", CultureInfo.InvariantCulture));
                text.AppendLine("Timestamp kind: " + timestamp.Kind);
                text.AppendLine("Timestamp behavior: " + (timestamp.Behavior ?? "[Not established]"));
                text.AppendLine("Timezone provenance: " + (timestamp.TimezoneProvenance ?? "[Not established]"));
                text.AppendLine("Verified instant: " + (timestamp.VerifiedInstant?.ToString("o", CultureInfo.InvariantCulture) ?? "[Not established]"));
            }
            var owner = (object)field.Value as ComponentOwnerEvidence;
            if (owner != null)
            {
                text.AppendLine("Owner lookup table (local audit): " + owner.LookupLogicalTable);
                text.AppendLine("Owner ID (local audit): " + owner.LocalOwnerId);
                text.AppendLine("Owner name verified: " + owner.DisplayNameVerified);
                text.AppendLine("Owner correlation: " + owner.Correlation);
            }
            var status = (object)field.Value as ComponentObjectStatusEvidence;
            if (status != null)
            {
                text.AppendLine("Object status field: " + status.SourceField);
                text.AppendLine("Object status raw numeric value: " + status.RawNumericValue);
                text.AppendLine("Object status label verified: " + status.LabelVerified);
                text.AppendLine("Object status interpretation: " + (status.Interpretation ?? "[Not established]"));
            }
        }

        private static void AppendField(StringBuilder text, string name, ComponentDisplayValue field)
        {
            text.AppendLine(name + " availability: " + field.Availability);
            text.AppendLine(name + " provenance: " + (field.DisplayEvidenceSource ?? "[No provenance retained]"));
            if (field.Availability != DisplayAvailability.Available) return;
            if (field.Value == null) { text.AppendLine(name + " value: [Retrieved null]"); return; }
            text.AppendLine(name + " length: " + field.Value.Length);
            text.AppendLine("BEGIN " + name + " value").AppendLine(field.Value).AppendLine("END " + name + " value");
        }

        private static string DisplayState(MembershipEnvironmentResult result)
        {
            switch (result.State)
            {
                case MembershipSnapshotState.Complete:
                    return "Present (" + result.Diagnostics.RawMembershipCount.GetValueOrDefault() + " raw component(s))";
                case MembershipSnapshotState.SolutionAbsent:
                    return "Solution Absent";
                default:
                    return "Unavailable - " + result.Diagnostics.Diagnostic;
            }
        }

        internal static string FormatSummary(MembershipPresentationSummary summary) =>
            "Total Comparison Rows: " + summary.TotalComparisonRows +
            "  |  Present in Both: " + summary.PresentInBoth + "  |  Source Only: " + summary.SourceOnly +
            "  |  Target Only: " + summary.TargetOnly + Environment.NewLine +
            "Indeterminate rows — Unsupported: " + summary.IndeterminateUnsupported +
            "  |  Unresolved: " + summary.IndeterminateUnresolved + "  |  Ambiguous: " + summary.IndeterminateAmbiguous +
            "  |  Resolved / Indeterminate: " + summary.IndeterminateResolved +
            (summary.IndeterminateOtherOrUnavailable == 0 ? string.Empty :
                "  |  Other / Unavailable: " + summary.IndeterminateOtherOrUnavailable);

        private static string FormatDiagnostics(string side, MembershipEnvironmentDiagnostics diagnostics) =>
            side + " (" + diagnostics.EnvironmentName + ") diagnostics: requests=" + diagnostics.RequestCount +
            ", elapsed=" + diagnostics.Elapsed.TotalSeconds.ToString("0.00", CultureInfo.InvariantCulture) + "s" +
            ", Raw " + side + " Components=" + Number(diagnostics.RawMembershipCount) +
            ", resolved=" + Number(diagnostics.ResolvedCount) +
            ", unsupported=" + Number(diagnostics.UnsupportedCount) +
            ", unresolved=" + Number(diagnostics.UnresolvedCount) +
            ", ambiguous=" + Number(diagnostics.AmbiguousCount);

        private static string Number(int? value) => value.HasValue
            ? value.Value.ToString(CultureInfo.InvariantCulture) : "n/a";
    }
}
