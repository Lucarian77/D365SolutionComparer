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
        private bool showExplorerMetadata;
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

            var viewBar = new ToolStrip { Dock = DockStyle.Top, GripStyle = ToolStripGripStyle.Hidden, Name = "MembershipViewBar" };
            viewSelector = new ToolStripComboBox { Name = "MembershipViewSelector", DropDownStyle = ComboBoxStyle.DropDownList,
                AutoSize = false, Width = 145, AccessibleName = "Membership presentation view" };
            viewSelector.Items.AddRange(new object[] { "Flat View", "Hierarchy View", "Differences Only" }); viewSelector.SelectedIndex = 0;
            membershipFilter = new ToolStripComboBox { Name = "MembershipPresenceFilter", DropDownStyle = ComboBoxStyle.DropDownList,
                AutoSize = false, Width = 165, AccessibleName = "Filter authoritative membership rows" };
            membershipFilter.Items.AddRange(new object[] { "All", "Present in Both", "Source Only", "Target Only", "Indeterminate" });
            membershipFilter.SelectedIndex = 0;
            viewStatus = new ToolStripLabel { Name = "MembershipViewStatus", Text = gridRows.Count + " comparison rows", AutoToolTip = false };
            legacyViewLabel = new ToolStripLabel("View:");
            viewBar.Items.Add(legacyViewLabel); viewBar.Items.Add(viewSelector);
            viewBar.Items.Add(new ToolStripLabel("Membership:")); viewBar.Items.Add(membershipFilter); viewBar.Items.Add(viewStatus);
            explorerButton = new ToolStripButton("Solution Objects Explorer") { Name = "SolutionObjectsExplorerAction",
                CheckOnClick = false, AccessibleName = "Open read-only Solution Objects Explorer" };
            explorerButton.Click += async (sender, args) =>
            {
                if (explorerButton.Checked)
                { RetainExplorerSelection(); explorerButton.Checked = false; viewRequestVersion++; if (viewSelector.SelectedIndex == 1) await ShowHierarchyAsync(); else ShowFlatView(); }
                else { explorerButton.Checked = true; await ShowExplorerAsync(); }
            };
            viewBar.Items.Add(explorerButton);
            ProjectHierarchy = token => completedComparison == null ? ComponentHierarchyProjector.Project(presentation, cancellationToken: token) :
                ComponentHierarchyProjector.Project(completedComparison, completedComparison.SourceHierarchyParents,
                    completedComparison.TargetHierarchyParents, token);
            ProjectExplorer = (hierarchy, token) => SolutionObjectsExplorerProjector.Project(presentation, completedComparison, hierarchy, token);
            bool differencesOnlyFilter = false;
            resultsGrid.CellClick += (sender, args) =>
            { if (args.RowIndex >= 0) retainedSelectionRow = (resultsGrid.Rows[args.RowIndex].DataBoundItem as MembershipGridRow)?.OriginalRowIndex; };
            resultsGrid.KeyUp += (sender, args) =>
            { retainedSelectionRow = (resultsGrid.CurrentRow?.DataBoundItem as MembershipGridRow)?.OriginalRowIndex; };
            viewSelector.SelectedIndexChanged += async (sender, args) =>
            {
                RetainExplorerSelection(); explorerButton.Checked = false;
                if (differencesOnlyFilter != (viewSelector.SelectedIndex == 2))
                { differencesOnlyFilter = viewSelector.SelectedIndex == 2; ApplyMembershipFilter(); }
                if (viewSelector.SelectedIndex == 1) await ShowHierarchyAsync();
                else { viewRequestVersion++; ShowFlatView(); }
            };
            membershipFilter.SelectedIndexChanged += async (sender, args) =>
            {
                ApplyMembershipFilter();
                if (explorerButton.Checked) await ShowExplorerAsync();
                else if (viewSelector.SelectedIndex == 1) await ShowHierarchyAsync();
            };

            Controls.Add(resultsGrid);
            Controls.Add(viewBar);
            Controls.Add(header);
            ResumeLayout(true);
        }

        private MembershipGridRow SelectedComparisonRow() => explorerControl?.Visible == true ? explorerControl.SelectedComparison : hierarchyControl?.Visible == true ?
            hierarchyControl.SelectedComparison : resultsGrid.CurrentRow?.DataBoundItem as MembershipGridRow;

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
            resultsGrid.Visible = true;
            if (selected != null)
                foreach (DataGridViewRow row in resultsGrid.Rows)
                    if ((row.DataBoundItem as MembershipGridRow)?.OriginalRowIndex == selected.OriginalRowIndex)
                    { resultsGrid.CurrentCell = row.Cells[0]; break; }
            displayEvidenceButton.Enabled = SelectedComparisonRow() != null;
            viewStatus.Text = FilteredRows().Count + " comparison rows";
        }

        internal async Task ShowHierarchyAsync()
        {
            if (IsDisposed || Disposing) return;
            legacyViewLabel.Text = "View:";
            if (explorerControl != null) explorerControl.Visible = false;
            int version = ++viewRequestVersion;
            var token = hierarchyLifetime.Token;
            var filtered = FilteredRows(); int filter = membershipFilter.SelectedIndex; bool all = filter == 0;
            try
            {
                viewStatus.Text = "Preparing local hierarchy...";
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
                        Controls.Add(hierarchyControl);
                        Controls.SetChildIndex(hierarchyControl, 0);
                    }
                    var selected = retainedSelectionRow ?? hierarchyControl.SelectedComparison?.OriginalRowIndex;
                    var side = retainedExplorerSide ?? hierarchyControl.SelectedSide;
                    if (!ReferenceEquals(HierarchyProjection, projection) || boundHierarchyFilter != filter)
                    { hierarchyControl.Bind(model); currentHierarchyModel = model; boundHierarchyFilter = filter; }
                    if (selected.HasValue && hierarchyControl.SelectedComparison?.OriginalRowIndex != selected.Value)
                        hierarchyControl.SelectComparison(selected.Value, side);
                    HierarchyProjection = projection;
                    retainedExplorerSide = null;
                    if (viewSelector.SelectedIndex != 1 || version != viewRequestVersion) return;
                    resultsGrid.Visible = false; hierarchyControl.Visible = true;
                    displayEvidenceButton.Enabled = SelectedComparisonRow() != null;
                    viewStatus.Text = model.ComparisonRowCount + " comparison rows; " + model.VisibleOccurrenceCount + " occurrences (including context)";
                }, token).ConfigureAwait(false);
            }
            catch (OperationCanceledException)
            { if (!token.IsCancellationRequested) await ReportHierarchyFailureAsync(null, version, token).ConfigureAwait(false); }
            catch (Exception ex)
            { if (!token.IsCancellationRequested) await ReportHierarchyFailureAsync(ex, version, token).ConfigureAwait(false); }
        }

        private void RetainExplorerSelection()
        {
            if (explorerControl?.Visible == true && explorerControl.SelectedComparison != null)
            { retainedSelectionRow = explorerControl.SelectedComparison.OriginalRowIndex; retainedExplorerSide = explorerControl.SelectedSide; }
        }
        internal async Task ShowExplorerAsync()
        {
            if (IsDisposed || Disposing || !explorerButton.Checked) return;
            int version = ++viewRequestVersion; var token = hierarchyLifetime.Token;
            var selected = SelectedComparisonRow();
            if (selected != null) retainedSelectionRow = selected.OriginalRowIndex;
            try
            {
                viewStatus.Text = "Preparing local object navigation...";
                if (hierarchyTask == null) hierarchyTask = Task.Run(() => ProjectHierarchy(token), token);
                var hierarchy = await hierarchyTask.ConfigureAwait(false); token.ThrowIfCancellationRequested();
                if (explorerTask == null) explorerTask = Task.Run(() => ProjectExplorer(hierarchy, token), token);
                var projection = await explorerTask.ConfigureAwait(false); token.ThrowIfCancellationRequested();
                await OnHierarchyUiAsync(() =>
                {
                    if (!explorerButton.Checked || version != viewRequestVersion) return;
                    if (explorerControl == null)
                    {
                        explorerControl = new SolutionObjectsExplorerControl(FormatDisplayEvidenceDetails, row =>
                        {
                            retainedSelectionRow = row.OriginalRowIndex; explorerButton.Checked = false; viewSelector.SelectedIndex = 0; ShowFlatView();
                            foreach (DataGridViewRow native in resultsGrid.Rows)
                                if ((native.DataBoundItem as MembershipGridRow)?.OriginalRowIndex == row.OriginalRowIndex)
                                { resultsGrid.CurrentCell = native.Cells[0]; break; }
                        }, OpenDefinitionDetails) { Visible = false };
                        explorerControl.SelectedComparisonChanged = row => displayEvidenceButton.Enabled = row != null;
                        explorerControl.RenderingFailed = ExplorerFailed; Controls.Add(explorerControl); Controls.SetChildIndex(explorerControl, 0);
                        explorerControl.Bind(projection, FilteredRows().Select(r => r.OriginalRowIndex), membershipFilter.SelectedIndex == 0 && viewSelector.SelectedIndex != 2);
                        boundExplorerFilter = membershipFilter.SelectedIndex; boundExplorerDifferences = viewSelector.SelectedIndex == 2;
                    }
                    else if (boundExplorerFilter != membershipFilter.SelectedIndex || boundExplorerDifferences != (viewSelector.SelectedIndex == 2))
                    {
                        explorerControl.ApplyFilter(FilteredRows().Select(r => r.OriginalRowIndex), membershipFilter.SelectedIndex == 0 && viewSelector.SelectedIndex != 2);
                        boundExplorerFilter = membershipFilter.SelectedIndex; boundExplorerDifferences = viewSelector.SelectedIndex == 2;
                    }
                    explorerControl.ShowMetadata(showExplorerMetadata);
                    if (retainedSelectionRow.HasValue) explorerControl.SelectComparison(retainedSelectionRow.Value, retainedExplorerSide ?? hierarchyControl?.SelectedSide);
                    if (version != viewRequestVersion || !explorerButton.Checked) return;
                    ExplorerProjection = projection; HierarchyProjection = hierarchy;
                    legacyViewLabel.Text = "Details scope:";
                    resultsGrid.Visible = false; if (hierarchyControl != null) hierarchyControl.Visible = false;
                    explorerControl.Visible = true; displayEvidenceButton.Enabled = explorerControl.SelectedComparison != null;
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
            viewSelector.SelectedIndex = 0; ShowFlatView(); viewStatus.Text = diagnostic;
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
