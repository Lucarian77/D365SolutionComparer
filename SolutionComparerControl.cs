using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using ClosedXML.Excel;
using XrmToolBox.Extensibility;
using McTools.Xrm.Connection;
using D365SolutionComparer.Services;
using ModelSolutionInfo = D365SolutionComparer.Models.SolutionInfo;
using OrgService = Microsoft.Xrm.Sdk.IOrganizationService;
using CompareResult = D365SolutionComparer.Models.SolutionCompareResult;

namespace D365SolutionComparer
{
    public class SolutionComparerControl : MultipleConnectionsPluginControlBase
    {
        private Label lblTitle;
        private FlowLayoutPanel topPanel;
        private Button btnLoadSource;
        private Button btnConnectTarget;
        private Button btnLoadTarget;
        private Button btnCompare;
        private Button btnExportCsv;
        private Button btnFilter;
        private ContextMenuStrip filterMenu;
        private ToolStripMenuItem miAll;
        private ToolStripMenuItem miMatch;
        private ToolStripMenuItem miVersionMismatch;
        private ToolStripMenuItem miPublisherMismatch;
        private ToolStripMenuItem miDisplayNameMismatch;
        private ToolStripMenuItem miPackageTypeDifference;
        private ToolStripMenuItem miMultipleDifferences;
        private ToolStripMenuItem miMissingInSource;
        private ToolStripMenuItem miMissingInTarget;
        private CheckBox chkPackageTypeMismatchOnly;

        private DataGridView dgvResults;
        private Label lblSourceEnv;
        private Label lblTargetEnv;
        private Label lblStatusMessage;
        private Label lblSummary;
        private Label lblLegend;

        private ConnectionDetail targetConnectionDetail;
        private OrgService targetService;

        private string sourceConnectionName = "Current XrmToolBox connection";
        private string targetConnectionName = "Not connected";

        private bool sourceLoaded;
        private bool targetLoaded;

        private List<ModelSolutionInfo> sourceSolutions = new List<ModelSolutionInfo>();
        private List<ModelSolutionInfo> targetSolutions = new List<ModelSolutionInfo>();
        private List<CompareResult> comparisonResults = new List<CompareResult>();

        public SolutionComparerControl()
        {
            Dock = DockStyle.Fill;
            BackColor = Color.White;
            BuildUi();
        }

        private void BuildUi()
        {
            Controls.Clear();

            lblTitle = new Label
            {
                Text = "D365 Solution Comparer",
                Dock = DockStyle.Top,
                Height = 40,
                Font = new Font("Segoe UI", 12F, FontStyle.Bold),
                Padding = new Padding(10, 8, 0, 0),
                BackColor = Color.White
            };

            topPanel = new FlowLayoutPanel
            {
                Dock = DockStyle.Top,
                Height = 45,
                Padding = new Padding(10, 5, 10, 5),
                BackColor = Color.White,
                AutoScroll = true,
                WrapContents = false
            };

            btnLoadSource = new Button { Text = "Load Source", Width = 110, Height = 30 };
            btnConnectTarget = new Button { Text = "Connect Target", Width = 120, Height = 30 };
            btnLoadTarget = new Button { Text = "Load Target", Width = 110, Height = 30 };
            btnCompare = new Button { Text = "Compare", Width = 110, Height = 30 };
            btnExportCsv = new Button { Text = "Export Excel", Width = 110, Height = 30 };
            btnFilter = new Button { Text = "Filter: All", Width = 220, Height = 30 };

            chkPackageTypeMismatchOnly = new CheckBox
            {
                Text = "Show only managed/unmanaged differences",
                AutoSize = true,
                Height = 30,
                Margin = new Padding(10, 6, 0, 0),
                BackColor = Color.White
            };

            BuildFilterMenu();

            btnLoadSource.Click += BtnLoadSource_Click;
            btnConnectTarget.Click += BtnConnectTarget_Click;
            btnLoadTarget.Click += BtnLoadTarget_Click;
            btnCompare.Click += BtnCompare_Click;
            btnExportCsv.Click += BtnExportCsv_Click;
            btnFilter.Click += BtnFilter_Click;
            chkPackageTypeMismatchOnly.CheckedChanged += ChkPackageTypeMismatchOnly_CheckedChanged;

            topPanel.Controls.Add(btnLoadSource);
            topPanel.Controls.Add(btnConnectTarget);
            topPanel.Controls.Add(btnLoadTarget);
            topPanel.Controls.Add(btnCompare);
            topPanel.Controls.Add(btnExportCsv);
            topPanel.Controls.Add(btnFilter);
            topPanel.Controls.Add(chkPackageTypeMismatchOnly);

            lblSourceEnv = new Label
            {
                Text = "Source: Current XrmToolBox connection",
                Dock = DockStyle.Top,
                Height = 24,
                Padding = new Padding(10, 0, 0, 0),
                BackColor = Color.White
            };

            lblTargetEnv = new Label
            {
                Text = "Target: Not connected",
                Dock = DockStyle.Top,
                Height = 24,
                Padding = new Padding(10, 0, 0, 0),
                BackColor = Color.White
            };

            lblStatusMessage = new Label
            {
                Text = "Status: Ready",
                Dock = DockStyle.Top,
                Height = 24,
                Padding = new Padding(10, 0, 0, 0),
                BackColor = Color.White,
                ForeColor = Color.Green,
                Font = new Font("Segoe UI", 9F, FontStyle.Bold)
            };

            lblSummary = new Label
            {
                Text = "Summary: No comparison results yet",
                Dock = DockStyle.Top,
                Height = 44,
                Padding = new Padding(10, 4, 10, 4),
                BackColor = Color.White,
                AutoEllipsis = false
            };

            lblLegend = new Label
            {
                Text = "Legend: Match=Green | Version=Orange | Publisher=Purple | Display Name=Blue | Package Type=Teal | Multiple=Magenta | Missing in Source=Red | Missing in Target=Brick Red",
                Dock = DockStyle.Top,
                Height = 32,
                Padding = new Padding(10, 0, 0, 0),
                BackColor = Color.White,
                ForeColor = Color.DimGray,
                Font = new Font("Segoe UI", 8.5F, FontStyle.Italic)
            };

            dgvResults = new DataGridView
            {
                Dock = DockStyle.Fill,
                ReadOnly = true,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                AutoGenerateColumns = true,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                BackgroundColor = Color.White,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                MultiSelect = false
            };

            dgvResults.CellFormatting += DgvResults_CellFormatting;

            Controls.Add(dgvResults);
            Controls.Add(lblLegend);
            Controls.Add(lblSummary);
            Controls.Add(lblStatusMessage);
            Controls.Add(lblTargetEnv);
            Controls.Add(lblSourceEnv);
            Controls.Add(topPanel);
            Controls.Add(lblTitle);
        }

        private void BuildFilterMenu()
        {
            filterMenu = new ContextMenuStrip();

            miAll = CreateFilterItem("All", true);
            miMatch = CreateFilterItem("Match");
            miVersionMismatch = CreateFilterItem("Version Mismatch");
            miPublisherMismatch = CreateFilterItem("Publisher Mismatch");
            miDisplayNameMismatch = CreateFilterItem("Display Name Mismatch");
            miPackageTypeDifference = CreateFilterItem("Package Type Differences");
            miMultipleDifferences = CreateFilterItem("Multiple Differences");
            miMissingInSource = CreateFilterItem("Missing in Source");
            miMissingInTarget = CreateFilterItem("Missing in Target");

            filterMenu.Items.AddRange(new ToolStripItem[]
            {
                miAll,
                new ToolStripSeparator(),
                miMatch,
                miVersionMismatch,
                miPublisherMismatch,
                miDisplayNameMismatch,
                miPackageTypeDifference,
                miMultipleDifferences,
                miMissingInSource,
                miMissingInTarget
            });
        }

        private ToolStripMenuItem CreateFilterItem(string text, bool isChecked = false)
        {
            var item = new ToolStripMenuItem(text)
            {
                CheckOnClick = true,
                Checked = isChecked
            };

            item.CheckedChanged += FilterItem_CheckedChanged;
            return item;
        }

        public override void UpdateConnection(OrgService newService, ConnectionDetail detail, string actionName, object parameter)
        {
            base.UpdateConnection(newService, detail, actionName, parameter);

            if (!sourceLoaded)
            {
                if (detail != null && !string.IsNullOrWhiteSpace(detail.ConnectionName))
                {
                    sourceConnectionName = detail.ConnectionName;
                }
                else
                {
                    sourceConnectionName = "Current XrmToolBox connection";
                }
            }

            RefreshEnvironmentLabels();
            SetStatusMessage("Ready", Color.Green);
        }

        protected override void ConnectionDetailsUpdated(NotifyCollectionChangedEventArgs e)
        {
            if (e.Action == NotifyCollectionChangedAction.Add && e.NewItems != null && e.NewItems.Count > 0)
            {
                var addedConnection = e.NewItems[0] as ConnectionDetail;

                if (addedConnection != null)
                {
                    targetConnectionDetail = addedConnection;
                    targetService = addedConnection.GetCrmServiceClient();
                    targetSolutions = new List<ModelSolutionInfo>();
                    targetLoaded = false;

                    targetConnectionName = !string.IsNullOrWhiteSpace(addedConnection.ConnectionName)
                        ? addedConnection.ConnectionName
                        : "Target environment";

                    comparisonResults = new List<CompareResult>();
                    dgvResults.DataSource = null;
                    lblSummary.Text = "Summary: No comparison results yet";

                    RefreshEnvironmentLabels();
                    SetStatusMessage("Target environment connected successfully.", Color.Green);
                }
            }
            else if (e.Action == NotifyCollectionChangedAction.Remove)
            {
                if (targetConnectionDetail != null && !AdditionalConnectionDetails.Contains(targetConnectionDetail))
                {
                    targetConnectionDetail = null;
                    targetService = null;
                    targetSolutions = new List<ModelSolutionInfo>();
                    targetLoaded = false;
                    targetConnectionName = "Not connected";

                    comparisonResults = new List<CompareResult>();
                    dgvResults.DataSource = null;
                    lblSummary.Text = "Summary: No comparison results yet";

                    RefreshEnvironmentLabels();
                    SetStatusMessage("Target connection removed.", Color.DarkOrange);
                }
            }
        }

        private void RefreshEnvironmentLabels()
        {
            lblSourceEnv.Text = sourceLoaded && sourceSolutions.Count > 0
                ? $"Source: {sourceConnectionName} ({sourceSolutions.Count} solutions)"
                : "Source: " + sourceConnectionName;

            lblTargetEnv.Text = targetLoaded && targetSolutions.Count > 0
                ? $"Target: {targetConnectionName} ({targetSolutions.Count} solutions)"
                : "Target: " + targetConnectionName;
        }

        private void BtnLoadSource_Click(object sender, EventArgs e)
        {
            if (Service == null)
            {
                SetStatusMessage("Please connect to a Dataverse environment first.", Color.DarkOrange);

                MessageBox.Show(
                    "Please connect to a Dataverse environment first.",
                    "Load Source",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
                return;
            }

            try
            {
                var dataverseService = new DataverseSolutionService();
                sourceSolutions = dataverseService.GetSolutions(Service);

                comparisonResults = new List<CompareResult>();
                sourceLoaded = true;

                ResetFilterSelection();
                chkPackageTypeMismatchOnly.Checked = false;

                sourceConnectionName = (ConnectionDetail != null && !string.IsNullOrWhiteSpace(ConnectionDetail.ConnectionName))
                    ? ConnectionDetail.ConnectionName
                    : "Current XrmToolBox connection";

                dgvResults.DataSource = null;
                dgvResults.DataSource = sourceSolutions;

                RefreshEnvironmentLabels();
                lblSummary.Text = "Summary: No comparison results yet";

                ApplySolutionListGridLayout();
                ResetGridScrollPosition();

                SetStatusMessage($"Loaded {sourceSolutions.Count} source solutions from the source environment.", Color.Green);
            }
            catch (Exception ex)
            {
                SetStatusMessage("Failed to load source solutions.", Color.Red);

                MessageBox.Show(
                    "Failed to load source solutions.\n\n" + ex.Message,
                    "Load Source",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        private void BtnConnectTarget_Click(object sender, EventArgs e)
        {
            try
            {
                if (targetConnectionDetail != null)
                {
                    RemoveAdditionalOrganization(targetConnectionDetail);
                    targetConnectionDetail = null;
                    targetService = null;
                    targetSolutions = new List<ModelSolutionInfo>();
                    targetLoaded = false;
                    targetConnectionName = "Not connected";
                    comparisonResults = new List<CompareResult>();
                    dgvResults.DataSource = null;
                    lblSummary.Text = "Summary: No comparison results yet";
                    RefreshEnvironmentLabels();
                }

                AddAdditionalOrganization();
            }
            catch (Exception ex)
            {
                SetStatusMessage("Failed to connect target environment.", Color.Red);

                MessageBox.Show(
                    "Failed to connect target environment.\n\n" + ex.Message,
                    "Connect Target",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        private void BtnLoadTarget_Click(object sender, EventArgs e)
        {
            if (targetService == null)
            {
                SetStatusMessage("Please connect a target Dataverse environment first.", Color.DarkOrange);

                MessageBox.Show(
                    "Please connect a target Dataverse environment first.",
                    "Load Target",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
                return;
            }

            try
            {
                var dataverseService = new DataverseSolutionService();
                targetSolutions = dataverseService.GetSolutions(targetService);

                comparisonResults = new List<CompareResult>();
                targetLoaded = true;

                ResetFilterSelection();
                chkPackageTypeMismatchOnly.Checked = false;

                targetConnectionName = (targetConnectionDetail != null && !string.IsNullOrWhiteSpace(targetConnectionDetail.ConnectionName))
                    ? targetConnectionDetail.ConnectionName
                    : "Target environment";

                dgvResults.DataSource = null;
                dgvResults.DataSource = targetSolutions;

                RefreshEnvironmentLabels();
                lblSummary.Text = "Summary: No comparison results yet";

                ApplySolutionListGridLayout();
                ResetGridScrollPosition();

                SetStatusMessage($"Loaded {targetSolutions.Count} target solutions from the target environment.", Color.Green);
            }
            catch (Exception ex)
            {
                SetStatusMessage("Failed to load target solutions.", Color.Red);

                MessageBox.Show(
                    "Failed to load target solutions.\n\n" + ex.Message,
                    "Load Target",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        private void BtnCompare_Click(object sender, EventArgs e)
        {
            if (sourceSolutions.Count == 0)
            {
                SetStatusMessage("Please load source solutions first.", Color.DarkOrange);

                MessageBox.Show(
                    "Please load source solutions first.",
                    "Compare",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
                return;
            }

            if (targetSolutions.Count == 0)
            {
                SetStatusMessage("Please load target solutions first.", Color.DarkOrange);

                MessageBox.Show(
                    "Please load target solutions first.",
                    "Compare",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
                return;
            }

            try
            {
                var comparisonService = new SolutionComparisonService();
                comparisonResults = comparisonService.Compare(sourceSolutions, targetSolutions);

                ResetFilterSelection();
                chkPackageTypeMismatchOnly.Checked = false;
                BindFilteredResults();

                SetStatusMessage("Comparison completed successfully.", Color.Green);
            }
            catch (Exception ex)
            {
                SetStatusMessage("Failed to compare solutions.", Color.Red);

                MessageBox.Show(
                    "Failed to compare solutions.\n\n" + ex.Message,
                    "Compare",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        private void BtnFilter_Click(object sender, EventArgs e)
        {
            filterMenu.Show(btnFilter, new Point(0, btnFilter.Height));
        }

        private void ChkPackageTypeMismatchOnly_CheckedChanged(object sender, EventArgs e)
        {
            if (HasComparisonResults())
            {
                BindFilteredResults();
            }
        }

        private void FilterItem_CheckedChanged(object sender, EventArgs e)
        {
            if (!(sender is ToolStripMenuItem changedItem))
                return;

            changedItem.CheckedChanged -= FilterItem_CheckedChanged;

            try
            {
                if (changedItem == miAll)
                {
                    if (miAll.Checked)
                    {
                        SetNonAllItemsChecked(false);
                    }
                    else if (!AnySpecificFilterChecked())
                    {
                        miAll.Checked = true;
                    }
                }
                else
                {
                    if (changedItem.Checked)
                    {
                        miAll.CheckedChanged -= FilterItem_CheckedChanged;
                        miAll.Checked = false;
                        miAll.CheckedChanged += FilterItem_CheckedChanged;
                    }
                    else if (!AnySpecificFilterChecked())
                    {
                        miAll.CheckedChanged -= FilterItem_CheckedChanged;
                        miAll.Checked = true;
                        miAll.CheckedChanged += FilterItem_CheckedChanged;
                    }
                }
            }
            finally
            {
                changedItem.CheckedChanged += FilterItem_CheckedChanged;
            }

            UpdateFilterButtonText();

            if (HasComparisonResults())
            {
                BindFilteredResults();
            }
        }

        private void SetNonAllItemsChecked(bool isChecked)
        {
            var items = new[]
            {
                miMatch,
                miVersionMismatch,
                miPublisherMismatch,
                miDisplayNameMismatch,
                miPackageTypeDifference,
                miMultipleDifferences,
                miMissingInSource,
                miMissingInTarget
            };

            foreach (var item in items)
            {
                item.CheckedChanged -= FilterItem_CheckedChanged;
                item.Checked = isChecked;
                item.CheckedChanged += FilterItem_CheckedChanged;
            }
        }

        private bool AnySpecificFilterChecked()
        {
            return miMatch.Checked
                   || miVersionMismatch.Checked
                   || miPublisherMismatch.Checked
                   || miDisplayNameMismatch.Checked
                   || miPackageTypeDifference.Checked
                   || miMultipleDifferences.Checked
                   || miMissingInSource.Checked
                   || miMissingInTarget.Checked;
        }

        private List<string> GetSelectedStatuses()
        {
            if (miAll.Checked || !AnySpecificFilterChecked())
            {
                return new List<string>();
            }

            var selected = new List<string>();

            if (miMatch.Checked) selected.Add("Match");
            if (miVersionMismatch.Checked) selected.Add("Version Mismatch");
            if (miPublisherMismatch.Checked) selected.Add("Publisher Mismatch");
            if (miDisplayNameMismatch.Checked) selected.Add("Display Name Mismatch");
            if (miMultipleDifferences.Checked) selected.Add("Multiple Differences");
            if (miMissingInSource.Checked) selected.Add("Missing in Source");
            if (miMissingInTarget.Checked) selected.Add("Missing in Target");

            return selected;
        }

        private void UpdateFilterButtonText()
        {
            var labels = new List<string>();

            if (miMatch.Checked) labels.Add("Match");
            if (miVersionMismatch.Checked) labels.Add("Version Mismatch");
            if (miPublisherMismatch.Checked) labels.Add("Publisher Mismatch");
            if (miDisplayNameMismatch.Checked) labels.Add("Display Name Mismatch");
            if (miPackageTypeDifference.Checked) labels.Add("Package Type Differences");
            if (miMultipleDifferences.Checked) labels.Add("Multiple Differences");
            if (miMissingInSource.Checked) labels.Add("Missing in Source");
            if (miMissingInTarget.Checked) labels.Add("Missing in Target");

            if (miAll.Checked || labels.Count == 0)
            {
                btnFilter.Text = "Filter: All";
            }
            else if (labels.Count == 1)
            {
                btnFilter.Text = "Filter: " + labels[0];
            }
            else if (labels.Count == 2)
            {
                btnFilter.Text = $"Filter: {labels[0]}, {labels[1]}";
            }
            else
            {
                btnFilter.Text = $"Filter: {labels[0]} + {labels.Count - 1} more";
            }
        }

        private void ResetFilterSelection()
        {
            miAll.CheckedChanged -= FilterItem_CheckedChanged;
            miAll.Checked = true;
            miAll.CheckedChanged += FilterItem_CheckedChanged;

            SetNonAllItemsChecked(false);
            UpdateFilterButtonText();
        }

        private void BindFilteredResults()
        {
            var selectedStatuses = GetSelectedStatuses();

            IEnumerable<CompareResult> filteredResults = comparisonResults ?? new List<CompareResult>();

            if (selectedStatuses.Count > 0)
            {
                filteredResults = filteredResults.Where(r => selectedStatuses.Contains(r.Status ?? string.Empty));
            }

            if (miPackageTypeDifference.Checked && !miAll.Checked)
            {
                filteredResults = filteredResults.Where(IsAnyPackageTypeDifference);
            }

            if (chkPackageTypeMismatchOnly != null && chkPackageTypeMismatchOnly.Checked)
            {
                filteredResults = filteredResults.Where(IsManagedUnmanagedDifference);
            }

            var finalResults = filteredResults.ToList();

            dgvResults.DataSource = null;
            dgvResults.DataSource = finalResults;

            ApplyComparisonGridLayout();
            ResetGridScrollPosition();
            UpdateSummary(finalResults);
        }

        private bool HasComparisonResults()
        {
            return comparisonResults != null && comparisonResults.Count > 0;
        }

        private void ApplySolutionListGridLayout()
        {
            dgvResults.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None;

            HideColumnIfExists("IsManaged");

            if (dgvResults.Columns["UniqueName"] != null)
            {
                dgvResults.Columns["UniqueName"].Width = 260;
                dgvResults.Columns["UniqueName"].HeaderText = "Unique Name";
                dgvResults.Columns["UniqueName"].DisplayIndex = 0;
            }

            if (dgvResults.Columns["DisplayName"] != null)
            {
                dgvResults.Columns["DisplayName"].Width = 300;
                dgvResults.Columns["DisplayName"].HeaderText = "Display Name";
                dgvResults.Columns["DisplayName"].DisplayIndex = 1;
            }

            if (dgvResults.Columns["Version"] != null)
            {
                dgvResults.Columns["Version"].Width = 120;
                dgvResults.Columns["Version"].HeaderText = "Version";
                dgvResults.Columns["Version"].DisplayIndex = 2;
            }

            if (dgvResults.Columns["Publisher"] != null)
            {
                dgvResults.Columns["Publisher"].Width = 260;
                dgvResults.Columns["Publisher"].HeaderText = "Publisher";
                dgvResults.Columns["Publisher"].DisplayIndex = 3;
            }

            if (dgvResults.Columns["PackageType"] != null)
            {
                dgvResults.Columns["PackageType"].Visible = true;
                dgvResults.Columns["PackageType"].Width = 140;
                dgvResults.Columns["PackageType"].HeaderText = "Package Type";
                dgvResults.Columns["PackageType"].DisplayIndex = 4;
            }

            foreach (DataGridViewColumn column in dgvResults.Columns)
            {
                if (column.Visible)
                {
                    column.SortMode = DataGridViewColumnSortMode.Automatic;
                }
            }
        }

        private void ApplyComparisonGridLayout()
        {
            dgvResults.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None;

            HideColumnIfExists("IsManagedUnmanagedMismatch");
            HideColumnIfExists("IsPackageTypeMismatch");

            if (dgvResults.Columns["UniqueName"] != null)
            {
                dgvResults.Columns["UniqueName"].Width = 220;
                dgvResults.Columns["UniqueName"].HeaderText = "Solution Unique Name";
                dgvResults.Columns["UniqueName"].DisplayIndex = 0;
            }

            if (dgvResults.Columns["SourceDisplayName"] != null)
            {
                dgvResults.Columns["SourceDisplayName"].Width = 220;
                dgvResults.Columns["SourceDisplayName"].HeaderText = "Source Display Name";
                dgvResults.Columns["SourceDisplayName"].DisplayIndex = 1;
            }

            if (dgvResults.Columns["TargetDisplayName"] != null)
            {
                dgvResults.Columns["TargetDisplayName"].Width = 220;
                dgvResults.Columns["TargetDisplayName"].HeaderText = "Target Display Name";
                dgvResults.Columns["TargetDisplayName"].DisplayIndex = 2;
            }

            if (dgvResults.Columns["SourceVersion"] != null)
            {
                dgvResults.Columns["SourceVersion"].Width = 120;
                dgvResults.Columns["SourceVersion"].HeaderText = "Source Version";
                dgvResults.Columns["SourceVersion"].DisplayIndex = 3;
            }

            if (dgvResults.Columns["TargetVersion"] != null)
            {
                dgvResults.Columns["TargetVersion"].Width = 120;
                dgvResults.Columns["TargetVersion"].HeaderText = "Target Version";
                dgvResults.Columns["TargetVersion"].DisplayIndex = 4;
            }

            if (dgvResults.Columns["SourcePublisher"] != null)
            {
                dgvResults.Columns["SourcePublisher"].Width = 180;
                dgvResults.Columns["SourcePublisher"].HeaderText = "Source Publisher";
                dgvResults.Columns["SourcePublisher"].DisplayIndex = 5;
            }

            if (dgvResults.Columns["TargetPublisher"] != null)
            {
                dgvResults.Columns["TargetPublisher"].Width = 180;
                dgvResults.Columns["TargetPublisher"].HeaderText = "Target Publisher";
                dgvResults.Columns["TargetPublisher"].DisplayIndex = 6;
            }

            if (dgvResults.Columns["SourcePackageType"] != null)
            {
                dgvResults.Columns["SourcePackageType"].Width = 140;
                dgvResults.Columns["SourcePackageType"].HeaderText = "Source Package Type";
                dgvResults.Columns["SourcePackageType"].DisplayIndex = 7;
            }

            if (dgvResults.Columns["TargetPackageType"] != null)
            {
                dgvResults.Columns["TargetPackageType"].Width = 140;
                dgvResults.Columns["TargetPackageType"].HeaderText = "Target Package Type";
                dgvResults.Columns["TargetPackageType"].DisplayIndex = 8;
            }

            if (dgvResults.Columns["PackageTypeStatus"] != null)
            {
                dgvResults.Columns["PackageTypeStatus"].Width = 170;
                dgvResults.Columns["PackageTypeStatus"].HeaderText = "Package Type Status";
                dgvResults.Columns["PackageTypeStatus"].DisplayIndex = 9;
            }

            if (dgvResults.Columns["Status"] != null)
            {
                dgvResults.Columns["Status"].Width = 170;
                dgvResults.Columns["Status"].HeaderText = "Overall Status";
                dgvResults.Columns["Status"].DisplayIndex = 10;
            }

            foreach (DataGridViewColumn column in dgvResults.Columns)
            {
                if (column.Visible)
                {
                    column.SortMode = DataGridViewColumnSortMode.Automatic;
                }
            }
        }

        private void HideColumnIfExists(string columnName)
        {
            if (dgvResults.Columns[columnName] != null)
            {
                dgvResults.Columns[columnName].Visible = false;
            }
        }

        private void ResetGridScrollPosition()
        {
            if (dgvResults.Rows.Count == 0 || dgvResults.Columns.Count == 0)
                return;

            try
            {
                dgvResults.ClearSelection();

                var firstVisibleColumn = dgvResults.Columns
                    .Cast<DataGridViewColumn>()
                    .Where(c => c.Visible)
                    .OrderBy(c => c.DisplayIndex)
                    .FirstOrDefault();

                if (firstVisibleColumn != null)
                {
                    dgvResults.FirstDisplayedScrollingColumnIndex = firstVisibleColumn.Index;

                    if (dgvResults.Rows.Count > 0)
                    {
                        dgvResults.CurrentCell = dgvResults.Rows[0].Cells[firstVisibleColumn.Index];
                    }
                }

                dgvResults.FirstDisplayedScrollingRowIndex = 0;
            }
            catch
            {
            }
        }

        private void UpdateSummary(List<CompareResult> results)
        {
            if (results == null || results.Count == 0)
            {
                lblSummary.Text = "Summary: No comparison results";
                return;
            }

            int total = results.Count;
            int match = results.Count(r => string.Equals(r.Status, "Match", StringComparison.OrdinalIgnoreCase));
            int versionMismatch = results.Count(r => string.Equals(r.Status, "Version Mismatch", StringComparison.OrdinalIgnoreCase));
            int publisherMismatch = results.Count(r => string.Equals(r.Status, "Publisher Mismatch", StringComparison.OrdinalIgnoreCase));
            int displayNameMismatch = results.Count(r => string.Equals(r.Status, "Display Name Mismatch", StringComparison.OrdinalIgnoreCase));
            int packageTypeDifference = results.Count(IsAnyPackageTypeDifference);
            int managedUnmanagedDifference = results.Count(IsManagedUnmanagedDifference);
            int multipleDifferences = results.Count(r => string.Equals(r.Status, "Multiple Differences", StringComparison.OrdinalIgnoreCase));
            int missingInSource = results.Count(r => string.Equals(r.Status, "Missing in Source", StringComparison.OrdinalIgnoreCase));
            int missingInTarget = results.Count(r => string.Equals(r.Status, "Missing in Target", StringComparison.OrdinalIgnoreCase));

            lblSummary.Text =
                $"Summary: Total={total} | Match={match} | Version={versionMismatch} | Publisher={publisherMismatch} | Display Name={displayNameMismatch}\r\n" +
                $"Package Type Differences={packageTypeDifference} | Managed/Unmanaged Differences={managedUnmanagedDifference} | Multiple={multipleDifferences} | Missing in Source={missingInSource} | Missing in Target={missingInTarget}";
        }

        private bool IsAnyPackageTypeDifference(CompareResult result)
        {
            if (result == null)
                return false;

            if (IsManagedUnmanagedDifference(result))
                return true;

            var packageTypeStatus = (result.PackageTypeStatus ?? string.Empty).Trim();

            if (string.Equals(packageTypeStatus, "Package Type Mismatch", StringComparison.OrdinalIgnoreCase))
                return true;

            var sourceType = (result.SourcePackageType ?? string.Empty).Trim();
            var targetType = (result.TargetPackageType ?? string.Empty).Trim();

            if (string.IsNullOrWhiteSpace(sourceType) || string.IsNullOrWhiteSpace(targetType))
                return false;

            return !string.Equals(sourceType, targetType, StringComparison.OrdinalIgnoreCase);
        }

        private bool IsManagedUnmanagedDifference(CompareResult result)
        {
            if (result == null)
                return false;

            var packageTypeStatus = (result.PackageTypeStatus ?? string.Empty).Trim();

            if (string.Equals(packageTypeStatus, "Managed/Unmanaged Mismatch", StringComparison.OrdinalIgnoreCase))
                return true;

            var sourceType = (result.SourcePackageType ?? string.Empty).Trim();
            var targetType = (result.TargetPackageType ?? string.Empty).Trim();

            if (string.IsNullOrWhiteSpace(sourceType) || string.IsNullOrWhiteSpace(targetType))
                return false;

            bool sourceManagedState = IsManagedOrUnmanagedLabel(sourceType);
            bool targetManagedState = IsManagedOrUnmanagedLabel(targetType);

            if (!sourceManagedState || !targetManagedState)
                return false;

            return !string.Equals(sourceType, targetType, StringComparison.OrdinalIgnoreCase);
        }

        private bool IsManagedOrUnmanagedLabel(string value)
        {
            return string.Equals(value, "Managed", StringComparison.OrdinalIgnoreCase)
                   || string.Equals(value, "Unmanaged", StringComparison.OrdinalIgnoreCase);
        }

        private void SetStatusMessage(string message, Color color)
        {
            lblStatusMessage.Text = "Status: " + message;
            lblStatusMessage.ForeColor = color;
        }

        private void DgvResults_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            if (e.RowIndex < 0 || e.ColumnIndex < 0)
                return;

            var column = dgvResults.Columns[e.ColumnIndex];
            var dataPropertyName = column.DataPropertyName;

            var rowData = dgvResults.Rows[e.RowIndex].DataBoundItem as CompareResult;
            bool packageTypeDifference = IsAnyPackageTypeDifference(rowData);

            if (e.Value == null)
                return;

            var style = dgvResults.Rows[e.RowIndex].Cells[e.ColumnIndex].Style;
            style.SelectionBackColor = SystemColors.Highlight;
            style.SelectionForeColor = SystemColors.HighlightText;

            if (dataPropertyName == "SourcePackageType" || dataPropertyName == "TargetPackageType")
            {
                if (packageTypeDifference)
                {
                    style.Font = new Font(dgvResults.Font, FontStyle.Bold);
                    style.ForeColor = Color.Teal;
                    style.BackColor = Color.LightCyan;
                }

                return;
            }

            if (dataPropertyName != "Status" && dataPropertyName != "PackageTypeStatus")
                return;

            var status = e.Value.ToString();
            if (string.IsNullOrWhiteSpace(status))
                return;

            style.Font = new Font(dgvResults.Font, FontStyle.Bold);

            switch (status)
            {
                case "Match":
                    style.ForeColor = Color.Green;
                    style.BackColor = Color.Honeydew;
                    break;

                case "Version Mismatch":
                    style.ForeColor = Color.DarkOrange;
                    style.BackColor = Color.Moccasin;
                    break;

                case "Publisher Mismatch":
                    style.ForeColor = Color.DarkViolet;
                    style.BackColor = Color.Lavender;
                    break;

                case "Display Name Mismatch":
                    style.ForeColor = Color.SteelBlue;
                    style.BackColor = Color.AliceBlue;
                    break;

                case "Package Type Mismatch":
                case "Managed/Unmanaged Mismatch":
                    style.ForeColor = Color.Teal;
                    style.BackColor = Color.LightCyan;
                    break;

                case "Multiple Differences":
                    style.ForeColor = Color.DarkMagenta;
                    style.BackColor = Color.MistyRose;
                    break;

                case "Missing in Source":
                    style.ForeColor = Color.Red;
                    style.BackColor = Color.MistyRose;
                    break;

                case "Missing in Target":
                    style.ForeColor = Color.Firebrick;
                    style.BackColor = Color.Linen;
                    break;

                default:
                    style.ForeColor = dgvResults.ForeColor;
                    style.BackColor = Color.White;
                    break;
            }
        }

        private void BtnExportCsv_Click(object sender, EventArgs e)
        {
            if (!HasComparisonResults() || dgvResults.DataSource == null || dgvResults.Rows.Count == 0)
            {
                SetStatusMessage("There is no comparison data to export.", Color.DarkOrange);

                MessageBox.Show(
                    "There is no comparison data to export.",
                    "Export Excel",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
                return;
            }

            using (var saveDialog = new SaveFileDialog())
            {
                saveDialog.Filter = "Excel Workbook (*.xlsx)|*.xlsx";
                saveDialog.FileName = $"D365SolutionComparer_Source_vs_Target_{DateTime.Now:yyyyMMdd_HHmm}.xlsx";

                if (saveDialog.ShowDialog() != DialogResult.OK)
                {
                    SetStatusMessage("Export cancelled.", Color.DarkOrange);
                    return;
                }

                try
                {
                    using (var workbook = new XLWorkbook())
                    {
                        var worksheet = workbook.Worksheets.Add("Comparison Results");
                        ExportGridToWorksheet(worksheet);
                        workbook.SaveAs(saveDialog.FileName);
                    }

                    SetStatusMessage("Excel exported successfully.", Color.Green);
                }
                catch (Exception ex)
                {
                    SetStatusMessage("Failed to export Excel.", Color.Red);

                    MessageBox.Show(
                        "Failed to export Excel.\n\n" + ex.Message,
                        "Export Excel",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                }
            }
        }

        private void ExportGridToWorksheet(IXLWorksheet worksheet)
        {
            var rowsToExport = dgvResults.Rows
                .Cast<DataGridViewRow>()
                .Where(r => !r.IsNewRow && r.Visible)
                .Select(r => r.DataBoundItem as CompareResult)
                .Where(r => r != null)
                .ToList();

            int headerRow = 1;
            int currentRow = 2;

            var headers = new[]
            {
                "Solution Unique Name",
                "Source Display Name",
                "Target Display Name",
                "Source Version",
                "Target Version",
                "Source Publisher",
                "Target Publisher",
                "Source Package Type",
                "Target Package Type",
                "Package Type Status",
                "Overall Status"
            };

            for (int colIndex = 0; colIndex < headers.Length; colIndex++)
            {
                var cell = worksheet.Cell(headerRow, colIndex + 1);

                cell.Value = headers[colIndex];
                cell.Style.Font.Bold = true;
                cell.Style.Fill.BackgroundColor = XLColor.LightGray;
                cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                cell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                cell.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                cell.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
            }

            foreach (var item in rowsToExport)
            {
                worksheet.Cell(currentRow, 1).Value = item.UniqueName ?? string.Empty;
                worksheet.Cell(currentRow, 2).Value = item.SourceDisplayName ?? string.Empty;
                worksheet.Cell(currentRow, 3).Value = item.TargetDisplayName ?? string.Empty;
                worksheet.Cell(currentRow, 4).Value = item.SourceVersion ?? string.Empty;
                worksheet.Cell(currentRow, 5).Value = item.TargetVersion ?? string.Empty;
                worksheet.Cell(currentRow, 6).Value = item.SourcePublisher ?? string.Empty;
                worksheet.Cell(currentRow, 7).Value = item.TargetPublisher ?? string.Empty;
                worksheet.Cell(currentRow, 8).Value = item.SourcePackageType ?? string.Empty;
                worksheet.Cell(currentRow, 9).Value = item.TargetPackageType ?? string.Empty;
                worksheet.Cell(currentRow, 10).Value = item.PackageTypeStatus ?? string.Empty;
                worksheet.Cell(currentRow, 11).Value = item.Status ?? string.Empty;

                for (int colIndex = 1; colIndex <= headers.Length; colIndex++)
                {
                    var cell = worksheet.Cell(currentRow, colIndex);
                    cell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                    cell.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                    cell.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
                }

                ApplyExcelStatusStyle(worksheet.Cell(currentRow, 10), item.PackageTypeStatus ?? string.Empty);
                ApplyExcelStatusStyle(worksheet.Cell(currentRow, 11), item.Status ?? string.Empty);

                if (IsAnyPackageTypeDifference(item))
                {
                    ApplyExcelStatusStyle(worksheet.Cell(currentRow, 8), "Package Type Mismatch");
                    ApplyExcelStatusStyle(worksheet.Cell(currentRow, 9), "Package Type Mismatch");
                }

                currentRow++;
            }

            var usedRange = worksheet.RangeUsed();
            if (usedRange != null)
            {
                usedRange.SetAutoFilter();
            }

            worksheet.SheetView.FreezeRows(1);
            worksheet.Columns().AdjustToContents();

            foreach (var column in worksheet.ColumnsUsed())
            {
                if (column.Width > 40)
                    column.Width = 40;
            }

            worksheet.Row(1).Height = 22;
        }

        private void ApplyExcelStatusStyle(IXLCell cell, string status)
        {
            cell.Style.Font.Bold = true;

            switch (status)
            {
                case "Match":
                    cell.Style.Font.FontColor = XLColor.Green;
                    cell.Style.Fill.BackgroundColor = XLColor.Honeydew;
                    break;

                case "Version Mismatch":
                    cell.Style.Font.FontColor = XLColor.DarkOrange;
                    cell.Style.Fill.BackgroundColor = XLColor.Moccasin;
                    break;

                case "Publisher Mismatch":
                    cell.Style.Font.FontColor = XLColor.DarkViolet;
                    cell.Style.Fill.BackgroundColor = XLColor.Lavender;
                    break;

                case "Display Name Mismatch":
                    cell.Style.Font.FontColor = XLColor.SteelBlue;
                    cell.Style.Fill.BackgroundColor = XLColor.AliceBlue;
                    break;

                case "Package Type Mismatch":
                case "Managed/Unmanaged Mismatch":
                    cell.Style.Font.FontColor = XLColor.Teal;
                    cell.Style.Fill.BackgroundColor = XLColor.LightCyan;
                    break;

                case "Multiple Differences":
                    cell.Style.Font.FontColor = XLColor.DarkMagenta;
                    cell.Style.Fill.BackgroundColor = XLColor.MistyRose;
                    break;

                case "Missing in Source":
                    cell.Style.Font.FontColor = XLColor.Red;
                    cell.Style.Fill.BackgroundColor = XLColor.MistyRose;
                    break;

                case "Missing in Target":
                    cell.Style.Font.FontColor = XLColor.Firebrick;
                    cell.Style.Fill.BackgroundColor = XLColor.Linen;
                    break;
            }
        }
    }
}