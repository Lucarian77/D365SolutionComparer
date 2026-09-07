using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Windows.Forms;
using System.Xml;
using X = DocumentFormat.OpenXml;
using Xp = DocumentFormat.OpenXml.Packaging;
using Xs = DocumentFormat.OpenXml.Spreadsheet;
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
        private Button btnExport;
        private Button btnFilter;
        private Button btnResetFilters;
        private Button btnAbout;
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
        private CheckBox chkChangedOnly;

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
        private bool suppressSettingsSave;
        private bool pendingTargetConnectedStatus;

        private Settings userSettings;

        private List<ModelSolutionInfo> sourceSolutions = new List<ModelSolutionInfo>();
        private List<ModelSolutionInfo> targetSolutions = new List<ModelSolutionInfo>();
        private List<CompareResult> comparisonResults = new List<CompareResult>();

        public SolutionComparerControl()
        {
            Dock = DockStyle.Fill;
            BackColor = Color.White;
            userSettings = Settings.Load();
            BuildUi();
        }

        private void BuildUi()
        {
            Controls.Clear();

            lblTitle = new Label
            {
                Text = "D365 Solution Comparer v" + GetProductVersion(),
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
            btnExport = new Button { Text = "Export", Width = 110, Height = 30 };
            btnFilter = new Button { Text = "Filter: All", Width = 220, Height = 30 };
            btnResetFilters = new Button { Text = "Reset Filters", Width = 110, Height = 30 };
            btnAbout = new Button { Text = "About", Width = 90, Height = 30 };

            chkPackageTypeMismatchOnly = new CheckBox
            {
                Text = "Managed/unmanaged differences only",
                AutoSize = true,
                Height = 30,
                Margin = new Padding(10, 6, 0, 0),
                BackColor = Color.White
            };

            chkChangedOnly = new CheckBox
            {
                Text = "Changed only",
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
            btnExport.Click += BtnExport_Click;
            btnFilter.Click += BtnFilter_Click;
            btnResetFilters.Click += BtnResetFilters_Click;
            btnAbout.Click += BtnAbout_Click;
            chkPackageTypeMismatchOnly.CheckedChanged += ChkPackageTypeMismatchOnly_CheckedChanged;
            chkChangedOnly.CheckedChanged += ChkChangedOnly_CheckedChanged;

            topPanel.Controls.Add(btnLoadSource);
            topPanel.Controls.Add(btnConnectTarget);
            topPanel.Controls.Add(btnLoadTarget);
            topPanel.Controls.Add(btnCompare);
            topPanel.Controls.Add(btnExport);
            topPanel.Controls.Add(btnFilter);
            topPanel.Controls.Add(btnResetFilters);
            topPanel.Controls.Add(btnAbout);
            topPanel.Controls.Add(chkChangedOnly);
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
            dgvResults.CellDoubleClick += DgvResults_CellDoubleClick;
            dgvResults.CellToolTipTextNeeded += DgvResults_CellToolTipTextNeeded;
            Resize += SolutionComparerControl_Resize;

            Controls.Add(dgvResults);
            Controls.Add(lblLegend);
            Controls.Add(lblSummary);
            Controls.Add(lblStatusMessage);
            Controls.Add(lblTargetEnv);
            Controls.Add(lblSourceEnv);
            Controls.Add(topPanel);
            Controls.Add(lblTitle);

            ApplySavedFilterState();
            ApplyResponsiveTopPanelLayout();
            UpdateActionButtonStates();
            UpdateLayoutHeights();
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

        private void UpdateActionButtonStates()
        {
            if (btnLoadSource != null)
            {
                btnLoadSource.Enabled = Service != null;
            }

            if (btnLoadTarget != null)
            {
                btnLoadTarget.Enabled = targetService != null;
            }

            if (btnCompare != null)
            {
                btnCompare.Enabled = sourceLoaded && targetLoaded && sourceSolutions.Count > 0 && targetSolutions.Count > 0;
            }

            if (btnExport != null)
            {
                btnExport.Enabled = GetVisibleComparisonResultCount() > 0;
            }

            var hasComparisonResults = comparisonResults != null && comparisonResults.Count > 0;

            if (btnFilter != null)
            {
                btnFilter.Enabled = hasComparisonResults;
            }

            if (btnResetFilters != null)
            {
                btnResetFilters.Enabled = hasComparisonResults && HasActiveComparisonFilters();
            }

            if (chkChangedOnly != null)
            {
                chkChangedOnly.Enabled = hasComparisonResults;
            }

            if (chkPackageTypeMismatchOnly != null)
            {
                chkPackageTypeMismatchOnly.Enabled = hasComparisonResults;
            }
        }

        private int GetVisibleComparisonResultCount()
        {
            if (dgvResults == null)
            {
                return 0;
            }

            return dgvResults.Rows
                .Cast<DataGridViewRow>()
                .Count(r => !r.IsNewRow && r.Visible && r.DataBoundItem is CompareResult);
        }

        private void ApplyResponsiveTopPanelLayout()
        {
            if (topPanel == null)
            {
                return;
            }

            var compact = ClientSize.Width > 0 && ClientSize.Width < 1450;
            var wrapped = ClientSize.Width > 0 && ClientSize.Width < 1280;

            btnLoadSource.Width = compact ? 105 : 110;
            btnConnectTarget.Width = compact ? 115 : 120;
            btnLoadTarget.Width = compact ? 105 : 110;
            btnCompare.Width = compact ? 105 : 110;
            btnExport.Width = compact ? 105 : 110;
            btnFilter.Width = compact ? 165 : 220;
            btnResetFilters.Width = compact ? 100 : 110;
            btnAbout.Width = compact ? 75 : 90;

            if (chkPackageTypeMismatchOnly != null)
            {
                chkPackageTypeMismatchOnly.Text = compact
                    ? "Managed/unmanaged only"
                    : "Managed/unmanaged differences only";
            }

            topPanel.WrapContents = wrapped;
            topPanel.AutoScroll = !wrapped;
            topPanel.Height = wrapped ? 78 : 45;
        }

        public override void UpdateConnection(OrgService newService, ConnectionDetail detail, string actionName, object parameter)
        {
            base.UpdateConnection(newService, detail, actionName, parameter);

            if (!sourceLoaded)
            {
                sourceConnectionName = detail != null && !string.IsNullOrWhiteSpace(detail.ConnectionName)
                    ? detail.ConnectionName
                    : "Current XrmToolBox connection";
            }

            RefreshEnvironmentLabels();
            ApplyResponsiveTopPanelLayout();
            UpdateActionButtonStates();

            if (pendingTargetConnectedStatus && targetConnectionDetail != null && targetService != null && !targetLoaded)
            {
                SetStatusMessage("Target environment connected successfully. Load target solutions when ready.", Color.Green);
            }
            else
            {
                SetStatusMessage("Ready", Color.Green);
            }
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
                    SetSummaryText("Summary: No comparison results yet");

                    pendingTargetConnectedStatus = true;
                    RefreshEnvironmentLabels();
                    UpdateActionButtonStates();
                    BeginInvoke(new Action(() =>
                    {
                        SetStatusMessage("Target environment connected successfully. Load target solutions when ready.", Color.Green);
                    }));
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
                    SetSummaryText("Summary: No comparison results yet");
                    pendingTargetConnectedStatus = false;

                    RefreshEnvironmentLabels();
                    UpdateActionButtonStates();
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

                sourceConnectionName = ConnectionDetail != null && !string.IsNullOrWhiteSpace(ConnectionDetail.ConnectionName)
                    ? ConnectionDetail.ConnectionName
                    : "Current XrmToolBox connection";

                dgvResults.DataSource = null;
                dgvResults.DataSource = sourceSolutions;

                RefreshEnvironmentLabels();
                SetSummaryText("Summary: No comparison results yet");

                ApplySolutionListGridLayout();
                ResetGridScrollPosition();

                SetStatusMessage($"Loaded {sourceSolutions.Count} source solutions from the source environment.", Color.Green);
                UpdateActionButtonStates();
            }
            catch (Exception ex)
            {
                SetStatusMessage("Failed to load source solutions.", Color.Red);
                UpdateActionButtonStates();

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
                    SetSummaryText("Summary: No comparison results yet");
                    RefreshEnvironmentLabels();
                    UpdateActionButtonStates();
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
                pendingTargetConnectedStatus = false;

                targetConnectionName = targetConnectionDetail != null && !string.IsNullOrWhiteSpace(targetConnectionDetail.ConnectionName)
                    ? targetConnectionDetail.ConnectionName
                    : "Target environment";

                dgvResults.DataSource = null;
                dgvResults.DataSource = targetSolutions;

                RefreshEnvironmentLabels();
                SetSummaryText("Summary: No comparison results yet");

                ApplySolutionListGridLayout();
                ResetGridScrollPosition();

                SetStatusMessage($"Loaded {targetSolutions.Count} target solutions from the target environment.", Color.Green);
                UpdateActionButtonStates();
            }
            catch (Exception ex)
            {
                SetStatusMessage("Failed to load target solutions.", Color.Red);
                UpdateActionButtonStates();

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

                BindFilteredResults();
                SetStatusMessage($"Comparison completed successfully. {GetVisibleComparisonRows().Count} visible result(s).", Color.Green);
            }
            catch (Exception ex)
            {
                SetStatusMessage("Failed to compare solutions.", Color.Red);
                UpdateActionButtonStates();

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

        private void BtnResetFilters_Click(object sender, EventArgs e)
        {
            ResetFilterSelection();

            if (HasComparisonResults())
            {
                BindFilteredResults();
            }
            else
            {
                UpdateActionButtonStates();
            }
        }

        private void BtnAbout_Click(object sender, EventArgs e)
        {
            var version = GetProductVersion();
            var visibleRowCount = GetVisibleComparisonRows().Count;

            var message =
                "D365 Solution Comparer\n\n" +
                "Version: " + version + "\n" +
                "Source loaded: " + (sourceLoaded ? "Yes" : "No") + "\n" +
                "Target loaded: " + (targetLoaded ? "Yes" : "No") + "\n" +
                "Current comparison rows: " + visibleRowCount + "\n" +
                "Export format: XLSX, Excel XML, or CSV\n\n" +
                "Double-click a comparison row to view details.\n" +
                "XLSX is the default Excel workbook export format.\n" +
                "Excel XML remains available for readable Excel-friendly output without extra libraries.\n" +
                "CSV remains available for plain-text exports.\n" +
                "On some machines, .xml files open more reliably when opened from inside Excel.";

            MessageBox.Show(
                message,
                "About D365 Solution Comparer",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
        }

        private void ChkPackageTypeMismatchOnly_CheckedChanged(object sender, EventArgs e)
        {
            PersistUiSettings();

            if (HasComparisonResults())
            {
                BindFilteredResults();
            }
        }

        private void ChkChangedOnly_CheckedChanged(object sender, EventArgs e)
        {
            PersistUiSettings();

            if (HasComparisonResults())
            {
                BindFilteredResults();
            }
        }

        private void FilterItem_CheckedChanged(object sender, EventArgs e)
        {
            if (!(sender is ToolStripMenuItem changedItem))
            {
                return;
            }

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
            PersistUiSettings();

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
            suppressSettingsSave = true;

            try
            {
                miAll.CheckedChanged -= FilterItem_CheckedChanged;
                miAll.Checked = true;
                miAll.CheckedChanged += FilterItem_CheckedChanged;

                SetNonAllItemsChecked(false);
                chkPackageTypeMismatchOnly.Checked = false;
                chkChangedOnly.Checked = false;
                UpdateFilterButtonText();
            }
            finally
            {
                suppressSettingsSave = false;
            }

            PersistUiSettings();
            UpdateActionButtonStates();
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

            if (chkChangedOnly != null && chkChangedOnly.Checked)
            {
                filteredResults = filteredResults.Where(IsChangedResult);
            }

            var finalResults = filteredResults.ToList();

            dgvResults.DataSource = null;
            dgvResults.DataSource = finalResults;

            ApplyComparisonGridLayout();
            ResetGridScrollPosition();
            UpdateSummary(finalResults);
            UpdateActionButtonStates();
            UpdateFilteredResultsStatus(finalResults);
        }

        private void UpdateFilteredResultsStatus(List<CompareResult> results)
        {
            if (!HasComparisonResults())
            {
                return;
            }

            var visibleCount = results?.Count ?? 0;

            if (visibleCount == 0)
            {
                SetStatusMessage("No comparison results match the current filters.", Color.DarkOrange);
                return;
            }

            if (!HasActiveComparisonFilters())
            {
                SetStatusMessage($"Showing all {visibleCount} comparison result(s).", Color.Green);
                return;
            }

            SetStatusMessage($"Showing {visibleCount} filtered comparison result(s).", Color.Green);
        }

        private bool HasActiveComparisonFilters()
        {
            return (chkChangedOnly != null && chkChangedOnly.Checked)
                   || (chkPackageTypeMismatchOnly != null && chkPackageTypeMismatchOnly.Checked)
                   || !miAll.Checked;
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
            {
                return;
            }

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
                // Ignore grid positioning issues.
            }
        }

        private void SetSummaryText(string text)
        {
            lblSummary.Text = text;
            UpdateLayoutHeights();
        }

        private void SolutionComparerControl_Resize(object sender, EventArgs e)
        {
            ApplyResponsiveTopPanelLayout();
            UpdateLayoutHeights();
        }

        private void UpdateLayoutHeights()
        {
            if (lblSummary == null || lblLegend == null)
            {
                return;
            }

            lblSummary.Height = GetPreferredLabelHeight(lblSummary, 36);
            lblLegend.Height = GetPreferredLabelHeight(lblLegend, 24);
        }

        private int GetPreferredLabelHeight(Label label, int minimumHeight)
        {
            if (label == null)
            {
                return minimumHeight;
            }

            var availableWidth = Math.Max(200, ClientSize.Width - label.Padding.Left - label.Padding.Right - 24);
            var proposedSize = new Size(availableWidth, int.MaxValue);
            var flags = TextFormatFlags.WordBreak;
            var measured = TextRenderer.MeasureText(label.Text ?? string.Empty, label.Font, proposedSize, flags);
            var height = measured.Height + label.Padding.Top + label.Padding.Bottom + 6;
            return Math.Max(minimumHeight, height);
        }

        private void DgvResults_CellToolTipTextNeeded(object sender, DataGridViewCellToolTipTextNeededEventArgs e)
        {
            if (e.RowIndex < 0 || e.ColumnIndex < 0)
            {
                return;
            }

            var value = dgvResults.Rows[e.RowIndex].Cells[e.ColumnIndex].Value;
            if (value == null)
            {
                return;
            }

            var text = Convert.ToString(value);
            if (!string.IsNullOrWhiteSpace(text) && text.Length > 24)
            {
                e.ToolTipText = text;
            }
        }

        private void UpdateSummary(List<CompareResult> results)
        {
            if (results == null || results.Count == 0)
            {
                SetSummaryText("Summary: No comparison results");
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

            SetSummaryText(
                $"Summary: Total={total} | Match={match} | Version={versionMismatch} | Publisher={publisherMismatch} | Display Name={displayNameMismatch}\r\n" +
                $"Package Type Differences={packageTypeDifference} | Managed/Unmanaged Differences={managedUnmanagedDifference} | Multiple={multipleDifferences} | Missing in Source={missingInSource} | Missing in Target={missingInTarget}");
        }

        private bool IsAnyPackageTypeDifference(CompareResult result)
        {
            if (result == null)
            {
                return false;
            }

            if (IsManagedUnmanagedDifference(result))
            {
                return true;
            }

            var packageTypeStatus = (result.PackageTypeStatus ?? string.Empty).Trim();

            if (string.Equals(packageTypeStatus, "Package Type Mismatch", StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            var sourceType = (result.SourcePackageType ?? string.Empty).Trim();
            var targetType = (result.TargetPackageType ?? string.Empty).Trim();

            if (string.IsNullOrWhiteSpace(sourceType) || string.IsNullOrWhiteSpace(targetType))
            {
                return false;
            }

            return !string.Equals(sourceType, targetType, StringComparison.OrdinalIgnoreCase);
        }

        private bool IsManagedUnmanagedDifference(CompareResult result)
        {
            if (result == null)
            {
                return false;
            }

            var packageTypeStatus = (result.PackageTypeStatus ?? string.Empty).Trim();

            if (string.Equals(packageTypeStatus, "Managed/Unmanaged Mismatch", StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            var sourceType = (result.SourcePackageType ?? string.Empty).Trim();
            var targetType = (result.TargetPackageType ?? string.Empty).Trim();

            if (string.IsNullOrWhiteSpace(sourceType) || string.IsNullOrWhiteSpace(targetType))
            {
                return false;
            }

            var sourceManagedState = IsManagedOrUnmanagedLabel(sourceType);
            var targetManagedState = IsManagedOrUnmanagedLabel(targetType);

            if (!sourceManagedState || !targetManagedState)
            {
                return false;
            }

            return !string.Equals(sourceType, targetType, StringComparison.OrdinalIgnoreCase);
        }

        private bool IsManagedOrUnmanagedLabel(string value)
        {
            return string.Equals(value, "Managed", StringComparison.OrdinalIgnoreCase)
                   || string.Equals(value, "Unmanaged", StringComparison.OrdinalIgnoreCase);
        }

        private bool IsChangedResult(CompareResult result)
        {
            if (result == null)
            {
                return false;
            }

            if (IsAnyPackageTypeDifference(result))
            {
                return true;
            }

            var status = (result.Status ?? string.Empty).Trim();
            return !string.Equals(status, "Match", StringComparison.OrdinalIgnoreCase);
        }

        private void SetStatusMessage(string message, Color color)
        {
            lblStatusMessage.Text = "Status: " + message;
            lblStatusMessage.ForeColor = color;
        }

        private void DgvResults_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            if (e.RowIndex < 0 || e.ColumnIndex < 0)
            {
                return;
            }

            var column = dgvResults.Columns[e.ColumnIndex];
            var dataPropertyName = column.DataPropertyName;

            var rowData = dgvResults.Rows[e.RowIndex].DataBoundItem as CompareResult;
            var packageTypeDifference = IsAnyPackageTypeDifference(rowData);

            if (e.Value == null)
            {
                return;
            }

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
            {
                return;
            }

            var status = e.Value.ToString();
            if (string.IsNullOrWhiteSpace(status))
            {
                return;
            }

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

        private void DgvResults_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0)
            {
                return;
            }

            var rowData = dgvResults.Rows[e.RowIndex].DataBoundItem as CompareResult;
            if (rowData == null)
            {
                return;
            }

            ShowRowDetails(rowData);
        }

        private void BtnExport_Click(object sender, EventArgs e)
        {
            var rowsToExport = GetVisibleComparisonRows();

            if (!HasComparisonResults() || rowsToExport.Count == 0)
            {
                SetStatusMessage("There is no comparison data to export.", Color.DarkOrange);

                MessageBox.Show(
                    "There is no comparison data to export.",
                    "Export",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
                return;
            }

            using (var saveDialog = CreateExportSaveDialog())
            {
                if (saveDialog.ShowDialog() != DialogResult.OK)
                {
                    SetStatusMessage("Export cancelled.", Color.DarkOrange);
                    return;
                }

                var extension = (Path.GetExtension(saveDialog.FileName) ?? string.Empty).ToLowerInvariant();

                switch (extension)
                {
                    case ".csv":
                        TryExportCsv(saveDialog.FileName);
                        break;

                    case ".xml":
                        TryExportSpreadsheetMl(saveDialog.FileName);
                        break;

                    default:
                        TryExportXlsx(saveDialog.FileName);
                        break;
                }
            }
        }

        private SaveFileDialog CreateExportSaveDialog()
        {
            return new SaveFileDialog
            {
                AddExtension = true,
                OverwritePrompt = true,
                SupportMultiDottedExtensions = true,
                Filter = "Excel Workbook (*.xlsx)|*.xlsx|Excel XML Spreadsheet 2003 (*.xml)|*.xml|CSV File (*.csv)|*.csv",
                DefaultExt = "xlsx",
                FileName = $"D365SolutionComparer_Source_vs_Target_{DateTime.Now:yyyyMMdd_HHmm}.xlsx",
                FilterIndex = 1
            };
        }

        private void TryExportXlsx(string filePath)
        {
            try
            {
                ExportVisibleRowsToXlsxOpenXml(filePath);
                SetStatusMessage("XLSX exported successfully.", Color.Green);

                MessageBox.Show(
                    "XLSX export completed successfully.\n\n" +
                    "This workbook opens directly in Excel and keeps the richer layout, filtering, and status styling of the comparison export.\n\n" +
                    filePath,
                    "Export",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                SetStatusMessage("Failed to export XLSX.", Color.Red);

                MessageBox.Show(
                    "Failed to export XLSX.\n\n" + GetExceptionSummary(ex),
                    "Export",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        private void TryExportSpreadsheetMl(string filePath)
        {
            try
            {
                ExportVisibleRowsToSpreadsheetMl(filePath);
                SetStatusMessage("Excel XML exported successfully.", Color.Green);

                MessageBox.Show(
                    "Excel XML export completed successfully.\n\n" +
                    "This format opens in Excel with better column layout and styling than CSV, without requiring external Excel libraries.\n\n" +
                    "On some machines, .xml files may not be associated with Excel. If the file does not open directly from File Explorer, open Excel first and then open the exported XML file.\n\n" +
                    filePath,
                    "Export",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                SetStatusMessage("Failed to export Excel XML.", Color.Red);

                MessageBox.Show(
                    "Failed to export Excel XML.\n\n" + GetExceptionSummary(ex),
                    "Export",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        private void TryExportCsv(string filePath)
        {
            try
            {
                ExportVisibleRowsToCsv(filePath);
                SetStatusMessage("CSV exported successfully.", Color.Green);

                MessageBox.Show(
                    "CSV export completed successfully.\n\n" +
                    "CSV is a plain-text format, so column widths and cell styling are controlled by the application used to open it.\n\n" +
                    filePath,
                    "Export",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                SetStatusMessage("Failed to export CSV.", Color.Red);

                MessageBox.Show(
                    "Failed to export CSV.\n\n" + GetExceptionSummary(ex),
                    "Export",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        private void ExportVisibleRowsToXlsxOpenXml(string filePath)
        {
            var rowsToExport = GetVisibleComparisonRows();
            var activeFilterLabel = btnFilter != null
                ? btnFilter.Text.Replace("Filter: ", string.Empty)
                : "All";

            using (var spreadsheet = Xp.SpreadsheetDocument.Create(filePath, X.SpreadsheetDocumentType.Workbook))
            {
                var workbookPart = spreadsheet.AddWorkbookPart();
                workbookPart.Workbook = new Xs.Workbook();
                workbookPart.Workbook.AppendChild(new Xs.BookViews(new Xs.WorkbookView()));

                var stylesPart = workbookPart.AddNewPart<Xp.WorkbookStylesPart>();
                stylesPart.Stylesheet = CreateXlsxStylesheet();
                stylesPart.Stylesheet.Save();

                var worksheetPart = workbookPart.AddNewPart<Xp.WorksheetPart>();
                var worksheet = new Xs.Worksheet();
                var columns = CreateXlsxColumns();
                var sheetData = new Xs.SheetData();

                worksheet.Append(CreateXlsxSheetViews());
                worksheet.Append(columns);
                worksheet.Append(sheetData);

                var currentRowIndex = 1u;

                var titleRow = new Xs.Row { RowIndex = currentRowIndex };
                titleRow.Append(CreateInlineStringCell("D365 Solution Comparer Export", 1));
                for (var i = 0; i < 10; i++)
                {
                    titleRow.Append(CreateInlineStringCell(string.Empty, 1));
                }
                sheetData.Append(titleRow);
                currentRowIndex++;

                sheetData.Append(CreateMetadataRow(currentRowIndex++, new[]
                {
                    Tuple.Create("Source", sourceConnectionName),
                    Tuple.Create("Target", targetConnectionName),
                    Tuple.Create("Visible Rows", rowsToExport.Count.ToString())
                }));

                sheetData.Append(CreateMetadataRow(currentRowIndex++, new[]
                {
                    Tuple.Create("Exported On", DateTime.Now.ToString("yyyy-MM-dd HH:mm")),
                    Tuple.Create("Status Filter", activeFilterLabel),
                    Tuple.Create("Changed Only", chkChangedOnly != null && chkChangedOnly.Checked ? "Yes" : "No")
                }));

                sheetData.Append(CreateMetadataRow(currentRowIndex++, new[]
                {
                    Tuple.Create("Managed/Unmanaged Only", chkPackageTypeMismatchOnly != null && chkPackageTypeMismatchOnly.Checked ? "Yes" : "No")
                }));

                sheetData.Append(new Xs.Row { RowIndex = currentRowIndex++ });

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

                var headerRowIndex = currentRowIndex;
                var dataStartRowIndex = headerRowIndex + 1;
                var headerRow = new Xs.Row { RowIndex = currentRowIndex++ };
                foreach (var header in headers)
                {
                    headerRow.Append(CreateInlineStringCell(header, 4));
                }
                sheetData.Append(headerRow);

                foreach (var item in rowsToExport)
                {
                    var row = new Xs.Row { RowIndex = currentRowIndex++ };

                    row.Append(CreateInlineStringCell(NormalizeCsvValue(item.UniqueName), 0));
                    row.Append(CreateInlineStringCell(NormalizeCsvValue(item.SourceDisplayName), 5));
                    row.Append(CreateInlineStringCell(NormalizeCsvValue(item.TargetDisplayName), 5));
                    row.Append(CreateInlineStringCell(NormalizeCsvValue(item.SourceVersion), 6));
                    row.Append(CreateInlineStringCell(NormalizeCsvValue(item.TargetVersion), 6));
                    row.Append(CreateInlineStringCell(NormalizeCsvValue(item.SourcePublisher), 5));
                    row.Append(CreateInlineStringCell(NormalizeCsvValue(item.TargetPublisher), 5));

                    var packageTypeStyleIndex = IsAnyPackageTypeDifference(item) ? 11u : 6u;
                    row.Append(CreateInlineStringCell(NormalizeCsvValue(item.SourcePackageType), packageTypeStyleIndex));
                    row.Append(CreateInlineStringCell(NormalizeCsvValue(item.TargetPackageType), packageTypeStyleIndex));
                    row.Append(CreateInlineStringCell(NormalizeCsvValue(item.PackageTypeStatus), GetXlsxStatusStyleIndex(item.PackageTypeStatus)));
                    row.Append(CreateInlineStringCell(NormalizeCsvValue(item.Status), GetXlsxStatusStyleIndex(item.Status)));

                    sheetData.Append(row);
                }

                if (rowsToExport.Count > 0)
                {
                    var lastDataRowIndex = currentRowIndex - 1;

                    worksheet.Append(new Xs.AutoFilter
                    {
                        Reference = new X.StringValue($"A{headerRowIndex}:K{lastDataRowIndex}")
                    });
                }

                var mergeCells = new Xs.MergeCells();
                mergeCells.Append(new Xs.MergeCell { Reference = new X.StringValue("A1:K1") });
                worksheet.Append(mergeCells);

                worksheetPart.Worksheet = worksheet;
                worksheetPart.Worksheet.Save();

                var sheets = workbookPart.Workbook.AppendChild(new Xs.Sheets());
                sheets.Append(new Xs.Sheet
                {
                    Id = workbookPart.GetIdOfPart(worksheetPart),
                    SheetId = 1u,
                    Name = "Comparison Results"
                });

                workbookPart.Workbook.Save();
            }
        }

        private Xs.Row CreateMetadataRow(uint rowIndex, IEnumerable<Tuple<string, string>> pairs)
        {
            var row = new Xs.Row { RowIndex = rowIndex };
            var values = new List<Xs.Cell>();

            foreach (var pair in pairs)
            {
                values.Add(CreateInlineStringCell(pair.Item1, 2));
                values.Add(CreateInlineStringCell(pair.Item2, 3));
                values.Add(CreateInlineStringCell(string.Empty, 3));
            }

            while (values.Count < 11)
            {
                values.Add(CreateInlineStringCell(string.Empty, 3));
            }

            foreach (var cell in values.Take(11))
            {
                row.Append(cell);
            }

            return row;
        }

        private Xs.Cell CreateInlineStringCell(string value, uint styleIndex)
        {
            return new Xs.Cell
            {
                DataType = Xs.CellValues.InlineString,
                StyleIndex = styleIndex,
                InlineString = new Xs.InlineString(new Xs.Text(value ?? string.Empty) { Space = X.SpaceProcessingModeValues.Preserve })
            };
        }

        private Xs.Columns CreateXlsxColumns()
        {
            var columns = new Xs.Columns();
            var widths = new[] { 28d, 32d, 32d, 14d, 14d, 24d, 24d, 16d, 16d, 18d, 18d };

            for (uint i = 0; i < widths.Length; i++)
            {
                columns.Append(new Xs.Column
                {
                    Min = i + 1,
                    Max = i + 1,
                    Width = widths[i],
                    CustomWidth = true
                });
            }

            return columns;
        }

        private Xs.SheetViews CreateXlsxSheetViews()
        {
            var sheetView = new Xs.SheetView { WorkbookViewId = 0u, TabSelected = true };
            sheetView.Append(new Xs.Pane
            {
                VerticalSplit = 6d,
                TopLeftCell = "A7",
                ActivePane = Xs.PaneValues.BottomLeft,
                State = Xs.PaneStateValues.Frozen
            });

            return new Xs.SheetViews(sheetView);
        }

        private Xs.Stylesheet CreateXlsxStylesheet()
        {
            var fonts = new Xs.Fonts(
                new Xs.Font(new Xs.FontName { Val = "Calibri" }, new Xs.FontSize { Val = 11d }),
                new Xs.Font(new Xs.Bold(), new Xs.FontName { Val = "Calibri" }, new Xs.FontSize { Val = 14d }),
                new Xs.Font(new Xs.Bold(), new Xs.Color { Rgb = "FFFFFF" }, new Xs.FontName { Val = "Calibri" }, new Xs.FontSize { Val = 11d }),
                new Xs.Font(new Xs.Bold(), new Xs.FontName { Val = "Calibri" }, new Xs.FontSize { Val = 11d }, new Xs.Color { Rgb = "008000" }),
                new Xs.Font(new Xs.Bold(), new Xs.FontName { Val = "Calibri" }, new Xs.FontSize { Val = 11d }, new Xs.Color { Rgb = "FF8C00" }),
                new Xs.Font(new Xs.Bold(), new Xs.FontName { Val = "Calibri" }, new Xs.FontSize { Val = 11d }, new Xs.Color { Rgb = "9400D3" }),
                new Xs.Font(new Xs.Bold(), new Xs.FontName { Val = "Calibri" }, new Xs.FontSize { Val = 11d }, new Xs.Color { Rgb = "4682B4" }),
                new Xs.Font(new Xs.Bold(), new Xs.FontName { Val = "Calibri" }, new Xs.FontSize { Val = 11d }, new Xs.Color { Rgb = "008080" }),
                new Xs.Font(new Xs.Bold(), new Xs.FontName { Val = "Calibri" }, new Xs.FontSize { Val = 11d }, new Xs.Color { Rgb = "8B008B" }),
                new Xs.Font(new Xs.Bold(), new Xs.FontName { Val = "Calibri" }, new Xs.FontSize { Val = 11d }, new Xs.Color { Rgb = "FF0000" }),
                new Xs.Font(new Xs.Bold(), new Xs.FontName { Val = "Calibri" }, new Xs.FontSize { Val = 11d }, new Xs.Color { Rgb = "B22222" }));
            fonts.Count = (uint)fonts.ChildElements.Count;

            var fills = new Xs.Fills(
                new Xs.Fill(new Xs.PatternFill { PatternType = Xs.PatternValues.None }),
                new Xs.Fill(new Xs.PatternFill { PatternType = Xs.PatternValues.Gray125 }),
                new Xs.Fill(new Xs.PatternFill(new Xs.ForegroundColor { Rgb = "D9E2F3" }) { PatternType = Xs.PatternValues.Solid }),
                new Xs.Fill(new Xs.PatternFill(new Xs.ForegroundColor { Rgb = "EEF3F8" }) { PatternType = Xs.PatternValues.Solid }),
                new Xs.Fill(new Xs.PatternFill(new Xs.ForegroundColor { Rgb = "44546A" }) { PatternType = Xs.PatternValues.Solid }),
                new Xs.Fill(new Xs.PatternFill(new Xs.ForegroundColor { Rgb = "F0FFF0" }) { PatternType = Xs.PatternValues.Solid }),
                new Xs.Fill(new Xs.PatternFill(new Xs.ForegroundColor { Rgb = "FFE4B5" }) { PatternType = Xs.PatternValues.Solid }),
                new Xs.Fill(new Xs.PatternFill(new Xs.ForegroundColor { Rgb = "E6E6FA" }) { PatternType = Xs.PatternValues.Solid }),
                new Xs.Fill(new Xs.PatternFill(new Xs.ForegroundColor { Rgb = "F0F8FF" }) { PatternType = Xs.PatternValues.Solid }),
                new Xs.Fill(new Xs.PatternFill(new Xs.ForegroundColor { Rgb = "E0FFFF" }) { PatternType = Xs.PatternValues.Solid }),
                new Xs.Fill(new Xs.PatternFill(new Xs.ForegroundColor { Rgb = "FFE4E1" }) { PatternType = Xs.PatternValues.Solid }),
                new Xs.Fill(new Xs.PatternFill(new Xs.ForegroundColor { Rgb = "FAF0E6" }) { PatternType = Xs.PatternValues.Solid }));
            fills.Count = (uint)fills.ChildElements.Count;

            var borders = new Xs.Borders(
                new Xs.Border(),
                new Xs.Border(
                    new Xs.LeftBorder { Style = Xs.BorderStyleValues.Thin, Color = new Xs.Color { Auto = true } },
                    new Xs.RightBorder { Style = Xs.BorderStyleValues.Thin, Color = new Xs.Color { Auto = true } },
                    new Xs.TopBorder { Style = Xs.BorderStyleValues.Thin, Color = new Xs.Color { Auto = true } },
                    new Xs.BottomBorder { Style = Xs.BorderStyleValues.Thin, Color = new Xs.Color { Auto = true } },
                    new Xs.DiagonalBorder()));
            borders.Count = (uint)borders.ChildElements.Count;

            var cellStyleFormats = new Xs.CellStyleFormats(new Xs.CellFormat());
            cellStyleFormats.Count = 1u;

            var cellFormats = new Xs.CellFormats(
                CreateCellFormat(0, 0, 1, Xs.HorizontalAlignmentValues.Left, false),
                CreateCellFormat(1, 2, 1, Xs.HorizontalAlignmentValues.Left, false),
                CreateCellFormat(1, 3, 1, Xs.HorizontalAlignmentValues.Left, false),
                CreateCellFormat(0, 0, 1, Xs.HorizontalAlignmentValues.Left, false),
                CreateCellFormat(2, 4, 1, Xs.HorizontalAlignmentValues.Center, true),
                CreateCellFormat(0, 0, 1, Xs.HorizontalAlignmentValues.Left, true),
                CreateCellFormat(0, 0, 1, Xs.HorizontalAlignmentValues.Center, false),
                CreateCellFormat(3, 5, 1, Xs.HorizontalAlignmentValues.Center, false),
                CreateCellFormat(4, 6, 1, Xs.HorizontalAlignmentValues.Center, false),
                CreateCellFormat(5, 7, 1, Xs.HorizontalAlignmentValues.Center, false),
                CreateCellFormat(6, 8, 1, Xs.HorizontalAlignmentValues.Center, false),
                CreateCellFormat(7, 9, 1, Xs.HorizontalAlignmentValues.Center, false),
                CreateCellFormat(8, 10, 1, Xs.HorizontalAlignmentValues.Center, false),
                CreateCellFormat(9, 10, 1, Xs.HorizontalAlignmentValues.Center, false),
                CreateCellFormat(10, 11, 1, Xs.HorizontalAlignmentValues.Center, false));
            cellFormats.Count = (uint)cellFormats.ChildElements.Count;

            var cellStyles = new Xs.CellStyles(new Xs.CellStyle { Name = "Normal", FormatId = 0u, BuiltinId = 0u });
            cellStyles.Count = 1u;

            return new Xs.Stylesheet(fonts, fills, borders, cellStyleFormats, cellFormats, cellStyles);
        }

        private Xs.CellFormat CreateCellFormat(uint fontId, uint fillId, uint borderId, Xs.HorizontalAlignmentValues horizontal, bool wrapText)
        {
            return new Xs.CellFormat
            {
                FontId = fontId,
                FillId = fillId,
                BorderId = borderId,
                ApplyFont = true,
                ApplyFill = true,
                ApplyBorder = true,
                ApplyAlignment = true,
                Alignment = new Xs.Alignment
                {
                    Horizontal = horizontal,
                    Vertical = Xs.VerticalAlignmentValues.Center,
                    WrapText = wrapText
                }
            };
        }

        private uint GetXlsxStatusStyleIndex(string status)
        {
            switch ((status ?? string.Empty).Trim())
            {
                case "Match": return 7u;
                case "Version Mismatch": return 8u;
                case "Publisher Mismatch": return 9u;
                case "Display Name Mismatch": return 10u;
                case "Package Type Mismatch":
                case "Managed/Unmanaged Mismatch": return 11u;
                case "Multiple Differences": return 12u;
                case "Missing in Source": return 13u;
                case "Missing in Target": return 14u;
                default: return 6u;
            }
        }

        private void ExportVisibleRowsToSpreadsheetMl(string filePath)
        {
            var rowsToExport = GetVisibleComparisonRows();
            var activeFilterLabel = btnFilter != null
                ? btnFilter.Text.Replace("Filter: ", string.Empty)
                : "All";

            const string spreadsheetNamespace = "urn:schemas-microsoft-com:office:spreadsheet";
            const string officeNamespace = "urn:schemas-microsoft-com:office:office";
            const string excelNamespace = "urn:schemas-microsoft-com:office:excel";
            const string htmlNamespace = "http://www.w3.org/TR/REC-html40";

            var settings = new XmlWriterSettings
            {
                Indent = true,
                Encoding = new UTF8Encoding(true)
            };

            using (var stream = new FileStream(filePath, FileMode.Create, FileAccess.Write, FileShare.None))
            using (var writer = XmlWriter.Create(stream, settings))
            {
                writer.WriteStartDocument();
                writer.WriteProcessingInstruction("mso-application", "progid=\"Excel.Sheet\"");

                writer.WriteStartElement("Workbook", spreadsheetNamespace);
                writer.WriteAttributeString("xmlns", "o", null, officeNamespace);
                writer.WriteAttributeString("xmlns", "x", null, excelNamespace);
                writer.WriteAttributeString("xmlns", "ss", null, spreadsheetNamespace);
                writer.WriteAttributeString("xmlns", "html", null, htmlNamespace);

                WriteSpreadsheetMlStyles(writer, spreadsheetNamespace);

                writer.WriteStartElement("Worksheet", spreadsheetNamespace);
                writer.WriteAttributeString("ss", "Name", spreadsheetNamespace, "Comparison Results");

                writer.WriteStartElement("Table", spreadsheetNamespace);

                WriteSpreadsheetMlColumn(writer, spreadsheetNamespace, 220);
                WriteSpreadsheetMlColumn(writer, spreadsheetNamespace, 250);
                WriteSpreadsheetMlColumn(writer, spreadsheetNamespace, 250);
                WriteSpreadsheetMlColumn(writer, spreadsheetNamespace, 110);
                WriteSpreadsheetMlColumn(writer, spreadsheetNamespace, 110);
                WriteSpreadsheetMlColumn(writer, spreadsheetNamespace, 190);
                WriteSpreadsheetMlColumn(writer, spreadsheetNamespace, 190);
                WriteSpreadsheetMlColumn(writer, spreadsheetNamespace, 130);
                WriteSpreadsheetMlColumn(writer, spreadsheetNamespace, 130);
                WriteSpreadsheetMlColumn(writer, spreadsheetNamespace, 150);
                WriteSpreadsheetMlColumn(writer, spreadsheetNamespace, 150);

                writer.WriteStartElement("Row", spreadsheetNamespace);
                WriteSpreadsheetMlCell(writer, spreadsheetNamespace, "D365 Solution Comparer Export", "Title", 10);
                writer.WriteEndElement();

                writer.WriteStartElement("Row", spreadsheetNamespace);
                WriteSpreadsheetMlCell(writer, spreadsheetNamespace, "Source", "Label");
                WriteSpreadsheetMlCell(writer, spreadsheetNamespace, sourceConnectionName, "MetaValue");
                WriteSpreadsheetMlCell(writer, spreadsheetNamespace, string.Empty, "MetaValue");
                WriteSpreadsheetMlCell(writer, spreadsheetNamespace, "Target", "Label");
                WriteSpreadsheetMlCell(writer, spreadsheetNamespace, targetConnectionName, "MetaValue");
                WriteSpreadsheetMlCell(writer, spreadsheetNamespace, string.Empty, "MetaValue");
                WriteSpreadsheetMlCell(writer, spreadsheetNamespace, "Visible Rows", "Label");
                WriteSpreadsheetMlCell(writer, spreadsheetNamespace, rowsToExport.Count.ToString(), "MetaValue");
                writer.WriteEndElement();

                writer.WriteStartElement("Row", spreadsheetNamespace);
                WriteSpreadsheetMlCell(writer, spreadsheetNamespace, "Exported On", "Label");
                WriteSpreadsheetMlCell(writer, spreadsheetNamespace, DateTime.Now.ToString("yyyy-MM-dd HH:mm"), "MetaValue");
                WriteSpreadsheetMlCell(writer, spreadsheetNamespace, string.Empty, "MetaValue");
                WriteSpreadsheetMlCell(writer, spreadsheetNamespace, "Status Filter", "Label");
                WriteSpreadsheetMlCell(writer, spreadsheetNamespace, activeFilterLabel, "MetaValue");
                WriteSpreadsheetMlCell(writer, spreadsheetNamespace, string.Empty, "MetaValue");
                WriteSpreadsheetMlCell(writer, spreadsheetNamespace, "Changed Only", "Label");
                WriteSpreadsheetMlCell(writer, spreadsheetNamespace, chkChangedOnly != null && chkChangedOnly.Checked ? "Yes" : "No", "MetaValue");
                writer.WriteEndElement();

                writer.WriteStartElement("Row", spreadsheetNamespace);
                WriteSpreadsheetMlCell(writer, spreadsheetNamespace, "Managed/Unmanaged Only", "Label");
                WriteSpreadsheetMlCell(writer, spreadsheetNamespace, chkPackageTypeMismatchOnly != null && chkPackageTypeMismatchOnly.Checked ? "Yes" : "No", "MetaValue");
                writer.WriteEndElement();

                writer.WriteStartElement("Row", spreadsheetNamespace);
                writer.WriteEndElement();

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

                writer.WriteStartElement("Row", spreadsheetNamespace);
                foreach (var header in headers)
                {
                    WriteSpreadsheetMlCell(writer, spreadsheetNamespace, header, "Header");
                }
                writer.WriteEndElement();

                foreach (var item in rowsToExport)
                {
                    writer.WriteStartElement("Row", spreadsheetNamespace);

                    WriteSpreadsheetMlCell(writer, spreadsheetNamespace, NormalizeCsvValue(item.UniqueName), "Text");
                    WriteSpreadsheetMlCell(writer, spreadsheetNamespace, NormalizeCsvValue(item.SourceDisplayName), "TextWrap");
                    WriteSpreadsheetMlCell(writer, spreadsheetNamespace, NormalizeCsvValue(item.TargetDisplayName), "TextWrap");
                    WriteSpreadsheetMlCell(writer, spreadsheetNamespace, NormalizeCsvValue(item.SourceVersion), "CenterText");
                    WriteSpreadsheetMlCell(writer, spreadsheetNamespace, NormalizeCsvValue(item.TargetVersion), "CenterText");
                    WriteSpreadsheetMlCell(writer, spreadsheetNamespace, NormalizeCsvValue(item.SourcePublisher), "TextWrap");
                    WriteSpreadsheetMlCell(writer, spreadsheetNamespace, NormalizeCsvValue(item.TargetPublisher), "TextWrap");

                    var packageTypeStyle = IsAnyPackageTypeDifference(item) ? "StatusPackageType" : "CenterText";
                    WriteSpreadsheetMlCell(writer, spreadsheetNamespace, NormalizeCsvValue(item.SourcePackageType), packageTypeStyle);
                    WriteSpreadsheetMlCell(writer, spreadsheetNamespace, NormalizeCsvValue(item.TargetPackageType), packageTypeStyle);
                    WriteSpreadsheetMlCell(writer, spreadsheetNamespace, NormalizeCsvValue(item.PackageTypeStatus), GetSpreadsheetMlStatusStyleId(item.PackageTypeStatus));
                    WriteSpreadsheetMlCell(writer, spreadsheetNamespace, NormalizeCsvValue(item.Status), GetSpreadsheetMlStatusStyleId(item.Status));

                    writer.WriteEndElement();
                }

                writer.WriteEndElement();

                writer.WriteStartElement("WorksheetOptions", excelNamespace);
                writer.WriteStartElement("FreezePanes", excelNamespace);
                writer.WriteEndElement();
                writer.WriteStartElement("FrozenNoSplit", excelNamespace);
                writer.WriteEndElement();
                writer.WriteElementString("SplitHorizontal", excelNamespace, "6");
                writer.WriteElementString("TopRowBottomPane", excelNamespace, "6");
                writer.WriteElementString("ActivePane", excelNamespace, "2");
                writer.WriteEndElement();

                writer.WriteEndElement();
                writer.WriteEndElement();
                writer.WriteEndDocument();
            }
        }

        private void WriteSpreadsheetMlStyles(XmlWriter writer, string spreadsheetNamespace)
        {
            writer.WriteStartElement("Styles", spreadsheetNamespace);

            WriteSpreadsheetMlStyle(writer, spreadsheetNamespace, "Default", "Vertical", null, false, null, null, null, false);
            WriteSpreadsheetMlStyle(writer, spreadsheetNamespace, "Title", "Left", "14", true, "#D9E2F3", null, null, false);
            WriteSpreadsheetMlStyle(writer, spreadsheetNamespace, "Label", "Left", null, true, "#EEF3F8", null, null, false);
            WriteSpreadsheetMlStyle(writer, spreadsheetNamespace, "MetaValue", "Left", null, false, null, null, null, false);
            WriteSpreadsheetMlStyle(writer, spreadsheetNamespace, "Header", "Center", null, true, "#44546A", "#FFFFFF", null, true);
            WriteSpreadsheetMlStyle(writer, spreadsheetNamespace, "Text", "Left", null, false, null, null, null, false);
            WriteSpreadsheetMlStyle(writer, spreadsheetNamespace, "TextWrap", "Left", null, false, null, null, null, true);
            WriteSpreadsheetMlStyle(writer, spreadsheetNamespace, "CenterText", "Center", null, false, null, null, null, false);
            WriteSpreadsheetMlStyle(writer, spreadsheetNamespace, "StatusMatch", "Center", null, true, "#F0FFF0", "#008000", null, false);
            WriteSpreadsheetMlStyle(writer, spreadsheetNamespace, "StatusVersion", "Center", null, true, "#FFE4B5", "#FF8C00", null, false);
            WriteSpreadsheetMlStyle(writer, spreadsheetNamespace, "StatusPublisher", "Center", null, true, "#E6E6FA", "#9400D3", null, false);
            WriteSpreadsheetMlStyle(writer, spreadsheetNamespace, "StatusDisplayName", "Center", null, true, "#F0F8FF", "#4682B4", null, false);
            WriteSpreadsheetMlStyle(writer, spreadsheetNamespace, "StatusPackageType", "Center", null, true, "#E0FFFF", "#008080", null, false);
            WriteSpreadsheetMlStyle(writer, spreadsheetNamespace, "StatusMultiple", "Center", null, true, "#FFE4E1", "#8B008B", null, false);
            WriteSpreadsheetMlStyle(writer, spreadsheetNamespace, "StatusMissingSource", "Center", null, true, "#FFE4E1", "#FF0000", null, false);
            WriteSpreadsheetMlStyle(writer, spreadsheetNamespace, "StatusMissingTarget", "Center", null, true, "#FAF0E6", "#B22222", null, false);

            writer.WriteEndElement();
        }

        private void WriteSpreadsheetMlStyle(
            XmlWriter writer,
            string spreadsheetNamespace,
            string styleId,
            string horizontalAlignment,
            string fontSize,
            bool bold,
            string backgroundColor,
            string fontColor,
            string numberFormat,
            bool wrapText)
        {
            writer.WriteStartElement("Style", spreadsheetNamespace);
            writer.WriteAttributeString("ss", "ID", spreadsheetNamespace, styleId);

            writer.WriteStartElement("Alignment", spreadsheetNamespace);
            if (!string.IsNullOrWhiteSpace(horizontalAlignment) && !string.Equals(horizontalAlignment, "Vertical", StringComparison.OrdinalIgnoreCase))
            {
                writer.WriteAttributeString("ss", "Horizontal", spreadsheetNamespace, horizontalAlignment);
            }
            writer.WriteAttributeString("ss", "Vertical", spreadsheetNamespace, "Center");
            if (wrapText)
            {
                writer.WriteAttributeString("ss", "WrapText", spreadsheetNamespace, "1");
            }
            writer.WriteEndElement();

            writer.WriteStartElement("Borders", spreadsheetNamespace);
            WriteSpreadsheetMlBorder(writer, spreadsheetNamespace, "Bottom");
            WriteSpreadsheetMlBorder(writer, spreadsheetNamespace, "Left");
            WriteSpreadsheetMlBorder(writer, spreadsheetNamespace, "Right");
            WriteSpreadsheetMlBorder(writer, spreadsheetNamespace, "Top");
            writer.WriteEndElement();

            writer.WriteStartElement("Font", spreadsheetNamespace);
            if (!string.IsNullOrWhiteSpace(fontSize))
            {
                writer.WriteAttributeString("ss", "Size", spreadsheetNamespace, fontSize);
            }
            if (bold)
            {
                writer.WriteAttributeString("ss", "Bold", spreadsheetNamespace, "1");
            }
            if (!string.IsNullOrWhiteSpace(fontColor))
            {
                writer.WriteAttributeString("ss", "Color", spreadsheetNamespace, fontColor);
            }
            writer.WriteEndElement();

            if (!string.IsNullOrWhiteSpace(backgroundColor))
            {
                writer.WriteStartElement("Interior", spreadsheetNamespace);
                writer.WriteAttributeString("ss", "Color", spreadsheetNamespace, backgroundColor);
                writer.WriteAttributeString("ss", "Pattern", spreadsheetNamespace, "Solid");
                writer.WriteEndElement();
            }

            if (!string.IsNullOrWhiteSpace(numberFormat))
            {
                writer.WriteStartElement("NumberFormat", spreadsheetNamespace);
                writer.WriteAttributeString("ss", "Format", spreadsheetNamespace, numberFormat);
                writer.WriteEndElement();
            }

            writer.WriteEndElement();
        }

        private void WriteSpreadsheetMlBorder(XmlWriter writer, string spreadsheetNamespace, string position)
        {
            writer.WriteStartElement("Border", spreadsheetNamespace);
            writer.WriteAttributeString("ss", "Position", spreadsheetNamespace, position);
            writer.WriteAttributeString("ss", "LineStyle", spreadsheetNamespace, "Continuous");
            writer.WriteAttributeString("ss", "Weight", spreadsheetNamespace, "1");
            writer.WriteAttributeString("ss", "Color", spreadsheetNamespace, "#D9D9D9");
            writer.WriteEndElement();
        }

        private void WriteSpreadsheetMlColumn(XmlWriter writer, string spreadsheetNamespace, double width)
        {
            writer.WriteStartElement("Column", spreadsheetNamespace);
            writer.WriteAttributeString("ss", "AutoFitWidth", spreadsheetNamespace, "0");
            writer.WriteAttributeString("ss", "Width", spreadsheetNamespace, width.ToString(System.Globalization.CultureInfo.InvariantCulture));
            writer.WriteEndElement();
        }

        private void WriteSpreadsheetMlCell(
            XmlWriter writer,
            string spreadsheetNamespace,
            string value,
            string styleId,
            int mergeAcross = 0)
        {
            writer.WriteStartElement("Cell", spreadsheetNamespace);

            if (!string.IsNullOrWhiteSpace(styleId))
            {
                writer.WriteAttributeString("ss", "StyleID", spreadsheetNamespace, styleId);
            }

            if (mergeAcross > 0)
            {
                writer.WriteAttributeString("ss", "MergeAcross", spreadsheetNamespace, mergeAcross.ToString());
            }

            writer.WriteStartElement("Data", spreadsheetNamespace);
            writer.WriteAttributeString("ss", "Type", spreadsheetNamespace, "String");
            writer.WriteString(value ?? string.Empty);
            writer.WriteEndElement();

            writer.WriteEndElement();
        }

        private string GetSpreadsheetMlStatusStyleId(string status)
        {
            switch ((status ?? string.Empty).Trim())
            {
                case "Match":
                    return "StatusMatch";
                case "Version Mismatch":
                    return "StatusVersion";
                case "Publisher Mismatch":
                    return "StatusPublisher";
                case "Display Name Mismatch":
                    return "StatusDisplayName";
                case "Package Type Mismatch":
                case "Managed/Unmanaged Mismatch":
                    return "StatusPackageType";
                case "Multiple Differences":
                    return "StatusMultiple";
                case "Missing in Source":
                    return "StatusMissingSource";
                case "Missing in Target":
                    return "StatusMissingTarget";
                default:
                    return "CenterText";
            }
        }

        private void ExportVisibleRowsToCsv(string filePath)
        {
            var rowsToExport = GetVisibleComparisonRows();
            var builder = new StringBuilder();

            var activeFilterLabel = btnFilter != null
                ? btnFilter.Text.Replace("Filter: ", string.Empty)
                : "All";

            builder.AppendLine("sep=,");
            builder.AppendLine(string.Join(",", EscapeCsvValue("D365 Solution Comparer Export")));
            builder.AppendLine(string.Join(",", EscapeCsvValue("Source"), EscapeCsvValue(sourceConnectionName), EscapeCsvValue(string.Empty), EscapeCsvValue("Target"), EscapeCsvValue(targetConnectionName), EscapeCsvValue(string.Empty), EscapeCsvValue("Visible Rows"), EscapeCsvValue(rowsToExport.Count.ToString())));
            builder.AppendLine(string.Join(",", EscapeCsvValue("Exported On"), EscapeCsvValue(DateTime.Now.ToString("yyyy-MM-dd HH:mm")), EscapeCsvValue(string.Empty), EscapeCsvValue("Status Filter"), EscapeCsvValue(activeFilterLabel), EscapeCsvValue(string.Empty), EscapeCsvValue("Changed Only"), EscapeCsvValue(chkChangedOnly != null && chkChangedOnly.Checked ? "Yes" : "No")));
            builder.AppendLine(string.Join(",", EscapeCsvValue("Managed/Unmanaged Only"), EscapeCsvValue(chkPackageTypeMismatchOnly != null && chkPackageTypeMismatchOnly.Checked ? "Yes" : "No")));
            builder.AppendLine();

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

            builder.AppendLine(string.Join(",", headers.Select(EscapeCsvValue)));

            foreach (var item in rowsToExport)
            {
                var values = new[]
                {
                    NormalizeCsvValue(item.UniqueName),
                    NormalizeCsvValue(item.SourceDisplayName),
                    NormalizeCsvValue(item.TargetDisplayName),
                    NormalizeCsvValue(item.SourceVersion),
                    NormalizeCsvValue(item.TargetVersion),
                    NormalizeCsvValue(item.SourcePublisher),
                    NormalizeCsvValue(item.TargetPublisher),
                    NormalizeCsvValue(item.SourcePackageType),
                    NormalizeCsvValue(item.TargetPackageType),
                    NormalizeCsvValue(item.PackageTypeStatus),
                    NormalizeCsvValue(item.Status)
                };

                builder.AppendLine(string.Join(",", values.Select(EscapeCsvValue)));
            }

            File.WriteAllText(filePath, builder.ToString(), new UTF8Encoding(true));
        }

        private List<CompareResult> GetVisibleComparisonRows()
        {
            return dgvResults.Rows
                .Cast<DataGridViewRow>()
                .Where(r => !r.IsNewRow && r.Visible)
                .Select(r => r.DataBoundItem as CompareResult)
                .Where(r => r != null)
                .ToList();
        }

        private string EscapeCsvValue(string value)
        {
            var safe = value ?? string.Empty;
            safe = safe.Replace("\"", "\"\"");
            return "\"" + safe + "\"";
        }

        private string NormalizeCsvValue(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return string.Empty;
            }

            return value
                .Replace("\r\n", " / ")
                .Replace("\n", " / ")
                .Replace("\r", " / ")
                .Replace("\t", " ")
                .Trim();
        }

        private string GetExceptionSummary(Exception ex)
        {
            if (ex == null)
            {
                return "Unknown error.";
            }

            var messages = new List<string>();
            var current = ex;

            while (current != null && messages.Count < 3)
            {
                if (!string.IsNullOrWhiteSpace(current.Message))
                {
                    messages.Add(current.Message.Trim());
                }

                current = current.InnerException;
            }

            var distinctMessages = messages
                .Where(m => !string.IsNullOrWhiteSpace(m))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();

            if (distinctMessages.Count == 0)
            {
                return ex.GetType().FullName;
            }

            return string.Join("\n\n", distinctMessages);
        }

        private void ShowRowDetails(CompareResult item)
        {
            var details =
                "Solution comparison details\r\n\r\n" +
                "Unique Name: " + Safe(item.UniqueName) + "\r\n" +
                "Source Display Name: " + Safe(item.SourceDisplayName) + "\r\n" +
                "Target Display Name: " + Safe(item.TargetDisplayName) + "\r\n" +
                "Source Version: " + Safe(item.SourceVersion) + "\r\n" +
                "Target Version: " + Safe(item.TargetVersion) + "\r\n" +
                "Source Publisher: " + Safe(item.SourcePublisher) + "\r\n" +
                "Target Publisher: " + Safe(item.TargetPublisher) + "\r\n" +
                "Source Package Type: " + Safe(item.SourcePackageType) + "\r\n" +
                "Target Package Type: " + Safe(item.TargetPackageType) + "\r\n" +
                "Package Type Status: " + Safe(item.PackageTypeStatus) + "\r\n" +
                "Overall Status: " + Safe(item.Status) + "\r\n\r\n" +
                "Changed: " + (IsChangedResult(item) ? "Yes" : "No") + "\r\n" +
                "Managed/Unmanaged Difference: " + (IsManagedUnmanagedDifference(item) ? "Yes" : "No") + "\r\n" +
                "Any Package Type Difference: " + (IsAnyPackageTypeDifference(item) ? "Yes" : "No");

            using (var detailsForm = new Form())
            using (var txtDetails = new TextBox())
            using (var bottomPanel = new Panel())
            using (var btnClose = new Button())
            {
                detailsForm.Text = "Row Details";
                detailsForm.StartPosition = FormStartPosition.CenterParent;
                detailsForm.Size = new Size(760, 520);
                detailsForm.MinimumSize = new Size(680, 420);
                detailsForm.MaximizeBox = false;
                detailsForm.MinimizeBox = false;
                detailsForm.ShowInTaskbar = false;
                detailsForm.FormBorderStyle = FormBorderStyle.Sizable;
                detailsForm.BackColor = Color.White;

                txtDetails.Multiline = true;
                txtDetails.ReadOnly = true;
                txtDetails.ScrollBars = ScrollBars.Both;
                txtDetails.WordWrap = false;
                txtDetails.Dock = DockStyle.Fill;
                txtDetails.Font = new Font("Consolas", 10F);
                txtDetails.BackColor = Color.White;
                txtDetails.Text = details;

                bottomPanel.Dock = DockStyle.Bottom;
                bottomPanel.Height = 52;
                bottomPanel.Padding = new Padding(10);
                bottomPanel.BackColor = Color.WhiteSmoke;

                btnClose.Text = "Close";
                btnClose.Width = 100;
                btnClose.Height = 30;
                btnClose.Anchor = AnchorStyles.Right | AnchorStyles.Top;
                btnClose.Left = bottomPanel.Width - btnClose.Width - 10;
                btnClose.Top = 10;
                btnClose.DialogResult = DialogResult.OK;

                bottomPanel.Controls.Add(btnClose);
                bottomPanel.Resize += (sender, args) =>
                {
                    btnClose.Left = bottomPanel.ClientSize.Width - btnClose.Width;
                };

                detailsForm.Controls.Add(txtDetails);
                detailsForm.Controls.Add(bottomPanel);
                detailsForm.AcceptButton = btnClose;
                detailsForm.CancelButton = btnClose;

                detailsForm.ShowDialog(this);
            }
        }

        private string Safe(string value)
        {
            return string.IsNullOrWhiteSpace(value) ? "(blank)" : value;
        }

        private string GetProductVersion()
        {
            try
            {
                var version = Assembly.GetExecutingAssembly().GetName().Version;
                return version != null ? version.ToString() : "unknown";
            }
            catch
            {
                return "unknown";
            }
        }

        private void ApplySavedFilterState()
        {
            suppressSettingsSave = true;

            try
            {
                var settings = userSettings ?? new Settings();

                miAll.Checked = settings.FilterAll;
                miMatch.Checked = settings.FilterMatch;
                miVersionMismatch.Checked = settings.FilterVersionMismatch;
                miPublisherMismatch.Checked = settings.FilterPublisherMismatch;
                miDisplayNameMismatch.Checked = settings.FilterDisplayNameMismatch;
                miPackageTypeDifference.Checked = settings.FilterPackageTypeDifference;
                miMultipleDifferences.Checked = settings.FilterMultipleDifferences;
                miMissingInSource.Checked = settings.FilterMissingInSource;
                miMissingInTarget.Checked = settings.FilterMissingInTarget;

                chkPackageTypeMismatchOnly.Checked = settings.ShowManagedUnmanagedOnly;
                chkChangedOnly.Checked = settings.ShowChangedOnly;

                if (!miAll.Checked && !AnySpecificFilterChecked())
                {
                    miAll.Checked = true;
                }

                UpdateFilterButtonText();
            }
            finally
            {
                suppressSettingsSave = false;
            }
        }

        private void PersistUiSettings()
        {
            if (suppressSettingsSave)
            {
                return;
            }

            if (userSettings == null)
            {
                userSettings = new Settings();
            }

            userSettings.FilterAll = miAll.Checked;
            userSettings.FilterMatch = miMatch.Checked;
            userSettings.FilterVersionMismatch = miVersionMismatch.Checked;
            userSettings.FilterPublisherMismatch = miPublisherMismatch.Checked;
            userSettings.FilterDisplayNameMismatch = miDisplayNameMismatch.Checked;
            userSettings.FilterPackageTypeDifference = miPackageTypeDifference.Checked;
            userSettings.FilterMultipleDifferences = miMultipleDifferences.Checked;
            userSettings.FilterMissingInSource = miMissingInSource.Checked;
            userSettings.FilterMissingInTarget = miMissingInTarget.Checked;
            userSettings.ShowManagedUnmanagedOnly = chkPackageTypeMismatchOnly.Checked;
            userSettings.ShowChangedOnly = chkChangedOnly.Checked;

            userSettings.Save();
        }
    }
}
