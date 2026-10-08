using System;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using D365SolutionComparer.Models.Membership;
using D365SolutionComparer.Services.Membership;

namespace D365SolutionComparer
{
    internal sealed class MembershipCoverageDetailsForm : Form
    {
        private readonly MembershipComparisonPresentation presentation;
        private readonly string sourceSolutionVersion;
        private readonly string targetSolutionVersion;
        private readonly ComboBox lifecycleOperation;
        private readonly Action captureAppSettingEvidence;
#if DEBUG
        private readonly Action captureType92Evidence;
#endif

        public MembershipCoverageDetailsForm(string sourceName, MembershipCoverageDiagnostics source,
            string targetName, MembershipCoverageDiagnostics target,
            MembershipComparisonPresentation presentation = null,
            string sourceSolutionVersion = null, string targetSolutionVersion = null,
            Microsoft.Xrm.Sdk.IOrganizationService sourceService = null,
            Microsoft.Xrm.Sdk.IOrganizationService targetService = null,
            Action captureAppSettingEvidence = null
#if DEBUG
            , Action captureType92Evidence = null, Action captureType59Evidence = null,
            Action discoverType59Evidence = null, Action captureCloudFlowEvidence = null,
            Action captureSavedQueryEvidence = null, Action captureUnsupportedCoverageInventory = null,
            Action captureType36Evidence = null, Action captureType31Evidence = null, Action captureType10072Evidence = null, Action captureType300Evidence = null, Action captureType74Evidence = null, Action captureType511Evidence = null, Action captureType10276Evidence = null, Action captureAppActionEvidence = null
#endif
            )
        {
            if (source == null) throw new ArgumentNullException(nameof(source));
            if (target == null) throw new ArgumentNullException(nameof(target));
            this.presentation = presentation;
            this.sourceSolutionVersion = sourceSolutionVersion ?? string.Empty;
            this.targetSolutionVersion = targetSolutionVersion ?? string.Empty;
            this.captureAppSettingEvidence = captureAppSettingEvidence;
#if DEBUG
            this.captureType92Evidence = captureType92Evidence;
#endif
            Text = "Membership Coverage Details";
            StartPosition = FormStartPosition.CenterParent;
            MinimumSize = new Size(800, 480);
            Size = new Size(1100, 700);
            Font = new Font("Segoe UI", 9F);

            var explanation = new Label
            {
                Dock = DockStyle.Top,
                Height = 52,
                Padding = new Padding(10, 8, 10, 4),
                Text = "Same-kind blockers make only that semantic kind incomplete. Isolated unsupported " +
                    "component-type buckets do not block other kinds. Broad / Unclassifiable blockers make " +
                    "every semantic kind incomplete."
            };
            var tabs = new TabControl { Dock = DockStyle.Fill };
            tabs.TabPages.Add(CreatePage("Source", sourceName, source));
            tabs.TabPages.Add(CreatePage("Target", targetName, target));
            lifecycleOperation = new ComboBox
            {
                Width = 430,
                DropDownStyle = ComboBoxStyle.DropDown
            };
            lifecycleOperation.Items.AddRange(new object[]
            {
                "Repeated import of the same unmanaged solution",
                "Updated Canvas App import",
                "Unmanaged DEV to managed UAT deployment",
                "Managed Update",
                "Managed Upgrade",
                "Patch followed by upgrade",
                "Canvas App rename",
                "Solution clone or supported equivalent",
                "Delete and recreate",
                "Same/similar app names under different publisher prefixes"
            });
            var lifecyclePanel = new FlowLayoutPanel
            {
                Dock = DockStyle.Top,
                Height = 38,
                Padding = new Padding(8, 5, 8, 3),
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = false
            };
            lifecyclePanel.Controls.Add(new Label
            {
                AutoSize = true,
                Margin = new Padding(2, 5, 8, 0),
                Text = "Lifecycle operation:"
            });
            lifecyclePanel.Controls.Add(lifecycleOperation);
            var close = new Button { Text = "Close", DialogResult = DialogResult.OK, Width = 90 };
            var export = new Button
            {
                Text = "Export CSV...",
                Width = 110,
                Enabled = presentation != null
            };
            export.Click += Export_Click;
            var compare = new Button
            {
                Text = "Compare lifecycle CSVs...",
                Width = 165
            };
            compare.Click += CompareLifecycleCsvs_Click;
            var appSetting = new Button
            {
                Text = "Capture AppSetting Evidence...",
                Width = 170,
                Enabled = captureAppSettingEvidence != null && sourceService != null && targetService != null &&
                    presentation != null && presentation.Source.Snapshot != null &&
                    presentation.Target.Snapshot != null &&
                    presentation.Source.Snapshot.State == MembershipSnapshotState.Complete &&
                    presentation.Target.Snapshot.State == MembershipSnapshotState.Complete
            };
            appSetting.Click += (sender, args) => this.captureAppSettingEvidence?.Invoke();
            var buttons = new FlowLayoutPanel
            {
                Dock = DockStyle.Bottom,
                Height = 42,
                Padding = new Padding(6),
                FlowDirection = FlowDirection.RightToLeft
            };
            buttons.Controls.Add(close);
            buttons.Controls.Add(appSetting);
#if DEBUG
            // Evidence actions may wrap on smaller windows; keep every action reachable.
            buttons.AutoSize = true;
            buttons.AutoSizeMode = AutoSizeMode.GrowAndShrink;
            var type92 = new Button
            {
                Text = "Capture Type 92 Evidence...",
                Width = 175,
                Enabled = captureType92Evidence != null && sourceService != null && targetService != null &&
                    presentation != null && presentation.Source.Snapshot?.State == MembershipSnapshotState.Complete &&
                    presentation.Target.Snapshot?.State == MembershipSnapshotState.Complete
            };
            type92.Click += (sender, args) => this.captureType92Evidence?.Invoke();
            buttons.Controls.Add(type92);
            var type59 = new Button
            {
                Text = "Capture Type 59 Evidence...",
                Width = 175,
                Enabled = captureType59Evidence != null && sourceService != null && targetService != null &&
                    presentation?.Source.Snapshot?.State == MembershipSnapshotState.Complete &&
                    presentation.Target.Snapshot?.State == MembershipSnapshotState.Complete
            };
            var type59Menu = new ContextMenuStrip();
            type59Menu.Items.Add("Capture selected solution", null, (sender, args) => captureType59Evidence?.Invoke());
            var discovery = type59Menu.Items.Add("Scan shared solutions for portability candidates", null,
                (sender, args) => discoverType59Evidence?.Invoke());
            discovery.Enabled = discoverType59Evidence != null;
            type59.Click += (sender, args) => type59Menu.Show(type59, new Point(0, type59.Height));
            type59.Disposed += (sender, args) => type59Menu.Dispose();
            buttons.Controls.Add(type59);
            var processQuery = new Button
            {
                Text = "Flow / View Evidence...",
                Width = 160,
                Enabled = sourceService != null && targetService != null &&
                    presentation?.Source.Snapshot?.State == MembershipSnapshotState.Complete &&
                    presentation.Target.Snapshot?.State == MembershipSnapshotState.Complete &&
                    (captureCloudFlowEvidence != null || captureSavedQueryEvidence != null)
            };
            var processQueryMenu = new ContextMenuStrip();
            var flow = processQueryMenu.Items.Add("Capture Cloud Flow evidence (all Type 29)", null,
                (sender, args) => captureCloudFlowEvidence?.Invoke());
            flow.Enabled = captureCloudFlowEvidence != null;
            var view = processQueryMenu.Items.Add("Capture Type 26 Saved Query evidence", null,
                (sender, args) => captureSavedQueryEvidence?.Invoke());
            view.Enabled = captureSavedQueryEvidence != null;
            processQuery.Click += (sender, args) => processQueryMenu.Show(processQuery, new Point(0, processQuery.Height));
            processQuery.Disposed += (sender, args) => processQueryMenu.Dispose();
            buttons.Controls.Add(processQuery);
            var inventory = new Button
            {
                Text = "Unsupported Coverage Inventory...",
                Width = 225,
                Enabled = captureUnsupportedCoverageInventory != null
            };
            inventory.Click += (sender, args) => captureUnsupportedCoverageInventory?.Invoke();
            buttons.Controls.Add(inventory);
            var type36 = new Button
            {
                Text = "Capture Type 36 Email Template Evidence...",
                Width = 285,
                Enabled = captureType36Evidence != null &&
                    presentation?.Source.Snapshot?.State == MembershipSnapshotState.Complete &&
                    presentation.Target.Snapshot?.State == MembershipSnapshotState.Complete
            };
            type36.Click += (sender, args) => captureType36Evidence?.Invoke();
            buttons.Controls.Add(type36);
            var type31 = new Button
            {
                Text = "Capture Type 31 Report Evidence...",
                Width = 255,
                Enabled = captureType31Evidence != null &&
                    presentation?.Source.Snapshot?.State == MembershipSnapshotState.Complete &&
                    presentation.Target.Snapshot?.State == MembershipSnapshotState.Complete
            };
            type31.Click += (sender, args) => captureType31Evidence?.Invoke();
            buttons.Controls.Add(type31);
            var type10072 = new Button
            {
                Text = "Capture Type 10072 AppElement Evidence...",
                Width = 255,
                Enabled = captureType10072Evidence != null &&
                    presentation?.Source.Snapshot?.State == MembershipSnapshotState.Complete &&
                    presentation.Target.Snapshot?.State == MembershipSnapshotState.Complete
            };
            type10072.Click += (sender, args) => captureType10072Evidence?.Invoke();
            buttons.Controls.Add(type10072);
            var type300 = new Button
            {
                Text = "Capture Type 300 Canvas App Evidence...",
                Width = 255,
                Enabled = captureType300Evidence != null &&
                    presentation?.Source.Snapshot?.State == MembershipSnapshotState.Complete &&
                    presentation.Target.Snapshot?.State == MembershipSnapshotState.Complete
            };
            type300.Click += (sender, args) => captureType300Evidence?.Invoke();
            buttons.Controls.Add(type300);
            var type74 = new Button
            {
                Text = "Capture Type 74 MaskingRule Evidence...",
                Width = 255,
                Enabled = captureType74Evidence != null &&
                    presentation?.Source.Snapshot?.State == MembershipSnapshotState.Complete &&
                    presentation.Target.Snapshot?.State == MembershipSnapshotState.Complete
            };
            type74.Click += (sender, args) => captureType74Evidence?.Invoke();
            buttons.Controls.Add(type74);
            var type511 = new Button
            {
                Text = "Capture Type 511 Team Template Evidence...",
                Width = 255,
                Enabled = captureType511Evidence != null &&
                    presentation?.Source.Snapshot?.State == MembershipSnapshotState.Complete &&
                    presentation.Target.Snapshot?.State == MembershipSnapshotState.Complete
            };
            type511.Click += (sender, args) => captureType511Evidence?.Invoke();
            buttons.Controls.Add(type511);
            var type10276 = new Button
            {
                Text = "Capture Type 10276 AI Skill Config Evidence...",
                Width = 255,
                Enabled = captureType10276Evidence != null &&
                    presentation?.Source.Snapshot?.State == MembershipSnapshotState.Complete &&
                    presentation.Target.Snapshot?.State == MembershipSnapshotState.Complete
            };
            type10276.Click += (sender, args) => captureType10276Evidence?.Invoke();
            buttons.Controls.Add(type10276);
            var appAction = new Button
            {
                Text = "Capture Types 10266 / 10267 App Action Evidence...",
                Width = 300,
                Enabled = captureAppActionEvidence != null &&
                    presentation?.Source.Snapshot?.State == MembershipSnapshotState.Complete &&
                    presentation.Target.Snapshot?.State == MembershipSnapshotState.Complete
            };
            appAction.Click += (sender, args) => captureAppActionEvidence?.Invoke();
            buttons.Controls.Add(appAction);
#endif
            buttons.Controls.Add(export);
            buttons.Controls.Add(compare);
            Controls.Add(tabs);
            Controls.Add(buttons);
            Controls.Add(lifecyclePanel);
            Controls.Add(explanation);
            AcceptButton = close;
            CancelButton = close;
        }

        private void CompareLifecycleCsvs_Click(object sender, EventArgs e)
        {
            string beforePath = SelectCsv("Select the Before membership coverage CSV");
            if (beforePath == null) return;
            string afterPath = SelectCsv("Select the After membership coverage CSV");
            if (afterPath == null) return;
            using (var dialog = new SaveFileDialog
            {
                Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*",
                DefaultExt = "csv",
                AddExtension = true,
                FileName = "canvas-app-lifecycle-comparison-" +
                    DateTime.UtcNow.ToString("yyyyMMdd-HHmmssfff", CultureInfo.InvariantCulture) + ".csv",
                Title = "Save Canvas App Lifecycle Comparison"
            })
            {
                if (dialog.ShowDialog(this) != DialogResult.OK) return;
                try
                {
                    new CanvasAppLifecycleCsvComparer().CompareFiles(beforePath, afterPath, dialog.FileName);
                    MessageBox.Show(this, "Canvas App lifecycle evidence was compared successfully.",
                        "Compare Lifecycle Evidence", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                catch (CanvasAppLifecycleChronologyException ex)
                {
                    MessageBox.Show(this, ex.Message, "Compare Lifecycle Evidence",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(this, "Failed to compare Canvas App lifecycle evidence.\n\n" + ex.Message,
                        "Compare Lifecycle Evidence", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private string SelectCsv(string title)
        {
            using (var dialog = new OpenFileDialog
            {
                Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*",
                CheckFileExists = true,
                Multiselect = false,
                Title = title
            })
                return dialog.ShowDialog(this) == DialogResult.OK ? dialog.FileName : null;
        }

        private void Export_Click(object sender, EventArgs e)
        {
            var operation = lifecycleOperation.Text == null ? string.Empty : lifecycleOperation.Text.Trim();
            if (operation.Length == 0)
            {
                MessageBox.Show(this, "Enter or select the lifecycle operation for this evidence checkpoint.",
                    "Export Membership Coverage", MessageBoxButtons.OK, MessageBoxIcon.Information);
                lifecycleOperation.Focus();
                return;
            }
            using (var dialog = new SaveFileDialog
            {
                Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*",
                DefaultExt = "csv",
                AddExtension = true,
                FileName = SafeFileName(presentation.SolutionUniqueName) + "-" + SafeFileName(operation) +
                    "-" + DateTime.UtcNow.ToString("yyyyMMdd-HHmmssfff", CultureInfo.InvariantCulture) +
                    "-membership-coverage.csv",
                Title = "Export Membership Coverage Details"
            })
            {
                if (dialog.ShowDialog(this) != DialogResult.OK) return;
                try
                {
                    new MembershipCoverageCsvExporter().WriteCsv(dialog.FileName, presentation,
                        sourceSolutionVersion, targetSolutionVersion, operation);
                    MessageBox.Show(this, "Membership coverage details were exported successfully.",
                        "Export Membership Coverage", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(this, "Failed to export membership coverage details.\n\n" + ex.Message,
                        "Export Membership Coverage", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private static string SafeFileName(string value)
        {
            var invalid = System.IO.Path.GetInvalidFileNameChars();
            return new string((value ?? "solution").Select(character =>
                invalid.Contains(character) ? '_' : character).ToArray());
        }

        private static TabPage CreatePage(string side, string environmentName,
            MembershipCoverageDiagnostics diagnostics)
        {
            var page = new TabPage(side + " - " + (environmentName ?? string.Empty));
            page.Controls.Add(new RichTextBox
            {
                Dock = DockStyle.Fill,
                ReadOnly = true,
                WordWrap = false,
                DetectUrls = false,
                BackColor = Color.White,
                Font = new Font("Consolas", 9F),
                Text = Format(environmentName, diagnostics)
            });
            return page;
        }

        private static string Format(string environmentName, MembershipCoverageDiagnostics diagnostics)
        {
            var text = new StringBuilder();
            text.AppendLine("Environment: " + (environmentName ?? string.Empty));
            text.AppendLine("Snapshot state: " + diagnostics.SnapshotState);
            text.AppendLine("Broad / Unclassifiable blockers: " +
                diagnostics.BroadUnclassifiable.TotalCandidates.ToString(CultureInfo.InvariantCulture));
            text.AppendLine();
            text.AppendLine(string.Format(CultureInfo.InvariantCulture,
                "{0,-43} {1,-12} {2,6} {3,6} {4,6} {5,6} {6,6} {7,-10}",
                "Kind / Bucket", "Scope", "Total", "Res", "Unsup", "Unres", "Ambig", "Coverage"));
            text.AppendLine(new string('-', 108));
            foreach (var bucket in diagnostics.SemanticKinds)
                AppendBucket(text, bucket);
            AppendBucket(text, diagnostics.BroadUnclassifiable);

            text.AppendLine();
            text.AppendLine("Broad / Unclassifiable raw component types:");
            if (diagnostics.BroadRawComponentTypes.Count == 0)
                text.AppendLine("None");
            foreach (var rawType in diagnostics.BroadRawComponentTypes)
            {
                text.Append("Raw ComponentType ").Append(rawType.ComponentType.ToString(CultureInfo.InvariantCulture))
                    .Append("  Count=").AppendLine(rawType.Count.ToString(CultureInfo.InvariantCulture));
                foreach (var group in rawType.DiagnosticGroups)
                {
                    AppendDiagnostic(text, "  ", group);
                    foreach (var evidence in rawType.Evidence.Where(item =>
                        item.ResolutionStatus == group.ResolutionStatus &&
                        string.Equals(item.Diagnostic, group.Diagnostic, StringComparison.Ordinal)))
                        AppendRawEvidence(text, evidence);
                }
            }

            text.AppendLine();
            text.AppendLine("Dynamically classified registered component families:");
            if (diagnostics.DynamicComponentTypes.Count == 0)
                text.AppendLine("None");
            foreach (var dynamicType in diagnostics.DynamicComponentTypes)
            {
                text.Append("Raw ComponentType ")
                    .Append(dynamicType.ComponentType.ToString(CultureInfo.InvariantCulture))
                    .Append("  Count=").Append(dynamicType.Count.ToString(CultureInfo.InvariantCulture))
                    .Append("  Definition=").Append(dynamicType.Definition.Name)
                    .Append("  PrimaryEntity=")
                    .Append(dynamicType.Definition.PrimaryEntityName.Length == 0
                        ? "(not supplied)" : dynamicType.Definition.PrimaryEntityName)
                    .Append("  Bucket=").AppendLine(dynamicType.SemanticKind);
            }

            text.AppendLine();
            text.AppendLine("Same-kind and isolated blocker diagnostics (original diagnostic text):");
            bool wroteDiagnostic = false;
            foreach (var bucket in diagnostics.SemanticKinds)
            {
                foreach (var group in bucket.DiagnosticGroups)
                {
                    wroteDiagnostic = true;
                    AppendDiagnostic(text, "[" + bucket.DisplayName + "] ", group);
                }
            }
            if (!wroteDiagnostic) text.AppendLine("None");

            text.AppendLine();
            text.AppendLine("Component identity diagnostic evidence:");
            bool wroteEvidence = false;
            foreach (var bucket in diagnostics.SemanticKinds.Where(item => item.AuditEvidence.Count > 0))
            {
                text.AppendLine("[" + bucket.DisplayName + "]");
                foreach (var evidence in bucket.AuditEvidence)
                {
                    wroteEvidence = true;
                    AppendRawEvidence(text, evidence);
                }
            }
            if (!wroteEvidence) text.AppendLine("None");
            return text.ToString();
        }

        private static void AppendDiagnostic(StringBuilder text, string prefix,
            MembershipCoverageDiagnosticGroup group)
        {
            text.Append(prefix).Append(group.ResolutionStatus).Append(" x")
                .Append(group.Count.ToString(CultureInfo.InvariantCulture)).Append(": ")
                .AppendLine(group.Diagnostic.Length == 0 ? "(empty diagnostic)" : group.Diagnostic);
        }

        private static void AppendRawEvidence(StringBuilder text,
            MembershipCoverageRawComponentEvidence evidence)
        {
            text.Append("    solutioncomponentid=").Append(evidence.SolutionComponentId.ToString("D"))
                .Append("  objectid=").AppendLine(evidence.ObjectId.HasValue
                    ? evidence.ObjectId.Value.ToString("D") : "(null)");
            foreach (var detail in evidence.DiagnosticEvidence)
                text.Append("      ").AppendLine(detail);
        }

        private static void AppendBucket(StringBuilder text, MembershipCoverageBucket bucket)
        {
            text.AppendLine(string.Format(CultureInfo.InvariantCulture,
                "{0,-43} {1,-12} {2,6} {3,6} {4,6} {5,6} {6,6} {7,-10}",
                Limit(bucket.DisplayName, 43), Scope(bucket.BucketType), bucket.TotalCandidates,
                bucket.Resolved, bucket.Unsupported, bucket.Unresolved, bucket.Ambiguous,
                bucket.CoverageStatus));
        }

        private static string Scope(MembershipCoverageBucketType bucketType)
        {
            switch (bucketType)
            {
                case MembershipCoverageBucketType.SemanticKind: return "Same kind";
                case MembershipCoverageBucketType.KnownUnsupportedIsolatedType: return "Isolated";
                case MembershipCoverageBucketType.DynamicallyClassifiedIsolatedFamily: return "Dynamic";
                default: return "Broad";
            }
        }

        private static string Limit(string value, int length) => value.Length <= length
            ? value : value.Substring(0, length - 3) + "...";
    }
}
