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

        public MembershipCoverageDetailsForm(string sourceName, MembershipCoverageDiagnostics source,
            string targetName, MembershipCoverageDiagnostics target,
            MembershipComparisonPresentation presentation = null,
            string sourceSolutionVersion = null, string targetSolutionVersion = null)
        {
            if (source == null) throw new ArgumentNullException(nameof(source));
            if (target == null) throw new ArgumentNullException(nameof(target));
            this.presentation = presentation;
            this.sourceSolutionVersion = sourceSolutionVersion ?? string.Empty;
            this.targetSolutionVersion = targetSolutionVersion ?? string.Empty;
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
            var buttons = new FlowLayoutPanel
            {
                Dock = DockStyle.Bottom,
                Height = 42,
                Padding = new Padding(6),
                FlowDirection = FlowDirection.RightToLeft
            };
            buttons.Controls.Add(close);
            buttons.Controls.Add(export);
            Controls.Add(tabs);
            Controls.Add(buttons);
            Controls.Add(lifecyclePanel);
            Controls.Add(explanation);
            AcceptButton = close;
            CancelButton = close;
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
