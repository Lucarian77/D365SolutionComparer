using System;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using D365SolutionComparer.Models.Membership;

namespace D365SolutionComparer
{
    internal sealed class AppSettingEvidenceResultsForm : Form
    {
        private readonly AppSettingEvidenceComparison comparison;

        internal AppSettingEvidenceResultsForm(AppSettingEvidenceComparison comparison)
        {
            this.comparison = comparison ?? throw new ArgumentNullException(nameof(comparison));
            Text = "AppSetting Evidence - Evidence Only";
            StartPosition = FormStartPosition.CenterParent;
            MinimumSize = new Size(900, 560);
            Size = new Size(1280, 780);
            Font = new Font("Segoe UI", 9F);

            var banner = new Label
            {
                Dock = DockStyle.Top,
                Height = 38,
                Padding = new Padding(10, 8, 10, 4),
                Text = "EVIDENCE ONLY - APPSETTING COMPARISON NOT ENABLED",
                ForeColor = Color.DarkOrange,
                Font = new Font("Segoe UI", 9F, FontStyle.Bold)
            };
            var tabs = new TabControl { Dock = DockStyle.Fill };
            tabs.TabPages.Add(Page("Component Types", FormatComponentTypes(comparison.Source, comparison.Target)));
            tabs.TabPages.Add(Page("Candidate Correlations", FormatCandidates(comparison.Source, comparison.Target)));
            tabs.TabPages.Add(Page("Backing Entity Metadata", FormatMetadata(comparison.Source, comparison.Target)));
            tabs.TabPages.Add(Page("Setting Definitions", FormatSettingDefinitions(comparison.Source, comparison.Target)));
            tabs.TabPages.Add(Page("Parent Apps", FormatParentApps(comparison.Source, comparison.Target)));
            tabs.TabPages.Add(Page("DEV vs UAT", FormatComparison(comparison)));
            tabs.TabPages.Add(Page("Request Counts", FormatRequests(comparison.Source, comparison.Target)));
            tabs.TabPages.Add(Page("Diagnostics", FormatDiagnostics(comparison.Source, comparison.Target)));

            var close = new Button { Text = "Close", DialogResult = DialogResult.OK, Width = 90 };
            var save = new Button { Text = "Save Evidence...", Width = 120 };
            save.Click += (sender, args) => SaveEvidence();
            var buttons = new FlowLayoutPanel
            {
                Dock = DockStyle.Bottom,
                Height = 42,
                Padding = new Padding(6),
                FlowDirection = FlowDirection.RightToLeft
            };
            buttons.Controls.Add(close);
            buttons.Controls.Add(save);
            Controls.Add(tabs);
            Controls.Add(buttons);
            Controls.Add(banner);
            AcceptButton = close;
            CancelButton = close;
        }

        private void SaveEvidence()
        {
            using (var dialog = new SaveFileDialog
            {
                Filter = "Text files (*.txt)|*.txt|All files (*.*)|*.*",
                DefaultExt = "txt",
                AddExtension = true,
                FileName = "appsetting-evidence-" + DateTime.UtcNow.ToString("yyyyMMdd-HHmmssfff",
                    CultureInfo.InvariantCulture) + ".txt",
                Title = "Save AppSetting Evidence"
            })
            {
                if (dialog.ShowDialog(this) != DialogResult.OK) return;
                try
                {
                    System.IO.File.WriteAllText(dialog.FileName, BuildFullText(), new UTF8EncodingWithBom());
                    MessageBox.Show(this, "AppSetting evidence was saved.", "AppSetting Evidence",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(this, "Failed to save AppSetting evidence.\n\n" + ex.Message,
                        "AppSetting Evidence", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private string BuildFullText() =>
            "EVIDENCE ONLY - APPSETTING COMPARISON NOT ENABLED" + Environment.NewLine +
            "\nCOMPONENT TYPES\n" + FormatComponentTypes(comparison.Source, comparison.Target) +
            "\nCANDIDATE CORRELATIONS\n" + FormatCandidates(comparison.Source, comparison.Target) +
            "\nBACKING ENTITY METADATA\n" + FormatMetadata(comparison.Source, comparison.Target) +
            "\nSETTING DEFINITIONS\n" + FormatSettingDefinitions(comparison.Source, comparison.Target) +
            "\nPARENT APPS\n" + FormatParentApps(comparison.Source, comparison.Target) +
            "\nDEV VS UAT\n" + FormatComparison(comparison) +
            "\nREQUEST COUNTS\n" + FormatRequests(comparison.Source, comparison.Target) +
            "\nDIAGNOSTICS\n" + FormatDiagnostics(comparison.Source, comparison.Target);

        private static TabPage Page(string title, string text)
        {
            var page = new TabPage(title);
            page.Controls.Add(new RichTextBox
            {
                Dock = DockStyle.Fill,
                ReadOnly = true,
                WordWrap = false,
                DetectUrls = false,
                BackColor = Color.White,
                Font = new Font("Consolas", 9F),
                Text = text ?? string.Empty
            });
            return page;
        }

        private static string FormatComponentTypes(AppSettingEvidenceReport source, AppSettingEvidenceReport target)
        {
            var text = new StringBuilder();
            AppendTypeSummary(text, source);
            AppendTypeSummary(text, target);
            return text.ToString();
        }

        private static void AppendTypeSummary(StringBuilder text, AppSettingEvidenceReport report)
        {
            text.AppendLine("Environment: " + report.EnvironmentName + "  Solution: " + report.SolutionUniqueName);
            text.AppendLine("ComponentType | FormattedLabel | Count | DistinctObjectIds | BlankObjectIds");
            foreach (var item in report.ComponentTypes)
                text.AppendLine(string.Format(CultureInfo.InvariantCulture, "{0} | {1} | {2} | {3} | {4}",
                    item.ComponentType, item.FormattedLabel, item.Count, item.DistinctObjectIdCount,
                    item.BlankObjectIdCount));
            text.AppendLine();
        }

        private static string FormatCandidates(AppSettingEvidenceReport source, AppSettingEvidenceReport target)
        {
            var text = new StringBuilder();
            AppendCandidates(text, source);
            AppendCandidates(text, target);
            return text.ToString();
        }

        private static void AppendCandidates(StringBuilder text, AppSettingEvidenceReport report)
        {
            text.AppendLine("Environment: " + report.EnvironmentName);
            var stats = report.CandidateStatistics;
            text.AppendLine("RawCandidateCount=" + stats.RawCandidateCount +
                " CompleteCompositeCandidateCount=" + stats.CompleteCompositeCandidateCount +
                " BlankParentAppCount=" + stats.BlankParentAppCount +
                " BlankSettingDefinitionNameCount=" + stats.BlankSettingDefinitionNameCount +
                " DuplicateCompositeCandidateCount=" + stats.DuplicateCompositeCandidateCount +
                " DistinctCompositeCandidateCount=" + stats.DistinctCompositeCandidateCount);
            if (report.Candidates.Count == 0) text.AppendLine("No candidate AppSetting rows identified.");
            foreach (var item in report.Candidates)
            {
                text.AppendLine("Type=" + item.ComponentType.ToString(CultureInfo.InvariantCulture) +
                    " Label=" + item.FormattedLabel + " State=" + item.State +
                    " SolutionComponentId=" + item.SolutionComponentId.ToString("D") +
                    " ObjectId=" + (item.ObjectId.HasValue ? item.ObjectId.Value.ToString("D") : "(null)") +
                    " RootComponentBehavior=" + (item.RootComponentBehavior.HasValue
                        ? item.RootComponentBehavior.Value.ToString(CultureInfo.InvariantCulture) : "(null)") +
                    " RootSolutionComponentId=" + (item.RootSolutionComponentId.HasValue
                        ? item.RootSolutionComponentId.Value.ToString("D") : "(null)") +
                    " IsMetadata=" + (item.IsMetadata.HasValue ? item.IsMetadata.Value.ToString() : "(null)"));
                text.AppendLine("  SettingDefinitionId=" + (item.SettingDefinitionId.Length == 0 ? "(none)" : item.SettingDefinitionId) +
                    " Name=" + (item.SettingDefinitionName.Length == 0 ? "(blank)" : item.SettingDefinitionName));
                text.AppendLine("  ParentAppModuleId=" + (item.ParentAppModuleId.Length == 0 ? "(none)" : item.ParentAppModuleId) +
                    " UniqueName=" + (item.ParentAppModuleUniqueName.Length == 0 ? "(blank)" : item.ParentAppModuleUniqueName) +
                    " Name=" + (item.ParentAppModuleName.Length == 0 ? "(blank)" : item.ParentAppModuleName));
                text.AppendLine("  CandidateCompositeIdentity=" + (item.CandidateCompositeIdentity.Length == 0
                    ? "(incomplete)" : item.CandidateCompositeIdentity) + "  Diagnostic=" + item.Diagnostic);
                foreach (var correlation in item.Correlations)
                {
                    if (correlation.MatchedRecordCount == 0 &&
                        ((correlation.EntityLogicalName == "settingdefinition" && item.SettingDefinitionName.Length != 0) ||
                        (correlation.EntityLogicalName == "appmodule" && item.ParentAppModuleUniqueName.Length != 0)))
                    {
                        text.AppendLine("  " + correlation.EntityLogicalName + " => Not applicable for direct objectid correlation. " +
                            "Resolved through appsetting." + (correlation.EntityLogicalName == "settingdefinition"
                                ? "settingdefinitionid." : "parentappmoduleid."));
                        continue;
                    }
                    text.AppendLine("  " + correlation.EntityLogicalName + " => " + correlation.State +
                        " rows=" + correlation.MatchedRecordCount + " (" + correlation.Diagnostic + ")");
                    foreach (var record in correlation.Records)
                    {
                        text.AppendLine("    matchedRecordId=" + record.RecordId.ToString("D"));
                        foreach (var field in record.Fields)
                            text.AppendLine("      " + field.Key + "=" + field.Value);
                    }
                }
            }
            text.AppendLine();
        }

        private static string FormatMetadata(AppSettingEvidenceReport source, AppSettingEvidenceReport target)
        {
            var text = new StringBuilder();
            foreach (var report in new[] { source, target })
            {
                text.AppendLine("Environment: " + report.EnvironmentName);
                foreach (var item in report.EntityMetadata)
                {
                    text.AppendLine(item.LogicalName + " => " + item.State +
                        ", primaryId=" + item.PrimaryIdAttribute + ", primaryName=" + item.PrimaryNameAttribute);
                    text.AppendLine("  Attributes: " + string.Join(", ", item.Attributes));
                    text.AppendLine("  Lookups: " + string.Join(", ", item.LookupAttributes));
                    text.AppendLine("  Relationships: " + string.Join(", ", item.Relationships));
                    text.AppendLine("  Diagnostic: " + item.Diagnostic);
                }
                text.AppendLine();
            }
            return text.ToString();
        }

        private static string FormatSettingDefinitions(AppSettingEvidenceReport source,
            AppSettingEvidenceReport target)
        {
            var text = new StringBuilder();
            foreach (var report in new[] { source, target })
            {
                text.AppendLine("Environment: " + report.EnvironmentName);
                foreach (var item in report.Candidates.Where(item => item.SettingDefinitionId.Length != 0))
                    text.AppendLine("SettingDefinitionId=" + item.SettingDefinitionId +
                        " Name=" + (item.SettingDefinitionName.Length == 0 ? "(blank)" : item.SettingDefinitionName) +
                        " State=" + item.State + " Diagnostic=" + item.Diagnostic);
                if (!report.Candidates.Any(item => item.SettingDefinitionId.Length != 0))
                    text.AppendLine("No Setting Definition correlations.");
                text.AppendLine();
            }
            return text.ToString();
        }

        private static string FormatParentApps(AppSettingEvidenceReport source, AppSettingEvidenceReport target)
        {
            var text = new StringBuilder();
            foreach (var report in new[] { source, target })
            {
                text.AppendLine("Environment: " + report.EnvironmentName);
                foreach (var item in report.Candidates.Where(item => item.ParentAppModuleId.Length != 0))
                    text.AppendLine("ParentAppModuleId=" + item.ParentAppModuleId +
                        " UniqueName=" + (item.ParentAppModuleUniqueName.Length == 0
                            ? "(blank)" : item.ParentAppModuleUniqueName) +
                        " Name=" + (item.ParentAppModuleName.Length == 0 ? "(blank)" : item.ParentAppModuleName));
                if (!report.Candidates.Any(item => item.ParentAppModuleId.Length != 0))
                    text.AppendLine("No parent AppModule correlations.");
                text.AppendLine();
            }
            return text.ToString();
        }

        private static string FormatComparison(AppSettingEvidenceComparison value)
        {
            var text = new StringBuilder();
            text.AppendLine("Candidate identities are diagnostic only and compared case-insensitively.");
            text.AppendLine("DEV candidate count: " + value.Source.Candidates.Count +
                "  UAT candidate count: " + value.Target.Candidates.Count);
            if (value.Candidates.Count == 0) text.AppendLine("No complete composite candidates.");
            foreach (var item in value.Candidates)
                text.AppendLine(item.Outcome + " | " + item.Identity + " | DEV=" + item.SourceCount +
                    " | UAT=" + item.TargetCount);
            return text.ToString();
        }

        private static string FormatRequests(AppSettingEvidenceReport source, AppSettingEvidenceReport target)
        {
            var text = new StringBuilder();
            AppendRequests(text, source);
            AppendRequests(text, target);
            text.AppendLine("Dataverse writes: 0 (this diagnostic operation is read-only).");
            return text.ToString();
        }

        private static void AppendRequests(StringBuilder text, AppSettingEvidenceReport report)
        {
            var item = report.Requests;
            text.AppendLine(string.Format(CultureInfo.InvariantCulture,
                "{0}: total={1}, WhoAmI={2}, solutioncomponent={3}, metadata={4}, candidate backing={5}, settingdefinition={6}, appmodule={7}, other={8}",
                report.EnvironmentName, item.Total, item.WhoAmI, item.SolutionComponent,
                item.MetadataDiscovery, item.CandidateBackingEntity, item.SettingDefinition,
                item.AppModule, item.Other));
        }

        private static string FormatDiagnostics(AppSettingEvidenceReport source, AppSettingEvidenceReport target)
        {
            var text = new StringBuilder();
            foreach (var report in new[] { source, target })
            {
                text.AppendLine(report.EnvironmentName + ":");
                if (report.Diagnostics.Count == 0) text.AppendLine("  None");
                foreach (var diagnostic in report.Diagnostics) text.AppendLine("  " + diagnostic);
            }
            return text.ToString();
        }

        private sealed class UTF8EncodingWithBom : System.Text.UTF8Encoding
        {
            internal UTF8EncodingWithBom() : base(true) { }
        }
    }
}
