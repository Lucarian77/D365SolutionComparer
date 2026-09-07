using System;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Windows.Forms;
using D365SolutionComparer.Models.Membership;
using D365SolutionComparer.Services.Membership;

namespace D365SolutionComparer
{
    internal sealed class MembershipResultsForm : Form
    {
        private readonly MembershipComparisonPresentation presentation;
        private readonly DataGridView resultsGrid;

        public MembershipResultsForm(MembershipComparisonPresentation presentation)
        {
            this.presentation = presentation ?? throw new ArgumentNullException(nameof(presentation));
            Text = "Solution Membership Compare - " + presentation.SolutionUniqueName;
            StartPosition = FormStartPosition.CenterParent;
            MinimumSize = new Size(900, 520);
            Size = new Size(1400, 760);
            Font = new Font("Segoe UI", 9F);

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
                Height = 24,
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
                Dock = DockStyle.Bottom,
                Height = 24,
                Text = "Coverage Details...",
                TextAlign = ContentAlignment.MiddleRight,
                LinkBehavior = LinkBehavior.HoverUnderline
            };
            coverageDetails.LinkClicked += CoverageDetails_LinkClicked;
            header.Controls.Add(diagnostics);
            header.Controls.Add(coverageDetails);
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
                AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.DisplayedCellsExceptHeaders,
                ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize,
                BackgroundColor = Color.White,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                MultiSelect = false,
                RowHeadersVisible = false
            };
            AddColumn("ComponentKind", "Component Kind", 130);
            AddColumn("PortableKey", "Component Identity / Portable Key", 210);
            AddColumn("SourcePresence", "Source Presence", 95);
            AddColumn("TargetPresence", "Target Presence", 95);
            AddColumn("MembershipStatus", "Membership Status", 145);
            AddColumn("SourceResolutionStatus", "Source Resolution Status", 120);
            AddColumn("TargetResolutionStatus", "Target Resolution Status", 120);
            AddColumn("Diagnostic", "Diagnostic / Reason", 260);
            AddColumn("SourceRawComponentType", "Source Raw Component Type", 90);
            AddColumn("TargetRawComponentType", "Target Raw Component Type", 90);
            resultsGrid.CellFormatting += ResultsGrid_CellFormatting;
            resultsGrid.DataSource = presentation.Rows.ToList();

            Controls.Add(resultsGrid);
            Controls.Add(header);
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
                presentation.Target.Diagnostics.EnvironmentName, target))
                form.ShowDialog(this);
        }

        private void AddColumn(string propertyName, string title, float fillWeight)
        {
            resultsGrid.Columns.Add(new DataGridViewTextBoxColumn
            {
                DataPropertyName = propertyName,
                HeaderText = title,
                Name = propertyName,
                AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill,
                FillWeight = fillWeight,
                SortMode = DataGridViewColumnSortMode.Automatic,
                DefaultCellStyle = propertyName == "Diagnostic"
                    ? new DataGridViewCellStyle { WrapMode = DataGridViewTriState.True }
                    : new DataGridViewCellStyle()
            });
        }

        private void ResultsGrid_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            if (e.RowIndex < 0) return;
            var row = resultsGrid.Rows[e.RowIndex].DataBoundItem as MembershipResultRow;
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

        private static string FormatSummary(MembershipPresentationSummary summary) =>
            "Present in Both: " + summary.PresentInBoth + "  |  Source Only: " + summary.SourceOnly +
            "  |  Target Only: " + summary.TargetOnly + "  |  Unsupported: " + summary.Unsupported +
            "  |  Unresolved: " + summary.Unresolved + "  |  Ambiguous: " + summary.Ambiguous;

        private static string FormatDiagnostics(string side, MembershipEnvironmentDiagnostics diagnostics) =>
            side + " (" + diagnostics.EnvironmentName + ") diagnostics: requests=" + diagnostics.RequestCount +
            ", elapsed=" + diagnostics.Elapsed.TotalSeconds.ToString("0.00", CultureInfo.InvariantCulture) + "s" +
            ", raw=" + Number(diagnostics.RawMembershipCount) +
            ", resolved=" + Number(diagnostics.ResolvedCount) +
            ", unsupported=" + Number(diagnostics.UnsupportedCount) +
            ", unresolved=" + Number(diagnostics.UnresolvedCount) +
            ", ambiguous=" + Number(diagnostics.AmbiguousCount);

        private static string Number(int? value) => value.HasValue
            ? value.Value.ToString(CultureInfo.InvariantCulture) : "n/a";
    }
}
