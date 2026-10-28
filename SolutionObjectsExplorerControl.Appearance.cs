using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using D365SolutionComparer.Models.ComponentDetails;
using D365SolutionComparer.Models.Presentation;

namespace D365SolutionComparer
{
    internal sealed partial class SolutionObjectsExplorerControl
    {
        private bool professionalAppearance, arrangingProfessionalDetails;

        internal static Color[] ProfessionalTreeColors(bool selected) => selected ? new[] { Color.FromArgb(29, 78, 216), Color.White } :
            new[] { Color.White, Color.FromArgb(51, 65, 85) };

        internal int[] ProfessionalHeaderCounts => Projection == null || !makerNavigation || ObjectInventoryUnavailable ? null :
            ProfessionalExplorerCounts(scopeObjects.Select(item => item.Status(Projection)));

        internal static int[] ProfessionalExplorerCounts(IEnumerable<string> statuses)
        {
            var counts = new int[5];
            foreach (string status in statuses ?? Enumerable.Empty<string>())
            {
                counts[0]++;
                if (status == "Match") counts[1]++;
                else if (status == "Different") counts[2]++;
                else if (status == "Needs Review") counts[3]++;
                else if (status == "Target Only") counts[4]++;
            }
            return counts;
        }

        internal static void ProfessionalExplorerCellStyle(DataGridViewCellStyle style, string status, bool statusCell, bool alternate)
        {
            SolutionComparerControl.ApplyComparisonCellStyle(style, "ExplorerCell", null, alternate);
            if (!statusCell) return;
            switch (status)
            {
                case "Match": style.ForeColor = Color.FromArgb(22, 101, 52); style.BackColor = Color.FromArgb(240, 253, 244); break;
                case "Different": style.ForeColor = Color.FromArgb(146, 64, 14); style.BackColor = Color.FromArgb(255, 251, 235); break;
                case "Needs Review": case "Unsupported": case "Indeterminate":
                    style.ForeColor = Color.FromArgb(146, 64, 14); style.BackColor = Color.FromArgb(255, 251, 235); break;
                case "Source Only": case "Target Only":
                    style.ForeColor = Color.FromArgb(185, 28, 28); style.BackColor = Color.FromArgb(254, 242, 242); break;
                default: style.ForeColor = Color.FromArgb(71, 85, 105); style.BackColor = Color.FromArgb(241, 245, 249); break;
            }
        }

        internal static void ProfessionalPropertyCellStyle(DataGridViewCellStyle style, ComponentPropertyPresentation property, bool resultCell, bool alternate)
        {
            string comparison = SolutionObjectsExplorerRow.PropertyComparison(property);
            ProfessionalExplorerCellStyle(style, comparison, resultCell, alternate);
            if (!resultCell && property.Changed && property.Comparison == "Different") style.BackColor = Color.FromArgb(255, 251, 235);
        }

        // Identity and explanation remain scrollable/tooltip-accessible. Never consume a complete
        // property line when the minimum readable identity/explanation allocations can coexist.
        internal static int[] ProfessionalDetailsAllocation(int available, int requiredProperty, int identityPreferred,
            int explanationPreferred, int identityMinimum, int explanationMinimum)
        {
            int budget = Math.Max(0, available - requiredProperty);
            if (budget < identityMinimum + explanationMinimum) return new[] { identityMinimum, explanationMinimum };
            int explanation = Math.Min(explanationPreferred, Math.Max(explanationMinimum, budget - identityMinimum));
            int identity = Math.Min(identityPreferred, Math.Max(identityMinimum, budget - explanation));
            return new[] { identity, explanation };
        }

        private Font ProfessionalExplorerFont(float size, FontStyle style = FontStyle.Regular)
        {
            var font = new Font("Segoe UI", size, style, GraphicsUnit.Point);
            Disposed += (sender, args) => font.Dispose(); return font;
        }
        private int ProfessionalPixels(int logical) => (int)Math.Ceiling(logical * Math.Max(96, DeviceDpi) / 96d);

        private void ConfigureProfessionalExplorerAppearance()
        {
            if (professionalAppearance) return; professionalAppearance = true;
            BackColor = Color.White;
            navigation.Font = ProfessionalExplorerFont(9F);
            navigation.BackColor = Color.White; navigation.ForeColor = Color.FromArgb(51, 65, 85);
            navigation.BorderStyle = BorderStyle.None;
            navigation.ItemHeight = ProfessionalPixels(26); navigation.Indent = ProfessionalPixels(22);
            navigation.DrawMode = TreeViewDrawMode.OwnerDrawText;
            navigation.DrawNode += (sender, args) =>
            {
                if ((args.State & TreeNodeStates.Selected) == 0) { args.DrawDefault = true; return; }
                var colors = ProfessionalTreeColors(true);
                using (var background = new SolidBrush(colors[0])) args.Graphics.FillRectangle(background, args.Bounds);
                TextRenderer.DrawText(args.Graphics, args.Node.Text, args.Node.NodeFont ?? navigation.Font, args.Bounds, colors[1],
                    TextFormatFlags.NoPrefix | TextFormatFlags.VerticalCenter | TextFormatFlags.PreserveGraphicsClipping);
                if ((args.State & TreeNodeStates.Focused) != 0) ControlPaint.DrawFocusRectangle(args.Graphics, args.Bounds, Color.White, Color.FromArgb(29, 78, 216));
            };
            foreach (var view in new[] { grid, comparedProperties })
            {
                view.Font = ProfessionalExplorerFont(9F);
                view.EnableHeadersVisualStyles = false; view.BackgroundColor = Color.White;
                view.GridColor = Color.FromArgb(226, 232, 240); view.BorderStyle = BorderStyle.FixedSingle;
                view.CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal;
                view.ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.Single;
                view.DefaultCellStyle = SolutionComparerControl.CreateComparisonRowStyle(false);
                view.RowsDefaultCellStyle = SolutionComparerControl.CreateComparisonRowStyle(false);
                view.AlternatingRowsDefaultCellStyle = SolutionComparerControl.CreateComparisonRowStyle(true);
                view.ColumnHeadersDefaultCellStyle = SolutionComparerControl.CreateComparisonHeaderStyle();
                view.ColumnHeadersDefaultCellStyle.Font = ProfessionalExplorerFont(9F, FontStyle.Bold);
                view.ColumnHeadersDefaultCellStyle.WrapMode = DataGridViewTriState.False;
                view.ShowCellToolTips = true;
            }
            grid.CellFormatting += (sender, args) =>
            {
                if (!makerNavigation || args.RowIndex < 0 || args.RowIndex >= visibleObjects.Count || args.ColumnIndex < 0) return;
                ProfessionalExplorerCellStyle(args.CellStyle, visibleObjects[args.RowIndex].Status(Projection),
                    grid.Columns[args.ColumnIndex].Name == "ComparisonStatus", (args.RowIndex & 1) != 0);
            };
            comparedProperties.CellFormatting += (sender, args) =>
            {
                if (args.RowIndex < 0 || args.RowIndex >= comparedPropertyRows.Count || args.ColumnIndex < 0) return;
                ProfessionalPropertyCellStyle(args.CellStyle, comparedPropertyRows[args.RowIndex],
                    comparedProperties.Columns[args.ColumnIndex].Name == "Result", (args.RowIndex & 1) != 0);
            };
            counts.Font = ProfessionalExplorerFont(9F); counts.Padding = new Padding(8, 4, 8, 4);
            counts.ForeColor = Color.FromArgb(71, 85, 105); counts.BackColor = Color.FromArgb(248, 250, 252);
            simplifiedIdentity.Font = ProfessionalExplorerFont(9F); simplifiedIdentity.BorderStyle = BorderStyle.None;
            simplifiedIdentity.BackColor = Color.FromArgb(248, 250, 252); simplifiedIdentity.ForeColor = Color.FromArgb(15, 23, 42);
            comparedPropertiesMessage.Font = ProfessionalExplorerFont(9F);
            comparedPropertiesMessage.BackColor = Color.FromArgb(241, 245, 249);
            comparedPropertiesMessage.ForeColor = Color.FromArgb(51, 65, 85); comparedPropertiesMessage.Padding = new Padding(8, 4, 8, 4);
            detailsLinks.BackColor = Color.FromArgb(248, 250, 252);
            foreach (var link in new[] { openRow, definition, expandDetails })
            { link.Font = ProfessionalExplorerFont(9F); link.LinkColor = Color.FromArgb(29, 78, 216); link.ActiveLinkColor = Color.FromArgb(30, 58, 138); }
            simplifiedIdentity.TextChanged += (sender, args) =>
            { countToolTip.SetToolTip(simplifiedIdentity, simplifiedIdentity.Text); ArrangeProfessionalDetails(); };
            comparedPropertiesMessage.TextChanged += (sender, args) => ArrangeProfessionalDetails();
            comparedPropertiesPanel.SizeChanged += (sender, args) => ArrangeProfessionalDetails();
            evidenceSplit.SplitterMoved += (sender, args) => ArrangeProfessionalDetails();
            VisibleChanged += (sender, args) => ArrangeProfessionalDetails();
            UpdateGridMetrics();
        }

        private void ArrangeProfessionalDetails()
        {
            if (!makerNavigation || arrangingProfessionalDetails || !Visible || evidenceSplit.Panel2Collapsed) return;
            arrangingProfessionalDetails = true;
            try
            {
                int font = simplifiedIdentity.Font.Height;
                int required = comparedProperties.ColumnHeadersHeight + comparedProperties.RowTemplate.Height + SystemInformation.BorderSize.Height * 2;
                int explanation = Math.Max(font + 8, TextRenderer.MeasureText(comparedPropertiesMessage.Text, comparedPropertiesMessage.Font,
                    new Size(Math.Max(1, comparedPropertiesPanel.ClientSize.Width - 16), int.MaxValue), TextFormatFlags.WordBreak).Height + 8);
                var allocation = ProfessionalDetailsAllocation(comparedPropertiesPanel.ClientSize.Height, required,
                    font * 4 + SystemInformation.HorizontalScrollBarHeight + 4, explanation,
                    font + SystemInformation.HorizontalScrollBarHeight + 4, font + 8);
                simplifiedIdentity.Height = allocation[0]; comparedPropertiesMessage.Height = allocation[1];
                UpdateObjectViewport();
            }
            finally { arrangingProfessionalDetails = false; }
        }
    }
}
