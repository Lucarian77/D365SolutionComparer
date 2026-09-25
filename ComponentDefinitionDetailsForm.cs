using System;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using D365SolutionComparer.Models.ComponentDetails;

namespace D365SolutionComparer
{
    internal sealed class ComponentDefinitionDetailsForm : Form
    {
        public ComponentDefinitionDetailsForm(ComponentDefinitionDetailPresentation detail)
        {
            if (detail == null) throw new ArgumentNullException(nameof(detail));
            Text = "Component Definition Details - " + detail.ComponentKind;
            StartPosition = FormStartPosition.CenterParent;
            MinimumSize = new Size(760, 460);
            Size = new Size(1050, 680);
            Font = new Font("Segoe UI", 9F);

            var header = new Label
            {
                Dock = DockStyle.Top,
                Height = 76,
                Padding = new Padding(10),
                BackColor = Color.White,
                AutoEllipsis = true,
                Text = detail.ComponentKind + ": " + detail.PortableKey + Environment.NewLine +
                    "Membership: " + detail.MembershipStatus + "    |    Definition: " +
                    detail.DefinitionStatus + Environment.NewLine + detail.Diagnostic
            };

            var grid = new DataGridView
            {
                Dock = DockStyle.Fill,
                ReadOnly = true,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                AutoGenerateColumns = false,
                BackgroundColor = Color.White,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                MultiSelect = false,
                RowHeadersVisible = false,
                DataSource = detail.Properties.ToList()
            };
            AddColumn(grid, "PropertyName", "Comparable Property", 170);
            AddColumn(grid, "SourceValue", "Source Value", 330);
            AddColumn(grid, "TargetValue", "Target Value", 330);
            AddColumn(grid, "Comparison", "Result", 90);
            grid.CellFormatting += (sender, args) =>
            {
                if (args.RowIndex < 0) return;
                var item = grid.Rows[args.RowIndex].DataBoundItem as ComponentPropertyPresentation;
                if (item == null || !item.Changed) return;
                args.CellStyle.BackColor = Color.MistyRose;
                args.CellStyle.ForeColor = Color.Firebrick;
            };

            var evidence = new TextBox
            {
                Dock = DockStyle.Fill,
                Multiline = true,
                ReadOnly = true,
                ScrollBars = ScrollBars.Vertical,
                BackColor = Color.White,
                Text = "Source diagnostic evidence:" + Environment.NewLine +
                    Evidence(detail.SourceEvidence) + Environment.NewLine + Environment.NewLine +
                    "Target diagnostic evidence:" + Environment.NewLine + Evidence(detail.TargetEvidence)
            };

            var sections = new SplitContainer
            {
                Dock = DockStyle.Fill,
                Orientation = Orientation.Horizontal,
                FixedPanel = FixedPanel.None,
                Panel1MinSize = 150,
                Panel2MinSize = 130,
                SplitterWidth = 7,
                BackColor = Color.Gainsboro
            };
            sections.Panel1.Controls.Add(grid);
            sections.Panel2.Controls.Add(evidence);

            Controls.Add(sections);
            Controls.Add(header);

            Shown += (sender, args) =>
            {
                var availableHeight = sections.ClientSize.Height - sections.SplitterWidth;
                if (availableHeight >= sections.Panel1MinSize + sections.Panel2MinSize)
                {
                    sections.SplitterDistance = Math.Max(sections.Panel1MinSize,
                        Math.Min(availableHeight - sections.Panel2MinSize,
                            (int)Math.Round(availableHeight * 0.6)));
                }
            };
        }

        private static string Evidence(System.Collections.Generic.IEnumerable<string> values)
        {
            var items = values.Where(item => !string.IsNullOrWhiteSpace(item)).ToList();
            return items.Count == 0 ? "(none)" : string.Join(Environment.NewLine, items);
        }

        private static void AddColumn(DataGridView grid, string propertyName, string title,
            float fillWeight)
        {
            grid.Columns.Add(new DataGridViewTextBoxColumn
            {
                DataPropertyName = propertyName,
                HeaderText = title,
                Name = propertyName,
                AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill,
                FillWeight = fillWeight,
                SortMode = DataGridViewColumnSortMode.Automatic,
                DefaultCellStyle = new DataGridViewCellStyle
                {
                    NullValue = "(not available)",
                    WrapMode = DataGridViewTriState.True
                }
            });
        }
    }
}
