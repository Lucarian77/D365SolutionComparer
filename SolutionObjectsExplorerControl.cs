using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using D365SolutionComparer.Models.Membership;
using D365SolutionComparer.Models.Presentation;

namespace D365SolutionComparer
{
    /// <summary>Local read-only object navigation. No service, cache or resolver dependency.</summary>
    internal sealed class SolutionObjectsExplorerControl : UserControl
    {
        private const int PageSize = 200;
        private readonly TreeView navigation;
        private readonly DataGridView grid;
        private readonly TextBox details;
        private readonly Label counts;
        private readonly LinkLabel openRow, definition;
        private readonly Func<MembershipGridRow, string> describe;
        private readonly Action<MembershipGridRow> openOriginal, openDefinition;
        private readonly Dictionary<int, List<int>> parents = new Dictionary<int, List<int>>();
        private HashSet<int> allowedRows, includedOccurrences, matchingOccurrences;
        private IReadOnlyList<int> visibleRows = new int[0];
        private bool preparing;
        private TreeNode inspectionNode;
        private MembershipGridRow selected;
        internal SolutionObjectsExplorerProjection Projection { get; private set; }
        internal TreeView Navigation => navigation;
        internal DataGridView Grid => grid;
        internal string EvidenceText => details.Text;
        internal IReadOnlyList<int> VisibleRows => visibleRows;
        internal MembershipGridRow SelectedComparison => selected;
        internal DisplayEnvironmentSide? SelectedSide { get; private set; }
        internal Action<MembershipGridRow> SelectedComparisonChanged { get; set; }
        internal Action<Exception> RenderingFailed { get; set; }

        internal SolutionObjectsExplorerControl(Func<MembershipGridRow, string> describe,
            Action<MembershipGridRow> openOriginal, Action<MembershipGridRow> openDefinition)
        {
            this.describe = describe; this.openOriginal = openOriginal; this.openDefinition = openDefinition;
            Dock = DockStyle.Fill; Name = "SolutionObjectsExplorer"; AccessibleName = "Solution Objects Explorer";
            var outer = new SplitContainer { Dock = DockStyle.Fill, Size = new Size(1000, 550), SplitterDistance = 255,
                Panel1MinSize = 180, Panel2MinSize = 450, FixedPanel = FixedPanel.Panel1, Name = "ExplorerNavigationSplit" };
            navigation = new TreeView { Dock = DockStyle.Fill, HideSelection = false, ShowNodeToolTips = true,
                Name = "ExplorerObjectsNavigation", AccessibleName = "Objects, Source and Target category navigation" };
            var right = new SplitContainer { Dock = DockStyle.Fill, Size = new Size(735, 550), Orientation = Orientation.Horizontal,
                SplitterDistance = 325, Panel1MinSize = 170, Panel2MinSize = 100, Name = "ExplorerEvidenceSplit" };
            grid = new DataGridView { Dock = DockStyle.Fill, Name = "ExplorerObjectGrid", AccessibleName = "Object membership comparison",
                ReadOnly = true, VirtualMode = true, AutoGenerateColumns = false, AllowUserToAddRows = false, AllowUserToDeleteRows = false,
                AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None, MultiSelect = false, SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                RowHeadersVisible = false, BackgroundColor = Color.White, RowTemplate = new ExplorerAccessibleRow { Tag = this } };
            grid.RowTemplate.Height = Math.Max(46, Font.Height * 3 + 4);
            foreach (var column in new[] { new[] { "DisplayName", "Display Name", "180" }, new[] { "Name", "Name", "210" },
                new[] { "Type", "Type", "155" }, new[] { "Source", "Source", "105" }, new[] { "Target", "Target", "105" },
                new[] { "ComparisonStatus", "Comparison Status", "250" }, new[] { "Managed", "Managed", "110" },
                new[] { "Customized", "Customized", "110" }, new[] { "LastModified", "Last Modified", "185" },
                new[] { "Owner", "Owner", "185" }, new[] { "Status", "Status", "170" } })
                grid.Columns.Add(new DataGridViewTextBoxColumn { Name = column[0], HeaderText = column[1],
                    Width = MembershipResultsForm.ScaleColumnWidth(int.Parse(column[2]), DeviceDpi),
                    MinimumWidth = 65, SortMode = DataGridViewColumnSortMode.NotSortable, Visible = grid.Columns.Count < 6,
                    DefaultCellStyle = new DataGridViewCellStyle { WrapMode = DataGridViewTriState.True } });
            grid.CellValueNeeded += (s, e) => { if (e.RowIndex < visibleRows.Count) e.Value = Value(Projection.Rows[visibleRows[e.RowIndex]], grid.Columns[e.ColumnIndex].Name); };
            grid.CellClick += (s, e) => Guard(() => { if (e.RowIndex >= 0) SelectGridRow(e.RowIndex); });
            grid.KeyUp += (s, e) => Guard(() => { if (grid.CurrentRow != null) SelectGridRow(grid.CurrentRow.Index); });
            grid.CellDoubleClick += (s, e) => Guard(() => { if (e.RowIndex >= 0) { SelectGridRow(e.RowIndex); InspectOccurrences(visibleRows[e.RowIndex]); } });
            grid.CellFormatting += (s, e) =>
            {
                if (e.RowIndex < 0 || e.RowIndex >= visibleRows.Count || grid.Columns[e.ColumnIndex].Name != "ComparisonStatus" || grid.Rows[e.RowIndex].Selected) return;
                var presence = Projection.Rows[visibleRows[e.RowIndex]].Original.AuthoritativeRow.Presence;
                e.CellStyle.BackColor = presence == MembershipPresence.PresentInBoth ? Color.FromArgb(243, 247, 244) :
                    presence == MembershipPresence.Indeterminate ? Color.FromArgb(255, 247, 232) : Color.FromArgb(251, 239, 239);
                e.CellStyle.ForeColor = Color.FromArgb(55, 60, 66);
            };
            grid.CellToolTipTextNeeded += (s, e) =>
            {
                if (e.RowIndex >= 0 && e.RowIndex < visibleRows.Count)
                    e.ToolTipText = "Original comparison row " + visibleRows[e.RowIndex] + ". " +
                        MembershipGridRow.DefinitionMeaning + " Missing means selected solution membership, not missing environment metadata.";
            };
            counts = new Label { Dock = DockStyle.Top, Height = 48, Name = "ExplorerAccounting", AccessibleName = "Explorer count units", AutoEllipsis = true };
            details = new TextBox { Dock = DockStyle.Fill, Multiline = true, ReadOnly = true, ScrollBars = ScrollBars.Both,
                WordWrap = false, BackColor = Color.White, Name = "ExplorerEvidence", AccessibleName = "Selected Source and Target object evidence" };
            openRow = new LinkLabel { Dock = DockStyle.Left, Width = 235, Text = "Open original row in Component Details", Enabled = false };
            definition = new LinkLabel { Dock = DockStyle.Fill, Text = "Open existing definition details", Enabled = false };
            openRow.LinkClicked += (s, e) => { if (selected != null) openOriginal(selected); };
            definition.LinkClicked += (s, e) => { if (selected?.AuthoritativeRow.DefinitionDetail != null) openDefinition(selected); };
            var links = new Panel { Dock = DockStyle.Bottom, Height = 28 }; links.Controls.Add(definition); links.Controls.Add(openRow);
            right.Panel1.Controls.Add(grid); right.Panel1.Controls.Add(counts); right.Panel2.Controls.Add(details); right.Panel2.Controls.Add(links);
            outer.Panel1.Controls.Add(navigation); outer.Panel2.Controls.Add(right); Controls.Add(outer);
            navigation.BeforeExpand += (s, e) => Guard(() => Populate(e.Node));
            navigation.AfterSelect += (s, e) => Guard(() => SelectNavigation(e.Node));
        }

        internal void Bind(SolutionObjectsExplorerProjection projection, IEnumerable<int> filteredRows, bool all)
        {
            Projection = projection ?? throw new ArgumentNullException(nameof(projection)); parents.Clear();
            foreach (var parent in projection.Occurrences)
                foreach (var link in parent.Children)
                {
                    if (!parents.TryGetValue(link.ChildOccurrenceIndex, out var list)) parents[link.ChildOccurrenceIndex] = list = new List<int>();
                    list.Add(parent.Original.OccurrenceIndex);
                }
            ApplyFilter(filteredRows, all);
        }
        internal void ApplyFilter(IEnumerable<int> filteredRows, bool all)
        {
            allowedRows = new HashSet<int>(filteredRows);
            if (allowedRows.Any(i => i < 0 || i >= Projection.Rows.Count)) throw new ArgumentOutOfRangeException(nameof(filteredRows));
            matchingOccurrences = new HashSet<int>(Projection.Occurrences.Where(o => all || o.Original.ComparisonRowIndices.Any(allowedRows.Contains))
                .Select(o => o.Original.OccurrenceIndex));
            includedOccurrences = new HashSet<int>(matchingOccurrences);
            var queue = new Queue<int>(includedOccurrences);
            while (queue.Count > 0)
                if (parents.TryGetValue(queue.Dequeue(), out var parentIds))
                    foreach (int id in parentIds) if (includedOccurrences.Add(id)) queue.Enqueue(id);
            var prior = selected?.OriginalRowIndex; var side = SelectedSide;
            navigation.BeginUpdate(); preparing = true;
            try
            {
                navigation.Nodes.Clear();
                inspectionNode = null;
                foreach (var key in new[] { "objects", "all", "review", "unmapped" })
                {
                    var category = Projection.Category(key);
                    navigation.Nodes.Add(new TreeNode(category.Name + " (" + category.RowIndices.Count(allowedRows.Contains) + " rows)")
                        { Tag = new Selection { Rows = category.RowIndices, Description = category.CountText + "; categories may overlap. Not a Power Apps top-level count." } });
                }
                foreach (var role in new[] { DisplayEnvironmentSide.Source, DisplayEnvironmentSide.Target })
                {
                    var sideNode = new TreeNode(role + ": " + (role == DisplayEnvironmentSide.Source ? Projection.Owner.Source.Diagnostics.EnvironmentName : Projection.Owner.Target.Diagnostics.EnvironmentName));
                    foreach (var category in Projection.Categories.Where(c => c.Key != "objects" && c.Key != "all" && c.Key != "review" && c.Key != "unmapped"))
                    {
                        var ids = (role == DisplayEnvironmentSide.Source ? category.SourceOccurrences : category.TargetOccurrences).Where(includedOccurrences.Contains).ToList();
                        if (ids.Count == 0) continue;
                        var entries = ids.Select(i => new Selection { Occurrence = i, Side = role }).ToList();
                        int contextCount = ids.Count(i => !matchingOccurrences.Contains(i));
                        var node = new TreeNode(category.Name + " (" + ids.Count + " raw" + (contextCount == 0 ? "" : "; " + contextCount + " context") + ")")
                        { Tag = new Selection { Rows = ids.SelectMany(i => Projection.Occurrences[i].Original.ComparisonRowIndices).Distinct().ToArray(), Side = role,
                            Entries = entries.AsReadOnly(), Description = "Side-specific raw occurrences, not object or comparison totals." },
                            ToolTipText = category.CountText + "; visible " + ids.Count + " raw occurrences, including " + contextCount + " navigation context occurrences" };
                        node.Nodes.Add(new TreeNode("Expand to inspect...")); sideNode.Nodes.Add(node);
                    }
                    navigation.Nodes.Add(sideNode); sideNode.Expand();
                }
            }
            finally { preparing = false; navigation.EndUpdate(); }
            SetRows(Projection.Category("objects").RowIndices, "All Objects (candidates), not verified maker top-level objects");
            if (prior.HasValue && allowedRows.Contains(prior.Value)) SelectComparison(prior.Value, side);
        }
        internal void ShowMetadata(bool show) { foreach (DataGridViewColumn c in grid.Columns) if (c.Index >= 6) c.Visible = show; }
        internal void SelectComparison(int index, DisplayEnvironmentSide? side = null)
        {
            if (!allowedRows.Contains(index)) return;
            if (!visibleRows.Contains(index)) SetRows(Projection.Category("all").RowIndices, "All Components (original comparison rows)");
            SelectedSide = side; int position = visibleRows.ToList().IndexOf(index);
            grid.CurrentCell = grid.Rows[position].Cells[0]; SelectGridRow(position);
        }
        internal void InspectOccurrences(int rowIndex)
        {
            var row = Projection.Rows[rowIndex];
            var choices = row.SourceOccurrences.Concat(row.TargetOccurrences).Select(i => new Selection { Occurrence = i, Side = Projection.Occurrences[i].Original.Side }).ToList();
            var node = new TreeNode("Original row " + rowIndex + " — all " + choices.Count + " mapped raw occurrences")
                { Tag = new Selection { Rows = new[] { rowIndex }, Entries = choices.AsReadOnly(), Description = "Choose a side/occurrence explicitly; no GUID-based pairing." } };
            inspectionNode?.Remove(); inspectionNode = node;
            node.Nodes.Add(new TreeNode("Expand to inspect...")); navigation.Nodes.Add(node); node.Expand(); navigation.SelectedNode = node;
        }
        internal void SelectOccurrence(int index)
        {
            if (!includedOccurrences.Contains(index)) return;
            var item = Projection.Occurrences[index]; var node = MakeNode(new Selection { Occurrence = index, Side = item.Original.Side });
            inspectionNode?.Remove(); inspectionNode = node;
            navigation.Nodes.Add(node); navigation.SelectedNode = node; node.Expand();
        }
        private sealed class Selection
        {
            internal int? Occurrence; internal DisplayEnvironmentSide? Side; internal IReadOnlyList<int> Rows;
            internal IReadOnlyList<Selection> Entries; internal string Description;
        }
        private sealed class Page { internal IReadOnlyList<Selection> Entries; internal int Start, Count; }
        private TreeNode MakeNode(Selection selection)
        {
            var item = Projection.Occurrences[selection.Occurrence.Value];
            var node = new TreeNode(item.Original.Side + ": " + item.Label + " [raw " + item.Original.OriginalRawIndex + "]" +
                (matchingOccurrences.Contains(item.Original.OccurrenceIndex) ? "" : " (navigation context)")) { Tag = selection,
                    ToolTipText = item.Original.Placement + "; raw occurrence, not top-level proof" };
            if (item.Children.Count > 0 || item.Original.Record.ComponentType == 1 || item.Original.Record.ComponentType == 80)
                node.Nodes.Add(new TreeNode("Expand to inspect...")); return node;
        }
        private void AddEntries(TreeNode node, IReadOnlyList<Selection> entries)
        {
            if (entries.Count <= PageSize) foreach (var entry in entries) node.Nodes.Add(MakeNode(entry));
            else for (int start = 0; start < entries.Count; start += PageSize)
            {
                var page = new TreeNode("Items " + (start + 1) + "–" + Math.Min(entries.Count, start + PageSize))
                    { Tag = new Page { Entries = entries, Start = start, Count = Math.Min(PageSize, entries.Count - start) } };
                page.Nodes.Add(new TreeNode("Expand to inspect...")); node.Nodes.Add(page);
            }
        }
        private void Populate(TreeNode node)
        {
            if (node.Nodes.Count != 1 || node.Nodes[0].Text != "Expand to inspect...") return;
            node.Nodes.Clear();
            if (node.Tag is Page page)
            { for (int i = page.Start; i < page.Start + page.Count; i++) node.Nodes.Add(MakeNode(page.Entries[i])); return; }
            var selection = node.Tag as Selection; if (selection == null) return;
            if (selection.Entries != null) { AddEntries(node, selection.Entries); return; }
            if (!selection.Occurrence.HasValue) return;
            var item = Projection.Occurrences[selection.Occurrence.Value];
            var presentGroups = new HashSet<string>();
            foreach (var group in item.Children.Where(l => includedOccurrences.Contains(l.ChildOccurrenceIndex))
                .GroupBy(l => new { l.Group, l.OriginalAssociation.Kind }).OrderBy(g => g.Key.Group, StringComparer.Ordinal).ThenBy(g => g.Key.Kind))
            {
                var entries = group.Select(l => new Selection { Occurrence = l.ChildOccurrenceIndex, Side = item.Original.Side }).ToList();
                presentGroups.Add(group.Key.Group);
                string evidence = group.Key.Kind == HierarchyRelationshipKind.StructuralParent ? "Verified structural parent" : "Inclusion-root only — not ownership";
                int rawCount = group.Select(l => l.ChildOccurrenceIndex).Distinct().Count();
                var child = new TreeNode(group.Key.Group + " — " + (group.Key.Kind == HierarchyRelationshipKind.StructuralParent ? "structural" : "inclusion root") +
                    " (" + rawCount + " raw" + (entries.Count == rawCount ? "" : "; " + entries.Count + " links") + ")")
                { Tag = new Selection { Entries = entries.AsReadOnly(), Side = item.Original.Side,
                    Rows = entries.SelectMany(e => Projection.Occurrences[e.Occurrence.Value].Original.ComparisonRowIndices).Distinct().ToArray(), Description = evidence },
                    ToolTipText = evidence + "; " + rawCount + " distinct raw occurrences; " + entries.Count + " association references" };
                child.Nodes.Add(new TreeNode("Expand to inspect...")); node.Nodes.Add(child);
            }
            var expected = item.Original.Record.ComponentType == 1 ? new[] { "Columns", "Relationships", "Forms", "Views", "Other Components" } :
                item.Original.Record.ComponentType == 80 ? new[] { "AppSettings", "Other Components" } : new string[0];
            foreach (var name in expected.Where(n => !presentGroups.Contains(n)))
                node.Nodes.Add(new TreeNode(name + " — no retained verified links") { Tag = new Selection { Rows = new int[0], Side = item.Original.Side,
                    Description = "No retained verified " + name + " links. This is not proof of no components in the environment." } });
        }
        private void SelectNavigation(TreeNode node)
        {
            if (preparing || !(node.Tag is Selection selection)) return;
            SelectedSide = selection.Side;
            if (selection.Occurrence.HasValue)
            {
                var item = Projection.Occurrences[selection.Occurrence.Value];
                SetRows(item.Original.ComparisonRowIndices, "Selected raw occurrence; all original row mappings");
                details.Text = OccurrenceEvidence(item);
                if (visibleRows.Count == 1) { SelectGridRow(0); details.Text = OccurrenceEvidence(item) + Environment.NewLine + details.Text; }
                else details.AppendText(Environment.NewLine + "Choose a comparison row explicitly; no arbitrary first mapping selected.");
            }
            else { SetRows(selection.Rows ?? new int[0], selection.Description); details.Text = selection.Description; }
        }
        private void SetRows(IEnumerable<int> indices, string context)
        {
            preparing = true;
            try
            {
                visibleRows = indices.Where(allowedRows.Contains).Distinct().OrderBy(i => i).ToList().AsReadOnly();
                grid.CurrentCell = null; grid.RowCount = visibleRows.Count; selected = null; openRow.Enabled = definition.Enabled = false;
                counts.Text = context + Environment.NewLine + visibleRows.Count + " visible comparison rows / " + Projection.Rows.Count +
                    " total; Source " + Projection.SourceOccurrenceCount + ", Target " + Projection.TargetOccurrenceCount + " raw occurrences";
            }
            finally { preparing = false; }
            SelectedComparisonChanged?.Invoke(null);
        }
        private void SelectGridRow(int position)
        {
            if (preparing || position < 0 || position >= visibleRows.Count) return;
            var row = Projection.Rows[visibleRows[position]]; selected = row.Original;
            details.Text = "Original comparison row: " + row.OriginalRowIndex + Environment.NewLine +
                "Source raw occurrence mappings: " + string.Join(",", row.SourceOccurrences) + Environment.NewLine +
                "Target raw occurrence mappings: " + string.Join(",", row.TargetOccurrences) + Environment.NewLine + describe(selected);
            openRow.Enabled = true; definition.Enabled = selected.AuthoritativeRow.DefinitionDetail != null;
            SelectedComparisonChanged?.Invoke(selected);
        }
        private static string OccurrenceEvidence(SolutionObjectsExplorerOccurrence item)
        {
            var o = item.Original; var r = o.Record; var text = new StringBuilder();
            text.AppendLine(o.Side + ": " + o.Snapshot.Environment.DisplayName).AppendLine("Solution: " + o.Snapshot.SolutionUniqueName);
            text.AppendLine("Original raw index: " + o.OriginalRawIndex).AppendLine("H1 occurrence index: " + o.OccurrenceIndex);
            text.AppendLine("SolutionComponentId: " + r.SolutionComponentId).AppendLine("ObjectId (local audit): " + r.ObjectId);
            text.AppendLine("ComponentType: " + r.ComponentType).AppendLine("RootSolutionComponentId: " + r.RootSolutionComponentId);
            text.AppendLine("RootComponentBehavior: " + r.RootComponentBehavior).AppendLine("IsMetadata: " + r.IsMetadata);
            text.AppendLine("Identity resolution: " + o.Identity.Status).AppendLine("Portable identity: " + o.Identity.ComparisonKey);
            text.AppendLine("Placement: " + o.Placement + "; multiple parents: " + o.HasDifferentParentAssociations + "; not top-level proof");
            foreach (var a in o.Associations) text.AppendLine(a.Kind + ": " + a.Status + "; parent occurrence=" + a.ParentOccurrenceIndex + "; provenance=" + a.Provenance + "; " + a.Diagnostic);
            text.AppendLine("Original identity diagnostic: " + o.Identity.Diagnostic);
            return text.ToString();
        }
        private static string Value(SolutionObjectsExplorerRow row, string column)
        {
            switch (column)
            {
                case "DisplayName": return row.DisplayName; case "Name": return row.Name; case "Type": return row.Type;
                case "Source": return row.Source; case "Target": return row.Target; case "ComparisonStatus": return row.ComparisonStatus;
                default: return row.ObjectField(column);
            }
        }
        private void Guard(Action action)
        { try { action(); } catch (Exception error) { selected = null; RenderingFailed?.Invoke(error); } }
        private sealed class ExplorerAccessibleRow : DataGridViewRow
        {
            protected override AccessibleObject CreateAccessibilityInstance() => new ExplorerAccessibleObject(this);
            private sealed class ExplorerAccessibleObject : DataGridViewRowAccessibleObject
            {
                private readonly DataGridViewRow row;
                internal ExplorerAccessibleObject(DataGridViewRow row) : base(row) { this.row = row; }
                public override Rectangle Bounds => row.DataGridView == null || !row.Displayed || row.Index < 0 || !row.DataGridView.IsHandleCreated
                    ? Rectangle.Empty : row.DataGridView.RectangleToScreen(row.DataGridView.GetRowDisplayRectangle(row.Index, true));
                public override string Description
                {
                    get
                    {
                        var owner = row.Tag as SolutionObjectsExplorerControl;
                        return owner == null || row.Index < 0 || row.Index >= owner.visibleRows.Count ? base.Description :
                            "Original comparison row " + owner.visibleRows[row.Index] + "; " + owner.Projection.Rows[owner.visibleRows[row.Index]].ComparisonStatus;
                    }
                }
            }
        }
    }
}
