using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using D365SolutionComparer.Models.ComponentDetails;
using D365SolutionComparer.Models.Membership;
using D365SolutionComparer.Models.Presentation;

namespace D365SolutionComparer
{
    internal sealed partial class SolutionObjectsExplorerControl
    {
        private bool makerNavigation;
        private SolutionObjectNavigation makerModel;
        private SolutionObjectNavigation.Item selectedObject;
        private IReadOnlyList<SolutionObjectNavigation.Item> visibleObjects = new SolutionObjectNavigation.Item[0];
        private IReadOnlyList<SolutionObjectNavigation.Item> scopeObjects = new SolutionObjectNavigation.Item[0];
        private string objectStatus = "All";
        private bool objectDifferences;
        private bool arrangingObjectViewport;
        internal IReadOnlyList<SolutionObjectNavigation.Item> VisibleObjects => visibleObjects;
        internal SolutionObjectNavigation.Item SelectedObject => selectedObject;
        internal bool UsesObjectNavigation => makerNavigation;
        internal event Action ObjectViewChanged;
        private sealed class ObjectBranch
        {
            internal IReadOnlyList<SolutionObjectNavigation.Item> Items;
            internal int Start, Count;
        }

        internal void UseObjectNavigation()
        {
            if (Projection != null) throw new InvalidOperationException("Choose object navigation before binding.");
            makerNavigation = true;
            comparedProperties.SizeChanged += (sender, args) => UpdateObjectViewport();
            comparedProperties.FontChanged += (sender, args) => UpdateObjectViewport();
            navigation.AccessibleName = "Solution objects by category";
            counts.AccessibleName = "Displayed solution object count";
        }
        internal void SetObjectFilters(string status, bool differences)
        { objectStatus = status ?? "All"; objectDifferences = differences; }
        private bool ObjectInventoryUnavailable => Projection != null && workspaceScope != ObjectsWorkspaceScope.Compare &&
            (workspaceScope == ObjectsWorkspaceScope.Source ? Projection.Owner.Source.State : Projection.Owner.Target.State) == MembershipSnapshotState.Unavailable;
        private string ObjectUnit => workspaceScope == ObjectsWorkspaceScope.Compare ? "object entries" : "objects";
        internal string ObjectSummary => ObjectInventoryUnavailable ? workspaceScope + " object inventory unavailable." :
            scopeObjects.Count + " " + ObjectUnit + " in " + workspaceScope + " scope. " +
            string.Join(" · ", scopeObjects.GroupBy(i => i.Status(Projection)).OrderBy(g => g.Key, StringComparer.Ordinal).Select(g => g.Key + ": " + g.Count())) +
            (workspaceScope == ObjectsWorkspaceScope.Compare ? " · Unconfirmed counterparts remain separate." : "") +
            (makerModel?.MappingLimited == true ? " · Coverage is limited to supported retained metadata." : "");

        private bool ObjectMatches(SolutionObjectNavigation.Item item) =>
            (allowedRows == null || item.Rows.Any(allowedRows.Contains)) &&
            (objectStatus == "All" || item.Status(Projection) == objectStatus) &&
            (!objectDifferences || item.Status(Projection) != "Match") &&
            (string.IsNullOrWhiteSpace(searchText) || new[] { item.Value(false), item.Value(true), item.Type }
                .Any(s => s != null && s.IndexOf(searchText, StringComparison.OrdinalIgnoreCase) >= 0));

        private void ApplyObjectFilter(IEnumerable<int> filteredRows, bool all)
        {
            var bookmark = CaptureNavigation();
            baseAllowedRows = filteredRows.ToList().AsReadOnly(); filterAll = all;
            if (baseAllowedRows.Any(i => i < 0 || i >= Projection.Rows.Count)) throw new ArgumentOutOfRangeException(nameof(filteredRows));
            allowedRows = new HashSet<int>(baseAllowedRows);
            scopeObjects = makerModel.Items(workspaceScope);
            var items = scopeObjects.Where(ObjectMatches).ToList().AsReadOnly();
            preparing = true; navigation.BeginUpdate();
            try
            {
                navigation.Nodes.Clear();
                navigation.Nodes.Add(ObjectCategory("workspace-all", "All", items, false));
                foreach (var group in scopeObjects.GroupBy(i => i.Category).OrderBy(g => g.Key, StringComparer.Ordinal))
                    navigation.Nodes.Add(ObjectCategory("objects:" + group.Key, group.Key, group.ToList().AsReadOnly(), true));
            }
            finally { navigation.EndUpdate(); preparing = false; }
            SetObjectRows(items);
            if (bookmark != null) RestoreObjectNavigation(bookmark, bookmark.Row, bookmark.Side);
            ObjectViewChanged?.Invoke();
        }
        private TreeNode ObjectCategory(string key, string label, IReadOnlyList<SolutionObjectNavigation.Item> items, bool expandable)
        {
            var node = new TreeNode(label + " (" + (ObjectInventoryUnavailable ? "not available" : items.Count.ToString()) + ")") { Name = key,
                Tag = new ObjectBranch { Items = items, Count = items.Count }, ToolTipText = ObjectInventoryUnavailable ? "Object inventory unavailable" :
                    items.Count + " " + ObjectUnit + " in " + workspaceScope + " scope; category totals are before search/status filtering" };
            if (expandable && items.Count > 0) node.Nodes.Add(new TreeNode("Expand to inspect..."));
            return node;
        }
        private TreeNode ObjectNode(SolutionObjectNavigation.Item item)
        {
            var node = new TreeNode(item.Value(false)) { Name = "object:" + item.Key, Tag = item, ToolTipText = item.Type + "; " + item.Status(Projection) };
            var children = item.Children.Where(i => SolutionObjectNavigation.InScope(i, workspaceScope)).ToList().AsReadOnly();
            if (children.Count > 0)
                node.Nodes.Add(ObjectCategory("children:Charts", "Charts", children, true));
            return node;
        }
        private void PopulateObjects(TreeNode node)
        {
            if (!(node.Tag is ObjectBranch branch) || node.Nodes.Count != 1 || node.Nodes[0].Text != "Expand to inspect...") return;
            node.Nodes.Clear();
            var entries = branch.Items.Skip(branch.Start).Take(branch.Count).Where(i => ObjectMatches(i) ||
                i.Children.Any(c => SolutionObjectNavigation.InScope(c, workspaceScope) && ObjectMatches(c))).ToList().AsReadOnly();
            if (entries.Count > PageSize)
                for (int start = 0; start < entries.Count; start += PageSize)
                {
                    int count = Math.Min(PageSize, entries.Count - start);
                    var page = new TreeNode("Items " + (start + 1) + "–" + (start + count)) { Name = "page:" + start,
                        Tag = new ObjectBranch { Items = entries, Start = start, Count = count } };
                    page.Nodes.Add(new TreeNode("Expand to inspect...")); node.Nodes.Add(page);
                }
            else foreach (var item in entries) node.Nodes.Add(ObjectNode(item));
        }
        private void SelectObjectNode(TreeNode node)
        {
            if (preparing) return;
            if (node?.Tag is ObjectBranch branch)
                SetObjectRows(branch.Items.Skip(branch.Start).Take(branch.Count).Where(ObjectMatches).ToList().AsReadOnly());
            else if (node?.Tag is SolutionObjectNavigation.Item item)
            {
                SetObjectRows(ObjectMatches(item) ? new[] { item } : new SolutionObjectNavigation.Item[0]);
                if (visibleObjects.Count == 1) grid.CurrentCell = grid.Rows[0].Cells[0];
                SelectObjectItem(item);
            }
            else SetObjectRows(new SolutionObjectNavigation.Item[0]);
        }
        private void SetObjectRows(IReadOnlyList<SolutionObjectNavigation.Item> items)
        {
            ClearTechnicalContext(); selectedObject = null; selected = null; SelectedOccurrenceIndex = null; SelectedSide = null;
            visibleObjects = items; visibleRows = items.Select(i => i.Rows[0]).ToList().AsReadOnly();
            preparing = true;
            try { grid.CurrentCell = null; grid.RowCount = items.Count; }
            finally { preparing = false; }
            grid.Invalidate(); // Same-size category replacement must repaint values from the new objects.
            counts.Text = ObjectInventoryUnavailable ? workspaceScope + " object inventory unavailable." : items.Count + " " + ObjectUnit + " displayed · " + workspaceScope + " scope";
            counts.AccessibleDescription = counts.Text; countToolTip.SetToolTip(counts, counts.Text);
            RenderObjectDetails(); SelectedComparisonChanged?.Invoke(null);
        }
        private string ObjectValue(SolutionObjectNavigation.Item item, string column)
        {
            switch (column)
            {
                case "DisplayName": return item.Value(false);
                case "Name": return item.Value(true);
                case "Type": return item.Type;
                case "Source": return item.Presence(DisplayEnvironmentSide.Source, Projection);
                case "Target": return item.Presence(DisplayEnvironmentSide.Target, Projection);
                case "ComparisonStatus": return item.Status(Projection);
                default: return "[Not available]";
            }
        }
        private void SelectObjectGridRow(int position)
        {
            if (preparing || position < 0 || position >= visibleObjects.Count) return;
            SelectObjectItem(visibleObjects[position]);
        }
        private void SelectObjectItem(SolutionObjectNavigation.Item item)
        {
            selectedObject = item;
            selected = selectedObject.Rows.Count == 1 ? Projection.Rows[selectedObject.Rows[0]].Original : null;
            SelectedSide = ScopeSide() ?? (selectedObject.Locals.Count == 1 ? (DisplayEnvironmentSide?)selectedObject.Locals[0].Side : null);
            SelectedOccurrenceIndex = selectedObject.Locals.Count == 1 && selectedObject.Locals[0].Occurrences.Count == 1 ?
                (int?)selectedObject.Locals[0].Occurrences[0] : null;
            ClearTechnicalContext();
            if (selected != null) SetTechnicalContext(TechnicalRoute.Row, selected, SelectedOccurrenceIndex);
            RenderObjectDetails(); SelectedComparisonChanged?.Invoke(selected);
        }
        private void RenderObjectDetails()
        {
            details.Visible = false; comparedPropertiesPanel.Visible = true;
            if (simplifiedIdentity == null) return;
            simplifiedIdentity.Text = selectedObject == null ? "Select an object to inspect its details." :
                "Display Name: " + selectedObject.Value(false) + Environment.NewLine +
                (selectedObject.Category == "Dashboards" || selectedObject.Category == "Reports" || selectedObject.Category == "Team template" ||
                    selectedObject.Category == "Secured Masking Rule" || selectedObject.Category == "Processes" ||
                    selectedObject.Category == "Email templates" ? "Name: " : "Logical / Schema Name: ") + selectedObject.Value(true) + Environment.NewLine +
                "Component Type: " + selectedObject.Type + " · Result: " + selectedObject.Status(Projection) + Environment.NewLine +
                "Source: " + selectedObject.Presence(DisplayEnvironmentSide.Source, Projection) + " · Target: " + selectedObject.Presence(DisplayEnvironmentSide.Target, Projection);
            var row = selected == null ? null : Projection.Rows[selected.OriginalRowIndex];
            bool ownerChanged = !ReferenceEquals(row, comparedPropertiesOwner);
            comparedPropertiesOwner = row;
            comparedPropertyRows = row == null ? (IReadOnlyList<ComponentPropertyPresentation>)new ComponentPropertyPresentation[0] : row.OrderedComparedProperties;
            comparedProperties.RowCount = comparedPropertyRows.Count; comparedProperties.Tag = selected?.AuthoritativeRow.DefinitionDetail;
            comparedPropertiesMessage.Text = selectedObject == null ? "No object selected." : selectedObject.Rows.Count != 1 ?
                "The retained results do not establish one property comparison for this object." :
                row.ComparedPropertiesExplanation;
            comparedProperties.AccessibleDescription = comparedPropertiesMessage.Text;
            countToolTip.SetToolTip(comparedPropertiesMessage, comparedPropertiesMessage.Text);
            if (ownerChanged && comparedProperties.RowCount > 0)
            {
                comparedProperties.ClearSelection();
                comparedProperties.CurrentCell = comparedProperties.Rows[0].Cells[0];
                comparedProperties.FirstDisplayedScrollingRowIndex = 0;
            }
            comparedProperties.Invalidate();
            openRow.Enabled = selected != null; definition.Enabled = false;
            UpdateObjectViewport();
        }
        private void ClearObjectOwner()
        {
            // The current-cell commit can request its old virtual value. Keep that backing
            // collection valid until the grid has finished clearing the current cell and rows.
            grid.CurrentCell = null; grid.RowCount = 0;
            selectedObject = null; visibleObjects = new SolutionObjectNavigation.Item[0]; scopeObjects = visibleObjects;
            SelectedSide = null; SelectedOccurrenceIndex = null; visibleRows = new int[0];
            comparedPropertyRows = new ComponentPropertyPresentation[0]; comparedProperties.RowCount = 0; comparedProperties.Tag = null;
            simplifiedIdentity?.Clear();
        }
        private void UpdateObjectViewport()
        {
            if (!makerNavigation || arrangingObjectViewport || !Visible || evidenceSplit.Panel2Collapsed || simplifiedIdentity == null) return;
            arrangingObjectViewport = true;
            try
            {
                if (comparedProperties.RowCount == 0) { detailsScrollCue.Visible = false; return; }
                evidenceSplit.Panel2.PerformLayout(); comparedPropertiesPanel.PerformLayout();
                int required = comparedProperties.ColumnHeadersHeight + comparedProperties.RowTemplate.Height + SystemInformation.BorderSize.Height * 2;
                int deficit = required - comparedProperties.ClientSize.Height;
                if (deficit > 0 && !evidenceSplit.Panel1Collapsed && evidenceSplit.SplitterDistance > evidenceSplit.Panel1MinSize)
                {
                    evidenceSplit.SplitterDistance = Math.Max(evidenceSplit.Panel1MinSize, evidenceSplit.SplitterDistance - deficit);
                    evidenceSplit.Panel2.PerformLayout(); comparedPropertiesPanel.PerformLayout();
                }
                detailsScrollCue.Text = "Expand Details to view more properties.";
                detailsScrollCue.Visible = comparedProperties.ClientSize.Height < required;
            }
            finally { arrangingObjectViewport = false; }
        }
        private void RestoreObjectNavigation(NavigationBookmark bookmark, int? row, DisplayEnvironmentSide? side)
        {
            if (bookmark == null || !ReferenceEquals(bookmark.Owner, Projection) || IsDisposed || Disposing) return;
            bool prior = restoringNavigation; restoringNavigation = true;
            try
            {
                foreach (string path in bookmark.ExpandedPaths ?? new string[0]) ResolveNavigationPath(path)?.Expand();
                var node = ResolveNavigationPath(bookmark.SelectedPath);
                if (node != null) { navigation.SelectedNode = node; SelectObjectNode(node); }
                if (bookmark.ObjectKey != null)
                {
                    var item = visibleObjects.SingleOrDefault(i => i.Locals.Any(l => l.Key == bookmark.ObjectKey));
                    if (item != null)
                    { int position = visibleObjects.ToList().IndexOf(item); grid.CurrentCell = grid.Rows[position].Cells[0]; SelectObjectGridRow(position); }
                    else
                    {
                        // Retain a valid inactive selection without inserting it into the other side's inventory.
                        var retained = makerModel.Items(ObjectsWorkspaceScope.Compare).SingleOrDefault(i =>
                            i.Locals.Any(l => l.Key == bookmark.ObjectKey) && (!row.HasValue || i.Rows.Contains(row.Value)));
                        if (retained != null && !SolutionObjectNavigation.InScope(retained, workspaceScope))
                        {
                            selectedObject = retained;
                            selected = retained.Rows.Count == 1 ? Projection.Rows[retained.Rows[0]].Original : null;
                            SelectedSide = side;
                            var local = retained.Locals.SingleOrDefault(l => l.Key == bookmark.ObjectKey);
                            SelectedOccurrenceIndex = local != null && local.Occurrences.Count == 1 ? (int?)local.Occurrences[0] : null;
                            ClearTechnicalContext();
                            if (selected != null) SetTechnicalContext(TechnicalRoute.RowOnly, selected);
                            RenderObjectDetails(); counts.Text += " · Selected object is outside this list/scope/filter.";
                            SelectedComparisonChanged?.Invoke(selected);
                        }
                    }
                }
                var top = ResolveNavigationPath(bookmark.TopPath); if (top != null) navigation.TopNode = top;
                RestoreGridScroll(bookmark);
            }
            finally { restoringNavigation = prior; }
        }
    }
}
