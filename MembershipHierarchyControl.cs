using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using D365SolutionComparer.Models.Presentation;

namespace D365SolutionComparer
{
    /// <summary>Native WinForms tree, lazily populated in bounded pages. No SDK or comparison operations.</summary>
    internal sealed class MembershipHierarchyControl : UserControl
    {
        internal const int PageSize = 200;
        private readonly TreeView tree;
        private readonly TextBox details;
        private readonly LinkLabel definition;
        private readonly Font contextFont;
        private readonly Func<MembershipGridRow, string> describeRow;
        private readonly Action<MembershipGridRow> openDefinition;
        private MembershipHierarchyView currentView;
        internal Action<Exception> RenderingFailed { get; set; }
        internal event Action<MembershipGridRow> SelectedComparisonChanged;
        internal MembershipGridRow SelectedComparison { get; private set; }
        internal TreeView Tree => tree;
        internal D365SolutionComparer.Models.Membership.DisplayEnvironmentSide? SelectedSide =>
            (tree.SelectedNode?.Tag as MembershipHierarchyNode)?.Occurrence?.Side;

        internal MembershipHierarchyControl(Func<MembershipGridRow, string> describeRow, Action<MembershipGridRow> openDefinition)
        {
            this.describeRow = describeRow; this.openDefinition = openDefinition;
            Name = "MembershipHierarchy"; Size = new Size(1100, 500); Dock = DockStyle.Fill;
            contextFont = new Font(Font, FontStyle.Italic);
            var split = new SplitContainer { Size = new Size(1100, 500), Dock = DockStyle.Fill, SplitterWidth = 6,
                SplitterDistance = 550, Panel1MinSize = 200, Panel2MinSize = 200 };
            tree = new TreeView { Name = "ComponentHierarchyTree", Dock = DockStyle.Fill, HideSelection = false,
                ShowNodeToolTips = true, AccessibleName = "Read-only component hierarchy",
                AccessibleDescription = "Raw occurrences remain separate. Inclusion-root links do not establish structural ownership. Context ancestors are not filtered comparison results." };
            details = new TextBox { Name = "HierarchyEvidence", Dock = DockStyle.Fill, Multiline = true, ReadOnly = true,
                ScrollBars = ScrollBars.Both, WordWrap = false, BackColor = Color.White,
                AccessibleName = "Selected occurrence and authoritative comparison evidence" };
            definition = new LinkLabel { Text = "Open existing definition details...", Dock = DockStyle.Bottom, Height = 28, Enabled = false };
            definition.LinkClicked += (s, e) => Guard(() => { if (SelectedComparison != null) openDefinition(SelectedComparison); });
            tree.BeforeExpand += (s, e) => Guard(() => Populate(e.Node));
            tree.AfterSelect += (s, e) => Guard(() => Select(e.Node));
            tree.NodeMouseDoubleClick += (s, e) => Guard(() =>
            { if ((e.Node.Tag as MembershipHierarchyNode)?.Row?.AuthoritativeRow.DefinitionDetail != null) openDefinition(((MembershipHierarchyNode)e.Node.Tag).Row); });
            split.Panel1.Controls.Add(tree); split.Panel2.Controls.Add(details); split.Panel2.Controls.Add(definition);
            Controls.Add(split);
        }

        internal void Bind(MembershipHierarchyView view)
        {
            currentView = view;
            tree.BeginUpdate();
            try
            {
                tree.Nodes.Clear(); SelectedComparison = null; definition.Enabled = false;
                details.Text = "Select an occurrence or linked comparison row. Inclusion-root links and structural parent links are distinct evidence.";
                var root = MakeNode(view.Root); tree.Nodes.Add(root); root.Expand();
                foreach (TreeNode side in root.Nodes) if (side.Tag is MembershipHierarchyNode &&
                    (side.Text.StartsWith("Source (", StringComparison.Ordinal) || side.Text.StartsWith("Source: ", StringComparison.Ordinal))) side.Expand();
            }
            finally { tree.EndUpdate(); }
            SelectedComparisonChanged?.Invoke(null);
        }

        private sealed class Page
        {
            internal IReadOnlyList<MembershipHierarchyNode> Items;
            internal int Start, Count;
        }
        internal void ClearHierarchy()
        { tree.Nodes.Clear(); currentView = null; SelectedComparison = null; definition.Enabled = false; details.Clear(); }

        // Locate an original comparison row in the detached model; materialize only its bounded
        // native navigation path. No authoritative projection or comparison is rebuilt.
        internal void SelectComparison(int rowIndex, D365SolutionComparer.Models.Membership.DisplayEnvironmentSide? preferredSide = null)
        {
            if (currentView == null || tree.Nodes.Count == 0) return;
            var parents = new Dictionary<MembershipHierarchyNode, MembershipHierarchyNode>();
            var stack = new Stack<MembershipHierarchyNode>(); stack.Push(currentView.Root);
            MembershipHierarchyNode selected = null, fallback = null;
            while (stack.Count > 0)
            {
                var node = stack.Pop();
                if (!node.NavigationContext && node.Row?.OriginalRowIndex == rowIndex)
                {
                    if (fallback == null) fallback = node;
                    if (node.Occurrence != null && (!preferredSide.HasValue || node.Occurrence.Side == preferredSide.Value))
                    { selected = node; break; }
                }
                foreach (var child in node.Children.Reverse()) { parents[child] = node; stack.Push(child); }
            }
            selected = selected ?? fallback; if (selected == null) return;
            var route = new Stack<MembershipHierarchyNode>(); var current = selected;
            while (current != currentView.Root) { route.Push(current); current = parents[current]; }
            var native = tree.Nodes[0]; var model = currentView.Root;
            tree.BeginUpdate();
            try
            {
                while (route.Count > 0)
                {
                    var next = route.Pop(); native.Expand(); Populate(native);
                    int index = 0; while (model.Children[index] != next) index++;
                    var pageNode = native.Nodes.Cast<TreeNode>().FirstOrDefault(n => n.Tag is Page page && index >= page.Start && index < page.Start + page.Count);
                    if (pageNode != null) { pageNode.Expand(); Populate(pageNode); native = pageNode; }
                    native = native.Nodes.Cast<TreeNode>().First(n => ReferenceEquals(n.Tag, next)); model = next;
                }
                tree.SelectedNode = native; native.EnsureVisible();
            }
            finally { tree.EndUpdate(); }
        }
        private TreeNode MakeNode(MembershipHierarchyNode item)
        {
            var node = new TreeNode(item.Text) { Tag = item, ToolTipText = item.NavigationContext ?
                "Navigation context only — not a filtered comparison result." : item.Occurrence == null ? item.Text :
                "Raw occurrence " + (item.Occurrence.OccurrenceIndex + 1) + "; placement: " + item.Occurrence.Placement +
                "; linked comparison rows: " + item.Occurrence.ComparisonRowIndices.Count };
            if (item.NavigationContext) { node.ForeColor = Color.DimGray; node.NodeFont = contextFont; }
            if (item.Children.Count > 0) node.Nodes.Add(new TreeNode("Expand to inspect..."));
            return node;
        }
        private void Populate(TreeNode node)
        {
            if (node.Nodes.Count != 1 || node.Nodes[0].Tag != null || node.Nodes[0].Text != "Expand to inspect...") return;
            var item = node.Tag as MembershipHierarchyNode; var page = node.Tag as Page;
            var items = item?.Children ?? page?.Items;
            if (items == null) return;
            node.Nodes.Clear();
            if (page != null)
                for (int i = page.Start; i < page.Start + page.Count; i++) node.Nodes.Add(MakeNode(items[i]));
            else if (items.Count <= PageSize)
                foreach (var child in items) node.Nodes.Add(MakeNode(child));
            else
                for (int start = 0; start < items.Count; start += PageSize)
                {
                    var bucket = new TreeNode("Items " + (start + 1) + "–" + Math.Min(items.Count, start + PageSize))
                    { Tag = new Page { Items = items, Start = start, Count = Math.Min(PageSize, items.Count - start) } };
                    bucket.Nodes.Add(new TreeNode("Expand to inspect...")); node.Nodes.Add(bucket);
                }
        }
        private void Select(TreeNode selected)
        {
            var item = selected?.Tag as MembershipHierarchyNode;
            SelectedComparison = item?.NavigationContext == true ? null : item?.Row;
            definition.Enabled = SelectedComparison?.AuthoritativeRow.DefinitionDetail != null;
            var text = new StringBuilder(item == null ? "Navigation page. Expand to inspect the retained entries." : MembershipHierarchyView.Describe(item));
            if (SelectedComparison != null) text.AppendLine().AppendLine(describeRow(SelectedComparison));
            details.Text = text.ToString();
            SelectedComparisonChanged?.Invoke(SelectedComparison);
        }
        private void Guard(Action action)
        {
            try { action(); }
            catch (Exception ex) { tree.Nodes.Clear(); SelectedComparison = null; RenderingFailed?.Invoke(ex); }
        }
        protected override void Dispose(bool disposing)
        { base.Dispose(disposing); if (disposing) contextFont.Dispose(); }
    }
}
