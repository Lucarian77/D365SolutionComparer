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
        private object navigationOwner;
        private bool detailsDirty;
        private string lastNavigationFocus = "tree";
        private bool restoringNavigation, explicitContextIntent;
        private long navigationIntentRevision;
        private NavigationBookmark pendingContext;
        internal Action NavigationIntentChanged { get; set; }
        internal sealed class NavigationBookmark
        {
            internal object Owner;
            internal MembershipHierarchyView View;
            internal object[] SelectedPath, TopPath;
            internal List<object[]> ExpandedPaths;
            internal int? Row, Occurrence;
            internal D365SolutionComparer.Models.Membership.DisplayEnvironmentSide? Side;
            internal D365SolutionComparer.Models.Membership.MembershipSnapshot Snapshot;
            internal string Focus;
            internal bool ExplicitContext;
            internal long IntentRevision;
            internal int ContextDepth;
            internal int? ContextPageStart;
            internal ComponentHierarchyOccurrence ContextAnchor;
        }
        internal Action<Exception> RenderingFailed { get; set; }
        internal event Action<MembershipGridRow> SelectedComparisonChanged;
        internal MembershipGridRow SelectedComparison { get; private set; }
        internal TreeView Tree => tree;
        private ComponentHierarchyOccurrence SelectedOccurrence
        {
            get
            {
                for (var node = tree.SelectedNode; node != null; node = node.Parent)
                    if ((node.Tag as MembershipHierarchyNode)?.Occurrence != null) return ((MembershipHierarchyNode)node.Tag).Occurrence;
                return null;
            }
        }
        internal int? SelectedOccurrenceIndex => SelectedOccurrence?.OccurrenceIndex;
        internal D365SolutionComparer.Models.Membership.DisplayEnvironmentSide? SelectedSide => SelectedOccurrence?.Side;

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
            tree.AfterSelect += (s, e) => Guard(() =>
            {
                Select(e.Node);
                if (!restoringNavigation)
                {
                    pendingContext = null; explicitContextIntent = SelectedComparison == null; navigationIntentRevision++;
                    NavigationIntentChanged?.Invoke();
                }
            });
            tree.NodeMouseDoubleClick += (s, e) => Guard(() =>
            { if ((e.Node.Tag as MembershipHierarchyNode)?.Row?.AuthoritativeRow.DefinitionDetail != null) openDefinition(((MembershipHierarchyNode)e.Node.Tag).Row); });
            split.Panel1.Controls.Add(tree); split.Panel2.Controls.Add(details); split.Panel2.Controls.Add(definition);
            Controls.Add(split);
            tree.Enter += (s, e) => lastNavigationFocus = "tree";
            details.Enter += (s, e) => lastNavigationFocus = "details";
            tree.KeyDown += (s, e) =>
            {
                if (e.KeyCode != Keys.Enter || tree.SelectedNode == null || SelectedComparison != null) return;
                if (tree.SelectedNode.IsExpanded) tree.SelectedNode.Collapse(); else tree.SelectedNode.Expand();
                e.Handled = e.SuppressKeyPress = true;
            };
        }

        internal void Bind(MembershipHierarchyView view, object owner = null)
        {
            bool wasRestoring = restoringNavigation; restoringNavigation = true;
            currentView = view; navigationOwner = owner ?? view; details.Clear(); detailsDirty = false;
            explicitContextIntent = false; pendingContext = null;
            tree.BeginUpdate();
            try
            {
                tree.Nodes.Clear(); SelectedComparison = null; definition.Enabled = false;
                details.Text = "Select an occurrence or linked comparison row. Inclusion-root links and structural parent links are distinct evidence.";
                var root = MakeNode(view.Root); tree.Nodes.Add(root); root.Expand();
                foreach (TreeNode side in root.Nodes) if (side.Tag is MembershipHierarchyNode &&
                    (side.Text.StartsWith("Source (", StringComparison.Ordinal) || side.Text.StartsWith("Source: ", StringComparison.Ordinal))) side.Expand();
            }
            finally { tree.EndUpdate(); restoringNavigation = wasRestoring; }
            SelectedComparisonChanged?.Invoke(null);
        }

        private sealed class Page
        {
            internal IReadOnlyList<MembershipHierarchyNode> Items;
            internal int Start, Count;
        }
        internal void ClearHierarchy()
        {
            bool wasRestoring = restoringNavigation; restoringNavigation = true;
            try { tree.Nodes.Clear(); currentView = null; navigationOwner = null; SelectedComparison = null; definition.Enabled = false; details.Clear(); detailsDirty = false; explicitContextIntent = false; pendingContext = null; }
            finally { restoringNavigation = wasRestoring; }
        }

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
                explicitContextIntent = false; pendingContext = null; SelectRestoredNode(native); native.EnsureVisible();
                if (Visible) { navigationIntentRevision++; NavigationIntentChanged?.Invoke(); }
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
            details.Clear(); detailsDirty = true;
            if (Visible) RenderSelectedEvidence();
            SelectedComparisonChanged?.Invoke(SelectedComparison);
        }
        private void RenderSelectedEvidence()
        {
            if (!Visible || IsDisposed || Disposing || currentView == null || !detailsDirty) return;
            var view = currentView; var native = tree.SelectedNode; var row = SelectedComparison;
            var item = native?.Tag as MembershipHierarchyNode; detailsDirty = false;
            var text = new StringBuilder(item == null ? "Navigation page. Expand to inspect the retained entries." : MembershipHierarchyView.Describe(item));
            if (row != null) text.AppendLine().AppendLine(describeRow(row));
            if (!Visible || IsDisposed || Disposing || !ReferenceEquals(view, currentView) || !ReferenceEquals(native, tree.SelectedNode) || !ReferenceEquals(row, SelectedComparison)) return;
            details.Text = text.ToString();
        }
        protected override void OnVisibleChanged(EventArgs e)
        {
            base.OnVisibleChanged(e);
            if (details == null || tree == null) return;
            if (!Visible) { details.Clear(); detailsDirty = currentView != null; }
            else Guard(RenderSelectedEvidence);
        }
        private static IEnumerable<TreeNode> NativeNodes(TreeNodeCollection nodes)
        {
            foreach (TreeNode node in nodes)
            {
                if (node.Tag == null) continue;
                yield return node;
                foreach (var child in NativeNodes(node.Nodes)) yield return child;
            }
        }
        private static object[] NativePath(TreeNode node)
        {
            if (node == null) return null;
            var path = new Stack<object>();
            while (node != null) { path.Push(node.Tag); node = node.Parent; }
            return path.ToArray();
        }
        private TreeNode ResolvePath(object[] path)
        {
            if (path == null) return null;
            var nodes = tree.Nodes; TreeNode found = null;
            for (int i = 0; i < path.Length; i++)
            {
                var key = path[i];
                found = nodes.Cast<TreeNode>().FirstOrDefault(n => ReferenceEquals(n.Tag, key) ||
                    n.Tag is Page current && key is Page saved && ReferenceEquals(current.Items, saved.Items) && current.Start == saved.Start && current.Count == saved.Count);
                if (found == null) return null;
                if (i < path.Length - 1) { found.Expand(); Populate(found); nodes = found.Nodes; }
            }
            return found;
        }
        internal NavigationBookmark CaptureNavigation()
        {
            if (pendingContext != null && explicitContextIntent && SelectedComparison == null)
            { pendingContext.Focus = lastNavigationFocus; return pendingContext; }
            var path = NativePath(tree.SelectedNode); var model = tree.SelectedNode?.Tag as MembershipHierarchyNode;
            var page = tree.SelectedNode?.Tag as Page;
            var anchor = model == null ? null : FindModel(model, n => n.Occurrence != null)?.Occurrence;
            if (anchor == null && page != null)
                anchor = page.Items.Skip(page.Start).Take(page.Count).Select(n => FindModel(n, child => child.Occurrence != null)?.Occurrence).FirstOrDefault(o => o != null);
            anchor = anchor ?? SelectedOccurrence;
            return new NavigationBookmark
            {
                Owner = navigationOwner, View = currentView, SelectedPath = path, TopPath = NativePath(tree.TopNode),
                ExpandedPaths = NativeNodes(tree.Nodes).Where(n => n.IsExpanded).Select(NativePath).ToList(),
                Row = SelectedComparison?.OriginalRowIndex, Side = SelectedSide, Occurrence = SelectedOccurrenceIndex,
                Snapshot = SelectedOccurrence?.Snapshot, Focus = lastNavigationFocus, ExplicitContext = explicitContextIntent,
                IntentRevision = navigationIntentRevision, ContextDepth = path == null ? -1 : path.OfType<MembershipHierarchyNode>().Count() - 1,
                ContextPageStart = page == null ? (int?)null : page.Start, ContextAnchor = anchor
            };
        }
        internal void RestoreNavigation(NavigationBookmark bookmark, int? row, D365SolutionComparer.Models.Membership.DisplayEnvironmentSide? side)
        {
            if (bookmark == null || !ReferenceEquals(bookmark.Owner, navigationOwner) || currentView == null || IsDisposed || Disposing ||
                bookmark.IntentRevision != navigationIntentRevision || bookmark.Row != row || bookmark.Side != side) return;
            bool wasRestoring = restoringNavigation; restoringNavigation = true;
            try
            {
                if (ReferenceEquals(bookmark.View, currentView))
                {
                    foreach (var path in bookmark.ExpandedPaths) ResolvePath(path)?.Expand();
                    var native = ResolvePath(bookmark.SelectedPath);
                    if (native != null) SelectRestoredNode(native);
                    else if (bookmark.ExplicitContext) SelectRestoredNode(null);
                    var top = ResolvePath(bookmark.TopPath); if (top != null) tree.TopNode = top;
                }
                else if (bookmark.ExplicitContext && !row.HasValue)
                {
                    var context = ResolveContext(bookmark);
                    var contextModel = context?.Tag as MembershipHierarchyNode;
                    if (contextModel?.Row != null && !contextModel.NavigationContext) context = null;
                    SelectRestoredNode(context);
                    pendingContext = context == null && bookmark.SelectedPath != null ? bookmark : null;
                }
                else if (row.HasValue && bookmark.Occurrence.HasValue)
                {
                    var target = FindModel(currentView.Root, n => n.Occurrence != null && n.Occurrence.OccurrenceIndex == bookmark.Occurrence &&
                        n.Occurrence.Side == side && ReferenceEquals(n.Occurrence.Snapshot, bookmark.Snapshot) && !n.NavigationContext);
                    if (target != null)
                    {
                        if (target.Row?.OriginalRowIndex != row) target = target.Children.FirstOrDefault(n => n.Row?.OriginalRowIndex == row);
                        if (target != null) SelectModel(target);
                    }
                }
                explicitContextIntent = bookmark.ExplicitContext;
                lastNavigationFocus = bookmark.Focus ?? "tree";
            }
            finally { restoringNavigation = wasRestoring; }
        }
        private TreeNode ResolveContext(NavigationBookmark bookmark)
        {
            if (bookmark.ContextDepth < 0 || tree.Nodes.Count == 0) return null;
            if (bookmark.ContextDepth == 0) return tree.Nodes[0];
            var anchor = bookmark.ContextAnchor; if (anchor == null) return null;
            if (bookmark.ContextDepth == 1 && currentView.Root.Children.Count >= 2)
                return MaterializeModel(currentView.Root.Children[anchor.Side == D365SolutionComparer.Models.Membership.DisplayEnvironmentSide.Source ? 0 : 1]);
            var target = FindModel(currentView.Root, n => n.Occurrence != null && n.Occurrence.OccurrenceIndex == anchor.OccurrenceIndex &&
                n.Occurrence.Side == anchor.Side && ReferenceEquals(n.Occurrence.Snapshot, anchor.Snapshot));
            if (target == null) return null;
            var path = ModelPath(target);
            if (path == null || bookmark.ContextDepth >= path.Count) return null;
            var native = MaterializeModel(path[bookmark.ContextDepth]);
            if (native == null || !bookmark.ContextPageStart.HasValue) return native;
            native.Expand(); Populate(native);
            return native.Nodes.Cast<TreeNode>().FirstOrDefault(n => n.Tag is Page page && page.Start == bookmark.ContextPageStart.Value);
        }
        private List<MembershipHierarchyNode> ModelPath(MembershipHierarchyNode target)
        {
            var parents = new Dictionary<MembershipHierarchyNode, MembershipHierarchyNode>();
            FindModel(currentView.Root, n => { foreach (var child in n.Children) parents[child] = n; return false; });
            var path = new Stack<MembershipHierarchyNode>(); var model = target;
            while (!ReferenceEquals(model, currentView.Root))
            {
                path.Push(model); if (!parents.TryGetValue(model, out model)) return null;
            }
            path.Push(currentView.Root); return path.ToList();
        }
        private void SelectRestoredNode(TreeNode native)
        {
            bool wasRestoring = restoringNavigation; restoringNavigation = true;
            try
            {
                if (native == null) { tree.SelectedNode = null; Select(null); }
                else if (ReferenceEquals(tree.SelectedNode, native)) Select(native);
                else tree.SelectedNode = native;
            }
            finally { restoringNavigation = wasRestoring; }
        }
        internal void ClearNavigationSelection()
        {
            pendingContext = null; explicitContextIntent = true; navigationIntentRevision++;
            SelectRestoredNode(null); NavigationIntentChanged?.Invoke();
        }
        private static MembershipHierarchyNode FindModel(MembershipHierarchyNode root, Func<MembershipHierarchyNode, bool> match)
        {
            var stack = new Stack<MembershipHierarchyNode>(); stack.Push(root);
            while (stack.Count > 0)
            {
                var node = stack.Pop(); if (match(node)) return node;
                foreach (var child in node.Children.Reverse()) stack.Push(child);
            }
            return null;
        }
        private TreeNode MaterializeModel(MembershipHierarchyNode target)
        {
            if (currentView == null || tree.Nodes.Count == 0) return null;
            var path = ModelPath(target); if (path == null) return null;
            var native = tree.Nodes[0]; var model = currentView.Root;
            foreach (var next in path.Skip(1))
            {
                native.Expand(); Populate(native);
                int index = model.Children.ToList().IndexOf(next);
                var page = native.Nodes.Cast<TreeNode>().FirstOrDefault(n => n.Tag is Page p && index >= p.Start && index < p.Start + p.Count);
                if (page != null) { page.Expand(); Populate(page); native = page; }
                native = native.Nodes.Cast<TreeNode>().First(n => ReferenceEquals(n.Tag, next)); model = next;
            }
            return native;
        }
        private void SelectModel(MembershipHierarchyNode target)
        {
            var native = MaterializeModel(target); if (native != null) SelectRestoredNode(native);
        }
        internal void ResetNavigation()
        {
            if (tree.Nodes.Count == 0) return;
            pendingContext = null; explicitContextIntent = true;
            SelectRestoredNode(tree.Nodes[0]);
            if (Visible) { navigationIntentRevision++; NavigationIntentChanged?.Invoke(); }
        }
        internal void RestoreNavigationFocus()
        {
            if (!Visible || IsDisposed || Disposing) return;
            if (lastNavigationFocus == "details") details.Focus(); else tree.Focus();
        }

        private void Guard(Action action)
        {
            try { action(); }
            catch (Exception ex) { tree.Nodes.Clear(); SelectedComparison = null; details.Clear(); detailsDirty = false; RenderingFailed?.Invoke(ex); }
        }
        protected override void Dispose(bool disposing)
        { base.Dispose(disposing); if (disposing) contextFont.Dispose(); }
    }
}
