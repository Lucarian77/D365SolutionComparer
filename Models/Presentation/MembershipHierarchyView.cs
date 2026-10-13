using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using D365SolutionComparer.Models.Membership;

namespace D365SolutionComparer.Models.Presentation
{
    internal sealed class MembershipHierarchyNode
    {
        private readonly List<MembershipHierarchyNode> children = new List<MembershipHierarchyNode>();
        private readonly string navigationDetails;
        private readonly Func<string> describePresentation;
        internal MembershipHierarchyNode(string text, ComponentHierarchyOccurrence occurrence = null,
            MembershipGridRow row = null, bool context = false, string details = null, string displayName = null,
            Func<string> describePresentation = null)
        { Text = text; Occurrence = occurrence; Row = row; NavigationContext = context; navigationDetails = details;
            DisplayName = displayName; this.describePresentation = describePresentation; }
        internal string Text { get; private set; }
        internal string NavigationDetails => describePresentation == null ? navigationDetails : describePresentation();
        internal string DisplayName { get; }
        internal void LabelCount() { Text += " (" + children.Count + ")"; }
        internal ComponentHierarchyOccurrence Occurrence { get; }
        internal MembershipGridRow Row { get; }
        internal bool NavigationContext { get; }
        internal IReadOnlyList<MembershipHierarchyNode> Children => children.AsReadOnly();
        internal void Add(MembershipHierarchyNode node) => children.Add(node);
    }

    /// <summary>H2A navigation only. Each occurrence and row retains its H1 locator and authority.</summary>
    internal sealed class MembershipHierarchyView
    {
        private MembershipHierarchyView(MembershipHierarchyNode root, int rows, int occurrences)
        { Root = root; ComparisonRowCount = rows; VisibleOccurrenceCount = occurrences; }
        internal MembershipHierarchyNode Root { get; }
        internal int ComparisonRowCount { get; }
        internal int VisibleOccurrenceCount { get; }

        internal static MembershipHierarchyView Create(ComponentHierarchyProjection projection,
            IReadOnlyList<MembershipGridRow> rows, IEnumerable<int> visibleRows, bool unfiltered,
            CancellationToken token = default(CancellationToken))
            => CreateCore(projection, rows, visibleRows, unfiltered, token, false);

        // Explicit administrator presentation; the original H2A adapter contract remains compatible.
        internal static MembershipHierarchyView CreateAdministratorView(ComponentHierarchyProjection projection,
            IReadOnlyList<MembershipGridRow> rows, IEnumerable<int> visibleRows, bool unfiltered,
            CancellationToken token = default(CancellationToken))
            => CreateCore(projection, rows, visibleRows, unfiltered, token, true);

        private static MembershipHierarchyView CreateCore(ComponentHierarchyProjection projection,
            IReadOnlyList<MembershipGridRow> rows, IEnumerable<int> visibleRows, bool unfiltered,
            CancellationToken token, bool administrator)
        {
            if (projection == null || projection.Outcome != HierarchyProjectionOutcome.Completed)
                throw new InvalidOperationException("Hierarchy projection is unavailable.");
            if (rows == null || rows.Count != projection.ComparisonRows.Count ||
                rows.Where((r, i) => r.OriginalRowIndex != i || !ReferenceEquals(r.AuthoritativeRow, projection.ComparisonRows[i].Row)).Any())
                throw new InvalidOperationException("Hierarchy row scope mismatch.");
            var visible = new HashSet<int>(visibleRows);
            if (visible.Any(i => i < 0 || i >= rows.Count)) throw new ArgumentOutOfRangeException(nameof(visibleRows));
            var parent = new Dictionary<int, int>();
            foreach (var n in projection.Occurrences)
            {
                token.ThrowIfCancellationRequested();
                if (n.Placement == HierarchyPlacement.ParentNotEstablished || n.HasDifferentParentAssociations) continue;
                var kind = n.Placement == HierarchyPlacement.VerifiedInclusionRoot
                    ? HierarchyRelationshipKind.InclusionRoot : HierarchyRelationshipKind.StructuralParent;
                var p = n.Associations.Where(a => a.Kind == kind && a.Status == HierarchyAssociationStatus.Verified)
                    .Select(a => a.ParentOccurrenceIndex).Distinct().ToList();
                if (p.Count == 1 && p[0].HasValue) parent.Add(n.OccurrenceIndex, p[0].Value);
            }
            var matching = new HashSet<int>(projection.Occurrences.Where(n =>
                n.ComparisonRowIndices.Any(visible.Contains)).Select(n => n.OccurrenceIndex));
            var included = unfiltered ? new HashSet<int>(projection.Occurrences.Select(n => n.OccurrenceIndex)) : new HashSet<int>();
            if (!unfiltered)
                foreach (int match in matching)
                {
                    int current = match;
                    while (included.Add(current))
                    { token.ThrowIfCancellationRequested(); if (!parent.TryGetValue(current, out current)) break; }
                }
            var root = new MembershipHierarchyNode("Solution Components — raw occurrences, not comparison totals");
            var sides = new Dictionary<DisplayEnvironmentSide, MembershipHierarchyNode>();
            foreach (var side in new[] { DisplayEnvironmentSide.Source, DisplayEnvironmentSide.Target })
            {
                var name = side == DisplayEnvironmentSide.Source ? projection.Presentation.Source.Diagnostics.EnvironmentName :
                    projection.Presentation.Target.Diagnostics.EnvironmentName;
                var node = new MembershipHierarchyNode(administrator ? side + ": " + MembershipGridRow.EscapeForGrid(name) :
                    side + " (" + MembershipGridRow.EscapeForGrid(name) + ")");
                root.Add(node); sides.Add(side, node);
            }
            var nodes = new Dictionary<int, MembershipHierarchyNode>();
            foreach (var n in projection.Occurrences.Where(n => included.Contains(n.OccurrenceIndex)))
            {
                token.ThrowIfCancellationRequested();
                var linked = n.ComparisonRowIndices.Where(visible.Contains).Select(i => rows[i]).ToList();
                var selected = linked.Count == 1 ? linked[0] : null;
                var context = !unfiltered && !matching.Contains(n.OccurrenceIndex);
                string status = linked.Count == 1 ? " — " + linked[0].MembershipStatus + " | " + linked[0].DefinitionStatusText :
                    linked.Count > 1 ? " — " + linked.Count + " linked comparison rows" : " — Comparison mapping unavailable";
                if (context) status = " — Navigation context (not a filtered result)";
                var parentTable = parent.ContainsKey(n.OccurrenceIndex) &&
                    projection.Occurrences[parent[n.OccurrenceIndex]].Identity.SemanticKind == ComponentSemanticKinds.Table
                    ? projection.Occurrences[parent[n.OccurrenceIndex]].Identity.ComparisonKey : null;
                var componentName = administrator ? AdministratorName(n, rows, parentTable) : Name(n, rows);
                var node = new MembershipHierarchyNode(componentName + status + " [occurrence " + (n.OccurrenceIndex + 1) + "]",
                    n, selected, context, displayName: componentName,
                    describePresentation: administrator ? (Func<string>)(() => DescribePresentation(n, projection, rows, componentName)) : null);
                foreach (var r in linked)
                    node.Add(new MembershipHierarchyNode("Comparison row " + (r.OriginalRowIndex + 1) + ": " +
                        r.MembershipStatus + " | " + r.DefinitionStatusText, row: r));
                nodes.Add(n.OccurrenceIndex, node);
            }
            var groups = new Dictionary<string, MembershipHierarchyNode>();
            foreach (var n in projection.Occurrences.Where(n => included.Contains(n.OccurrenceIndex)))
            {
                token.ThrowIfCancellationRequested(); int p;
                if (parent.TryGetValue(n.OccurrenceIndex, out p) && nodes.ContainsKey(p))
                {
                    if (administrator)
                    {
                        var kind = n.Placement == HierarchyPlacement.VerifiedInclusionRoot ? HierarchyRelationshipKind.InclusionRoot : HierarchyRelationshipKind.StructuralParent;
                        var childLabel = ChildCategory(n.Identity);
                        var key = "parent:" + p + ":" + kind + ":" + childLabel;
                        Group(groups, key, nodes[p], childLabel, "Linked raw occurrences; relationship evidence: " + kind +
                            ". InclusionRoot is inclusion only, not structural ownership. Counts include navigation context where shown.").Add(nodes[n.OccurrenceIndex]);
                        continue;
                    }
                    var label = n.Placement == HierarchyPlacement.VerifiedInclusionRoot
                        ? "Inclusion-root links — not structural ownership" : "Verified structural parent links";
                    Group(groups, "parent:" + p + ":" + label, nodes[p], label).Add(nodes[n.OccurrenceIndex]);
                }
                else
                {
                    if (administrator)
                    {
                        var familyLabel = Family(n.Identity);
                        Group(groups, "family:" + n.Side + ":" + familyLabel, sides[n.Side], familyLabel,
                            "Navigation category only. Count represents direct child raw occurrences, including context; not descendant totals. These categories are not verified top-level Power Apps objects or comparison rows.")
                            .Add(nodes[n.OccurrenceIndex]);
                        continue;
                    }
                    var label = n.HasDifferentParentAssociations ? "Multiple verified associations — no single display parent" : "Parent Not Established";
                    var placement = Group(groups, "side:" + n.Side + ":" + label, sides[n.Side], label);
                    var category = (n.Identity.SemanticKind == ComponentSemanticKinds.Table ? "Tables" :
                        n.Identity.SemanticKind == ComponentSemanticKinds.AppModule ? "Model-driven Apps" : "Other Components") +
                        (n.HasDifferentParentAssociations ? " — multiple associations retained" : " — placement not established");
                    Group(groups, "category:" + n.Side + ":" + label + ":" + category, placement, category).Add(nodes[n.OccurrenceIndex]);
                }
            }
            if (administrator) foreach (var group in groups.Values) group.LabelCount();
            var orphanRows = projection.ComparisonRows.Where(l => visible.Contains(l.RowIndex) &&
                l.SourceOccurrences.Count + l.TargetOccurrences.Count == 0).ToList();
            if (orphanRows.Count > 0)
            {
                var group = new MembershipHierarchyNode("Comparison rows without retained raw mapping"); root.Add(group);
                foreach (var l in orphanRows)
                    group.Add(new MembershipHierarchyNode("Comparison row " + (l.RowIndex + 1) + ": " + rows[l.RowIndex].ComponentText +
                        " — " + rows[l.RowIndex].MembershipStatus + " | " + rows[l.RowIndex].DefinitionStatusText, row: rows[l.RowIndex]));
            }
            if (projection.Diagnostics.Count > 0)
            {
                var diagnostics = new MembershipHierarchyNode("Hierarchy diagnostics / retrieval limitations"); root.Add(diagnostics);
                foreach (var diagnostic in projection.Diagnostics) diagnostics.Add(new MembershipHierarchyNode(diagnostic));
            }
            return new MembershipHierarchyView(root, visible.Count, included.Count);
        }

        private static MembershipHierarchyNode Group(Dictionary<string, MembershipHierarchyNode> groups, string key,
            MembershipHierarchyNode parent, string label, string details = null)
        {
            MembershipHierarchyNode result;
            if (!groups.TryGetValue(key, out result)) { result = new MembershipHierarchyNode(label, details: details); groups.Add(key, result); parent.Add(result); }
            return result;
        }

        private static string Family(ComponentIdentity identity)
        {
            switch (identity.SemanticKind)
            {
                case ComponentSemanticKinds.Table: return "Tables";
                case ComponentSemanticKinds.AppModule: return "Model-driven Apps";
                case ComponentSemanticKinds.EnvironmentVariableDefinition: return "Environment Variables";
                case ComponentSemanticKinds.GlobalChoice: return "Choices";
                case ComponentSemanticKinds.WebResource: return "Web Resources";
                default: return "Other Components";
            }
        }
        private static string ChildCategory(ComponentIdentity identity)
        {
            switch (identity.SemanticKind)
            {
                case ComponentSemanticKinds.AppSetting: return "AppSettings";
                case ComponentSemanticKinds.Column: return "Columns";
                case ComponentSemanticKinds.Relationship: return "Relationships";
                case ComponentSemanticKinds.SystemForm: return "Forms";
                case ComponentSemanticKinds.SavedQuery: return "Views";
                case ComponentSemanticKinds.SavedQueryVisualization: return "Charts";
                default: return "Components";
            }
        }

        private static string AdministratorName(ComponentHierarchyOccurrence n, IReadOnlyList<MembershipGridRow> rows, string parentTable)
        {
            var side = n.ComparisonRowIndices.Select(i => n.Side == DisplayEnvironmentSide.Source ? rows[i].SourceDisplayEvidence : rows[i].TargetDisplayEvidence)
                .FirstOrDefault(s => s?.HasLocalObservation == true && s.EnvironmentSide == n.Side &&
                    s.SolutionComponentId == n.Record.SolutionComponentId && s.LocalObjectId == n.Record.ObjectId && s.RawComponentType == n.Record.ComponentType);
            var refined = MembershipGridRow.RefinedComponentName(side, n.Identity);
            if (refined != null)
            {
                var qualified = MembershipGridRow.QualifiedColumnIdentity(n.Identity);
                return MembershipGridRow.EscapeForGrid(qualified != null && string.Equals(parentTable,
                    qualified.Substring(0, qualified.IndexOf('.')), StringComparison.OrdinalIgnoreCase)
                    ? qualified.Substring(qualified.IndexOf('.') + 1) : refined.Value);
            }
            return Name(n, rows);
        }

        private static string DescribePresentation(ComponentHierarchyOccurrence n, ComponentHierarchyProjection projection,
            IReadOnlyList<MembershipGridRow> rows, string name)
        {
            var text = new StringBuilder();
            var environment = n.Side == DisplayEnvironmentSide.Source ? projection.Presentation.Source.Diagnostics.EnvironmentName : projection.Presentation.Target.Diagnostics.EnvironmentName;
            text.AppendLine("Component name: " + name).AppendLine(n.Side + ": " + environment);
            text.AppendLine("Component semantic type: " + (n.Identity.SemanticKind ?? "[Not established]"));
            text.AppendLine("Selected solution: " + n.Snapshot.SolutionUniqueName);
            text.AppendLine("Complete portable identity: " + (n.Identity.ComparisonKey ?? "[Not established]"));
            text.AppendLine("Parent placement status: " + n.Placement);
            text.AppendLine("Original raw occurrence reference: " + n.Side + " snapshot index " + n.OriginalRawIndex);
            foreach (var index in n.ComparisonRowIndices)
                text.AppendLine("Original comparison row " + index + ": Membership status = " + rows[index].MembershipStatus +
                    "; Definition status = " + rows[index].DefinitionStatusText);
            foreach (var association in n.Associations.Where(a => a.Status == HierarchyAssociationStatus.Verified && a.ParentOccurrenceIndex.HasValue))
            {
                var parent = projection.Occurrences[association.ParentOccurrenceIndex.Value];
                text.AppendLine("Verified parent component: " + Name(parent, rows));
                text.AppendLine("Verified parent identity: " + (parent.Identity.ComparisonKey ?? "[Portable identity not established]"));
                text.AppendLine("Parent relationship evidence type: " + association.Kind);
            }
            string qualifier, definition;
            if (MembershipGridRow.AppSettingIdentityNames(n.Identity, out qualifier, out definition))
            {
                text.AppendLine("Setting definition identity (from verified membership key): " + definition);
                text.AppendLine("Application identity qualifier (from verified membership key): " + qualifier);
                text.AppendLine("Identity labels do not independently establish parent association or structural ownership.");
            }
            return text.ToString();
        }

        private static string Name(ComponentHierarchyOccurrence n, IReadOnlyList<MembershipGridRow> rows)
        {
            var observed = n.ComparisonRowIndices.Select(i => n.Side == DisplayEnvironmentSide.Source ?
                rows[i].SourceDisplayEvidence : rows[i].TargetDisplayEvidence).Where(s => s != null && s.HasLocalObservation &&
                s.EnvironmentSide == n.Side && s.SolutionComponentId == n.Record.SolutionComponentId &&
                s.LocalObjectId == n.Record.ObjectId && s.RawComponentType == n.Record.ComponentType).ToList();
            foreach (var selector in new Func<ComponentSideDisplayEvidence, ComponentDisplayValue>[] { s => s.FriendlyName, s => s.InternalName })
            {
                var names = observed.Select(selector).Where(v => v?.Availability == DisplayAvailability.Available &&
                    !string.IsNullOrWhiteSpace(v.Value)).Select(v => v.Value).Distinct(StringComparer.Ordinal).ToList();
                if (names.Count == 1) return MembershipGridRow.EscapeForGrid(names[0]);
                if (names.Count > 1) return "Component Type " + n.Record.ComponentType + " — Conflicting retained names";
            }
            return "Component Type " + n.Record.ComponentType + " — Name unavailable";
        }

        internal static string Describe(MembershipHierarchyNode node)
        {
            var text = new StringBuilder();
            if (node.NavigationContext) text.AppendLine("Navigation context only; not counted as a filtered comparison result.");
            var presentationDetails = node.NavigationDetails;
            if (presentationDetails != null) text.AppendLine(presentationDetails);
            var n = node.Occurrence;
            if (n == null) return text.AppendLine(node.Text).ToString();
            text.AppendLine("Environment side: " + n.Side).AppendLine("Original raw occurrence index: " + n.OriginalRawIndex);
            text.AppendLine("Hierarchy placement: " + n.Placement);
            text.AppendLine("SolutionComponentId (local audit): " + n.Record.SolutionComponentId);
            text.AppendLine("ComponentType: " + n.Record.ComponentType).AppendLine("ObjectId (local audit): " + n.Record.ObjectId);
            text.AppendLine("RootSolutionComponentId (local audit): " + n.Record.RootSolutionComponentId);
            text.AppendLine("RootComponentBehavior: " + n.Record.RootComponentBehavior).AppendLine("IsMetadata: " + n.Record.IsMetadata);
            text.AppendLine("Raw identity resolution (retained audit): " + n.Identity.Status);
            text.AppendLine("Raw portable identity (retained audit): " + n.Identity.ComparisonKey);
            text.AppendLine("Raw diagnostic: " + n.Identity.Diagnostic);
            foreach (var item in n.Identity.DiagnosticEvidence) text.AppendLine(item);
            if (n.Placement == HierarchyPlacement.ParentNotEstablished)
                text.AppendLine("Parent relationship not established from available evidence. This is not top-level proof.");
            if (n.HasDifferentParentAssociations) text.AppendLine("Multiple verified parent associations retained; no single parent selected for display.");
            foreach (var a in n.Associations)
            {
                text.AppendLine().AppendLine(a.Kind + ": " + a.Status);
                text.AppendLine("Parent occurrence index: " + a.ParentOccurrenceIndex).AppendLine("Referenced local ID: " + a.ReferencedLocalId);
                text.AppendLine("Provenance: " + a.Provenance).AppendLine(a.Diagnostic);
            }
            text.AppendLine().AppendLine("Linked authoritative comparison row indices: " + string.Join(", ", n.ComparisonRowIndices));
            if (n.ComparisonRowIndices.Count == 0) text.AppendLine("Comparison mapping unavailable; no classification invented.");
            if (n.ComparisonRowIndices.Count > 1) text.AppendLine("Expand this occurrence and select a linked comparison row for its details.");
            return text.ToString();
        }
    }
}
