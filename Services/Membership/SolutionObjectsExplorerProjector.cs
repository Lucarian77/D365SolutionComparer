using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using D365SolutionComparer.Models.Membership;
using D365SolutionComparer.Models.Presentation;

namespace D365SolutionComparer.Services.Membership
{
    /// <summary>Pure local projection of an existing H1 result. No SDK/cache/diagnostic parsing.</summary>
    internal static class SolutionObjectsExplorerProjector
    {
        internal static SolutionObjectsExplorerProjection Project(MembershipComparisonPresentation presentation,
            CompletedMembershipComparison completed, ComponentHierarchyProjection hierarchy,
            CancellationToken token = default(CancellationToken))
        {
            token.ThrowIfCancellationRequested();
            if (presentation == null) throw new ArgumentNullException(nameof(presentation));
            if (completed != null && !ReferenceEquals(completed.Presentation, presentation))
                throw new InvalidOperationException("Explorer display owner mismatch.");
            if (hierarchy == null || hierarchy.Outcome != HierarchyProjectionOutcome.Completed ||
                !ReferenceEquals(hierarchy.Presentation, presentation) || hierarchy.ComparisonRows.Count != presentation.Rows.Count)
                throw new InvalidOperationException("Explorer requires the completed hierarchy for this comparison.");
            var grid = MembershipGridRow.CreateRows(presentation, completed, refinedPresentation: true);
            var rows = new List<SolutionObjectsExplorerRow>();
            for (int index = 0; index < grid.Count; index++)
            {
                token.ThrowIfCancellationRequested(); var link = hierarchy.ComparisonRows[index];
                if (link.RowIndex != index || !ReferenceEquals(link.Row, presentation.Rows[index]))
                    throw new InvalidOperationException("Explorer row mapping mismatch.");
                rows.Add(new SolutionObjectsExplorerRow(grid[index], link.SourceOccurrences, link.TargetOccurrences));
            }
            var children = hierarchy.Occurrences.Select(o => new List<SolutionObjectsExplorerLink>()).ToArray();
            var categories = new string[children.Length];
            var labels = new string[children.Length]; var candidates = new bool[children.Length];
            for (int index = 0; index < hierarchy.Occurrences.Count; index++)
            {
                token.ThrowIfCancellationRequested(); var item = hierarchy.Occurrences[index];
                if (item.OccurrenceIndex != index || item.Snapshot != (item.Side == DisplayEnvironmentSide.Source ? presentation.Source.Snapshot : presentation.Target.Snapshot))
                    throw new InvalidOperationException("Explorer occurrence owner mismatch.");
                categories[index] = Category(item, out var candidate); candidates[index] = candidate;
                var mapped = item.ComparisonRowIndices.Select(i => rows[i]).ToList();
                var knownLabels = mapped.Select(r => r.Label(item.Side)).Distinct(StringComparer.Ordinal).ToList();
                labels[index] = knownLabels.Count == 1 ? knownLabels[0] : knownLabels.Count > 1
                    ? "[Conflicting retained labels]" : "Component Type " + item.Record.ComponentType + " [Name not retrieved]";
                if (item.Placement == HierarchyPlacement.ParentNotEstablished || item.HasDifferentParentAssociations) continue;
                foreach (var association in item.Associations.Where(a => a.Status == HierarchyAssociationStatus.Verified && a.ParentOccurrenceIndex.HasValue))
                {
                    int parent = association.ParentOccurrenceIndex.Value;
                    if (parent < 0 || parent >= children.Length || hierarchy.Occurrences[parent].Side != item.Side ||
                        !ReferenceEquals(hierarchy.Occurrences[parent].Snapshot, item.Snapshot))
                        throw new InvalidOperationException("Explorer parent evidence crosses its selected side.");
                    children[parent].Add(new SolutionObjectsExplorerLink(index, association, ChildGroup(item.Record.ComponentType)));
                }
            }
            var occurrences = hierarchy.Occurrences.Select((item, index) => new SolutionObjectsExplorerOccurrence(item,
                categories[index], labels[index], candidates[index], children[index])).ToList();
            var resultCategories = new List<SolutionObjectsExplorerCategory>();
            Action<string, string, IEnumerable<SolutionObjectsExplorerOccurrence>, IEnumerable<int>> add = (key, name, selected, extra) =>
            {
                var local = selected.ToList();
                resultCategories.Add(new SolutionObjectsExplorerCategory(key, name,
                    local.SelectMany(o => o.Original.ComparisonRowIndices).Concat(extra),
                    local.Where(o => o.Original.Side == DisplayEnvironmentSide.Source).Select(o => o.Original.OccurrenceIndex),
                    local.Where(o => o.Original.Side == DisplayEnvironmentSide.Target).Select(o => o.Original.OccurrenceIndex)));
            };
            add("objects", "All Objects (candidates)", occurrences.Where(o => o.ObjectCandidate), new int[0]);
            add("all", "All Components", occurrences, Enumerable.Range(0, rows.Count));
            foreach (var group in occurrences.GroupBy(o => o.Category).OrderBy(g => g.Key, StringComparer.Ordinal))
                add(group.Key, group.Key, group, new int[0]);
            add("review", "Review / placement not established", occurrences.Where(o => o.PlacementNotEstablished ||
                o.Original.Identity.Status != IdentityResolutionStatus.Resolved ||
                o.Original.ComparisonRowIndices.Any(i => rows[i].Original.AuthoritativeRow.Presence == MembershipPresence.Indeterminate)),
                rows.Where(r => r.Original.AuthoritativeRow.Presence == MembershipPresence.Indeterminate).Select(r => r.OriginalRowIndex));
            add("unmapped", "Raw / unmapped evidence", occurrences.Where(o => o.Original.ComparisonRowIndices.Count == 0),
                rows.Where(r => r.SourceOccurrences.Count + r.TargetOccurrences.Count == 0).Select(r => r.OriginalRowIndex));
            token.ThrowIfCancellationRequested();
            return new SolutionObjectsExplorerProjection(presentation, rows, occurrences, resultCategories, hierarchy.Diagnostics);
        }

        private static string Category(ComponentHierarchyOccurrence item, out bool objectCandidate)
        {
            objectCandidate = true;
            switch (item.Record.ComponentType)
            {
                case 1: return "Tables";
                case 80: return "Model-driven Apps";
                case 9: return "Choices (supported subset / review)";
                case 61: return "Web Resources";
                case 380: return "Environment Variables";
                case 29: return "Processes / Workflows";
                case 62: return "Site Maps";
                case 31: return "Reports";
                case 36: return "Email Templates";
                case 74: return Registered(item, "maskingrule") ? "Masking Rules (scope unverified)" : "Other / unclassified components";
                case 300: return "Canvas Apps / Libraries (kind unclassified)";
                case 511: return "Team Templates";
                case 10072: return Registered(item, "appelement") ? "AppElements (type unverified)" : "Other / unclassified components";
                case 10276: return Registered(item, "aiskillconfig") ? "AI Skill Config" : "Other / unclassified components";
                case 10266: case 10267: return Registered(item, "appaction") ? "App Actions (diagnostic family only)" : "Other / unclassified components";
                case 91: return "Plug-in Assemblies";
                case 92: objectCandidate = false; return "SDK Message Processing Steps";
                case 2: objectCandidate = false; return "Columns";
                case 10: objectCandidate = false; return "Relationships";
                case 26: objectCandidate = false; return "Views";
                case 59: objectCandidate = false; return "Charts";
                case 60: objectCandidate = false; return "Forms / Dashboards (subtype unclassified)";
                case 10075: objectCandidate = false; return "AppSettings";
                default:
                    return item.Identity.SemanticKind == ComponentSemanticKinds.ConnectionReference ? "Connection References" : "Other / unclassified components";
            }
        }
        private static bool Registered(ComponentHierarchyOccurrence item, string backing) =>
            StringComparer.OrdinalIgnoreCase.Equals(item.Identity.RegisteredDefinition?.PrimaryEntityName, backing);
        internal static string ChildGroup(int rawType)
        {
            switch (rawType)
            {
                case 2: return "Columns"; case 10: return "Relationships"; case 60: return "Forms";
                case 26: return "Views"; case 10075: return "AppSettings";
                default: return "Other Components";
            }
        }
    }
}
