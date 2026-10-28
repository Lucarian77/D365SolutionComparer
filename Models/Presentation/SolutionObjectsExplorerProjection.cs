using System;
using System.Collections.Generic;
using System.Linq;
using D365SolutionComparer.Models.Membership;

namespace D365SolutionComparer.Models.Presentation
{
    internal enum ObjectsWorkspaceScope { Target, Source, Compare }

    /// <summary>Detached navigation only. Original rows and raw occurrences remain authoritative.</summary>
    internal sealed class SolutionObjectsExplorerProjection
    {
        internal SolutionObjectsExplorerProjection(MembershipComparisonPresentation owner,
            IEnumerable<SolutionObjectsExplorerRow> rows, IEnumerable<SolutionObjectsExplorerOccurrence> occurrences,
            IEnumerable<SolutionObjectsExplorerCategory> categories, IEnumerable<string> diagnostics)
        {
            Owner = owner;
            Rows = rows.ToList().AsReadOnly(); Occurrences = occurrences.ToList().AsReadOnly();
            Categories = categories.ToList().AsReadOnly(); Diagnostics = diagnostics.ToList().AsReadOnly();
            SourceOccurrenceCount = Occurrences.Count(o => o.Original.Side == DisplayEnvironmentSide.Source);
            TargetOccurrenceCount = Occurrences.Count - SourceOccurrenceCount;
        }
        internal MembershipComparisonPresentation Owner { get; }
        internal IReadOnlyList<SolutionObjectsExplorerRow> Rows { get; }
        internal IReadOnlyList<SolutionObjectsExplorerOccurrence> Occurrences { get; }
        internal IReadOnlyList<SolutionObjectsExplorerCategory> Categories { get; }
        internal IReadOnlyList<string> Diagnostics { get; }
        internal int SourceOccurrenceCount { get; }
        internal int TargetOccurrenceCount { get; }
        internal SolutionObjectsExplorerCategory Category(string key) => Categories.Single(c => c.Key == key);
        internal IReadOnlyList<SolutionObjectsExplorerCategory> WorkspaceCategories(ObjectsWorkspaceScope scope,
            IEnumerable<int> filteredRows, bool all)
        {
            var allowed = new HashSet<int>(filteredRows);
            var local = Occurrences.Where(o => (scope == ObjectsWorkspaceScope.Compare || o.Original.Side ==
                (scope == ObjectsWorkspaceScope.Target ? DisplayEnvironmentSide.Target : DisplayEnvironmentSide.Source)) &&
                (all || o.Original.ComparisonRowIndices.Any(allowed.Contains))).ToList();
            var result = new List<SolutionObjectsExplorerCategory>();
            Action<string, string, IEnumerable<SolutionObjectsExplorerOccurrence>, IEnumerable<int>> add = (key, name, items, extra) =>
            {
                var entries = items.ToList(); result.Add(new SolutionObjectsExplorerCategory(key, name,
                    entries.SelectMany(o => o.Original.ComparisonRowIndices).Where(allowed.Contains).Concat(extra),
                    entries.Where(o => o.Original.Side == DisplayEnvironmentSide.Source).Select(o => o.Original.OccurrenceIndex),
                    entries.Where(o => o.Original.Side == DisplayEnvironmentSide.Target).Select(o => o.Original.OccurrenceIndex)));
            };
            add("workspace-all", "All", local.Where(o => o.PrimaryCategory != null), Enumerable.Empty<int>());
            foreach (var group in local.Where(o => o.PrimaryCategory != null).GroupBy(o => o.PrimaryCategory).OrderBy(g => g.Key, StringComparer.Ordinal))
                add("workspace-" + group.Key, group.Key, group, Enumerable.Empty<int>());
            add("workspace-technical", "Technical Membership / unclassified", local.Where(o => o.PrimaryCategory == null), Enumerable.Empty<int>());
            return result.AsReadOnly();
        }

    }

    internal sealed class SolutionObjectsExplorerCategory
    {
        internal SolutionObjectsExplorerCategory(string key, string name, IEnumerable<int> rows,
            IEnumerable<int> source, IEnumerable<int> target)
        {
            Key = key; Name = name;
            RowIndices = rows.Distinct().OrderBy(i => i).ToList().AsReadOnly();
            SourceOccurrences = source.Distinct().OrderBy(i => i).ToList().AsReadOnly();
            TargetOccurrences = target.Distinct().OrderBy(i => i).ToList().AsReadOnly();
        }
        internal string Key { get; }
        internal string Name { get; }
        internal IReadOnlyList<int> RowIndices { get; }
        internal IReadOnlyList<int> SourceOccurrences { get; }
        internal IReadOnlyList<int> TargetOccurrences { get; }
        internal string CountText => RowIndices.Count + " comparison rows; Source " + SourceOccurrences.Count +
            " / Target " + TargetOccurrences.Count + " raw occurrences";
    }

    internal sealed class SolutionObjectsExplorerOccurrence
    {
        internal SolutionObjectsExplorerOccurrence(ComponentHierarchyOccurrence original, string category,
            string label, bool objectCandidate, IEnumerable<SolutionObjectsExplorerLink> children, string primaryCategory = null)
        {
            Original = original; Category = category; Label = label; ObjectCandidate = objectCandidate;
            PrimaryCategory = primaryCategory; Children = children.ToList().AsReadOnly();
        }
        internal ComponentHierarchyOccurrence Original { get; }
        internal string Category { get; }
        internal string Label { get; }
        internal bool ObjectCandidate { get; }
        internal string PrimaryCategory { get; }
        internal IReadOnlyList<SolutionObjectsExplorerLink> Children { get; }
        internal bool PlacementNotEstablished => Original.Placement == HierarchyPlacement.ParentNotEstablished ||
            Original.HasDifferentParentAssociations;
    }

    internal sealed class SolutionObjectsExplorerLink
    {
        internal SolutionObjectsExplorerLink(int child, ComponentHierarchyAssociation original, string group)
        { ChildOccurrenceIndex = child; OriginalAssociation = original; Group = group; }
        internal int ChildOccurrenceIndex { get; }
        internal ComponentHierarchyAssociation OriginalAssociation { get; }
        internal string Group { get; }
    }
}
