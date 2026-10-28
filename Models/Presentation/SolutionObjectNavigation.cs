using System;
using System.Collections.Generic;
using System.Linq;
using D365SolutionComparer.Models.Membership;

namespace D365SolutionComparer.Models.Presentation
{
    // Display-only object navigation. Never creates an identity, absence inference or comparison row.
    internal sealed class SolutionObjectNavigation
    {
        internal sealed class LocalObject
        {
            internal string Key, Category, Type, Name, LogicalName, ParentName;
            internal DisplayEnvironmentSide Side;
            internal Guid Id;
            internal readonly List<int> Rows = new List<int>();
            internal readonly List<int> Occurrences = new List<int>();
            internal bool NameConflict, Child;
        }
        internal sealed class Item
        {
            internal string Key, Category, Type;
            internal IReadOnlyList<LocalObject> Locals;
            internal IReadOnlyList<int> Rows;
            internal readonly List<Item> Children = new List<Item>();
            internal bool Has(DisplayEnvironmentSide side) => Locals.Any(l => l.Side == side);
            internal string Value(bool logical)
            {
                var source = Locals.Where(l => l.Side == DisplayEnvironmentSide.Source).Select(l => logical ? l.LogicalName : l.Name).Distinct().ToArray();
                var target = Locals.Where(l => l.Side == DisplayEnvironmentSide.Target).Select(l => logical ? l.LogicalName : l.Name).Distinct().ToArray();
                string left = source.Length == 1 ? source[0] : null, right = target.Length == 1 ? target[0] : null;
                if (logical && (Category == "Dashboards" || Category == "Reports" || Category == "Team template" || Category == "Secured Masking Rule" || Category == "Processes" || Category == "Email templates"))
                    return left != null && right != null && left == right ? left :
                        "Source: " + (left ?? "[Not available]") + " · Target: " + (right ?? "[Not available]");
                if (left != null && (target.Length == 0 || left == right)) return left;
                if (right != null && source.Length == 0) return right;
                return "Source: " + (left ?? "[Not available]") + " · Target: " + (right ?? "[Not available]");
            }
            internal string Presence(DisplayEnvironmentSide side, SolutionObjectsExplorerProjection owner)
            {
                if (Has(side)) return "Present";
                bool absent = Rows.Count > 0 && Rows.All(i => owner.Rows[i].Original.AuthoritativeRow.Presence ==
                    (side == DisplayEnvironmentSide.Source ? MembershipPresence.OnlyInTarget : MembershipPresence.OnlyInSource));
                return absent ? "Absent from solution" : "Not established";
            }
            internal string Status(SolutionObjectsExplorerProjection owner)
            {
                if (Locals.Any(l => l.NameConflict)) return "Needs Review";
                var statuses = Rows.Select(i => SolutionObjectsExplorerControl.SimplifiedStatus(owner.Rows[i])).Distinct().ToArray();
                return statuses.Length == 1 ? statuses[0] : "Needs Review";
            }
        }
        internal SolutionObjectsExplorerProjection Owner { get; }
        private readonly IReadOnlyList<LocalObject> locals;
        internal bool MappingLimited { get; }

        internal SolutionObjectNavigation(SolutionObjectsExplorerProjection owner)
        {
            Owner = owner ?? throw new ArgumentNullException(nameof(owner));
            var groups = new Dictionary<string, LocalObject>(StringComparer.Ordinal);
            foreach (var occurrence in owner.Occurrences)
            {
                var identity = occurrence.Original.Identity; var record = identity.Record;
                if (!record.ObjectId.HasValue || record.ObjectId == Guid.Empty || occurrence.Original.ComparisonRowIndices.Count == 0) continue;
                var candidates = occurrence.Original.ComparisonRowIndices.Select(i => owner.Rows[i]).ToList();
                var row = candidates[0]; var side = occurrence.Original.Side;
                var display = side == DisplayEnvironmentSide.Source ? row.SourceDisplay : row.TargetDisplay;
                var maker = display?.MakerObject;
                if (maker != null && !maker.IsFor(identity)) maker = null;
                bool cloudFlow = RetainedCloudFlow(identity, display, side);
                string category = maker != null ? maker.Category : cloudFlow ? "Cloud flows" : Category(identity);
                bool backing = display != null && (display.BackingCorrelationStatus == DisplayBackingCorrelationStatus.Unique ||
                    display.BackingCorrelationStatus == DisplayBackingCorrelationStatus.VerifiedMetadata || display.BackingCorrelationStatus == DisplayBackingCorrelationStatus.VerifiedDefinition);
                if (category == null || identity.Status == IdentityResolutionStatus.Ambiguous && maker == null && record.ComponentType != 10139 ||
                    identity.Status != IdentityResolutionStatus.Resolved && !backing) continue;
                var names = candidates.Select(r => r.PrimaryName(side, true)).Where(n => !string.IsNullOrWhiteSpace(n)).Distinct(StringComparer.Ordinal).ToArray();
                string name = names.Length == 1 ? names[0] : names.Length > 1 ? "[Conflicting retained names]" : null;
                if (string.IsNullOrWhiteSpace(name)) continue; // No invented friendly object from a raw type/diagnostic string.
                Guid localId = maker?.CanonicalId ?? record.ObjectId.Value;
                string key = side + ":" + identity.SemanticKind + ":" + localId.ToString("D");
                if (!groups.TryGetValue(key, out var item))
                {
                    item = new LocalObject { Key = key, Id = localId, Side = side, Category = category,
                        Type = maker != null ? maker.Type : cloudFlow ? "Cloud flow" :
                            record.ComponentType == 91 ? "Plug-in assembly" : record.ComponentType == 92 ? "Plug-in step" :
                            record.ComponentType == 511 ? "Team Template" : record.ComponentType == 59 ? "Chart" :
                            record.ComponentType == 10139 ? "MetadataForArchival" : row.PrimaryType,
                        Name = name, LogicalName = row.PrimaryName(side, false), Child = record.ComponentType == 59 };
                    groups.Add(key, item);
                }
                else if (!StringComparer.Ordinal.Equals(item.Name, name)) item.NameConflict = true;
                if (names.Length > 1) item.NameConflict = true;
                item.Rows.AddRange(occurrence.Original.ComparisonRowIndices); item.Occurrences.Add(occurrence.Original.OccurrenceIndex);
                if (record.ComponentType == 59 && display?.ParentDisplayContext?.Availability == DisplayAvailability.Available &&
                    display.ParentDisplayContext.DisplayEvidenceSource == "NormalMembershipDiagnostic.savedqueryvisualization.primaryentitytypecode")
                {
                    if (item.ParentName != null && !StringComparer.OrdinalIgnoreCase.Equals(item.ParentName, display.ParentDisplayContext.Value)) item.NameConflict = true;
                    item.ParentName = display.ParentDisplayContext.Value;
                }
            }
            foreach (var group in groups.Values.Where(l => l.NameConflict)) { group.Name = "[Conflicting retained names]"; group.LogicalName = null; }
            locals = groups.Values.ToList().AsReadOnly();
            var mapped = new HashSet<int>(locals.SelectMany(l => l.Occurrences));
            MappingLimited = owner.Occurrences.Any(o => !mapped.Contains(o.Original.OccurrenceIndex)) ||
                locals.Any(l => l.NameConflict || l.Child && string.IsNullOrWhiteSpace(l.ParentName));
        }

        // This consumes resolver-retained category and detached, side-owned name evidence only.
        // It never reads a cache/service or establishes a portable identity or counterpart.
        private static bool RetainedCloudFlow(ComponentIdentity identity, ComponentSideDisplayEvidence display,
            DisplayEnvironmentSide side) => identity.Record.ComponentType == 29 &&
            identity.SemanticKind == ComponentSemanticKinds.Process && identity.Status == IdentityResolutionStatus.Resolved &&
            identity.WorkflowCategory == 5 && display != null && display.EnvironmentSide == side &&
            display.SolutionComponentId == identity.Record.SolutionComponentId && display.LocalObjectId == identity.Record.ObjectId &&
            display.RawComponentType == 29 &&
            (display.BackingCorrelationStatus == DisplayBackingCorrelationStatus.Unique ||
             display.BackingCorrelationStatus == DisplayBackingCorrelationStatus.VerifiedDefinition) &&
            display.FriendlyName?.Availability == DisplayAvailability.Available &&
            display.FriendlyName.DisplayEvidenceSource == "workflow.name" && !string.IsNullOrWhiteSpace(display.FriendlyName.Value);

        private static string Category(ComponentIdentity identity)
        {
            switch (identity.Record.ComponentType)
            {
                case 1: return "Tables";
                case 9: return identity.Status == IdentityResolutionStatus.Resolved ? "Choices" : null;
                case 31: return "Reports";
                case 36: return "Email templates";
                case 61: return "Web resources";
                case 62: return "Site maps";
                case 80: return "Apps";
                case 10139: return "MetadataForArchival";
                case 91: return "Plug-in assemblies";
                case 92: return "Plug-in steps";
                case 380: return "Environment variables";
                case 59: return "Charts"; // Child navigation only, never a top-level category.
                case 511: return identity.SemanticKind == ComponentSemanticKinds.TeamTemplate ? "Team template" : null;
                default: return identity.SemanticKind == ComponentSemanticKinds.ConnectionReference ? "Connection references" : null;
            }
        }

        internal static bool InScope(Item item, ObjectsWorkspaceScope scope) => scope == ObjectsWorkspaceScope.Compare ||
            item.Has(scope == ObjectsWorkspaceScope.Source ? DisplayEnvironmentSide.Source : DisplayEnvironmentSide.Target);

        internal IReadOnlyList<Item> Items(ObjectsWorkspaceScope scope)
        {
            var local = locals.ToList();
            var parent = Enumerable.Range(0, local.Count).ToArray();
            Func<int, int> root = null; root = i => parent[i] == i ? i : parent[i] = root(parent[i]);
            // Only an existing authoritative, resolved shared comparison may join two environments.
            // Equal display names, GUIDs and diagnostic strings do not introduce portable identity.
            {
                // Retain approved opposite-side observations for the Source/Target columns in every scope.
                var byRow = new Dictionary<int, List<int>>();
                for (int i = 0; i < local.Count; i++) foreach (int row in local[i].Rows.Distinct())
                { if (!byRow.TryGetValue(row, out var ids)) byRow[row] = ids = new List<int>(); ids.Add(i); }
                foreach (var link in byRow)
                {
                    var row = Owner.Rows[link.Key].Original.AuthoritativeRow;
                    if (row.Presence != MembershipPresence.PresentInBoth || row.SourceIdentity?.Status != IdentityResolutionStatus.Resolved ||
                        row.TargetIdentity?.Status != IdentityResolutionStatus.Resolved || link.Value.Count != 2) continue;
                    int a = link.Value[0], b = link.Value[1];
                    if (local[a].Side != local[b].Side && local[a].Category == local[b].Category && local[a].Child == local[b].Child)
                        parent[root(b)] = root(a);
                }
            }
            var items = Enumerable.Range(0, local.Count).GroupBy(root).Select(g =>
            {
                var members = g.Select(i => local[i]).ToList().AsReadOnly();
                string type = (members[0].Category == "Dashboards" || members[0].Category == "Processes") && members.Select(l => l.Type).Distinct().Count() > 1 ?
                    string.Join(" · ", members.Select(l => l.Side + ": " + l.Type).Distinct()) : members[0].Type;
                return new Item { Key = members[0].Key, Category = members[0].Category, Type = type, Locals = members,
                    Rows = members.SelectMany(l => l.Rows).Distinct().OrderBy(i => i).ToList().AsReadOnly() };
            }).ToList();
            foreach (var chart in items.Where(i => i.Category == "Charts" && InScope(i, scope)))
            {
                var parents = chart.Locals.Select(l => items.Where(i => i.Category == "Tables" && i.Locals.Any(p => p.Side == l.Side &&
                    !string.IsNullOrWhiteSpace(p.LogicalName) && StringComparer.OrdinalIgnoreCase.Equals(p.LogicalName, l.ParentName))).ToList()).ToList();
                if (parents.All(p => p.Count == 1) && parents.SelectMany(p => p).Distinct().Count() == 1 && !chart.Locals.Any(l => l.NameConflict))
                    parents[0][0].Children.Add(chart);
            }
            return items.Where(i => i.Category != "Charts" && InScope(i, scope)).ToList().AsReadOnly();
        }
    }
}
