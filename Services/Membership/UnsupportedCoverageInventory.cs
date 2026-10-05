#if DEBUG
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading;
using D365SolutionComparer.Models.Membership;

namespace D365SolutionComparer.Services.Membership
{
    // Observational labels from the existing response, not identity/equality properties.
    // Weak ownership preserves the public snapshot model and does not retain discarded records.
    internal static class InventoryFormattedLabels
    {
        private sealed class Label { internal string Value; }
        private static readonly ConditionalWeakTable<SolutionComponentRecord, Label> Labels =
            new ConditionalWeakTable<SolutionComponentRecord, Label>();

        internal static void Capture(SolutionComponentRecord record, string label)
        {
            if (!string.IsNullOrWhiteSpace(label)) Labels.GetValue(record, _ => new Label()).Value = label;
        }
        internal static string Get(SolutionComponentRecord record) =>
            Labels.TryGetValue(record, out var label) ? label.Value : null;
    }

    internal sealed class InventoryCheckpoint
    {
        internal InventoryCheckpoint(MembershipSnapshot snapshot, string displayName = null, string version = null)
        {
            Snapshot = snapshot ?? throw new ArgumentNullException(nameof(snapshot));
            DisplayName = displayName;
            Version = version;
        }
        internal MembershipSnapshot Snapshot { get; }
        internal string DisplayName { get; }
        internal string Version { get; }
    }

    // UI-thread owned. Retain the latest checkpoint for each solution, not all runs.
    // Environment-pair changes reset the session. Loaded solution lists are NOT inventories.
    internal sealed class UnsupportedInventorySession
    {
        private Guid? sourceEnvironment, targetEnvironment;
        private readonly Dictionary<string, InventoryCheckpoint> source =
            new Dictionary<string, InventoryCheckpoint>(StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<string, InventoryCheckpoint> target =
            new Dictionary<string, InventoryCheckpoint>(StringComparer.OrdinalIgnoreCase);

        internal void Record(MembershipSnapshot sourceSnapshot, MembershipSnapshot targetSnapshot,
            string sourceDisplayName, string sourceVersion, string targetDisplayName, string targetVersion)
        {
            if (sourceSnapshot == null || targetSnapshot == null ||
                sourceSnapshot != null && sourceEnvironment.HasValue &&
                sourceEnvironment != sourceSnapshot.Environment.OrganizationId ||
                targetSnapshot != null && targetEnvironment.HasValue &&
                targetEnvironment != targetSnapshot.Environment.OrganizationId)
            {
                source.Clear(); target.Clear(); sourceEnvironment = null; targetEnvironment = null;
            }
            if (sourceSnapshot != null) sourceEnvironment = sourceSnapshot.Environment.OrganizationId;
            if (targetSnapshot != null) targetEnvironment = targetSnapshot.Environment.OrganizationId;
            Record(source, sourceSnapshot, sourceDisplayName, sourceVersion);
            Record(target, targetSnapshot, targetDisplayName, targetVersion);
        }

        private static void Record(Dictionary<string, InventoryCheckpoint> checkpoints,
            MembershipSnapshot snapshot, string displayName, string version)
        {
            if (snapshot == null) return;
            if (!checkpoints.TryGetValue(snapshot.SolutionUniqueName, out var existing) ||
                snapshot.CapturedAt >= existing.Snapshot.CapturedAt)
                checkpoints[snapshot.SolutionUniqueName] = new InventoryCheckpoint(snapshot, displayName, version);
        }

        internal InventoryCheckpoint[] Source => source.Values.ToArray();
        internal InventoryCheckpoint[] Target => target.Values.ToArray();
    }

    internal sealed class InventoryCounts
    {
        internal InventoryCounts(IEnumerable<ComponentIdentity> components)
        {
            var rows = components.ToArray();
            Raw = rows.Length;
            Ids = new HashSet<Guid>(rows.Where(row => row.Record.ObjectId.HasValue &&
                row.Record.ObjectId.Value != Guid.Empty).Select(row => row.Record.ObjectId.Value));
            BlankIds = rows.Count(row => !row.Record.ObjectId.HasValue || row.Record.ObjectId.Value == Guid.Empty);
            Resolved = rows.Count(row => row.Status == IdentityResolutionStatus.Resolved);
            Unsupported = rows.Count(row => row.Status == IdentityResolutionStatus.Unsupported);
            Unresolved = rows.Count(row => row.Status == IdentityResolutionStatus.Unresolved);
            Ambiguous = rows.Count(row => row.Status == IdentityResolutionStatus.Ambiguous);
        }
        internal int Raw { get; }
        internal HashSet<Guid> Ids { get; }
        internal int BlankIds { get; }
        internal int Resolved { get; }
        internal int Unsupported { get; }
        internal int Unresolved { get; }
        internal int Ambiguous { get; }
        internal string Values => string.Join("\t", Raw, Ids.Count, BlankIds, Resolved, Unsupported, Unresolved, Ambiguous);
    }

    internal sealed class InventoryFamily
    {
        internal int Type { get; set; }
        internal string FormattedLabel { get; set; }
        internal string CatalogLabel { get; set; }
        internal string BackingEntity { get; set; }
        internal string ResolverSupport { get; set; }
        internal string EvidenceStatus { get; set; }
        internal bool ExistingMetadataEvidence { get; set; }
        internal InventoryCounts Source { get; set; }
        internal InventoryCounts Target { get; set; }
        internal int SourceSolutions { get; set; }
        internal int TargetSolutions { get; set; }
        internal int SharedSolutions { get; set; }
        internal int AffectedSolutions { get; set; }
        internal string NextAction { get; set; }
        internal bool Both => Source.Raw > 0 && Target.Raw > 0;
        internal bool KnownBacking => BackingEntity != "Unknown";
        internal int TotalRaw => Source.Raw + Target.Raw;
        // Counts are environment-scoped: GUID equality across organizations does not merge identities.
        internal int TotalDistinct => Source.Ids.Count + Target.Ids.Count;
    }

    /// <summary>Local, read-only diagnostic aggregation. Deliberately accepts no Dataverse service.</summary>
    internal sealed class UnsupportedCoverageInventory
    {
        internal IReadOnlyList<InventoryFamily> Families { get; private set; }
        internal IReadOnlyList<InventoryFamily> Candidates { get; private set; }
        internal string Text { get; private set; }

        internal static UnsupportedCoverageInventory Build(IEnumerable<InventoryCheckpoint> source,
            IEnumerable<InventoryCheckpoint> target, CancellationToken cancellationToken = default(CancellationToken))
        {
            cancellationToken.ThrowIfCancellationRequested();
            var allSource = Normalize(source);
            var allTarget = Normalize(target);
            var sourceComplete = allSource.Where(c => c.Snapshot.State == MembershipSnapshotState.Complete).ToArray();
            var targetComplete = allTarget.Where(c => c.Snapshot.State == MembershipSnapshotState.Complete).ToArray();
            var sourceRows = sourceComplete.SelectMany(c => c.Snapshot.Components).ToLookup(c => c.Record.ComponentType);
            var targetRows = targetComplete.SelectMany(c => c.Snapshot.Components).ToLookup(c => c.Record.ComponentType);
            var families = new List<InventoryFamily>();
            foreach (var type in sourceRows.Select(g => g.Key).Union(targetRows.Select(g => g.Key)).OrderBy(t => t))
            {
                cancellationToken.ThrowIfCancellationRequested();
                var rows = sourceRows[type].Concat(targetRows[type]).ToArray();
                var s = new InventoryCounts(sourceRows[type]); var t = new InventoryCounts(targetRows[type]);
                var sourceSolutions = Solutions(sourceComplete, type); var targetSolutions = Solutions(targetComplete, type);
                var backing = Join(rows.Select(row => row.RegisteredDefinition?.PrimaryEntityName)
                    .Concat(new[] { KnownBacking(type, rows) }));
                bool known = ComponentSemanticKinds.IsKnownBuiltInType(type) ||
                    rows.Any(row => row.RegisteredDefinition != null ||
                        !string.IsNullOrWhiteSpace(row.SemanticKind)) || backing != "Unknown";
                var family = new InventoryFamily
                {
                    Type = type, FormattedLabel = Join(rows.Select(row => InventoryFormattedLabels.Get(row.Record))),
                    CatalogLabel = Join(rows.Select(row => row.RegisteredDefinition?.Name)
                        .Concat(rows.Select(row => row.SemanticKind)).Concat(new[] { DiagnosticFamilyLabel(type) })),
                    BackingEntity = backing, ResolverSupport = ResolverSupport(type, rows),
                    Source = s, Target = t, SourceSolutions = sourceSolutions.Count, TargetSolutions = targetSolutions.Count,
                    SharedSolutions = sourceSolutions.Intersect(targetSolutions, StringComparer.OrdinalIgnoreCase).Count(),
                    AffectedSolutions = sourceSolutions.Union(targetSolutions, StringComparer.OrdinalIgnoreCase).Count(),
                    ExistingMetadataEvidence = rows.Any(row => row.RegisteredDefinition != null || row.DiagnosticEvidence.Count > 0) ||
                        new[] { 9, 14, 26, 31, 36, 59, 60, 62, 80, 91, 92, 300, 511 }.Contains(type),
                    EvidenceStatus = EvidenceStatus(s, t, known)
                };
                family.NextAction = new[] { 3, 11, 12 }.Contains(type) ? "System/internal family - defer" :
                    family.KnownBacking && family.ExistingMetadataEvidence ? "Investigate portable identity" :
                    family.TotalRaw <= 1 && !family.Both ? "Low priority" : "Collect backing-row evidence";
                families.Add(family);
            }
            // Transparent lexicographic ranking: frequency, solution breadth, both sides,
            // known backing, existing evidence, then raw type. This is NOT a portability verdict.
            var candidates = families.Where(f => f.Source.Unsupported + f.Target.Unsupported > 0)
                .OrderByDescending(f => f.TotalRaw).ThenByDescending(f => f.AffectedSolutions)
                .ThenByDescending(f => f.Both).ThenByDescending(f => f.KnownBacking)
                .ThenByDescending(f => f.ExistingMetadataEvidence).ThenBy(f => f.Type).ToList();
            var result = new UnsupportedCoverageInventory
            {
                Families = families.AsReadOnly(), Candidates = candidates.AsReadOnly()
            };
            result.Text = Format(result, allSource, allTarget, cancellationToken);
            return result;
        }

        private static InventoryCheckpoint[] Normalize(IEnumerable<InventoryCheckpoint> checkpoints)
        {
            var rows = (checkpoints ?? throw new ArgumentNullException(nameof(checkpoints))).ToArray();
            if (rows.Any(row => row == null)) throw new ArgumentException("Null checkpoints are not permitted.");
            if (rows.Select(row => row.Snapshot.Environment.OrganizationId).Distinct().Count() > 1)
                throw new ArgumentException("A side must contain snapshots from only one organization.");
            if (rows.GroupBy(row => row.Snapshot.SolutionUniqueName, StringComparer.OrdinalIgnoreCase).Any(g => g.Count() > 1))
                throw new ArgumentException("Use only one checkpoint per solution and side.");
            return rows.OrderBy(row => row.Snapshot.SolutionUniqueName, StringComparer.OrdinalIgnoreCase).ToArray();
        }

        private static HashSet<string> Solutions(IEnumerable<InventoryCheckpoint> checkpoints, int type) =>
            new HashSet<string>(checkpoints.Where(c => c.Snapshot.Components.Any(row => row.Record.ComponentType == type))
                .Select(c => c.Snapshot.SolutionUniqueName), StringComparer.OrdinalIgnoreCase);

        private static string EvidenceStatus(InventoryCounts s, InventoryCounts t, bool known)
        {
            int resolved = s.Resolved + t.Resolved;
            if (resolved > 0 && resolved < s.Raw + t.Raw) return "PartiallyResolved";
            if (s.Ambiguous + t.Ambiguous > 0) return "Ambiguous";
            if (s.Unresolved + t.Unresolved > 0) return "Unresolved";
            if (s.Unsupported + t.Unsupported > 0) return known ? "UnsupportedKnownFamily" : "UnsupportedUnknownFamily";
            return "Supported";
        }

        // Descriptive inventory only. This never controls resolution/coverage/absence decisions.
        private static string ResolverSupport(int type, ComponentIdentity[] rows)
        {
            if (type == 9 || type == 31 || type == 20 || type == 29) return "Partial (approved subsets only)";
            if (new[] { 1, 2, 10, 14, 26, 59, 60, 61, 62, 80, 91, 92, 380 }.Contains(type) ||
                rows.Any(row => row.SemanticKind == ComponentSemanticKinds.ConnectionReference ||
                    row.SemanticKind == ComponentSemanticKinds.AppSetting)) return "Supported";
            return ComponentSemanticKinds.IsKnownBuiltInType(type) ||
                rows.Any(row => row.RegisteredDefinition != null || row.SemanticKind == ComponentSemanticKinds.TeamTemplate)
                ? "Unsupported" : "Unknown";
        }

        private static string DiagnosticFamilyLabel(int type)
        {
            switch (type)
            {
                case 9: return "Global Choice (verified subset)";
                case 31: return "Signed Report (verified subset)";
                case 36: return "Email Template";
                case 300: return "Canvas App";
                case 511: return "Team Template (correlation must be verified)";
                case 90: return "Plug-in Type";
                default: return null;
            }
        }

        private static string KnownBacking(int type, ComponentIdentity[] rows)
        {
            switch (type)
            {
                case 1: return "EntityMetadata (API)";
                case 2: return "AttributeMetadata (API)";
                case 10: return "RelationshipMetadata (API)";
                case 9: return "OptionSetMetadata (API)";
                case 14: return "EntityKeyMetadata (API)";
                case 20: return "role";
                case 26: return "savedquery";
                case 29: return "workflow";
                case 31: return "report";
                case 36: return "template";
                case 59: return "savedqueryvisualization";
                case 60: return "systemform";
                case 61: return "webresource";
                case 62: return "sitemap";
                case 80: return "appmodule";
                case 90: return "plugintype"; // Already read as Type 92 parent evidence.
                case 91: return "pluginassembly";
                case 92: return "sdkmessageprocessingstep";
                case 300: return "canvasapp";
                case 380: return "environmentvariabledefinition";
                case 511: return rows.Any(row => row.SemanticKind == ComponentSemanticKinds.TeamTemplate) ? "teamtemplate" : null;
                default: return rows.Any(row => row.SemanticKind == ComponentSemanticKinds.ConnectionReference)
                    ? "connectionreference" : null;
            }
        }

        private static string Join(IEnumerable<string> values) =>
            string.Join("; ", values.Where(value => !string.IsNullOrWhiteSpace(value))
                .Distinct(StringComparer.Ordinal).OrderBy(value => value, StringComparer.OrdinalIgnoreCase)
                .ThenBy(value => value, StringComparer.Ordinal).DefaultIfEmpty("Unknown"));
        private static string Cell(object value) => (Convert.ToString(value, CultureInfo.InvariantCulture) ?? "Unknown")
            .Replace("\\", "\\\\").Replace("\t", "\\t").Replace("\r", "\\r").Replace("\n", "\\n");
        private static void Line(StringBuilder text, params object[] values) =>
            text.AppendLine(string.Join("\t", values.Select(Cell)));
        private static string Ids(IEnumerable<Guid> ids) => string.Join(",", ids.OrderBy(id => id).Select(id => id.ToString("D")));

        private static string Format(UnsupportedCoverageInventory inventory, InventoryCheckpoint[] source,
            InventoryCheckpoint[] target, CancellationToken token)
        {
            var text = new StringBuilder();
            text.AppendLine("UNSUPPORTED COMPONENT COVERAGE INVENTORY - DEBUG ONLY");
            text.AppendLine("Scope: latest captured checkpoints per solution in this session; counts use Complete snapshots only; NOT an environment-wide inventory.");
            text.AppendLine("Unavailable/absent checkpoints are listed but excluded from counts. Zero counts do not prove absence.");
            text.AppendLine("Raw ObjectId overlaps are audit observations only, NOT membership matches or portable identity evidence.");
            text.AppendLine("Distinct IDs are deduplicated within each environment; candidate totals sum the two environment counts.");
            text.AppendLine("Ranking is investigation priority only. No identity is approved by this report.");
            text.AppendLine("Dataverse requests: 0; WhoAmI: 0; Writes: 0");
            text.AppendLine("CHECKPOINTS");
            Line(text, "Side", "Environment", "SolutionUniqueName", "DisplayName", "Version", "CapturedAtUtc", "SnapshotState");
            foreach (var side in new[] { new { Name = "Source", Rows = source }, new { Name = "Target", Rows = target } })
                foreach (var c in side.Rows)
                    Line(text, side.Name, c.Snapshot.Environment.DisplayName, c.Snapshot.SolutionUniqueName,
                        c.DisplayName, c.Version, c.Snapshot.CapturedAt.UtcDateTime.ToString("O", CultureInfo.InvariantCulture), c.Snapshot.State);
            text.AppendLine("RAW COMPONENT TYPE SUMMARY");
            Line(text, "ComponentType", "FormattedLabel", "CatalogLabel", "ResolverCatalogSupport", "KnownBackingEntity", "EvidenceStatus",
                "SourceRaw", "SourceDistinctObjectIds", "SourceBlankObjectIds", "SourceResolved", "SourceUnsupported", "SourceUnresolved", "SourceAmbiguous",
                "TargetRaw", "TargetDistinctObjectIds", "TargetBlankObjectIds", "TargetResolved", "TargetUnsupported", "TargetUnresolved", "TargetAmbiguous",
                "SourceSolutions", "TargetSolutions", "SharedSolutions");
            foreach (var f in inventory.Families)
            {
                token.ThrowIfCancellationRequested();
                text.Append(string.Join("\t", new object[] { f.Type, f.FormattedLabel, f.CatalogLabel, f.ResolverSupport, f.BackingEntity, f.EvidenceStatus }.Select(Cell)));
                text.Append('\t').Append(f.Source.Values).Append('\t').Append(f.Target.Values);
                text.Append('\t').Append(f.SourceSolutions).Append('\t').Append(f.TargetSolutions).Append('\t').Append(f.SharedSolutions).AppendLine();
            }
            text.AppendLine("NON-RESOLVED COVERAGE BY SOLUTION");
            Line(text, "Side", "SolutionUniqueName", "DisplayName", "Version", "ComponentType", "FormattedLabel",
                "RawCount", "DistinctObjectIds", "BlankObjectIds", "Resolved", "Unsupported", "Unresolved", "Ambiguous");
            foreach (var side in new[] { new { Name = "Source", Rows = source }, new { Name = "Target", Rows = target } })
                foreach (var c in side.Rows.Where(row => row.Snapshot.State == MembershipSnapshotState.Complete))
                    foreach (var group in c.Snapshot.Components.GroupBy(row => row.Record.ComponentType).OrderBy(g => g.Key))
                    {
                        token.ThrowIfCancellationRequested();
                        if (group.All(row => row.Status == IdentityResolutionStatus.Resolved)) continue;
                        text.Append(string.Join("\t", new object[] { side.Name, c.Snapshot.SolutionUniqueName, c.DisplayName, c.Version,
                            group.Key, Join(group.Select(row => InventoryFormattedLabels.Get(row.Record))) }.Select(Cell)));
                        text.Append('\t').Append(new InventoryCounts(group).Values).AppendLine();
                    }
            text.AppendLine("RAW OBJECT ID OBSERVATIONS - NOT MEMBERSHIP RESULTS");
            Line(text, "ComponentType", "SourceOnlyObservedIds", "TargetOnlyObservedIds", "SameRawObjectIdsObservedBoth");
            foreach (var f in inventory.Candidates.OrderBy(f => f.Type))
                Line(text, f.Type, Ids(f.Source.Ids.Except(f.Target.Ids)), Ids(f.Target.Ids.Except(f.Source.Ids)), Ids(f.Source.Ids.Intersect(f.Target.Ids)));
            text.AppendLine("NEXT RESOLVER CANDIDATES");
            text.AppendLine("Order: total raw references, affected solution unique names, both environments, known backing, existing metadata evidence, raw type.");
            Line(text, "Rank", "ComponentType", "CatalogLabel", "TotalRawReferences", "DistinctIdsEnvironmentScoped", "AffectedSolutions",
                "OccursBothEnvironments", "KnownBackingEntity", "ExistingMetadataEvidence", "RecommendedNextAction");
            int rank = 0;
            foreach (var f in inventory.Candidates)
            {
                token.ThrowIfCancellationRequested();
                Line(text, ++rank, f.Type, f.CatalogLabel, f.TotalRaw, f.TotalDistinct, f.AffectedSolutions,
                    f.Both ? "Yes" : "No", f.KnownBacking ? "Yes" : "Unknown", f.ExistingMetadataEvidence ? "Yes" : "No", f.NextAction);
            }
            token.ThrowIfCancellationRequested();
            return text.ToString();
        }
    }
}
#endif
