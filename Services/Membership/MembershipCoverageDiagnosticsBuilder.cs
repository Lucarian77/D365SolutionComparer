using System;
using System.Collections.Generic;
using System.Linq;
using D365SolutionComparer.Models.Membership;

namespace D365SolutionComparer.Services.Membership
{
    /// <summary>Builds read-only coverage diagnostics without changing comparison results.</summary>
    public sealed class MembershipCoverageDiagnosticsBuilder
    {
        private static readonly string[] SupportedKinds =
        {
            ComponentSemanticKinds.Table,
            ComponentSemanticKinds.Column,
            ComponentSemanticKinds.Relationship,
            ComponentSemanticKinds.WebResource,
            ComponentSemanticKinds.Process,
            ComponentSemanticKinds.SecurityRole,
            ComponentSemanticKinds.EnvironmentVariableDefinition,
            ComponentSemanticKinds.ConnectionReference
        };

        public MembershipCoverageDiagnostics Build(MembershipSnapshot snapshot)
        {
            if (snapshot == null) throw new ArgumentNullException(nameof(snapshot));
            return Build(snapshot.State, snapshot.Components);
        }

        public MembershipCoverageDiagnostics BuildUnavailable()
        {
            return Build(MembershipSnapshotState.Unavailable, new ComponentIdentity[0]);
        }

        private static MembershipCoverageDiagnostics Build(MembershipSnapshotState state,
            IReadOnlyList<ComponentIdentity> components)
        {
            var broadCandidates = components.Where(item => string.IsNullOrWhiteSpace(item.SemanticKind)).ToList();
            bool hasBroadBlockers = broadCandidates.Count > 0;
            var broad = CreateBucket(null, "Broad / Unclassifiable blockers",
                MembershipCoverageBucketType.BroadUnclassifiable, broadCandidates, state,
                hasBroadBlockers: false);
            var broadRawComponentTypes = broadCandidates
                .GroupBy(item => item.Record.ComponentType)
                .OrderBy(group => group.Key)
                .Select(group => new MembershipCoverageRawComponentTypeGroup(group.Key, group.Count(),
                    CreateDiagnosticGroups(group)))
                .ToList();
            var dynamicComponentTypes = components.Where(item => item.RegisteredDefinition != null)
                .GroupBy(item => item.Record.ComponentType)
                .OrderBy(group => group.Key)
                .Select(CreateDynamicComponentTypeGroup)
                .ToList();

            var kinds = new HashSet<string>(SupportedKinds, StringComparer.OrdinalIgnoreCase);
            foreach (var kind in components.Where(item => !string.IsNullOrWhiteSpace(item.SemanticKind))
                .Select(item => item.SemanticKind))
                kinds.Add(kind);

            var orderedKinds = SupportedKinds.Concat(kinds.Except(SupportedKinds, StringComparer.OrdinalIgnoreCase)
                .OrderBy(kind => kind, StringComparer.OrdinalIgnoreCase));
            var summaries = orderedKinds.Select(kind =>
            {
                var candidates = components.Where(item => string.Equals(item.SemanticKind, kind,
                    StringComparison.OrdinalIgnoreCase)).ToList();
                var bucketType = kind.StartsWith("unsupported:componenttype:",
                    StringComparison.OrdinalIgnoreCase)
                    ? MembershipCoverageBucketType.KnownUnsupportedIsolatedType
                    : ComponentSemanticKinds.IsRegisteredDefinitionKind(kind)
                    ? MembershipCoverageBucketType.DynamicallyClassifiedIsolatedFamily
                    : MembershipCoverageBucketType.SemanticKind;
                return CreateBucket(kind, DisplayName(kind, candidates), bucketType, candidates, state,
                    hasBroadBlockers);
            }).ToList();

            return new MembershipCoverageDiagnostics(state, summaries, broad, broadRawComponentTypes,
                dynamicComponentTypes);
        }

        private static MembershipCoverageDynamicComponentTypeGroup CreateDynamicComponentTypeGroup(
            IGrouping<int, ComponentIdentity> group)
        {
            var definitions = group.Select(item => item.RegisteredDefinition).ToList();
            var first = definitions[0];
            if (definitions.Any(item => item == null ||
                !string.Equals(item.SemanticKind, first.SemanticKind, StringComparison.OrdinalIgnoreCase) ||
                !string.Equals(item.Name, first.Name, StringComparison.Ordinal) ||
                !string.Equals(item.PrimaryEntityName, first.PrimaryEntityName, StringComparison.Ordinal)))
                throw new InvalidOperationException(
                    "A raw component type has conflicting registered definition diagnostics.");
            return new MembershipCoverageDynamicComponentTypeGroup(group.Key, definitions.Count, first);
        }

        private static MembershipCoverageBucket CreateBucket(string semanticKind, string displayName,
            MembershipCoverageBucketType bucketType, IReadOnlyList<ComponentIdentity> candidates,
            MembershipSnapshotState state, bool hasBroadBlockers)
        {
            int resolved = candidates.Count(item => item.Status == IdentityResolutionStatus.Resolved);
            int unsupported = candidates.Count(item => item.Status == IdentityResolutionStatus.Unsupported);
            int unresolved = candidates.Count(item => item.Status == IdentityResolutionStatus.Unresolved);
            int ambiguous = candidates.Count(item => item.Status == IdentityResolutionStatus.Ambiguous);
            bool localBlocker = unsupported > 0 || unresolved > 0 || ambiguous > 0;
            bool complete = state == MembershipSnapshotState.SolutionAbsent ||
                state == MembershipSnapshotState.Complete && !localBlocker && !hasBroadBlockers;
            if (bucketType == MembershipCoverageBucketType.BroadUnclassifiable)
                complete = state == MembershipSnapshotState.SolutionAbsent ||
                    state == MembershipSnapshotState.Complete && candidates.Count == 0;

            return new MembershipCoverageBucket(semanticKind, displayName, bucketType, candidates.Count,
                resolved, unsupported, unresolved, ambiguous,
                complete ? MembershipCoverageStatus.Complete : MembershipCoverageStatus.Incomplete,
                CreateDiagnosticGroups(candidates));
        }

        private static IReadOnlyList<MembershipCoverageDiagnosticGroup> CreateDiagnosticGroups(
            IEnumerable<ComponentIdentity> candidates)
        {
            return candidates.Where(item => item.Status != IdentityResolutionStatus.Resolved)
                .GroupBy(item => new DiagnosticKey(item.Status, item.Diagnostic))
                .OrderBy(group => group.Key.Status)
                .ThenBy(group => group.Key.Diagnostic, StringComparer.Ordinal)
                .Select(group => new MembershipCoverageDiagnosticGroup(group.Key.Status,
                    group.Key.Diagnostic, group.Count())).ToList();
        }

        private static string DisplayName(string semanticKind, IReadOnlyList<ComponentIdentity> candidates)
        {
            switch (semanticKind)
            {
                case ComponentSemanticKinds.Table: return "Table";
                case ComponentSemanticKinds.Column: return "Column";
                case ComponentSemanticKinds.Relationship: return "Relationship";
                case ComponentSemanticKinds.WebResource: return "Web Resource";
                case ComponentSemanticKinds.Process: return "Process / Workflow";
                case ComponentSemanticKinds.SecurityRole: return "Security Role";
                case ComponentSemanticKinds.EnvironmentVariableDefinition: return "Environment Variable Definition";
                case ComponentSemanticKinds.ConnectionReference: return "Connection Reference";
            }
            const string prefix = "unsupported:componenttype:";
            if (semanticKind.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
                return "Isolated unsupported component type " + semanticKind.Substring(prefix.Length);
            if (ComponentSemanticKinds.IsRegisteredDefinitionKind(semanticKind))
            {
                var definition = candidates.Select(item => item.RegisteredDefinition)
                    .FirstOrDefault(item => item != null);
                return "Registered family " + (definition == null
                    ? semanticKind.Substring(ComponentSemanticKinds.RegisteredDefinitionPrefix.Length)
                    : definition.Name);
            }
            return semanticKind;
        }

        private struct DiagnosticKey : IEquatable<DiagnosticKey>
        {
            public DiagnosticKey(IdentityResolutionStatus status, string diagnostic)
            {
                Status = status;
                Diagnostic = diagnostic ?? string.Empty;
            }

            public IdentityResolutionStatus Status { get; }
            public string Diagnostic { get; }
            public bool Equals(DiagnosticKey other) => Status == other.Status &&
                string.Equals(Diagnostic, other.Diagnostic, StringComparison.Ordinal);
            public override bool Equals(object obj) => obj is DiagnosticKey && Equals((DiagnosticKey)obj);
            public override int GetHashCode() => unchecked(((int)Status * 397) ^
                StringComparer.Ordinal.GetHashCode(Diagnostic));
        }
    }
}
