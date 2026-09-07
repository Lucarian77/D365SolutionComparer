using System;
using System.Collections.Generic;
using System.Linq;

namespace D365SolutionComparer.Models.Membership
{
    public enum MembershipCoverageBucketType
    {
        SemanticKind,
        KnownUnsupportedIsolatedType,
        DynamicallyClassifiedIsolatedFamily,
        BroadUnclassifiable
    }

    public enum MembershipCoverageStatus { Complete, Incomplete }

    public sealed class MembershipCoverageDiagnosticGroup
    {
        public MembershipCoverageDiagnosticGroup(IdentityResolutionStatus resolutionStatus,
            string diagnostic, int count)
        {
            if (resolutionStatus == IdentityResolutionStatus.Resolved)
                throw new ArgumentException("Resolved candidates are not diagnostic blockers.", nameof(resolutionStatus));
            if (count <= 0) throw new ArgumentOutOfRangeException(nameof(count));
            ResolutionStatus = resolutionStatus;
            Diagnostic = diagnostic ?? string.Empty;
            Count = count;
        }

        public IdentityResolutionStatus ResolutionStatus { get; }
        public string Diagnostic { get; }
        public int Count { get; }
    }

    public sealed class MembershipCoverageBucket
    {
        public MembershipCoverageBucket(string semanticKind, string displayName,
            MembershipCoverageBucketType bucketType, int totalCandidates, int resolved,
            int unsupported, int unresolved, int ambiguous, MembershipCoverageStatus coverageStatus,
            IEnumerable<MembershipCoverageDiagnosticGroup> diagnosticGroups)
        {
            if (bucketType != MembershipCoverageBucketType.BroadUnclassifiable &&
                string.IsNullOrWhiteSpace(semanticKind))
                throw new ArgumentException("A semantic coverage bucket requires a kind.", nameof(semanticKind));
            if (bucketType == MembershipCoverageBucketType.BroadUnclassifiable && semanticKind != null)
                throw new ArgumentException("A broad coverage bucket cannot have a semantic kind.", nameof(semanticKind));
            if (string.IsNullOrWhiteSpace(displayName))
                throw new ArgumentException("A coverage bucket display name is required.", nameof(displayName));
            if (totalCandidates < 0 || resolved < 0 || unsupported < 0 || unresolved < 0 || ambiguous < 0)
                throw new ArgumentOutOfRangeException(nameof(totalCandidates));
            if (resolved + unsupported + unresolved + ambiguous != totalCandidates)
                throw new ArgumentException("Coverage status counts must equal the total candidate count.");
            SemanticKind = semanticKind;
            DisplayName = displayName;
            BucketType = bucketType;
            TotalCandidates = totalCandidates;
            Resolved = resolved;
            Unsupported = unsupported;
            Unresolved = unresolved;
            Ambiguous = ambiguous;
            CoverageStatus = coverageStatus;
            DiagnosticGroups = new List<MembershipCoverageDiagnosticGroup>(diagnosticGroups ??
                throw new ArgumentNullException(nameof(diagnosticGroups))).AsReadOnly();
        }

        public string SemanticKind { get; }
        public string DisplayName { get; }
        public MembershipCoverageBucketType BucketType { get; }
        public int TotalCandidates { get; }
        public int Resolved { get; }
        public int Unsupported { get; }
        public int Unresolved { get; }
        public int Ambiguous { get; }
        public MembershipCoverageStatus CoverageStatus { get; }
        public IReadOnlyList<MembershipCoverageDiagnosticGroup> DiagnosticGroups { get; }
    }

    public sealed class MembershipCoverageRawComponentTypeGroup
    {
        public MembershipCoverageRawComponentTypeGroup(int componentType, int count,
            IEnumerable<MembershipCoverageDiagnosticGroup> diagnosticGroups)
        {
            if (count <= 0) throw new ArgumentOutOfRangeException(nameof(count));
            ComponentType = componentType;
            Count = count;
            DiagnosticGroups = new List<MembershipCoverageDiagnosticGroup>(diagnosticGroups ??
                throw new ArgumentNullException(nameof(diagnosticGroups))).AsReadOnly();
            if (DiagnosticGroups.Sum(group => group.Count) > count)
                throw new ArgumentException("Diagnostic counts cannot exceed the raw component-type count.",
                    nameof(diagnosticGroups));
        }

        public int ComponentType { get; }
        public int Count { get; }
        public IReadOnlyList<MembershipCoverageDiagnosticGroup> DiagnosticGroups { get; }
    }

    public sealed class MembershipCoverageDynamicComponentTypeGroup
    {
        public MembershipCoverageDynamicComponentTypeGroup(int componentType, int count,
            SolutionComponentDefinitionIdentity definition)
        {
            if (count <= 0) throw new ArgumentOutOfRangeException(nameof(count));
            Definition = definition ?? throw new ArgumentNullException(nameof(definition));
            if (definition.ObjectTypeCode != componentType)
                throw new ArgumentException("The definition belongs to a different raw component type.",
                    nameof(definition));
            ComponentType = componentType;
            Count = count;
        }

        public int ComponentType { get; }
        public int Count { get; }
        public SolutionComponentDefinitionIdentity Definition { get; }
        public string SemanticKind => Definition.SemanticKind;
    }

    public sealed class MembershipCoverageDiagnostics
    {
        public MembershipCoverageDiagnostics(MembershipSnapshotState snapshotState,
            IEnumerable<MembershipCoverageBucket> semanticKinds,
            MembershipCoverageBucket broadUnclassifiable,
            IEnumerable<MembershipCoverageRawComponentTypeGroup> broadRawComponentTypes,
            IEnumerable<MembershipCoverageDynamicComponentTypeGroup> dynamicComponentTypes)
        {
            SnapshotState = snapshotState;
            SemanticKinds = new List<MembershipCoverageBucket>(semanticKinds ??
                throw new ArgumentNullException(nameof(semanticKinds))).AsReadOnly();
            BroadUnclassifiable = broadUnclassifiable ??
                throw new ArgumentNullException(nameof(broadUnclassifiable));
            if (BroadUnclassifiable.BucketType != MembershipCoverageBucketType.BroadUnclassifiable)
                throw new ArgumentException("The broad blocker summary has the wrong bucket type.",
                    nameof(broadUnclassifiable));
            BroadRawComponentTypes = new List<MembershipCoverageRawComponentTypeGroup>(
                broadRawComponentTypes ?? throw new ArgumentNullException(nameof(broadRawComponentTypes)))
                .AsReadOnly();
            if (BroadRawComponentTypes.Select(group => group.ComponentType).Distinct().Count() !=
                BroadRawComponentTypes.Count)
                throw new ArgumentException("Broad raw component types must be unique.",
                    nameof(broadRawComponentTypes));
            if (BroadRawComponentTypes.Sum(group => group.Count) != BroadUnclassifiable.TotalCandidates)
                throw new ArgumentException("Broad raw component-type counts must equal the broad candidate count.",
                    nameof(broadRawComponentTypes));
            DynamicComponentTypes = new List<MembershipCoverageDynamicComponentTypeGroup>(
                dynamicComponentTypes ?? throw new ArgumentNullException(nameof(dynamicComponentTypes)))
                .AsReadOnly();
            if (DynamicComponentTypes.Select(group => group.ComponentType).Distinct().Count() !=
                DynamicComponentTypes.Count)
                throw new ArgumentException("Dynamically classified raw component types must be unique.",
                    nameof(dynamicComponentTypes));
            var dynamicBuckets = SemanticKinds.Where(bucket =>
                bucket.BucketType == MembershipCoverageBucketType.DynamicallyClassifiedIsolatedFamily).ToList();
            if (DynamicComponentTypes.Sum(group => group.Count) !=
                dynamicBuckets.Sum(bucket => bucket.TotalCandidates))
                throw new ArgumentException("Dynamic component-type counts must equal dynamic bucket candidates.",
                    nameof(dynamicComponentTypes));
            foreach (var bucket in dynamicBuckets)
                if (DynamicComponentTypes.Where(group => string.Equals(group.SemanticKind, bucket.SemanticKind,
                        StringComparison.OrdinalIgnoreCase)).Sum(group => group.Count) != bucket.TotalCandidates)
                    throw new ArgumentException("Dynamic component-type counts do not reconcile by family.",
                        nameof(dynamicComponentTypes));
            TotalCandidates = SemanticKinds.Sum(bucket => bucket.TotalCandidates) +
                BroadUnclassifiable.TotalCandidates;
        }

        public MembershipSnapshotState SnapshotState { get; }
        public IReadOnlyList<MembershipCoverageBucket> SemanticKinds { get; }
        public MembershipCoverageBucket BroadUnclassifiable { get; }
        public IReadOnlyList<MembershipCoverageRawComponentTypeGroup> BroadRawComponentTypes { get; }
        public IReadOnlyList<MembershipCoverageDynamicComponentTypeGroup> DynamicComponentTypes { get; }
        public int TotalCandidates { get; }
        public bool HasBroadUnclassifiableBlockers => BroadUnclassifiable.TotalCandidates > 0;
    }
}
