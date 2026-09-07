using System;
using System.Collections.Generic;
using System.Linq;

namespace D365SolutionComparer.Models.Membership
{
    public sealed class MembershipEnvironmentDiagnostics
    {
        public MembershipEnvironmentDiagnostics(string environmentName, MembershipSnapshotState state,
            int requestCount, TimeSpan elapsed, int? rawMembershipCount, int? resolvedCount,
            int? unsupportedCount, int? unresolvedCount, int? ambiguousCount, string diagnostic)
        {
            EnvironmentName = environmentName ?? string.Empty;
            State = state;
            RequestCount = requestCount;
            Elapsed = elapsed;
            RawMembershipCount = rawMembershipCount;
            ResolvedCount = resolvedCount;
            UnsupportedCount = unsupportedCount;
            UnresolvedCount = unresolvedCount;
            AmbiguousCount = ambiguousCount;
            Diagnostic = diagnostic ?? string.Empty;
        }

        public string EnvironmentName { get; }
        public MembershipSnapshotState State { get; }
        public int RequestCount { get; }
        public TimeSpan Elapsed { get; }
        public int? RawMembershipCount { get; }
        public int? ResolvedCount { get; }
        public int? UnsupportedCount { get; }
        public int? UnresolvedCount { get; }
        public int? AmbiguousCount { get; }
        public string Diagnostic { get; }
    }

    /// <summary>An environment read may be unavailable without fabricating an EnvironmentIdentity.</summary>
    public sealed class MembershipEnvironmentResult
    {
        private MembershipEnvironmentResult(string solutionUniqueName, MembershipSnapshot snapshot,
            MembershipEnvironmentDiagnostics diagnostics)
        {
            if (string.IsNullOrWhiteSpace(solutionUniqueName))
                throw new ArgumentException("A solution Unique Name is required.", nameof(solutionUniqueName));
            SolutionUniqueName = solutionUniqueName;
            Snapshot = snapshot;
            Diagnostics = diagnostics ?? throw new ArgumentNullException(nameof(diagnostics));
            if (snapshot != null && !string.Equals(snapshot.SolutionUniqueName, solutionUniqueName,
                StringComparison.OrdinalIgnoreCase))
                throw new ArgumentException("The snapshot refers to a different solution.", nameof(snapshot));
            if (snapshot == null && diagnostics.State != MembershipSnapshotState.Unavailable)
                throw new ArgumentException("A missing snapshot must be reported as unavailable.", nameof(diagnostics));
        }

        public string SolutionUniqueName { get; }
        public MembershipSnapshot Snapshot { get; }
        public MembershipEnvironmentDiagnostics Diagnostics { get; }
        public MembershipSnapshotState State => Snapshot == null ? MembershipSnapshotState.Unavailable : Snapshot.State;

        public static MembershipEnvironmentResult FromSnapshot(string environmentName, MembershipSnapshot snapshot,
            int requestCount, TimeSpan elapsed)
        {
            if (snapshot == null) throw new ArgumentNullException(nameof(snapshot));
            var components = snapshot.Components;
            var hasInventory = snapshot.State != MembershipSnapshotState.Unavailable;
            var diagnostics = new MembershipEnvironmentDiagnostics(environmentName, snapshot.State, requestCount, elapsed,
                hasInventory ? (int?)components.Count : null,
                hasInventory ? (int?)components.Count(item => item.Status == IdentityResolutionStatus.Resolved) : null,
                hasInventory ? (int?)components.Count(item => item.Status == IdentityResolutionStatus.Unsupported) : null,
                hasInventory ? (int?)components.Count(item => item.Status == IdentityResolutionStatus.Unresolved) : null,
                hasInventory ? (int?)components.Count(item => item.Status == IdentityResolutionStatus.Ambiguous) : null,
                snapshot.Diagnostic);
            return new MembershipEnvironmentResult(snapshot.SolutionUniqueName, snapshot, diagnostics);
        }

        public static MembershipEnvironmentResult Unavailable(string environmentName, string solutionUniqueName,
            int requestCount, TimeSpan elapsed, string diagnostic)
        {
            if (string.IsNullOrWhiteSpace(diagnostic))
                throw new ArgumentException("An unavailable read requires a diagnostic.", nameof(diagnostic));
            return new MembershipEnvironmentResult(solutionUniqueName, null,
                new MembershipEnvironmentDiagnostics(environmentName, MembershipSnapshotState.Unavailable,
                    requestCount, elapsed, null, null, null, null, null, diagnostic));
        }
    }

    public sealed class MembershipPresentationSummary
    {
        internal MembershipPresentationSummary(IEnumerable<MembershipResultRow> rows)
        {
            var items = rows.ToList();
            PresentInBoth = items.Count(item => item.Presence == MembershipPresence.PresentInBoth);
            SourceOnly = items.Count(item => item.Presence == MembershipPresence.OnlyInSource);
            TargetOnly = items.Count(item => item.Presence == MembershipPresence.OnlyInTarget);
            Unsupported = items.Count(item => item.HasResolutionStatus(IdentityResolutionStatus.Unsupported));
            Unresolved = items.Count(item => item.HasResolutionStatus(IdentityResolutionStatus.Unresolved));
            Ambiguous = items.Count(item => item.HasResolutionStatus(IdentityResolutionStatus.Ambiguous));
        }

        public int PresentInBoth { get; }
        public int SourceOnly { get; }
        public int TargetOnly { get; }
        public int Unsupported { get; }
        public int Unresolved { get; }
        public int Ambiguous { get; }
    }

    public sealed class MembershipResultRow
    {
        internal MembershipResultRow(MembershipPresence presence, ComponentIdentity source, ComponentIdentity target,
            string componentKind, string portableKey, string sourcePresence, string targetPresence,
            string membershipStatus, string sourceResolutionStatus, string targetResolutionStatus,
            string diagnostic)
        {
            Presence = presence;
            SourceIdentity = source;
            TargetIdentity = target;
            ComponentKind = componentKind;
            PortableKey = portableKey;
            SourcePresence = sourcePresence;
            TargetPresence = targetPresence;
            MembershipStatus = membershipStatus;
            SourceResolutionStatus = sourceResolutionStatus;
            TargetResolutionStatus = targetResolutionStatus;
            Diagnostic = diagnostic;
            SourceRawComponentType = source?.Record.ComponentType;
            TargetRawComponentType = target?.Record.ComponentType;
        }

        internal ComponentIdentity SourceIdentity { get; }
        internal ComponentIdentity TargetIdentity { get; }
        internal MembershipPresence Presence { get; }
        public string ComponentKind { get; }
        public string PortableKey { get; }
        public string SourcePresence { get; }
        public string TargetPresence { get; }
        public string MembershipStatus { get; }
        public string SourceResolutionStatus { get; }
        public string TargetResolutionStatus { get; }
        public string Diagnostic { get; }
        public int? SourceRawComponentType { get; }
        public int? TargetRawComponentType { get; }

        internal bool HasResolutionStatus(IdentityResolutionStatus status) =>
            SourceIdentity?.Status == status || TargetIdentity?.Status == status;
    }

    public sealed class MembershipComparisonPresentation
    {
        public MembershipComparisonPresentation(string solutionUniqueName, MembershipEnvironmentResult source,
            MembershipEnvironmentResult target, IEnumerable<MembershipResultRow> rows)
        {
            if (string.IsNullOrWhiteSpace(solutionUniqueName))
                throw new ArgumentException("A solution Unique Name is required.", nameof(solutionUniqueName));
            SolutionUniqueName = solutionUniqueName;
            Source = source ?? throw new ArgumentNullException(nameof(source));
            Target = target ?? throw new ArgumentNullException(nameof(target));
            var copy = (rows ?? throw new ArgumentNullException(nameof(rows))).ToList();
            Rows = copy.AsReadOnly();
            Summary = new MembershipPresentationSummary(copy);
        }

        public string SolutionUniqueName { get; }
        public MembershipEnvironmentResult Source { get; }
        public MembershipEnvironmentResult Target { get; }
        public IReadOnlyList<MembershipResultRow> Rows { get; }
        public MembershipPresentationSummary Summary { get; }
    }
}
