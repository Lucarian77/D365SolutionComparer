using System;
using System.Collections.Generic;
using System.Linq;
using D365SolutionComparer.Models.Identity;

namespace D365SolutionComparer.Models.Membership
{
    public enum MembershipSnapshotState { Complete, SolutionAbsent, Unavailable }

    /// <summary>Complete means retrieval finished, not that every component identity was resolved.</summary>
    public sealed class MembershipSnapshot
    {
        private MembershipSnapshot(EnvironmentIdentity environment, string uniqueName, SolutionIdentity solution,
            MembershipSnapshotState state, IEnumerable<ComponentIdentity> components, DateTimeOffset capturedAt, string diagnostic)
        {
            Environment = environment ?? throw new ArgumentNullException(nameof(environment));
            if (string.IsNullOrWhiteSpace(uniqueName)) throw new ArgumentException("A unique name is required.", nameof(uniqueName));
            var copy = (components ?? throw new ArgumentNullException(nameof(components))).ToList();
            if (copy.Any(c => c == null)) throw new ArgumentException("Null components are not allowed.", nameof(components));
            SolutionUniqueName = uniqueName;
            Solution = solution;
            State = state;
            Components = copy.AsReadOnly();
            CapturedAt = capturedAt;
            Diagnostic = diagnostic ?? string.Empty;
        }

        public EnvironmentIdentity Environment { get; }
        public string SolutionUniqueName { get; }
        public SolutionIdentity Solution { get; }
        public MembershipSnapshotState State { get; }
        public IReadOnlyList<ComponentIdentity> Components { get; }
        public DateTimeOffset CapturedAt { get; }
        public string Diagnostic { get; }

        public static MembershipSnapshot Complete(SolutionIdentity solution, IEnumerable<ComponentIdentity> components, DateTimeOffset capturedAt)
        {
            if (solution == null) throw new ArgumentNullException(nameof(solution));
            return new MembershipSnapshot(solution.Environment, solution.UniqueName, solution,
                MembershipSnapshotState.Complete, components, capturedAt, null);
        }

        /// <summary>Use only after positively establishing that the selected solution does not exist.</summary>
        public static MembershipSnapshot Absent(EnvironmentIdentity environment, string uniqueName, DateTimeOffset capturedAt)
        {
            return new MembershipSnapshot(environment, uniqueName, null, MembershipSnapshotState.SolutionAbsent,
                new ComponentIdentity[0], capturedAt, null);
        }

        public static MembershipSnapshot Unavailable(EnvironmentIdentity environment, string uniqueName,
            DateTimeOffset capturedAt, string diagnostic)
        {
            if (string.IsNullOrWhiteSpace(diagnostic)) throw new ArgumentException("Explain why retrieval is unavailable.", nameof(diagnostic));
            return new MembershipSnapshot(environment, uniqueName, null, MembershipSnapshotState.Unavailable,
                new ComponentIdentity[0], capturedAt, diagnostic);
        }
    }
}
