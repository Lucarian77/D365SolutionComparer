using System;

namespace D365SolutionComparer.Models.Membership
{
    public enum MembershipPresence { Indeterminate, PresentInBoth, OnlyInSource, OnlyInTarget }
    public enum MembershipAbsenceEvidence { None, OppositeSolutionAbsent, CompleteResolvedInventory }

    /// <summary>Presence says nothing about definition equality. Missing requires explicit evidence.</summary>
    public sealed class MembershipCompareResult
    {
        public MembershipCompareResult(ComponentIdentity source, ComponentIdentity target, MembershipPresence presence,
            MembershipAbsenceEvidence absenceEvidence = MembershipAbsenceEvidence.None)
        {
            if (!Enum.IsDefined(typeof(MembershipPresence), presence)) throw new ArgumentOutOfRangeException(nameof(presence));
            if (!Enum.IsDefined(typeof(MembershipAbsenceEvidence), absenceEvidence)) throw new ArgumentOutOfRangeException(nameof(absenceEvidence));
            if (source == null && target == null) throw new ArgumentException("At least one component is required.");
            bool sourceResolved = source != null && source.Status == IdentityResolutionStatus.Resolved;
            bool targetResolved = target != null && target.Status == IdentityResolutionStatus.Resolved;
            if (presence == MembershipPresence.PresentInBoth && (!sourceResolved || !targetResolved))
                throw new ArgumentException("Both identities must be resolved to establish shared presence.");
            if (presence == MembershipPresence.PresentInBoth && !string.Equals(source.ComponentTypeKey, target.ComponentTypeKey, StringComparison.OrdinalIgnoreCase))
                throw new ArgumentException("Different component types cannot establish shared presence.");
            bool oneSided = presence == MembershipPresence.OnlyInSource || presence == MembershipPresence.OnlyInTarget;
            if (oneSided)
            {
                bool validSide = presence == MembershipPresence.OnlyInSource
                    ? sourceResolved && target == null : targetResolved && source == null;
                if (!validSide || absenceEvidence == MembershipAbsenceEvidence.None)
                    throw new ArgumentException("Missing requires a resolved identity, an empty opposite side, and absence evidence.");
            }
            else if (absenceEvidence != MembershipAbsenceEvidence.None)
                throw new ArgumentException("Absence evidence applies only to one-sided results.", nameof(absenceEvidence));
            Source = source;
            Target = target;
            Presence = presence;
            AbsenceEvidence = absenceEvidence;
        }

        public ComponentIdentity Source { get; }
        public ComponentIdentity Target { get; }
        public MembershipPresence Presence { get; }
        public MembershipAbsenceEvidence AbsenceEvidence { get; }
    }
}
