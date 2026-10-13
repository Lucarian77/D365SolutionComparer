using System;
using System.Collections.Generic;
using System.Linq;
using D365SolutionComparer.Models.Presentation;

namespace D365SolutionComparer.Models.Membership
{
    public enum DisplayProjectionOutcome { Completed, Failed }
    public enum DisplayCacheAssociation { Verified, NotAvailable, Rejected }

    /// <summary>Detached display evidence for exactly this final presentation's row order.</summary>
    public sealed class CompletedMembershipComparison
    {
        public CompletedMembershipComparison(MembershipComparisonPresentation presentation,
            IEnumerable<ComponentDisplayEvidence> displayEvidence, DisplayProjectionOutcome outcome,
            DisplayCacheAssociation sourceCacheAssociation = DisplayCacheAssociation.NotAvailable,
            DisplayCacheAssociation targetCacheAssociation = DisplayCacheAssociation.NotAvailable)
        {
            Presentation = presentation ?? throw new ArgumentNullException(nameof(presentation));
            if (!Enum.IsDefined(typeof(DisplayProjectionOutcome), outcome))
                throw new ArgumentOutOfRangeException(nameof(outcome));
            if (!Enum.IsDefined(typeof(DisplayCacheAssociation), sourceCacheAssociation) ||
                !Enum.IsDefined(typeof(DisplayCacheAssociation), targetCacheAssociation))
                throw new ArgumentException("Unknown display cache association.");
            var copy = (displayEvidence ?? throw new ArgumentNullException(nameof(displayEvidence))).ToList();
            if (outcome == DisplayProjectionOutcome.Failed && copy.Count != 0)
                throw new ArgumentException("Failed projection cannot retain partial evidence.", nameof(displayEvidence));
            if (outcome == DisplayProjectionOutcome.Completed)
            {
                if (copy.Count != presentation.Rows.Count)
                    throw new ArgumentException("Display evidence must cover every final row.", nameof(displayEvidence));
                for (var index = 0; index < copy.Count; index++)
                    if (copy[index] == null || copy[index].RowIndex != index)
                        throw new ArgumentException("Display evidence must preserve final row order.", nameof(displayEvidence));
            }
            DisplayEvidence = copy.AsReadOnly();
            DisplayProjectionOutcome = outcome;
            SourceCacheAssociation = sourceCacheAssociation;
            TargetCacheAssociation = targetCacheAssociation;
        }

        internal CompletedMembershipComparison(CompletedMembershipComparison completed,
            ComponentHierarchyParentEvidence sourceParents, ComponentHierarchyParentEvidence targetParents)
            : this(completed.Presentation, completed.DisplayEvidence, completed.DisplayProjectionOutcome,
                completed.SourceCacheAssociation, completed.TargetCacheAssociation)
        {
            SourceHierarchyParents = ValidParent(sourceParents, Presentation.Source.Snapshot, DisplayEnvironmentSide.Source);
            TargetHierarchyParents = ValidParent(targetParents, Presentation.Target.Snapshot, DisplayEnvironmentSide.Target);
        }
        private static ComponentHierarchyParentEvidence ValidParent(ComponentHierarchyParentEvidence parent,
            MembershipSnapshot owner, DisplayEnvironmentSide side) => parent != null &&
            ReferenceEquals(parent.Owner, owner) && parent.Side == side ? parent : null;
        internal ComponentHierarchyParentEvidence SourceHierarchyParents { get; }
        internal ComponentHierarchyParentEvidence TargetHierarchyParents { get; }

        public MembershipComparisonPresentation Presentation { get; }
        public IReadOnlyList<ComponentDisplayEvidence> DisplayEvidence { get; }
        public DisplayProjectionOutcome DisplayProjectionOutcome { get; }
        public DisplayCacheAssociation SourceCacheAssociation { get; }
        public DisplayCacheAssociation TargetCacheAssociation { get; }
    }
}
