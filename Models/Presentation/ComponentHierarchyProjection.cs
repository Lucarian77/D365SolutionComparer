using System;
using System.Collections.Generic;
using System.Linq;
using D365SolutionComparer.Models.Membership;

namespace D365SolutionComparer.Models.Presentation
{
    public enum HierarchyProjectionOutcome { Completed, Cancelled, Failed }
    public enum HierarchyPlacement { VerifiedInclusionRoot, VerifiedStructuralParent, ParentNotEstablished }
    public enum HierarchyRelationshipKind { InclusionRoot, StructuralParent }
    public enum HierarchyAssociationStatus
    {
        Verified, NotSupplied, MissingParent, DuplicateParent, SelfReference,
        CyclicReference, UnverifiedAncestor, InvalidReference, UnverifiedEvidence, ConflictingEvidence,
        ParentIdentityNotEstablished
    }

    /// <summary>A local association, never a new membership identity or schema ownership inference.</summary>
    public sealed class ComponentHierarchyAssociation
    {
        internal ComponentHierarchyAssociation(HierarchyRelationshipKind kind, HierarchyAssociationStatus status,
            int? parentOccurrenceIndex, Guid? referencedLocalId, string provenance, string diagnostic)
        {
            Kind = kind; Status = status; ParentOccurrenceIndex = parentOccurrenceIndex;
            ReferencedLocalId = referencedLocalId; Provenance = provenance; Diagnostic = diagnostic;
        }
        public HierarchyRelationshipKind Kind { get; }
        public HierarchyAssociationStatus Status { get; }
        public int? ParentOccurrenceIndex { get; }
        public Guid? ReferencedLocalId { get; }
        public string Provenance { get; }
        public string Diagnostic { get; }
    }

    /// <summary>Exactly one node per raw occurrence, including identical/repeated records.</summary>
    public sealed class ComponentHierarchyOccurrence
    {
        internal ComponentHierarchyOccurrence(int index, DisplayEnvironmentSide side, int rawIndex,
            MembershipSnapshot snapshot, ComponentIdentity identity, IEnumerable<int> comparisons,
            IEnumerable<ComponentHierarchyAssociation> associations)
        {
            OccurrenceIndex = index; Side = side; OriginalRawIndex = rawIndex; Snapshot = snapshot;
            Identity = identity; ComparisonRowIndices = comparisons.ToList().AsReadOnly();
            Associations = associations.ToList().AsReadOnly();
            Placement = Associations.Any(a => a.Status == HierarchyAssociationStatus.ConflictingEvidence ||
                a.Status == HierarchyAssociationStatus.CyclicReference)
                ? HierarchyPlacement.ParentNotEstablished
                : Associations.Any(a => a.Kind == HierarchyRelationshipKind.InclusionRoot && a.Status == HierarchyAssociationStatus.Verified)
                    ? HierarchyPlacement.VerifiedInclusionRoot
                    : !Record.RootSolutionComponentId.HasValue && Associations.Any(a =>
                        a.Kind == HierarchyRelationshipKind.StructuralParent && a.Status == HierarchyAssociationStatus.Verified) &&
                        Associations.Where(a => a.Kind == HierarchyRelationshipKind.StructuralParent)
                            .All(a => a.Status == HierarchyAssociationStatus.Verified)
                        ? HierarchyPlacement.VerifiedStructuralParent : HierarchyPlacement.ParentNotEstablished;
        }
        public int OccurrenceIndex { get; }
        public DisplayEnvironmentSide Side { get; }
        public int OriginalRawIndex { get; }
        public MembershipSnapshot Snapshot { get; }
        public ComponentIdentity Identity { get; }
        public SolutionComponentRecord Record => Identity.Record;
        public IReadOnlyList<int> ComparisonRowIndices { get; }
        public IReadOnlyList<ComponentHierarchyAssociation> Associations { get; }
        public HierarchyPlacement Placement { get; }
        // ParentNotEstablished is deliberately not "top-level".
        public bool HasDifferentParentAssociations => Associations.Where(a => a.Status == HierarchyAssociationStatus.Verified)
            .Select(a => a.ParentOccurrenceIndex).Distinct().Count() > 1;
    }

    public sealed class HierarchyComparisonRowLink
    {
        internal HierarchyComparisonRowLink(int index, MembershipResultRow row, IEnumerable<int> source, IEnumerable<int> target)
        { RowIndex = index; Row = row; SourceOccurrences = source.ToList().AsReadOnly(); TargetOccurrences = target.ToList().AsReadOnly(); }
        public int RowIndex { get; }
        public MembershipResultRow Row { get; }
        public IReadOnlyList<int> SourceOccurrences { get; }
        public IReadOnlyList<int> TargetOccurrences { get; }
    }

    /// <summary>Evidence of an EXISTING authoritative shared parent row; nodes remain side-specific.</summary>
    public sealed class HierarchyParentCorrelation
    {
        internal HierarchyParentCorrelation(int row, int source, int target, string kind, string key)
        { ComparisonRowIndex = row; SourceOccurrenceIndex = source; TargetOccurrenceIndex = target; ComponentTypeKey = kind; PortableKey = key; }
        public int ComparisonRowIndex { get; }
        public int SourceOccurrenceIndex { get; }
        public int TargetOccurrenceIndex { get; }
        public string ComponentTypeKey { get; }
        public string PortableKey { get; }
    }

    /// <summary>Detached immutable result. The original snapshots, rows and diagnostics remain authoritative.</summary>
    public sealed class ComponentHierarchyProjection
    {
        internal ComponentHierarchyProjection(MembershipComparisonPresentation presentation, HierarchyProjectionOutcome outcome,
            IEnumerable<ComponentHierarchyOccurrence> occurrences, IEnumerable<HierarchyComparisonRowLink> links,
            IEnumerable<HierarchyParentCorrelation> correlations, IEnumerable<string> diagnostics)
        {
            Presentation = presentation; Outcome = outcome;
            Occurrences = occurrences.ToList().AsReadOnly(); ComparisonRows = links.ToList().AsReadOnly();
            ParentCorrelations = correlations.ToList().AsReadOnly(); Diagnostics = diagnostics.ToList().AsReadOnly();
        }
        public MembershipComparisonPresentation Presentation { get; }
        public HierarchyProjectionOutcome Outcome { get; }
        public IReadOnlyList<ComponentHierarchyOccurrence> Occurrences { get; }
        public IReadOnlyList<HierarchyComparisonRowLink> ComparisonRows { get; }
        public IReadOnlyList<HierarchyParentCorrelation> ParentCorrelations { get; }
        public IReadOnlyList<string> Diagnostics { get; }
    }

    /// <summary>Detached observations bound to the exact completed snapshot and role that supplied them.</summary>
    public sealed class ComponentHierarchyParentEvidence
    {
        internal ComponentHierarchyParentEvidence(MembershipSnapshot owner, DisplayEnvironmentSide side,
            IEnumerable<HierarchyStructuralObservation> observations, string diagnostic = null)
        { Owner = owner; Side = side; Observations = observations.ToList().AsReadOnly(); Diagnostic = diagnostic; }
        public MembershipSnapshot Owner { get; }
        public DisplayEnvironmentSide Side { get; }
        public string Diagnostic { get; }
        internal IReadOnlyList<HierarchyStructuralObservation> Observations { get; }
    }

    internal sealed class HierarchyStructuralObservation
    {
        internal HierarchyStructuralObservation(Guid child, Guid? childObject, Guid? parentObject,
            string parentKind, string parentKey, string provenance, bool verified)
        { Child = child; ChildObject = childObject; ParentObject = parentObject; ParentKind = parentKind;
            ParentKey = parentKey; Provenance = provenance; Verified = verified; }
        internal Guid Child { get; }
        internal Guid? ChildObject { get; }
        internal Guid? ParentObject { get; }
        internal string ParentKind { get; }
        internal string ParentKey { get; }
        internal string Provenance { get; }
        internal bool Verified { get; }
    }
}
