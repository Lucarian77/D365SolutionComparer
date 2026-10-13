using System;

namespace D365SolutionComparer.Models.Membership
{
    public enum DisplayAvailability
    {
        Available, Unavailable, NotRetrieved, AttributeOmitted, Ambiguous,
        RetrievalFailed, NotApplicable
    }

    public enum DisplayEnvironmentSide { Source, Target }

    public enum DisplayBackingCorrelationStatus
    {
        Unavailable, NotAttempted, Unique, Ambiguous, Missing, RetrievalFailed,
        VerifiedMetadata, VerifiedDefinition
    }

    /// <summary>
    /// An immutable display observation. Available null and Available empty are actual
    /// values; neither means omitted, unrequested, unavailable, or failed.
    /// </summary>
    public sealed class ComponentDisplayValue
    {
        public ComponentDisplayValue(DisplayAvailability availability, string value = null,
            string displayEvidenceSource = null)
        {
            if (!Enum.IsDefined(typeof(DisplayAvailability), availability))
                throw new ArgumentOutOfRangeException(nameof(availability));
            if (availability != DisplayAvailability.Available && value != null)
                throw new ArgumentException("Unavailable display evidence cannot carry a value.", nameof(value));
            if (availability == DisplayAvailability.Available && string.IsNullOrWhiteSpace(displayEvidenceSource))
                throw new ArgumentException("An observed value requires provenance.", nameof(displayEvidenceSource));
            Availability = availability;
            Value = value;
            DisplayEvidenceSource = displayEvidenceSource;
        }

        public DisplayAvailability Availability { get; }
        public string Value { get; }
        public string DisplayEvidenceSource { get; }
    }

    /// <summary>
    /// Local display evidence only. IDs are audit locators, never portable keys.
    /// Contains no authoritative identity, comparison, SDK entity, or payload reference.
    /// </summary>
    public sealed class ComponentSideDisplayEvidence
    {
        internal ComponentSideDisplayEvidence(DisplayEnvironmentSide side, Guid? solutionComponentId,
            Guid? localObjectId, int? rawComponentType, ComponentDisplayValue category,
            ComponentDisplayValue friendlyName, ComponentDisplayValue internalName,
            DisplayBackingCorrelationStatus backingCorrelationStatus,
            ComponentDisplayValue parentDisplayContext, ComponentObjectMetadataEvidence objectMetadata = null)
        {
            if (objectMetadata != null && objectMetadata.EnvironmentSide != side)
                throw new ArgumentException("Object metadata belongs to a different environment side.", nameof(objectMetadata));
            EnvironmentSide = side;
            SolutionComponentId = solutionComponentId;
            LocalObjectId = localObjectId;
            RawComponentType = rawComponentType;
            ComponentCategory = category;
            FriendlyName = friendlyName;
            InternalName = internalName;
            BackingCorrelationStatus = backingCorrelationStatus;
            ParentDisplayContext = parentDisplayContext;
            ObjectMetadata = objectMetadata ?? ComponentObjectMetadataEvidence.NotRetrieved(side);
        }

        public DisplayEnvironmentSide EnvironmentSide { get; }
        public bool HasLocalObservation => SolutionComponentId.HasValue;
        public Guid? SolutionComponentId { get; }
        public Guid? LocalObjectId { get; }
        public int? RawComponentType { get; }
        public ComponentDisplayValue ComponentCategory { get; }
        public ComponentDisplayValue FriendlyName { get; }
        public ComponentDisplayValue InternalName { get; }
        public DisplayBackingCorrelationStatus BackingCorrelationStatus { get; }
        public ComponentDisplayValue ParentDisplayContext { get; }
        public ComponentObjectMetadataEvidence ObjectMetadata { get; }
    }

    /// <summary>
    /// One projection per existing presentation row, in the same order. Two local
    /// observations here do not establish Present in Both. The original results stay authoritative.
    /// </summary>
    public sealed class ComponentDisplayEvidence
    {
        internal ComponentDisplayEvidence(int rowIndex, ComponentSideDisplayEvidence source,
            ComponentSideDisplayEvidence target)
        {
            RowIndex = rowIndex;
            SourceDisplayEvidence = source;
            TargetDisplayEvidence = target;
        }

        public int RowIndex { get; }
        public ComponentSideDisplayEvidence SourceDisplayEvidence { get; }
        public ComponentSideDisplayEvidence TargetDisplayEvidence { get; }
    }
}
