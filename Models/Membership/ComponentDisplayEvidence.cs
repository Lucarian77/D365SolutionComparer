using System;

namespace D365SolutionComparer.Models.Membership
{
    internal enum MakerObjectKind { Dashboard, Process, ComponentLibrary, AiSkillConfig, SecuredMaskingRule }

    // Optional local display classification; never a portable identity or authoritative row.
    internal sealed class MakerObjectDisplayEvidence
    {
        internal MakerObjectDisplayEvidence(ComponentIdentity identity, MakerObjectKind kind, Guid canonicalId,
            ComponentDisplayValue friendly, ComponentDisplayValue internalName, DisplayBackingCorrelationStatus correlation,
            int? dashboardFormType = null, int? processCategory = null)
        {
            SolutionComponentId = identity.Record.SolutionComponentId; ObjectId = identity.Record.ObjectId;
            RawType = identity.Record.ComponentType; Kind = kind; CanonicalId = canonicalId;
            FriendlyName = friendly; InternalName = internalName; Correlation = correlation;
            DashboardFormType = kind == MakerObjectKind.Dashboard ? dashboardFormType : null;
            ProcessCategory = kind == MakerObjectKind.Process ? processCategory : null;
        }
        internal Guid SolutionComponentId { get; }
        internal Guid? ObjectId { get; }
        internal int RawType { get; }
        internal MakerObjectKind Kind { get; }
        internal Guid CanonicalId { get; }
        internal ComponentDisplayValue FriendlyName { get; }
        internal ComponentDisplayValue InternalName { get; }
        internal DisplayBackingCorrelationStatus Correlation { get; }
        internal int? DashboardFormType { get; }
        internal int? ProcessCategory { get; }
        private string ProcessType => ProcessCategory == 0 ? "Process (Workflow)" : ProcessCategory == 1 ? "Process (Dialog)" :
            ProcessCategory == 3 ? "Process (Action)" : ProcessCategory == 4 ? "Process (Business Process Flow)" : "Process";
        internal bool IsFor(ComponentIdentity identity) => identity != null && CanonicalId != Guid.Empty &&
            identity.Record.SolutionComponentId == SolutionComponentId && identity.Record.ObjectId == ObjectId &&
            identity.Record.ComponentType == RawType &&
            (CanonicalId == ObjectId || Kind == MakerObjectKind.Process && Correlation == DisplayBackingCorrelationStatus.VerifiedDefinition) &&
            (Kind == MakerObjectKind.Dashboard && RawType == 60 || Kind == MakerObjectKind.Process && RawType == 29 ||
             Kind == MakerObjectKind.ComponentLibrary && RawType == 300 ||
             identity.RegisteredDefinition != null && identity.RegisteredDefinition.ObjectTypeCode == RawType &&
             (Kind == MakerObjectKind.AiSkillConfig && string.Equals(identity.RegisteredDefinition.Name.Replace(" ", ""), "aiskillconfig", StringComparison.OrdinalIgnoreCase) ||
              Kind == MakerObjectKind.SecuredMaskingRule && string.Equals(identity.RegisteredDefinition.Name.Replace(" ", ""), "maskingrule", StringComparison.OrdinalIgnoreCase)));
        internal string Category => Kind == MakerObjectKind.Dashboard ? "Dashboards" : Kind == MakerObjectKind.Process ? "Processes" :
            Kind == MakerObjectKind.ComponentLibrary ? "Component libraries" : Kind == MakerObjectKind.AiSkillConfig ? "AI Skill Config" : "Secured Masking Rule";
        internal string Type => Kind == MakerObjectKind.Dashboard ? (DashboardFormType == 103 ? "Power BI Embedded" : "Dashboard") : Kind == MakerObjectKind.Process ? ProcessType :
            Kind == MakerObjectKind.ComponentLibrary ? "Component library" : Category;
    }

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
            ComponentDisplayValue parentDisplayContext, ComponentObjectMetadataEvidence objectMetadata = null,
            MakerObjectDisplayEvidence makerObject = null, MetadataForArchivalInformation archivalInformation = null)
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
            if (archivalInformation != null && (rawComponentType != 10139 || archivalInformation.Side != side ||
                archivalInformation.SolutionComponentId != solutionComponentId || archivalInformation.ObjectId != localObjectId))
                throw new ArgumentException("Archival information belongs to a different side or component.", nameof(archivalInformation));
            ArchivalInformation = archivalInformation;
            MakerObject = makerObject;
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
        internal MetadataForArchivalInformation ArchivalInformation { get; }
        internal MakerObjectDisplayEvidence MakerObject { get; }
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
