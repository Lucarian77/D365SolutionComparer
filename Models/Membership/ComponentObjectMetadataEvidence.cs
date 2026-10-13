using System;

namespace D365SolutionComparer.Models.Membership
{
    /// <summary>A detached observation. Requestedness is evidence, not a retrieval instruction.</summary>
    public sealed class ComponentObjectMetadataValue<T>
    {
        public ComponentObjectMetadataValue(DisplayAvailability availability, T value = default(T),
            string provenance = null, bool wasRequested = false)
        {
            if (!Enum.IsDefined(typeof(DisplayAvailability), availability))
                throw new ArgumentOutOfRangeException(nameof(availability));
            if (availability != DisplayAvailability.Available && !Equals(value, default(T)))
                throw new ArgumentException("Unestablished evidence cannot carry a value.", nameof(value));
            if (availability == DisplayAvailability.Available &&
                (!wasRequested || string.IsNullOrWhiteSpace(provenance)))
                throw new ArgumentException("An observed value requires requestedness and provenance.");
            if ((availability == DisplayAvailability.AttributeOmitted ||
                availability == DisplayAvailability.RetrievalFailed) && !wasRequested)
                throw new ArgumentException("Omitted or failed retrieval requires requestedness.");
            if (availability == DisplayAvailability.NotRetrieved && wasRequested)
                throw new ArgumentException("NotRetrieved cannot claim a requested property.");
            if (availability == DisplayAvailability.NotApplicable && string.IsNullOrWhiteSpace(provenance))
                throw new ArgumentException("NotApplicable requires applicability evidence.");
            Availability = availability;
            Value = value;
            Provenance = provenance;
            WasRequested = wasRequested;
        }

        public DisplayAvailability Availability { get; }
        public T Value { get; }
        public string Provenance { get; }
        public bool WasRequested { get; }
    }

    public sealed class ComponentTimestampEvidence
    {
        public ComponentTimestampEvidence(DateTime rawTimestamp, string behavior = null,
            string timezoneProvenance = null, DateTimeOffset? verifiedInstant = null)
        {
            if (verifiedInstant.HasValue && string.IsNullOrWhiteSpace(timezoneProvenance))
                throw new ArgumentException("An interpreted instant requires independent timezone provenance.");
            RawTimestamp = rawTimestamp;
            Behavior = behavior;
            TimezoneProvenance = timezoneProvenance;
            VerifiedInstant = verifiedInstant;
        }

        public DateTime RawTimestamp { get; }
        public DateTimeKind Kind => RawTimestamp.Kind;
        public string Behavior { get; }
        public string TimezoneProvenance { get; }
        public DateTimeOffset? VerifiedInstant { get; }
    }

    public sealed class ComponentOwnerEvidence
    {
        public ComponentOwnerEvidence(string lookupLogicalTable, Guid? localOwnerId,
            string displayName = null, bool displayNameVerified = false,
            DisplayBackingCorrelationStatus correlation = DisplayBackingCorrelationStatus.NotAttempted)
        {
            if (localOwnerId == Guid.Empty || localOwnerId.HasValue && string.IsNullOrWhiteSpace(lookupLogicalTable))
                throw new ArgumentException("A local owner locator requires its logical table and nonblank ID.");
            if (!displayNameVerified && displayName != null)
                throw new ArgumentException("An owner display name requires independent verification.");
            if (!Enum.IsDefined(typeof(DisplayBackingCorrelationStatus), correlation))
                throw new ArgumentOutOfRangeException(nameof(correlation));
            LookupLogicalTable = lookupLogicalTable;
            LocalOwnerId = localOwnerId;
            DisplayName = displayName;
            DisplayNameVerified = displayNameVerified;
            Correlation = correlation;
        }

        public string LookupLogicalTable { get; }
        public Guid? LocalOwnerId { get; }
        public string DisplayName { get; }
        public bool DisplayNameVerified { get; }
        public DisplayBackingCorrelationStatus Correlation { get; }
    }

    public sealed class ComponentObjectStatusEvidence
    {
        public ComponentObjectStatusEvidence(string sourceField, int? rawNumericValue,
            string displayLabel = null, bool labelVerified = false, string interpretation = null)
        {
            if (string.IsNullOrWhiteSpace(sourceField))
                throw new ArgumentException("Object state requires field provenance.", nameof(sourceField));
            if (string.Equals(sourceField, "componentstate", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(sourceField, "MembershipStatus", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(sourceField, "DefinitionStatus", StringComparison.OrdinalIgnoreCase))
                throw new ArgumentException("Comparison or component-layer state is not object activation.");
            if (!labelVerified && displayLabel != null)
                throw new ArgumentException("An object state label requires independent verification.");
            SourceField = sourceField;
            RawNumericValue = rawNumericValue;
            DisplayLabel = displayLabel;
            LabelVerified = labelVerified;
            Interpretation = interpretation;
        }

        public string SourceField { get; }
        public int? RawNumericValue { get; }
        public string DisplayLabel { get; }
        public bool LabelVerified { get; }
        public string Interpretation { get; }
    }

    /// <summary>Immutable, per-environment object observations, separate from comparison authority.</summary>
    public sealed class ComponentObjectMetadataEvidence
    {
        private static readonly ComponentObjectMetadataEvidence SourceNotRetrieved =
            new ComponentObjectMetadataEvidence(DisplayEnvironmentSide.Source);
        private static readonly ComponentObjectMetadataEvidence TargetNotRetrieved =
            new ComponentObjectMetadataEvidence(DisplayEnvironmentSide.Target);

        public ComponentObjectMetadataEvidence(DisplayEnvironmentSide environmentSide,
            ComponentObjectMetadataValue<bool?> managed = null,
            ComponentObjectMetadataValue<bool?> customized = null,
            ComponentObjectMetadataValue<ComponentTimestampEvidence> lastModified = null,
            ComponentObjectMetadataValue<ComponentOwnerEvidence> owner = null,
            ComponentObjectMetadataValue<ComponentObjectStatusEvidence> objectStatus = null)
        {
            if (!Enum.IsDefined(typeof(DisplayEnvironmentSide), environmentSide))
                throw new ArgumentOutOfRangeException(nameof(environmentSide));
            EnvironmentSide = environmentSide;
            Managed = managed ?? new ComponentObjectMetadataValue<bool?>(DisplayAvailability.NotRetrieved);
            Customized = customized ?? new ComponentObjectMetadataValue<bool?>(DisplayAvailability.NotRetrieved);
            LastModified = lastModified ?? new ComponentObjectMetadataValue<ComponentTimestampEvidence>(DisplayAvailability.NotRetrieved);
            Owner = owner ?? new ComponentObjectMetadataValue<ComponentOwnerEvidence>(DisplayAvailability.NotRetrieved);
            ObjectStatus = objectStatus ?? new ComponentObjectMetadataValue<ComponentObjectStatusEvidence>(DisplayAvailability.NotRetrieved);
        }

        public DisplayEnvironmentSide EnvironmentSide { get; }
        public ComponentObjectMetadataValue<bool?> Managed { get; }
        public ComponentObjectMetadataValue<bool?> Customized { get; }
        public ComponentObjectMetadataValue<ComponentTimestampEvidence> LastModified { get; }
        public ComponentObjectMetadataValue<ComponentOwnerEvidence> Owner { get; }
        public ComponentObjectMetadataValue<ComponentObjectStatusEvidence> ObjectStatus { get; }

        public static ComponentObjectMetadataEvidence NotRetrieved(DisplayEnvironmentSide side)
        {
            if (!Enum.IsDefined(typeof(DisplayEnvironmentSide), side))
                throw new ArgumentOutOfRangeException(nameof(side));
            return side == DisplayEnvironmentSide.Source ? SourceNotRetrieved : TargetNotRetrieved;
        }
    }
}
