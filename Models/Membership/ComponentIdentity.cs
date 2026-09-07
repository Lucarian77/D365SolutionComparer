using System;

namespace D365SolutionComparer.Models.Membership
{
    public enum IdentityResolutionStatus { Unresolved, Resolved, Unsupported, Ambiguous }

    /// <summary>A type-specific key plus the original record. Display names alone are not identity keys.</summary>
    public sealed class ComponentIdentity
    {
        public ComponentIdentity(SolutionComponentRecord record, IdentityResolutionStatus status,
            string comparisonKey = null, string diagnostic = null, string componentTypeKey = null,
            string semanticKind = null, SolutionComponentDefinitionIdentity registeredDefinition = null)
        {
            Record = record ?? throw new ArgumentNullException(nameof(record));
            if (!Enum.IsDefined(typeof(IdentityResolutionStatus), status)) throw new ArgumentOutOfRangeException(nameof(status));
            if (status == IdentityResolutionStatus.Resolved && string.IsNullOrWhiteSpace(comparisonKey))
                throw new ArgumentException("Resolved identities require a type-specific key.", nameof(comparisonKey));
            if (status != IdentityResolutionStatus.Resolved && comparisonKey != null)
                throw new ArgumentException("Unresolved identities cannot carry a trusted comparison key.", nameof(comparisonKey));
            Status = status;
            if (componentTypeKey != null && string.IsNullOrWhiteSpace(componentTypeKey))
                throw new ArgumentException("A canonical component type key cannot be blank.", nameof(componentTypeKey));
            if (semanticKind != null && string.IsNullOrWhiteSpace(semanticKind))
                throw new ArgumentException("A semantic component kind cannot be blank.", nameof(semanticKind));
            if (registeredDefinition != null && registeredDefinition.ObjectTypeCode != record.ComponentType)
                throw new ArgumentException("The registered definition belongs to a different raw component type.",
                    nameof(registeredDefinition));
            if (registeredDefinition != null && semanticKind != null &&
                !string.Equals(registeredDefinition.SemanticKind, semanticKind, StringComparison.OrdinalIgnoreCase))
                throw new ArgumentException("The semantic kind does not match the registered definition.",
                    nameof(semanticKind));
            ComponentTypeKey = componentTypeKey ?? "componenttype:" + record.ComponentType.ToString(System.Globalization.CultureInfo.InvariantCulture);
            SemanticKind = semanticKind ?? registeredDefinition?.SemanticKind ??
                ComponentSemanticKinds.FromCanonicalTypeKey(componentTypeKey) ??
                ComponentSemanticKinds.FromRawComponentType(record.ComponentType);
            ComparisonKey = comparisonKey;
            Diagnostic = diagnostic ?? string.Empty;
            RegisteredDefinition = registeredDefinition;
        }

        public SolutionComponentRecord Record { get; }
        /// <summary>Stable resolver-defined kind; raw numeric codes can vary between environments.</summary>
        public string ComponentTypeKey { get; }
        /// <summary>Candidate coverage kind; null means the raw type cannot be classified safely.</summary>
        public string SemanticKind { get; }
        public IdentityResolutionStatus Status { get; }
        public string ComparisonKey { get; }
        public string Diagnostic { get; }
        public SolutionComponentDefinitionIdentity RegisteredDefinition { get; }
    }
}
