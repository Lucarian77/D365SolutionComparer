using System;

namespace D365SolutionComparer.Models.Membership
{
    public enum IdentityResolutionStatus { Unresolved, Resolved, Unsupported, Ambiguous }

    /// <summary>A type-specific key plus the original record. Display names alone are not identity keys.</summary>
    public sealed class ComponentIdentity
    {
        public ComponentIdentity(SolutionComponentRecord record, IdentityResolutionStatus status,
            string comparisonKey = null, string diagnostic = null, string componentTypeKey = null)
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
            ComponentTypeKey = componentTypeKey ?? "componenttype:" + record.ComponentType.ToString(System.Globalization.CultureInfo.InvariantCulture);
            ComparisonKey = comparisonKey;
            Diagnostic = diagnostic ?? string.Empty;
        }

        public SolutionComponentRecord Record { get; }
        /// <summary>Stable resolver-defined kind; raw numeric codes can vary between environments.</summary>
        public string ComponentTypeKey { get; }
        public IdentityResolutionStatus Status { get; }
        public string ComparisonKey { get; }
        public string Diagnostic { get; }
    }
}
