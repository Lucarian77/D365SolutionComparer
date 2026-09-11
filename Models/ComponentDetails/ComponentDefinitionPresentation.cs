using System;
using System.Collections.Generic;
using System.Linq;
using D365SolutionComparer.Models.Membership;

namespace D365SolutionComparer.Models.ComponentDetails
{
    internal sealed class ComponentPropertyPresentation
    {
        public ComponentPropertyPresentation(string propertyName, string sourceValue,
            string targetValue, bool changed, string comparison)
        {
            if (string.IsNullOrWhiteSpace(propertyName))
                throw new ArgumentException("A property name is required.", nameof(propertyName));
            PropertyName = propertyName;
            SourceValue = sourceValue;
            TargetValue = targetValue;
            Changed = changed;
            Comparison = comparison ?? string.Empty;
        }

        public string PropertyName { get; }
        public string SourceValue { get; }
        public string TargetValue { get; }
        public bool Changed { get; }
        public string Comparison { get; }
    }

    internal sealed class ComponentDefinitionDetailPresentation
    {
        public ComponentDefinitionDetailPresentation(string componentKind, string portableKey,
            string membershipStatus, string definitionStatus,
            IEnumerable<ComponentPropertyPresentation> properties, string diagnostic,
            IEnumerable<string> sourceEvidence = null, IEnumerable<string> targetEvidence = null)
        {
            ComponentKind = componentKind ?? string.Empty;
            PortableKey = portableKey ?? string.Empty;
            MembershipStatus = membershipStatus ?? string.Empty;
            DefinitionStatus = definitionStatus ?? string.Empty;
            Properties = new List<ComponentPropertyPresentation>(properties ??
                Enumerable.Empty<ComponentPropertyPresentation>()).AsReadOnly();
            Diagnostic = diagnostic ?? string.Empty;
            SourceEvidence = new List<string>(sourceEvidence ??
                Enumerable.Empty<string>()).AsReadOnly();
            TargetEvidence = new List<string>(targetEvidence ??
                Enumerable.Empty<string>()).AsReadOnly();
        }

        public string ComponentKind { get; }
        public string PortableKey { get; }
        public string MembershipStatus { get; }
        public string DefinitionStatus { get; }
        public IReadOnlyList<ComponentPropertyPresentation> Properties { get; }
        public string Diagnostic { get; }
        public IReadOnlyList<string> SourceEvidence { get; }
        public IReadOnlyList<string> TargetEvidence { get; }
    }
}
