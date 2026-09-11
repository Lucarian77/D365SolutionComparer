using System;
using System.Collections.Generic;

namespace D365SolutionComparer.Models.Membership
{
    public enum IdentityResolutionStatus { Unresolved, Resolved, Unsupported, Ambiguous }
    public enum ResolutionBlockerScope { None, PortableIdentity, SemanticKind }

    /// <summary>A type-specific key plus the original record. Display names alone are not identity keys.</summary>
    public sealed class ComponentIdentity
    {
        public ComponentIdentity(SolutionComponentRecord record, IdentityResolutionStatus status,
            string comparisonKey = null, string diagnostic = null, string componentTypeKey = null,
            string semanticKind = null, SolutionComponentDefinitionIdentity registeredDefinition = null,
            IEnumerable<string> diagnosticEvidence = null, string workflowCandidateKey = null,
            string blockerPortableIdentity = null, ResolutionBlockerScope blockerScope = ResolutionBlockerScope.None)
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
            WorkflowCandidateKey = workflowCandidateKey;
            BlockerPortableIdentity = blockerPortableIdentity;
            BlockerScope = blockerScope == ResolutionBlockerScope.None && status != IdentityResolutionStatus.Resolved
                ? (string.IsNullOrWhiteSpace(blockerPortableIdentity)
                    ? ResolutionBlockerScope.SemanticKind : ResolutionBlockerScope.PortableIdentity)
                : blockerScope;
            if (BlockerScope == ResolutionBlockerScope.PortableIdentity &&
                string.IsNullOrWhiteSpace(BlockerPortableIdentity))
                throw new ArgumentException("A portable-identity blocker requires its identity.", nameof(blockerPortableIdentity));
            if (status == IdentityResolutionStatus.Resolved && BlockerScope != ResolutionBlockerScope.None)
                throw new ArgumentException("Resolved identities cannot carry a resolution blocker.", nameof(blockerScope));
            ComparisonKey = comparisonKey;
            Diagnostic = diagnostic ?? string.Empty;
            RegisteredDefinition = registeredDefinition;
            DiagnosticEvidence = new List<string>(diagnosticEvidence ?? new string[0]).AsReadOnly();
        }

        public SolutionComponentRecord Record { get; }
        /// <summary>Stable resolver-defined kind; raw numeric codes can vary between environments.</summary>
        public string ComponentTypeKey { get; }
        /// <summary>Candidate coverage kind; null means the raw type cannot be classified safely.</summary>
        public string SemanticKind { get; }
        public IdentityResolutionStatus Status { get; }
        public string ComparisonKey { get; }
        /// <summary>Process-only semantic candidate evidence; never overrides an established uniquename.</summary>
        public string WorkflowCandidateKey { get; }
        public string BlockerPortableIdentity { get; }
        public ResolutionBlockerScope BlockerScope { get; }
        public string Diagnostic { get; }
        public SolutionComponentDefinitionIdentity RegisteredDefinition { get; }
        /// <summary>Per-record audit data that is never used as an identity or diagnostic grouping key.</summary>
        public IReadOnlyList<string> DiagnosticEvidence { get; }
    }
}
