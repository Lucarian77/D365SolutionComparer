using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using D365SolutionComparer.Models.Membership;

namespace D365SolutionComparer.Models.ComponentDetails
{
    internal enum ComponentDefinitionReadStatus
    {
        Available,
        Unsupported,
        Unresolved,
        Ambiguous
    }

    /// <summary>
    /// A read-only projection of portable component configuration. Environment-local identifiers
    /// belong in DiagnosticEvidence and are never part of equality.
    /// </summary>
    internal sealed class ComponentDefinition
    {
        public ComponentDefinition(ComponentIdentity identity, ComponentDefinitionReadStatus status,
            IEnumerable<KeyValuePair<string, string>> comparableProperties = null,
            string diagnostic = null, IEnumerable<string> diagnosticEvidence = null)
        {
            Identity = identity ?? throw new ArgumentNullException(nameof(identity));
            if (!Enum.IsDefined(typeof(ComponentDefinitionReadStatus), status))
                throw new ArgumentOutOfRangeException(nameof(status));
            var properties = (comparableProperties ?? Enumerable.Empty<KeyValuePair<string, string>>()).ToList();
            if (properties.Any(item => string.IsNullOrWhiteSpace(item.Key)))
                throw new ArgumentException("Comparable property names cannot be blank.", nameof(comparableProperties));
            if (properties.GroupBy(item => item.Key, StringComparer.OrdinalIgnoreCase).Any(group => group.Count() > 1))
                throw new ArgumentException("Comparable property names must be unique.", nameof(comparableProperties));
            if (status == ComponentDefinitionReadStatus.Available &&
                identity.Status != IdentityResolutionStatus.Resolved)
                throw new ArgumentException("Only a resolved component identity can have an available definition.",
                    nameof(identity));
            if (status == ComponentDefinitionReadStatus.Available)
            {
                var contract = ComponentDefinitionContractCatalog.For(identity.SemanticKind);
                if (contract == null)
                    throw new ArgumentException("No definition equality contract exists for this component kind.",
                        nameof(identity));
                var supplied = new HashSet<string>(properties.Select(item => item.Key),
                    StringComparer.OrdinalIgnoreCase);
                if (supplied.Count != contract.ComparableProperties.Count ||
                    contract.ComparableProperties.Any(item => !supplied.Contains(item)))
                    throw new ArgumentException("An available definition must contain every property in its family equality contract.",
                        nameof(comparableProperties));
            }

            Status = status;
            ComparableProperties = new ReadOnlyDictionary<string, string>(properties
                .OrderBy(item => item.Key, StringComparer.OrdinalIgnoreCase)
                .ToDictionary(item => item.Key, item => item.Value, StringComparer.OrdinalIgnoreCase));
            Diagnostic = diagnostic ?? string.Empty;
            DiagnosticEvidence = new List<string>(diagnosticEvidence ?? Enumerable.Empty<string>()).AsReadOnly();
        }

        public ComponentIdentity Identity { get; }
        public ComponentDefinitionReadStatus Status { get; }
        public IReadOnlyDictionary<string, string> ComparableProperties { get; }
        public string Diagnostic { get; }
        public IReadOnlyList<string> DiagnosticEvidence { get; }
    }
}
