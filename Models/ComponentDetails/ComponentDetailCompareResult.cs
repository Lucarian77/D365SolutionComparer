using System;
using System.Collections.Generic;
using System.Linq;
using D365SolutionComparer.Models.Membership;

namespace D365SolutionComparer.Models.ComponentDetails
{
    internal enum ComponentDetailComparisonStatus
    {
        Match,
        Different,
        SourceOnly,
        TargetOnly,
        Unsupported,
        Unresolved,
        Ambiguous
    }

    internal sealed class ComponentPropertyDifference
    {
        public ComponentPropertyDifference(string propertyName, string sourceValue, string targetValue)
        {
            if (string.IsNullOrWhiteSpace(propertyName))
                throw new ArgumentException("A property name is required.", nameof(propertyName));
            PropertyName = propertyName;
            SourceValue = sourceValue;
            TargetValue = targetValue;
        }

        public string PropertyName { get; }
        public string SourceValue { get; }
        public string TargetValue { get; }
    }

    internal sealed class ComponentDetailCompareResult
    {
        public ComponentDetailCompareResult(MembershipCompareResult membership,
            ComponentDefinition source, ComponentDefinition target,
            ComponentDetailComparisonStatus status,
            IEnumerable<ComponentPropertyDifference> differences = null, string diagnostic = null)
        {
            Membership = membership ?? throw new ArgumentNullException(nameof(membership));
            if (!Enum.IsDefined(typeof(ComponentDetailComparisonStatus), status))
                throw new ArgumentOutOfRangeException(nameof(status));
            Source = source;
            Target = target;
            Status = status;
            Differences = new List<ComponentPropertyDifference>(differences ??
                Enumerable.Empty<ComponentPropertyDifference>()).AsReadOnly();
            Diagnostic = diagnostic ?? string.Empty;
        }

        public MembershipCompareResult Membership { get; }
        public ComponentDefinition Source { get; }
        public ComponentDefinition Target { get; }
        public ComponentDetailComparisonStatus Status { get; }
        public IReadOnlyList<ComponentPropertyDifference> Differences { get; }
        public string Diagnostic { get; }
    }
}
