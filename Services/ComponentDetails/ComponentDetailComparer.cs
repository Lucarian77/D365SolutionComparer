using System;
using System.Collections.Generic;
using System.Linq;
using D365SolutionComparer.Models.ComponentDetails;
using D365SolutionComparer.Models.Membership;
using D365SolutionComparer.Services.Contracts;

namespace D365SolutionComparer.Services.ComponentDetails
{
    /// <summary>
    /// Adds definition equality to authoritative membership results. It never re-evaluates
    /// portable identity, duplicate identity, or absence evidence.
    /// </summary>
    internal sealed class ComponentDetailComparer : IComponentDetailComparer
    {
        public IReadOnlyList<ComponentDetailCompareResult> Compare(
            IReadOnlyList<MembershipCompareResult> membershipResults,
            ComponentDefinitionSnapshot source, ComponentDefinitionSnapshot target)
        {
            if (membershipResults == null) throw new ArgumentNullException(nameof(membershipResults));
            if (source == null) throw new ArgumentNullException(nameof(source));
            if (target == null) throw new ArgumentNullException(nameof(target));
            if (!string.Equals(source.Membership.SolutionUniqueName,
                target.Membership.SolutionUniqueName, StringComparison.OrdinalIgnoreCase))
                throw new ArgumentException("Definition snapshots must refer to the same solution Unique Name.");

            var sourceDefinitions = Index(source.Definitions);
            var targetDefinitions = Index(target.Definitions);
            var results = membershipResults.Select(item => CompareOne(item,
                Find(sourceDefinitions, item.Source), Find(targetDefinitions, item.Target))).ToList();
            return results.AsReadOnly();
        }

        private static ComponentDetailCompareResult CompareOne(MembershipCompareResult membership,
            ComponentDefinition source, ComponentDefinition target)
        {
            if (membership.Presence == MembershipPresence.OnlyInSource)
                return new ComponentDetailCompareResult(membership, source, null,
                    ComponentDetailComparisonStatus.SourceOnly);
            if (membership.Presence == MembershipPresence.OnlyInTarget)
                return new ComponentDetailCompareResult(membership, null, target,
                    ComponentDetailComparisonStatus.TargetOnly);
            if (membership.Presence != MembershipPresence.PresentInBoth)
                return Conservative(membership, source, target,
                    "Membership does not provide definitive shared presence or absence.");

            if (source == null || target == null)
                return new ComponentDetailCompareResult(membership, source, target,
                    ComponentDetailComparisonStatus.Unresolved,
                    diagnostic: "A component definition was not retrieved for both sides.");
            if (source.Status != ComponentDefinitionReadStatus.Available ||
                target.Status != ComponentDefinitionReadStatus.Available)
                return Conservative(membership, source, target,
                    "One or both component definitions are not complete and uniquely correlated.");

            var differences = CompareProperties(source.ComparableProperties,
                target.ComparableProperties);
            return new ComponentDetailCompareResult(membership, source, target,
                differences.Count == 0 ? ComponentDetailComparisonStatus.Match :
                    ComponentDetailComparisonStatus.Different, differences);
        }

        private static ComponentDetailCompareResult Conservative(MembershipCompareResult membership,
            ComponentDefinition source, ComponentDefinition target, string diagnostic)
        {
            var statuses = new[] { source?.Status, target?.Status };
            var identities = new[] { membership.Source?.Status, membership.Target?.Status };
            ComponentDetailComparisonStatus status;
            if (statuses.Contains(ComponentDefinitionReadStatus.Ambiguous) ||
                identities.Contains(IdentityResolutionStatus.Ambiguous))
                status = ComponentDetailComparisonStatus.Ambiguous;
            else if (statuses.Contains(ComponentDefinitionReadStatus.Unsupported) ||
                identities.Contains(IdentityResolutionStatus.Unsupported))
                status = ComponentDetailComparisonStatus.Unsupported;
            else status = ComponentDetailComparisonStatus.Unresolved;
            return new ComponentDetailCompareResult(membership, source, target, status,
                diagnostic: diagnostic);
        }

        private static IReadOnlyList<ComponentPropertyDifference> CompareProperties(
            IReadOnlyDictionary<string, string> source,
            IReadOnlyDictionary<string, string> target)
        {
            var names = source.Keys.Concat(target.Keys).Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(item => item, StringComparer.OrdinalIgnoreCase);
            var differences = new List<ComponentPropertyDifference>();
            foreach (var name in names)
            {
                string sourceValue;
                string targetValue;
                source.TryGetValue(name, out sourceValue);
                target.TryGetValue(name, out targetValue);
                if (!string.Equals(sourceValue, targetValue, StringComparison.Ordinal))
                    differences.Add(new ComponentPropertyDifference(name, sourceValue, targetValue));
            }
            return differences.AsReadOnly();
        }

        private static IDictionary<Guid, List<ComponentDefinition>> Index(
            IEnumerable<ComponentDefinition> definitions) => definitions
                .GroupBy(item => item.Identity.Record.SolutionComponentId)
                .ToDictionary(group => group.Key, group => group.ToList());

        private static ComponentDefinition Find(IDictionary<Guid, List<ComponentDefinition>> index,
            ComponentIdentity identity)
        {
            if (identity == null) return null;
            List<ComponentDefinition> matches;
            if (!index.TryGetValue(identity.Record.SolutionComponentId, out matches) || matches.Count == 0)
                return null;
            if (matches.Count == 1) return matches[0];
            return new ComponentDefinition(identity, ComponentDefinitionReadStatus.Ambiguous,
                diagnostic: "Multiple definition records refer to the same membership record.");
        }
    }
}
