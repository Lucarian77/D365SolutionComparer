using System;
using System.Collections.Generic;
using System.Linq;
using D365SolutionComparer.Models.ComponentDetails;
using D365SolutionComparer.Models.Membership;

namespace D365SolutionComparer.Services.ComponentDetails
{
    /// <summary>
    /// Maps definition comparisons onto membership rows without changing authoritative
    /// membership presence, resolution, diagnostics, or summary counts.
    /// </summary>
    internal sealed class ComponentDefinitionResultPresenter
    {
        private readonly ComponentDetailComparer comparer = new ComponentDetailComparer();

        public MembershipComparisonPresentation Apply(MembershipComparisonPresentation membership,
            ComponentDefinitionSnapshot source, ComponentDefinitionSnapshot target)
        {
            if (membership == null) throw new ArgumentNullException(nameof(membership));
            var comparisonRows = membership.Rows.Select(item => item.Comparison).ToList();
            if (comparisonRows.Any(item => item == null))
                throw new ArgumentException("Membership rows do not retain their authoritative comparison.",
                    nameof(membership));

            IReadOnlyList<ComponentDetailCompareResult> details;
            if (source != null && target != null)
                details = comparer.Compare(comparisonRows, source, target);
            else
                details = comparisonRows.Select(item => MissingEnvironmentResult(item,
                    source, target)).ToList().AsReadOnly();

            var rows = membership.Rows.Zip(details, (row, detail) =>
            {
                var detailPresentation = CreateDetail(row, detail);
                return row.WithDefinition(DisplayStatus(detail.Status),
                    DisplayChangedProperties(detail), detailPresentation);
            }).ToList();
            return new MembershipComparisonPresentation(membership.SolutionUniqueName,
                membership.Source, membership.Target, rows);
        }

        private static ComponentDetailCompareResult MissingEnvironmentResult(
            MembershipCompareResult membership, ComponentDefinitionSnapshot source,
            ComponentDefinitionSnapshot target)
        {
            var sourceDefinition = Find(source, membership.Source);
            var targetDefinition = Find(target, membership.Target);
            var identities = new[] { membership.Source?.Status, membership.Target?.Status };
            var status = identities.Contains(IdentityResolutionStatus.Ambiguous)
                ? ComponentDetailComparisonStatus.Ambiguous
                : identities.Contains(IdentityResolutionStatus.Unsupported)
                    ? ComponentDetailComparisonStatus.Unsupported
                    : ComponentDetailComparisonStatus.Unresolved;
            return new ComponentDetailCompareResult(membership, sourceDefinition,
                targetDefinition, status,
                diagnostic: "Definition comparison is incomplete because an environment retrieval is unavailable.");
        }

        private static ComponentDefinition Find(ComponentDefinitionSnapshot snapshot,
            ComponentIdentity identity)
        {
            if (snapshot == null || identity == null) return null;
            var matches = snapshot.Definitions.Where(item =>
                item.Identity.Record.SolutionComponentId == identity.Record.SolutionComponentId).ToList();
            return matches.Count == 1 ? matches[0] : null;
        }

        private static ComponentDefinitionDetailPresentation CreateDetail(MembershipResultRow row,
            ComponentDetailCompareResult detail)
        {
            var contract = ComponentDefinitionContractCatalog.For(
                (detail.Membership.Source ?? detail.Membership.Target).SemanticKind);
            var names = contract == null
                ? detail.Source?.ComparableProperties.Keys.Concat(
                    detail.Target?.ComparableProperties.Keys ?? Enumerable.Empty<string>()) ??
                    Enumerable.Empty<string>()
                : contract.ComparableProperties;
            var changed = new HashSet<string>(detail.Differences.Select(item => item.PropertyName),
                StringComparer.OrdinalIgnoreCase);
            var properties = names.Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(item => item, StringComparer.OrdinalIgnoreCase)
                .Select(name => new ComponentPropertyPresentation(name,
                    Value(detail.Source, name), Value(detail.Target, name), changed.Contains(name),
                    PropertyStatus(detail.Status, changed.Contains(name))))
                .ToList();
            return new ComponentDefinitionDetailPresentation(row.ComponentKind, row.PortableKey,
                row.MembershipStatus, DisplayStatus(detail.Status), properties,
                BuildDiagnostic(detail), detail.Source?.DiagnosticEvidence,
                detail.Target?.DiagnosticEvidence);
        }

        private static string Value(ComponentDefinition definition, string name)
        {
            string value;
            return definition != null && definition.ComparableProperties.TryGetValue(name, out value)
                ? value : null;
        }

        private static string BuildDiagnostic(ComponentDetailCompareResult detail)
        {
            var messages = new[] { detail.Diagnostic, detail.Source?.Diagnostic,
                detail.Target?.Diagnostic }.Where(item => !string.IsNullOrWhiteSpace(item));
            return string.Join(" ", messages.Distinct(StringComparer.Ordinal));
        }

        private static string DisplayChangedProperties(ComponentDetailCompareResult detail) =>
            detail.Status == ComponentDetailComparisonStatus.Different
                ? string.Join(", ", detail.Differences.Select(item => item.PropertyName)
                    .OrderBy(item => item, StringComparer.OrdinalIgnoreCase))
                : string.Empty;

        private static string PropertyStatus(ComponentDetailComparisonStatus status, bool changed)
        {
            if (status == ComponentDetailComparisonStatus.Different)
                return changed ? "Different" : "Match";
            if (status == ComponentDetailComparisonStatus.Match) return "Match";
            if (status == ComponentDetailComparisonStatus.SourceOnly) return "Source only";
            if (status == ComponentDetailComparisonStatus.TargetOnly) return "Target only";
            return "Not compared";
        }

        private static string DisplayStatus(ComponentDetailComparisonStatus status)
        {
            switch (status)
            {
                case ComponentDetailComparisonStatus.SourceOnly: return "SourceOnly";
                case ComponentDetailComparisonStatus.TargetOnly: return "TargetOnly";
                default: return status.ToString();
            }
        }
    }
}
