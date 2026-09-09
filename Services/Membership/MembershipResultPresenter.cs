using System;
using System.Collections.Generic;
using System.Linq;
using D365SolutionComparer.Models.Membership;

namespace D365SolutionComparer.Services.Membership
{
    /// <summary>Creates display rows while preserving the comparison engine's absence safeguards.</summary>
    public sealed class MembershipResultPresenter
    {
        private readonly SolutionMembershipComparer comparer = new SolutionMembershipComparer();

        public MembershipComparisonPresentation Create(MembershipEnvironmentResult source,
            MembershipEnvironmentResult target)
        {
            if (source == null) throw new ArgumentNullException(nameof(source));
            if (target == null) throw new ArgumentNullException(nameof(target));
            if (!string.Equals(source.SolutionUniqueName, target.SolutionUniqueName, StringComparison.OrdinalIgnoreCase))
                throw new ArgumentException("Membership results must refer to the same solution Unique Name.");

            IReadOnlyList<MembershipCompareResult> comparisons;
            if (source.Snapshot != null && target.Snapshot != null)
                comparisons = comparer.Compare(source.Snapshot, target.Snapshot);
            else
                comparisons = CreateUnavailableComparisons(source.Snapshot, target.Snapshot);

            var rows = comparisons.Select(item => CreateRow(item, source, target)).ToList();
            return new MembershipComparisonPresentation(source.SolutionUniqueName, source, target, rows);
        }

        private static IReadOnlyList<MembershipCompareResult> CreateUnavailableComparisons(
            MembershipSnapshot source, MembershipSnapshot target)
        {
            var rows = new List<MembershipCompareResult>();
            if (source != null)
                rows.AddRange(MarkDuplicates(source.Components).Select(item =>
                    new MembershipCompareResult(item, null, MembershipPresence.Indeterminate)));
            if (target != null)
                rows.AddRange(MarkDuplicates(target.Components).Select(item =>
                    new MembershipCompareResult(null, item, MembershipPresence.Indeterminate)));
            return rows.OrderBy(item => (item.Source ?? item.Target).ComponentTypeKey, StringComparer.OrdinalIgnoreCase)
                .ThenBy(item => (item.Source ?? item.Target).ComparisonKey, StringComparer.OrdinalIgnoreCase)
                .ThenBy(item => (item.Source ?? item.Target).Record.SolutionComponentId).ToList().AsReadOnly();
        }

        private static IEnumerable<ComponentIdentity> MarkDuplicates(IEnumerable<ComponentIdentity> components)
        {
            var items = components.ToList();
            var duplicateKeys = new HashSet<string>(items
                .Where(item => item.Status == IdentityResolutionStatus.Resolved)
                .GroupBy(IdentityKey, StringComparer.OrdinalIgnoreCase)
                .Where(group => group.Count() > 1)
                .Select(group => group.Key), StringComparer.OrdinalIgnoreCase);
            return items.Select(item => item.Status == IdentityResolutionStatus.Resolved &&
                duplicateKeys.Contains(IdentityKey(item))
                ? new ComponentIdentity(item.Record, IdentityResolutionStatus.Ambiguous,
                    diagnostic: "Multiple membership records resolve to the same identity key: " + item.ComparisonKey,
                    componentTypeKey: item.ComponentTypeKey, semanticKind: item.SemanticKind)
                : item);
        }

        private static string IdentityKey(ComponentIdentity identity) => identity.ComponentTypeKey.Length + ":" +
            identity.ComponentTypeKey + identity.ComparisonKey;

        private static MembershipResultRow CreateRow(MembershipCompareResult item,
            MembershipEnvironmentResult source, MembershipEnvironmentResult target)
        {
            var identity = item.Source ?? item.Target;
            return new MembershipResultRow(item.Presence, item.Source, item.Target,
                DisplayKind(identity), DisplayPortableKey(identity),
                DisplayPresence(true, item, source), DisplayPresence(false, item, target),
                DisplayMembershipStatus(item, source, target),
                DisplayResolution(item.Source, source), DisplayResolution(item.Target, target),
                BuildDiagnostic(item, source, target));
        }

        private static string DisplayKind(ComponentIdentity identity)
        {
            switch (identity.ComponentTypeKey)
            {
                case "table": return "Table";
                case "column": return "Column";
                case "relationship": return "Relationship";
                case "webresource": return "Web Resource";
                case "process": return "Process / Workflow";
                case "securityrole": return "Security Role";
                case "environmentvariabledefinition": return "Environment Variable Definition";
                case "connectionreference": return "Connection Reference";
                case "globalchoice": return "Global Choice";
                case "appmodule": return "Model-driven App / AppModule";
                case "teamtemplate": return "Team Template";
                default: return "Component Type " + identity.Record.ComponentType;
            }
        }

        private static string DisplayPortableKey(ComponentIdentity identity) =>
            identity.Status == IdentityResolutionStatus.Resolved ? identity.ComparisonKey : "(identity not resolved)";

        private static string DisplayPresence(bool isSource, MembershipCompareResult item,
            MembershipEnvironmentResult side)
        {
            var identity = isSource ? item.Source : item.Target;
            if (identity != null) return "Present";
            if (side.State == MembershipSnapshotState.Unavailable) return "Unavailable";
            if ((isSource && item.Presence == MembershipPresence.OnlyInTarget) ||
                (!isSource && item.Presence == MembershipPresence.OnlyInSource)) return "Missing";
            return "Indeterminate";
        }

        private static string DisplayMembershipStatus(MembershipCompareResult item,
            MembershipEnvironmentResult source, MembershipEnvironmentResult target)
        {
            if (source.State == MembershipSnapshotState.Unavailable || target.State == MembershipSnapshotState.Unavailable)
                return "Indeterminate - Environment Unavailable";
            switch (item.Presence)
            {
                case MembershipPresence.PresentInBoth: return "Present in Both";
                case MembershipPresence.OnlyInSource: return "Source Only";
                case MembershipPresence.OnlyInTarget: return "Target Only";
            }
            var identity = item.Source ?? item.Target;
            switch (identity.Status)
            {
                case IdentityResolutionStatus.Unsupported: return "Indeterminate - Unsupported";
                case IdentityResolutionStatus.Ambiguous: return "Indeterminate - Ambiguous";
                case IdentityResolutionStatus.Unresolved: return "Indeterminate - Unresolved";
                default: return "Indeterminate";
            }
        }

        private static string DisplayResolution(ComponentIdentity identity, MembershipEnvironmentResult side)
        {
            if (identity != null) return identity.Status.ToString();
            if (side.State == MembershipSnapshotState.Unavailable) return "Unavailable";
            return "Not present";
        }

        private static string BuildDiagnostic(MembershipCompareResult item,
            MembershipEnvironmentResult source, MembershipEnvironmentResult target)
        {
            var messages = new List<string>();
            Add(messages, item.Source?.Diagnostic);
            Add(messages, item.Target?.Diagnostic);
            if (source.State == MembershipSnapshotState.Unavailable)
                Add(messages, "Source retrieval unavailable: " + source.Diagnostics.Diagnostic);
            if (target.State == MembershipSnapshotState.Unavailable)
                Add(messages, "Target retrieval unavailable: " + target.Diagnostics.Diagnostic);
            if (item.AbsenceEvidence == MembershipAbsenceEvidence.OppositeSolutionAbsent)
                Add(messages, "Missing is established because the opposite solution is absent.");
            else if (item.AbsenceEvidence == MembershipAbsenceEvidence.CompleteResolvedInventory)
                Add(messages, "Missing is established from complete identity coverage for the opposite " +
                    DisplaySemanticKind(item.Source ?? item.Target) + " component kind.");
            else if (item.Presence == MembershipPresence.Indeterminate &&
                (item.Source ?? item.Target).Status == IdentityResolutionStatus.Resolved &&
                source.State != MembershipSnapshotState.Unavailable && target.State != MembershipSnapshotState.Unavailable)
                Add(messages, "Absence is not established because the opposite " +
                    DisplaySemanticKind(item.Source ?? item.Target) + " component kind is not fully resolved.");
            return string.Join(" ", messages.Distinct(StringComparer.Ordinal));
        }

        private static string DisplaySemanticKind(ComponentIdentity identity)
        {
            switch (identity.SemanticKind)
            {
                case ComponentSemanticKinds.Table: return "Table";
                case ComponentSemanticKinds.Column: return "Column";
                case ComponentSemanticKinds.Relationship: return "Relationship";
                case ComponentSemanticKinds.WebResource: return "Web Resource";
                case ComponentSemanticKinds.Process: return "Process / Workflow";
                case ComponentSemanticKinds.SecurityRole: return "Security Role";
                case ComponentSemanticKinds.EnvironmentVariableDefinition: return "Environment Variable Definition";
                case ComponentSemanticKinds.ConnectionReference: return "Connection Reference";
                case ComponentSemanticKinds.GlobalChoice: return "Global Choice";
                case ComponentSemanticKinds.AppModule: return "Model-driven App / AppModule";
                case ComponentSemanticKinds.TeamTemplate: return "Team Template";
                default: return DisplayKind(identity);
            }
        }

        private static void Add(ICollection<string> messages, string message)
        {
            if (!string.IsNullOrWhiteSpace(message)) messages.Add(message.Trim());
        }
    }
}
