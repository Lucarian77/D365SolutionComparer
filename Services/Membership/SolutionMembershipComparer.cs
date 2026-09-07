using System;
using System.Collections.Generic;
using System.Linq;
using D365SolutionComparer.Models.Membership;
using D365SolutionComparer.Services.Contracts;

namespace D365SolutionComparer.Services.Membership
{
    /// <summary>Pure presence comparison. Unknown coverage blocks absence, not independently proven matches.</summary>
    public sealed class SolutionMembershipComparer : ISolutionMembershipComparer
    {
        public IReadOnlyList<MembershipCompareResult> Compare(MembershipSnapshot source, MembershipSnapshot target)
        {
            if (source == null) throw new ArgumentNullException(nameof(source));
            if (target == null) throw new ArgumentNullException(nameof(target));
            if (!string.Equals(source.SolutionUniqueName, target.SolutionUniqueName, StringComparison.OrdinalIgnoreCase))
                throw new ArgumentException("Membership snapshots must refer to the same solution Unique Name.");
            var sourceItems = MarkDuplicates(source.Components);
            var targetItems = MarkDuplicates(target.Components);
            var sourceCoverage = IdentityCoverage.From(sourceItems);
            var targetCoverage = IdentityCoverage.From(targetItems);
            var targetLookup = targetItems.Where(IsResolved).ToDictionary(Key, StringComparer.OrdinalIgnoreCase);
            var usedTargets = new HashSet<ComponentIdentity>();
            var results = new List<MembershipCompareResult>();
            foreach (var item in sourceItems)
            {
                ComponentIdentity match;
                if (IsResolved(item) && targetLookup.TryGetValue(Key(item), out match))
                {
                    usedTargets.Add(match);
                    results.Add(new MembershipCompareResult(item, match, MembershipPresence.PresentInBoth));
                }
                else results.Add(Unmatched(item, true, target.State, targetCoverage));
            }
            foreach (var item in targetItems.Where(i => !usedTargets.Contains(i)))
                results.Add(Unmatched(item, false, source.State, sourceCoverage));
            return results.OrderBy(r => (r.Source ?? r.Target).ComponentTypeKey, StringComparer.OrdinalIgnoreCase)
                .ThenBy(r => (r.Source ?? r.Target).ComparisonKey, StringComparer.OrdinalIgnoreCase)
                .ThenBy(r => (r.Source ?? r.Target).Record.SolutionComponentId).ToList().AsReadOnly();
        }

        private static MembershipCompareResult Unmatched(ComponentIdentity item, bool source,
            MembershipSnapshotState oppositeState, IdentityCoverage oppositeCoverage)
        {
            var evidence = MembershipAbsenceEvidence.None;
            if (IsResolved(item))
            {
                if (oppositeState == MembershipSnapshotState.SolutionAbsent)
                    evidence = MembershipAbsenceEvidence.OppositeSolutionAbsent;
                else if (oppositeState == MembershipSnapshotState.Complete &&
                    oppositeCoverage.IsComplete(item.SemanticKind))
                    evidence = MembershipAbsenceEvidence.CompleteResolvedInventory;
            }
            var presence = evidence == MembershipAbsenceEvidence.None ? MembershipPresence.Indeterminate
                : source ? MembershipPresence.OnlyInSource : MembershipPresence.OnlyInTarget;
            return new MembershipCompareResult(source ? item : null, source ? null : item, presence, evidence);
        }

        private static IReadOnlyList<ComponentIdentity> MarkDuplicates(IReadOnlyList<ComponentIdentity> items)
        {
            var duplicates = new HashSet<string>(items.Where(IsResolved).GroupBy(Key, StringComparer.OrdinalIgnoreCase)
                .Where(g => g.Count() > 1).Select(g => g.Key), StringComparer.OrdinalIgnoreCase);
            return items.Select(i => IsResolved(i) && duplicates.Contains(Key(i))
                ? new ComponentIdentity(i.Record, IdentityResolutionStatus.Ambiguous,
                    diagnostic: "Multiple membership records resolve to the same identity key: " + i.ComparisonKey,
                    componentTypeKey: i.ComponentTypeKey, semanticKind: i.SemanticKind) : i).ToList().AsReadOnly();
        }

        private sealed class IdentityCoverage
        {
            private readonly bool blocksAllKinds;
            private readonly HashSet<string> incompleteKinds;

            private IdentityCoverage(bool blocksAllKinds, HashSet<string> incompleteKinds)
            {
                this.blocksAllKinds = blocksAllKinds;
                this.incompleteKinds = incompleteKinds;
            }

            public static IdentityCoverage From(IEnumerable<ComponentIdentity> items)
            {
                bool blocksAll = false;
                var incomplete = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                foreach (var item in items)
                {
                    if (string.IsNullOrWhiteSpace(item.SemanticKind)) blocksAll = true;
                    else if (!IsResolved(item)) incomplete.Add(item.SemanticKind);
                }
                return new IdentityCoverage(blocksAll, incomplete);
            }

            public bool IsComplete(string semanticKind) => !blocksAllKinds &&
                !string.IsNullOrWhiteSpace(semanticKind) && !incompleteKinds.Contains(semanticKind);
        }

        private static bool IsResolved(ComponentIdentity identity) => identity.Status == IdentityResolutionStatus.Resolved;
        private static string Key(ComponentIdentity identity) => identity.ComponentTypeKey.Length + ":" + identity.ComponentTypeKey + identity.ComparisonKey;
    }
}
