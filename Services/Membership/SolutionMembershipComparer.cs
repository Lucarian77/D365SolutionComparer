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
            var sourceItems = MarkDuplicates(GuardWorkflowAlternatives(source.Components, target.Components));
            var targetItems = MarkDuplicates(GuardWorkflowAlternatives(target.Components, source.Components));
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

        // A semantic candidate cannot prove absence of a differently keyed uniquename record
        // with the same semantic evidence. Keep strategies separate and report the uncertainty.
        private static IReadOnlyList<ComponentIdentity> GuardWorkflowAlternatives(
            IReadOnlyList<ComponentIdentity> items, IReadOnlyList<ComponentIdentity> opposite)
        {
            return items.Select(item =>
            {
                if (!IsResolved(item) || item.SemanticKind != ComponentSemanticKinds.Process) return item;
                bool fallback = item.WorkflowCandidateKey != null &&
                    StringComparer.OrdinalIgnoreCase.Equals(item.WorkflowCandidateKey, item.ComparisonKey);
                bool uncertain = opposite.Where(other => IsResolved(other) && other.SemanticKind == ComponentSemanticKinds.Process &&
                    !StringComparer.OrdinalIgnoreCase.Equals(item.ComparisonKey, other.ComparisonKey)).Any(other =>
                {
                    bool otherFallback = other.WorkflowCandidateKey != null &&
                        StringComparer.OrdinalIgnoreCase.Equals(other.WorkflowCandidateKey, other.ComparisonKey);
                    return (fallback || otherFallback) && (item.WorkflowCandidateKey == null || other.WorkflowCandidateKey == null ||
                        StringComparer.OrdinalIgnoreCase.Equals(item.WorkflowCandidateKey, other.WorkflowCandidateKey));
                });
                return uncertain ? new ComponentIdentity(item.Record, IdentityResolutionStatus.Unresolved,
                    diagnostic: "Workflow uniquename and semantic fallback evidence cannot establish a unique cross-strategy correlation or absence.",
                    componentTypeKey: item.ComponentTypeKey, semanticKind: item.SemanticKind,
                    diagnosticEvidence: item.DiagnosticEvidence, workflowCandidateKey: item.WorkflowCandidateKey,
                    blockerPortableIdentity: item.WorkflowCandidateKey,
                    blockerScope: ResolutionBlockerScope.PortableIdentity) : item;
            }).ToList().AsReadOnly();
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
                    oppositeCoverage.CanEstablishAbsence(item))
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
                    diagnostic: i.SemanticKind == ComponentSemanticKinds.Process
                        ? "Duplicate Process / Workflow fallback candidates share this portable semantic identity. Only this identity is ambiguous; unrelated workflow identities remain eligible for definitive absence checks."
                        : "Multiple membership records resolve to the same identity key: " + i.ComparisonKey,
                    componentTypeKey: i.ComponentTypeKey, semanticKind: i.SemanticKind,
                    diagnosticEvidence: i.DiagnosticEvidence,
                    workflowCandidateKey: i.WorkflowCandidateKey,
                    blockerPortableIdentity: i.SemanticKind == ComponentSemanticKinds.Process ? i.ComparisonKey : null,
                    blockerScope: i.SemanticKind == ComponentSemanticKinds.Process
                        ? ResolutionBlockerScope.PortableIdentity : ResolutionBlockerScope.SemanticKind) : i).ToList().AsReadOnly();
        }

        private sealed class IdentityCoverage
        {
            private readonly bool blocksAllKinds;
            private readonly HashSet<string> incompleteKinds;
            private readonly Dictionary<string, HashSet<string>> identityBlockers;

            private IdentityCoverage(bool blocksAllKinds, HashSet<string> incompleteKinds,
                Dictionary<string, HashSet<string>> identityBlockers)
            {
                this.blocksAllKinds = blocksAllKinds;
                this.incompleteKinds = incompleteKinds;
                this.identityBlockers = identityBlockers;
            }

            public static IdentityCoverage From(IEnumerable<ComponentIdentity> items)
            {
                bool blocksAll = false;
                var incomplete = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                var identityBlockers = new Dictionary<string, HashSet<string>>(StringComparer.OrdinalIgnoreCase);
                foreach (var item in items)
                {
                    if (string.IsNullOrWhiteSpace(item.SemanticKind)) blocksAll = true;
                    else if (!IsResolved(item) && item.BlockerScope == ResolutionBlockerScope.SemanticKind)
                        incomplete.Add(item.SemanticKind);
                    if (!IsResolved(item) && item.BlockerScope == ResolutionBlockerScope.PortableIdentity &&
                        !string.IsNullOrWhiteSpace(item.SemanticKind) && !string.IsNullOrWhiteSpace(item.BlockerPortableIdentity))
                    {
                        HashSet<string> keys;
                        if (!identityBlockers.TryGetValue(item.SemanticKind, out keys))
                            identityBlockers[item.SemanticKind] = keys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                        keys.Add(item.BlockerPortableIdentity);
                    }
                    if (!IsResolved(item) &&
                        ComponentSemanticKinds.IsGlobalChoiceCandidate(item.ComponentTypeKey))
                        incomplete.Add(ComponentSemanticKinds.GlobalChoice);
                    if (!IsResolved(item) &&
                        ComponentSemanticKinds.IsReportCandidate(item.ComponentTypeKey))
                        incomplete.Add(ComponentSemanticKinds.Report);
                }
                return new IdentityCoverage(blocksAll, incomplete, identityBlockers);
            }

            public bool CanEstablishAbsence(ComponentIdentity identity)
            {
                if (blocksAllKinds || identity == null || string.IsNullOrWhiteSpace(identity.SemanticKind) ||
                    incompleteKinds.Contains(identity.SemanticKind)) return false;
                HashSet<string> keys;
                return !identityBlockers.TryGetValue(identity.SemanticKind, out keys) ||
                    string.IsNullOrWhiteSpace(identity.ComparisonKey) || !keys.Contains(identity.ComparisonKey);
            }
        }

        private static bool IsResolved(ComponentIdentity identity) => identity.Status == IdentityResolutionStatus.Resolved;
        private static string Key(ComponentIdentity identity) => identity.ComponentTypeKey.Length + ":" + identity.ComponentTypeKey + identity.ComparisonKey;
    }
}
