using System;
using System.Collections.Generic;
using System.Linq;
using D365SolutionComparer.Models.Membership;
using D365SolutionComparer.Services.Membership;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace D365SolutionComparer.Tests
{
    [TestClass]
    public class MembershipComparisonPhase2G5Tests
    {
        [TestMethod, TestCategory("Phase2G5")]
        public void UnrelatedIdentityScopedAmbiguityDoesNotBlockTargetOnly()
        {
            var source = Snapshot(Blocked("B"));
            var target = Snapshot(Resolved("A"));
            var result = Compare(source, target).Single(item => item.Target != null && item.Target.Status == IdentityResolutionStatus.Resolved);
            Assert.AreEqual(MembershipPresence.OnlyInTarget, result.Presence);
            Assert.AreEqual(MembershipAbsenceEvidence.CompleteResolvedInventory, result.AbsenceEvidence);
        }

        [TestMethod, TestCategory("Phase2G5")]
        public void UnrelatedIdentityScopedAmbiguityDoesNotBlockSourceOnly()
        {
            var source = Snapshot(Resolved("A"));
            var target = Snapshot(Blocked("B"));
            var result = Compare(source, target).Single(item => item.Source != null && item.Source.Status == IdentityResolutionStatus.Resolved);
            Assert.AreEqual(MembershipPresence.OnlyInSource, result.Presence);
        }

        [TestMethod, TestCategory("Phase2G5")]
        public void SameIdentityScopedAmbiguityBlocksTargetOnly()
        {
            var results = Compare(Snapshot(Blocked("A")), Snapshot(Resolved("A")));
            var result = results
                .Single(item => item.Target != null && item.Target.Status == IdentityResolutionStatus.Resolved);
            Assert.AreEqual(MembershipPresence.Indeterminate, result.Presence);
            Assert.AreEqual(ResolutionBlockerScope.PortableIdentity, results
                .Single(item => item.Source != null && item.Source.Status == IdentityResolutionStatus.Ambiguous)
                .Source.BlockerScope);
        }

        [TestMethod, TestCategory("Phase2G5")]
        public void SameIdentityScopedAmbiguityBlocksSourceOnly()
        {
            var result = Compare(Snapshot(Resolved("A")), Snapshot(Blocked("A")))
                .Single(item => item.Source != null && item.Source.Status == IdentityResolutionStatus.Resolved);
            Assert.AreEqual(MembershipPresence.Indeterminate, result.Presence);
        }

        [TestMethod, TestCategory("Phase2G5")]
        public void KindWideUnresolvedEvidenceBlocksAbsence()
        {
            var results = Compare(Snapshot(Unresolved()), Snapshot(Resolved("A")));
            var result = results
                .Single(item => item.Target != null && item.Target.Status == IdentityResolutionStatus.Resolved);
            Assert.AreEqual(MembershipPresence.Indeterminate, result.Presence);
            Assert.AreEqual(ResolutionBlockerScope.SemanticKind, results
                .Single(item => item.Source != null && item.Source.Status == IdentityResolutionStatus.Unresolved)
                .Source.BlockerScope);
        }

        [TestMethod, TestCategory("Phase2G5")]
        public void PortableBlockerKeysUseOrdinalIgnoreCase()
        {
            var result = Compare(Snapshot(Blocked("workflow-semantic:v1:A")),
                Snapshot(Resolved("WORKFLOW-SEMANTIC:V1:A")))
                .Single(item => item.Target != null && item.Target.Status == IdentityResolutionStatus.Resolved);
            Assert.AreEqual(MembershipPresence.Indeterminate, result.Presence);
        }

        [TestMethod, TestCategory("Phase2G5")]
        public void DuplicateWorkflowFallbackCandidatesRemainAmbiguous()
        {
            var source = Snapshot(Resolved("A"), Resolved("a"));
            var results = Compare(source, SnapshotAbsent());
            Assert.IsTrue(results.All(item => item.Source.Status == IdentityResolutionStatus.Ambiguous));
            Assert.IsTrue(results.All(item => item.Source.BlockerScope == ResolutionBlockerScope.PortableIdentity));
            Assert.IsTrue(results.All(item => item.Source.BlockerPortableIdentity != null));
        }

        [TestMethod, TestCategory("Phase2G5")]
        public void ResolvedSameKeyProducesPresentInBoth()
        {
            var result = Compare(Snapshot(Resolved("A")), Snapshot(Resolved("a"))).Single();
            Assert.AreEqual(MembershipPresence.PresentInBoth, result.Presence);
        }

        [TestMethod, TestCategory("Phase2G5")]
        public void OneSidedDefinitionsRemainDefinitiveWhenCoverageAllowsIt()
        {
            var source = Snapshot(Resolved("A"));
            var result = Compare(source, SnapshotAbsent()).Single();
            Assert.AreEqual(MembershipPresence.OnlyInSource, result.Presence);
        }

        [TestMethod, TestCategory("Phase2G5")]
        public void WorkflowConfigurationChangesDoNotChangePortablePresence()
        {
            var result = Compare(Snapshot(Resolved("A")), Snapshot(Resolved("B"))).ToList();
            Assert.IsTrue(result.All(item => item.Presence != MembershipPresence.PresentInBoth));
        }

        [TestMethod, TestCategory("Phase2G5")]
        public void BlockerEvaluationIsDeterministicRegardlessOfInputOrder()
        {
            var first = Compare(Snapshot(Blocked("B"), Resolved("A")), Snapshot(Resolved("C")))
                .Select(item => item.Presence + ":" + (item.Source ?? item.Target).ComparisonKey).ToArray();
            var second = Compare(Snapshot(Resolved("A"), Blocked("B")), Snapshot(Resolved("C")))
                .Select(item => item.Presence + ":" + (item.Source ?? item.Target).ComparisonKey).ToArray();
            CollectionAssert.AreEqual(first, second);
        }

        [TestMethod, TestCategory("Phase2G5")]
        public void LocalGuidDifferencesDoNotAffectPortablePresence()
        {
            var a = Resolved("A");
            var b = Resolved("A");
            Assert.AreNotEqual(a.Record.ObjectId, b.Record.ObjectId);
            Assert.AreEqual(MembershipPresence.PresentInBoth, Compare(Snapshot(a), Snapshot(b)).Single().Presence);
        }

        [TestMethod, TestCategory("Phase2G5")]
        public void ManagedStateDoesNotAffectPortablePresence()
        {
            var a = Resolved("A", workflowCandidate: "A");
            var b = Resolved("A", workflowCandidate: "A");
            Assert.AreEqual(MembershipPresence.PresentInBoth, Compare(Snapshot(a), Snapshot(b)).Single().Presence);
        }

        [TestMethod, TestCategory("Phase2G5")]
        public void UnrelatedFamiliesRetainKindAwareAbsenceBehavior()
        {
            var source = Snapshot(new ComponentIdentity(new SolutionComponentRecord(Guid.NewGuid(), 1, Guid.NewGuid()),
                IdentityResolutionStatus.Resolved, "account", componentTypeKey: "table", semanticKind: ComponentSemanticKinds.Table));
            var target = Snapshot(Blocked("workflow"));
            Assert.AreEqual(MembershipPresence.OnlyInSource, Compare(source, target).Single(item => item.Source != null).Presence);
        }

        [TestMethod, TestCategory("Phase2G5")]
        public void BroadUnknownBlockerStillBlocksEveryKind()
        {
            var broad = new ComponentIdentity(new SolutionComponentRecord(Guid.NewGuid(), 99999, Guid.NewGuid()),
                IdentityResolutionStatus.Unsupported, semanticKind: null);
            var result = Compare(Snapshot(broad), Snapshot(Resolved("A"))).Single(item => item.Target != null);
            Assert.AreEqual(MembershipPresence.Indeterminate, result.Presence);
        }

        [TestMethod, TestCategory("Phase2G5")]
        public void SolutionAbsentStillProvesAbsenceDespiteBlockers()
        {
            var result = Compare(Snapshot(Resolved("A")), SnapshotAbsent()).Single();
            Assert.AreEqual(MembershipPresence.OnlyInSource, result.Presence);
            Assert.AreEqual(MembershipAbsenceEvidence.OppositeSolutionAbsent, result.AbsenceEvidence);
        }

        private static IReadOnlyList<MembershipCompareResult> Compare(MembershipSnapshot source, MembershipSnapshot target) =>
            new SolutionMembershipComparer().Compare(source, target);

        private static MembershipSnapshot Snapshot(params ComponentIdentity[] identities) =>
            MembershipSnapshot.Complete(MembershipTestData.Solution(), identities, DateTimeOffset.UtcNow);

        private static MembershipSnapshot SnapshotAbsent() =>
            MembershipSnapshot.Absent(MembershipTestData.Solution().Environment, "sample", DateTimeOffset.UtcNow);

        private static ComponentIdentity Resolved(string key, string workflowCandidate = null) =>
            new ComponentIdentity(new SolutionComponentRecord(Guid.NewGuid(), 29, Guid.NewGuid()),
                IdentityResolutionStatus.Resolved, key, componentTypeKey: "process",
                semanticKind: ComponentSemanticKinds.Process, workflowCandidateKey: workflowCandidate);

        private static ComponentIdentity Blocked(string key) =>
            new ComponentIdentity(new SolutionComponentRecord(Guid.NewGuid(), 29, Guid.NewGuid()),
                IdentityResolutionStatus.Ambiguous, componentTypeKey: "process",
                semanticKind: ComponentSemanticKinds.Process, blockerPortableIdentity: key,
                blockerScope: ResolutionBlockerScope.PortableIdentity);

        private static ComponentIdentity Unresolved() =>
            new ComponentIdentity(new SolutionComponentRecord(Guid.NewGuid(), 29, Guid.NewGuid()),
                IdentityResolutionStatus.Unresolved, componentTypeKey: "process",
                semanticKind: ComponentSemanticKinds.Process);
    }
}
