using System;
using System.Linq;
using System.Threading;
using D365SolutionComparer.Models.Membership;
using D365SolutionComparer.Services.Membership;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Microsoft.Xrm.Sdk;
using static D365SolutionComparer.Tests.MembershipTestData;

namespace D365SolutionComparer.Tests
{
    [TestClass]
    public class MembershipComparisonTests
    {
        [TestMethod]
        public void EmptyMembershipsProduceNoComponentRows()
        {
            Assert.AreEqual(0, new SolutionMembershipComparer().Compare(Snapshot(), Snapshot()).Count);
        }

        [TestMethod]
        public void ResolvedIdentityMatchesAcrossDifferentLocalIdsAndKeyCase()
        {
            var source = Identity("new_/script.js"); var target = Identity("NEW_/SCRIPT.JS");
            Assert.AreNotEqual(source.Record.ObjectId, target.Record.ObjectId);
            var result = new SolutionMembershipComparer().Compare(Snapshot(source), Snapshot(target)).Single();
            Assert.AreEqual(MembershipPresence.PresentInBoth, result.Presence);
            Assert.AreSame(source, result.Source);
            Assert.AreSame(target, result.Target);
        }

        [DataTestMethod]
        [DataRow(true)]
        [DataRow(false)]
        public void ReadResolveCompareSupportsOneSidedSolutionEndToEnd(bool sourcePresent)
        {
            var present = Solution(); var absent = Solution(); var member = ComponentRow(present);
            var presentService = Service(present, query =>
            {
                if (query.EntityName == "solution") return Rows(SolutionRow(present));
                if (query.EntityName == "solutioncomponent") return Rows(member);
                Assert.AreEqual("webresource", query.EntityName);
                return Rows(new Entity("webresource", member.GetAttributeValue<Guid>("objectid")) { ["name"] = "new_/script.js" });
            });
            var absentService = Service(absent, query => Rows());
            var reader = new DataverseSolutionMembershipReader(); var resolver = new DataverseComponentIdentityResolver();
            var presentSnapshot = resolver.ResolveSnapshot(presentService, reader.Read(presentService, present, CancellationToken.None), CancellationToken.None);
            var absentSnapshot = reader.Read(absentService, absent.Environment, absent.UniqueName, CancellationToken.None);
            var result = new SolutionMembershipComparer().Compare(sourcePresent ? presentSnapshot : absentSnapshot,
                sourcePresent ? absentSnapshot : presentSnapshot).Single();
            Assert.AreEqual(sourcePresent ? MembershipPresence.OnlyInSource : MembershipPresence.OnlyInTarget, result.Presence);
            Assert.AreEqual(MembershipAbsenceEvidence.OppositeSolutionAbsent, result.AbsenceEvidence);
        }

        [DataTestMethod]
        [DataRow(true)]
        [DataRow(false)]
        public void CompleteEmptyOppositeInventoryEstablishesMembershipAbsence(bool sourcePresent)
        {
            var present = Snapshot(Identity("item")); var empty = Snapshot();
            var result = new SolutionMembershipComparer().Compare(sourcePresent ? present : empty, sourcePresent ? empty : present).Single();
            Assert.AreEqual(sourcePresent ? MembershipPresence.OnlyInSource : MembershipPresence.OnlyInTarget, result.Presence);
            Assert.AreEqual(MembershipAbsenceEvidence.CompleteResolvedInventory, result.AbsenceEvidence);
        }

        [DataTestMethod]
        [DataRow(IdentityResolutionStatus.Unresolved)]
        [DataRow(IdentityResolutionStatus.Unsupported)]
        [DataRow(IdentityResolutionStatus.Ambiguous)]
        public void UnknownIdentityBlocksFalseMissingInBothDirections(IdentityResolutionStatus status)
        {
            var unknown = Identity(null, status: status); var known = Identity("key");
            foreach (var results in new[] { new SolutionMembershipComparer().Compare(Snapshot(known), Snapshot(unknown)),
                new SolutionMembershipComparer().Compare(Snapshot(unknown), Snapshot(known)) })
            {
                Assert.AreEqual(2, results.Count);
                Assert.IsTrue(results.All(r => r.Presence == MembershipPresence.Indeterminate));
            }
        }

        [TestMethod]
        public void UnsupportedIdentityAgainstAbsentSolutionIsStillIndeterminate()
        {
            var source = Snapshot(Identity(null, 98765, IdentityResolutionStatus.Unsupported));
            var target = MembershipSnapshot.Absent(Solution().Environment, "sample", DateTimeOffset.UtcNow);
            var result = new SolutionMembershipComparer().Compare(source, target).Single();
            Assert.AreEqual(MembershipPresence.Indeterminate, result.Presence);
            Assert.AreEqual(IdentityResolutionStatus.Unsupported, result.Source.Status);
        }

        [TestMethod]
        public void DuplicateKeysBecomeAmbiguousWithoutChoosingFirstOrMutatingSnapshots()
        {
            var first = Identity("key"); var second = Identity("KEY"); var source = Snapshot(first, second); var target = Snapshot(Identity("key"));
            var results = new SolutionMembershipComparer().Compare(source, target);
            Assert.AreEqual(3, results.Count);
            Assert.IsTrue(results.All(r => r.Presence == MembershipPresence.Indeterminate));
            Assert.IsTrue(results.Where(r => r.Source != null).All(r => r.Source.Status == IdentityResolutionStatus.Ambiguous));
            Assert.AreEqual(IdentityResolutionStatus.Resolved, first.Status);
            Assert.AreEqual(IdentityResolutionStatus.Resolved, second.Status);
        }

        [TestMethod]
        public void UnavailableOppositeSideNeverEstablishesAbsence()
        {
            var unavailable = MembershipSnapshot.Unavailable(Solution().Environment, "sample", DateTimeOffset.UtcNow, "Read denied");
            var comparer = new SolutionMembershipComparer();
            Assert.AreEqual(MembershipPresence.Indeterminate, comparer.Compare(Snapshot(Identity("key")), unavailable).Single().Presence);
            Assert.AreEqual(MembershipPresence.Indeterminate, comparer.Compare(unavailable, Snapshot(Identity("key"))).Single().Presence);
        }

        [TestMethod]
        public void UnknownCoverageDoesNotSuppressAnIndependentlyResolvedMatch()
        {
            var source = Snapshot(Identity("key"), Identity(null, 9999, IdentityResolutionStatus.Unsupported));
            var target = Snapshot(Identity("key"));
            var results = new SolutionMembershipComparer().Compare(source, target);
            Assert.AreEqual(1, results.Count(r => r.Presence == MembershipPresence.PresentInBoth));
            Assert.AreEqual(1, results.Count(r => r.Presence == MembershipPresence.Indeterminate));
        }

        [TestMethod]
        public void SameNameAcrossDifferentComponentKindsIsNotAMatch()
        {
            var results = new SolutionMembershipComparer().Compare(Snapshot(Identity("same", 61, kind: "webresource")),
                Snapshot(Identity("same", 29, kind: "process")));
            Assert.AreEqual(2, results.Count);
            Assert.IsFalse(results.Any(r => r.Presence == MembershipPresence.PresentInBoth));
        }

        [TestMethod]
        public void SameCanonicalKindWithDifferentRawCodesIsSupported()
        {
            var results = new SolutionMembershipComparer().Compare(Snapshot(Identity("new_shared", 10027, kind: "connectionreference")),
                Snapshot(Identity("new_shared", 10150, kind: "connectionreference")));
            Assert.AreEqual(MembershipPresence.PresentInBoth, results.Single().Presence);
        }

        [TestMethod]
        public void WrongSolutionNamesAreRejectedAndBothAbsentRetainNoComponentRows()
        {
            var first = MembershipSnapshot.Absent(Solution().Environment, "sample", DateTimeOffset.UtcNow);
            var second = MembershipSnapshot.Absent(Solution().Environment, "SAMPLE", DateTimeOffset.UtcNow);
            var wrong = MembershipSnapshot.Absent(Solution().Environment, "different", DateTimeOffset.UtcNow);
            var comparer = new SolutionMembershipComparer();
            Assert.AreEqual(0, comparer.Compare(first, second).Count);
            Assert.ThrowsException<ArgumentException>(() => comparer.Compare(first, wrong));
        }
    }
}
