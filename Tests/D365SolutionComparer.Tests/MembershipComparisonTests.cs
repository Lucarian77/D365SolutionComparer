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
        public void UnresolvedWorkflowDoesNotBlockColumnSourceOnly()
        {
            var column = Identity("activitypointer.regardingobjectid", 2, kind: ComponentSemanticKinds.Column);
            var workflow = Identity(null, 29, IdentityResolutionStatus.Unresolved,
                ComponentSemanticKinds.Process);

            var results = new SolutionMembershipComparer().Compare(Snapshot(column), Snapshot(workflow));

            var columnResult = results.Single(result => result.Source == column);
            Assert.AreEqual(MembershipPresence.OnlyInSource, columnResult.Presence);
            Assert.AreEqual(MembershipAbsenceEvidence.CompleteResolvedInventory, columnResult.AbsenceEvidence);
            Assert.AreEqual(MembershipPresence.Indeterminate,
                results.Single(result => result.Target == workflow).Presence);
        }

        [TestMethod]
        public void UnresolvedColumnBlocksColumnSourceOnly()
        {
            var sourceColumn = Identity("activitypointer.regardingobjectid", 2,
                kind: ComponentSemanticKinds.Column);
            var targetColumn = Identity(null, 2, IdentityResolutionStatus.Unresolved,
                ComponentSemanticKinds.Column);

            var results = new SolutionMembershipComparer().Compare(Snapshot(sourceColumn), Snapshot(targetColumn));

            Assert.AreEqual(MembershipPresence.Indeterminate,
                results.Single(result => result.Source == sourceColumn).Presence);
        }

        [TestMethod]
        public void KnownUnsupportedUnrelatedTypeDoesNotBlockColumnAbsence()
        {
            var column = Identity("account.name", 2, kind: ComponentSemanticKinds.Column);
            var optionSet = Identity(null, 9, IdentityResolutionStatus.Unsupported);
            Assert.IsFalse(string.IsNullOrWhiteSpace(optionSet.SemanticKind));
            Assert.AreNotEqual(ComponentSemanticKinds.Column, optionSet.SemanticKind);

            var results = new SolutionMembershipComparer().Compare(Snapshot(column), Snapshot(optionSet));

            Assert.AreEqual(MembershipPresence.OnlyInSource,
                results.Single(result => result.Source == column).Presence);
        }

        [TestMethod]
        public void UnresolvedRelationshipDoesNotBlockWebResourceAbsence()
        {
            var webResource = Identity("new_/script.js", 61, kind: ComponentSemanticKinds.WebResource);
            var relationship = Identity(null, 10, IdentityResolutionStatus.Unresolved,
                ComponentSemanticKinds.Relationship);

            var results = new SolutionMembershipComparer().Compare(Snapshot(webResource), Snapshot(relationship));

            Assert.AreEqual(MembershipPresence.OnlyInSource,
                results.Single(result => result.Source == webResource).Presence);
        }

        [TestMethod]
        public void DuplicateColumnKeysBecomeAmbiguousAndBlockColumnAbsence()
        {
            var sourceColumn = Identity("contact.emailaddress1", 2, kind: ComponentSemanticKinds.Column);
            var firstTargetColumn = Identity("account.name", 2, kind: ComponentSemanticKinds.Column);
            var secondTargetColumn = Identity("ACCOUNT.NAME", 2, kind: ComponentSemanticKinds.Column);

            var results = new SolutionMembershipComparer().Compare(Snapshot(sourceColumn),
                Snapshot(firstTargetColumn, secondTargetColumn));

            Assert.AreEqual(MembershipPresence.Indeterminate,
                results.Single(result => result.Source == sourceColumn).Presence);
            Assert.IsTrue(results.Where(result => result.Target != null)
                .All(result => result.Target.Status == IdentityResolutionStatus.Ambiguous));
        }

        [TestMethod]
        public void AmbiguousWorkflowDoesNotBlockColumnAbsence()
        {
            var column = Identity("account.name", 2, kind: ComponentSemanticKinds.Column);
            var firstWorkflow = Identity("new_process", 29, kind: ComponentSemanticKinds.Process);
            var secondWorkflow = Identity("NEW_PROCESS", 29, kind: ComponentSemanticKinds.Process);

            var results = new SolutionMembershipComparer().Compare(Snapshot(column),
                Snapshot(firstWorkflow, secondWorkflow));

            Assert.AreEqual(MembershipPresence.OnlyInSource,
                results.Single(result => result.Source == column).Presence);
            Assert.IsTrue(results.Where(result => result.Target != null)
                .All(result => result.Target.Status == IdentityResolutionStatus.Ambiguous));
        }

        [TestMethod]
        public void SolutionAbsentProvidesUniversalAbsenceEvidenceForSupportedKinds()
        {
            var present = Snapshot(
                Identity("account", 1, kind: ComponentSemanticKinds.Table),
                Identity("account.name", 2, kind: ComponentSemanticKinds.Column),
                Identity("account_contact", 10, kind: ComponentSemanticKinds.Relationship),
                Identity("new_/script.js", 61, kind: ComponentSemanticKinds.WebResource),
                Identity("new_process", 29, kind: ComponentSemanticKinds.Process),
                Identity(Guid.NewGuid().ToString("D"), 20, kind: ComponentSemanticKinds.SecurityRole),
                Identity("new_setting", 380, kind: ComponentSemanticKinds.EnvironmentVariableDefinition),
                Identity("new_shared", 10027, kind: ComponentSemanticKinds.ConnectionReference));
            var absent = MembershipSnapshot.Absent(Solution().Environment, "sample", DateTimeOffset.UtcNow);

            var results = new SolutionMembershipComparer().Compare(present, absent);

            Assert.AreEqual(8, results.Count);
            Assert.IsTrue(results.All(result => result.Presence == MembershipPresence.OnlyInSource));
            Assert.IsTrue(results.All(result =>
                result.AbsenceEvidence == MembershipAbsenceEvidence.OppositeSolutionAbsent));
        }

        [TestMethod]
        public void UnknownUnclassifiableRawTypeBlocksAbsenceConservatively()
        {
            var column = Identity("account.name", 2, kind: ComponentSemanticKinds.Column);
            var unknown = Identity(null, 98765, IdentityResolutionStatus.Unsupported);
            Assert.IsNull(unknown.SemanticKind);

            var results = new SolutionMembershipComparer().Compare(Snapshot(column), Snapshot(unknown));

            Assert.AreEqual(MembershipPresence.Indeterminate,
                results.Single(result => result.Source == column).Presence);
        }

        [TestMethod]
        public void OneSidedMixedKindsUseIndependentCoveragePerKind()
        {
            var table = Identity("account", 1, kind: ComponentSemanticKinds.Table);
            var column = Identity("account.name", 2, kind: ComponentSemanticKinds.Column);
            var workflow = Identity(null, 29, IdentityResolutionStatus.Unresolved,
                ComponentSemanticKinds.Process);
            var webResource = Identity("new_/target.js", 61, kind: ComponentSemanticKinds.WebResource);

            var results = new SolutionMembershipComparer().Compare(Snapshot(table, column),
                Snapshot(workflow, webResource));

            Assert.AreEqual(MembershipPresence.OnlyInSource,
                results.Single(result => result.Source == table).Presence);
            Assert.AreEqual(MembershipPresence.OnlyInSource,
                results.Single(result => result.Source == column).Presence);
            Assert.AreEqual(MembershipPresence.OnlyInTarget,
                results.Single(result => result.Target == webResource).Presence);
            Assert.AreEqual(MembershipPresence.Indeterminate,
                results.Single(result => result.Target == workflow).Presence);
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
