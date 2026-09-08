using System;
using System.Collections.Generic;
using D365SolutionComparer.Models.Identity;
using D365SolutionComparer.Models.Membership;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace D365SolutionComparer.Tests
{
    [TestClass]
    public class MembershipModelTests
    {
        private static EnvironmentIdentity Environment() => new EnvironmentIdentity(Guid.NewGuid(), "Display label");
        private static ComponentIdentity Identity(IdentityResolutionStatus status = IdentityResolutionStatus.Resolved) =>
            new ComponentIdentity(new SolutionComponentRecord(Guid.NewGuid(), 1, Guid.NewGuid()), status,
                status == IdentityResolutionStatus.Resolved ? "account" : null, "Diagnostic");

        [TestMethod]
        public void SolutionIdentityRetainsLocalEnvironmentAndUnmodifiedUniqueName()
        {
            var environment = Environment(); var id = Guid.NewGuid();
            var solution = new SolutionIdentity(environment, id, " Sample ");
            Assert.AreSame(environment, solution.Environment);
            Assert.AreEqual(id, solution.SolutionId);
            Assert.AreEqual(" Sample ", solution.UniqueName);
            var other = new SolutionIdentity(Environment(), Guid.NewGuid(), " Sample ");
            Assert.AreNotEqual(solution.Environment.OrganizationId, other.Environment.OrganizationId);
        }

        [TestMethod]
        public void IdentityRejectsMissingRequiredIdentifiers()
        {
            Assert.ThrowsException<ArgumentException>(() => new EnvironmentIdentity(Guid.Empty, "Label"));
            Assert.ThrowsException<ArgumentNullException>(() => new SolutionIdentity(null, Guid.NewGuid(), "sample"));
            Assert.ThrowsException<ArgumentException>(() => new SolutionIdentity(Environment(), Guid.Empty, "sample"));
            Assert.ThrowsException<ArgumentException>(() => new SolutionIdentity(Environment(), Guid.NewGuid(), " "));
        }

        [TestMethod]
        public void RawRecordsPreserveUnknownTypesBehaviorsAndOptionalData()
        {
            var parent = Guid.NewGuid();
            var record = new SolutionComponentRecord(Guid.NewGuid(), 987654, null, 1234, parent, null);
            Assert.AreEqual(987654, record.ComponentType);
            Assert.AreEqual(1234, record.RootComponentBehavior);
            Assert.AreEqual(parent, record.RootSolutionComponentId);
            Assert.IsNull(record.ObjectId);
            Assert.IsNull(record.IsMetadata);
        }

        [DataTestMethod]
        [DataRow(1, ComponentSemanticKinds.Table)]
        [DataRow(2, ComponentSemanticKinds.Column)]
        [DataRow(3, ComponentSemanticKinds.Relationship)]
        [DataRow(8, ComponentSemanticKinds.Relationship)]
        [DataRow(10, ComponentSemanticKinds.Relationship)]
        [DataRow(11, ComponentSemanticKinds.Relationship)]
        [DataRow(12, ComponentSemanticKinds.Relationship)]
        [DataRow(61, ComponentSemanticKinds.WebResource)]
        [DataRow(80, ComponentSemanticKinds.AppModule)]
        [DataRow(29, ComponentSemanticKinds.Process)]
        [DataRow(20, ComponentSemanticKinds.SecurityRole)]
        [DataRow(380, ComponentSemanticKinds.EnvironmentVariableDefinition)]
        public void RawComponentTypeProvidesSemanticKindWithoutAResolvedPortableIdentity(int componentType,
            string expectedKind)
        {
            var identity = new ComponentIdentity(
                new SolutionComponentRecord(Guid.NewGuid(), componentType, Guid.NewGuid()),
                IdentityResolutionStatus.Unresolved);

            Assert.AreEqual(expectedKind, identity.SemanticKind);
        }

        [TestMethod]
        public void KnownUnsupportedAndUnknownRawTypesHaveDifferentCoverageScopes()
        {
            var knownUnsupported = new ComponentIdentity(
                new SolutionComponentRecord(Guid.NewGuid(), 9, Guid.NewGuid()),
                IdentityResolutionStatus.Unsupported);
            var unknown = new ComponentIdentity(
                new SolutionComponentRecord(Guid.NewGuid(), 98765, Guid.NewGuid()),
                IdentityResolutionStatus.Unsupported);
            var dynamicConnectionReference = new ComponentIdentity(
                new SolutionComponentRecord(Guid.NewGuid(), 10027, Guid.NewGuid()),
                IdentityResolutionStatus.Unresolved, componentTypeKey: ComponentSemanticKinds.ConnectionReference);

            Assert.AreEqual("unsupported:componenttype:9", knownUnsupported.SemanticKind);
            Assert.IsNull(unknown.SemanticKind);
            Assert.AreEqual(ComponentSemanticKinds.ConnectionReference, dynamicConnectionReference.SemanticKind);
        }

        [DataTestMethod]
        [DataRow(IdentityResolutionStatus.Unresolved)]
        [DataRow(IdentityResolutionStatus.Unsupported)]
        [DataRow(IdentityResolutionStatus.Ambiguous)]
        public void UntrustedIdentitiesCannotAcquireKeysOrMissingClassifications(IdentityResolutionStatus status)
        {
            var identity = Identity(status);
            Assert.AreEqual(status, identity.Status);
            Assert.AreEqual("Diagnostic", identity.Diagnostic);
            Assert.IsNull(identity.ComparisonKey);
            Assert.ThrowsException<ArgumentException>(() => new ComponentIdentity(identity.Record, status, "account"));
            foreach (var evidence in new[] { MembershipAbsenceEvidence.None,
                MembershipAbsenceEvidence.OppositeSolutionAbsent, MembershipAbsenceEvidence.CompleteResolvedInventory })
            {
                Assert.ThrowsException<ArgumentException>(() => new MembershipCompareResult(identity, null, MembershipPresence.OnlyInSource, evidence));
                Assert.ThrowsException<ArgumentException>(() => new MembershipCompareResult(null, identity, MembershipPresence.OnlyInTarget, evidence));
            }
            Assert.ThrowsException<ArgumentException>(() => new MembershipCompareResult(identity, Identity(), MembershipPresence.PresentInBoth));
            Assert.AreEqual(MembershipPresence.Indeterminate,
                new MembershipCompareResult(identity, null, MembershipPresence.Indeterminate).Presence);
        }

        [TestMethod]
        public void MissingRequiresExplicitEvidenceEvenForResolvedIdentity()
        {
            Assert.ThrowsException<ArgumentException>(() => new MembershipCompareResult(Identity(), null, MembershipPresence.OnlyInSource));
            Assert.ThrowsException<ArgumentException>(() => new MembershipCompareResult(null, Identity(), MembershipPresence.OnlyInTarget));
            Assert.ThrowsException<ArgumentException>(() => new MembershipCompareResult(Identity(), Identity(), MembershipPresence.OnlyInSource,
                MembershipAbsenceEvidence.CompleteResolvedInventory));
        }

        [DataTestMethod]
        [DataRow(true, MembershipAbsenceEvidence.OppositeSolutionAbsent)]
        [DataRow(false, MembershipAbsenceEvidence.OppositeSolutionAbsent)]
        [DataRow(true, MembershipAbsenceEvidence.CompleteResolvedInventory)]
        [DataRow(false, MembershipAbsenceEvidence.CompleteResolvedInventory)]
        public void ResolvedOneSidedMembershipIsRepresentableInEitherDirection(bool source, MembershipAbsenceEvidence evidence)
        {
            var identity = Identity();
            var presence = source ? MembershipPresence.OnlyInSource : MembershipPresence.OnlyInTarget;
            var result = new MembershipCompareResult(source ? identity : null, source ? null : identity, presence, evidence);
            Assert.AreEqual(presence, result.Presence);
            Assert.AreEqual(evidence, result.AbsenceEvidence);
        }

        [TestMethod]
        public void CompleteEmptyAbsentAndUnavailableSnapshotsRemainDistinct()
        {
            var environment = Environment(); var time = DateTimeOffset.UtcNow;
            var complete = MembershipSnapshot.Complete(new SolutionIdentity(environment, Guid.NewGuid(), "sample"), new ComponentIdentity[0], time);
            var absent = MembershipSnapshot.Absent(environment, "sample", time);
            var unavailable = MembershipSnapshot.Unavailable(environment, "sample", time, "Read denied");
            Assert.AreEqual(MembershipSnapshotState.Complete, complete.State);
            Assert.IsNotNull(complete.Solution);
            Assert.AreEqual(MembershipSnapshotState.SolutionAbsent, absent.State);
            Assert.IsNull(absent.Solution);
            Assert.AreEqual(MembershipSnapshotState.Unavailable, unavailable.State);
            Assert.AreEqual("Read denied", unavailable.Diagnostic);
            Assert.AreSame(environment, absent.Environment);
            Assert.AreEqual(time, complete.CapturedAt);
            Assert.ThrowsException<ArgumentException>(() => MembershipSnapshot.Unavailable(environment, "sample", time, ""));
        }

        [TestMethod]
        public void CompleteSnapshotRetainsUnresolvedRowsAndDefensivelyCopiesCollection()
        {
            var identity = Identity(IdentityResolutionStatus.Unsupported);
            var items = new List<ComponentIdentity> { identity };
            var snapshot = MembershipSnapshot.Complete(new SolutionIdentity(Environment(), Guid.NewGuid(), "sample"), items, DateTimeOffset.UtcNow);
            items.Clear();
            Assert.AreEqual(1, snapshot.Components.Count);
            Assert.AreSame(identity, snapshot.Components[0]);
            Assert.AreEqual(MembershipSnapshotState.Complete, snapshot.State);
            Assert.ThrowsException<NotSupportedException>(() => ((IList<ComponentIdentity>)snapshot.Components).Clear());
        }

        [TestMethod]
        public void InvalidResultStatesAndKeysAreRejected()
        {
            var record = Identity().Record;
            Assert.ThrowsException<ArgumentException>(() => new ComponentIdentity(record, IdentityResolutionStatus.Resolved, " "));
            Assert.ThrowsException<ArgumentOutOfRangeException>(() => new ComponentIdentity(record, (IdentityResolutionStatus)999));
            Assert.ThrowsException<ArgumentException>(() => new MembershipCompareResult(null, null, MembershipPresence.Indeterminate));
            Assert.ThrowsException<ArgumentOutOfRangeException>(() => new MembershipCompareResult(Identity(), null, (MembershipPresence)999));
            Assert.ThrowsException<ArgumentOutOfRangeException>(() => new MembershipCompareResult(Identity(), null, MembershipPresence.OnlyInSource, (MembershipAbsenceEvidence)999));
            Assert.ThrowsException<ArgumentException>(() => new MembershipCompareResult(Identity(), null, MembershipPresence.Indeterminate,
                MembershipAbsenceEvidence.OppositeSolutionAbsent));
        }

        [TestMethod]
        public void SharedPresenceRequiresResolvedSameTypeIdentities()
        {
            Assert.AreEqual(MembershipPresence.PresentInBoth,
                new MembershipCompareResult(Identity(), Identity(), MembershipPresence.PresentInBoth).Presence);
            var otherType = new ComponentIdentity(new SolutionComponentRecord(Guid.NewGuid(), 61, Guid.NewGuid()), IdentityResolutionStatus.Resolved, "account");
            Assert.ThrowsException<ArgumentException>(() => new MembershipCompareResult(Identity(), otherType, MembershipPresence.PresentInBoth));
        }
    }
}
