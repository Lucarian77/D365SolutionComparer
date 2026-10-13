using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Reflection;
using System.ServiceModel;
using System.Threading;
using System.Windows.Forms;
using D365SolutionComparer.Infrastructure;
using D365SolutionComparer.Models.ComponentDetails;
using D365SolutionComparer.Models.Identity;
using D365SolutionComparer.Models.Membership;
using D365SolutionComparer.Models.Presentation;
using D365SolutionComparer.Services.ComponentDetails;
using D365SolutionComparer.Services.Membership;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Query;
using static D365SolutionComparer.Tests.MembershipTestData;

namespace D365SolutionComparer.Tests
{
    [TestClass, TestCategory("H2BParentHandoff")]
    public class AppSettingParentEvidenceHandoffTests
    {
        public TestContext TestContext { get; set; }
        private const string Secret = "NEVER_RETAIN_SETTING_VALUES";

        [DataTestMethod]
        [DataRow(0), DataRow(1), DataRow(200), DataRow(201)]
        public void RealNormalReadPreservesExactRequestsIdentitiesDefinitionsAndExports(int count)
        {
            var f = new Fixture(count); var operation = new DataverseComponentDefinitionOperation();
            var oldCounter = new DataverseRequestCounter(); var newCounter = new DataverseRequestCounter();
            var old = operation.ReadAndResolveWithCache(f.Service(), "Source", "edu", CancellationToken.None, null, oldCounter);
            var before = f.Ledger.ToArray(); f.Ledger.Clear();
            var current = operation.ReadAndResolveWithCache(f.Service(), "Source", "edu", CancellationToken.None, null, newCounter, DisplayEnvironmentSide.Source);
            CollectionAssert.AreEqual(before, f.Ledger.ToArray()); Assert.AreEqual(oldCounter.TotalRequests, newCounter.TotalRequests);
            Assert.AreEqual(1, newCounter.GetExecuteCount("WhoAmI")); Assert.AreEqual(0, f.ServiceInstance.WriteCalls);
            CollectionAssert.AreEqual(Signatures(old.Definitions), Signatures(current.Definitions));
            CollectionAssert.AreEqual(old.Definitions.Definitions.Select(DefinitionSignature).ToArray(), current.Definitions.Definitions.Select(DefinitionSignature).ToArray());
            // Independent executions legitimately capture at different times. Freeze only that
            // audit timestamp; all exported semantic fields and diagnostics must match byte for byte.
            var fixedTime = new DateTimeOffset(2026, 10, 9, 0, 0, 0, TimeSpan.Zero);
            var oldExport = MembershipSnapshot.Complete(old.Definitions.Membership.Solution, old.Definitions.Membership.Components, fixedTime);
            var newExport = MembershipSnapshot.Complete(current.Definitions.Membership.Solution, current.Definitions.Membership.Components, fixedTime);
            Assert.AreEqual(new MembershipCoverageCsvExporter().CreateCsv(Present(oldExport)), new MembershipCoverageCsvExporter().CreateCsv(Present(newExport)));
            Assert.AreEqual(count == 0 ? ParentCaptureDisposition.NotApplicable : ParentCaptureDisposition.Completed, current.ParentCapture.Disposition);
            Assert.AreSame(current.Definitions.Membership, current.ParentCapture.Owner); Assert.IsNull(old.ParentCapture);
            if (count > 0)
            {
                Assert.AreEqual(count, current.ParentCapture.Observations.Count);
                Assert.AreEqual((count + 199) / 200, newCounter.GetQueryCount("appsetting"));
                Assert.AreEqual(1, newCounter.GetQueryCount("appmodule"));
                var hierarchy = ComponentHierarchyProjector.Project(Present(current.Definitions.Membership), Retain(current));
                Assert.AreEqual(count, hierarchy.Occurrences.Count(o => o.Record.ComponentType == 10075 && o.Placement == HierarchyPlacement.VerifiedStructuralParent));
            }
            TestContext.WriteLine("count={0}; old_requests={1}; emitted_requests={2}; ledger={3}", count, oldCounter.TotalRequests, newCounter.TotalRequests, string.Join(";", before));
        }

        [DataTestMethod]
        [DataRow("appsetting", "missing"), DataRow("appsetting", "duplicate"), DataRow("appsetting", "conflict")]
        [DataRow("appsetting", "paged"), DataRow("appsetting", "fault"), DataRow("appsetting", "null")]
        [DataRow("settingdefinition", "missing"), DataRow("settingdefinition", "duplicate"), DataRow("settingdefinition", "conflict")]
        [DataRow("settingdefinition", "paged"), DataRow("settingdefinition", "fault")]
        [DataRow("appmodule", "missing"), DataRow("appmodule", "duplicate"), DataRow("appmodule", "conflict")]
        [DataRow("appmodule", "paged"), DataRow("appmodule", "fault"), DataRow("structure", "fault")]
        public void FailuresPreserveNormalResultsAndNeverCreateTrustedPartialLinks(string table, string failure)
        {
            var f = new Fixture(2) { FailureTable = table, Failure = failure }; var operation = new DataverseComponentDefinitionOperation();
            var old = operation.ReadAndResolveWithCache(f.Service(), "Source", "edu", CancellationToken.None, null);
            var ledger = f.Ledger.ToArray(); f.Ledger.Clear();
            var current = operation.ReadAndResolveWithCache(f.Service(), "Source", "edu", CancellationToken.None, null, null, DisplayEnvironmentSide.Source);
            CollectionAssert.AreEqual(ledger, f.Ledger.ToArray()); CollectionAssert.AreEqual(Signatures(old.Definitions), Signatures(current.Definitions));
            CollectionAssert.AreEqual(old.Definitions.Definitions.Select(DefinitionSignature).ToArray(), current.Definitions.Definitions.Select(DefinitionSignature).ToArray());
            var parents = Retain(current);
            Assert.IsTrue(table == "structure" ? parents.Observations.All(o => o.Verified) : parents == null || parents.Observations.All(o => !o.Verified));
            if (table == "structure") Assert.AreEqual(ParentCaptureDisposition.Completed, current.ParentCapture.Disposition);
            Assert.IsFalse(current.ParentCapture.DiagnosticCode.Contains(Secret)); Assert.AreEqual(0, f.ServiceInstance.WriteCalls);
        }

        [DataTestMethod]
        [DataRow("blankparent"), DataRow("blankdefinition"), DataRow("wrongparentlookup"), DataRow("wrongdefinitionlookup")]
        [DataRow("parentname"), DataRow("definitionname"), DataRow("noobject"), DataRow("emptyobject")]
        [DataRow("wrongregistration"), DataRow("otherraw")]
        public void InvalidScalarsCannotManufactureConfirmedLinks(string state)
        {
            var f = new Fixture(1);
            if (state == "blankparent") f.Settings[0].Attributes.Remove("parentappmoduleid");
            if (state == "blankdefinition") f.Settings[0].Attributes.Remove("settingdefinitionid");
            if (state == "wrongparentlookup") f.Settings[0]["parentappmoduleid"] = new EntityReference("other", f.Parent);
            if (state == "wrongdefinitionlookup") f.Settings[0]["settingdefinitionid"] = Guid.NewGuid();
            if (state == "parentname") f.App["uniquename"] = " ";
            if (state == "definitionname") f.Definitions[0]["uniquename"] = " ";
            if (state == "noobject") f.Raw[1].Attributes.Remove("objectid");
            if (state == "emptyobject") f.Raw[1]["objectid"] = Guid.Empty;
            if (state == "wrongregistration") f.RegisteredName = "Other";
            if (state == "otherraw") f.Raw[1]["componenttype"] = new OptionSetValue(54321);
            var read = new DataverseComponentDefinitionOperation().ReadAndResolveWithCache(f.Service(), "Source", "edu", CancellationToken.None, null, null, DisplayEnvironmentSide.Source);
            var parents = Retain(read); Assert.IsTrue(parents == null || parents.Observations.All(o => !o.Verified));
        }

        [DataTestMethod]
        [DataRow("Snapshot"), DataRow("Role"), DataRow("Invocation"), DataRow("Read")]
        [DataRow("Organization"), DataRow("SolutionId"), DataRow("SolutionName")]
        public void ExactOwnershipRejectsCopiedAndStaleEvidence(string mismatch)
        {
            var f = new Fixture(1); var operation = new DataverseComponentDefinitionOperation();
            var read = operation.ReadAndResolveWithCache(f.Service(), "Source", "edu", CancellationToken.None, null, null, DisplayEnvironmentSide.Source);
            var capture = read.ParentCapture; var owner = capture.Owner; var side = capture.Side;
            var invocation = capture.ComparisonInvocationOwner; var token = capture.ReadOwner;
            if (mismatch == "Snapshot") owner = MembershipSnapshot.Complete(owner.Solution, owner.Components, owner.CapturedAt);
            if (mismatch == "Role") side = DisplayEnvironmentSide.Target;
            if (mismatch == "Invocation") invocation = new object();
            if (mismatch == "Read") token = new object();
            if (mismatch == "Organization" || mismatch == "SolutionId" || mismatch == "SolutionName")
                capture = new AppSettingParentCapture(owner, side, mismatch == "Organization" ? Guid.NewGuid() : capture.OrganizationId,
                    mismatch == "SolutionId" ? Guid.NewGuid() : capture.SolutionId, mismatch == "SolutionName" ? "other" : capture.SolutionUniqueName,
                    invocation, token, capture.Disposition, capture.Observations);
            Assert.IsNull(ComponentHierarchyProjector.RetainAppSettingParentsFromNormalRead(capture, owner, side, invocation, token));
        }

        [TestMethod]
        public void SuccessiveReadsAndRolesRemainIsolatedEvenWithinSameEnvironment()
        {
            var f = new Fixture(1); var operation = new DataverseComponentDefinitionOperation();
            var source = operation.ReadAndResolveWithCache(f.Service(), "same", "edu", CancellationToken.None, null, null, DisplayEnvironmentSide.Source);
            var stale = operation.ReadAndResolveWithCache(f.Service(), "same", "edu", CancellationToken.None, null, null, DisplayEnvironmentSide.Source);
            var target = operation.ReadAndResolveWithCache(f.Service(), "same", "edu", CancellationToken.None, null, null, DisplayEnvironmentSide.Target);
            Assert.AreNotSame(source.ReadOwner, stale.ReadOwner); Assert.AreNotSame(source.ReadOwner, target.ReadOwner);
            Assert.IsFalse(source.EvidenceBelongsTo(operation, stale.Definitions.Membership, DisplayEnvironmentSide.Source));
            Assert.IsFalse(source.EvidenceBelongsTo(operation, source.Definitions.Membership, DisplayEnvironmentSide.Target));
            var p = Present(source.Definitions.Membership, target.Definitions.Membership);
            var completed = SolutionComparerControl.CompleteMembershipComparison(p, operation, source, target, CancellationToken.None);
            Assert.IsNotNull(completed.SourceHierarchyParents); Assert.IsNotNull(completed.TargetHierarchyParents); Assert.AreSame(p, completed.Presentation);
            var swapped = SolutionComparerControl.CompleteMembershipComparison(p, operation, target, source, CancellationToken.None);
            Assert.IsNull(swapped.SourceHierarchyParents); Assert.IsNull(swapped.TargetHierarchyParents);
            var noHandoff = SolutionComparerControl.CompleteMembershipComparison(p, operation, null, null, CancellationToken.None);
            Assert.IsNull(noHandoff.SourceHierarchyParents); Assert.IsNull(noHandoff.TargetHierarchyParents);
        }

        [DataTestMethod]
        [DataRow("Reference"), DataRow("SolutionComponentId"), DataRow("ComponentType"), DataRow("ObjectId")]
        [DataRow("RootId"), DataRow("RootBehavior"), DataRow("IsMetadata"), DataRow("Order")]
        [DataRow("RepeatedObject"), DataRow("RepeatedRecord")]
        public void BindingRejectsRawReplacementReorderingAndRepeatedOccurrences(string mutation)
        {
            var f = Local(2); var items = f.Raw.Components.ToList(); var record = items[1].Record;
            if (mutation == "Order") items.Reverse();
            else if (mutation == "RepeatedRecord") items[2] = items[1];
            else items[1] = SettingIdentity(new SolutionComponentRecord(
                mutation == "SolutionComponentId" ? Guid.NewGuid() : record.SolutionComponentId,
                mutation == "ComponentType" ? 36 : record.ComponentType,
                mutation == "ObjectId" ? Guid.NewGuid() : mutation == "RepeatedObject" ? items[2].Record.ObjectId : record.ObjectId,
                mutation == "RootBehavior" ? 1 : record.RootComponentBehavior,
                mutation == "RootId" ? Guid.NewGuid() : record.RootSolutionComponentId,
                mutation == "IsMetadata" ? true : record.IsMetadata));
            var final = MembershipSnapshot.Complete(f.Raw.Solution, items, f.Raw.CapturedAt);
            var capture = f.Emission.SealAgainstFinalSnapshot(final, CancellationToken.None);
            Assert.AreEqual(ParentCaptureDisposition.Incomplete, capture.Disposition); Assert.AreEqual(0, capture.Observations.Count);
        }

        [TestMethod]
        public void SnapshotReplacementBindsOnlyFinalOwnerAndCannotBeResealed()
        {
            var f = Local(1); var intermediate = MembershipSnapshot.Complete(f.Raw.Solution, f.Raw.Components, f.Raw.CapturedAt);
            var context = new DataverseReadContext(Service(f.Raw.Solution), f.Raw.Environment, CancellationToken.None);
            var final = new Type92MembershipResolver().Apply(context, intermediate, CancellationToken.None);
            var capture = f.Emission.SealAgainstFinalSnapshot(final, CancellationToken.None);
            Assert.AreEqual(ParentCaptureDisposition.Completed, capture.Disposition); Assert.AreSame(final, capture.Owner); Assert.AreNotSame(f.Raw, capture.Owner);
            Assert.IsNull(ComponentHierarchyProjector.RetainAppSettingParentsFromNormalRead(capture, f.Raw, capture.Side, f.Invocation, f.Read));
            Assert.IsNotNull(ComponentHierarchyProjector.RetainAppSettingParentsFromNormalRead(capture, final, capture.Side, f.Invocation, f.Read));
            Assert.IsNull(f.Emission.SealAgainstFinalSnapshot(final, CancellationToken.None));
        }

        [DataTestMethod]
        [DataRow(0), DataRow(4), DataRow(5), DataRow(6), DataRow(7)]
        public void FailedSourceEventsSuppressAllPartialLinks(int state)
        {
            var f = Local(2, false); EmitGood(f.Emission, f.Raw.Components[1].Record, f.Parent, f.Definition);
            f.Emission.Observe(new AppSettingReadObservation(AppSettingReadKind.Child, f.Raw.Components[2].Record.ObjectId.Value, (NormalReadDisposition)state, 0));
            var capture = f.Emission.SealAgainstFinalSnapshot(f.Raw, CancellationToken.None);
            Assert.AreEqual(ParentCaptureDisposition.Incomplete, capture.Disposition); Assert.AreEqual(0, capture.Observations.Count);
        }

        [TestMethod]
        public void EmissionFailureAndConflictingEventsDoNotAlterNormalResolution()
        {
            var f = Local(1); f.Emission.Observe(new AppSettingReadObservation(AppSettingReadKind.Parent, f.Parent, NormalReadDisposition.Unique, 1, f.Parent, "other"));
            Assert.AreEqual(ParentCaptureDisposition.Faulted, f.Emission.SealAgainstFinalSnapshot(f.Raw, CancellationToken.None).Disposition);
            var live = new Fixture(1); var context = new DataverseReadContext(live.Service(), live.Solution.Environment, CancellationToken.None);
            var raw = MembershipSnapshot.Complete(live.Solution, live.Raw.Select(r => new ComponentIdentity(ToRecord(r), IdentityResolutionStatus.Unresolved)), DateTimeOffset.UtcNow);
            var result = new DataverseComponentIdentityResolver().ResolveSnapshot(context, raw, CancellationToken.None, new ThrowingSink());
            Assert.IsTrue(result.Components.All(i => i.Status == IdentityResolutionStatus.Resolved)); Assert.AreEqual(0, live.ServiceInstance.WriteCalls);
            var other = Local(1, false); AppSettingParentEvidenceEmission.Emit(other.Emission, () => { throw new InvalidOperationException(Secret); });
            Assert.AreEqual(ParentCaptureDisposition.Faulted, other.Emission.SealAgainstFinalSnapshot(other.Raw, CancellationToken.None).Disposition);
        }

        [DataTestMethod]
        [DataRow("solutioncomponent"), DataRow("appsetting"), DataRow("settingdefinition"), DataRow("appmodule"), DataRow("structure")]
        public void RequestedCancellationPreservesExistingBehavior(string stage)
        {
            foreach (var enabled in new[] { false, true })
            using (var cancel = new CancellationTokenSource())
            {
                var f = new Fixture(1) { Cancel = cancel, CancelStage = stage }; var operation = new DataverseComponentDefinitionOperation();
                Assert.ThrowsException<OperationCanceledException>(() =>
                {
                    if (enabled) operation.ReadAndResolveWithCache(f.Service(), "Source", "edu", cancel.Token, null, null, DisplayEnvironmentSide.Source);
                    else operation.ReadAndResolveWithCache(f.Service(), "Source", "edu", cancel.Token, null);
                });
                Assert.AreEqual(0, f.ServiceInstance.WriteCalls);
            }
        }

        [TestMethod]
        public void CancelledSealCannotPublishPartialCapture()
        { var f = Local(1); Assert.ThrowsException<OperationCanceledException>(() => f.Emission.SealAgainstFinalSnapshot(f.Raw, new CancellationToken(true))); }

        [DataTestMethod]
        [DataRow("Missing"), DataRow("Duplicate"), DataRow("Unresolved"), DataRow("Ambiguous"), DataRow("WrongId"), DataRow("WrongName")]
        public void ExistingSelectedParentGuardsRemainAuthoritative(string state)
        {
            var f = Local(1); var items = f.Raw.Components.ToList();
            if (state == "Missing") items.RemoveAt(0);
            else if (state == "Duplicate") items.Add(new ComponentIdentity(new SolutionComponentRecord(Guid.NewGuid(), 80, f.Parent), IdentityResolutionStatus.Resolved, "app"));
            else if (state == "Unresolved" || state == "Ambiguous") items[0] = new ComponentIdentity(items[0].Record, state == "Unresolved" ? IdentityResolutionStatus.Unresolved : IdentityResolutionStatus.Ambiguous);
            else items[0] = new ComponentIdentity(new SolutionComponentRecord(items[0].Record.SolutionComponentId, 80, state == "WrongId" ? Guid.NewGuid() : f.Parent), IdentityResolutionStatus.Resolved, state == "WrongName" ? "other" : "app");
            var snapshot = MembershipSnapshot.Complete(f.Raw.Solution, items, f.Raw.CapturedAt);
            var emitter = new AppSettingParentEvidenceEmission(snapshot, DisplayEnvironmentSide.Source, f.Invocation, f.Read);
            EmitGood(emitter, items.Single(i => i.Record.ComponentType == 10075).Record, f.Parent, f.Definition);
            var capture = emitter.SealAgainstFinalSnapshot(snapshot, CancellationToken.None);
            var parents = ComponentHierarchyProjector.RetainAppSettingParentsFromNormalRead(capture, snapshot, capture.Side, f.Invocation, f.Read);
            var projection = ComponentHierarchyProjector.Project(Present(snapshot), parents);
            Assert.IsFalse(projection.Occurrences.Single(i => i.Record.ComponentType == 10075).Associations.Any(a => a.Kind == HierarchyRelationshipKind.StructuralParent && a.Status == HierarchyAssociationStatus.Verified));
        }

        [TestMethod]
        public void ContractIsImmutableAndCannotRetainSdkPayloads()
        {
            foreach (var type in new[] { typeof(AppSettingParentCapture), typeof(AppSettingParentObservation), typeof(AppSettingReadObservation) })
            foreach (var property in type.GetProperties(BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic))
            {
                Assert.IsFalse(property.CanWrite, property.Name); Assert.IsFalse(typeof(Entity).IsAssignableFrom(property.PropertyType));
                Assert.IsFalse(typeof(IOrganizationService).IsAssignableFrom(property.PropertyType));
            }
            var f = Local(1); var capture = f.Emission.SealAgainstFinalSnapshot(f.Raw, CancellationToken.None);
            Assert.ThrowsException<NotSupportedException>(() => ((IList<AppSettingParentObservation>)capture.Observations).Clear());
            Assert.AreSame(f.Raw.Components[1].Record, capture.Observations[0].OriginalRecord);
        }

        [DataTestMethod]
        [DataRow("ChildPrimary"), DataRow("ParentPrimary"), DataRow("DefinitionPrimary")]
        [DataRow("ChildCount"), DataRow("ParentCount"), DataRow("DefinitionCount")]
        [DataRow("ParentLookup"), DataRow("DefinitionLookup"), DataRow("BlankParentName"), DataRow("BlankDefinitionName")]
        public void StrictAdapterRequiresActualChildDefinitionAndParentProof(string fault)
        {
            var f = Local(1, false); var record = f.Raw.Components[1].Record;
            f.Emission.Observe(new AppSettingReadObservation(AppSettingReadKind.Child, record.ObjectId.Value,
                NormalReadDisposition.Unique, fault == "ChildCount" ? 2 : 1,
                fault == "ChildPrimary" ? Guid.NewGuid() : record.ObjectId,
                parentId: f.Parent, parentTableVerified: fault != "ParentLookup",
                definitionId: f.Definition, definitionTableVerified: fault != "DefinitionLookup"));
            f.Emission.Observe(new AppSettingReadObservation(AppSettingReadKind.Parent, f.Parent, NormalReadDisposition.Unique,
                fault == "ParentCount" ? 2 : 1, fault == "ParentPrimary" ? Guid.NewGuid() : f.Parent,
                fault == "BlankParentName" ? " " : "app"));
            f.Emission.Observe(new AppSettingReadObservation(AppSettingReadKind.Definition, f.Definition, NormalReadDisposition.Unique,
                fault == "DefinitionCount" ? 2 : 1, fault == "DefinitionPrimary" ? Guid.NewGuid() : f.Definition,
                fault == "BlankDefinitionName" ? " " : "setting"));
            var capture = f.Emission.SealAgainstFinalSnapshot(f.Raw, CancellationToken.None);
            var retained = ComponentHierarchyProjector.RetainAppSettingParentsFromNormalRead(capture, f.Raw, capture.Side, f.Invocation, f.Read);
            Assert.IsTrue(retained == null || !retained.Observations.Single().Verified);
        }

        [TestMethod]
        public void WarmCacheIsNotHarvestedOrReadAgainToManufactureEvidence()
        {
            var f = new Fixture(1); var context = new DataverseReadContext(f.Service(), f.Solution.Environment, CancellationToken.None);
            var raw = MembershipSnapshot.Complete(f.Solution, f.Raw.Select(r => new ComponentIdentity(ToRecord(r), IdentityResolutionStatus.Unresolved)), DateTimeOffset.UtcNow);
            var resolver = new DataverseComponentIdentityResolver(); var resolved = resolver.ResolveSnapshot(context, raw, CancellationToken.None);
            var before = f.Ledger.Count; var emitter = new AppSettingParentEvidenceEmission(raw, DisplayEnvironmentSide.Source, new object(), new object());
            var again = resolver.ResolveSnapshot(context, resolved, CancellationToken.None, emitter);
            var capture = emitter.SealAgainstFinalSnapshot(again, CancellationToken.None);
            Assert.AreEqual(ParentCaptureDisposition.Incomplete, capture.Disposition); Assert.AreEqual(0, capture.Observations.Count);
            // Resolver may reread a selected AppModule exactly as it already did; it must not
            // reread cached AppSetting/SettingDefinition identities to satisfy this handoff.
            Assert.IsFalse(f.Ledger.Skip(before).Any(l => l.StartsWith("Q:appsetting:") || l.StartsWith("Q:settingdefinition:")));
            Assert.IsTrue(again.Components.All(i => i.Status == IdentityResolutionStatus.Resolved));
        }

        [DataTestMethod]
        [DataRow("RepeatedObject"), DataRow("RepeatedSCID"), DataRow("Rooted")]
        public void NormalRawDuplicationAndInclusionRootRemainUnchanged(string kind)
        {
            var f = new Fixture(1);
            if (kind == "RepeatedObject") f.Raw.Add(new Entity("solutioncomponent", Guid.NewGuid())
                { ["componenttype"] = new OptionSetValue(10075), ["objectid"] = f.Settings[0].Id });
            if (kind == "RepeatedSCID") f.Raw.Add(f.Raw[1]);
            if (kind == "Rooted") f.Raw[1]["rootsolutioncomponentid"] = f.Raw[0].Id;
            var operation = new DataverseComponentDefinitionOperation();
            var old = operation.ReadAndResolveWithCache(f.Service(), "Source", "edu", CancellationToken.None, null);
            var current = operation.ReadAndResolveWithCache(f.Service(), "Source", "edu", CancellationToken.None, null, null, DisplayEnvironmentSide.Source);
            CollectionAssert.AreEqual(Signatures(old.Definitions), Signatures(current.Definitions));
            if (kind != "Rooted")
            { Assert.AreEqual(ParentCaptureDisposition.Incomplete, current.ParentCapture.Disposition); Assert.IsNull(Retain(current)); }
            else
            {
                var projection = ComponentHierarchyProjector.Project(Present(current.Definitions.Membership), Retain(current));
                var child = projection.Occurrences.Single(o => o.Record.ComponentType == 10075);
                Assert.AreEqual(HierarchyPlacement.VerifiedInclusionRoot, child.Placement);
                Assert.IsTrue(child.Associations.Any(a => a.Kind == HierarchyRelationshipKind.StructuralParent && a.Status == HierarchyAssociationStatus.Verified));
            }
        }

        [TestMethod]
        public void DifferentLocalIdsAcrossSidesUseOnlyExistingAuthoritativeParentMatch()
        {
            var sourceFixture = new Fixture(1); var targetFixture = new Fixture(1); var operation = new DataverseComponentDefinitionOperation();
            var source = operation.ReadAndResolveWithCache(sourceFixture.Service(), "Source", "edu", CancellationToken.None, null, null, DisplayEnvironmentSide.Source);
            var target = operation.ReadAndResolveWithCache(targetFixture.Service(), "Target", "edu", CancellationToken.None, null, null, DisplayEnvironmentSide.Target);
            var p = Present(source.Definitions.Membership, target.Definitions.Membership);
            Assert.AreNotEqual(sourceFixture.Parent, targetFixture.Parent);
            Assert.AreNotEqual(sourceFixture.Settings[0].Id, targetFixture.Settings[0].Id);
            var completed = SolutionComparerControl.CompleteMembershipComparison(p, operation, source, target, CancellationToken.None);
            var projection = ComponentHierarchyProjector.Project(completed, completed.SourceHierarchyParents, completed.TargetHierarchyParents);
            Assert.AreEqual(2, p.Summary.PresentInBoth); Assert.AreEqual(1, projection.ParentCorrelations.Count);
            Assert.AreSame(p, completed.Presentation);
            for (int index = 0; index < p.Rows.Count; index++)
            { Assert.AreSame(p.Rows[index], completed.Presentation.Rows[index]); Assert.AreSame(p.Rows[index].DefinitionDetail, completed.Presentation.Rows[index].DefinitionDetail); }
        }

        [TestMethod]
        public void Type92ReplacementIsBoundAfterApplyWithoutChangingType92()
        {
            var f = Local(1, false); var items = f.Raw.Components.ToList();
            items.Add(new ComponentIdentity(new SolutionComponentRecord(Guid.NewGuid(), 92, null), IdentityResolutionStatus.Unresolved));
            var raw = MembershipSnapshot.Complete(f.Raw.Solution, items, f.Raw.CapturedAt);
            var emission = new AppSettingParentEvidenceEmission(raw, DisplayEnvironmentSide.Source, f.Invocation, f.Read);
            EmitGood(emission, items[1].Record, f.Parent, f.Definition);
            var service = Service(raw.Solution, query => Rows());
            var context = new DataverseReadContext(service, raw.Environment, CancellationToken.None);
            var intermediate = MembershipSnapshot.Complete(raw.Solution, items, raw.CapturedAt);
            var final = new Type92MembershipResolver().Apply(context, intermediate, CancellationToken.None);
            Assert.AreNotSame(intermediate, final);
            var capture = emission.SealAgainstFinalSnapshot(final, CancellationToken.None);
            Assert.AreSame(final, capture.Owner); Assert.AreEqual(ParentCaptureDisposition.Completed, capture.Disposition);
            Assert.IsNull(ComponentHierarchyProjector.RetainAppSettingParentsFromNormalRead(capture, intermediate, capture.Side, f.Invocation, f.Read));
            Assert.IsNotNull(ComponentHierarchyProjector.RetainAppSettingParentsFromNormalRead(capture, final, capture.Side, f.Invocation, f.Read));
        }

        [DataTestMethod]
        [DataRow(840), DataRow(10000)]
        public void IncrementalCaptureMemoryExcludesRawSnapshotAndSdkPayloads(int count)
        {
            var f = Local(count, false); f.Emission = null; GC.Collect(); long before = GC.GetTotalMemory(true);
            var clock = Stopwatch.StartNew(); var emission = new AppSettingParentEvidenceEmission(f.Raw, DisplayEnvironmentSide.Source, f.Invocation, f.Read);
            emission.Observe(new AppSettingReadObservation(AppSettingReadKind.Parent, f.Parent, NormalReadDisposition.Unique, 1, f.Parent, "app"));
            emission.Observe(new AppSettingReadObservation(AppSettingReadKind.Definition, f.Definition, NormalReadDisposition.Unique, 1, f.Definition, "setting"));
            foreach (var row in f.Raw.Components.Skip(1)) emission.Observe(Child(row.Record, f.Parent, f.Definition));
            var capture = emission.SealAgainstFinalSnapshot(f.Raw, CancellationToken.None); emission = null;
            var retained = ComponentHierarchyProjector.RetainAppSettingParentsFromNormalRead(capture, f.Raw, capture.Side, f.Invocation, f.Read);
            clock.Stop(); long after = GC.GetTotalMemory(true);
            TestContext.WriteLine("H2B incremental count={0}; emit_seal_adapter_ms={1:F3}; retained_scalar_capture_and_H1_bytes={2}; bytes_per_setting={3:F1}", count, clock.Elapsed.TotalMilliseconds, after - before, (after - before) / (double)count);
            Assert.AreEqual(count, capture.Observations.Count); Assert.AreEqual(count, retained.Observations.Count);
            GC.KeepAlive(capture); GC.KeepAlive(retained); GC.KeepAlive(f.Raw);
        }

        [DataTestMethod]
        [DataRow(840), DataRow(10000)]
        public void SyntheticProducerSealAdapterHierarchyAndNativeUiOverhead(int count)
        {
            GC.Collect(); long start = GC.GetTotalMemory(true); var clock = Stopwatch.StartNew(); var f = Local(count);
            var emitMs = clock.Elapsed.TotalMilliseconds; clock.Restart(); var capture = f.Emission.SealAgainstFinalSnapshot(f.Raw, CancellationToken.None);
            var sealMs = clock.Elapsed.TotalMilliseconds; clock.Restart(); var parents = ComponentHierarchyProjector.RetainAppSettingParentsFromNormalRead(capture, f.Raw, capture.Side, f.Invocation, f.Read);
            var adapterMs = clock.Elapsed.TotalMilliseconds; long retained = GC.GetTotalMemory(true) - start;
            var p = Present(f.Raw); var csv = new MembershipCoverageCsvExporter().CreateCsv(p);
            clock.Restart(); var noHandoff = ComponentHierarchyProjector.Project(p); var oldMs = clock.Elapsed.TotalMilliseconds;
            clock.Restart(); var hierarchy = ComponentHierarchyProjector.Project(p, parents); var newMs = clock.Elapsed.TotalMilliseconds;
            Assert.AreEqual(count, hierarchy.Occurrences.Count(o => o.Placement == HierarchyPlacement.VerifiedStructuralParent));
            Assert.AreEqual(count + 1, noHandoff.Occurrences.Count); Assert.AreEqual(csv, new MembershipCoverageCsvExporter().CreateCsv(p));
            long nativeBytes = 0; double nativeMs = 0; Exception error = null;
            var thread = new Thread(() =>
            {
                try
                {
                    var display = ComponentDisplayEvidenceProjector.Project(p, null, null, CancellationToken.None);
                    var completed = new CompletedMembershipComparison(new CompletedMembershipComparison(p, display, DisplayProjectionOutcome.Completed), parents, null);
                    using (var form = new MembershipResultsForm(completed))
                    {
                        var before = GC.GetTotalMemory(true); var timer = Stopwatch.StartNew(); form.Show(); Application.DoEvents();
                        var selector = form.Controls.OfType<ToolStrip>().SelectMany(t => t.Items.OfType<ToolStripComboBox>()).First(c => c.Name == "MembershipViewSelector");
                        selector.SelectedIndex = 1; var timeout = Stopwatch.StartNew();
                        while (form.HierarchyProjection == null && timeout.Elapsed.TotalSeconds < 15) { Application.DoEvents(); Thread.Sleep(1); }
                        timer.Stop(); nativeMs = timer.Elapsed.TotalMilliseconds; nativeBytes = GC.GetTotalMemory(true) - before;
                        Assert.IsNotNull(form.HierarchyProjection); Assert.AreEqual(count + 1, form.HierarchyProjection.Occurrences.Count);
                        Assert.AreEqual(count, form.HierarchyProjection.Occurrences.Count(o => o.Placement == HierarchyPlacement.VerifiedStructuralParent));
                    }
                }
                catch (Exception ex) { error = ex; }
            });
            thread.SetApartmentState(ApartmentState.STA); thread.Start(); Assert.IsTrue(thread.Join(30000)); if (error != null) throw error;
            TestContext.WriteLine("H2B synthetic records={0}; producer_ms={1:F3}; seal_ms={2:F3}; adapter_ms={3:F3}; retained_fixture_capture_bytes={4}; no_handoff_hierarchy_ms={5:F3}; emitted_hierarchy_ms={6:F3}; native_open_hierarchy_ms={7:F3}; native_managed_bytes_delta={8}; requests=0; writes=0", count, emitMs, sealMs, adapterMs, retained, oldMs, newMs, nativeMs, nativeBytes);
        }

        [TestMethod]
        public void Current840AccountingAndNoHandoffRemainUnchanged()
        {
            var method = typeof(MembershipHierarchyUiTests).GetMethod("Baseline", BindingFlags.NonPublic | BindingFlags.Static);
            var p = (MembershipComparisonPresentation)method.Invoke(null, null);
            var csv = new MembershipCoverageCsvExporter().CreateCsv(p); var projection = ComponentHierarchyProjector.Project(p);
            Assert.AreEqual(840, p.Summary.TotalComparisonRows); Assert.AreEqual(672, p.Summary.PresentInBoth);
            Assert.AreEqual(103, p.Summary.SourceOnly); Assert.AreEqual(32, p.Summary.TargetOnly);
            Assert.AreEqual(26, p.Summary.IndeterminateUnsupported); Assert.AreEqual(4, p.Summary.Ambiguous); Assert.AreEqual(3, p.Summary.IndeterminateResolved);
            Assert.AreEqual(csv, new MembershipCoverageCsvExporter().CreateCsv(p)); Assert.AreSame(p, projection.Presentation);
        }

        private sealed class ThrowingSink : IAppSettingParentEvidenceSink
        { public void Observe(AppSettingReadObservation observation) { throw new InvalidOperationException(Secret); } }
        private sealed class LocalFixture
        {
            internal MembershipSnapshot Raw; internal Guid Parent, Definition;
            internal readonly object Invocation = new object(), Read = new object(); internal AppSettingParentEvidenceEmission Emission;
        }
        private static LocalFixture Local(int count, bool emit = true)
        {
            var f = new LocalFixture { Parent = Guid.NewGuid(), Definition = Guid.NewGuid() };
            var items = new List<ComponentIdentity> { new ComponentIdentity(new SolutionComponentRecord(Guid.NewGuid(), 80, f.Parent), IdentityResolutionStatus.Resolved, "app") };
            for (int index = 0; index < count; index++) items.Add(SettingIdentity(new SolutionComponentRecord(Guid.NewGuid(), 10075, Guid.NewGuid())));
            f.Raw = Snapshot(items.ToArray()); f.Emission = new AppSettingParentEvidenceEmission(f.Raw, DisplayEnvironmentSide.Source, f.Invocation, f.Read);
            if (emit)
            {
                f.Emission.Observe(new AppSettingReadObservation(AppSettingReadKind.Parent, f.Parent, NormalReadDisposition.Unique, 1, f.Parent, "app"));
                f.Emission.Observe(new AppSettingReadObservation(AppSettingReadKind.Definition, f.Definition, NormalReadDisposition.Unique, 1, f.Definition, "setting"));
                foreach (var identity in items.Skip(1)) f.Emission.Observe(Child(identity.Record, f.Parent, f.Definition));
            }
            return f;
        }
        private static ComponentIdentity SettingIdentity(SolutionComponentRecord record) => new ComponentIdentity(record, IdentityResolutionStatus.Resolved,
            AppSettingResolutionOperation.PortableKey("app", "setting"), componentTypeKey: ComponentSemanticKinds.AppSetting,
            registeredDefinition: record.ComponentType == 10075 ? new SolutionComponentDefinitionIdentity(10075, "AppSetting", "appsetting") : null);
        private static AppSettingReadObservation Child(SolutionComponentRecord record, Guid parent, Guid definition) => new AppSettingReadObservation(
            AppSettingReadKind.Child, record.ObjectId.Value, NormalReadDisposition.Unique, 1, record.ObjectId, parentId: parent, parentTableVerified: true, definitionId: definition, definitionTableVerified: true);
        private static void EmitGood(AppSettingParentEvidenceEmission sink, SolutionComponentRecord record, Guid parent, Guid definition)
        { sink.Observe(Child(record, parent, definition)); sink.Observe(new AppSettingReadObservation(AppSettingReadKind.Parent, parent, NormalReadDisposition.Unique, 1, parent, "app")); sink.Observe(new AppSettingReadObservation(AppSettingReadKind.Definition, definition, NormalReadDisposition.Unique, 1, definition, "setting")); }
        private static ComponentHierarchyParentEvidence Retain(DataverseComponentDefinitionOperation.DefinitionReadWithCacheResult read) => ComponentHierarchyProjector.RetainAppSettingParentsFromNormalRead(read.ParentCapture, read.Definitions.Membership, read.Side.Value, read.ComparisonInvocationOwner, read.ReadOwner);
        private static MembershipComparisonPresentation Present(MembershipSnapshot source, MembershipSnapshot target = null) => new MembershipResultPresenter().Create(MembershipEnvironmentResult.FromSnapshot("Source", source, 0, TimeSpan.Zero), MembershipEnvironmentResult.FromSnapshot("Target", target ?? MembershipSnapshot.Complete(source.Solution, new ComponentIdentity[0], source.CapturedAt), 0, TimeSpan.Zero));
        private static string[] Signatures(ComponentDefinitionSnapshot data) => data.Membership.Components.Select(i =>
            i.Record.SolutionComponentId + "|" + i.Record.ObjectId + "|" + i.Record.ComponentType + "|" +
            i.Record.RootSolutionComponentId + "|" + i.Record.RootComponentBehavior + "|" + i.Record.IsMetadata + "|" +
            i.Status + "|" + i.ComparisonKey + "|" + i.ComponentTypeKey + "|" + i.SemanticKind + "|" +
            i.BlockerPortableIdentity + "|" + i.BlockerScope + "|" + i.InventoryAbsencePolicy + "|" +
            i.Diagnostic + "|" + string.Join(";", i.DiagnosticEvidence)).ToArray();
        private static string DefinitionSignature(ComponentDefinition d) => d.Status + "|" + d.Diagnostic + "|" +
            string.Join(";", d.DiagnosticEvidence) + "|" + string.Join(";", d.ComparableProperties.OrderBy(p => p.Key).Select(p => p.Key + "=" + p.Value));
        private static SolutionComponentRecord ToRecord(Entity row) => new SolutionComponentRecord(row.Id, row.GetAttributeValue<OptionSetValue>("componenttype").Value, row.GetAttributeValue<Guid>("objectid"));

        private sealed class Fixture
        {
            internal readonly SolutionIdentity Solution = MembershipTestData.Solution("edu"); internal readonly Guid Parent = Guid.NewGuid();
            internal readonly List<Entity> Raw = new List<Entity>(), Settings = new List<Entity>(), Definitions = new List<Entity>(); internal readonly List<string> Ledger = new List<string>();
            internal Entity App; internal string FailureTable, Failure, CancelStage, RegisteredName = "AppSetting"; internal CancellationTokenSource Cancel; internal FakeOrganizationService ServiceInstance;
            internal Fixture(int count)
            {
                App = new Entity("appmodule", Parent) { ["appmoduleid"] = Parent, ["uniquename"] = "app", ["name"] = "Application" };
                if (count > 0) Raw.Add(new Entity("solutioncomponent", Guid.NewGuid()) { ["componenttype"] = new OptionSetValue(80), ["objectid"] = Parent });
                for (int n = 0; n < count; n++)
                {
                    var id = Guid.NewGuid(); var definition = Guid.NewGuid(); Raw.Add(new Entity("solutioncomponent", Guid.NewGuid()) { ["componenttype"] = new OptionSetValue(10075), ["objectid"] = id });
                    Settings.Add(new Entity("appsetting", id) { ["appsettingid"] = id, ["parentappmoduleid"] = new EntityReference("appmodule", Parent), ["settingdefinitionid"] = new EntityReference("settingdefinition", definition), ["value"] = Secret });
                    Definitions.Add(new Entity("settingdefinition", definition) { ["settingdefinitionid"] = definition, ["uniquename"] = "setting" + n, ["datatype"] = new OptionSetValue(1), ["isoverridable"] = true, ["overridablelevel"] = new OptionSetValue(1), ["releaselevel"] = new OptionSetValue(1), ["defaultvalue"] = Secret });
                }
            }
            internal FakeOrganizationService Service() => ServiceInstance = new FakeOrganizationService { RetrievePage = Query, ExecuteRequest = Execute };
            private EntityCollection Query(QueryExpression query)
            {
                Ledger.Add("Q:" + query.EntityName + ":" + string.Join(",", query.ColumnSet.Columns) + ":" + query.TopCount + ":" + query.PageInfo?.PageNumber + ":" + query.PageInfo?.Count + ":" + string.Join(";", query.Criteria.Conditions.Select(c => c.AttributeName + "=" + c.Operator + ":" + string.Join(",", c.Values))));
                var stage = query.EntityName == "settingdefinition" && query.ColumnSet.Columns.Contains("datatype") ? "structure" : query.EntityName;
                if (CancelStage == stage) Cancel.Cancel();
                if (query.EntityName == "solution") return Rows(new Entity("solution", Solution.SolutionId) { ["uniquename"] = "edu" });
                if (query.EntityName == "solutioncomponent") return Rows(Raw.ToArray());
                if (query.EntityName == "solutioncomponentdefinition") return query.Criteria.Conditions.Any(c => c.AttributeName == "primaryentityname") ? Rows(new Entity("solutioncomponentdefinition") { ["objecttypecode"] = 19999 }) : Rows(new Entity("solutioncomponentdefinition") { ["objecttypecode"] = 10075, ["name"] = RegisteredName, ["primaryentityname"] = "appsetting" });
                var ids = query.Criteria.Conditions.Single().Values.Cast<Guid>().ToList();
                var data = (query.EntityName == "appmodule" ? new[] { App } : query.EntityName == "appsetting" ? Settings.ToArray() : Definitions.ToArray()).Where(r => ids.Contains(r.Id)).Select(r =>
                { var copy = new Entity(r.LogicalName, r.Id); foreach (var field in r.Attributes) copy[field.Key] = field.Value; return copy; }).ToList();
                if (FailureTable == stage)
                { if (Failure == "fault") throw new FaultException(Secret); if (Failure == "null") return null; if (Failure == "missing") data.Clear(); if (Failure == "duplicate") data.AddRange(data.ToArray()); if (Failure == "conflict") data[0][query.EntityName + "id"] = Guid.NewGuid(); }
                var result = Rows(data.ToArray()); if (FailureTable == stage && Failure == "paged") result.MoreRecords = true; return result;
            }
            private OrganizationResponse Execute(OrganizationRequest request)
            {
                Ledger.Add("E:" + request.RequestName + ":" + (request is RetrieveEntityRequest e ? e.LogicalName + ":" + e.EntityFilters + ":" + e.RetrieveAsIfPublished : ""));
                if (request is WhoAmIRequest) return WhoAmI(Solution.Environment.OrganizationId);
                if (request is RetrieveEntityRequest)
                {
                    var metadata = new EntityMetadata { LogicalName = "settingdefinition" };
                    var fields = AppSettingResolutionOperation.StructuralFields.Select(field => field == "isoverridable" ? (AttributeMetadata)new BooleanAttributeMetadata { LogicalName = field } : new PicklistAttributeMetadata { LogicalName = field }).ToArray();
                    foreach (var field in fields) typeof(AttributeMetadata).GetProperty("IsValidForRead").SetValue(field, true, null);
                    typeof(EntityMetadata).GetProperty("PrimaryIdAttribute").SetValue(metadata, "settingdefinitionid", null); typeof(EntityMetadata).GetProperty("Attributes").SetValue(metadata, fields, null);
                    return new RetrieveEntityResponse { Results = new ParameterCollection { ["EntityMetadata"] = metadata } };
                }
                throw new FaultException("Synthetic metadata unavailable");
            }
        }
    }
}
