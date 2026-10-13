using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Reflection;
using System.Threading;
using D365SolutionComparer.Models.Identity;
using D365SolutionComparer.Models.Membership;
using D365SolutionComparer.Models.Presentation;
using D365SolutionComparer.Services.Membership;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace D365SolutionComparer.Tests
{
    [TestClass, TestCategory("H1Hierarchy")]
    public class ComponentHierarchyProjectionTests
    {
        public TestContext TestContext { get; set; }

        [DataTestMethod]
        [DataRow(2), DataRow(10), DataRow(26), DataRow(59), DataRow(60)]
        public void VerifiedRootLinkedTableSubcomponentsKeepOriginalClassification(int type)
        {
            var table = Identity(1, 1, "account"); var child = Identity(2, type, "child", table.Record.SolutionComponentId);
            var p = Present(new[] { table, child }); var csv = new MembershipCoverageCsvExporter().CreateCsv(p);
            var result = ComponentHierarchyProjector.Project(p); var node = Node(result, child);
            Assert.AreEqual(HierarchyPlacement.VerifiedInclusionRoot, node.Placement);
            Assert.AreEqual(Node(result, table).OccurrenceIndex, node.Associations.Single().ParentOccurrenceIndex);
            Assert.AreSame(child, node.Identity); Assert.AreSame(child.Record, node.Record);
            Assert.AreEqual(csv, new MembershipCoverageCsvExporter().CreateCsv(p));
            Assert.AreEqual(HierarchyPlacement.ParentNotEstablished, Node(result, table).Placement);
            AssertLossless(result);
        }

        [DataTestMethod]
        [DataRow("missing", HierarchyAssociationStatus.MissingParent)]
        [DataRow("duplicate", HierarchyAssociationStatus.DuplicateParent)]
        [DataRow("self", HierarchyAssociationStatus.SelfReference)]
        [DataRow("empty", HierarchyAssociationStatus.InvalidReference)]
        [DataRow("cycle", HierarchyAssociationStatus.CyclicReference)]
        [DataRow("ancestor", HierarchyAssociationStatus.UnverifiedAncestor)]
        public void InvalidParentReferencesRemainExplicitAndLossless(string mode, HierarchyAssociationStatus expected)
        {
            var child = Identity(1, 2, "a.c", mode == "self" ? Id(1) : mode == "empty" ? Guid.Empty : Id(2));
            var items = new List<ComponentIdentity> { child };
            if (mode == "duplicate") { items.Add(Identity(2, 1, "a")); items.Add(Identity(2, 1, "b", objectId: Id(80002))); }
            if (mode == "cycle") items.Add(Identity(2, 1, "a", Id(1)));
            if (mode == "ancestor") items.Add(Identity(2, 1, "a", Id(900)));
            var result = ComponentHierarchyProjector.Project(Present(items)); var node = Node(result, child);
            Assert.AreEqual(expected, node.Associations.Single().Status);
            Assert.AreEqual(HierarchyPlacement.ParentNotEstablished, node.Placement); AssertLossless(result);
        }

        [TestMethod]
        public void RootCannotResolveAgainstAnotherEnvironmentEvenWhenLocalGuidsOverlap()
        {
            var child = Identity(1, 2, "account.c", Id(2)); var targetTable = Identity(2, 1, "account");
            var result = ComponentHierarchyProjector.Project(Present(new[] { child }, new[] { targetTable }));
            Assert.AreEqual(HierarchyAssociationStatus.MissingParent, Node(result, child).Associations.Single().Status);
            Assert.AreEqual(0, result.ParentCorrelations.Count); AssertLossless(result);
        }

        [TestMethod]
        public void PopulatedBehaviorAndMetadataFlagNeverCreateTopLevelOrParentProof()
        {
            var items = new[] { Identity(1, 1, "a", behavior: 0), Identity(2, 1, "b", behavior: 1), Identity(3, 1, "c", behavior: 2) };
            var result = ComponentHierarchyProjector.Project(Present(items));
            Assert.IsTrue(result.Occurrences.All(n => n.Placement == HierarchyPlacement.ParentNotEstablished));
            Assert.IsTrue(result.Occurrences.All(n => n.Record.IsMetadata == true)); AssertLossless(result);
        }

        [DataTestMethod]
        [DataRow(false), DataRow(true)]
        public void AppSettingStructuralAndRootEvidenceRemainSeparate(bool rootLinked)
        {
            var app = Identity(1, 80, "app"); var setting = Identity(2, 10075, null,
                rootLinked ? (Guid?)app.Record.SolutionComponentId : null, IdentityResolutionStatus.Unsupported);
            var p = Present(new[] { app, setting });
            var evidence = ComponentHierarchyProjector.RetainAppSettingParents(p.Source.Snapshot,
                DisplayEnvironmentSide.Source, AppReport(p.Source.Snapshot, setting, app));
            var result = ComponentHierarchyProjector.Project(p, evidence); var node = Node(result, setting);
            Assert.AreEqual(rootLinked ? HierarchyPlacement.VerifiedInclusionRoot : HierarchyPlacement.VerifiedStructuralParent, node.Placement);
            Assert.AreEqual(2, node.Associations.Count);
            Assert.AreEqual(HierarchyRelationshipKind.StructuralParent, node.Associations[1].Kind);
            Assert.AreEqual(HierarchyAssociationStatus.Verified, node.Associations[1].Status);
            Assert.AreEqual(IdentityResolutionStatus.Unsupported, node.Identity.Status);
            Assert.AreEqual("Indeterminate - Unsupported", result.ComparisonRows.Single(l => l.Row.SourceIdentity == setting).Row.MembershipStatus);
            AssertLossless(result);
        }

        [DataTestMethod]
        [DataRow("scope"), DataRow("environment"), DataRow("primary"), DataRow("reference"), DataRow("correlation"), DataRow("state")]
        [DataRow("root"), DataRow("metadata")]
        public void InvalidAppSettingEvidenceCannotEstablishStructuralParent(string mode)
        {
            var app = Identity(1, 80, "app"); var child = Identity(2, 10075, "setting"); var p = Present(new[] { app, child });
            var report = AppReport(p.Source.Snapshot, child, app, mode);
            var evidence = ComponentHierarchyProjector.RetainAppSettingParents(p.Source.Snapshot, DisplayEnvironmentSide.Source, report);
            var result = ComponentHierarchyProjector.Project(p, evidence);
            Assert.AreEqual(HierarchyPlacement.ParentNotEstablished, Node(result, child).Placement);
            Assert.IsFalse(Node(result, child).Associations.Any(a => a.Kind == HierarchyRelationshipKind.StructuralParent && a.Status == HierarchyAssociationStatus.Verified));
            AssertLossless(result);
        }

        [DataTestMethod]
        [DataRow("clone"), DataRow("role")]
        public void RetainedEvidenceRequiresExactSnapshotAndRole(string mode)
        {
            var app = Identity(1, 80, "app"); var child = Identity(2, 10075, "setting"); var p = Present(new[] { app, child });
            var owner = mode == "clone" ? Snapshot(new[] { app, child }, false) : p.Source.Snapshot;
            var evidence = ComponentHierarchyProjector.RetainAppSettingParents(owner,
                mode == "role" ? DisplayEnvironmentSide.Target : DisplayEnvironmentSide.Source, AppReport(owner, child, app));
            var result = ComponentHierarchyProjector.Project(p, evidence);
            Assert.AreEqual(HierarchyPlacement.ParentNotEstablished, Node(result, child).Placement);
            Assert.IsTrue(result.Diagnostics.Any(d => d.Contains("scope rejected"))); AssertLossless(result);
        }

        [TestMethod]
        public void ParentAmbiguityDoesNotEraseLocalRootOrCreatePortableCorrelation()
        {
            var parent = Identity(1, 1, null, status: IdentityResolutionStatus.Ambiguous);
            var child = Identity(2, 2, "account.c", parent.Record.SolutionComponentId);
            var result = ComponentHierarchyProjector.Project(Present(new[] { parent, child }, new[] { Identity(3, 1, "account") }));
            Assert.AreEqual(HierarchyPlacement.VerifiedInclusionRoot, Node(result, child).Placement);
            Assert.AreEqual(IdentityResolutionStatus.Ambiguous, Node(result, parent).Identity.Status);
            Assert.AreEqual(0, result.ParentCorrelations.Count); AssertLossless(result);
        }

        [TestMethod]
        public void StructuralParentMustBeUniquelyResolvedInSelectedMembership()
        {
            var app = Identity(1, 80, "app"); var child = Identity(2, 10075, "setting"); var duplicate = Identity(3, 80, "app");
            var p = Present(new[] { app, child, duplicate });
            var result = ComponentHierarchyProjector.Project(p, ComponentHierarchyProjector.RetainAppSettingParents(
                p.Source.Snapshot, DisplayEnvironmentSide.Source, AppReport(p.Source.Snapshot, child, app)));
            Assert.AreEqual(HierarchyAssociationStatus.DuplicateParent, Node(result, child).Associations.Last().Status);
            Assert.AreEqual(HierarchyPlacement.ParentNotEstablished, Node(result, child).Placement); AssertLossless(result);
        }

        [DataTestMethod]
        [DataRow(IdentityResolutionStatus.Ambiguous), DataRow(IdentityResolutionStatus.Unresolved), DataRow(IdentityResolutionStatus.Unsupported)]
        public void LocatedButUnresolvedParentIsNotMisreportedAsMissing(IdentityResolutionStatus status)
        {
            var app = Identity(1, 80, null, status: status); var child = Identity(2, 10075, "setting");
            var p = Present(new[] { app, child });
            var result = ComponentHierarchyProjector.Project(p, ComponentHierarchyProjector.RetainAppSettingParents(
                p.Source.Snapshot, DisplayEnvironmentSide.Source, AppReport(p.Source.Snapshot, child, app)));
            var node = Node(result, child);
            Assert.AreEqual(HierarchyAssociationStatus.ParentIdentityNotEstablished, node.Associations.Last().Status);
            Assert.AreEqual(Node(result, app).OccurrenceIndex, node.Associations.Last().ParentOccurrenceIndex);
            Assert.AreEqual(HierarchyPlacement.ParentNotEstablished, node.Placement);
            Assert.AreEqual(0, result.ParentCorrelations.Count); AssertLossless(result);
        }

        [TestMethod]
        public void ConflictingParentPortableNameCannotOverrideTheSelectedIdentity()
        {
            var app = Identity(1, 80, "app"); var child = Identity(2, 10075, "setting"); var p = Present(new[] { app, child });
            var evidence = new ComponentHierarchyParentEvidence(p.Source.Snapshot, DisplayEnvironmentSide.Source,
                new[] { new HierarchyStructuralObservation(child.Record.SolutionComponentId, child.Record.ObjectId,
                    app.Record.ObjectId, ComponentSemanticKinds.AppModule, "different-app", "Conflicting retained name", true) });
            var result = ComponentHierarchyProjector.Project(p, evidence);
            Assert.AreEqual(HierarchyAssociationStatus.ConflictingEvidence, Node(result, child).Associations.Last().Status);
            Assert.AreEqual(HierarchyPlacement.ParentNotEstablished, Node(result, child).Placement); AssertLossless(result);
        }

        [TestMethod]
        public void UnknownTableIdentityCannotBeReportedAsMissingFromTheSelectedScope()
        {
            var table = Identity(1, 1, null, status: IdentityResolutionStatus.Ambiguous);
            var child = Identity(2, 2, "account.c"); var p = Present(new[] { table, child });
            var result = ComponentHierarchyProjector.Project(Display(p, "account"));
            Assert.AreEqual(HierarchyAssociationStatus.ParentIdentityNotEstablished, Node(result, child).Associations.Last().Status);
            Assert.AreEqual(HierarchyPlacement.ParentNotEstablished, Node(result, child).Placement); AssertLossless(result);
        }

        [TestMethod]
        public void VerifiedColumnMetadataScopeUsesExistingTableIdentityWithoutKeyParsing()
        {
            var table = Identity(1, 1, "account"); var child = Identity(2, 2, "opaque-column-key"); var p = Present(new[] { table, child });
            var completed = Display(p, "account"); var result = ComponentHierarchyProjector.Project(completed);
            var node = Node(result, child);
            Assert.AreEqual(HierarchyPlacement.VerifiedStructuralParent, node.Placement);
            Assert.AreEqual("AttributeMetadata.EntityLogicalName", node.Associations.Last().Provenance);
            Assert.AreSame(child, node.Identity); AssertLossless(result);
        }

        [DataTestMethod]
        [DataRow("account", DisplayCacheAssociation.Rejected, "AttributeMetadata.EntityLogicalName")]
        [DataRow("account", DisplayCacheAssociation.NotAvailable, "AttributeMetadata.EntityLogicalName")]
        [DataRow("account", DisplayCacheAssociation.Verified, "guessed from key")]
        [DataRow("missing", DisplayCacheAssociation.Verified, "AttributeMetadata.EntityLogicalName")]
        public void UnverifiedColumnScopeOrUnavailableParentRemainsIncomplete(string scope, DisplayCacheAssociation association, string provenance)
        {
            var table = Identity(1, 1, "account"); var child = Identity(2, 2, "account.c"); var p = Present(new[] { table, child });
            var result = ComponentHierarchyProjector.Project(Display(p, scope, association, provenance));
            Assert.AreEqual(HierarchyPlacement.ParentNotEstablished, Node(result, child).Placement); AssertLossless(result);
        }

        [TestMethod]
        public void ConflictingStructuralEvidenceIsRetainedWithoutChoosingAParent()
        {
            var a = Identity(1, 1, "account"); var b = Identity(2, 1, "contact"); var child = Identity(3, 2, "opaque");
            var p = Present(new[] { a, b, child });
            var evidence = new ComponentHierarchyParentEvidence(p.Source.Snapshot, DisplayEnvironmentSide.Source, new[] {
                Observation(child, a), Observation(child, b) });
            var result = ComponentHierarchyProjector.Project(p, evidence); var node = Node(result, child);
            Assert.AreEqual(HierarchyPlacement.ParentNotEstablished, node.Placement);
            Assert.AreEqual(2, node.Associations.Count(x => x.Kind == HierarchyRelationshipKind.StructuralParent && x.Status == HierarchyAssociationStatus.Verified));
            Assert.IsTrue(node.Associations.Any(x => x.Status == HierarchyAssociationStatus.ConflictingEvidence));
            Assert.IsTrue(node.HasDifferentParentAssociations); Assert.AreEqual(IdentityResolutionStatus.Resolved, node.Identity.Status);
            AssertLossless(result);
        }

        [TestMethod]
        public void InclusionAndStructuralAssociationsDoNotImplyEquivalentOwnership()
        {
            var a = Identity(1, 1, "account"); var b = Identity(2, 1, "contact"); var child = Identity(3, 10, "relationship", a.Record.SolutionComponentId);
            var p = Present(new[] { a, b, child });
            var result = ComponentHierarchyProjector.Project(p, new ComponentHierarchyParentEvidence(p.Source.Snapshot,
                DisplayEnvironmentSide.Source, new[] { Observation(child, b) }));
            var node = Node(result, child);
            Assert.IsTrue(node.HasDifferentParentAssociations);
            Assert.AreEqual(2, node.Associations.Count);
            StringAssert.Contains(node.Associations[0].Diagnostic, "not schema ownership");
            StringAssert.Contains(node.Associations[1].Diagnostic, "not a raw inclusion-root"); AssertLossless(result);
        }

        [TestMethod]
        public void MixedAssociationCycleCannotBecomeAConfirmedDisplayParentPath()
        {
            var table = Identity(1, 1, "account", Id(2)); var column = Identity(2, 2, "account.c");
            var p = Present(new[] { table, column });
            var evidence = new ComponentHierarchyParentEvidence(p.Source.Snapshot, DisplayEnvironmentSide.Source,
                new[] { Observation(column, table) });
            var result = ComponentHierarchyProjector.Project(p, evidence);
            Assert.IsTrue(result.Occurrences.All(n => n.Associations.Any(a => a.Status == HierarchyAssociationStatus.CyclicReference)));
            Assert.IsTrue(result.Occurrences.All(n => n.Placement == HierarchyPlacement.ParentNotEstablished)); AssertLossless(result);
        }

        [TestMethod]
        public void PartialStructuralEvidenceCannotSelectTheSuccessfulAlternative()
        {
            var table = Identity(1, 1, "account"); var column = Identity(2, 2, "account.c"); var p = Present(new[] { table, column });
            var evidence = new ComponentHierarchyParentEvidence(p.Source.Snapshot, DisplayEnvironmentSide.Source,
                new[] { Observation(column, table), new HierarchyStructuralObservation(column.Record.SolutionComponentId,
                    column.Record.ObjectId, Id(999), ComponentSemanticKinds.Table, "unknown", "Incomplete retained scope", false) });
            var result = ComponentHierarchyProjector.Project(p, evidence);
            Assert.AreEqual(HierarchyPlacement.ParentNotEstablished, Node(result, column).Placement);
            Assert.AreEqual(3, Node(result, column).Associations.Count); AssertLossless(result);
        }

        [TestMethod]
        public void SameLocalParentGuidAcrossEnvironmentsDoesNotManufactureSemanticMatch()
        {
            var source = Identity(1, 1, "account"); var target = Identity(1, 1, "contact");
            var result = ComponentHierarchyProjector.Project(Present(new[] { source, Identity(2, 2, "column", Id(1)) },
                new[] { target, Identity(22, 2, "column", Id(1)) }));
            Assert.AreEqual(0, result.ParentCorrelations.Count); AssertLossless(result);
        }

        [TestMethod]
        public void IdenticalRawRecordReferencesRemainSeparateOccurrencesAndMappings()
        {
            var identity = Identity(1, 2, "account.c"); var p = Present(new[] { identity, identity });
            var result = ComponentHierarchyProjector.Project(p);
            Assert.AreEqual(2, result.Occurrences.Count);
            Assert.AreEqual(2, result.ComparisonRows.Count);
            Assert.IsTrue(result.ComparisonRows.All(r => r.SourceOccurrences.Count == 2));
            Assert.IsTrue(result.Occurrences.All(n => n.ComparisonRowIndices.Count == 2)); AssertLossless(result);
        }

        [TestMethod]
        public void OriginalDefinitionDetailsDiagnosticsAndEvidenceAreRetainedByReference()
        {
            var p = Present(new[] { Identity(1, 1, "account") }, new[] { Identity(21, 1, "account") });
            var detail = new Models.ComponentDetails.ComponentDefinitionDetailPresentation("Table", "account", "Present in Both", "Match",
                new[] { new Models.ComponentDetails.ComponentPropertyPresentation("SchemaName", "Account", "Account", false, "Match") },
                "Original definition diagnostic", new[] { "Source definition evidence" }, new[] { "Target definition evidence" });
            var row = p.Rows[0].WithDefinition("Match", "", detail);
            var supplied = new MembershipComparisonPresentation("sample", p.Source, p.Target, new[] { row });
            var result = ComponentHierarchyProjector.Project(supplied);
            Assert.AreSame(row, result.ComparisonRows[0].Row); Assert.AreSame(detail, result.ComparisonRows[0].Row.DefinitionDetail);
            Assert.AreEqual("Match", result.ComparisonRows[0].Row.DefinitionStatus);
            Assert.AreSame(p.Source.Snapshot.Components[0].DiagnosticEvidence, result.Occurrences[0].Identity.DiagnosticEvidence);
            AssertLossless(result);
        }

        [TestMethod]
        public void ExistingAuthoritativeParentMatchIsTheOnlyCrossEnvironmentCorrelation()
        {
            var sourceParent = Identity(1, 1, "account"); var targetParent = Identity(21, 1, "ACCOUNT");
            var child = Identity(2, 2, "account.c", sourceParent.Record.SolutionComponentId);
            var targetChild = Identity(22, 2, "account.c", targetParent.Record.SolutionComponentId);
            var p = Present(new[] { sourceParent, child }, new[] { targetParent, targetChild });
            var result = ComponentHierarchyProjector.Project(p);
            Assert.AreEqual(1, result.ParentCorrelations.Count);
            var pair = result.ParentCorrelations.Single();
            Assert.AreEqual(Node(result, sourceParent).OccurrenceIndex, pair.SourceOccurrenceIndex);
            Assert.AreEqual(Node(result, targetParent).OccurrenceIndex, pair.TargetOccurrenceIndex);
            Assert.AreEqual(MembershipPresence.PresentInBoth, p.Rows[pair.ComparisonRowIndex].Presence);
            Assert.AreNotEqual(sourceParent.Record.SolutionComponentId, targetParent.Record.SolutionComponentId); AssertLossless(result);
        }

        [TestMethod]
        public void DifferentSourceAndTargetParentAssociationsRemainIndependent()
        {
            var a = Identity(1, 1, "account"); var b = Identity(21, 1, "contact");
            var source = Identity(2, 26, "view", a.Record.SolutionComponentId); var target = Identity(22, 26, "view", b.Record.SolutionComponentId);
            var result = ComponentHierarchyProjector.Project(Present(new[] { a, source }, new[] { b, target }));
            Assert.AreEqual(Node(result, a).OccurrenceIndex, Node(result, source).Associations.Single().ParentOccurrenceIndex);
            Assert.AreEqual(Node(result, b).OccurrenceIndex, Node(result, target).Associations.Single().ParentOccurrenceIndex);
            Assert.AreEqual(0, result.ParentCorrelations.Count); AssertLossless(result);
        }

        [DataTestMethod]
        [DataRow(60), DataRow(91), DataRow(92)]
        public void ApprovedCollapsedReferencesKeepEveryRawOccurrenceAndOneExistingRow(int type)
        {
            var a = Identity(1, type, "same", objectId: Id(999)); var b = Identity(2, type, "same", objectId: Id(999));
            var p = Present(new[] { a, b }); var result = ComponentHierarchyProjector.Project(p);
            Assert.AreEqual(1, p.Rows.Count); Assert.AreEqual(2, result.Occurrences.Count);
            Assert.AreEqual(2, result.ComparisonRows.Single().SourceOccurrences.Count);
            Assert.IsTrue(result.Occurrences.All(n => n.ComparisonRowIndices.SequenceEqual(new[] { 0 })));
            Assert.AreEqual(0, result.ParentCorrelations.Count); AssertLossless(result);
        }

        [TestMethod]
        public void NonCollapsedDuplicatesAndMultipleRowsAreNotSuppressed()
        {
            var a = Identity(1, 26, "same", objectId: Id(999)); var b = Identity(2, 26, "same", objectId: Id(999));
            var p = Present(new[] { a, b });
            var rows = p.Rows.Concat(new[] { p.Rows[0] }).ToList();
            var supplied = new MembershipComparisonPresentation("sample", p.Source, p.Target, rows);
            var result = ComponentHierarchyProjector.Project(supplied);
            Assert.AreEqual(2, result.Occurrences.Count); Assert.AreEqual(3, result.ComparisonRows.Count);
            Assert.AreEqual(2, Node(result, a).ComparisonRowIndices.Count);
            Assert.IsTrue(result.ComparisonRows.All(r => r.Row.SourceResolutionStatus == "Ambiguous")); AssertLossless(result);
        }

        [TestMethod]
        public void StableOrderingPreservesOriginalRawIndicesWithoutReorderingComparisonRows()
        {
            var identities = new[] { Identity(3, 1, "c"), Identity(1, 1, "a"), Identity(2, 1, "b") };
            var first = ComponentHierarchyProjector.Project(Present(identities));
            var second = ComponentHierarchyProjector.Project(Present(identities.Reverse()));
            CollectionAssert.AreEqual(first.Occurrences.Select(n => n.Record.SolutionComponentId).ToArray(), second.Occurrences.Select(n => n.Record.SolutionComponentId).ToArray());
            foreach (var node in first.Occurrences) Assert.AreSame(node.Identity, node.Snapshot.Components[node.OriginalRawIndex]);
            AssertLossless(first); AssertLossless(second);
        }

        [TestMethod]
        public void EmptyAndUnavailableInventoriesRemainExplicitWithoutAbsenceInference()
        {
            var empty = ComponentHierarchyProjector.Project(Present(new ComponentIdentity[0]));
            Assert.AreEqual(0, empty.Occurrences.Count); Assert.AreEqual(0, empty.ComparisonRows.Count);
            var target = Snapshot(new[] { Identity(1, 1, "a") }, true);
            var p = new MembershipResultPresenter().Create(MembershipEnvironmentResult.Unavailable("Source", "sample", 0, TimeSpan.Zero, "unavailable"),
                MembershipEnvironmentResult.FromSnapshot("Target", target, 0, TimeSpan.Zero));
            var result = ComponentHierarchyProjector.Project(p);
            Assert.IsTrue(result.Diagnostics.Any(d => d.Contains("not absence")));
            Assert.AreEqual(MembershipPresence.Indeterminate, p.Rows[0].Presence); AssertLossless(result);
        }

        [DataTestMethod]
        [DataRow("before"), DataRow("AccountingPrepared"), DataRow("ParentsValidated"), DataRow("Completed")]
        public void CancellationDiscardsPartialHierarchyButKeepsCompleteAuthority(string when)
        {
            var table = Identity(1, 1, "account"); var child = Identity(2, 2, "account.c", table.Record.SolutionComponentId);
            var p = Present(new[] { table, child }); var cts = new CancellationTokenSource();
            if (when == "before") cts.Cancel();
            var result = ComponentHierarchyProjector.ProjectCore(p, null, null, null, cts.Token, stage => { if (stage == when) cts.Cancel(); });
            Assert.AreEqual(HierarchyProjectionOutcome.Cancelled, result.Outcome);
            Assert.IsTrue(result.Occurrences.All(n => n.Placement == HierarchyPlacement.ParentNotEstablished));
            Assert.AreEqual(0, result.ParentCorrelations.Count); AssertLossless(result);
        }

        [TestMethod]
        public void FailureIsLocalAndDoesNotExposeExceptionPayloadOrPartialHierarchy()
        {
            var p = Present(new[] { Identity(1, 1, "account"), Identity(2, 2, "account.c", Id(1)) });
            var csv = new MembershipCoverageCsvExporter().CreateCsv(p);
            var result = ComponentHierarchyProjector.ProjectCore(p, null, null, null, CancellationToken.None,
                stage => { if (stage == "ParentsValidated") throw new InvalidOperationException("secret-token-url"); });
            Assert.AreEqual(HierarchyProjectionOutcome.Failed, result.Outcome);
            Assert.IsTrue(result.Occurrences.All(n => n.Placement == HierarchyPlacement.ParentNotEstablished));
            Assert.IsFalse(string.Join(" ", result.Diagnostics).Contains("secret-token-url"));
            Assert.AreEqual(csv, new MembershipCoverageCsvExporter().CreateCsv(p)); AssertLossless(result);
        }

        [TestMethod]
        public void WrongSelectedSolutionFailsWithoutAttachingRelationships()
        {
            var p = Present(new[] { Identity(1, 1, "account"), Identity(2, 2, "account.c", Id(1)) });
            var wrong = new MembershipComparisonPresentation("other", p.Source, p.Target, p.Rows);
            var result = ComponentHierarchyProjector.Project(wrong);
            Assert.AreEqual(HierarchyProjectionOutcome.Failed, result.Outcome); AssertLossless(result);
        }

        [TestMethod]
        public void ProjectionCannotCallDataverseAndLeavesExistingRequestCountersAndExportsUntouched()
        {
            var p = Present(new[] { Identity(1, 1, "account"), Identity(2, 2, "account.c", Id(1)) });
            var before = new MembershipCoverageCsvExporter().CreateCsv(p);
            var result = ComponentHierarchyProjector.Project(p);
            foreach (var parameter in typeof(ComponentHierarchyProjector).GetMethods(BindingFlags.Public | BindingFlags.Static).SelectMany(m => m.GetParameters()))
                Assert.IsFalse(parameter.ParameterType.FullName.Contains("OrganizationService"));
            Assert.AreEqual(27, p.Source.Diagnostics.RequestCount); Assert.AreEqual(27, p.Target.Diagnostics.RequestCount);
            Assert.AreEqual(before, new MembershipCoverageCsvExporter().CreateCsv(p)); AssertLossless(result);
        }

        [DataTestMethod]
        [DataRow(840, false), DataRow(10000, false), DataRow(10000, true)]
        public void LargeInventoriesAndDeepChainsAreLocallyResponsive(int count, bool deep)
        {
            var items = Enumerable.Range(1, count).Select(i => Identity(i, deep ? 1 : 2, "key" + i,
                deep ? (i == count ? (Guid?)null : Id(i + 1)) : (i == 1 ? (Guid?)null : Id(1)))).ToList();
            var p = RawOnly(items); var watch = Stopwatch.StartNew();
            var result = ComponentHierarchyProjector.Project(p); watch.Stop();
            TestContext.WriteLine("H1 records={0}; deep={1}; projection_ms={2:0.000}; requests=0; writes=0", count, deep, watch.Elapsed.TotalMilliseconds);
            Assert.AreEqual(count - 1, result.Occurrences.Count(n => n.Placement == HierarchyPlacement.VerifiedInclusionRoot));
            Assert.IsTrue(watch.Elapsed < TimeSpan.FromSeconds(10)); AssertLossless(result);
        }

        [TestMethod]
        public void HistoricalEduAccountingShapeDoesNotSubstituteOldComparisonClassifications()
        {
            // Synthetic shape derived from saved evidence; actual local-archive replay is a separate validation.
            var source = Shape(792, 627); var target = Shape(720, 534);
            var p = RawOnly(source, target); var settings = target.Where(i => i.Record.ComponentType == 10075).ToList();
            var app = target.Single(i => i.Record.ComponentType == 80);
            var evidence = new ComponentHierarchyParentEvidence(p.Target.Snapshot, DisplayEnvironmentSide.Target,
                settings.Select(s => Observation(s, app)));
            var result = ComponentHierarchyProjector.Project(p, targetParents: evidence);
            Assert.AreEqual(627, result.Occurrences.Count(n => n.Side == DisplayEnvironmentSide.Source && n.Placement == HierarchyPlacement.VerifiedInclusionRoot));
            Assert.AreEqual(534, result.Occurrences.Count(n => n.Side == DisplayEnvironmentSide.Target && n.Placement == HierarchyPlacement.VerifiedInclusionRoot));
            Assert.AreEqual(15, result.Occurrences.Count(n => n.Side == DisplayEnvironmentSide.Target && n.Placement == HierarchyPlacement.VerifiedStructuralParent));
            Assert.AreEqual(165, result.Occurrences.Count(n => n.Side == DisplayEnvironmentSide.Source && n.Placement == HierarchyPlacement.ParentNotEstablished));
            Assert.AreEqual(171, result.Occurrences.Count(n => n.Side == DisplayEnvironmentSide.Target && n.Placement == HierarchyPlacement.ParentNotEstablished));
            Assert.AreEqual(0, result.ComparisonRows.Count); AssertLossless(result);
        }

        [TestMethod]
        public void Current840ComparisonBaselineIsIndependentlyPreserved()
        {
            var source = new List<ComponentIdentity>(); var target = new List<ComponentIdentity>(); int id = 1;
            for (int i = 0; i < 672; i++) { source.Add(Identity(id++, 61, "shared" + i)); target.Add(Identity(id++, 61, "shared" + i)); }
            for (int i = 0; i < 103; i++) source.Add(Identity(id++, 61, "source" + i));
            for (int i = 0; i < 32; i++) target.Add(Identity(id++, 61, "target" + i));
            for (int i = 0; i < 13; i++) { source.Add(Identity(id++, 300, null, status: IdentityResolutionStatus.Unsupported)); target.Add(Identity(id++, 300, null, status: IdentityResolutionStatus.Unsupported)); }
            for (int i = 0; i < 2; i++) { source.Add(Identity(id++, 92, null, status: IdentityResolutionStatus.Ambiguous)); target.Add(Identity(id++, 92, null, status: IdentityResolutionStatus.Ambiguous)); }
            Func<ComponentIdentity> indeterminate = () => new ComponentIdentity(new SolutionComponentRecord(Id(id++), 60, Id(id + 50000)),
                IdentityResolutionStatus.Resolved, "match-only" + id, componentTypeKey: ComponentSemanticKinds.SystemForm,
                inventoryAbsencePolicy: InventoryAbsencePolicy.MatchOnly);
            source.Add(indeterminate()); target.Add(indeterminate()); target.Add(indeterminate());
            var p = Present(source, target);
            var enriched = new MembershipComparisonPresentation("sample", p.Source, p.Target, p.Rows.Select(r => r.WithDefinition(
                r.SourceIdentity?.Status == IdentityResolutionStatus.Unsupported || r.TargetIdentity?.Status == IdentityResolutionStatus.Unsupported ? "Unsupported" : "Not Compared", "", null)));
            var csv = new MembershipCoverageCsvExporter().CreateCsv(enriched); var result = ComponentHierarchyProjector.Project(enriched);
            var s = result.Presentation.Summary;
            Assert.AreEqual(840, s.TotalComparisonRows); Assert.AreEqual(672, s.PresentInBoth); Assert.AreEqual(103, s.SourceOnly);
            Assert.AreEqual(32, s.TargetOnly); Assert.AreEqual(26, s.IndeterminateUnsupported); Assert.AreEqual(4, s.IndeterminateAmbiguous); Assert.AreEqual(3, s.IndeterminateResolved);
            Assert.AreEqual(csv, new MembershipCoverageCsvExporter().CreateCsv(enriched)); AssertLossless(result);
        }

        private static List<ComponentIdentity> Shape(int total, int roots)
        {
            var items = new List<ComponentIdentity> { Identity(1, 1, "account"), Identity(2, 80, "app") };
            for (int i = 0; i < roots; i++) items.Add(Identity(i + 3, 2, "account.c" + i, Id(1)));
            for (int i = 0; i < 15; i++) items.Add(Identity(roots + 3 + i, 10075, "setting" + i));
            while (items.Count < total) items.Add(Identity(items.Count + 1, 61, "web" + items.Count));
            return items;
        }

        private static HierarchyStructuralObservation Observation(ComponentIdentity child, ComponentIdentity parent) =>
            new HierarchyStructuralObservation(child.Record.SolutionComponentId, child.Record.ObjectId, parent.Record.ObjectId,
                parent.SemanticKind, parent.ComparisonKey, "Synthetic independently verified parent binding", true);
        private static ComponentHierarchyOccurrence Node(ComponentHierarchyProjection result, ComponentIdentity identity) =>
            result.Occurrences.Single(n => ReferenceEquals(n.Identity, identity));

        private static void AssertLossless(ComponentHierarchyProjection result)
        {
            var p = result.Presentation;
            Assert.AreEqual((p.Source.Snapshot?.Components.Count ?? 0) + (p.Target.Snapshot?.Components.Count ?? 0), result.Occurrences.Count);
            Assert.AreEqual(p.Rows.Count, result.ComparisonRows.Count);
            Assert.AreEqual(result.Occurrences.Count, result.Occurrences.Select(n => n.OccurrenceIndex).Distinct().Count());
            foreach (var n in result.Occurrences)
            {
                Assert.AreSame(n.Snapshot.Components[n.OriginalRawIndex], n.Identity); Assert.AreSame(n.Identity.Record, n.Record);
                foreach (var r in n.ComparisonRowIndices)
                    Assert.IsTrue((n.Side == DisplayEnvironmentSide.Source ? result.ComparisonRows[r].SourceOccurrences : result.ComparisonRows[r].TargetOccurrences).Contains(n.OccurrenceIndex));
            }
            foreach (var link in result.ComparisonRows)
            {
                Assert.AreSame(p.Rows[link.RowIndex], link.Row);
                foreach (var n in link.SourceOccurrences.Concat(link.TargetOccurrences)) Assert.IsTrue(result.Occurrences[n].ComparisonRowIndices.Contains(link.RowIndex));
            }
        }

        private static Guid Id(int n) => new Guid(n, 0, 0, new byte[8]);
        private static ComponentIdentity Identity(int id, int type, string key, Guid? root = null,
            IdentityResolutionStatus status = IdentityResolutionStatus.Resolved, Guid? objectId = null, int? behavior = 1) =>
            new ComponentIdentity(new SolutionComponentRecord(Id(id), type, objectId ?? Id(id + 50000), behavior, root, true), status,
                status == IdentityResolutionStatus.Resolved ? key : null, diagnostic: "Synthetic original diagnostic",
                componentTypeKey: Kind(type), diagnosticEvidence: new[] { "Original diagnostic evidence" });
        private static string Kind(int type) => type == 1 ? ComponentSemanticKinds.Table : type == 2 ? ComponentSemanticKinds.Column :
            type == 80 ? ComponentSemanticKinds.AppModule : type == 10075 ? ComponentSemanticKinds.AppSetting :
            type == 60 ? ComponentSemanticKinds.SystemForm : type == 91 ? ComponentSemanticKinds.PluginAssembly :
            type == 92 ? ComponentSemanticKinds.SdkMessageProcessingStep : "componenttype:" + type;
        private static MembershipSnapshot Snapshot(IEnumerable<ComponentIdentity> items, bool target) => MembershipSnapshot.Complete(
            new SolutionIdentity(new EnvironmentIdentity(Id(target ? 90002 : 90001), target ? "Target" : "Source"), Id(target ? 80002 : 80001), "sample"),
            items, new DateTimeOffset(2026, 9, 9, 0, 0, 0, TimeSpan.Zero));
        private static MembershipComparisonPresentation Present(IEnumerable<ComponentIdentity> source, IEnumerable<ComponentIdentity> target = null) =>
            new MembershipResultPresenter().Create(MembershipEnvironmentResult.FromSnapshot("Source", Snapshot(source, false), 27, TimeSpan.Zero),
                MembershipEnvironmentResult.FromSnapshot("Target", Snapshot(target ?? new ComponentIdentity[0], true), 27, TimeSpan.Zero));
        private static MembershipComparisonPresentation RawOnly(IEnumerable<ComponentIdentity> source, IEnumerable<ComponentIdentity> target = null) =>
            new MembershipComparisonPresentation("sample", MembershipEnvironmentResult.FromSnapshot("Source", Snapshot(source, false), 27, TimeSpan.Zero),
                MembershipEnvironmentResult.FromSnapshot("Target", Snapshot(target ?? new ComponentIdentity[0], true), 27, TimeSpan.Zero), new MembershipResultRow[0]);

        private static CompletedMembershipComparison Display(MembershipComparisonPresentation p, string scope,
            DisplayCacheAssociation association = DisplayCacheAssociation.Verified, string provenance = "AttributeMetadata.EntityLogicalName")
        {
            Func<DisplayEnvironmentSide, ComponentIdentity, ComponentSideDisplayEvidence> side = (role, identity) =>
                new ComponentSideDisplayEvidence(role, identity?.Record.SolutionComponentId, identity?.Record.ObjectId, identity?.Record.ComponentType,
                    new ComponentDisplayValue(DisplayAvailability.NotRetrieved), new ComponentDisplayValue(DisplayAvailability.NotRetrieved),
                    new ComponentDisplayValue(DisplayAvailability.NotRetrieved), DisplayBackingCorrelationStatus.VerifiedMetadata,
                    new ComponentDisplayValue(DisplayAvailability.Available, scope, provenance));
            return new CompletedMembershipComparison(p, p.Rows.Select((r, i) => new ComponentDisplayEvidence(i,
                side(DisplayEnvironmentSide.Source, r.SourceIdentity), side(DisplayEnvironmentSide.Target, r.TargetIdentity))),
                DisplayProjectionOutcome.Completed, association, association);
        }

        private static AppSettingEvidenceReport AppReport(MembershipSnapshot owner, ComponentIdentity child, ComponentIdentity parent, string mode = null)
        {
            var record = new AppSettingRecordEvidence(child.Record.ObjectId.Value, new[] {
                new KeyValuePair<string, string>("appsettingid", mode == "primary" ? Id(9999).ToString("D") : child.Record.ObjectId.Value.ToString("D")),
                new KeyValuePair<string, string>("parentappmoduleid", mode == "reference" ? Id(9998).ToString("D") : parent.Record.ObjectId.Value.ToString("D")) });
            var correlation = new AppSettingEntityCorrelation("appsetting", mode == "correlation" ? AppSettingEvidenceState.Ambiguous : AppSettingEvidenceState.Confirmed,
                "appsettingid", 1, new[] { record }, "Original completed correlation");
            var candidate = new AppSettingCandidateEvidence(owner.Environment.DisplayName, owner.SolutionUniqueName, owner.Solution.SolutionId,
                child.Record.SolutionComponentId, child.Record.ComponentType, "", child.Record.ObjectId, child.Record.RootComponentBehavior,
                mode == "root" ? Id(9999) : child.Record.RootSolutionComponentId,
                mode == "metadata" ? false : child.Record.IsMetadata, "Existing registered AppSetting", new[] { correlation },
                mode == "state" ? AppSettingEvidenceState.Unresolved : AppSettingEvidenceState.Confirmed, Id(900).ToString("D"), "definition",
                parent.Record.ObjectId.Value.ToString("D"), parent.ComparisonKey ?? "app", "Display name", "existing-composite", "Existing diagnostic");
            return new AppSettingEvidenceReport(mode == "environment" ? "Foreign" : owner.Environment.DisplayName,
                owner.SolutionUniqueName, mode == "scope" ? Id(9997) : owner.Solution.SolutionId, new AppSettingComponentTypeSummary[0],
                new AppSettingEntityMetadataEvidence[0], new[] { candidate }, new AppSettingRequestSummary(0, 0, 0, 0, 0, 0, 0, 0), new string[0]);
        }
    }
}
