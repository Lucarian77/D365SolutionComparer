using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using D365SolutionComparer.Infrastructure;
using D365SolutionComparer.Models.Membership;
using D365SolutionComparer.Services.Membership;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Microsoft.Xrm.Sdk;
using static D365SolutionComparer.Tests.MembershipTestData;

namespace D365SolutionComparer.Tests
{
    [TestClass]
    public class MembershipPresentationTests
    {
        [DataTestMethod]
        [DataRow(true, false)]
        [DataRow(false, true)]
        [DataRow(true, true)]
        public void CommandAllowsOneOrTwoSidedSelectedSolutions(bool sourcePresent, bool targetPresent)
        {
            Assert.IsTrue(MembershipCompareCommandEvaluator.CanExecute(true, true, true, true, true,
                sourcePresent, targetPresent, false));
        }

        [TestMethod]
        public void CommandRequiresSelectionConnectionsLoadedDataAndAtLeastOneSolution()
        {
            Assert.IsFalse(MembershipCompareCommandEvaluator.CanExecute(false, true, true, true, true, true, true, false));
            Assert.IsFalse(MembershipCompareCommandEvaluator.CanExecute(true, false, true, true, true, true, true, false));
            Assert.IsFalse(MembershipCompareCommandEvaluator.CanExecute(true, true, false, true, true, true, true, false));
            Assert.IsFalse(MembershipCompareCommandEvaluator.CanExecute(true, true, true, false, true, true, true, false));
            Assert.IsFalse(MembershipCompareCommandEvaluator.CanExecute(true, true, true, true, false, true, true, false));
            Assert.IsFalse(MembershipCompareCommandEvaluator.CanExecute(true, true, true, true, true, false, false, false));
            Assert.IsFalse(MembershipCompareCommandEvaluator.CanExecute(true, true, true, true, true, true, true, true));
        }

        [TestMethod]
        public void ResolvedMatchIsPresentedAsPresentInBoth()
        {
            var presentation = Present(Snapshot(Identity("new_/script.js", kind: "webresource")),
                Snapshot(Identity("NEW_/SCRIPT.JS", kind: "webresource")));
            var row = presentation.Rows.Single();
            Assert.AreEqual("Web Resource", row.ComponentKind);
            Assert.AreEqual("new_/script.js", row.PortableKey);
            Assert.AreEqual("Present", row.SourcePresence);
            Assert.AreEqual("Present", row.TargetPresence);
            Assert.AreEqual("Present in Both", row.MembershipStatus);
            Assert.AreEqual(1, presentation.Summary.PresentInBoth);
        }

        [TestMethod]
        public void AppModuleIdentityUsesModelDrivenAppPresentationLabel()
        {
            var presentation = Present(
                Snapshot(Identity("new_App", 80, kind: ComponentSemanticKinds.AppModule)),
                Snapshot(Identity("NEW_APP", 80, kind: ComponentSemanticKinds.AppModule)));

            Assert.AreEqual("Model-driven App / AppModule", presentation.Rows.Single().ComponentKind);
            Assert.AreEqual("Present in Both", presentation.Rows.Single().MembershipStatus);
        }

        [TestMethod]
        public void GlobalChoiceIdentityUsesGlobalChoicePresentationLabel()
        {
            var presentation = Present(
                Snapshot(Identity("new_Priority", 9, kind: ComponentSemanticKinds.GlobalChoice)),
                Snapshot(Identity("NEW_PRIORITY", 9, kind: ComponentSemanticKinds.GlobalChoice)));

            Assert.AreEqual("Global Choice", presentation.Rows.Single().ComponentKind);
            Assert.AreEqual("Present in Both", presentation.Rows.Single().MembershipStatus);
        }

        [TestMethod]
        public void SignedReportIdentityUsesSignedReportPresentationLabel()
        {
            var signatureId = Guid.NewGuid().ToString("D");
            var presentation = Present(
                Snapshot(Identity(signatureId, 31, kind: ComponentSemanticKinds.Report)),
                Snapshot(Identity(signatureId, 31, kind: ComponentSemanticKinds.Report)));

            Assert.AreEqual("Signed Report", presentation.Rows.Single().ComponentKind);
            Assert.AreEqual("Present in Both", presentation.Rows.Single().MembershipStatus);
        }

        [TestMethod]
        public void EntityKeyIdentityUsesEntityKeyPresentationLabel()
        {
            const string portableKey = "entitykey:v1:7:account:10:new_keyone";
            var presentation = Present(
                Snapshot(Identity(portableKey, 14, kind: ComponentSemanticKinds.EntityKey)),
                Snapshot(Identity(portableKey, 14, kind: ComponentSemanticKinds.EntityKey)));

            var row = presentation.Rows.Single();
            Assert.AreEqual("Entity Key", row.ComponentKind);
            Assert.AreEqual("Present in Both", row.MembershipStatus);
            Assert.AreEqual(14, row.SourceRawComponentType);
            Assert.AreEqual(14, row.TargetRawComponentType);
        }

        [TestMethod]
        public void UnsupportedTeamTemplateUsesTeamTemplatePresentationLabel()
        {
            var teamTemplate = Identity(null, 511, IdentityResolutionStatus.Unsupported,
                ComponentSemanticKinds.TeamTemplate);

            var row = Present(Snapshot(teamTemplate), Snapshot()).Rows.Single();

            Assert.AreEqual("Team Template", row.ComponentKind);
            Assert.AreEqual("Indeterminate - Unsupported", row.MembershipStatus);
            Assert.AreEqual("Unsupported", row.SourceResolutionStatus);
            Assert.AreEqual("Indeterminate", row.TargetPresence);
        }

        [TestMethod]
        public void CompleteResolvedOppositeInventoryPresentsGenuineMissing()
        {
            var presentation = Present(Snapshot(Identity("new_/script.js")), Snapshot());
            var row = presentation.Rows.Single();
            Assert.AreEqual("Source Only", row.MembershipStatus);
            Assert.AreEqual("Present", row.SourcePresence);
            Assert.AreEqual("Missing", row.TargetPresence);
            StringAssert.Contains(row.Diagnostic, "opposite Web Resource component kind");
            Assert.AreEqual(1, presentation.Summary.SourceOnly);
        }

        [TestMethod]
        public void OneSidedAbsentSolutionPresentsGenuineMissing()
        {
            var present = Snapshot(Identity("new_/script.js"));
            var absent = MembershipSnapshot.Absent(Solution().Environment, "sample", DateTimeOffset.UtcNow);
            var presentation = Present(present, absent);
            Assert.AreEqual("Source Only", presentation.Rows.Single().MembershipStatus);
            StringAssert.Contains(presentation.Rows.Single().Diagnostic, "opposite solution is absent");
        }

        [TestMethod]
        public void IndeterminateDiagnosticNamesTheIncompleteOppositeComponentKind()
        {
            var sourceColumn = Identity("account.name", 2, kind: ComponentSemanticKinds.Column);
            var targetColumn = Identity(null, 2, IdentityResolutionStatus.Unresolved,
                ComponentSemanticKinds.Column);

            var row = Present(Snapshot(sourceColumn), Snapshot(targetColumn)).Rows
                .Single(item => item.PortableKey == "account.name");

            StringAssert.Contains(row.Diagnostic,
                "opposite Column component kind is not fully resolved");
        }

        [TestMethod]
        public void TargetOnlyAbsentSolutionPresentsGenuineMissing()
        {
            var absent = MembershipSnapshot.Absent(Solution().Environment, "sample", DateTimeOffset.UtcNow);
            var present = Snapshot(Identity("new_/script.js"));
            var presentation = Present(absent, present);
            var row = presentation.Rows.Single();
            Assert.AreEqual("Target Only", row.MembershipStatus);
            Assert.AreEqual("Missing", row.SourcePresence);
            Assert.AreEqual("Present", row.TargetPresence);
            Assert.AreEqual(1, presentation.Summary.TargetOnly);
        }

        [TestMethod]
        public void EmptyPresentSolutionsProduceNoPresentationRows()
        {
            var presentation = Present(Snapshot(), Snapshot());
            Assert.AreEqual(0, presentation.Rows.Count);
            Assert.AreEqual(0, presentation.Summary.PresentInBoth);
            Assert.AreEqual(0, presentation.Summary.SourceOnly);
            Assert.AreEqual(0, presentation.Summary.TargetOnly);
        }

        [DataTestMethod]
        [DataRow(IdentityResolutionStatus.Unsupported, "Indeterminate - Unsupported")]
        [DataRow(IdentityResolutionStatus.Unresolved, "Indeterminate - Unresolved")]
        [DataRow(IdentityResolutionStatus.Ambiguous, "Indeterminate - Ambiguous")]
        public void UnknownIdentityIsNeverPresentedAsMissing(IdentityResolutionStatus status,
            string expectedStatus)
        {
            var unknown = new ComponentIdentity(
                new SolutionComponentRecord(Guid.NewGuid(), 98765, Guid.NewGuid()), status,
                diagnostic: "Identity unavailable");
            var absent = MembershipSnapshot.Absent(Solution().Environment, "sample", DateTimeOffset.UtcNow);
            var presentation = Present(Snapshot(unknown), absent);
            var row = presentation.Rows.Single();
            Assert.AreEqual(expectedStatus, row.MembershipStatus);
            Assert.AreEqual("Indeterminate", row.TargetPresence);
            Assert.AreEqual(0, presentation.Summary.SourceOnly);
            Assert.AreEqual(status == IdentityResolutionStatus.Unsupported ? 1 : 0, presentation.Summary.Unsupported);
            Assert.AreEqual(status == IdentityResolutionStatus.Unresolved ? 1 : 0, presentation.Summary.Unresolved);
            Assert.AreEqual(status == IdentityResolutionStatus.Ambiguous ? 1 : 0, presentation.Summary.Ambiguous);
        }

        [TestMethod]
        public void DuplicatePortableKeysArePresentedAsAmbiguousNotMatchedOrMissing()
        {
            var presentation = Present(Snapshot(Identity("same"), Identity("SAME")), Snapshot());
            Assert.AreEqual(2, presentation.Rows.Count);
            Assert.IsTrue(presentation.Rows.All(row => row.MembershipStatus == "Indeterminate - Ambiguous"));
            Assert.AreEqual(2, presentation.Summary.Ambiguous);
            Assert.AreEqual(0, presentation.Summary.SourceOnly);
        }

        [TestMethod]
        public void UnavailableSideKeepsAvailableRowsIndeterminateAndUsesNullableDiagnostics()
        {
            var unavailable = MembershipEnvironmentResult.Unavailable("Source", "sample", 3,
                TimeSpan.FromSeconds(2), "FaultException: Access denied");
            var target = MembershipEnvironmentResult.FromSnapshot("Target", Snapshot(Identity("key")),
                4, TimeSpan.FromSeconds(1));
            var presentation = new MembershipResultPresenter().Create(unavailable, target);
            var row = presentation.Rows.Single();
            Assert.AreEqual("Unavailable", row.SourcePresence);
            Assert.AreEqual("Present", row.TargetPresence);
            Assert.AreEqual("Indeterminate - Environment Unavailable", row.MembershipStatus);
            StringAssert.Contains(row.Diagnostic, "Source retrieval unavailable");
            Assert.AreEqual(0, presentation.Summary.TargetOnly);
            Assert.IsNull(unavailable.Diagnostics.RawMembershipCount);
            Assert.AreEqual(3, unavailable.Diagnostics.RequestCount);
        }

        [TestMethod]
        public void DuplicateKeysRemainAmbiguousWhenOppositeEnvironmentIsUnavailable()
        {
            var source = MembershipEnvironmentResult.FromSnapshot("Source",
                Snapshot(Identity("same"), Identity("SAME")), 4, TimeSpan.Zero);
            var target = MembershipEnvironmentResult.Unavailable("Target", "sample", 1,
                TimeSpan.Zero, "Connection failed");
            var presentation = new MembershipResultPresenter().Create(source, target);
            Assert.AreEqual(2, presentation.Summary.Ambiguous);
            Assert.IsTrue(presentation.Rows.All(row => row.TargetPresence == "Unavailable"));
            Assert.IsTrue(presentation.Rows.All(row => row.SourceResolutionStatus == "Ambiguous"));
            Assert.AreEqual(0, presentation.Summary.PresentInBoth);
            Assert.AreEqual(0, presentation.Summary.SourceOnly);
        }

        [TestMethod]
        public void SideDiagnosticsExposeRawAndResolutionCounts()
        {
            var snapshot = Snapshot(Identity("resolved"),
                Identity(null, status: IdentityResolutionStatus.Unsupported),
                Identity(null, status: IdentityResolutionStatus.Unresolved),
                Identity(null, status: IdentityResolutionStatus.Ambiguous));
            var result = MembershipEnvironmentResult.FromSnapshot("Source", snapshot, 9, TimeSpan.FromSeconds(3));
            Assert.AreEqual(4, result.Diagnostics.RawMembershipCount);
            Assert.AreEqual(1, result.Diagnostics.ResolvedCount);
            Assert.AreEqual(1, result.Diagnostics.UnsupportedCount);
            Assert.AreEqual(1, result.Diagnostics.UnresolvedCount);
            Assert.AreEqual(1, result.Diagnostics.AmbiguousCount);
        }

        [TestMethod]
        public void UnavailableSnapshotDoesNotReportAnEmptySuccessfulInventory()
        {
            var snapshot = MembershipSnapshot.Unavailable(Solution().Environment, "sample", DateTimeOffset.UtcNow,
                "Read denied");
            var result = MembershipEnvironmentResult.FromSnapshot("Source", snapshot, 2, TimeSpan.Zero);
            Assert.AreEqual(MembershipSnapshotState.Unavailable, result.State);
            Assert.IsNull(result.Diagnostics.RawMembershipCount);
            Assert.IsNull(result.Diagnostics.ResolvedCount);
        }

        [TestMethod]
        public void LiveOperationOverloadCapturesEnvironmentAndUsesOneWhoAmI()
        {
            var solution = Solution();
            var service = Service(solution, query => query.EntityName == "solution"
                ? Rows(SolutionRow(solution)) : Rows());
            var counter = new DataverseRequestCounter();
            var stages = new List<MembershipOperationStage>();
            var snapshot = new DataverseSolutionMembershipOperation().ReadAndResolve(service, "Live environment",
                solution.UniqueName, CancellationToken.None, progress => stages.Add(progress.Stage), counter);
            Assert.AreEqual(solution.Environment.OrganizationId, snapshot.Environment.OrganizationId);
            Assert.AreEqual("Live environment", snapshot.Environment.DisplayName);
            Assert.AreEqual(1, counter.GetExecuteCount("WhoAmI"));
            Assert.AreEqual(3, counter.TotalRequests);
            CollectionAssert.Contains(stages, MembershipOperationStage.ValidatingEnvironment);
            CollectionAssert.Contains(stages, MembershipOperationStage.ReadingMembership);
            CollectionAssert.Contains(stages, MembershipOperationStage.ResolvingIdentities);
            CollectionAssert.Contains(stages, MembershipOperationStage.Completed);
        }

        [TestMethod]
        public void LiveOperationOverloadCancellationCannotReturnAPartialSnapshot()
        {
            var solution = Solution();
            var service = Service(solution, query => Rows(SolutionRow(solution)));
            var counter = new DataverseRequestCounter();
            var cancellation = new CancellationTokenSource();
            cancellation.Cancel();
            Assert.ThrowsException<OperationCanceledException>(() =>
                new DataverseSolutionMembershipOperation().ReadAndResolve(service, "Live environment",
                    solution.UniqueName, cancellation.Token, progress => { }, counter));
            Assert.AreEqual(0, counter.TotalRequests);
            Assert.AreEqual(0, service.Calls);
            Assert.AreEqual(0, service.ExecuteCalls);
        }

        [DataTestMethod, TestCategory("MembershipSummaryAccounting"), TestCategory("Phase2GPresentation")]
        [DataRow(MembershipPresence.PresentInBoth)]
        [DataRow(MembershipPresence.OnlyInSource)]
        [DataRow(MembershipPresence.OnlyInTarget)]
        public void ExhaustiveSummaryCountsVerifiedPresenceOnce(MembershipPresence presence)
        {
            var presentation = Present(
                Snapshot(presence == MembershipPresence.OnlyInTarget ? new ComponentIdentity[0] : new[] { Identity("key") }),
                Snapshot(presence == MembershipPresence.OnlyInSource ? new ComponentIdentity[0] : new[] { Identity("key") }));
            Assert.AreEqual(presence, presentation.Rows.Single().Comparison.Presence);
            Assert.AreEqual(presence == MembershipPresence.PresentInBoth ? 1 : 0, presentation.Summary.PresentInBoth);
            Assert.AreEqual(presence == MembershipPresence.OnlyInSource ? 1 : 0, presentation.Summary.SourceOnly);
            Assert.AreEqual(presence == MembershipPresence.OnlyInTarget ? 1 : 0, presentation.Summary.TargetOnly);
            AssertPartition(presentation);
        }

        [DataTestMethod, TestCategory("MembershipSummaryAccounting"), TestCategory("Phase2GPresentation")]
        [DataRow(IdentityResolutionStatus.Unsupported)]
        [DataRow(IdentityResolutionStatus.Unresolved)]
        [DataRow(IdentityResolutionStatus.Ambiguous)]
        public void ExhaustiveSummaryCountsNonResolvedIndeterminateRows(IdentityResolutionStatus status)
        {
            var presentation = Present(Snapshot(Identity(null, status: status)), Snapshot());
            Assert.AreEqual(status == IdentityResolutionStatus.Unsupported ? 1 : 0, presentation.Summary.IndeterminateUnsupported);
            Assert.AreEqual(status == IdentityResolutionStatus.Unresolved ? 1 : 0, presentation.Summary.IndeterminateUnresolved);
            Assert.AreEqual(status == IdentityResolutionStatus.Ambiguous ? 1 : 0, presentation.Summary.IndeterminateAmbiguous);
            Assert.AreEqual(0, presentation.Summary.IndeterminateResolved);
            AssertPartition(presentation);
        }

        [DataTestMethod, TestCategory("MembershipSummaryAccounting"), TestCategory("Phase2GPresentation")]
        [DataRow(IdentityResolutionStatus.Resolved, IdentityResolutionStatus.Resolved, IdentityResolutionStatus.Resolved)]
        [DataRow(IdentityResolutionStatus.Resolved, IdentityResolutionStatus.Unsupported, IdentityResolutionStatus.Unsupported)]
        [DataRow(IdentityResolutionStatus.Resolved, IdentityResolutionStatus.Unresolved, IdentityResolutionStatus.Unresolved)]
        [DataRow(IdentityResolutionStatus.Resolved, IdentityResolutionStatus.Ambiguous, IdentityResolutionStatus.Ambiguous)]
        [DataRow(IdentityResolutionStatus.Unsupported, IdentityResolutionStatus.Resolved, IdentityResolutionStatus.Unsupported)]
        [DataRow(IdentityResolutionStatus.Unsupported, IdentityResolutionStatus.Unsupported, IdentityResolutionStatus.Unsupported)]
        [DataRow(IdentityResolutionStatus.Unsupported, IdentityResolutionStatus.Unresolved, IdentityResolutionStatus.Unresolved)]
        [DataRow(IdentityResolutionStatus.Unsupported, IdentityResolutionStatus.Ambiguous, IdentityResolutionStatus.Ambiguous)]
        [DataRow(IdentityResolutionStatus.Unresolved, IdentityResolutionStatus.Resolved, IdentityResolutionStatus.Unresolved)]
        [DataRow(IdentityResolutionStatus.Unresolved, IdentityResolutionStatus.Unsupported, IdentityResolutionStatus.Unresolved)]
        [DataRow(IdentityResolutionStatus.Unresolved, IdentityResolutionStatus.Unresolved, IdentityResolutionStatus.Unresolved)]
        [DataRow(IdentityResolutionStatus.Unresolved, IdentityResolutionStatus.Ambiguous, IdentityResolutionStatus.Ambiguous)]
        [DataRow(IdentityResolutionStatus.Ambiguous, IdentityResolutionStatus.Resolved, IdentityResolutionStatus.Ambiguous)]
        [DataRow(IdentityResolutionStatus.Ambiguous, IdentityResolutionStatus.Unsupported, IdentityResolutionStatus.Ambiguous)]
        [DataRow(IdentityResolutionStatus.Ambiguous, IdentityResolutionStatus.Unresolved, IdentityResolutionStatus.Ambiguous)]
        [DataRow(IdentityResolutionStatus.Ambiguous, IdentityResolutionStatus.Ambiguous, IdentityResolutionStatus.Ambiguous)]
        public void MixedResolutionSummaryUsesOnePartitionWithoutChangingEitherSide(
            IdentityResolutionStatus sourceStatus, IdentityResolutionStatus targetStatus, IdentityResolutionStatus expected)
        {
            // Synthetic valid comparison contracts; this is not a live EDU pairing.
            var source = Identity("key", status: sourceStatus);
            var target = Identity("key", status: targetStatus);
            var comparison = new MembershipCompareResult(source, target, MembershipPresence.Indeterminate);
            var row = SummaryRow(comparison);
            var presentation = new MembershipComparisonPresentation("sample",
                MembershipEnvironmentResult.FromSnapshot("Source", Snapshot(source), 7, TimeSpan.Zero),
                MembershipEnvironmentResult.FromSnapshot("Target", Snapshot(target), 9, TimeSpan.Zero), new[] { row });
            var summary = presentation.Summary;
            Assert.AreEqual(expected == IdentityResolutionStatus.Ambiguous ? 1 : 0, summary.IndeterminateAmbiguous);
            Assert.AreEqual(expected == IdentityResolutionStatus.Unresolved ? 1 : 0, summary.IndeterminateUnresolved);
            Assert.AreEqual(expected == IdentityResolutionStatus.Unsupported ? 1 : 0, summary.IndeterminateUnsupported);
            Assert.AreEqual(expected == IdentityResolutionStatus.Resolved ? 1 : 0, summary.IndeterminateResolved);
            // Legacy counters retain their original overlapping meaning for existing consumers.
            Assert.AreEqual(sourceStatus == IdentityResolutionStatus.Unsupported || targetStatus == IdentityResolutionStatus.Unsupported ? 1 : 0, summary.Unsupported);
            Assert.AreEqual(sourceStatus == IdentityResolutionStatus.Unresolved || targetStatus == IdentityResolutionStatus.Unresolved ? 1 : 0, summary.Unresolved);
            Assert.AreEqual(sourceStatus == IdentityResolutionStatus.Ambiguous || targetStatus == IdentityResolutionStatus.Ambiguous ? 1 : 0, summary.Ambiguous);
            Assert.AreSame(row, presentation.Rows.Single());
            Assert.AreSame(comparison, row.Comparison);
            Assert.AreSame(source, comparison.Source);
            Assert.AreSame(target, comparison.Target);
            Assert.AreEqual(sourceStatus, source.Status);
            Assert.AreEqual(targetStatus, target.Status);
            Assert.AreEqual(MembershipPresence.Indeterminate, comparison.Presence);
            Assert.AreEqual(MembershipAbsenceEvidence.None, comparison.AbsenceEvidence);
            Assert.AreEqual("synthetic unchanged reason", row.Diagnostic);
            Assert.AreEqual("Not Compared", row.DefinitionStatus);
            Assert.AreEqual("unchanged property", row.ChangedProperties);
            AssertPartition(presentation);
        }

        [TestMethod, TestCategory("MembershipSummaryAccounting"), TestCategory("Phase2GPresentation")]
        public void SyntheticOneSourceAndTwoTargetResolvedIndeterminateRowsAreAccountedSeparatelyFromOccurrences()
        {
            var presentation = Present(Snapshot(Identity("shared"), Identity("source-only"), MatchOnly("source-form")),
                Snapshot(Identity("shared"), Identity("target-only"), MatchOnly("target-form-one"), MatchOnly("target-form-two")));
            Assert.AreEqual(1, presentation.Summary.PresentInBoth);
            Assert.AreEqual(1, presentation.Summary.SourceOnly);
            Assert.AreEqual(1, presentation.Summary.TargetOnly);
            Assert.AreEqual(3, presentation.Summary.IndeterminateResolved);
            Assert.AreEqual(6, presentation.Summary.TotalComparisonRows);
            Assert.AreEqual(1, presentation.Rows.Count(row => row.Presence == MembershipPresence.Indeterminate && row.SourceIdentity != null));
            Assert.AreEqual(2, presentation.Rows.Count(row => row.Presence == MembershipPresence.Indeterminate && row.TargetIdentity != null));
            Assert.AreEqual(3, presentation.Source.Diagnostics.RawMembershipCount);
            Assert.AreEqual(4, presentation.Target.Diagnostics.RawMembershipCount);
            AssertOccurrencePartition(presentation.Source.Diagnostics);
            AssertOccurrencePartition(presentation.Target.Diagnostics);
            AssertPartition(presentation);
        }

        [DataTestMethod, TestCategory("MembershipSummaryAccounting"), TestCategory("Phase2GPresentation")]
        [DataRow(60, ComponentSemanticKinds.SystemForm)]
        [DataRow(91, ComponentSemanticKinds.PluginAssembly)]
        [DataRow(92, ComponentSemanticKinds.SdkMessageProcessingStep)]
        public void ApprovedRepeatedReferencesKeepRawOccurrencesAndOneComparisonRow(int type, string kind)
        {
            var objectId = Guid.NewGuid();
            Func<ComponentIdentity> reference = () => new ComponentIdentity(
                new SolutionComponentRecord(Guid.NewGuid(), type, objectId), IdentityResolutionStatus.Resolved,
                "same-portable-key", componentTypeKey: kind);
            var presentation = Present(Snapshot(reference(), reference()), Snapshot(reference(), reference()));
            Assert.AreEqual(2, presentation.Source.Diagnostics.RawMembershipCount);
            Assert.AreEqual(2, presentation.Target.Diagnostics.RawMembershipCount);
            Assert.AreEqual(1, presentation.Summary.PresentInBoth);
            Assert.AreEqual(1, presentation.Summary.TotalComparisonRows);
            AssertOccurrencePartition(presentation.Source.Diagnostics);
            AssertOccurrencePartition(presentation.Target.Diagnostics);
            AssertPartition(presentation);
        }

        [TestMethod, TestCategory("MembershipSummaryAccounting"), TestCategory("Phase2GPresentation")]
        public void DuplicatePortableIdentitiesRemainAmbiguousAndDoNotCreateMatches()
        {
            var presentation = Present(Snapshot(Identity("duplicate"), Identity("DUPLICATE")), Snapshot());
            Assert.AreEqual(2, presentation.Summary.IndeterminateAmbiguous);
            Assert.AreEqual(0, presentation.Summary.PresentInBoth);
            Assert.AreEqual(0, presentation.Summary.SourceOnly);
            Assert.IsTrue(presentation.Rows.All(row => row.SourceResolutionStatus == "Ambiguous"));
            // Environment diagnostics count the retained pre-comparison raw identities, unchanged.
            Assert.AreEqual(2, presentation.Source.Diagnostics.ResolvedCount);
            AssertPartition(presentation);
        }

        [TestMethod, TestCategory("MembershipSummaryAccounting"), TestCategory("Phase2GPresentation")]
        public void UnsupportedObservationsOnBothSidesStayTwoIndeterminateRows()
        {
            var presentation = Present(Snapshot(Identity(null, 300, IdentityResolutionStatus.Unsupported)),
                Snapshot(Identity(null, 300, IdentityResolutionStatus.Unsupported)));
            Assert.AreEqual(2, presentation.Summary.IndeterminateUnsupported);
            Assert.AreEqual(0, presentation.Summary.PresentInBoth);
            Assert.IsTrue(presentation.Rows.All(row => row.Presence == MembershipPresence.Indeterminate));
            AssertOccurrencePartition(presentation.Source.Diagnostics);
            AssertOccurrencePartition(presentation.Target.Diagnostics);
            AssertPartition(presentation);
        }

        [TestMethod, TestCategory("MembershipSummaryAccounting"), TestCategory("Phase2GPresentation")]
        public void IncompleteOppositeCoverageRetainsResolvedIndeterminateAndOriginalAbsenceState()
        {
            var source = Identity("account.name", 2, kind: ComponentSemanticKinds.Column);
            var presentation = Present(Snapshot(source), Snapshot(Identity(null, 2,
                IdentityResolutionStatus.Unresolved, ComponentSemanticKinds.Column)));
            Assert.AreEqual(1, presentation.Summary.IndeterminateResolved);
            Assert.AreEqual(1, presentation.Summary.IndeterminateUnresolved);
            Assert.AreEqual(0, presentation.Summary.SourceOnly);
            var row = presentation.Rows.Single(item => item.SourceIdentity == source);
            Assert.AreEqual(MembershipPresence.Indeterminate, row.Comparison.Presence);
            Assert.AreEqual(MembershipAbsenceEvidence.None, row.Comparison.AbsenceEvidence);
            StringAssert.Contains(row.Diagnostic, "not fully resolved");
            AssertPartition(presentation);
        }

        [DataTestMethod, TestCategory("MembershipSummaryAccounting"), TestCategory("Phase2GPresentation")]
        [DataRow(true)]
        [DataRow(false)]
        public void UnavailableOppositeEnvironmentCannotTurnResolvedAccountingIntoMissing(bool sourceAvailable)
        {
            var available = MembershipEnvironmentResult.FromSnapshot("Available", Snapshot(Identity("key")), 4, TimeSpan.Zero);
            var unavailable = MembershipEnvironmentResult.Unavailable("Unavailable", "sample", 2, TimeSpan.Zero, "Read unavailable");
            var presentation = new MembershipResultPresenter().Create(sourceAvailable ? available : unavailable,
                sourceAvailable ? unavailable : available);
            Assert.AreEqual(1, presentation.Summary.IndeterminateResolved);
            Assert.AreEqual(0, presentation.Summary.SourceOnly);
            Assert.AreEqual(0, presentation.Summary.TargetOnly);
            Assert.IsNull(unavailable.Diagnostics.RawMembershipCount);
            Assert.AreEqual(2, unavailable.Diagnostics.RequestCount);
            StringAssert.Contains(presentation.Rows.Single().MembershipStatus, "Unavailable");
            AssertPartition(presentation);
        }

        [DataTestMethod, TestCategory("MembershipSummaryAccounting"), TestCategory("Phase2GPresentation")]
        [DataRow(false)]
        [DataRow(true)]
        public void EmptyAuthoritativeResultsHaveZeroRowsNotFabricatedUnavailableComponents(bool unavailable)
        {
            var result = unavailable ? MembershipEnvironmentResult.Unavailable("Unavailable", "sample", 0, TimeSpan.Zero, "Unavailable")
                : MembershipEnvironmentResult.FromSnapshot("Complete", Snapshot(), 0, TimeSpan.Zero);
            var presentation = new MembershipResultPresenter().Create(result, result);
            Assert.AreEqual(0, presentation.Summary.TotalComparisonRows);
            Assert.AreEqual(0, presentation.Summary.IndeterminateOtherOrUnavailable);
            Assert.AreEqual(unavailable ? (int?)null : 0, result.Diagnostics.RawMembershipCount);
            AssertPartition(presentation);
        }

        [DataTestMethod, TestCategory("MembershipSummaryAccounting"), TestCategory("Phase2GPresentation")]
        [DataRow(-1)]
        [DataRow(99)]
        public void FutureResolutionRepresentationCannotSilentlyBecomeResolved(int value)
        {
            var unknown = (IdentityResolutionStatus)value;
            Assert.AreEqual(MembershipPresentationSummary.IndeterminatePartition.OtherOrUnavailable,
                MembershipPresentationSummary.ClassifyIndeterminate(unknown, IdentityResolutionStatus.Resolved));
            Assert.AreEqual(MembershipPresentationSummary.IndeterminatePartition.OtherOrUnavailable,
                MembershipPresentationSummary.ClassifyIndeterminate(null, unknown));
            // Today's authoritative contract rejects unknown resolution values; it is not weakened.
            Assert.ThrowsException<ArgumentOutOfRangeException>(() => Identity(null, status: unknown));
        }

        [TestMethod, TestCategory("MembershipSummaryAccounting"), TestCategory("Phase2GPresentation")]
        public void LegacyPresentationFallbackWithoutIdentityIsCountedInResidualBucket()
        {
            // Compatibility edge of the existing row constructor, not a fabricated live component.
            var row = new MembershipResultRow(MembershipPresence.Indeterminate, null, null,
                "Unavailable", "", "Unavailable", "Unavailable", "Indeterminate", "Unavailable", "Unavailable", "synthetic unavailable");
            var summary = new MembershipPresentationSummary(new[] { row });
            Assert.AreEqual(1, summary.TotalComparisonRows);
            Assert.AreEqual(1, summary.IndeterminateOtherOrUnavailable);
            Assert.AreEqual(0, summary.IndeterminateResolved);
            Assert.AreEqual(MembershipPresentationSummary.IndeterminatePartition.OtherOrUnavailable,
                MembershipPresentationSummary.ClassifyIndeterminate(null, null));
            Assert.AreEqual(summary.TotalComparisonRows, PartitionTotal(summary));
            StringAssert.Contains(MembershipResultsForm.FormatSummary(summary), "Other / Unavailable: 1");
        }

        [TestMethod, TestCategory("MembershipSummaryAccounting"), TestCategory("Phase2GPresentation")]
        public void LegacyCoverageExportKeepsOriginalSchemaAndMixedResolutionCounters()
        {
            var comparison = new MembershipCompareResult(Identity(null, status: IdentityResolutionStatus.Unsupported),
                Identity(null, status: IdentityResolutionStatus.Unresolved), MembershipPresence.Indeterminate);
            var presentation = new MembershipComparisonPresentation("sample",
                MembershipEnvironmentResult.FromSnapshot("Source", Snapshot(comparison.Source), 0, TimeSpan.Zero),
                MembershipEnvironmentResult.FromSnapshot("Target", Snapshot(comparison.Target), 0, TimeSpan.Zero), new[] { SummaryRow(comparison) });
            var csv = new MembershipCoverageCsvExporter().CreateCsv(presentation);
            var lines = csv.Split(new[] { "\r\n", "\n" }, StringSplitOptions.None);
            var headers = lines[0].Split(',').Select(value => value.Trim('"')).ToArray();
            var values = lines[1].Split(',').Select(value => value.Trim('"')).ToArray();
            Assert.AreEqual(55, headers.Length);
            Assert.AreEqual("1", values[Array.IndexOf(headers, "ComparisonUnsupported")]);
            Assert.AreEqual("1", values[Array.IndexOf(headers, "ComparisonUnresolved")]);
            CollectionAssert.DoesNotContain(headers, "TotalComparisonRows");
            CollectionAssert.DoesNotContain(headers, "IndeterminateResolved");
            Assert.AreEqual(1, presentation.Summary.IndeterminateUnresolved);
            Assert.AreEqual(0, presentation.Summary.IndeterminateUnsupported);
            StringAssert.Contains(MembershipResultsForm.FormatSummary(presentation.Summary), "Unsupported: 0");
            StringAssert.Contains(MembershipResultsForm.FormatSummary(presentation.Summary), "Unresolved: 1");
            AssertPartition(presentation);
        }

        [TestMethod, TestCategory("MembershipSummaryAccounting"), TestCategory("Phase2GPresentation")]
        public void PresentationAccountingPreservesComparerResultsAndRequestCounters()
        {
            var source = Snapshot(Identity("shared"), Identity("source-only"), MatchOnly("match-only"));
            var target = Snapshot(Identity("shared"), Identity("target-only"));
            var comparer = new SolutionMembershipComparer();
            var before = comparer.Compare(source, target);
            var presentation = Present(source, target);
            for (int i = 0; i < before.Count; i++)
            {
                Assert.AreSame(before[i].Source, presentation.Rows[i].Comparison.Source);
                Assert.AreSame(before[i].Target, presentation.Rows[i].Comparison.Target);
                Assert.AreEqual(before[i].Presence, presentation.Rows[i].Comparison.Presence);
                Assert.AreEqual(before[i].AbsenceEvidence, presentation.Rows[i].Comparison.AbsenceEvidence);
            }
            Assert.AreEqual(1, presentation.Source.Diagnostics.RequestCount);
            Assert.AreEqual(1, presentation.Target.Diagnostics.RequestCount);
            Assert.AreEqual(InventoryAbsencePolicy.MatchOnly, source.Components.Last().InventoryAbsencePolicy);
            Assert.AreEqual(1, presentation.Summary.IndeterminateResolved);
            AssertPartition(presentation);
        }

        private static ComponentIdentity MatchOnly(string key) => new ComponentIdentity(
            new SolutionComponentRecord(Guid.NewGuid(), 60, Guid.NewGuid()), IdentityResolutionStatus.Resolved,
            key, componentTypeKey: ComponentSemanticKinds.SystemForm, inventoryAbsencePolicy: InventoryAbsencePolicy.MatchOnly);

        private static MembershipResultRow SummaryRow(MembershipCompareResult comparison) => new MembershipResultRow(
            comparison.Presence, comparison.Source, comparison.Target, "Synthetic", "synthetic", "Present", "Present",
            "Indeterminate", comparison.Source.Status.ToString(), comparison.Target.Status.ToString(),
            "synthetic unchanged reason", comparison, "Not Compared", "unchanged property");

        private static int PartitionTotal(MembershipPresentationSummary summary) => summary.PresentInBoth + summary.SourceOnly +
            summary.TargetOnly + summary.IndeterminateAmbiguous + summary.IndeterminateUnresolved +
            summary.IndeterminateUnsupported + summary.IndeterminateResolved + summary.IndeterminateOtherOrUnavailable;

        private static void AssertPartition(MembershipComparisonPresentation presentation)
        {
            Assert.AreEqual(presentation.Rows.Count, presentation.Summary.TotalComparisonRows);
            Assert.AreEqual(presentation.Summary.TotalComparisonRows, PartitionTotal(presentation.Summary));
        }

        private static void AssertOccurrencePartition(MembershipEnvironmentDiagnostics diagnostics) =>
            Assert.AreEqual(diagnostics.RawMembershipCount, diagnostics.ResolvedCount + diagnostics.UnsupportedCount +
                diagnostics.UnresolvedCount + diagnostics.AmbiguousCount);

        private static MembershipComparisonPresentation Present(MembershipSnapshot source,
            MembershipSnapshot target)
        {
            return new MembershipResultPresenter().Create(
                MembershipEnvironmentResult.FromSnapshot("Source", source, 1, TimeSpan.Zero),
                MembershipEnvironmentResult.FromSnapshot("Target", target, 1, TimeSpan.Zero));
        }
    }
}
