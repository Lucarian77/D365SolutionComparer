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

        private static MembershipComparisonPresentation Present(MembershipSnapshot source,
            MembershipSnapshot target)
        {
            return new MembershipResultPresenter().Create(
                MembershipEnvironmentResult.FromSnapshot("Source", source, 1, TimeSpan.Zero),
                MembershipEnvironmentResult.FromSnapshot("Target", target, 1, TimeSpan.Zero));
        }
    }
}
