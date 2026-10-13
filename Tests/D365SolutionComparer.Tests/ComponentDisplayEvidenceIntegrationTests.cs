using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.ExceptionServices;
using System.Threading;
using System.Windows.Forms;
using Timer = System.Threading.Timer;
using Label = System.Windows.Forms.Label;
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
using Microsoft.Xrm.Sdk.Metadata;
using static D365SolutionComparer.Tests.MembershipTestData;

namespace D365SolutionComparer.Tests
{
    /// <summary>Synthetic integration fixtures only; these do not establish live SAA evidence.</summary>
    [TestClass, TestCategory("P1CDisplayIntegration")]
    public class ComponentDisplayEvidenceIntegrationTests
    {
        public TestContext TestContext { get; set; }

        [DataTestMethod]
        [DataRow(false)]
        [DataRow(true)]
        public void CacheAwareReadPreservesPipelineAndExactRequestLedger(bool empty)
        {
            var solution = Solution();
            var component = ComponentRow(solution);
            var ledgerOld = new List<string>(); var ledgerNew = new List<string>();
            var oldService = ReadService(solution, component, empty, ledgerOld);
            var newService = ReadService(solution, component, empty, ledgerNew);
            var oldCounter = new DataverseRequestCounter(); var newCounter = new DataverseRequestCounter();
            var oldProgress = new List<string>(); var newProgress = new List<string>();
            var operation = new DataverseComponentDefinitionOperation();
            var old = operation.ReadAndResolve(oldService, "Synthetic Source", solution.UniqueName,
                CancellationToken.None, p => oldProgress.Add(p.Stage + ":" + p.Message), oldCounter);
            var read = operation.ReadAndResolveWithCache(newService, "Synthetic Source", solution.UniqueName,
                CancellationToken.None, p => newProgress.Add(p.Stage + ":" + p.Message), newCounter);
            CollectionAssert.AreEqual(ledgerOld, ledgerNew);
            CollectionAssert.AreEqual(oldProgress, newProgress);
            Assert.AreEqual(oldCounter.TotalRequests, newCounter.TotalRequests);
            Assert.AreEqual(1, newCounter.GetExecuteCount("WhoAmI")); // Existing validation, no extra call.
            Assert.AreEqual(0, oldService.WriteCalls + newService.WriteCalls);
            Assert.AreEqual(old.Membership.State, read.Definitions.Membership.State);
            CollectionAssert.AreEqual(old.Membership.Components.Select(IdentitySignature).ToArray(),
                read.Definitions.Membership.Components.Select(IdentitySignature).ToArray());
            CollectionAssert.AreEqual(old.Definitions.Select(DefinitionSignature).ToArray(),
                read.Definitions.Definitions.Select(DefinitionSignature).ToArray());
            Assert.IsTrue(read.BelongsTo(operation, read.Definitions.Membership));
            Assert.IsNotNull(read.Cache);
            if (!empty)
            {
                Entity retained;
                Assert.IsTrue(read.Cache.TryGetEntityRow("webresource", (Guid)component["objectid"], out retained));
                Assert.AreEqual("synthetic.js", retained.GetAttributeValue<string>("name"));
            }
            TestContext.WriteLine("Unchanged synthetic ledger: " + string.Join(", ", ledgerNew));
        }

        [TestMethod]
        public void ExistingSolutionIdentityOverloadAndValidationRemainCompatible()
        {
            var solution = Solution(); var component = ComponentRow(solution);
            var operation = new DataverseComponentDefinitionOperation();
            var result = operation.ReadAndResolve(ReadService(solution, component, false, new List<string>()),
                solution, CancellationToken.None);
            Assert.AreEqual(IdentityResolutionStatus.Resolved, result.Membership.Components.Single().Status);
            Assert.AreEqual(ComponentDefinitionReadStatus.Available, result.Definitions.Single().Status);
            Assert.ThrowsException<ArgumentException>(() => operation.ReadAndResolveWithCache(null, "Test", " ", CancellationToken.None, null));
            Assert.ThrowsException<ArgumentNullException>(() => operation.ReadAndResolve(null, (SolutionIdentity)null, CancellationToken.None));
        }

        [DataTestMethod]
        [DataRow("Source")]
        [DataRow("Target")]
        public void RetrievalFailureStillEscapesReadWithoutReturningAnEnvelope(string side)
        {
            var service = Service(Solution(), query => throw new InvalidOperationException("Synthetic retrieval failure"));
            Assert.ThrowsException<InvalidOperationException>(() => new DataverseComponentDefinitionOperation()
                .ReadAndResolveWithCache(service, side, "sample", CancellationToken.None, null));
            Assert.AreEqual(0, service.WriteCalls);
        }

        [DataTestMethod]
        [DataRow("Swapped")]
        [DataRow("StaleInvocation")]
        [DataRow("WrongSolution")]
        [DataRow("WrongEnvironment")]
        [DataRow("CopiedSnapshot")]
        public void MisassociatedCachesAreRejectedWithoutBorrowingOrRetrieval(string problem)
        {
            var f = Fixture(); var source = f.Source; var target = f.Target;
            var token = Token(f.Operation);
            if (problem == "Swapped") { source = f.Target; target = f.Source; }
            if (problem == "StaleInvocation") source = Envelope(f.Presentation.Source.Snapshot, f.Source.Cache,
                new DataverseComponentDefinitionOperation());
            if (problem == "WrongSolution") source = new DataverseComponentDefinitionOperation.DefinitionReadWithCacheResult(
                f.Source.Definitions, f.Source.Cache, token, f.Presentation.Source.Snapshot.Environment.OrganizationId, "other");
            if (problem == "WrongEnvironment") source = new DataverseComponentDefinitionOperation.DefinitionReadWithCacheResult(
                f.Source.Definitions, f.Source.Cache, token, Guid.NewGuid(), "sample");
            if (problem == "CopiedSnapshot") source = Envelope(MembershipSnapshot.Complete(
                f.Presentation.Source.Snapshot.Solution, f.Presentation.Source.Snapshot.Components, DateTimeOffset.UtcNow),
                f.Source.Cache, f.Operation);
            var result = Complete(f, source, target);
            Assert.AreEqual(DisplayCacheAssociation.Rejected, result.SourceCacheAssociation);
            Assert.AreEqual(problem == "Swapped" ? DisplayCacheAssociation.Rejected : DisplayCacheAssociation.Verified,
                result.TargetCacheAssociation);
            Assert.AreEqual(DisplayAvailability.NotRetrieved, result.DisplayEvidence[0].SourceDisplayEvidence.FriendlyName.Availability);
            Assert.AreSame(f.Presentation, result.Presentation);
            Assert.AreEqual(1, result.Presentation.Summary.PresentInBoth);
        }

        [DataTestMethod]
        [DataRow(true, true)]
        [DataRow(true, false)]
        [DataRow(false, true)]
        [DataRow(false, false)]
        public void OptionalCachesRemainSideSpecific(bool sourceAvailable, bool targetAvailable)
        {
            var f = Fixture();
            var result = Complete(f, sourceAvailable ? f.Source : null, targetAvailable ? f.Target : null);
            Assert.AreEqual(DisplayProjectionOutcome.Completed, result.DisplayProjectionOutcome);
            Assert.AreEqual(sourceAvailable ? DisplayCacheAssociation.Verified : DisplayCacheAssociation.NotAvailable, result.SourceCacheAssociation);
            Assert.AreEqual(targetAvailable ? DisplayCacheAssociation.Verified : DisplayCacheAssociation.NotAvailable, result.TargetCacheAssociation);
            var item = result.DisplayEvidence.Single();
            Assert.AreEqual(sourceAvailable ? "Source display" : null, item.SourceDisplayEvidence.FriendlyName.Value);
            Assert.AreEqual(targetAvailable ? "Target display" : null, item.TargetDisplayEvidence.FriendlyName.Value);
            Assert.AreEqual(sourceAvailable ? "source_internal" : null, item.SourceDisplayEvidence.InternalName.Value);
            Assert.AreEqual(targetAvailable ? "target_internal" : null, item.TargetDisplayEvidence.InternalName.Value);
            Assert.AreEqual(DisplayEnvironmentSide.Source, item.SourceDisplayEvidence.EnvironmentSide);
            Assert.AreEqual(DisplayEnvironmentSide.Target, item.TargetDisplayEvidence.EnvironmentSide);
            Assert.AreNotEqual(item.SourceDisplayEvidence.LocalObjectId, item.TargetDisplayEvidence.LocalObjectId);
        }

        [DataTestMethod]
        [DataRow("Null")]
        [DataRow("Empty")]
        [DataRow("Omitted")]
        [DataRow("Malformed")]
        public void FieldAvailabilityDoesNotChangeAuthority(string fieldState)
        {
            var f = Fixture(); Entity row;
            f.Source.Cache.TryGetEntityRow("appmodule", f.Presentation.Rows[0].SourceIdentity.Record.ObjectId.Value, out row);
            if (fieldState == "Null") row["name"] = null;
            if (fieldState == "Empty") row["name"] = string.Empty;
            if (fieldState == "Omitted") row.Attributes.Remove("name");
            if (fieldState == "Malformed") row["name"] = new OptionSetValue(1);
            var result = Complete(f); var value = result.DisplayEvidence[0].SourceDisplayEvidence.FriendlyName;
            Assert.AreEqual(fieldState == "Omitted" ? DisplayAvailability.AttributeOmitted : fieldState == "Malformed"
                ? DisplayAvailability.Unavailable : DisplayAvailability.Available, value.Availability);
            Assert.AreEqual(fieldState == "Empty" ? string.Empty : null, value.Value);
            Assert.AreEqual(1, result.Presentation.Summary.PresentInBoth);
        }

        [TestMethod]
        public void RowOrderDuplicatesAndDifferentNamesFollowFinalDefinitionRows()
        {
            var source = new[] { Identity(80, "z"), Identity(80, "a") };
            var target = new[] { Identity(80, "a"), Identity(80, "z") };
            var f = Fixture(source, target);
            foreach (var item in source) f.Source.Cache.StoreEntityRow(App(item, "same friendly", "Source " + item.ComparisonKey), item.Record.ObjectId.Value);
            var result = Complete(f);
            CollectionAssert.AreEqual(new[] { 0, 1 }, result.DisplayEvidence.Select(e => e.RowIndex).ToArray());
            for (var i = 0; i < result.DisplayEvidence.Count; i++)
            {
                var row = f.Presentation.Rows[i]; var display = result.DisplayEvidence[i];
                Assert.AreEqual(row.SourceIdentity.Record.SolutionComponentId, display.SourceDisplayEvidence.SolutionComponentId);
                Assert.AreEqual(row.TargetIdentity.Record.SolutionComponentId, display.TargetDisplayEvidence.SolutionComponentId);
                Assert.AreEqual("same friendly", display.SourceDisplayEvidence.FriendlyName.Value);
                Assert.AreSame(row, result.Presentation.Rows[i]);
                Assert.AreSame(row.DefinitionDetail, result.Presentation.Rows[i].DefinitionDetail);
            }
            Assert.AreEqual(2, result.Presentation.Summary.PresentInBoth);
        }

        [DataTestMethod]
        [DataRow(300)]
        [DataRow(10072)]
        [DataRow(74)]
        [DataRow(36)]
        [DataRow(511)]
        [DataRow(10266)]
        [DataRow(10267)]
        public void UnsupportedFamiliesCannotBePromotedByDisplay(int type)
        {
            var f = Fixture(new[] { Identity(type, status: IdentityResolutionStatus.Unsupported) },
                new[] { Identity(type, status: IdentityResolutionStatus.Unsupported) });
            var result = Complete(f);
            Assert.AreEqual(2, result.DisplayEvidence.Count);
            Assert.IsTrue(result.Presentation.Rows.All(r => r.Presence == MembershipPresence.Indeterminate));
            Assert.IsTrue(f.Presentation.Source.Snapshot.Components.All(i => i.ComparisonKey == null));
            Assert.AreEqual(2, result.Presentation.Summary.IndeterminateUnsupported);
        }

        [TestMethod]
        public void AppActionLocalRawNamespacesAreNotNormalized()
        {
            var f = Fixture(new[] { Identity(10266, status: IdentityResolutionStatus.Unsupported) },
                new[] { Identity(10267, status: IdentityResolutionStatus.Unsupported) });
            var result = Complete(f);
            Assert.AreEqual(2, result.DisplayEvidence.Count);
            Assert.AreEqual(0, result.Presentation.Summary.PresentInBoth);
        }

        [TestMethod]
        public void AppSettingSupportAndOpaqueNameGapArePreserved()
        {
            var f = Fixture(new[] { Identity(10075, "appsetting:v1:3:app:7:setting", kind: ComponentSemanticKinds.AppSetting) },
                new[] { Identity(10075, "appsetting:v1:3:app:7:setting", kind: ComponentSemanticKinds.AppSetting) });
            var result = Complete(f);
            Assert.AreEqual(1, result.Presentation.Summary.PresentInBoth);
            Assert.AreEqual(IdentityResolutionStatus.Resolved, f.Presentation.Rows[0].SourceIdentity.Status);
            Assert.AreEqual(DisplayAvailability.NotRetrieved, result.DisplayEvidence[0].SourceDisplayEvidence.InternalName.Availability);
        }

        [DataTestMethod]
        [DataRow("SourceOnly")]
        [DataRow("TargetOnly")]
        [DataRow("MatchOnly")]
        [DataRow("Ambiguous")]
        [DataRow("DuplicateKey")]
        [DataRow("UnavailableSource")]
        [DataRow("UnavailableTarget")]
        public void MembershipDefinitionAccountingAndCsvStayUnchanged(string scenario)
        {
            var source = new[] { Identity(80) }; var target = new[] { Identity(80) };
            if (scenario == "SourceOnly" || scenario == "MatchOnly") target = new ComponentIdentity[0];
            if (scenario == "TargetOnly") source = new ComponentIdentity[0];
            if (scenario == "MatchOnly") source = new[] { Identity(80, policy: InventoryAbsencePolicy.MatchOnly) };
            if (scenario == "Ambiguous") source = new[] { Identity(80, status: IdentityResolutionStatus.Ambiguous) };
            if (scenario == "DuplicateKey") source = new[] { Identity(80), Identity(80) };
            var f = Fixture(source, target, scenario == "UnavailableSource", scenario == "UnavailableTarget");
            var csvBefore = new MembershipCoverageCsvExporter().CreateCsv(f.Presentation);
            var summaryBefore = Summary(f.Presentation);
            var result = Complete(f);
            Assert.AreSame(f.Presentation, result.Presentation);
            Assert.AreEqual(csvBefore, new MembershipCoverageCsvExporter().CreateCsv(result.Presentation));
            Assert.AreEqual(summaryBefore, Summary(result.Presentation));
            for (var i = 0; i < result.Presentation.Rows.Count; i++)
            {
                Assert.AreSame(f.Presentation.Rows[i].Comparison, result.Presentation.Rows[i].Comparison);
                Assert.AreSame(f.Presentation.Rows[i].DefinitionDetail, result.Presentation.Rows[i].DefinitionDetail);
            }
            if (scenario == "SourceOnly") Assert.AreEqual(1, result.Presentation.Summary.SourceOnly);
            if (scenario == "TargetOnly") Assert.AreEqual(1, result.Presentation.Summary.TargetOnly);
            if (scenario == "MatchOnly") Assert.AreEqual(1, result.Presentation.Summary.IndeterminateResolved);
            if (scenario == "DuplicateKey") Assert.AreEqual(0, result.Presentation.Summary.PresentInBoth);
            if (scenario == "UnavailableSource") Assert.AreEqual(DisplayAvailability.Unavailable, result.DisplayEvidence[0].SourceDisplayEvidence.FriendlyName.Availability);
            if (scenario == "UnavailableTarget") Assert.AreEqual(DisplayAvailability.Unavailable, result.DisplayEvidence[0].TargetDisplayEvidence.FriendlyName.Availability);
        }

        [TestMethod]
        public void RepeatedSystemFormReferencesRemainComparerOwned()
        {
            var shared = Guid.NewGuid();
            var a = Identity(60, "form-key", shared, kind: ComponentSemanticKinds.SystemForm);
            var b = Identity(60, "form-key", shared, kind: ComponentSemanticKinds.SystemForm);
            var f = Fixture(new[] { a, b }, new[] { Identity(60, "form-key", kind: ComponentSemanticKinds.SystemForm) });
            var result = Complete(f);
            Assert.AreEqual(2, f.Presentation.Source.Snapshot.Components.Count);
            Assert.AreEqual(1, result.DisplayEvidence.Count);
            Assert.AreEqual(1, result.Presentation.Summary.PresentInBoth);
        }

        [DataTestMethod]
        [DataRow(false)]
        [DataRow(true)]
        public void AvailableDefinitionEqualityAndDifferencesSurviveProjection(bool changed)
        {
            var f = Fixture();
            var source = f.Presentation.Source.Snapshot.Components.Single();
            var target = f.Presentation.Target.Snapshot.Components.Single();
            var fields = ComponentDefinitionContractCatalog.For(source.SemanticKind).ComparableProperties;
            var sd = new ComponentDefinition(source, ComponentDefinitionReadStatus.Available,
                fields.Select(name => new KeyValuePair<string, string>(name, "source")));
            var td = new ComponentDefinition(target, ComponentDefinitionReadStatus.Available,
                fields.Select(name => new KeyValuePair<string, string>(name, changed ? "target" : "source")));
            f.Presentation = new ComponentDefinitionResultPresenter().Apply(f.Presentation,
                new ComponentDefinitionSnapshot(f.Presentation.Source.Snapshot, new[] { sd }),
                new ComponentDefinitionSnapshot(f.Presentation.Target.Snapshot, new[] { td }));
            var result = Complete(f);
            Assert.AreEqual(changed ? "Different" : "Match", result.Presentation.Rows[0].DefinitionStatus);
            Assert.AreSame(f.Presentation.Rows[0].DefinitionDetail, result.Presentation.Rows[0].DefinitionDetail);
        }

        [TestMethod]
        public void ExistingCoverageAndDefinitionFormsStillConsumeAuthoritativePresentation()
        {
            OnSta(() =>
            {
                var f = Fixture(); var result = Complete(f); var builder = new MembershipCoverageDiagnosticsBuilder();
                using (var coverage = new MembershipCoverageDetailsForm("Source", builder.Build(result.Presentation.Source.Snapshot),
                    "Target", builder.Build(result.Presentation.Target.Snapshot), result.Presentation))
                using (var details = new ComponentDefinitionDetailsForm(result.Presentation.Rows[0].DefinitionDetail))
                {
                    coverage.Show(); details.Show(); Application.DoEvents();
                    Assert.AreEqual("Membership Coverage Details", coverage.Text);
                    Assert.IsTrue(details.Controls.OfType<SplitContainer>().Any());
                    Assert.AreSame(f.Presentation, typeof(MembershipCoverageDetailsForm).GetField("presentation",
                        BindingFlags.Instance | BindingFlags.NonPublic).GetValue(coverage));
                    details.Close(); coverage.Close();
                }
            });
        }

        [TestMethod]
        public void DisplayListIsCopiedAndCannotBeMutated()
        {
            var f = Fixture(); var original = Complete(f).DisplayEvidence.ToList();
            var result = new CompletedMembershipComparison(f.Presentation, original, DisplayProjectionOutcome.Completed);
            original.Clear();
            Assert.AreEqual(1, result.DisplayEvidence.Count);
            Assert.ThrowsException<NotSupportedException>(() => ((IList<ComponentDisplayEvidence>)result.DisplayEvidence).Clear());
        }

        [DataTestMethod]
        [DataRow("MissingRow")]
        [DataRow("WrongIndex")]
        [DataRow("NullRow")]
        [DataRow("PartialFailure")]
        public void InvalidCompletedHandoffsAreRejected(string problem)
        {
            var f = Fixture(); var rows = Complete(f).DisplayEvidence.ToList(); var outcome = DisplayProjectionOutcome.Completed;
            if (problem == "MissingRow") rows.Clear();
            if (problem == "WrongIndex") rows[0] = new ComponentDisplayEvidence(1, rows[0].SourceDisplayEvidence, rows[0].TargetDisplayEvidence);
            if (problem == "NullRow") rows[0] = null;
            if (problem == "PartialFailure") outcome = DisplayProjectionOutcome.Failed;
            Assert.ThrowsException<ArgumentException>(() => new CompletedMembershipComparison(f.Presentation, rows, outcome));
        }

        [TestMethod]
        public void DetachedHandoffHasNoSdkOrCacheFieldsAndSurvivesRowMutation()
        {
            var f = Fixture(); var result = Complete(f); Entity original;
            f.Source.Cache.TryGetEntityRow("appmodule", f.Presentation.Rows[0].SourceIdentity.Record.ObjectId.Value, out original);
            original["name"] = "mutated"; original["uniquename"] = "mutated_internal";
            original["large_unused_xml"] = new string('x', 10000);
            Assert.AreEqual("Source display", result.DisplayEvidence[0].SourceDisplayEvidence.FriendlyName.Value);
            foreach (var type in new[] { typeof(CompletedMembershipComparison), typeof(ComponentDisplayEvidence), typeof(ComponentSideDisplayEvidence), typeof(ComponentDisplayValue) })
                foreach (var field in type.GetFields(BindingFlags.Instance | BindingFlags.NonPublic | BindingFlags.Public))
                {
                    Assert.IsTrue(field.IsInitOnly);
                    Assert.IsFalse(typeof(IOrganizationService).IsAssignableFrom(field.FieldType));
                    Assert.IsFalse(typeof(Entity).IsAssignableFrom(field.FieldType));
                    Assert.IsFalse(typeof(EntityMetadata).IsAssignableFrom(field.FieldType));
                    Assert.AreNotEqual(typeof(DataverseComponentMetadataCache), field.FieldType);
                    Assert.AreNotEqual(typeof(DataverseReadContext), field.FieldType);
                }
        }

        [DataTestMethod]
        [DataRow("Exception")]
        [DataRow("UnexpectedEmpty")]
        [DataRow("UnexpectedNull")]
        [DataRow("UnrequestedCancellation")]
        public void OptionalFailureKeepsSuccessfulPresentationAndNoPayload(string problem)
        {
            var f = Fixture();
            var result = SolutionComparerControl.CompleteMembershipComparison(f.Presentation, f.Operation, f.Source, f.Target,
                CancellationToken.None, (p, s, t, token) =>
                {
                    if (problem == "UnexpectedEmpty") return new ComponentDisplayEvidence[0];
                    if (problem == "UnexpectedNull") return null;
                    if (problem == "UnrequestedCancellation") throw new OperationCanceledException();
                    throw new InvalidOperationException("Synthetic sensitive payload must not escape");
                });
            Assert.AreSame(f.Presentation, result.Presentation);
            Assert.AreEqual(DisplayProjectionOutcome.Failed, result.DisplayProjectionOutcome);
            Assert.AreEqual(0, result.DisplayEvidence.Count);
            Assert.AreEqual(1, result.Presentation.Summary.PresentInBoth);
            Assert.IsFalse(typeof(CompletedMembershipComparison).GetProperties().Any(p => p.PropertyType == typeof(Exception) || p.PropertyType == typeof(string)));
        }

        [DataTestMethod]
        [DataRow("Before")]
        [DataRow("During")]
        [DataRow("AfterReturn")]
        public void RequestedCancellationNeverPublishesAWrapper(string point)
        {
            var f = Fixture(); using (var cancel = new CancellationTokenSource())
            {
                var entered = false;
                if (point == "Before") cancel.Cancel();
                Assert.ThrowsException<OperationCanceledException>(() => SolutionComparerControl.CompleteMembershipComparison(
                    f.Presentation, f.Operation, f.Source, f.Target, cancel.Token, (p, s, t, token) =>
                    {
                        entered = true;
                        if (point == "During") { cancel.Cancel(); token.ThrowIfCancellationRequested(); }
                        var evidence = ComponentDisplayEvidenceProjector.Project(p, s, t, token);
                        cancel.Cancel(); // Actual post-projector token check must discard these completed values.
                        return evidence;
                    }));
                Assert.AreEqual(point != "Before", entered);
            }
        }

        [TestMethod]
        public void EmptyCompletedComparisonIsNotFailedProjection()
        {
            var f = Fixture(new ComponentIdentity[0], new ComponentIdentity[0]); var result = Complete(f);
            Assert.AreEqual(DisplayProjectionOutcome.Completed, result.DisplayProjectionOutcome);
            Assert.AreEqual(0, result.DisplayEvidence.Count);
        }

        [TestMethod]
        public void RepeatedInvocationsNeverReusePreviousDisplayEvidence()
        {
            var first = Fixture(); var old = Complete(first); var second = Fixture();
            var next = Complete(second, first.Source, first.Target);
            Assert.AreEqual(DisplayCacheAssociation.Rejected, next.SourceCacheAssociation);
            Assert.AreEqual(DisplayAvailability.NotRetrieved, next.DisplayEvidence[0].SourceDisplayEvidence.FriendlyName.Availability);
            Assert.AreEqual("Source display", old.DisplayEvidence[0].SourceDisplayEvidence.FriendlyName.Value);
            Assert.AreNotSame(old.Presentation, next.Presentation);
        }

        [DataTestMethod]
        [DataRow(false)]
        [DataRow(true)]
        public void BothFormConstructorsPreserveGridSummaryCoverageAndDetailBinding(bool useWrapper)
        {
            OnSta(() =>
            {
                var f = Fixture(); var completed = Complete(f);
                using (var form = useWrapper ? new MembershipResultsForm(completed) : new MembershipResultsForm(f.Presentation))
                {
                    Assert.AreSame(useWrapper ? completed : null, form.CompletedComparison);
                    form.Show(); Application.DoEvents();
                    var grid = form.Controls.OfType<DataGridView>().Single();
                    Assert.AreEqual(9, grid.Columns.Count);
                    Assert.AreEqual(f.Presentation.Rows.Count, grid.Rows.Count);
                    var adapter = (MembershipGridRow)grid.Rows[0].DataBoundItem;
                    Assert.AreSame(f.Presentation.Rows[0], adapter.AuthoritativeRow);
                    Assert.AreEqual(0, adapter.OriginalRowIndex);
                    Assert.AreSame(f.Presentation.Rows[0].DefinitionDetail, adapter.AuthoritativeRow.DefinitionDetail);
                    Assert.AreEqual(MembershipResultsForm.FormatSummary(f.Presentation.Summary),
                        form.Controls.OfType<Panel>().SelectMany(p => p.Controls.OfType<Label>()).Single(l => l.Name == "MembershipSummary").Text);
                    Assert.AreEqual(1, f.Presentation.Source.Diagnostics.RawMembershipCount);
                    form.Close();
                }
            });
        }

        [TestMethod]
        public void ProjectionAndFormRetentionAddNoServiceCallsOrReleaseCollectors()
        {
            var solution = Solution(); var service = ReadService(solution, ComponentRow(solution), false, new List<string>());
            var operation = new DataverseComponentDefinitionOperation(); var counter = new DataverseRequestCounter();
            var read = operation.ReadAndResolveWithCache(service, "Synthetic", "sample", CancellationToken.None, null, counter);
            var presentation = new MembershipResultPresenter().Create(Environment(read.Definitions.Membership), Environment(read.Definitions.Membership));
            var final = new ComponentDefinitionResultPresenter().Apply(presentation, read.Definitions, read.Definitions);
            var calls = service.Calls; var execute = service.ExecuteCalls; var count = counter.TotalRequests;
            var result = SolutionComparerControl.CompleteMembershipComparison(final, operation, read, read, CancellationToken.None);
            Assert.AreEqual(calls, service.Calls); Assert.AreEqual(execute, service.ExecuteCalls);
            Assert.AreEqual(count, counter.TotalRequests); Assert.AreEqual(0, service.WriteCalls);
            Assert.AreEqual(DisplayProjectionOutcome.Completed, result.DisplayProjectionOutcome);
#if !DEBUG
            foreach (var name in new[] { "Type300EvidenceCollector", "Type10072EvidenceCollector", "AppActionEvidenceCollector" })
                Assert.IsNull(typeof(SolutionComparerControl).Assembly.GetType("D365SolutionComparer.Services.Membership." + name));
#endif
        }

        [DataTestMethod]
        [DataRow("ExportVisibleRowsToCsv")]
        [DataRow("ExportVisibleRowsToSpreadsheetMl")]
        public void ExistingTextExportsAreUnaffectedByMembershipHandoff(string exportMethod)
        {
            OnSta(() =>
            {
                var firstPath = Path.GetTempFileName(); var secondPath = Path.GetTempFileName();
                try
                {
                    using (var host = new Form())
                    using (var control = new SolutionComparerControl())
                    {
                        host.Controls.Add(control); host.Show(); Application.DoEvents();
                        var grid = (DataGridView)typeof(SolutionComparerControl).GetField("dgvResults",
                            BindingFlags.Instance | BindingFlags.NonPublic).GetValue(control);
                        grid.DataSource = new List<Models.SolutionCompareResult> { new Models.SolutionCompareResult
                            { UniqueName = "synthetic_solution", SourceDisplayName = "Source solution", TargetDisplayName = "Target solution",
                                SourceVersion = "1.0", TargetVersion = "1.0", SourcePackageType = "Unmanaged", TargetPackageType = "Managed",
                                SourcePublisher = "Synthetic", TargetPublisher = "Synthetic", PackageTypeStatus = "Managed/Unmanaged Mismatch", Status = "Package Type Mismatch" } };
                        Application.DoEvents(); Assert.AreEqual(1, grid.Rows.Count);
                        var export = typeof(SolutionComparerControl).GetMethod(exportMethod, BindingFlags.Instance | BindingFlags.NonPublic);
                        export.Invoke(control, new object[] { firstPath });
                        var completed = Complete(Fixture());
                        using (var membership = new MembershipResultsForm(completed))
                        { membership.Show(); Application.DoEvents(); membership.Close(); }
                        export.Invoke(control, new object[] { secondPath });
                        Func<string, string> normalizeTime = text => System.Text.RegularExpressions.Regex.Replace(text,
                            @"\d{4}-\d{2}-\d{2} \d{2}:\d{2}", "export timestamp");
                        Assert.AreEqual(normalizeTime(File.ReadAllText(firstPath)), normalizeTime(File.ReadAllText(secondPath)));
                        StringAssert.Contains(File.ReadAllText(secondPath), "synthetic_solution");
                        Assert.IsFalse(File.ReadAllText(secondPath).Contains("DisplayEvidence")); host.Close();
                    }
                }
                finally { File.Delete(firstPath); File.Delete(secondPath); }
            });
        }

        [DataTestMethod]
        [DataRow(840)]
        [DataRow(10000)]
        public void SyntheticProjectionPerformanceAndCacheLifetime(int count)
        {
            // Warm JIT before measurement. No timing pass/fail thresholds are hardware assumptions.
            Complete(Fixture());
            var outcome = MeasureDetached(count);
            GC.Collect(); GC.WaitForPendingFinalizers(); GC.Collect();
            Assert.IsFalse(outcome.SourceCache.IsAlive); Assert.IsFalse(outcome.TargetCache.IsAlive);
            Assert.AreEqual(count, outcome.Completed.DisplayEvidence.Count);
            var afterCollection = GC.GetTotalMemory(true);
            var presentation = outcome.Completed.Presentation;
            outcome.Completed = null;
            var withoutDisplay = GC.GetTotalMemory(true);
            TestContext.WriteLine(outcome.Metrics + "; retained_display_heap_delta_bytes=" + Math.Max(0, afterCollection - withoutDisplay));
            GC.KeepAlive(presentation);
        }

        [DataTestMethod, TestCategory("P2AObjectMetadata")]
        [DataRow(false), DataRow(true)]
        public void B1CompletedProjectionKeepsObjectMetadataUnrequestedAndAuthorityIntact(bool mutateCache)
        {
            var f = Fixture();
            if (mutateCache)
                foreach (var cache in new[] { f.Source.Cache, f.Target.Cache })
                    foreach (var identity in f.Presentation.Source.Snapshot.Components.Concat(f.Presentation.Target.Snapshot.Components))
                        cache.StoreEntityRow(new Entity("appmodule", identity.Record.ObjectId.Value)
                        { ["ismanaged"] = true, ["modifiedon"] = DateTime.UtcNow,
                            ["ownerid"] = new EntityReference("team", Guid.NewGuid()) { Name = "Never borrowed" },
                            ["statecode"] = new OptionSetValue(1) }, identity.Record.ObjectId.Value);
            var before = new MembershipCoverageCsvExporter().CreateCsv(f.Presentation);
            var completed = Complete(f);
            Assert.AreSame(f.Presentation, completed.Presentation);
            foreach (var row in completed.DisplayEvidence)
            {
                ComponentDisplayEvidenceTests.AssertObjectNotRetrieved(row.SourceDisplayEvidence.ObjectMetadata);
                ComponentDisplayEvidenceTests.AssertObjectNotRetrieved(row.TargetDisplayEvidence.ObjectMetadata);
                Assert.AreEqual(DisplayEnvironmentSide.Source, row.SourceDisplayEvidence.ObjectMetadata.EnvironmentSide);
                Assert.AreEqual(DisplayEnvironmentSide.Target, row.TargetDisplayEvidence.ObjectMetadata.EnvironmentSide);
            }
            Assert.AreEqual(before, new MembershipCoverageCsvExporter().CreateCsv(completed.Presentation));
            var adapters = MembershipGridRow.CreateRows(f.Presentation, completed);
            for (int i = 0; i < adapters.Count; i++) Assert.AreSame(f.Presentation.Rows[i], adapters[i].AuthoritativeRow);
        }

        [TestMethod]
        public void ActualProjectorCancellationDuringLargeIndexingIsDiscardedAndMeasured()
        {
            var f = Fixture(); IndexMetadata(f.Source.Cache, 100000);
            using (var cancel = new CancellationTokenSource())
            {
                var elapsed = Stopwatch.StartNew();
                using (var timer = new Timer(_ => cancel.Cancel(), null, 1, Timeout.Infinite))
                {
                    Assert.ThrowsException<OperationCanceledException>(() => SolutionComparerControl.CompleteMembershipComparison(
                        f.Presentation, f.Operation, f.Source, f.Target, cancel.Token));
                }
                elapsed.Stop();
                TestContext.WriteLine("Actual P1B projection with 100000 indexed attributes; cancel scheduled after 1ms; observation_ms=" + elapsed.Elapsed.TotalMilliseconds);
            }
        }

        private static string IdentitySignature(ComponentIdentity i) => i.Record.SolutionComponentId + ":" + i.Record.ObjectId + ":" + i.Status + ":" + i.ComponentTypeKey + ":" + i.ComparisonKey;
        private static string DefinitionSignature(ComponentDefinition d) => IdentitySignature(d.Identity) + ":" + d.Status + ":" + string.Join(";", d.ComparableProperties.Select(p => p.Key + "=" + p.Value));
        private static string Summary(MembershipComparisonPresentation p) => MembershipResultsForm.FormatSummary(p.Summary);
        private static MembershipEnvironmentResult Environment(MembershipSnapshot snapshot) => MembershipEnvironmentResult.FromSnapshot(snapshot.Environment.DisplayName, snapshot, 0, TimeSpan.Zero);
        private static object Token(DataverseComponentDefinitionOperation operation) => typeof(DataverseComponentDefinitionOperation)
            .GetField("invocationIdentity", BindingFlags.Instance | BindingFlags.NonPublic).GetValue(operation);
        private static DataverseComponentDefinitionOperation.DefinitionReadWithCacheResult Envelope(MembershipSnapshot snapshot,
            DataverseComponentMetadataCache cache, DataverseComponentDefinitionOperation operation) =>
            new DataverseComponentDefinitionOperation.DefinitionReadWithCacheResult(new ComponentDefinitionSnapshot(snapshot, new ComponentDefinition[0]),
                cache, Token(operation), snapshot.Environment.OrganizationId, snapshot.SolutionUniqueName);
        private static ComponentIdentity Identity(int type, string key = "same-key", Guid? id = null,
            IdentityResolutionStatus status = IdentityResolutionStatus.Resolved, string kind = null,
            InventoryAbsencePolicy policy = InventoryAbsencePolicy.CompleteInventory) => new ComponentIdentity(
                new SolutionComponentRecord(Guid.NewGuid(), type, id ?? Guid.NewGuid()), status,
                status == IdentityResolutionStatus.Resolved ? key : null, semanticKind: kind, inventoryAbsencePolicy: policy);
        private static Entity App(ComponentIdentity identity, string display, string internalName) => new Entity("appmodule", identity.Record.ObjectId.Value)
            { ["name"] = display, ["uniquename"] = internalName };
        private static FixtureData Fixture(ComponentIdentity[] source = null, ComponentIdentity[] target = null,
            bool sourceUnavailable = false, bool targetUnavailable = false)
        {
            source = source ?? new[] { Identity(80) }; target = target ?? new[] { Identity(80) };
            var operation = new DataverseComponentDefinitionOperation();
            var s = MembershipSnapshot.Complete(Solution(), source, DateTimeOffset.UtcNow);
            var t = MembershipSnapshot.Complete(Solution(), target, DateTimeOffset.UtcNow);
            var sc = new DataverseComponentMetadataCache(); var sourceCreated = Stopwatch.GetTimestamp();
            foreach (var item in source.Where(i => i.Record.ComponentType == 80)) sc.StoreEntityRow(App(item, "Source display", "source_internal"), item.Record.ObjectId.Value);
            var tc = new DataverseComponentMetadataCache(); var targetCreated = Stopwatch.GetTimestamp();
            foreach (var item in target.Where(i => i.Record.ComponentType == 80)) tc.StoreEntityRow(App(item, "Target display", "target_internal"), item.Record.ObjectId.Value);
            var sr = sourceUnavailable ? MembershipEnvironmentResult.Unavailable("Source", "sample", 0, TimeSpan.Zero, "Synthetic fault") : Environment(s);
            var tr = targetUnavailable ? MembershipEnvironmentResult.Unavailable("Target", "sample", 0, TimeSpan.Zero, "Synthetic fault") : Environment(t);
            var presentation = new MembershipResultPresenter().Create(sr, tr);
            var sd = sourceUnavailable ? null : new ComponentDefinitionSnapshot(s, new ComponentDefinition[0]);
            var td = targetUnavailable ? null : new ComponentDefinitionSnapshot(t, new ComponentDefinition[0]);
            presentation = new ComponentDefinitionResultPresenter().Apply(presentation, sd, td);
            return new FixtureData { Operation = operation, Presentation = presentation, SourceCacheCreated = sourceCreated, TargetCacheCreated = targetCreated,
                Source = sourceUnavailable ? null : Envelope(s, sc, operation), Target = targetUnavailable ? null : Envelope(t, tc, operation) };
        }
        private static CompletedMembershipComparison Complete(FixtureData f) => Complete(f, f.Source, f.Target);
        private static CompletedMembershipComparison Complete(FixtureData f, DataverseComponentDefinitionOperation.DefinitionReadWithCacheResult source,
            DataverseComponentDefinitionOperation.DefinitionReadWithCacheResult target) => SolutionComparerControl.CompleteMembershipComparison(
                f.Presentation, f.Operation, source, target, CancellationToken.None);
        private static FakeOrganizationService ReadService(SolutionIdentity solution, Entity component, bool empty, List<string> ledger)
        {
            return new FakeOrganizationService
            {
                ExecuteRequest = request => { ledger.Add("Execute:" + request.RequestName); return request is WhoAmIRequest ? WhoAmI(solution.Environment.OrganizationId) : throw new NotSupportedException(request.RequestName); },
                RetrievePage = query =>
                {
                    ledger.Add("Query:" + query.EntityName + ":" + string.Join(",", query.ColumnSet.Columns) + ":page" + query.PageInfo.PageNumber);
                    if (query.EntityName == "solution") return Rows(SolutionRow(solution));
                    if (query.EntityName == "solutioncomponent") return empty ? Rows() : Rows(component);
                    if (query.EntityName == "webresource")
                    {
                        var id = (Guid)component["objectid"];
                        return Rows(new Entity("webresource", id) { ["webresourceid"] = id, ["name"] = "synthetic.js", ["displayname"] = "Synthetic script", ["webresourcetype"] = new OptionSetValue(3), ["content"] = "YWJj", ["description"] = "fixture" });
                    }
                    throw new NotSupportedException(query.EntityName);
                }
            };
        }
        private static void IndexMetadata(DataverseComponentMetadataCache cache, int count)
        {
            cache.ParentMetadata = new ParentEntityMetadataInventory();
            var parent = new EntityMetadata { MetadataId = Guid.NewGuid(), LogicalName = "synthetic_table" };
            typeof(EntityMetadata).GetProperty("Attributes").SetValue(parent, Enumerable.Range(0, count).Select(i => new StringAttributeMetadata
                { MetadataId = Guid.NewGuid(), LogicalName = "field_" + i }).Cast<AttributeMetadata>().ToArray());
            cache.Store(parent);
        }
        [MethodImpl(MethodImplOptions.NoInlining)]
        private static Measurement MeasureDetached(int count)
        {
            var lifetime = Stopwatch.StartNew();
            var source = Enumerable.Range(0, count).Select(i => Identity(80, "key-" + i)).ToArray();
            var target = Enumerable.Range(0, count).Select(i => Identity(80, "key-" + i)).ToArray();
            var f = Fixture(source, target); var sourceAliveAtTarget = f.Source.Cache != null && f.Target.Cache != null;
            IndexMetadata(f.Source.Cache, count); IndexMetadata(f.Target.Cache, count);
            var before = GC.GetTotalMemory(true); long peak = before;
            var sw = Stopwatch.StartNew();
            using (var timer = new Timer(_ => { var value = GC.GetTotalMemory(false); long old; do { old = Interlocked.Read(ref peak); if (value <= old) break; } while (Interlocked.CompareExchange(ref peak, value, old) != old); }, null, 0, 1))
            {
                var completed = Complete(f); sw.Stop();
                peak = Math.Max(Interlocked.Read(ref peak), GC.GetTotalMemory(false));
                return new Measurement { Completed = completed, SourceCache = new WeakReference(f.Source.Cache), TargetCache = new WeakReference(f.Target.Cache),
                    Metrics = "rows=" + count + "; indexed_attributes=" + count * 2 + "; projection_ms=" + sw.Elapsed.TotalMilliseconds +
                        "; temporary_plus_output_heap_delta_bytes=" + Math.Max(0, peak - before) + "; process_peak_working_set_bytes=" + Process.GetCurrentProcess().PeakWorkingSet64 +
                        "; cache_alive_through_target=" + sourceAliveAtTarget + "; fixture_lifetime_ms=" + lifetime.Elapsed.TotalMilliseconds +
                        "; source_cache_retained_ms=" + (Stopwatch.GetTimestamp() - f.SourceCacheCreated) * 1000.0 / Stopwatch.Frequency +
                        "; target_cache_retained_ms=" + (Stopwatch.GetTimestamp() - f.TargetCacheCreated) * 1000.0 / Stopwatch.Frequency };
            }
        }
        // The detached memory measurement is a heap observation, not an exact object allocation profiler.
        private sealed class Measurement
        {
            internal CompletedMembershipComparison Completed; internal WeakReference SourceCache, TargetCache;
            internal string Metrics;
        }
        private sealed class FixtureData
        {
            internal DataverseComponentDefinitionOperation Operation; internal MembershipComparisonPresentation Presentation;
            internal long SourceCacheCreated, TargetCacheCreated;
            internal DataverseComponentDefinitionOperation.DefinitionReadWithCacheResult Source, Target;
        }
        private static void OnSta(Action action)
        {
            Exception error = null; var thread = new Thread(() => { try { action(); } catch (Exception ex) { error = ex; } });
            thread.SetApartmentState(ApartmentState.STA); thread.Start(); thread.Join();
            if (error != null) ExceptionDispatchInfo.Capture(error).Throw();
        }
    }
}
