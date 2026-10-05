using System;
using System.Linq;
using System.Threading;
using System.Windows.Forms;
using D365SolutionComparer.Models.Identity;
using D365SolutionComparer.Models.Membership;
using D365SolutionComparer.Services.Membership;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using static D365SolutionComparer.Tests.MembershipTestData;

namespace D365SolutionComparer.Tests
{
    [TestClass, TestCategory("Phase2GInventory")]
    public class UnsupportedCoverageInventoryTests
    {
        [TestMethod]
        public void DiagnosticTypesAndUiAreExcludedFromRelease()
        {
            var assembly = typeof(SolutionComparerControl).Assembly;
#if DEBUG
            Assert.IsNotNull(assembly.GetType("D365SolutionComparer.Services.Membership.UnsupportedCoverageInventory"));
            Assert.IsNotNull(assembly.GetType("D365SolutionComparer.UnsupportedCoverageInventoryResultsForm"));
            Assert.IsNotNull(typeof(MembershipResultsForm).GetProperty("CaptureUnsupportedCoverageInventory",
                System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic));
#else
            Assert.IsNull(assembly.GetType("D365SolutionComparer.Services.Membership.UnsupportedCoverageInventory"));
            Assert.IsNull(assembly.GetType("D365SolutionComparer.Services.Membership.UnsupportedInventorySession"));
            Assert.IsNull(assembly.GetType("D365SolutionComparer.Services.Membership.InventoryFormattedLabels"));
            Assert.IsNull(assembly.GetType("D365SolutionComparer.UnsupportedCoverageInventoryResultsForm"));
            Assert.IsNull(typeof(MembershipResultsForm).GetProperty("CaptureUnsupportedCoverageInventory",
                System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic));
            Assert.IsFalse(typeof(MembershipCoverageDetailsForm).GetConstructors()
                .SelectMany(c => c.GetParameters()).Any(p => p.Name == "captureUnsupportedCoverageInventory"));
#endif
        }

#if DEBUG
        private static readonly Guid SourceOrg = new Guid("11111111-1111-1111-1111-111111111111");
        private static readonly Guid TargetOrg = new Guid("22222222-2222-2222-2222-222222222222");
        private static readonly DateTimeOffset Capture = new DateTimeOffset(2026, 10, 5, 12, 0, 0, TimeSpan.Zero);
        private static ComponentIdentity Row(int type, IdentityResolutionStatus status = IdentityResolutionStatus.Unsupported,
            Guid? id = null, SolutionComponentDefinitionIdentity definition = null) =>
            new ComponentIdentity(new SolutionComponentRecord(Guid.NewGuid(), type, id ?? Guid.NewGuid()), status,
                status == IdentityResolutionStatus.Resolved ? "portable:" + type : null,
                registeredDefinition: definition);
        private static InventoryCheckpoint Check(bool source, string name, params ComponentIdentity[] rows) =>
            new InventoryCheckpoint(MembershipSnapshot.Complete(new SolutionIdentity(
                new EnvironmentIdentity(source ? SourceOrg : TargetOrg, source ? "CSC-ROCS-UAT" : "CSC-ROCS-PROD"),
                Guid.NewGuid(), name), rows, Capture), "Display " + name, "1.2.3.4");
        private static UnsupportedCoverageInventory Report(ComponentIdentity[] source, ComponentIdentity[] target) =>
            UnsupportedCoverageInventory.Build(new[] { Check(true, "sample", source) }, new[] { Check(false, "sample", target) });

        [TestMethod]
        public void RawTypesRemainSeparateAndStatusesAccountForEveryRawRow()
        {
            var result = Report(new[] { Row(300), Row(300), Row(26, IdentityResolutionStatus.Resolved),
                Row(29, IdentityResolutionStatus.Unresolved), Row(92, IdentityResolutionStatus.Ambiguous) }, new[] { Row(300) });
            CollectionAssert.AreEqual(new[] { 26, 29, 92, 300 }, result.Families.Select(f => f.Type).ToArray());
            var canvas = result.Families.Single(f => f.Type == 300);
            Assert.AreEqual(2, canvas.Source.Raw); Assert.AreEqual(1, canvas.Target.Raw);
            Assert.AreEqual(2, canvas.Source.Unsupported); Assert.AreEqual(0, canvas.Source.Resolved);
            Assert.AreEqual(1, result.Families.Single(f => f.Type == 26).Source.Resolved);
            Assert.AreEqual(1, result.Families.Single(f => f.Type == 29).Source.Unresolved);
            Assert.AreEqual(1, result.Families.Single(f => f.Type == 92).Source.Ambiguous);
            CollectionAssert.AreEqual(new[] { 300 }, result.Candidates.Select(f => f.Type).ToArray());
        }

        [DataTestMethod]
        [DataRow(26, IdentityResolutionStatus.Resolved, "Supported")]
        [DataRow(300, IdentityResolutionStatus.Unsupported, "UnsupportedKnownFamily")]
        [DataRow(999999, IdentityResolutionStatus.Unsupported, "UnsupportedUnknownFamily")]
        [DataRow(29, IdentityResolutionStatus.Unresolved, "Unresolved")]
        [DataRow(92, IdentityResolutionStatus.Ambiguous, "Ambiguous")]
        public void EvidenceStatusReflectsObservedResolution(int type, IdentityResolutionStatus status, string expected)
        {
            Assert.AreEqual(expected, Report(new[] { Row(type, status) }, new ComponentIdentity[0]).Families.Single().EvidenceStatus);
        }

        [TestMethod]
        public void PartiallyResolvedFamilyDoesNotBecomeSupportedOrProveAbsence()
        {
            var result = Report(new[] { Row(31, IdentityResolutionStatus.Resolved), Row(31) }, new[] { Row(31, IdentityResolutionStatus.Unresolved) });
            Assert.AreEqual("PartiallyResolved", result.Families.Single().EvidenceStatus);
            StringAssert.Contains(result.Families.Single().ResolverSupport, "Partial");
            Assert.AreEqual(1, result.Candidates.Count);
            StringAssert.Contains(result.Text, "Zero counts do not prove absence");
        }

        [TestMethod]
        public void SharedRawGuidIsOnlyAnObservationAndOneSidedGuidsNeverInferAbsence()
        {
            var common = Guid.NewGuid(); var left = Guid.NewGuid(); var right = Guid.NewGuid();
            var source = Check(true, "sample", Row(300, id: common), Row(300, id: left));
            var target = Check(false, "sample", Row(300, id: common), Row(300, id: right));
            var result = UnsupportedCoverageInventory.Build(new[] { source }, new[] { target });
            StringAssert.Contains(result.Text, "300\t" + left.ToString("D") + "\t" + right.ToString("D") + "\t" + common.ToString("D"));
            StringAssert.Contains(result.Text, "NOT membership matches");
            Assert.AreEqual(4, result.Candidates.Single().TotalDistinct);
            var comparison = new SolutionMembershipComparer().Compare(source.Snapshot, target.Snapshot);
            Assert.AreEqual(4, comparison.Count);
            Assert.IsTrue(comparison.All(c => c.Presence == MembershipPresence.Indeterminate && c.AbsenceEvidence == MembershipAbsenceEvidence.None));
        }

        [TestMethod]
        public void MultipleSolutionsAndRepeatedReferencesDoNotInflateDistinctIds()
        {
            var id = Guid.NewGuid();
            var source = new[] { Check(true, "Alpha", Row(300, id: id), Row(300, id: id)), Check(true, "Beta", Row(300, id: id)) };
            var target = new[] { Check(false, "ALPHA", Row(300, id: id)), Check(false, "Gamma", Row(300)) };
            var f = UnsupportedCoverageInventory.Build(source, target).Families.Single();
            Assert.AreEqual(3, f.Source.Raw); Assert.AreEqual(1, f.Source.Ids.Count);
            Assert.AreEqual(2, f.Target.Raw); Assert.AreEqual(2, f.Target.Ids.Count);
            Assert.AreEqual(2, f.SourceSolutions); Assert.AreEqual(2, f.TargetSolutions);
            Assert.AreEqual(1, f.SharedSolutions); Assert.AreEqual(3, f.AffectedSolutions);
        }

        [TestMethod]
        public void EmptyGuidsAndNullObjectIdsAreExcludedFromDistinctCounts()
        {
            var nullRow = new ComponentIdentity(new SolutionComponentRecord(Guid.NewGuid(), 300, null), IdentityResolutionStatus.Unsupported);
            var f = Report(new[] { nullRow, Row(300, id: Guid.Empty) }, new ComponentIdentity[0]).Families.Single();
            Assert.AreEqual(2, f.Source.Raw); Assert.AreEqual(2, f.Source.BlankIds); Assert.AreEqual(0, f.Source.Ids.Count);
        }

        [DataTestMethod]
        [DataRow("AppElement", "appelement")]
        [DataRow("appaction", "appaction")]
        [DataRow("aiskillconfig", "aiskillconfig")]
        public void RegisteredFamilyNamesAndBackingEntitiesComeFromExistingDefinitions(string name, string entity)
        {
            var definition = new SolutionComponentDefinitionIdentity(10099, name, entity);
            var f = Report(new[] { Row(10099, definition: definition) }, new ComponentIdentity[0]).Families.Single();
            StringAssert.Contains(f.CatalogLabel, name);
            Assert.AreEqual(entity, f.BackingEntity); Assert.IsTrue(f.ExistingMetadataEvidence);
            Assert.AreEqual("Unsupported", f.ResolverSupport); Assert.AreEqual("UnsupportedKnownFamily", f.EvidenceStatus);
        }

        [TestMethod]
        public void UnknownBackingRemainsUnknownRatherThanGuessedFromDisplayLabel()
        {
            var row = Row(999999);
            InventoryFormattedLabels.Capture(row.Record, "Some Table Name");
            var f = Report(new[] { row }, new ComponentIdentity[0]).Families.Single();
            Assert.AreEqual("Some Table Name", f.FormattedLabel); Assert.AreEqual("Unknown", f.BackingEntity);
            Assert.IsFalse(f.ExistingMetadataEvidence);
        }

        [TestMethod]
        public void FormattedLabelsReuseExistingReadResponseAndNormalRequestShapeIsUnchanged()
        {
            var solution = Solution(); var raw = ComponentRow(solution, 90);
            raw.FormattedValues["componenttype"] = "Plug-in Type";
            var service = Service(solution, q =>
            {
                if (q.EntityName == "solution") return Rows(SolutionRow(solution));
                Assert.AreEqual("solutioncomponent", q.EntityName);
                CollectionAssert.AreEquivalent(new[] { "solutioncomponentid", "componenttype", "objectid", "rootcomponentbehavior",
                    "rootsolutioncomponentid", "ismetadata", "solutionid" }, q.ColumnSet.Columns.ToArray());
                return Rows(raw);
            });
            var snapshot = new DataverseSolutionMembershipOperation().ReadAndResolve(service, solution, CancellationToken.None);
            Assert.AreEqual(2, service.Calls); Assert.AreEqual(1, service.ExecuteCalls);
            Assert.AreEqual(IdentityResolutionStatus.Unsupported, snapshot.Components.Single().Status);
            var before = Present(snapshot, snapshot);
            var csv = new MembershipCoverageCsvExporter().CreateCsv(before);
            var report = UnsupportedCoverageInventory.Build(new[] { new InventoryCheckpoint(snapshot) }, new[] { new InventoryCheckpoint(snapshot) });
            Assert.AreEqual("Plug-in Type", report.Families.Single().FormattedLabel);
            Assert.AreEqual(2, service.Calls); Assert.AreEqual(1, service.ExecuteCalls); Assert.AreEqual(0, service.WriteCalls);
            Assert.AreEqual(csv, new MembershipCoverageCsvExporter().CreateCsv(Present(snapshot, snapshot)));
            Assert.AreEqual(2, before.Summary.Unsupported); Assert.AreEqual(0, before.Summary.PresentInBoth);
        }

        [TestMethod]
        public void InventoryPreservesIdentityEvidenceCoverageCountsAndMembershipResults()
        {
            var source = Check(true, "sample", Row(26, IdentityResolutionStatus.Resolved), Row(300), Row(29, IdentityResolutionStatus.Unresolved));
            var target = Check(false, "sample", Row(26, IdentityResolutionStatus.Resolved), Row(300), Row(92, IdentityResolutionStatus.Ambiguous));
            var identities = source.Snapshot.Components.Concat(target.Snapshot.Components).ToArray();
            var csv = new MembershipCoverageCsvExporter().CreateCsv(Present(source.Snapshot, target.Snapshot));
            var coverage = new MembershipCoverageDiagnosticsBuilder().Build(source.Snapshot);
            UnsupportedCoverageInventory.Build(new[] { source }, new[] { target });
            CollectionAssert.AreEqual(identities, source.Snapshot.Components.Concat(target.Snapshot.Components).ToArray());
            Assert.AreEqual(csv, new MembershipCoverageCsvExporter().CreateCsv(Present(source.Snapshot, target.Snapshot)));
            var after = new MembershipCoverageDiagnosticsBuilder().Build(source.Snapshot);
            CollectionAssert.AreEqual(coverage.SemanticKinds.Select(k => k.TotalCandidates).ToArray(), after.SemanticKinds.Select(k => k.TotalCandidates).ToArray());
        }

        [TestMethod]
        public void RankingAndReportAreDeterministicRegardlessOfInputEnumerationOrder()
        {
            var source = Check(true, "A", Row(36), Row(300), Row(26, IdentityResolutionStatus.Resolved));
            var more = Check(true, "B", Row(300));
            var target = Check(false, "A", Row(36), Row(300));
            var first = UnsupportedCoverageInventory.Build(new[] { source, more }, new[] { target });
            var reversed = new InventoryCheckpoint(MembershipSnapshot.Complete(source.Snapshot.Solution,
                source.Snapshot.Components.Reverse(), Capture), source.DisplayName, source.Version);
            var second = UnsupportedCoverageInventory.Build(new[] { more, reversed }, new[] { target });
            Assert.AreEqual(first.Text, second.Text);
            CollectionAssert.AreEqual(new[] { 300, 36 }, first.Candidates.Select(f => f.Type).ToArray());
            Assert.IsTrue(first.Candidates.All(f => f.Source.Unsupported + f.Target.Unsupported > 0));
        }

        [TestMethod]
        public void RankingTieBreaksBySolutionBreadthAndRawTypeNotInputOrder()
        {
            var source = new[] { Check(true, "A", Row(300), Row(300), Row(36)), Check(true, "B", Row(36)) };
            CollectionAssert.AreEqual(new[] { 36, 300 }, UnsupportedCoverageInventory.Build(source, new InventoryCheckpoint[0])
                .Candidates.Select(f => f.Type).ToArray());
            CollectionAssert.AreEqual(new[] { 36, 300 }, Report(new[] { Row(300), Row(36) }, new ComponentIdentity[0])
                .Candidates.Select(f => f.Type).ToArray());
        }

        [TestMethod]
        public void ReportSchemaIncludesSolutionMetadataStatusCountsScopeAndRequestLedger()
        {
            var report = Report(new[] { Row(300) }, new[] { Row(300) }).Text;
            StringAssert.Contains(report, "SourceDistinctObjectIds\tSourceBlankObjectIds\tSourceResolved\tSourceUnsupported\tSourceUnresolved\tSourceAmbiguous");
            StringAssert.Contains(report, "TargetSolutions\tSharedSolutions");
            StringAssert.Contains(report, "Source\tCSC-ROCS-UAT\tsample\tDisplay sample\t1.2.3.4\t2026-10-05T12:00:00.0000000Z\tComplete");
            StringAssert.Contains(report, "Source\tsample\tDisplay sample\t1.2.3.4\t300\tUnknown\t1\t1\t0\t0\t1\t0\t0");
            StringAssert.Contains(report, "NEXT RESOLVER CANDIDATES");
            StringAssert.Contains(report, "Dataverse requests: 0; WhoAmI: 0; Writes: 0");
            StringAssert.Contains(report, "NOT an environment-wide inventory");
        }

        [TestMethod]
        public void UnavailableAndAbsentSnapshotsAreListedButExcludedFromCounts()
        {
            var environment = new EnvironmentIdentity(SourceOrg, "Source");
            var result = UnsupportedCoverageInventory.Build(new[] {
                new InventoryCheckpoint(MembershipSnapshot.Unavailable(environment, "fault", Capture, "Denied")),
                new InventoryCheckpoint(MembershipSnapshot.Absent(environment, "absent", Capture)) }, new InventoryCheckpoint[0]);
            Assert.AreEqual(0, result.Families.Count); Assert.AreEqual(0, result.Candidates.Count);
            StringAssert.Contains(result.Text, "Unavailable"); StringAssert.Contains(result.Text, "SolutionAbsent");
        }

        [TestMethod]
        public void SessionReplacesRepeatedComparisonsAndKeepsMultiSolutionEvidence()
        {
            var session = new UnsupportedInventorySession();
            var first = Check(true, "A", Row(300)); var second = Check(true, "B", Row(300));
            session.Record(first.Snapshot, Check(false, "A", Row(300)).Snapshot, "A", "1", "A", "2");
            session.Record(first.Snapshot, Check(false, "A", Row(300)).Snapshot, "A", "3", "A", "2");
            session.Record(second.Snapshot, Check(false, "B").Snapshot, "B", "4", "B", "2");
            Assert.AreEqual(2, session.Source.Length); Assert.AreEqual(2, session.Target.Length);
            Assert.AreEqual("3", session.Source.Single(c => c.Snapshot.SolutionUniqueName == "A").Version);
            Assert.AreEqual(2, UnsupportedCoverageInventory.Build(session.Source, session.Target).Families.Single().Source.Raw);
        }

        [TestMethod]
        public void SessionEnvironmentPairChangeResetsBothSidesAndFrozenScopeIsPreserved()
        {
            var session = new UnsupportedInventorySession();
            var first = Check(true, "A", Row(300)); var target = Check(false, "A", Row(300));
            session.Record(first.Snapshot, target.Snapshot, null, null, null, null);
            var frozen = session.Source;
            session.Record(Snapshot(Row(36)), target.Snapshot, null, null, null, null);
            Assert.AreEqual(36, session.Source.Single().Snapshot.Components.Single().Record.ComponentType);
            Assert.AreEqual(1, session.Target.Length);
            Assert.AreEqual(300, frozen.Single().Snapshot.Components.Single().Record.ComponentType);
        }

        [TestMethod]
        public void MixingOrganizationsOrRepeatedSolutionCheckpointsIsRejected()
        {
            var c = Check(true, "A", Row(300));
            Assert.ThrowsException<ArgumentException>(() => UnsupportedCoverageInventory.Build(new[] { c, Check(false, "B", Row(300)) }, new InventoryCheckpoint[0]));
            Assert.ThrowsException<ArgumentException>(() => UnsupportedCoverageInventory.Build(new[] { c, c }, new InventoryCheckpoint[0]));
        }

        [TestMethod]
        public void MissingEnvironmentSnapshotCannotReuseHistoryFromAnUnverifiedPair()
        {
            var session = new UnsupportedInventorySession();
            session.Record(Check(true, "A", Row(300)).Snapshot, Check(false, "A", Row(300)).Snapshot, null, null, null, null);
            session.Record(Check(true, "B", Row(36)).Snapshot, null, null, null, null, null);
            Assert.AreEqual(0, session.Target.Length);
            Assert.AreEqual("B", session.Source.Single().Snapshot.SolutionUniqueName);
        }

        [TestMethod]
        public void CancellationPropagatesWithoutReturningAPartialReport()
        {
            using (var cancellation = new CancellationTokenSource())
            {
                cancellation.Cancel();
                Assert.ThrowsException<OperationCanceledException>(() => UnsupportedCoverageInventory.Build(
                    new[] { Check(true, "A", Row(300)) }, new InventoryCheckpoint[0], cancellation.Token));
            }
        }

        [TestMethod]
        public void DebugInventoryUiDoesNoAnalysisUntilClickedAndReportIsReadOnly()
        {
            Exception error = null;
            var thread = new Thread(() =>
            {
                try
                {
                    var snapshot = Snapshot(Enumerable.Range(0, 900).Select(_ => Row(300)).ToArray());
                    var coverage = new MembershipCoverageDiagnosticsBuilder().Build(snapshot);
                    int invoked = 0;
                    using (var form = new MembershipCoverageDetailsForm("Source", coverage, "Target", coverage,
                        captureUnsupportedCoverageInventory: () => invoked++))
                    {
                        Assert.AreEqual(0, invoked);
                        var button = form.Controls.OfType<FlowLayoutPanel>().SelectMany(p => p.Controls.OfType<Button>())
                            .Single(b => b.Text == "Unsupported Coverage Inventory...");
                        form.StartPosition = FormStartPosition.Manual;
                        form.Location = new System.Drawing.Point(-4000, -4000);
                        form.Show(); Application.DoEvents();
                        Assert.AreEqual(0, invoked);
                        button.PerformClick(); Assert.AreEqual(1, invoked);
                    }
                    using (var form = new UnsupportedCoverageInventoryResultsForm("complete evidence"))
                    {
                        var body = form.Controls.OfType<RichTextBox>().Single();
                        Assert.IsTrue(body.ReadOnly); Assert.IsFalse(body.WordWrap); Assert.AreEqual("complete evidence", body.Text);
                    }
                }
                catch (Exception ex) { error = ex; }
            }) { IsBackground = true };
            thread.SetApartmentState(ApartmentState.STA); thread.Start();
            Assert.IsTrue(thread.Join(TimeSpan.FromSeconds(30)), "Inventory UI initialization must remain responsive.");
            if (error != null) System.Runtime.ExceptionServices.ExceptionDispatchInfo.Capture(error).Throw();
        }

        private static MembershipComparisonPresentation Present(MembershipSnapshot source, MembershipSnapshot target) =>
            new MembershipResultPresenter().Create(MembershipEnvironmentResult.FromSnapshot("Source", source, 3, TimeSpan.Zero),
                MembershipEnvironmentResult.FromSnapshot("Target", target, 3, TimeSpan.Zero));
#endif
    }
}
