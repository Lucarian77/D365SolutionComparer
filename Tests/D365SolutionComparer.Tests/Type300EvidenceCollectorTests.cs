using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.ServiceModel;
using System.Threading;
using System.Windows.Forms;
using D365SolutionComparer.Models.Membership;
using D365SolutionComparer.Services.Membership;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Query;
using Label = Microsoft.Xrm.Sdk.Label;
using static D365SolutionComparer.Tests.MembershipTestData;

namespace D365SolutionComparer.Tests
{
    [TestClass, TestCategory("Phase2GType300Evidence")]
    public class Type300EvidenceCollectorTests
    {
        [TestMethod]
        public void NormalType300MembershipStaysUnsupportedIndeterminateWithoutPortableKeyOrDefinition()
        {
            var source = Snapshot(new ComponentIdentity(new SolutionComponentRecord(Guid.NewGuid(), 300, Guid.NewGuid()), IdentityResolutionStatus.Unsupported));
            var target = Snapshot(new ComponentIdentity(new SolutionComponentRecord(Guid.NewGuid(), 300, Guid.NewGuid()), IdentityResolutionStatus.Unsupported));
            Assert.AreEqual("unsupported:componenttype:300", ComponentSemanticKinds.FromRawComponentType(300));
            Assert.IsTrue(new SolutionMembershipComparer().Compare(source, target).All(r => r.Presence == MembershipPresence.Indeterminate));
            Assert.IsTrue(source.Components.Concat(target.Components).All(c => c.ComparisonKey == null));
            Assert.IsNull(D365SolutionComparer.Models.ComponentDetails.ComponentDefinitionContractCatalog.For("unsupported:componenttype:300"));
        }
        [TestMethod]
        public void CollectorAndCaptureActionAreExcludedFromRelease()
        {
            var assembly = typeof(SolutionComparerControl).Assembly;
#if DEBUG
            Assert.IsNotNull(assembly.GetType("D365SolutionComparer.Services.Membership.Type300EvidenceCollector"));
            Assert.IsNotNull(typeof(MembershipResultsForm).GetProperty("CaptureType300Evidence", BindingFlags.Instance | BindingFlags.NonPublic));
#else
            Assert.IsNull(assembly.GetType("D365SolutionComparer.Services.Membership.Type300EvidenceCollector"));
            Assert.IsNull(assembly.GetType("D365SolutionComparer.Type300EvidenceResultsForm"));
            Assert.IsNull(typeof(MembershipResultsForm).GetProperty("CaptureType300Evidence", BindingFlags.Instance | BindingFlags.NonPublic));
            Assert.IsNull(typeof(SolutionComparerControl).GetMethod("CaptureType300Evidence", BindingFlags.Instance | BindingFlags.NonPublic));
            Assert.IsFalse(typeof(MembershipCoverageDetailsForm).GetConstructors().SelectMany(c => c.GetParameters()).Any(p => p.Name == "captureType300Evidence"));
#endif
        }
#if DEBUG
        private const string Secret = "PRIVATE-CANVAS-JSON-AND-URI-NOT-FOR-REPORT";
        [TestMethod]
        public void DebugCaptureButtonIsExplicitAndDoesNotRunDuringCoverageFormInitialization()
        {
            Exception error = null;
            var thread = new Thread(() =>
            {
                try
                {
                    var pair = new Pair(); pair.Source.Add(); pair.Target.Add(); var source = pair.Source.Snapshot(); var target = pair.Target.Snapshot();
                    var presentation = new MembershipResultPresenter().Create(MembershipEnvironmentResult.FromSnapshot("Source", source, 4, TimeSpan.Zero),
                        MembershipEnvironmentResult.FromSnapshot("Target", target, 4, TimeSpan.Zero)); int calls = 0;
                    using (var form = new MembershipCoverageDetailsForm("Source", new MembershipCoverageDiagnosticsBuilder().Build(source), "Target",
                        new MembershipCoverageDiagnosticsBuilder().Build(target), presentation, captureType300Evidence: () => calls++))
                    {
                        form.StartPosition = FormStartPosition.Manual; form.Location = new System.Drawing.Point(-4000, -4000); form.Show(); Application.DoEvents();
                        var button = form.Controls.OfType<FlowLayoutPanel>().SelectMany(p => p.Controls.OfType<Button>()).Single(b => b.Text == "Capture Type 300 Canvas App Evidence...");
                        Assert.IsTrue(button.Enabled); Assert.AreEqual(0, calls); button.PerformClick(); Assert.AreEqual(1, calls);
                    }
                    Assert.AreEqual(0, pair.Source.Service.Calls + pair.Target.Service.Calls);
                    using (var form = new Type300EvidenceResultsForm("redacted")) Assert.IsTrue(form.Controls.OfType<RichTextBox>().Single().ReadOnly);
                }
                catch (Exception ex) { error = ex; }
            }) { IsBackground = true };
            thread.SetApartmentState(ApartmentState.STA); thread.Start(); Assert.IsTrue(thread.Join(TimeSpan.FromSeconds(30)));
            if (error != null) System.Runtime.ExceptionServices.ExceptionDispatchInfo.Capture(error).Throw();
        }

        [TestMethod]
        public void ZeroMemberSidesStopBeforeSchemaBackingAndDependencyReads()
        {
            var pair = new Pair(); pair.Source.AddElement(Guid.NewGuid()); pair.Target.AddElement(Guid.NewGuid()); var report = pair.Capture();
            Assert.AreEqual(0, pair.Source.Service.Calls + pair.Target.Service.Calls + pair.Source.Service.ExecuteCalls + pair.Target.Service.ExecuteCalls);
            Assert.AreEqual(0, report.Pairs.Count); StringAssert.Contains(report.Build(), "raw=0");
        }
        [TestMethod]
        public void ZeroMemberTargetDoesNotTurnDiagnosticOneSidedEvidenceIntoMembershipAbsence()
        {
            var pair = new Pair(); pair.Source.Add(); var report = pair.Capture();
            Assert.AreEqual("OneSidedEvidence", report.Pairs.Single().Outcome); Assert.AreEqual(0, pair.Target.Service.Calls + pair.Target.Service.ExecuteCalls);
        }
        [TestMethod]
        public void DirectCorrelationDifferentPrimaryIdsCaseInsensitiveCandidateAndManagedTransition()
        {
            var pair = new Pair(); var left = pair.Source.Add(name: " publisher_Canvas "); var right = pair.Target.Add(name: "PUBLISHER_canvas"); right["ismanaged"] = true;
            var report = pair.Capture(); var match = report.Pairs.Single();
            Assert.AreNotEqual(left.Id, right.Id); Assert.AreEqual("SemanticPair", match.Outcome);
            Assert.AreEqual(left.Id, match.Source.PrimaryId); Assert.AreEqual(right.Id, match.Target.PrimaryId);
            Assert.IsTrue(StringComparer.OrdinalIgnoreCase.Equals(match.Source.CandidateA, match.Target.CandidateA));
            Assert.IsTrue(match.Categories.Contains("DifferentPrimaryId")); Assert.IsTrue(match.Categories.Contains("DifferentUniqueId"));
            Assert.IsTrue(match.Categories.Contains("UnmanagedToManaged"));
            CollectionAssert.AreEquivalent(new[] { "canvasappid", "name" }, pair.Source.Queries.First().ColumnSet.Columns.ToArray());
            Assert.AreEqual(2, pair.Source.Service.Calls); Assert.AreEqual(2, pair.Source.Service.ExecuteCalls);
        }
        [TestMethod]
        public void DifferentCandidatesNeverPairUsingDisplayCompositeHashOrMatchingGuid()
        {
            var pair = new Pair(); var left = pair.Source.Add(name: "publisher_one"); pair.Target.Add(left.Id, "publisher_two");
            var report = pair.Capture(); Assert.AreEqual(2, report.Pairs.Count); Assert.IsTrue(report.Pairs.All(p => p.Outcome == "OneSidedEvidence"));
            Assert.AreEqual(report.Source.Rows.Values.Single().CandidateB, report.Target.Rows.Values.Single().CandidateB);
        }
        [TestMethod]
        public void CandidateACollisionCannotBeRepairedByDistinctCandidateBOrContent()
        {
            var pair = new Pair(); pair.Source.Add(); var duplicate = pair.Source.Add(name: "PUBLISHER_canvas"); duplicate["displayname"] = "Other Display"; duplicate["configuration"] = "Other config"; pair.Target.Add();
            var report = pair.Capture(); Assert.IsTrue(report.Source.Rows.Values.All(r => r.DuplicateA));
            Assert.IsTrue(report.Pairs.All(p => p.Outcome == "Ambiguous"));
            Assert.AreEqual(2, report.Source.Rows.Values.Select(r => r.CandidateB).Distinct().Count());
        }
        [TestMethod]
        public void CandidateBCollisionDoesNotInvalidateUniqueAOrCreatePairing()
        {
            var pair = new Pair(); pair.Source.Add(name: "one"); pair.Source.Add(name: "two"); pair.Target.Add(name: "one");
            var report = pair.Capture(); Assert.IsTrue(report.Source.Rows.Values.All(r => r.DuplicateB));
            Assert.AreEqual(1, report.Pairs.Count(p => p.Outcome == "SemanticPair")); Assert.AreEqual(1, report.Pairs.Count(p => p.Outcome == "OneSidedEvidence"));
        }
        [DataTestMethod, DataRow(""), DataRow(" "), DataRow("84d7cdda-0013-4305-a9b9-a69d741e88fe")]
        public void BlankOrGuidOnlyInternalNameDoesNotInventSemanticCandidate(string name)
        {
            var pair = new Pair(); pair.Source.Add(name: name); var row = pair.Capture().Source.Rows.Values.Single();
            Assert.AreEqual("Unique", row.Status); Assert.IsNull(row.CandidateA);
        }
        [TestMethod]
        public void StrongestSchemaIdentifierIsExplicitAndBlankStrongestNeverFallsBack()
        {
            var pair = new Pair(); var row = pair.Source.Add(); pair.Source.Attributes.Add(Attribute("uniquename", new StringAttributeMetadata())); row["uniquename"] = "";
            var report = pair.Capture(); Assert.AreEqual("uniquename", report.Source.CandidateField); Assert.IsNull(report.Source.Rows.Values.Single().CandidateA);
        }
        [TestMethod]
        public void MalformedRuntimeIdentifierCannotBecomeSemanticText()
        {
            var pair = new Pair(); var row = pair.Source.Add(); row["name"] = new OptionSetValue(1);
            Assert.IsNull(pair.Capture().Source.Rows.Values.Single().CandidateA);
        }
        [TestMethod]
        public void IsolationLimitKeepsUnretrievedContentUnavailable()
        {
            var pair = new Pair(); pair.Source.Add();
            var bad = Enumerable.Range(0, 80).Select(i => "audit" + i.ToString("D2")).ToArray();
            foreach (var field in bad) pair.Source.Attributes.Add(Attribute(field, new MemoAttributeMetadata()));
            var normal = pair.Source.Service.RetrievePage;
            pair.Source.Service.RetrievePage = q => { var rows = normal(q); if (bad.Any(q.ColumnSet.Columns.Contains)) throw new TimeoutException(Secret); return rows; };
            var report = pair.Capture(); Assert.IsTrue(pair.Source.Service.Calls <= Type300EvidenceCollector.MaxIsolationGroups + 1);
            Assert.IsTrue(bad.All(f => !report.Source.Rows.Values.Single().Content[f].Known));
            StringAssert.Contains(report.Build(), "after isolation limit=64");
        }
        [TestMethod]
        public void OneSidedBlankRawObjectIdRemainsIncompleteWithoutRequests()
        {
            var pair = new Pair(); pair.Source.Reference(300, null); var report = pair.Capture();
            Assert.AreEqual(1, report.Source.Raw.Count); Assert.AreEqual(0, report.Source.Rows.Count); StringAssert.Contains(report.Build(), "blank=1");
            Assert.AreEqual("Incomplete", report.Pairs.Single().Outcome);
            Assert.AreEqual(0, pair.Source.Service.Calls + pair.Source.Service.ExecuteCalls);
        }
        [TestMethod]
        public void OptionalRuntimeFaultIsIsolatedWithoutInvalidatingPrimaryOrCriticalCandidate()
        {
            var pair = new Pair(); pair.Source.Add(); Fail(pair.Source, "configuration"); var report = pair.Capture(); var row = report.Source.Rows.Values.Single();
            Assert.AreEqual("Unique", row.Status); Assert.IsTrue(row.CriticalComplete); Assert.IsNotNull(row.CandidateA);
            Assert.IsFalse(row.Content["configuration"].Known); Assert.IsTrue(row.Content["appopenuri"].Known);
            StringAssert.Contains(report.Build(), "Attribute faulted: configuration"); StringAssert.Contains(report.Build(), "SDKErrorCode=0x8004023B");
            Assert.IsFalse(report.Build().Contains(Secret)); Assert.IsFalse(report.Build().Contains("STACK-TRACE")); Assert.IsFalse(report.Build().Contains("Bearer"));
        }
        [TestMethod]
        public void CriticalRuntimeFaultKeepsPrimaryAuditButBlocksCandidateAndOptionalReads()
        {
            var pair = new Pair(); var raw = pair.Source.Add(); Fail(pair.Source, "name"); var report = pair.Capture(); var row = report.Source.Rows.Values.Single();
            Assert.AreEqual("Unique", row.Status); Assert.AreEqual(raw.Id, row.PrimaryId); Assert.IsFalse(row.CriticalComplete); Assert.IsNull(row.CandidateA);
            Assert.IsFalse(pair.Source.Queries.Any(q => q.ColumnSet.Columns.Contains("configuration")));
        }
        [TestMethod]
        public void PrimaryRuntimeFaultRemainsFaultedAndNoCandidateIsEvaluated()
        {
            var pair = new Pair(); pair.Source.Add(); Fail(pair.Source, "canvasappid"); var row = pair.Capture().Source.Rows.Values.Single();
            Assert.AreEqual("Faulted", row.Status); Assert.IsNull(row.PrimaryId); Assert.IsNull(row.CandidateA); Assert.AreEqual(2, pair.Source.Service.Calls);
        }
        [TestMethod]
        public void OptionalCorrelationContradictionPreservesPrimaryButConservativelyBlocksPairing()
        {
            var pair = new Pair(); pair.Source.Add(); var normal = pair.Source.Service.RetrievePage;
            pair.Source.Service.RetrievePage = q => q.ColumnSet.Columns.Contains("configuration") ? Rows() : normal(q);
            var row = pair.Capture().Source.Rows.Values.Single(); Assert.AreEqual("Unique", row.Status); Assert.IsNotNull(row.PrimaryId);
            Assert.IsFalse(row.CriticalComplete); Assert.IsNull(row.CandidateA);
        }
        [DataTestMethod, DataRow(200, 2), DataRow(201, 4)]
        public void ScopedBatchesNeverExceed200AndDoNotQueryPerApp(int count, int calls)
        {
            var pair = new Pair(); for (int i = 0; i < count; i++) pair.Source.Add(name: "app" + i);
            var report = pair.Capture(); Assert.AreEqual(calls, pair.Source.Service.Calls);
            Assert.IsTrue(pair.Source.Queries.All(q => q.Criteria.Conditions.Single().Values.Count <= 200));
            Assert.AreEqual(count, report.Source.Rows.Count); Assert.AreEqual(2, pair.Source.Service.ExecuteCalls);
        }
        [TestMethod]
        public void TerminalPagingCookieIsValidAndNotExposed()
        {
            var pair = new Pair(); pair.Source.Add(); var normal = pair.Source.Service.RetrievePage;
            pair.Source.Service.RetrievePage = q => { var rows = normal(q); rows.PagingCookie = "PRIVATE-COOKIE"; return rows; };
            var report = pair.Capture(); Assert.AreEqual("Unique", report.Source.Rows.Values.Single().Status);
            StringAssert.Contains(report.Build(), "PagingCookieSupplied=True"); Assert.IsFalse(report.Build().Contains("PRIVATE-COOKIE"));
        }
        [TestMethod]
        public void GenuinePagingDeduplicatesIdenticalCrossPageRowsAfterTerminalRetrieval()
        {
            var pair = new Pair(); var one = pair.Source.Add(name: "one"); pair.Source.Add(name: "two"); var normal = pair.Source.Service.RetrievePage;
            pair.Source.Service.RetrievePage = q => { var rows = normal(q); return new EntityCollection(rows.Entities.Where(r => q.PageInfo.PageNumber != 1 || r.Id == one.Id).ToList())
                { MoreRecords = q.PageInfo.PageNumber == 1, PagingCookie = q.PageInfo.PageNumber == 1 ? "first" : "terminal" }; };
            var report = pair.Capture(); Assert.IsTrue(report.Source.Rows.Values.All(r => r.Status == "Unique")); Assert.AreEqual(4, pair.Source.Service.Calls);
        }
        [TestMethod]
        public void StalledPagingCannotProduceCandidateOrMissingFinding()
        {
            var pair = new Pair(); pair.Source.Add(); var normal = pair.Source.Service.RetrievePage;
            pair.Source.Service.RetrievePage = q => { var rows = normal(q); rows.MoreRecords = true; rows.PagingCookie = "stalled"; return rows; };
            var row = pair.Capture().Source.Rows.Values.Single(); Assert.AreEqual("Incomplete", row.Status); Assert.IsNull(row.CandidateA);
        }
        [TestMethod]
        public void MissingBackingIsLocalCorrelationMissingAndNotMembershipAbsence()
        {
            var pair = new Pair(); pair.Source.Reference(300, Guid.NewGuid()); var report = pair.Capture(); Assert.AreEqual("Missing", report.Source.Rows.Values.Single().Status);
            Assert.AreEqual("Incomplete", report.Pairs.Single().Outcome);
        }
        [DataTestMethod, DataRow(false), DataRow(true)]
        public void DuplicateOrConflictingRowsStayConservative(bool conflict)
        {
            var pair = new Pair(); pair.Source.Add(); var normal = pair.Source.Service.RetrievePage;
            pair.Source.Service.RetrievePage = q => { var rows = normal(q); var extra = new Entity("canvasapp", rows.Entities[0].Id);
                foreach (var field in rows.Entities[0].Attributes) extra[field.Key] = field.Value;
                if (conflict) extra["name"] = "conflicting"; rows.Entities.Add(extra); return rows; };
            var row = pair.Capture().Source.Rows.Values.Single(); Assert.AreEqual("Duplicate", row.Status); Assert.IsNull(row.CandidateA);
        }
        [TestMethod]
        public void ConflictingPrimaryAttributeOrForeignRowMakesBatchIncomplete()
        {
            var pair = new Pair(); pair.Source.Add(); var normal = pair.Source.Service.RetrievePage;
            pair.Source.Service.RetrievePage = q => { var rows = normal(q); rows.Entities[0]["canvasappid"] = Guid.NewGuid(); return rows; };
            var row = pair.Capture().Source.Rows.Values.Single(); Assert.AreEqual("Incomplete", row.Status); Assert.IsNull(row.CandidateA);
        }
        [TestMethod]
        public void RepeatedRawReferencesCollapseForCandidateUniquenessOnly()
        {
            var pair = new Pair(); var one = pair.Source.Add(); pair.Source.Reference(300, one.Id); pair.Target.Add(); var report = pair.Capture();
            Assert.AreEqual(2, report.Source.Raw.Count); Assert.AreEqual(1, report.Source.Rows.Count); Assert.AreEqual("SemanticPair", report.Pairs.Single().Outcome);
            StringAssert.Contains(report.Build(), "RepeatedRawObjectId=");
        }
        [TestMethod]
        public void PayloadsAndLookupShadowsAreExcludedAndLargeTextAndUrisAreHashOnly()
        {
            var pair = new Pair(); pair.Source.Add(); foreach (var field in new[] { "document", "thumbnail", "binarypackage", "attachment", "media", "owneridname", "owneridyominame" })
                pair.Source.Attributes.Add(Attribute(field, new StringAttributeMetadata()));
            var report = pair.Capture(); Assert.IsTrue(pair.Source.Queries.All(q => !q.ColumnSet.Columns.Any(c => new[] { "document", "thumbnail", "binarypackage", "attachment", "media", "owneridname", "owneridyominame" }.Contains(c))));
            var row = report.Source.Rows.Values.Single(); Assert.IsTrue(row.Content["configuration"].Known); Assert.IsTrue(row.Content["appopenuri"].Known); Assert.IsFalse(report.Build().Contains(Secret));
        }
        [TestMethod]
        public void SchemaUnavailableNeverGuessesBackingColumns()
        {
            var pair = new Pair(); pair.Source.Add(); pair.Source.Service.ExecuteRequest = r => new RetrieveEntityResponse();
            var row = pair.Capture().Source.Rows.Values.Single(); Assert.AreEqual("Incomplete", row.Status); Assert.AreEqual(0, pair.Source.Service.Calls);
        }
        [TestMethod]
        public void AppElementReferencesProveLocalEqualityAndPotentialSemanticDependencyOnly()
        {
            var pair = new Pair(); var left = pair.Source.Add(); var right = pair.Target.Add(); pair.Source.AddElement(left.Id); pair.Target.AddElement(right.Id);
            var report = pair.Capture(); Assert.AreEqual(left.Id, report.Source.Links.Single().CanvasId); Assert.AreEqual(right.Id, report.Target.Links.Single().CanvasId);
            StringAssert.Contains(report.Build(), "localReferenceEqualsBacking=True"); StringAssert.Contains(report.Build(), "SemanticPairObserved; potential future reuse only");
            Assert.AreEqual(3, pair.Source.Service.Calls); Assert.AreEqual(3, pair.Source.Service.ExecuteCalls);
            Assert.IsTrue(pair.Source.Queries.Where(q => q.EntityName == "appelement").All(q => q.ColumnSet.Columns.Count == 2));
        }
        [TestMethod]
        public void AmbiguousCanvasCandidateCannotCompleteAppElementDependency()
        {
            var pair = new Pair(); var left = pair.Source.Add(); pair.Source.Add(); pair.Target.Add(); pair.Source.AddElement(left.Id);
            var report = pair.Capture(); Assert.IsFalse(report.Build().Contains("SemanticPairObserved; potential future reuse only"));
        }
        [TestMethod]
        public void CancellationDuringIsolationPropagatesWithoutFurtherQueries()
        {
            var pair = new Pair(); pair.Source.Add(); Fail(pair.Source, "configuration");
            using (var cancel = new CancellationTokenSource())
            {
                var normal = pair.Source.Service.RetrievePage;
                pair.Source.Service.RetrievePage = q => { if (q.ColumnSet.Columns.Count == 2 && q.ColumnSet.Columns.Contains("configuration")) cancel.Cancel(); return normal(q); };
                Assert.ThrowsException<OperationCanceledException>(() => pair.Capture(cancel.Token)); Assert.AreEqual(0, pair.Target.Service.Calls + pair.Target.Service.ExecuteCalls);
            }
        }
        [TestMethod]
        public void EvidenceCaptureCannotMutateMembershipIdentityTotalsOrCoverage()
        {
            var pair = new Pair(); pair.Source.Add(); pair.Target.Add(); var left = pair.Source.Snapshot(); var right = pair.Target.Snapshot();
            var comparer = new SolutionMembershipComparer(); var before = comparer.Compare(left, right);
            new Type300EvidenceCollector().Capture(pair.Source.Service, left, "1", pair.Target.Service, right, "2", CancellationToken.None);
            CollectionAssert.AreEqual(before.Select(r => r.Presence).ToArray(), comparer.Compare(left, right).Select(r => r.Presence).ToArray());
            Assert.IsTrue(left.Components.Concat(right.Components).All(c => c.Status == IdentityResolutionStatus.Unsupported && c.ComparisonKey == null));
            Assert.AreEqual(0, pair.Source.Service.WriteCalls + pair.Target.Service.WriteCalls);
            Assert.AreEqual(4, pair.Source.Service.ExecuteCalls + pair.Target.Service.ExecuteCalls); // Schema and targeted option metadata only, no WhoAmI.
        }
        [TestMethod]
        public void ReportHasEveryRequiredSectionAndDeterministicSafeRequestLedger()
        {
            var pair = new Pair(); pair.Source.Add(); pair.Target.Add(); var report = pair.Capture(); var text = report.Build();
            foreach (var heading in new[] { "RAW TYPE 300 MEMBERSHIP", "BACKING CANVAS APP CORRELATION", "READABLE / UNAVAILABLE SCHEMA", "RUNTIME-READABLE / FAULTED COLUMNS", "PARENT / REFERENCE RELATIONSHIP DISCOVERY", "CANDIDATE IDENTITY ANALYSIS", "DUPLICATE / REPEATED ANALYSIS", "SOURCE / TARGET FIELD COMPARISON", "LIFECYCLE CORRELATION MATRIX", "DIFFERING PRIMARY-ID SEMANTIC PAIRS", "TYPE 10072 DEPENDENCY ASSESSMENT", "PORTABILITY ASSESSMENT", "EXACT REQUEST LEDGER" }) StringAssert.Contains(text, heading);
            Assert.AreEqual(text, report.Build()); StringAssert.Contains(text, "TotalReads=4"); StringAssert.Contains(text, "AdditionalWhoAmI=0");
            StringAssert.Contains(text, "Writes=0"); StringAssert.Contains(text, "NormalMembershipEvidenceRequests=0");
        }
        [TestMethod]
        public void FixedNameCandidatePairsDifferingIdsWithoutChangingCandidateAOrReads()
        {
            var pair = new Pair(); var left = pair.Source.Add(name: " publisher_Canvas "); var right = pair.Target.Add(name: "PUBLISHER_canvas"); right["ismanaged"] = true;
            var report = pair.Capture(); var proposed = report.ProposedPairs.Single();
            Assert.AreNotEqual(left.Id, right.Id); Assert.AreEqual("FixedFieldPairHypothesis", proposed.Outcome);
            Assert.IsTrue(proposed.Source.CompleteP && proposed.Target.CompleteP);
            Assert.IsTrue(StringComparer.OrdinalIgnoreCase.Equals(proposed.Source.CandidateP, proposed.Target.CandidateP));
            StringAssert.Contains(proposed.Source.CandidateP, "name"); Assert.AreEqual(2, pair.Source.Service.Calls);
            StringAssert.Contains(report.Build(), "CandidateP=EqualObserved"); StringAssert.Contains(report.Build(), "ManagedTransition=True");
            Assert.AreEqual("SemanticPair", report.Pairs.Single().Outcome);
        }
        [TestMethod]
        public void FixedNameMismatchNeverPairsThroughOtherIdentifierDisplayGuidOrHash()
        {
            var pair = new Pair(); var left = pair.Source.Add(name: "one"); pair.Target.Add(left.Id, "two");
            var report = pair.Capture(); Assert.IsTrue(report.ProposedPairs.All(p => p.Outcome == "OneSidedEvidence"));
        }
        [DataTestMethod, DataRow(""), DataRow(" "), DataRow("d956ce77-da6f-4cb0-937d-1e0b5958a780")]
        public void FixedNameBlankOrGuidRemainsIncompleteWithoutFallback(string name)
        {
            var pair = new Pair(); var row = pair.Source.Add(name: name); pair.Source.Attributes.Add(Attribute("uniquename", new StringAttributeMetadata())); row["uniquename"] = "independent_candidate_a";
            var evidence = pair.Capture().Source.Rows.Values.Single(); Assert.IsNotNull(evidence.CandidateA); Assert.IsNull(evidence.CandidateP);
            StringAssert.Contains(evidence.ProposedBlockingReason, "no fallback");
        }
        [DataTestMethod, DataRow(false), DataRow(true)]
        public void FixedNameUnavailableOrRuntimeFaultDoesNotSelectUniquename(bool runtimeFault)
        {
            var pair = new Pair(); var row = pair.Source.Add(); pair.Source.Attributes.Add(Attribute("uniquename", new StringAttributeMetadata())); row["uniquename"] = "portable_stronger";
            if (runtimeFault) Fail(pair.Source, "name"); else pair.Source.Attributes.RemoveAll(a => a.LogicalName == "name");
            var evidence = pair.Capture().Source.Rows.Values.Single(); Assert.IsNotNull(evidence.CandidateA); Assert.IsNull(evidence.CandidateP);
            StringAssert.Contains(evidence.ProposedBlockingReason, "Fixed identifier name unavailable/faulted");
        }
        [TestMethod]
        public void StrongerFieldMetadataDifferenceCannotChangeFixedContractOrRepairCandidateA()
        {
            var pair = new Pair(); var left = pair.Source.Add(); pair.Target.Add();
            pair.Source.Attributes.Add(Attribute("uniquename", new StringAttributeMetadata())); left["uniquename"] = "different_contract";
            var report = pair.Capture(); Assert.IsTrue(report.Pairs.All(p => p.Outcome == "OneSidedEvidence"));
            Assert.AreEqual("FixedFieldPairHypothesis", report.ProposedPairs.Single().Outcome);
            StringAssert.Contains(report.Build(), "CandidateASelectedFieldEquality=DifferentObserved");
            Assert.IsTrue(report.Source.Rows.Values.Single().CandidateP.Contains("name"));
        }
        [DataTestMethod, DataRow(1, true), DataRow(2, false)]
        public void ExplicitKindHypothesisUsesExactOptionsWhileDefaultRoleRemainsUnresolved(int targetKind, bool equal)
        {
            var pair = new Pair(); var left = pair.Source.Add(); var right = pair.Target.Add();
            foreach (var side in new[] { pair.Source, pair.Target }) side.Attributes.Add(Attribute("canvasapptype", new PicklistAttributeMetadata()));
            left["canvasapptype"] = new OptionSetValue(1); right["canvasapptype"] = new OptionSetValue(targetKind);
            var report = pair.Capture(); Assert.AreEqual("FixedFieldPairHypothesis", report.ProposedPairs.Single().Outcome);
            StringAssert.Contains(report.Source.Rows.Values.Single().KindEvidence["canvasapptype"], "classification=Unresolved");
            Type300EvidenceCollector.EvaluateCandidateP(report.Source.Rows.Values.Single(), report.Source.Snapshot, requiredKind: "canvasapptype");
            Type300EvidenceCollector.EvaluateCandidateP(report.Target.Rows.Values.Single(), report.Target.Snapshot, requiredKind: "canvasapptype");
            Assert.AreEqual(equal, StringComparer.OrdinalIgnoreCase.Equals(report.Source.Rows.Values.Single().CandidateP, report.Target.Rows.Values.Single().CandidateP));
        }
        [TestMethod]
        public void ExplicitRequiredKindUnavailableCannotBeInferred()
        {
            var pair = new Pair(); pair.Source.Add(); var report = pair.Capture(); var row = report.Source.Rows.Values.Single();
            Type300EvidenceCollector.EvaluateCandidateP(row, report.Source.Snapshot, requiredKind: "canvasappsubtype");
            Assert.IsNull(row.CandidateP); StringAssert.Contains(row.ProposedBlockingReason, "Required kind unavailable");
        }
        [DataTestMethod, DataRow(true), DataRow(false)]
        public void ExplicitRequiredAppModuleScopeUsesOnlyIndependentSnapshotIdentity(bool verified)
        {
            var pair = new Pair();
            foreach (var side in new[] { pair.Source, pair.Target })
            {
                var row = side.Add(); var parent = Guid.NewGuid(); side.Attributes.Add(Attribute("appmoduleid", new LookupAttributeMetadata { Targets = new[] { "appmodule" } }));
                row["appmoduleid"] = new EntityReference("appmodule", parent);
                if (verified) side.Raw.Add(new ComponentIdentity(new SolutionComponentRecord(Guid.NewGuid(), 80, parent), IdentityResolutionStatus.Resolved, "publisher_app"));
            }
            var report = pair.Capture();
            foreach (var side in new[] { report.Source, report.Target })
            {
                var row = side.Rows.Values.Single(); Type300EvidenceCollector.EvaluateCandidateP(row, side.Snapshot, requiredScope: "appmoduleid");
                Assert.AreEqual(verified, row.CandidateP != null);
                if (!verified) StringAssert.Contains(row.ProposedBlockingReason, "Required portable AppModule scope unresolved");
            }
            if (verified) Assert.AreEqual(report.Source.Rows.Values.Single().CandidateP, report.Target.Rows.Values.Single().CandidateP);
        }
        [TestMethod]
        public void DefaultHypothesisDoesNotInventParentRequirementFromOwnerOrAppElement()
        {
            var pair = new Pair(); var left = pair.Source.Add(); pair.Target.Add(); left["ownerid"] = new EntityReference("systemuser", Guid.NewGuid()); pair.Source.AddElement(left.Id);
            var report = pair.Capture(); Assert.IsTrue(report.Source.Rows.Values.Single().CompleteP);
            StringAssert.Contains(report.Build(), "No structural scope has yet been independently proven required");
            Assert.IsFalse(report.Source.Rows.Values.Single().CandidateP.Contains("systemuser"));
            StringAssert.Contains(report.Build(), "SameSelectedBackingPair; no independent identifier proof or repair");
        }
        [TestMethod]
        public void CandidatePCollisionIsSeparateFromRepeatedMembershipAndCannotUseCandidateARepair()
        {
            var pair = new Pair(); var one = pair.Source.Add(); var two = pair.Source.Add(); pair.Target.Add();
            pair.Source.Reference(300, one.Id);
            pair.Source.Attributes.Add(Attribute("uniquename", new StringAttributeMetadata())); one["uniquename"] = "one"; two["uniquename"] = "two";
            var report = pair.Capture(); Assert.AreEqual(2, report.Source.Rows.Count); Assert.AreEqual(3, report.Source.Raw.Count);
            Assert.IsTrue(report.Source.Rows.Values.All(r => r.DuplicateP && !r.DuplicateA && !r.CompleteP));
            Assert.IsTrue(report.ProposedPairs.All(p => p.Outcome == "Ambiguous")); StringAssert.Contains(report.Build(), "CandidatePCollisionGroups=1");
        }
        [TestMethod]
        public void RepeatedMembershipDoesNotCreateCandidatePCollision()
        {
            var pair = new Pair(); var row = pair.Source.Add(); pair.Source.Reference(300, row.Id); pair.Target.Add();
            var report = pair.Capture(); Assert.AreEqual(1, report.ProposedPairs.Count); Assert.IsTrue(report.Source.Rows.Values.Single().CompleteP);
            Assert.AreEqual(2, pair.Source.Service.Calls);
        }
        [DataTestMethod, DataRow("ishidden"), DataRow("hascustommiddletier")]
        public void BehavioralDefinitionDifferenceDoesNotChangeEitherCandidate(string field)
        {
            var pair = new Pair(); foreach (var side in new[] { pair.Source, pair.Target }) { side.Attributes.Add(Attribute(field, new BooleanAttributeMetadata())); side.Add()[field] = side == pair.Target; }
            var report = pair.Capture(); Assert.AreEqual("FixedFieldPairHypothesis", report.ProposedPairs.Single().Outcome);
            Assert.AreEqual("False", report.Source.Rows.Values.Single().DefinitionValues[field]); Assert.AreEqual("True", report.Target.Rows.Values.Single().DefinitionValues[field]);
            StringAssert.Contains(report.Build(), "INITIAL SCALAR DEFINITION EVIDENCE"); StringAssert.Contains(report.Build(), "Comparison=DifferentObserved");
            Assert.AreEqual("SemanticPair", report.Pairs.Single().Outcome);
        }
        [TestMethod]
        public void OptionalBehaviorFaultDoesNotInvalidateFixedIdentifier()
        {
            var pair = new Pair(); pair.Source.Add(); pair.Source.Attributes.Add(Attribute("ishidden", new BooleanAttributeMetadata())); Fail(pair.Source, "ishidden");
            var row = pair.Capture().Source.Rows.Values.Single(); Assert.IsTrue(row.CompleteP); Assert.AreEqual("Unavailable", row.DefinitionValues["ishidden"]);
        }
        [TestMethod]
        public void AppElementDependencyAndCandidateACannotRepairFixedFieldUnavailable()
        {
            var pair = new Pair(); var left = pair.Source.Add(); var right = pair.Target.Add();
            foreach (var side in new[] { pair.Source, pair.Target }) { side.Attributes.Add(Attribute("uniquename", new StringAttributeMetadata())); side.Data.Single()["uniquename"] = "stronger"; side.Attributes.RemoveAll(a => a.LogicalName == "name"); }
            pair.Source.AddElement(left.Id); pair.Target.AddElement(right.Id); var report = pair.Capture();
            Assert.AreEqual("SemanticPair", report.Pairs.Single().Outcome); Assert.IsTrue(report.Source.Rows.Values.All(r => r.CandidateP == null));
            StringAssert.Contains(report.Build(), "dependency cannot repair P"); Assert.AreEqual(3, pair.Source.Service.Calls); // Critical, optional and scoped AppElement batch.
        }
        [TestMethod]
        public void AuthenticatedOptionMetadataCapturesReturnedLabelsAndProvenanceWithOneRequestPerPopulatedSide()
        {
            var pair = new Pair(); pair.Source.Add(); pair.Target.Add(); var report = pair.Capture();
            foreach (var side in new[] { report.Source, report.Target })
            {
                Assert.IsTrue(side.OptionMetadata.Complete); Assert.AreEqual(1, side.OptionMetadata.Attempted);
                Assert.AreEqual(1, side.OptionMetadata.Succeeded); Assert.AreEqual(0, side.OptionMetadata.Failed);
                Assert.AreEqual("App Component Library", side.OptionMetadata.Labels[1033]);
                Assert.IsNotNull(side.OptionMetadata.TimestampUtc); Assert.IsNull(side.OptionMetadata.Endpoint);
                Assert.AreEqual(2, side.Requests.Count(r => r.StartsWith("Execute")));
                Assert.AreEqual(2, side.Requests.Count(r => r.StartsWith("RetrieveMultiple")));
            }
            Assert.AreEqual(1, pair.Source.OptionRequests.Count); Assert.AreEqual(1, pair.Target.OptionRequests.Count);
            var text = report.Build(); StringAssert.Contains(text, "OptionSet.Name=canvasapp_canvasapptype");
            StringAssert.Contains(text, "OptionSet.IsGlobal=True"); StringAssert.Contains(text, "OptionSet.OptionSetType=Picklist");
            StringAssert.Contains(text, "Attribute.MetadataId="); StringAssert.Contains(text, "UserLocalizedLabel.LanguageCode=1033");
            StringAssert.Contains(text, "LocalizedLabel: LanguageCode=1033"); StringAssert.Contains(text, "OptionEvidenceGate=Complete");
            StringAssert.Contains(text, "StructuralScope=Unresolved"); StringAssert.Contains(text, "ConfirmedEndpoint=Unavailable");
            Assert.AreEqual(0, pair.Source.Service.WriteCalls + pair.Target.Service.WriteCalls);
        }
        [TestMethod]
        public void EnvironmentSpecificMetadataIdsAreAuditOnly()
        {
            var pair = new Pair(); pair.Source.Add(); pair.Target.Add(); var left = ValidOptionMetadata(); var right = ValidOptionMetadata();
            Assert.AreNotEqual(left.MetadataId, right.MetadataId); Assert.AreNotEqual(left.OptionSet.MetadataId, right.OptionSet.MetadataId);
            pair.Source.OptionResponse = r => OptionMetadataResponse(left); pair.Target.OptionResponse = r => OptionMetadataResponse(right);
            StringAssert.Contains(pair.Capture().Build(), "Comparison=EqualObservedForSharedLanguages");
        }
        [DataTestMethod]
        [DataRow("NullAttribute")]
        [DataRow("WrongType")]
        [DataRow("WrongName")]
        [DataRow("Unreadable")]
        [DataRow("NoOptionSet")]
        [DataRow("NoValue1")]
        [DataRow("DuplicateValue1")]
        [DataRow("NoLabels")]
        [DataRow("ConflictingLabels")]
        [DataRow("IncompleteProvenance")]
        public void IncompleteOptionMetadataNeverRepairsOrInvalidatesIdentity(string problem)
        {
            var pair = new Pair(); pair.Source.Add(); pair.Target.Add(); AttributeMetadata attribute = ValidOptionMetadata();
            var picklist = (PicklistAttributeMetadata)attribute;
            switch (problem)
            {
                case "NullAttribute": attribute = null; break;
                case "WrongType": attribute = Attribute("canvasapptype", new StringAttributeMetadata()); break;
                case "WrongName": attribute.LogicalName = "different"; break;
                case "Unreadable": Set(attribute, "IsValidForRead", false); break;
                case "NoOptionSet": picklist.OptionSet = null; break;
                case "NoValue1": picklist.OptionSet.Options.Clear(); break;
                case "DuplicateValue1": picklist.OptionSet.Options.Add(new OptionMetadata(new Label("duplicate", 1033), 1)); break;
                case "NoLabels": picklist.OptionSet.Options[0].Label = new Label(); break;
                case "ConflictingLabels": picklist.OptionSet.Options[0].Label.LocalizedLabels.Add(new LocalizedLabel("conflicting", 1033)); break;
                case "IncompleteProvenance": picklist.OptionSet.Name = null; break;
            }
            pair.Source.OptionResponse = r => OptionMetadataResponse(attribute); var report = pair.Capture();
            Assert.IsFalse(report.Source.OptionMetadata.Complete); Assert.IsTrue(report.Target.OptionMetadata.Complete);
            Assert.AreEqual(1, report.Source.OptionMetadata.Succeeded); Assert.AreEqual(0, report.Source.OptionMetadata.Failed);
            Assert.IsTrue(report.Source.Rows.Values.Single().CompleteP); Assert.AreEqual("Unique", report.Source.Rows.Values.Single().Status);
            Assert.AreEqual("FixedFieldPairHypothesis", report.ProposedPairs.Single().Outcome);
            StringAssert.Contains(report.Build(), "OptionEvidenceGate=Unresolved");
        }
        [TestMethod]
        public void MissingTranslationDoesNotConflictAndDoesNotFabricatePublishedLabel()
        {
            var pair = new Pair(); pair.Source.Add(); pair.Target.Add(); var metadata = ValidOptionMetadata();
            metadata.OptionSet.Options[0].Label.LocalizedLabels.Add(new LocalizedLabel("Bibliotheque", 1036));
            pair.Source.OptionResponse = r => OptionMetadataResponse(metadata); var text = pair.Capture().Build();
            StringAssert.Contains(text, "Comparison=EqualObservedForSharedLanguages");
            StringAssert.Contains(text, "SourceOnlyTranslation=1036"); StringAssert.Contains(text, "Not a conflicting option definition");
        }
        [TestMethod]
        public void DifferentLabelsForSameLanguageKeepCombinedGateUnresolved()
        {
            var pair = new Pair(); pair.Source.Add(); pair.Target.Add(); var metadata = ValidOptionMetadata();
            metadata.OptionSet.Options[0].Label = new Label("Different live definition", 1033);
            pair.Target.OptionResponse = r => OptionMetadataResponse(metadata); var text = pair.Capture().Build();
            StringAssert.Contains(text, "Comparison=DifferentObserved"); StringAssert.Contains(text, "OptionEvidenceGate=Unresolved");
        }
        [DataTestMethod]
        [DataRow(true, false)]
        [DataRow(false, true)]
        [DataRow(true, true)]
        public void OptionalMetadataFaultOnEitherSideNeverChangesSnapshotOrCandidateEvidence(bool sourceFault, bool targetFault)
        {
            var pair = new Pair(); pair.Source.Add(); pair.Target.Add();
            Func<RetrieveAttributeRequest, OrganizationResponse> fault = r => { throw new FaultException<OrganizationServiceFault>(
                new OrganizationServiceFault { ErrorCode = unchecked((int)0x8004023B), Message = Secret + " https://private.invalid/?token=secret", TraceText = "PRIVATE-STACK" }, new FaultReason(Secret)); };
            if (sourceFault) pair.Source.OptionResponse = fault; if (targetFault) pair.Target.OptionResponse = fault;
            var source = pair.Source.Snapshot(); var target = pair.Target.Snapshot(); var before = new SolutionMembershipComparer().Compare(source, target).Select(r => r.Presence).ToArray();
            var report = new Type300EvidenceCollector().Capture(pair.Source.Service, source, "1", pair.Target.Service, target, "2", CancellationToken.None);
            CollectionAssert.AreEqual(before, new SolutionMembershipComparer().Compare(source, target).Select(r => r.Presence).ToArray());
            Assert.AreEqual(sourceFault ? 1 : 0, report.Source.OptionMetadata.Failed); Assert.AreEqual(targetFault ? 1 : 0, report.Target.OptionMetadata.Failed);
            Assert.IsTrue(report.Source.Rows.Values.Single().CompleteP); Assert.IsTrue(report.Target.Rows.Values.Single().CompleteP);
            var text = report.Build(); StringAssert.Contains(text, "8004023B");
            Assert.IsFalse(text.Contains(Secret)); Assert.IsFalse(text.Contains("private.invalid")); Assert.IsFalse(text.Contains("PRIVATE-STACK"));
            Assert.AreEqual(1, pair.Source.OptionRequests.Count); Assert.AreEqual(1, pair.Target.OptionRequests.Count);
        }
        [TestMethod]
        public void ZeroMemberSideSkipsTargetedOptionMetadataWithoutDispatch()
        {
            var pair = new Pair(); pair.Source.Add(); var report = pair.Capture();
            Assert.AreEqual(1, report.Target.OptionMetadata.Skipped); Assert.AreEqual(0, report.Target.OptionMetadata.Attempted);
            Assert.AreEqual(0, pair.Target.OptionRequests.Count); Assert.AreEqual(0, pair.Target.Service.ExecuteCalls + pair.Target.Service.Calls);
        }
        [TestMethod]
        public void CancellationAfterMetadataDispatchDoesNotStartTargetMetadata()
        {
            var pair = new Pair(); pair.Source.Add(); pair.Target.Add();
            using (var cancellation = new CancellationTokenSource())
            {
                pair.Source.OptionResponse = r => { cancellation.Cancel(); return OptionMetadataResponse(ValidOptionMetadata()); };
                Assert.ThrowsException<OperationCanceledException>(() => pair.Capture(cancellation.Token));
                Assert.AreEqual(1, pair.Source.OptionRequests.Count); Assert.AreEqual(0, pair.Target.OptionRequests.Count);
            }
        }
        [DataTestMethod]
        [DataRow("Blank", "EmptyString", "0", "True")]
        [DataRow("Populated", "Populated", "7", "True")]
        [DataRow("Null", "Null", "Unknown", "True")]
        [DataRow("Omitted", "AttributeOmitted", "Unknown", "True")]
        [DataRow("Unavailable", "Unavailable", "Unknown", "False")]
        [DataRow("Faulted", "RetrievalFailed", "Unknown", "False")]
        public void StringSubtypeReportsRetrievalAndPresenceSeparatelyFromNumericApplicability(string scenario, string presence, string length, string runtimeRead)
        {
            var pair = new Pair(); var row = pair.Source.Add();
            if (scenario != "Unavailable") pair.Source.Attributes.Add(Attribute("canvasappsubtype", new StringAttributeMetadata()));
            if (scenario == "Blank") row["canvasappsubtype"] = "";
            if (scenario == "Null") row["canvasappsubtype"] = null;
            if (scenario == "Populated") row["canvasappsubtype"] = "Subtype";
            if (scenario == "Faulted") Fail(pair.Source, "canvasappsubtype");
            var evidence = pair.Capture().Source.Rows.Values.Single(); var text = evidence.KindEvidence["canvasappsubtype"];
            StringAssert.Contains(text, "runtimeRead=" + runtimeRead); StringAssert.Contains(text, "presence=" + presence);
            StringAssert.Contains(text, "length=" + length);
            StringAssert.Contains(text, scenario == "Unavailable" ? "numericClassification=Unknown" : "numericClassification=NotApplicable");
            if (scenario != "Unavailable") StringAssert.Contains(text, "exactNumericValue=NotApplicable");
            StringAssert.Contains(text, "eligibility=NotEstablished");
            StringAssert.Contains(evidence.Evidence("canvasappsubtype"), "presence=" + presence);
            Assert.IsTrue(evidence.CompleteP);
        }
        [DataTestMethod]
        [DataRow("EmptyString", "EmptyString", true)]
        [DataRow("Null", "Null", true)]
        [DataRow("AttributeOmitted", "AttributeOmitted", true)]
        [DataRow("EmptyString", "Null", false)]
        [DataRow("EmptyString", "AttributeOmitted", false)]
        [DataRow("Null", "AttributeOmitted", false)]
        [DataRow("Populated", "Populated", true)]
        public void SubtypeComparisonPreservesSdkRepresentationEvenWhenNormalizedHashesAgree(string sourceState, string targetState, bool equal)
        {
            var pair = new Pair(); var left = pair.Source.Add(); var right = pair.Target.Add();
            foreach (var side in new[] { pair.Source, pair.Target }) side.Attributes.Add(Attribute("canvasappsubtype", new StringAttributeMetadata()));
            SetSubtype(left, sourceState); SetSubtype(right, targetState); var report = pair.Capture();
            var source = report.Source.Rows.Values.Single(); var target = report.Target.Rows.Values.Single();
            StringAssert.Contains(source.KindEvidence["canvasappsubtype"], "presence=" + sourceState);
            StringAssert.Contains(target.KindEvidence["canvasappsubtype"], "presence=" + targetState);
            var comparison = report.Build().Split(new[] { "SOURCE / TARGET FIELD COMPARISON" }, StringSplitOptions.None).Last()
                .Split('\n').Single(line => line.StartsWith("canvasappsubtype\tSource=", StringComparison.Ordinal));
            StringAssert.Contains(comparison, "Source=presence=" + sourceState); StringAssert.Contains(comparison, "Target=presence=" + targetState);
            StringAssert.Contains(comparison, equal ? "EqualObserved" : "DifferentObserved");
            Assert.AreEqual(source.Content["canvasappsubtype"].Sha256, target.Content["canvasappsubtype"].Sha256);
            Assert.AreEqual(sourceState != "AttributeOmitted", source.StringValues["canvasappsubtype"].AttributeCollectionContains);
            Assert.IsTrue(source.StringValues["canvasappsubtype"].QueryColumnCompleted);
            Assert.IsTrue(source.CompleteP && target.CompleteP); StringAssert.Contains(comparison, "hash does not prove representation or eligibility");
        }
        private static void SetSubtype(Entity row, string state)
        {
            if (state == "EmptyString") row["canvasappsubtype"] = "";
            else if (state == "Null") row["canvasappsubtype"] = null;
            else if (state == "Populated") row["canvasappsubtype"] = "Subtype";
        }
        [TestMethod]
        public void UserLabelIsSeparateProvenanceAndDoesNotRenderAsSecondLocalizedCollectionEntry()
        {
            var pair = new Pair(); pair.Source.Add(); var metadata = ValidOptionMetadata();
            metadata.OptionSet.Options[0].Label.UserLocalizedLabel = metadata.OptionSet.Options[0].Label.LocalizedLabels.Single();
            pair.Source.OptionResponse = r => OptionMetadataResponse(metadata); var report = pair.Capture(); var fields = report.Source.OptionMetadata.Fields;
            Assert.IsTrue(report.Source.OptionMetadata.Complete); Assert.AreEqual(1, fields.Count(f => f.StartsWith("LocalizedLabel:")));
            CollectionAssert.Contains(fields, "LocalizedLabels.RawEntryCount=1"); CollectionAssert.Contains(fields, "UserLocalizedLabelAlsoInLocalizedLabels=True");
        }
        [TestMethod]
        public void ExactSdkLocalizedDuplicatesRetainRawMultiplicityAndMetadataIdsWithOneDisplayEntry()
        {
            var pair = new Pair(); pair.Source.Add(); var metadata = ValidOptionMetadata(); var labels = metadata.OptionSet.Options[0].Label;
            labels.LocalizedLabels.Add(new LocalizedLabel(labels.LocalizedLabels[0].Label, 1033) { MetadataId = Guid.NewGuid() });
            pair.Source.OptionResponse = r => OptionMetadataResponse(metadata); var report = pair.Capture(); var fields = report.Source.OptionMetadata.Fields;
            Assert.IsTrue(report.Source.OptionMetadata.Complete); CollectionAssert.Contains(fields, "LocalizedLabels.RawEntryCount=2");
            var display = fields.Single(f => f.StartsWith("LocalizedLabel:")); StringAssert.Contains(display, "RawCollectionOccurrences=2");
            StringAssert.Contains(display, labels.LocalizedLabels[1].MetadataId.ToString());
            CollectionAssert.Contains(fields, "Value1OptionCount=1");
        }
        [TestMethod]
        public void MultipleLanguagesArePresentedSeparatelyWithoutUserLabelDuplication()
        {
            var pair = new Pair(); pair.Source.Add(); var metadata = ValidOptionMetadata();
            metadata.OptionSet.Options[0].Label.LocalizedLabels.Add(new LocalizedLabel("Bibliotheque", 1036));
            pair.Source.OptionResponse = r => OptionMetadataResponse(metadata); var option = pair.Capture().Source.OptionMetadata;
            Assert.IsTrue(option.Complete); Assert.AreEqual(2, option.Labels.Count);
            Assert.AreEqual(2, option.Fields.Count(f => f.StartsWith("LocalizedLabel:")));
        }
        [TestMethod]
        public void ConflictingLabelsForSameLanguageAreRetainedAndDoNotBecomeDuplicateOptions()
        {
            var pair = new Pair(); pair.Source.Add(); var metadata = ValidOptionMetadata();
            metadata.OptionSet.Options[0].Label.LocalizedLabels.Add(new LocalizedLabel("conflicting", 1033));
            pair.Source.OptionResponse = r => OptionMetadataResponse(metadata); var option = pair.Capture().Source.OptionMetadata;
            Assert.IsFalse(option.Complete); Assert.AreEqual(2, option.Fields.Count(f => f.StartsWith("LocalizedLabel:")));
            StringAssert.Contains(option.Limitation, "conflicting"); CollectionAssert.Contains(option.Fields, "Value1OptionCount=1");
        }
        [TestMethod]
        public void MissingUserLabelDoesNotInvalidateLocalizedCollectionProvenance()
        {
            var pair = new Pair(); pair.Source.Add(); var metadata = ValidOptionMetadata(); metadata.OptionSet.Options[0].Label.UserLocalizedLabel = null;
            pair.Source.OptionResponse = r => OptionMetadataResponse(metadata); var option = pair.Capture().Source.OptionMetadata;
            Assert.IsTrue(option.Complete); Assert.AreEqual(1, option.Labels.Count); Assert.AreEqual(1, option.Fields.Count(f => f.StartsWith("LocalizedLabel:")));
        }
        [TestMethod]
        public void MissingLocalizedCollectionEntriesPreserveUserLabelAsSeparateProvenance()
        {
            var pair = new Pair(); pair.Source.Add(); var metadata = ValidOptionMetadata(); metadata.OptionSet.Options[0].Label.LocalizedLabels.Clear();
            pair.Source.OptionResponse = r => OptionMetadataResponse(metadata); var option = pair.Capture().Source.OptionMetadata;
            Assert.IsTrue(option.Complete); Assert.AreEqual(1, option.Labels.Count); Assert.AreEqual(0, option.Fields.Count(f => f.StartsWith("LocalizedLabel:")));
            CollectionAssert.Contains(option.Fields, "LocalizedLabels.RawEntryCount=0");
        }
        [TestMethod]
        public void ResolvedIndeterminateRowsAreRetainedEvenThoughExistingSummaryHasNoCounterForThem()
        {
            Func<ComponentIdentity> resolved = () => new ComponentIdentity(new SolutionComponentRecord(Guid.NewGuid(), 60, Guid.NewGuid()),
                IdentityResolutionStatus.Resolved, "systemform:" + Guid.NewGuid(), componentTypeKey: ComponentSemanticKinds.SystemForm,
                semanticKind: ComponentSemanticKinds.SystemForm, inventoryAbsencePolicy: InventoryAbsencePolicy.MatchOnly);
            var source = Snapshot(resolved()); var target = Snapshot(resolved(), resolved());
            var presentation = new MembershipResultPresenter().Create(MembershipEnvironmentResult.FromSnapshot("Source", source, 0, TimeSpan.Zero),
                MembershipEnvironmentResult.FromSnapshot("Target", target, 0, TimeSpan.Zero));
            Assert.AreEqual(1, presentation.Source.Diagnostics.ResolvedCount); Assert.AreEqual(2, presentation.Target.Diagnostics.ResolvedCount);
            Assert.AreEqual(3, presentation.Rows.Count); Assert.IsTrue(presentation.Rows.All(r => r.Presence == MembershipPresence.Indeterminate));
            Assert.AreEqual(0, presentation.Summary.PresentInBoth + presentation.Summary.SourceOnly + presentation.Summary.TargetOnly +
                presentation.Summary.Unsupported + presentation.Summary.Unresolved + presentation.Summary.Ambiguous);
            Assert.IsTrue(presentation.Rows.All(r => r.SourceResolutionStatus == "Resolved" || r.TargetResolutionStatus == "Resolved"));
        }
        private static PicklistAttributeMetadata ValidOptionMetadata()
        {
            var option = new OptionMetadata(new Label("App Component Library", 1033), 1);
            option.Label.UserLocalizedLabel = new LocalizedLabel("App Component Library", 1033) { MetadataId = Guid.NewGuid() };
            var set = new OptionSetMetadata { Name = "canvasapp_canvasapptype", MetadataId = Guid.NewGuid(), IsGlobal = true, OptionSetType = OptionSetType.Picklist };
            set.Options.Add(option);
            return (PicklistAttributeMetadata)Attribute("canvasapptype", new PicklistAttributeMetadata { MetadataId = Guid.NewGuid(), OptionSet = set });
        }
        private static RetrieveAttributeResponse OptionMetadataResponse(AttributeMetadata metadata)
        { var response = new RetrieveAttributeResponse(); response.Results["AttributeMetadata"] = metadata; return response; }
        private static void Fail(Fixture side, string column)
        {
            var normal = side.Service.RetrievePage; side.Service.RetrievePage = q => { var rows = normal(q); if (q.ColumnSet.Columns.Contains(column))
                throw new FaultException<OrganizationServiceFault>(new OrganizationServiceFault { ErrorCode = unchecked((int)0x8004023B), Message = Secret + " Bearer token", TraceText = "STACK-TRACE" }, new FaultReason(Secret)); return rows; };
        }
        private static AttributeMetadata Attribute(string name, AttributeMetadata attribute)
        { attribute.LogicalName = name; Set(attribute, "IsValidForRead", true); return attribute; }
        private static void Set(object obj, string name, object value) => obj.GetType().GetProperty(name).SetValue(obj, value, null);
        private sealed class Pair
        {
            internal readonly Fixture Source = new Fixture(), Target = new Fixture();
            internal CanvasAppEvidenceReport Capture(CancellationToken token = default(CancellationToken)) => new Type300EvidenceCollector().Capture(Source.Service, Source.Snapshot(), "1", Target.Service, Target.Snapshot(), "2", token);
        }
        private sealed class Fixture
        {
            internal readonly FakeOrganizationService Service;
            internal readonly List<Entity> Data = new List<Entity>();
            internal readonly List<ComponentIdentity> Raw = new List<ComponentIdentity>();
            internal readonly List<QueryExpression> Queries = new List<QueryExpression>();
            internal readonly List<AttributeMetadata> Attributes = new List<AttributeMetadata>();
            internal readonly List<RetrieveAttributeRequest> OptionRequests = new List<RetrieveAttributeRequest>();
            internal Func<RetrieveAttributeRequest, OrganizationResponse> OptionResponse = request => OptionMetadataResponse(ValidOptionMetadata());
            internal Fixture()
            {
                Attributes.Add(Attribute("canvasappid", new UniqueIdentifierAttributeMetadata()));
                foreach (var field in new[] { "name", "displayname", "uniquecanvasappid", "version", "appopenuri" }) Attributes.Add(Attribute(field, new StringAttributeMetadata()));
                Attributes.Add(Attribute("appcomponenttype", new PicklistAttributeMetadata())); Attributes.Add(Attribute("ismanaged", new BooleanAttributeMetadata()));
                Attributes.Add(Attribute("componentstate", new PicklistAttributeMetadata())); Attributes.Add(Attribute("configuration", new MemoAttributeMetadata()));
                Attributes.Add(Attribute("ownerid", new LookupAttributeMetadata { Targets = new[] { "systemuser", "team" } }));
                Service = new FakeOrganizationService {
                    ExecuteRequest = request => {
                        var optionRequest = request as RetrieveAttributeRequest;
                        if (optionRequest != null)
                        {
                            OptionRequests.Add(optionRequest); Assert.AreEqual("canvasapp", optionRequest.EntityLogicalName);
                            Assert.AreEqual("canvasapptype", optionRequest.LogicalName); Assert.IsFalse(optionRequest.RetrieveAsIfPublished);
                            return OptionResponse(optionRequest);
                        }
                        Assert.IsInstanceOfType(request, typeof(RetrieveEntityRequest)); var req = (RetrieveEntityRequest)request;
                        Assert.AreEqual(EntityFilters.Attributes | EntityFilters.Relationships, req.EntityFilters); Assert.IsFalse(req.RetrieveAsIfPublished);
                        Assert.IsTrue(new[] { "canvasapp", "appelement" }.Contains(req.LogicalName));
                        var metadata = new EntityMetadata { LogicalName = req.LogicalName }; Set(metadata, "PrimaryIdAttribute", req.LogicalName + "id");
                        Set(metadata, "Attributes", req.LogicalName == "canvasapp" ? Attributes.ToArray() : new[] {
                            Attribute("appelementid", new UniqueIdentifierAttributeMetadata()), Attribute("canvasappid", new LookupAttributeMetadata { Targets = new[] { "canvasapp" } }) });
                        var response = new RetrieveEntityResponse(); response.Results["EntityMetadata"] = metadata; return response;
                    },
                    RetrievePage = q => {
                        Queries.Add(q); var condition = q.Criteria.Conditions.Single(); Assert.AreEqual(ConditionOperator.In, condition.Operator);
                        Assert.AreEqual(q.EntityName + "id", condition.AttributeName); Assert.IsTrue(condition.Values.Count <= 200);
                        var ids = condition.Values.Cast<Guid>().ToArray();
                        return new EntityCollection(Data.Where(r => r.LogicalName == q.EntityName && ids.Contains(r.Id)).Select(r => {
                            var copy = new Entity(r.LogicalName, r.Id); foreach (var column in q.ColumnSet.Columns) if (r.Contains(column)) copy[column] = r[column]; return copy;
                        }).ToList());
                    }
                };
            }
            internal Entity Add(Guid? id = null, string name = "publisher_canvas")
            {
                var row = new Entity("canvasapp", id ?? Guid.NewGuid()) { ["name"] = name, ["displayname"] = "Canvas Display", ["uniquecanvasappid"] = Guid.NewGuid().ToString("D"),
                    ["version"] = "1", ["appcomponenttype"] = new OptionSetValue(1), ["ismanaged"] = false, ["componentstate"] = new OptionSetValue(0), ["configuration"] = Secret, ["appopenuri"] = Secret };
                row["canvasappid"] = row.Id; Data.Add(row); Reference(300, row.Id); return row;
            }
            internal void AddElement(Guid canvasId)
            { var row = new Entity("appelement", Guid.NewGuid()) { ["canvasappid"] = new EntityReference("canvasapp", canvasId) }; row["appelementid"] = row.Id; Data.Add(row); Reference(10072, row.Id); }
            internal void Reference(int type, Guid? id) => Raw.Add(new ComponentIdentity(new SolutionComponentRecord(Guid.NewGuid(), type, id), IdentityResolutionStatus.Unsupported));
            internal MembershipSnapshot Snapshot() => MembershipSnapshot.Complete(Solution(), Raw, DateTimeOffset.UtcNow);
        }
#endif
    }
}
