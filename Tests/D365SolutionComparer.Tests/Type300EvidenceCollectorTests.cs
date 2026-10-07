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
            Assert.AreEqual(2, pair.Source.Service.Calls); Assert.AreEqual(1, pair.Source.Service.ExecuteCalls);
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
            Assert.AreEqual(count, report.Source.Rows.Count); Assert.AreEqual(1, pair.Source.Service.ExecuteCalls);
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
            Assert.AreEqual(3, pair.Source.Service.Calls); Assert.AreEqual(2, pair.Source.Service.ExecuteCalls);
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
            Assert.AreEqual(2, pair.Source.Service.ExecuteCalls + pair.Target.Service.ExecuteCalls); // Schema only, no WhoAmI.
        }
        [TestMethod]
        public void ReportHasEveryRequiredSectionAndDeterministicSafeRequestLedger()
        {
            var pair = new Pair(); pair.Source.Add(); pair.Target.Add(); var report = pair.Capture(); var text = report.Build();
            foreach (var heading in new[] { "RAW TYPE 300 MEMBERSHIP", "BACKING CANVAS APP CORRELATION", "READABLE / UNAVAILABLE SCHEMA", "RUNTIME-READABLE / FAULTED COLUMNS", "PARENT / REFERENCE RELATIONSHIP DISCOVERY", "CANDIDATE IDENTITY ANALYSIS", "DUPLICATE / REPEATED ANALYSIS", "SOURCE / TARGET FIELD COMPARISON", "LIFECYCLE CORRELATION MATRIX", "DIFFERING PRIMARY-ID SEMANTIC PAIRS", "TYPE 10072 DEPENDENCY ASSESSMENT", "PORTABILITY ASSESSMENT", "EXACT REQUEST LEDGER" }) StringAssert.Contains(text, heading);
            Assert.AreEqual(text, report.Build()); StringAssert.Contains(text, "TotalReads=3"); StringAssert.Contains(text, "AdditionalWhoAmI=0");
            StringAssert.Contains(text, "Writes=0"); StringAssert.Contains(text, "NormalMembershipEvidenceRequests=0");
        }
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
            internal Fixture()
            {
                Attributes.Add(Attribute("canvasappid", new UniqueIdentifierAttributeMetadata()));
                foreach (var field in new[] { "name", "displayname", "uniquecanvasappid", "version", "appopenuri" }) Attributes.Add(Attribute(field, new StringAttributeMetadata()));
                Attributes.Add(Attribute("appcomponenttype", new PicklistAttributeMetadata())); Attributes.Add(Attribute("ismanaged", new BooleanAttributeMetadata()));
                Attributes.Add(Attribute("componentstate", new PicklistAttributeMetadata())); Attributes.Add(Attribute("configuration", new MemoAttributeMetadata()));
                Attributes.Add(Attribute("ownerid", new LookupAttributeMetadata { Targets = new[] { "systemuser", "team" } }));
                Service = new FakeOrganizationService {
                    ExecuteRequest = request => {
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
