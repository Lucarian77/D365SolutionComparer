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
    [TestClass, TestCategory("Phase2GAppActionEvidence")]
    public class AppActionEvidenceCollectorTests
    {
        [DataTestMethod, DataRow(10266), DataRow(10267)]
        public void DefaultResolverRemainsUnsupportedWithoutKeyContractOrAbsence(int type)
        {
            var solution = Solution();
            var service = Service(solution, q => q.EntityName == "solutioncomponentdefinition" &&
                q.Criteria.Conditions.Any(c => c.AttributeName == "objecttypecode" && c.Values.Contains(type))
                ? Rows(new Entity("solutioncomponentdefinition", Guid.NewGuid()) {
                    ["objecttypecode"] = type, ["name"] = "AppAction", ["primaryentityname"] = "appaction" }) : Rows());
            var identity = new DataverseComponentIdentityResolver().Resolve(service, solution.Environment,
                new SolutionComponentRecord(Guid.NewGuid(), type, Guid.NewGuid()), CancellationToken.None);
            Assert.AreEqual(IdentityResolutionStatus.Unsupported, identity.Status);
            Assert.IsNull(identity.ComparisonKey);
            Assert.IsNull(D365SolutionComparer.Models.ComponentDetails.ComponentDefinitionContractCatalog.For(identity.SemanticKind));
            Assert.IsTrue(new SolutionMembershipComparer().Compare(Snapshot(identity), Snapshot()).All(r => r.Presence == MembershipPresence.Indeterminate));
            Assert.AreEqual(0, service.WriteCalls);
        }
        [TestMethod]
        public void DebugOnlyCollectorActionAndFormAreExcludedFromRelease()
        {
            var assembly = typeof(SolutionComparerControl).Assembly;
#if DEBUG
            Assert.IsNotNull(assembly.GetType("D365SolutionComparer.Services.Membership.AppActionEvidenceCollector"));
            Assert.IsNotNull(typeof(MembershipResultsForm).GetProperty("CaptureAppActionEvidence", BindingFlags.Instance | BindingFlags.NonPublic));
#else
            Assert.IsNull(assembly.GetType("D365SolutionComparer.Services.Membership.AppActionEvidenceCollector"));
            Assert.IsNull(assembly.GetType("D365SolutionComparer.AppActionEvidenceResultsForm"));
            Assert.IsNull(typeof(MembershipResultsForm).GetProperty("CaptureAppActionEvidence", BindingFlags.Instance | BindingFlags.NonPublic));
            Assert.IsNull(typeof(SolutionComparerControl).GetMethod("CaptureAppActionEvidence", BindingFlags.Instance | BindingFlags.NonPublic));
            Assert.IsFalse(typeof(MembershipCoverageDetailsForm).GetConstructors().SelectMany(c => c.GetParameters()).Any(p => p.Name == "captureAppActionEvidence"));
#endif
        }
#if DEBUG
        private const string Backing = "fixture_action", Primary = "fixture_actionid", Secret = "PRIVATE-CONTRACT-PAYLOAD-URL-TOKEN";
        [TestMethod]
        public void ZeroMembersPerformZeroReads()
        {
            var pair = new Pair(); var report = pair.Capture();
            Assert.AreEqual(0, pair.Source.Reads + pair.Target.Reads); Assert.AreEqual(0, report.Pairs.Count);
            Assert.IsTrue(report.Sides.All(s => s.Requests.Count == 0));
        }
        [DataTestMethod, DataRow(10266), DataRow(10267)]
        public void PopulatedSubtypeDoesNotReadEmptySubtypeOrTarget(int type)
        {
            var pair = new Pair(); pair.Source.Add(type); var report = pair.Capture();
            Assert.AreEqual(3, pair.Source.Reads); Assert.AreEqual(0, pair.Target.Reads);
            Assert.IsTrue(report.Sides.Where(s => !s.IsSource || s.RawType != type).All(s => s.Requests.Count == 0 && s.Rows.Count == 0));
            Assert.AreEqual("OneSidedEvidence", report.Pairs.Single().Outcome);
        }
        [TestMethod]
        public void BothRegisteredDefinitionsCanMapToSameMetadataDiscoveredBacking()
        {
            var pair = new Pair(); pair.Source.Add(10266); pair.Target.Add(10267);
            var report = pair.Capture(); var populated = report.Sides.Where(s => s.Raw.Count > 0).ToArray();
            Assert.IsTrue(populated.All(s => s.EntityName == Backing && s.PrimaryId == Primary));
            Assert.IsTrue(report.Build().Contains("objecttypecode=10266") && report.Build().Contains("objecttypecode=10267"));
        }
        [TestMethod]
        public void MissingCompletedDefinitionUsesOnlyScopedRegisteredLookup()
        {
            var pair = new Pair(); var row = pair.Source.Add(10266); pair.Source.Raw.Clear();
            pair.Source.Raw.Add(Unsupported(10266, row.Id, null)); var report = pair.Capture();
            Assert.AreEqual(Backing, report.Sides.Single(s => s.IsSource && s.RawType == 10266).EntityName);
            Assert.AreEqual(4, pair.Source.Reads);
            Assert.AreEqual(10266, pair.Source.Queries.First().Criteria.Conditions.Single().Values.Single());
        }
        [TestMethod]
        public void ConflictingRegisteredMappingsDoNotGuessBackingOrReadMetadata()
        {
            var pair = new Pair(); var row = pair.Source.Add(10266);
            pair.Source.Raw.Add(Unsupported(10266, row.Id, "other_backing")); var report = pair.Capture();
            Assert.AreEqual(0, pair.Source.Reads); Assert.IsNull(Record(report, true, 10266).CandidateA);
        }
        [TestMethod]
        public void DirectCorrelationIsExactAndCandidateNeverUsesLocalGuid()
        {
            var pair = new Pair(); var left = pair.Source.Add(10266); var right = pair.Target.Add(10266);
            var report = pair.Capture(); var match = report.Pairs.Single();
            Assert.AreEqual("DiagnosticSemanticPairHypothesis", match.Outcome);
            Assert.AreEqual(left.Id, match.Source.PrimaryId); Assert.AreEqual(right.Id, match.Target.PrimaryId);
            Assert.AreNotEqual(left.Id, right.Id); Assert.IsFalse(match.Source.CandidateA.Contains(left.Id.ToString()));
            Assert.IsTrue(match.Categories.Contains("DifferentPrimaryId"));
        }
        [TestMethod]
        public void CrossTypeSameCompleteSemanticCandidateIsOnlyAHypothesis()
        {
            var pair = new Pair(); pair.Source.Add(10266, unique: " Publisher_Action "); pair.Target.Add(10267, unique: "publisher_action")["ismanaged"] = true;
            var report = pair.Capture(); var match = report.Pairs.Single();
            Assert.AreEqual("DiagnosticSemanticPairHypothesis", match.Outcome);
            Assert.AreEqual(10266, match.Source.RawType); Assert.AreEqual(10267, match.Target.RawType);
            Assert.IsTrue(match.Categories.Contains("UnmanagedToManaged"));
            StringAssert.Contains(report.Build(), "CrossTypeDiagnosticOnly=True");
            StringAssert.Contains(report.Build(), "no production key");
        }
        [TestMethod]
        public void SubtypeMismatchCannotSilentlyPair()
        {
            var pair = new Pair(); pair.Source.Add(10266, subtype: 1); pair.Target.Add(10267, subtype: 2);
            var report = pair.Capture(); Assert.IsTrue(report.Pairs.All(p => p.Outcome == "OneSidedEvidence"));
            StringAssert.Contains(report.Build(), "SubtypeMismatch - not paired");
        }
        [TestMethod]
        public void SameBackingTableAndDisplayNameCannotResolveSubtypeUncertainty()
        {
            var pair = new Pair(); pair.Source.Add(10266); pair.Target.Add(10267);
            foreach (var side in new[] { pair.Source, pair.Target }) side.Attributes.RemoveAll(a => a.LogicalName == "actiontype");
            var report = pair.Capture(); Assert.IsTrue(report.Pairs.All(p => p.Outcome == "Incomplete"));
            Assert.IsNull(Record(report, true, 10266).CandidateA);
            Assert.IsNotNull(Record(report, true, 10266).CandidateB);
            StringAssert.Contains(report.Build(), "Semantic subtype unavailable");
        }
        [TestMethod]
        public void CandidateACollisionsCannotBeRepairedByBHashDisplayOrRawType()
        {
            var pair = new Pair(); pair.Source.Add(10266); pair.Source.Add(10267)["name"] = "Different display";
            pair.Target.Add(10266); var report = pair.Capture();
            Assert.IsTrue(report.Pairs.All(p => p.Outcome == "Ambiguous"));
            Assert.IsTrue(report.Sides.Where(s => s.IsSource).SelectMany(s => s.Rows.Values).All(r => r.DuplicateA));
        }
        [DataTestMethod, DataRow(false), DataRow(true)]
        public void VerifiedOrIncompleteParentIsIndependentOfDisplayCandidate(bool missing)
        {
            var pair = new Pair(); var row = pair.Source.Add(10266);
            pair.Source.Attributes.Add(Attribute("parentappmoduleid", new LookupAttributeMetadata { Targets = new[] { "appmodule" } }));
            var id = Guid.NewGuid(); row["parentappmoduleid"] = new EntityReference("appmodule", id);
            if (!missing) pair.Source.Context.Add(new ComponentIdentity(new SolutionComponentRecord(Guid.NewGuid(), 80, id), IdentityResolutionStatus.Resolved, "publisher_app"));
            var found = Record(pair.Capture(), true, 10266);
            Assert.AreEqual(!missing, found.ParentComplete); Assert.AreEqual(!missing, found.CandidateA != null);
            Assert.IsNotNull(found.CandidateB); Assert.AreEqual(3, pair.Source.Reads);
        }
        [TestMethod]
        public void MultiplePortableParentCandidatesRemainAmbiguous()
        {
            var pair = new Pair(); var row = pair.Source.Add(10266); var id = Guid.NewGuid();
            pair.Source.Attributes.Add(Attribute("parentappmoduleid", new LookupAttributeMetadata { Targets = new[] { "appmodule" } }));
            row["parentappmoduleid"] = new EntityReference("appmodule", id);
            foreach (var key in new[] { "app_a", "app_b" }) pair.Source.Context.Add(new ComponentIdentity(new SolutionComponentRecord(Guid.NewGuid(), 80, id), IdentityResolutionStatus.Resolved, key));
            var found = Record(pair.Capture(), true, 10266); Assert.IsFalse(found.ParentComplete); Assert.IsNull(found.CandidateA);
            Assert.AreEqual("Ambiguous", found.ParentStatus);
        }
        [TestMethod]
        public void MissingBackingRecordIsNotMembershipAbsence()
        {
            var pair = new Pair(); pair.Source.Add(10266); pair.Source.Data.Clear();
            var found = Record(pair.Capture(), true, 10266); Assert.AreEqual("Missing", found.Status); Assert.IsNull(found.CandidateA);
        }
        [DataTestMethod, DataRow(false), DataRow(true)]
        public void DuplicateOrConflictingBackingRowsBlockA(bool conflict)
        {
            var pair = new Pair(); var row = pair.Source.Add(10266); var copy = new Entity(Backing, row.Id);
            foreach (var field in row.Attributes) copy[field.Key] = field.Value;
            if (conflict) copy["uniquename"] = "other_action"; pair.Source.Data.Add(copy);
            var found = Record(pair.Capture(), true, 10266); Assert.AreEqual("Duplicate", found.Status); Assert.IsNull(found.CandidateA);
        }
        [TestMethod]
        public void ConflictingPrimaryKeyNeverCreatesACandidate()
        {
            var pair = new Pair(); pair.Source.Add(10266)[Primary] = Guid.NewGuid();
            var found = Record(pair.Capture(), true, 10266); Assert.AreEqual("Incomplete", found.Status); Assert.IsNull(found.CandidateA);
        }
        [TestMethod]
        public void RepeatedRawMembershipIsAuditDuplicationNotBackingCollision()
        {
            var pair = new Pair(); var row = pair.Source.Add(10266); pair.Source.Raw.Add(Unsupported(10266, row.Id, Backing));
            pair.Target.Add(10266); var report = pair.Capture(); Assert.AreEqual(1, report.Pairs.Count);
            Assert.AreEqual("DiagnosticSemanticPairHypothesis", report.Pairs.Single().Outcome);
            StringAssert.Contains(report.Build(), "references=2"); Assert.AreEqual(3, pair.Source.Reads);
        }
        [DataTestMethod, DataRow(200, 2), DataRow(201, 4)]
        public void BackingQueriesBatchAtTwoHundredWithoutPerActionPattern(int count, int backingReads)
        {
            var pair = new Pair(); for (int i = 0; i < count; i++) pair.Source.Add(10266, unique: "publisher_action_" + i);
            pair.Capture(); Assert.AreEqual(backingReads, pair.Source.Queries.Count);
            Assert.AreEqual(1, pair.Source.Service.ExecuteCalls);
        }
        [TestMethod]
        public void TerminalCookieDoesNotInvalidateCompleteRetrieval()
        {
            var pair = new Pair(); pair.Source.Add(10266); var normal = pair.Source.Service.RetrievePage;
            pair.Source.Service.RetrievePage = q => { var result = normal(q); result.PagingCookie = "private-cookie"; return result; };
            var report = pair.Capture(); Assert.IsNotNull(Record(report, true, 10266).CandidateA);
            StringAssert.Contains(report.Build(), "PagingCookieSupplied=True"); Assert.IsFalse(report.Build().Contains("private-cookie"));
        }
        [TestMethod]
        public void MultiPageRowsDeduplicateIdenticalOverlapAfterTerminalPage()
        {
            var pair = new Pair(); var first = pair.Source.Add(10266); var second = pair.Source.Add(10266, unique: "second_action");
            pair.Source.Service.RetrievePage = q => q.PageInfo.PageNumber == 1
                ? new EntityCollection(new List<Entity> { Clone(first, q) }) { MoreRecords = true, PagingCookie = "page1" }
                : Rows(Clone(first, q), Clone(second, q));
            var report = pair.Capture(); Assert.IsTrue(report.Sides.SelectMany(s => s.Rows.Values).All(r => r.CandidateA != null && !r.DuplicateA));
            Assert.AreEqual(4, pair.Source.Service.Calls);
        }
        [TestMethod]
        public void StalledPagingRemainsIncomplete()
        {
            var pair = new Pair(); var row = pair.Source.Add(10266);
            pair.Source.Service.RetrievePage = q => new EntityCollection(new List<Entity> { Clone(row, q) }) { MoreRecords = true, PagingCookie = "same" };
            var found = Record(pair.Capture(), true, 10266); Assert.AreEqual("Incomplete", found.Status); Assert.IsNull(found.CandidateA);
        }
        [DataTestMethod, DataRow(false), DataRow(true)]
        public void RuntimeFaultIsolationPreservesPrimaryAndBlocksOnlyCriticalEvidence(bool critical)
        {
            var pair = new Pair(); pair.Source.Add(10266); Fail(pair.Source, critical ? "uniquename" : "configuration");
            var report = pair.Capture(); var found = Record(report, true, 10266);
            Assert.AreEqual("Unique", found.Status); Assert.AreEqual(!critical, found.CandidateA != null);
            StringAssert.Contains(report.Build(), "SDKErrorCode=0x8004023B"); Assert.IsFalse(report.Build().Contains(Secret));
        }
        [TestMethod]
        public void HashOnlyContentAndExcludedLookupShadowsNeverSupplyScope()
        {
            var pair = new Pair(); pair.Source.Add(10266);
            pair.Source.Attributes.Add(Attribute("parentappmoduleid", new LookupAttributeMetadata { Targets = new[] { "appmodule" } }));
            pair.Source.Attributes.Add(Attribute("parentappmoduleidname", new StringAttributeMetadata()));
            pair.Source.Attributes.Add(Attribute("binarypackage", new MemoAttributeMetadata()));
            var report = pair.Capture(); var found = Record(report, true, 10266);
            Assert.IsNotNull(found.CandidateA); Assert.IsTrue(found.Content["configuration"].Known);
            Assert.IsFalse(report.Build().Contains(Secret));
            Assert.IsTrue(pair.Source.Queries.All(q => !q.ColumnSet.Columns.Contains("parentappmoduleidname") && !q.ColumnSet.Columns.Contains("binarypackage")));
        }
        [TestMethod]
        public void BlankOrGuidInternalIdentifierCannotUseDisplayFallback()
        {
            foreach (var unique in new[] { " ", Guid.NewGuid().ToString("D") })
            {
                var pair = new Pair(); pair.Source.Add(10266, unique: unique); var found = Record(pair.Capture(), true, 10266);
                Assert.IsNull(found.CandidateA); Assert.IsNotNull(found.CandidateB);
            }
        }
        [TestMethod]
        public void SameGuidDisplayAndHashCannotCreateSemanticPairWithDifferentInternalIdentity()
        {
            var pair = new Pair(); var left = pair.Source.Add(10266); var right = pair.Target.Add(10267, unique: "other_action");
            pair.Target.Data.Clear(); var sameId = new Entity(Backing, left.Id);
            foreach (var field in right.Attributes) sameId[field.Key] = field.Value;
            sameId[Primary] = left.Id; pair.Target.Data.Add(sameId); pair.Target.Raw.Clear();
            pair.Target.Raw.Add(Unsupported(10267, left.Id, Backing));
            Assert.IsTrue(pair.Capture().Pairs.All(p => p.Outcome == "OneSidedEvidence"));
        }
        [TestMethod]
        public void UnreadableInternalIdentifierCannotBeRepairedByDisplayOrSubtype()
        {
            var pair = new Pair(); pair.Source.Add(10266);
            Set(pair.Source.Attributes.Single(a => a.LogicalName == "uniquename"), "IsValidForRead", false);
            var found = Record(pair.Capture(), true, 10266); Assert.IsNull(found.CandidateA); Assert.IsNotNull(found.CandidateB);
            Assert.IsFalse(found.RuntimeColumns.Contains("uniquename"));
        }
        [TestMethod]
        public void PrimaryQueryFaultLeavesEvidenceFaultedWithoutUnscopedFallback()
        {
            var pair = new Pair(); pair.Source.Add(10266); Fail(pair.Source, Primary);
            var found = Record(pair.Capture(), true, 10266); Assert.AreEqual("Faulted", found.Status); Assert.IsNull(found.CandidateA);
            Assert.AreEqual(2, pair.Source.Queries.Count == 0 ? pair.Source.Service.Calls : pair.Source.Queries.Count);
            Assert.AreEqual(0, pair.Target.Reads);
        }
        [TestMethod]
        public void CancellationBeforeCapturePerformsNoReads()
        {
            var pair = new Pair(); pair.Source.Add(10266);
            using (var cancel = new CancellationTokenSource()) {
                cancel.Cancel(); Assert.ThrowsException<OperationCanceledException>(() => pair.Capture(cancel.Token));
            }
            Assert.AreEqual(0, pair.Source.Reads + pair.Target.Reads);
        }
        [TestMethod]
        public void CancellationDuringIsolationStopsImmediately()
        {
            var pair = new Pair(); pair.Source.Add(10266); Fail(pair.Source, "configuration");
            using (var cancel = new CancellationTokenSource())
            {
                var normal = pair.Source.Service.RetrievePage;
                pair.Source.Service.RetrievePage = q => { if (pair.Source.Service.Calls > 2) cancel.Cancel(); return normal(q); };
                Assert.ThrowsException<OperationCanceledException>(() => pair.Capture(cancel.Token));
                Assert.AreEqual(0, pair.Target.Reads);
            }
        }
        [TestMethod]
        public void MetadataFailureCannotGuessColumnsOrBackingIdentity()
        {
            var pair = new Pair(); pair.Source.Add(10266); pair.Source.Service.ExecuteRequest = r => { throw new InvalidOperationException(Secret); };
            var report = pair.Capture(); Assert.AreEqual("Faulted", Record(report, true, 10266).Status);
            Assert.AreEqual(0, pair.Source.Service.Calls); Assert.IsFalse(report.Build().Contains(Secret));
        }
        [TestMethod]
        public void CaptureDoesNotMutateCompletedProductionMembershipAndUsesNoWhoAmIOrWrites()
        {
            var pair = new Pair(); pair.Source.Add(10266); pair.Target.Add(10267); var source = pair.Source.Snapshot(); var target = pair.Target.Snapshot();
            var comparer = new SolutionMembershipComparer(); var before = comparer.Compare(source, target).Select(r => r.Presence).ToArray();
            new AppActionEvidenceCollector().Capture(pair.Source.Service, source, "1", pair.Target.Service, target, "2", CancellationToken.None);
            CollectionAssert.AreEqual(before, comparer.Compare(source, target).Select(r => r.Presence).ToArray());
            Assert.IsTrue(source.Components.Concat(target.Components).All(c => c.Status == IdentityResolutionStatus.Unsupported && c.ComparisonKey == null));
            Assert.AreEqual(0, pair.Source.Service.WriteCalls + pair.Target.Service.WriteCalls);
        }
        [TestMethod]
        public void CaptureButtonIsExplicitAndFormInitializationAddsNoReads()
        {
            Exception error = null; var thread = new Thread(() => {
                try {
                    var pair = new Pair(); pair.Source.Add(10266); pair.Target.Add(10267); int calls = 0;
                    var source = pair.Source.Snapshot(); var target = pair.Target.Snapshot();
                    var presentation = new MembershipResultPresenter().Create(MembershipEnvironmentResult.FromSnapshot("Source", source, 0, TimeSpan.Zero),
                        MembershipEnvironmentResult.FromSnapshot("Target", target, 0, TimeSpan.Zero));
                    using (var form = new MembershipCoverageDetailsForm("Source", new MembershipCoverageDiagnosticsBuilder().Build(source),
                        "Target", new MembershipCoverageDiagnosticsBuilder().Build(target), presentation, captureAppActionEvidence: () => calls++)) {
                        form.StartPosition = FormStartPosition.Manual; form.Location = new System.Drawing.Point(-4000, -4000); form.Show(); Application.DoEvents();
                        var button = form.Controls.OfType<FlowLayoutPanel>().SelectMany(p => p.Controls.OfType<Button>()).Single(b => b.Text == "Capture Types 10266 / 10267 App Action Evidence...");
                        Assert.AreEqual(0, calls); button.PerformClick(); Assert.AreEqual(1, calls);
                    }
                    using (var form = new AppActionEvidenceResultsForm("read-only evidence")) Assert.IsTrue(form.Controls.OfType<RichTextBox>().Single().ReadOnly);
                    Assert.AreEqual(0, pair.Source.Reads + pair.Target.Reads);
                } catch (Exception ex) { error = ex; }
            }); thread.SetApartmentState(ApartmentState.STA); thread.Start(); Assert.IsTrue(thread.Join(TimeSpan.FromSeconds(30)));
            if (error != null) System.Runtime.ExceptionServices.ExceptionDispatchInfo.Capture(error).Throw();
        }
        private static ComponentIdentity Unsupported(int type, Guid id, string backing) => new ComponentIdentity(new SolutionComponentRecord(Guid.NewGuid(), type, id),
            IdentityResolutionStatus.Unsupported, registeredDefinition: backing == null ? null : new SolutionComponentDefinitionIdentity(type, "AppAction", backing));
        private static AppActionRecordEvidence Record(AppActionEvidenceReport report, bool source, int type) => report.Sides.Single(s => s.IsSource == source && s.RawType == type).Rows.Values.Single();
        private static void Set(object obj, string name, object value) => obj.GetType().GetProperty(name).SetValue(obj, value, null);
        private static AttributeMetadata Attribute(string name, AttributeMetadata attribute) { attribute.LogicalName = name; Set(attribute, "IsValidForRead", true); return attribute; }
        private static Entity Clone(Entity row, QueryExpression query) { var copy = new Entity(row.LogicalName, row.Id); foreach (var field in query.ColumnSet.Columns) if (row.Contains(field)) copy[field] = row[field]; return copy; }
        private static void Fail(Fixture fixture, string field) {
            var normal = fixture.Service.RetrievePage; fixture.Service.RetrievePage = q => { if (q.ColumnSet.Columns.Contains(field)) throw new FaultException<OrganizationServiceFault>(
                new OrganizationServiceFault { ErrorCode = unchecked((int)0x8004023B), Message = Secret, TraceText = Secret }, new FaultReason(Secret)); return normal(q); };
        }
        private sealed class Pair
        {
            internal readonly Fixture Source = new Fixture(), Target = new Fixture();
            internal AppActionEvidenceReport Capture(CancellationToken token = default(CancellationToken)) => new AppActionEvidenceCollector().Capture(Source.Service, Source.Snapshot(), "1", Target.Service, Target.Snapshot(), "2", token);
        }
        private sealed class Fixture
        {
            internal readonly List<Entity> Data = new List<Entity>();
            internal readonly List<ComponentIdentity> Raw = new List<ComponentIdentity>(), Context = new List<ComponentIdentity>();
            internal readonly List<AttributeMetadata> Attributes = new List<AttributeMetadata>();
            internal readonly List<QueryExpression> Queries = new List<QueryExpression>();
            internal readonly FakeOrganizationService Service;
            internal int Reads => Service.Calls + Service.ExecuteCalls;
            internal Fixture() {
                Attributes.Add(Attribute(Primary, new UniqueIdentifierAttributeMetadata()));
                foreach (var name in new[] { "name", "uniquename", "installationidunique" }) Attributes.Add(Attribute(name, new StringAttributeMetadata()));
                Attributes.Add(Attribute("actiontype", new PicklistAttributeMetadata())); Attributes.Add(Attribute("ismanaged", new BooleanAttributeMetadata()));
                Attributes.Add(Attribute("configuration", new MemoAttributeMetadata()));
                Service = new FakeOrganizationService {
                    ExecuteRequest = r => {
                        Assert.IsInstanceOfType(r, typeof(RetrieveEntityRequest)); var request = (RetrieveEntityRequest)r;
                        Assert.AreEqual(Backing, request.LogicalName); Assert.IsFalse(request.RetrieveAsIfPublished);
                        var metadata = new EntityMetadata { LogicalName = Backing }; Set(metadata, "PrimaryIdAttribute", Primary);
                        Set(metadata, "PrimaryNameAttribute", "name"); Set(metadata, "Attributes", Attributes.ToArray());
                        var response = new RetrieveEntityResponse(); response.Results["EntityMetadata"] = metadata; return response;
                    },
                    RetrievePage = q => {
                        Queries.Add(q);
                        if (q.EntityName == "solutioncomponentdefinition") return Rows(new Entity("solutioncomponentdefinition", Guid.NewGuid()) {
                            ["objecttypecode"] = q.Criteria.Conditions.Single().Values.Single(), ["name"] = "AppAction", ["primaryentityname"] = Backing });
                        Assert.AreEqual(Backing, q.EntityName); var condition = q.Criteria.Conditions.Single();
                        Assert.AreEqual(Primary, condition.AttributeName); Assert.AreEqual(ConditionOperator.In, condition.Operator); Assert.IsTrue(condition.Values.Count <= 200);
                        return new EntityCollection(Data.Where(r => condition.Values.Contains(r.Id)).Select(r => Clone(r, q)).ToList());
                    }
                };
            }
            internal Entity Add(int type, string unique = "publisher_action", int subtype = 1) {
                var row = new Entity(Backing, Guid.NewGuid()) { ["uniquename"] = unique, ["name"] = "Display action", ["actiontype"] = new OptionSetValue(subtype),
                    ["ismanaged"] = false, ["configuration"] = Secret, ["installationidunique"] = Guid.NewGuid().ToString("D") };
                row[Primary] = row.Id; Data.Add(row); Raw.Add(Unsupported(type, row.Id, Backing)); return row;
            }
            internal MembershipSnapshot Snapshot() => MembershipSnapshot.Complete(Solution(), Raw.Concat(Context), DateTimeOffset.UtcNow);
        }
#endif
    }
}
