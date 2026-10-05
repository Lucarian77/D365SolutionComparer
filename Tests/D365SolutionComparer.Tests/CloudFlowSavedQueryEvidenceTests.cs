using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.ServiceModel;
using System.Threading;
using D365SolutionComparer.Models.Identity;
using D365SolutionComparer.Models.Membership;
using D365SolutionComparer.Models.ComponentDetails;
using D365SolutionComparer.Services.Membership;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Query;

namespace D365SolutionComparer.Tests
{
    [TestClass]
    public class CloudFlowSavedQueryEvidenceTests
    {
        private const string Category = "Phase2GCloudFlowEvidence";

        [TestMethod, TestCategory(Category)]
        public void ProductionCatalogAndReleaseExclusionRemainUnchanged()
        {
            Assert.AreEqual(ComponentSemanticKinds.SavedQuery, ComponentSemanticKinds.FromRawComponentType(26));
            Assert.IsNull(ComponentDefinitionContractCatalog.For(ComponentSemanticKinds.SavedQuery));
            Assert.AreEqual("process", ComponentSemanticKinds.FromRawComponentType(29));
#if !DEBUG
            var assembly = typeof(DataverseComponentIdentityResolver).Assembly;
            Assert.IsNull(assembly.GetType("D365SolutionComparer.Services.Membership.CloudFlowSavedQueryEvidenceCollector"));
            Assert.IsNull(assembly.GetType("D365SolutionComparer.ProcessQueryEvidenceResultsForm"));
            foreach (var name in new[] { "CaptureCloudFlowEvidence", "CaptureSavedQueryEvidence" })
                Assert.IsNull(typeof(MembershipResultsForm).GetProperty(name, BindingFlags.Instance | BindingFlags.NonPublic));
            Assert.IsFalse(typeof(MembershipCoverageDetailsForm).GetConstructors().SelectMany(c => c.GetParameters())
                .Any(p => p.Name == "captureCloudFlowEvidence" || p.Name == "captureSavedQueryEvidence"));
#endif
        }

#if DEBUG
        [TestMethod, TestCategory(Category)]
        public void AllType29RowsUseSchemaValidatedFieldsAndBusinessProcessRemainsContextOnly()
        {
            var pair = new Pair(false);
            pair.Source.Add(); pair.Target.Add();
            foreach (var fixture in new[] { pair.Source, pair.Target })
            {
                var bpf = fixture.Add(); bpf["category"] = new OptionSetValue(4); bpf["name"] = "Case Process";
                bpf["uniquename"] = "ava_caseprocess";
                fixture.ResolveLast("ava_caseprocess", "Business Process resolved through uniquename.");
            }
            var report = pair.Capture();
            Assert.AreEqual(2, report.Source.Rows.Count);
            foreach (var fixture in new[] { pair.Source, pair.Target })
            {
                Assert.AreEqual(1, fixture.Service.ExecuteCalls); Assert.AreEqual(1, fixture.Service.Calls);
                CollectionAssert.AreEquivalent(CloudFlowSavedQueryEvidenceCollector.WorkflowFields, fixture.Queries.Single().ColumnSet.Columns.ToArray());
                Assert.IsFalse(fixture.Queries.Single().ColumnSet.AllColumns);
                Assert.AreEqual(0, fixture.Service.WriteCalls);
            }
            StringAssert.Contains(report.Build(), "category: Source=5; Target=5; observation=EqualObserved");
            StringAssert.Contains(report.Build(), "workflowid: Source=");
            StringAssert.Contains(report.Build(), "observation=DifferentObserved");
            StringAssert.Contains(report.Build(), "WhoAmI=0; writes=0; solutioncomponent=0");
            StringAssert.Contains(report.Build(), "analysisRole=NonCloudContext");
            StringAssert.Contains(report.Build(), "ava_caseprocess");
            StringAssert.Contains(report.Build(), "category=5 candidate counts: Source=1; Target=1");
            StringAssert.Contains(report.Build(), "CloudFlowContext=Unique");
            Assert.IsFalse(report.Build().Contains("category: Source=4; Target=4"));
        }

        [DataTestMethod, TestCategory(Category)]
        [DataRow(true)] [DataRow(false)]
        public void UnresolvedSourceAndResolvedTargetFlowPairDiagnosticallyWithoutChangingMembership(bool equalWorkflowIds)
        {
            var pair = new Pair(false);
            var source = pair.Source.Add(); pair.Target.Add(equalWorkflowIds ? source.Id : (Guid?)null);
            pair.Target.ResolveLast("workflow-semantic:v1:existing-target-key", "Existing semantic fallback resolved.");
            var comparer = new SolutionMembershipComparer();
            var before = comparer.Compare(pair.Source.Snapshot(), pair.Target.Snapshot()).ToList();
            var text = pair.Capture().Build();
            var after = comparer.Compare(pair.Source.Snapshot(), pair.Target.Snapshot()).ToList();
            CollectionAssert.AreEqual(before.Select(r => r.Presence).ToList(), after.Select(r => r.Presence).ToList());
            CollectionAssert.AreEqual(before.Select(r => r.Source?.Status).ToList(), after.Select(r => r.Source?.Status).ToList());
            CollectionAssert.AreEqual(before.Select(r => r.Target?.ComparisonKey).ToList(), after.Select(r => r.Target?.ComparisonKey).ToList());
            Assert.IsTrue(after.All(r => r.Presence == MembershipPresence.Indeterminate));
            StringAssert.Contains(text, "priorIdentityStatus=Resolved; priorIdentityDiagnostic=Existing semantic fallback resolved.");
            StringAssert.Contains(text, "CloudFlowContext=Unique");
            var workflowLine = text.Split(new[] { '\n' }).Single(line => line.TrimStart().StartsWith("workflowid: Source=", StringComparison.Ordinal));
            StringAssert.Contains(workflowLine, equalWorkflowIds ? "EqualObserved" : "DifferentObserved");
            Assert.AreEqual(1, pair.Source.Service.Calls); Assert.AreEqual(1, pair.Target.Service.Calls);
            Assert.AreEqual(0, pair.Source.Service.WriteCalls + pair.Target.Service.WriteCalls);
        }

        [TestMethod, TestCategory(Category)]
        public void ResolvedCategory5CandidatesStillCauseAmbiguityAndAreNotAutoPaired()
        {
            var pair = new Pair(false); pair.Source.Add(); pair.Target.Add();
            pair.Target.ResolveLast("existing-key-1", "Resolved.");
            pair.Target.Add(); pair.Target.ResolveLast("existing-key-2", "Resolved.");
            var text = pair.Capture().Build();
            StringAssert.Contains(text, "CloudFlowContext=Ambiguous");
            Assert.IsFalse(text.Contains("workflowid: Source="));
        }

        [DataTestMethod, TestCategory(Category)]
        [DataRow("missing")] [DataRow("unknownCategory")]
        public void UncertainAdditionalWorkflowContextBlocksAutomaticPairing(string failure)
        {
            var pair = new Pair(false); pair.Source.Add(); pair.Target.Add();
            var context = pair.Source.Add();
            if (failure == "missing") pair.Source.Rows.Remove(context);
            else context.Attributes.Remove("category");
            var text = pair.Capture().Build();
            Assert.IsFalse(text.Contains("workflowid: Source="));
            StringAssert.Contains(text, "no automatic pairing");
        }

        [TestMethod, TestCategory(Category)]
        public void HashesNeverExposeClientDataOrXamlContentAndEqualIdsAreEvidenceOnly()
        {
            var pair = new Pair(false); var id = Guid.NewGuid();
            var source = pair.Source.Add(id); var target = pair.Target.Add(id);
            var installation = Guid.NewGuid(); source["workflowidunique"] = installation; target["workflowidunique"] = installation;
            source["clientdata"] = "secret-json"; target["clientdata"] = "other-secret-json";
            source["xaml"] = "<secret-xaml/>"; target["xaml"] = "<secret-xaml/>";
            var before = new SolutionMembershipComparer().Compare(pair.Source.Snapshot(), pair.Target.Snapshot()).Select(r => r.Presence).ToList();
            var report = pair.Capture(); var text = report.Build();
            Assert.IsFalse(text.Contains("secret-json")); Assert.IsFalse(text.Contains("secret-xaml"));
            StringAssert.Contains(text, "rawSha256="); StringAssert.Contains(text, "installation-specific");
            StringAssert.Contains(text, "workflowidunique: Source=" + installation.ToString("D"));
            CollectionAssert.AreEqual(before, new SolutionMembershipComparer().Compare(pair.Source.Snapshot(), pair.Target.Snapshot()).Select(r => r.Presence).ToList());
            Assert.IsTrue(pair.Source.Raw.Concat(pair.Target.Raw).All(c => c.ComparisonKey == null));
            Assert.IsTrue(before.All(p => p == MembershipPresence.Indeterminate));
        }

        [TestMethod, TestCategory(Category)]
        public void UnavailableColumnsAreNotGuessedAndExtraExposedFlowIdentifierIsAuditOnly()
        {
            var pair = new Pair(false); pair.Source.Add(); pair.Target.Add();
            pair.Source.Fields.Remove("modernflowtype"); pair.Source.Fields.Remove("xaml");
            pair.Source.Fields.Add("new_cloudflowid"); pair.Source.Rows[0]["new_cloudflowid"] = Guid.NewGuid();
            var report = pair.Capture();
            Assert.IsFalse(pair.Source.Queries.Single().ColumnSet.Columns.Contains("modernflowtype"));
            Assert.IsTrue(pair.Source.Queries.Single().ColumnSet.Columns.Contains("new_cloudflowid"));
            StringAssert.Contains(report.Build(), "modernflowtype=UnavailableInReadableMetadata");
            StringAssert.Contains(report.Build(), "modernflowtype: Source=(not returned/null)");
            StringAssert.Contains(report.Build(), "observation=Unavailable");
        }

        [TestMethod, TestCategory(Category)]
        public void MultipleUnresolvedFlowsAreNotPairedByDisplayNameOrLocalGuid()
        {
            var pair = new Pair(false);
            pair.Source.Add(); pair.Source.Add(); pair.Target.Add(); pair.Target.Add();
            StringAssert.Contains(pair.Capture().Build(), "no automatic pairing");
        }

        [TestMethod, TestCategory(Category)]
        public void NonModernWorkflowIsNotAutomaticallyTreatedAsCloudFlow()
        {
            var pair = new Pair(false); pair.Source.Add()["category"] = new OptionSetValue(4); pair.Target.Add();
            StringAssert.Contains(pair.Capture().Build(), "Insufficient/ambiguous Cloud Flow context");
        }

        [DataTestMethod, TestCategory(Category)]
        [DataRow(false)] [DataRow(true)]
        public void ZeroPopulationUsesNoRequests(bool savedQuery)
        {
            var pair = new Pair(savedQuery); var report = pair.Capture();
            Assert.AreEqual(0, report.Source.Requests.Count + report.Target.Requests.Count);
            Assert.AreEqual(0, pair.Source.Service.Calls + pair.Source.Service.ExecuteCalls);
        }

        [DataTestMethod, TestCategory(Category)]
        [DataRow(false, 200, 1)] [DataRow(false, 201, 2)]
        [DataRow(true, 200, 1)] [DataRow(true, 201, 2)]
        public void BatchingDeduplicatesRawMembershipWithGuidTypedDeterministicFilters(bool savedQuery, int count, int batches)
        {
            var pair = new Pair(savedQuery);
            for (int i = 0; i < count; i++) pair.Source.Add();
            pair.Source.Raw.Add(pair.Source.Raw[0]);
            var report = pair.Capture();
            Assert.AreEqual(batches, pair.Source.Queries.Count);
            Assert.AreEqual(count, report.Source.Rows.Count); Assert.AreEqual(count + 1, report.Source.Raw.Count);
            var ids = new List<Guid>();
            foreach (var query in pair.Source.Queries)
            {
                var condition = query.Criteria.Conditions.Single();
                Assert.AreEqual(pair.Source.Primary, condition.AttributeName); Assert.AreEqual(ConditionOperator.In, condition.Operator);
                Assert.IsTrue(condition.Values.Count <= 200); Assert.IsTrue(condition.Values.All(v => v is Guid));
                ids.AddRange(condition.Values.Cast<Guid>());
            }
            CollectionAssert.AreEqual(ids.OrderBy(id => id).ToList(), ids);
            Assert.AreEqual(batches + 1, report.Source.Requests.Count);
        }

        [DataTestMethod, TestCategory(Category)]
        [DataRow("missing", "Missing")] [DataRow("duplicate", "DuplicateReturnedRow")]
        [DataRow("conflict", "ConflictingPrimaryKey")] [DataRow("unexpected", "ConflictingPrimaryKey")]
        [DataRow("paged", "IncompleteBatch")] [DataRow("null", "IncompleteBatch")]
        [DataRow("fault", "FaultedBatch (server details withheld)")]
        public void UncertainCorrelationNeverCreatesCloudFlowPairOrAbsence(string failure, string status)
        {
            var pair = new Pair(false); var row = pair.Source.Add(); pair.Target.Add();
            pair.Source.Service.RetrievePage = query =>
            {
                if (failure == "fault") throw new FaultException("sensitive server message");
                if (failure == "null") return null;
                if (failure == "missing") return new EntityCollection();
                if (failure == "conflict") row["workflowid"] = Guid.NewGuid();
                if (failure == "unexpected") row.LogicalName = "other";
                return new EntityCollection(failure == "duplicate" ? new[] { row, row } : new[] { row }) { MoreRecords = failure == "paged" };
            };
            var report = pair.Capture(); Assert.AreEqual(status, report.Source.Rows.Single().Value.Status);
            StringAssert.Contains(report.Build(), "no automatic pairing"); Assert.IsFalse(report.Build().Contains("sensitive server message"));
        }

        [DataTestMethod, TestCategory(Category)]
        [DataRow("fault")] [DataRow("null")] [DataRow("wrongPrimary")] [DataRow("duplicateAttribute")]
        public void MetadataFailureDoesNotGuessFieldsOrQueryBackingRows(string failure)
        {
            var pair = new Pair(false); pair.Source.Add();
            pair.Source.Service.ExecuteRequest = request =>
            {
                if (failure == "fault") throw new FaultException("secret");
                if (failure == "null") return null;
                var response = pair.Source.Schema(); var metadata = response.EntityMetadata;
                if (failure == "wrongPrimary") Set(metadata, "PrimaryIdAttribute", "other");
                else Set(metadata, "Attributes", metadata.Attributes.Concat(new[] { metadata.Attributes[0] }).ToArray());
                return response;
            };
            var report = pair.Capture(); Assert.AreEqual(0, pair.Source.Service.Calls);
            Assert.AreEqual("Unavailable", report.Source.Rows.Single().Value.Status);
            Assert.IsFalse(report.Build().Contains("secret"));
        }

        [DataTestMethod, TestCategory(Category)]
        [DataRow(false)] [DataRow(true)]
        public void CancellationDuringMetadataOrBackingReadPropagates(bool backing)
        {
            var pair = new Pair(false); pair.Source.Add();
            using (var cancellation = new CancellationTokenSource())
            {
                if (backing) pair.Source.Service.RetrievePage = query => { cancellation.Cancel(); return new EntityCollection(pair.Source.Rows); };
                else pair.Source.Service.ExecuteRequest = request => { cancellation.Cancel(); return pair.Source.Schema(); };
                Assert.ThrowsException<OperationCanceledException>(() => pair.Capture(cancellation.Token));
            }
        }

        [TestMethod, TestCategory(Category)]
        public void SavedQueryCandidatesRemainSeparateCaseInsensitiveAndIgnoreGuidXmlAndManagedDifferences()
        {
            var pair = new Pair(true); var source = pair.Source.Add(); var target = pair.Target.Add();
            source["name"] = " My View "; target["name"] = "my view"; target["ismanaged"] = true;
            target["layoutxml"] = "<layout>changed</layout>";
            var report = pair.Capture(); var l = report.Source.Rows.Single().Value; var r = report.Target.Rows.Single().Value;
            Assert.IsTrue(StringComparer.OrdinalIgnoreCase.Equals(l.Candidate(true), r.Candidate(true)));
            Assert.IsTrue(StringComparer.OrdinalIgnoreCase.Equals(l.Candidate(false), r.Candidate(false)));
            Assert.AreNotEqual(l.Candidate(true), l.Candidate(false)); Assert.AreNotEqual(source.Id, target.Id);
            StringAssert.Contains(report.Build(), "UniqueCandidatePair"); StringAssert.Contains(report.Build(), "canonicalXmlSha256=");
            Assert.IsFalse(report.Build().Contains("<layout>"));
        }

        [TestMethod, TestCategory(Category)]
        public void SavedQueryDistinctBackingCollisionsRemainAmbiguousButRepeatedRawMembershipDoesNot()
        {
            var pair = new Pair(true); pair.Source.Add(); pair.Target.Add();
            pair.Source.Raw.Add(pair.Source.Raw[0]);
            Assert.IsFalse(pair.Capture().Build().Contains("evidenceStatus=Ambiguous"));
            pair.Source.Add(); var report = pair.Capture();
            StringAssert.Contains(report.Build(), "evidenceStatus=Ambiguous");
            Assert.IsFalse(report.Build().Contains("UniqueCandidatePair"));
        }

        [TestMethod, TestCategory(Category)]
        public void QueryTypeDistinguishesCandidateAWhileCandidateBCollisionIsPreserved()
        {
            var pair = new Pair(true); pair.Source.Add(); pair.Source.Add()["querytype"] = 1;
            var report = pair.Capture(); var rows = report.Source.Rows.Values.ToList();
            Assert.AreNotEqual(rows[0].Candidate(true), rows[1].Candidate(true));
            Assert.AreEqual(rows[0].Candidate(false), rows[1].Candidate(false));
            StringAssert.Contains(report.Build(), "evidenceStatus=Ambiguous");
        }

        [DataTestMethod, TestCategory(Category)]
        [DataRow("name")] [DataRow("returnedtypecode")] [DataRow("querytype")]
        public void MissingSavedQueryIdentityEvidenceNeverForcesCandidateA(string field)
        {
            var pair = new Pair(true); pair.Source.Add().Attributes.Remove(field);
            Assert.IsNull(pair.Capture().Source.Rows.Single().Value.Candidate(true));
        }

        [TestMethod, TestCategory(Category)]
        public void NumericQueryScopeUsesOneGroupedMetadataReadAndNoPerViewQueries()
        {
            var pair = new Pair(true); pair.Source.Add()["returnedtypecode"] = 1; pair.Source.Add()["returnedtypecode"] = "1";
            var initial = pair.Source.Service.ExecuteRequest;
            pair.Source.Service.ExecuteRequest = request =>
            {
                if (request is RetrieveEntityRequest) return initial(request);
                Assert.IsInstanceOfType(request, typeof(RetrieveMetadataChangesRequest));
                var entity = new EntityMetadata { LogicalName = "account" }; Set(entity, "ObjectTypeCode", 1);
                var metadata = new EntityMetadataCollection(); metadata.Add(entity);
                var response = new RetrieveMetadataChangesResponse(); response.Results["EntityMetadata"] = metadata; return response;
            };
            var report = pair.Capture(); Assert.AreEqual(2, pair.Source.Service.ExecuteCalls); Assert.AreEqual(1, pair.Source.Service.Calls);
            Assert.IsTrue(report.Source.Rows.Values.All(r => r.Scope == "account"));
        }

        [TestMethod, TestCategory(Category)]
        public void AbsentOrUnavailableSnapshotsAreRejectedBeforeReads()
        {
            var pair = new Pair(false); var snapshot = pair.Source.Snapshot();
            Assert.ThrowsException<ArgumentException>(() => new CloudFlowSavedQueryEvidenceCollector().Capture(pair.Source.Service,
                MembershipSnapshot.Unavailable(snapshot.Environment, snapshot.SolutionUniqueName, DateTimeOffset.UtcNow, "fault"), "1",
                pair.Target.Service, pair.Target.Snapshot(), "1", false, CancellationToken.None));
            Assert.AreEqual(0, pair.Source.Service.Calls + pair.Source.Service.ExecuteCalls);
        }

        private static void Set(object target, string property, object value) => target.GetType().GetProperty(property).SetValue(target, value, null);
        private sealed class Pair
        {
            internal readonly Fixture Source, Target;
            internal readonly bool SavedQuery;
            internal Pair(bool savedQuery) { SavedQuery = savedQuery; Source = new Fixture(savedQuery, "Source"); Target = new Fixture(savedQuery, "Target"); }
            internal ProcessQueryEvidenceReport Capture(CancellationToken token = default(CancellationToken)) => new CloudFlowSavedQueryEvidenceCollector()
                .Capture(Source.Service, Source.Snapshot(), "1", Target.Service, Target.Snapshot(), "2", SavedQuery, token);
        }
        private sealed class Fixture
        {
            internal readonly string Entity, Primary;
            internal readonly bool SavedQuery;
            internal readonly SolutionIdentity Solution;
            internal readonly List<ComponentIdentity> Raw = new List<ComponentIdentity>();
            internal readonly List<Entity> Rows = new List<Entity>();
            internal readonly List<QueryExpression> Queries = new List<QueryExpression>();
            internal readonly List<string> Fields;
            internal readonly FakeOrganizationService Service;
            internal Fixture(bool savedQuery, string label)
            {
                SavedQuery = savedQuery; Entity = savedQuery ? "savedquery" : "workflow"; Primary = Entity + "id";
                Fields = (savedQuery ? CloudFlowSavedQueryEvidenceCollector.QueryFields : CloudFlowSavedQueryEvidenceCollector.WorkflowFields).ToList();
                Solution = new SolutionIdentity(new EnvironmentIdentity(Guid.NewGuid(), label), Guid.NewGuid(), "ICMSEnhancementRelease261");
                Service = new FakeOrganizationService
                {
                    ExecuteRequest = request =>
                    {
                        Assert.IsInstanceOfType(request, typeof(RetrieveEntityRequest));
                        var retrieval = (RetrieveEntityRequest)request;
                        Assert.AreEqual(Entity, retrieval.LogicalName); Assert.AreEqual(EntityFilters.Attributes, retrieval.EntityFilters);
                        Assert.IsFalse(retrieval.RetrieveAsIfPublished); return Schema();
                    },
                    RetrievePage = query =>
                    {
                        Queries.Add(query); Assert.AreEqual(Entity, query.EntityName);
                        var ids = query.Criteria.Conditions.Single().Values.Cast<Guid>().ToList();
                        return new EntityCollection(Rows.Where(r => ids.Contains(r.Id)).Select(r =>
                        {
                            var row = new Entity(Entity, r.Id);
                            foreach (var column in query.ColumnSet.Columns) if (r.Contains(column)) row[column] = r[column];
                            return row;
                        }).ToList());
                    }
                };
            }
            internal RetrieveEntityResponse Schema()
            {
                var metadata = new EntityMetadata { LogicalName = Entity };
                Set(metadata, "PrimaryIdAttribute", Primary);
                Set(metadata, "Attributes", Fields.Select(field =>
                {
                    AttributeMetadata attribute = field.EndsWith("id", StringComparison.Ordinal) || field.EndsWith("idunique", StringComparison.Ordinal) ?
                        (AttributeMetadata)new UniqueIdentifierAttributeMetadata() : new StringAttributeMetadata();
                    attribute.LogicalName = field; Set(attribute, "IsValidForRead", true); return attribute;
                }).ToArray());
                var response = new RetrieveEntityResponse(); response.Results["EntityMetadata"] = metadata; return response;
            }
            internal Entity Add(Guid? id = null)
            {
                var row = new Entity(Entity, id ?? Guid.NewGuid()) { [Primary] = id ?? Guid.Empty,
                    ["name"] = "Flow or View", ["ismanaged"] = false, ["componentstate"] = new OptionSetValue(0) };
                row[Primary] = row.Id;
                if (SavedQuery) { row["returnedtypecode"] = "account"; row["querytype"] = 0; row["savedqueryidunique"] = Guid.NewGuid(); row["layoutxml"] = "<layout/>"; }
                else { row["category"] = new OptionSetValue(5); row["type"] = new OptionSetValue(1); row["workflowidunique"] = Guid.NewGuid(); }
                Rows.Add(row); Raw.Add(new ComponentIdentity(new SolutionComponentRecord(Guid.NewGuid(), SavedQuery ? 26 : 29, row.Id),
                    SavedQuery ? IdentityResolutionStatus.Unsupported : IdentityResolutionStatus.Unresolved)); return row;
            }
            internal void ResolveLast(string key, string diagnostic)
            {
                var index = Raw.Count - 1;
                Raw[index] = new ComponentIdentity(Raw[index].Record, IdentityResolutionStatus.Resolved, key, diagnostic);
            }
            internal MembershipSnapshot Snapshot() => MembershipSnapshot.Complete(Solution, Raw, DateTimeOffset.UtcNow);
        }
#endif
    }
}
