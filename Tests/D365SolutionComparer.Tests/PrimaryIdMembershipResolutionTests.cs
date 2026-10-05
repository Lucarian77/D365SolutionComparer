using System;
using System.Collections.Generic;
using System.Linq;
using System.ServiceModel;
using System.Threading;
using D365SolutionComparer.Infrastructure;
using D365SolutionComparer.Models.Identity;
using D365SolutionComparer.Models.Membership;
using D365SolutionComparer.Models.ComponentDetails;
using D365SolutionComparer.Services.Membership;
using D365SolutionComparer.Services.ComponentDetails;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using static D365SolutionComparer.Tests.MembershipTestData;

namespace D365SolutionComparer.Tests
{
    [TestClass]
    public class PrimaryIdMembershipResolutionTests
    {
        private const string Category = "Phase2GPrimaryIdMembership";

        [DataTestMethod, TestCategory(Category)]
        [DataRow(29, "workflowidunique")]
        [DataRow(29, "resourceid")]
        [DataRow(29, "solutionid")]
        [DataRow(29, "ismanaged")]
        [DataRow(29, "name")]
        [DataRow(29, "uniquename")]
        [DataRow(29, "modernflowtype")]
        [DataRow(29, "clientdata")]
        [DataRow(29, "xaml")]
        [DataRow(26, "savedqueryidunique")]
        [DataRow(26, "fetchxml")]
        [DataRow(26, "layoutxml")]
        [DataRow(26, "columnsetxml")]
        [DataRow(26, "ismanaged")]
        [DataRow(26, "name")]
        [DataRow(26, "returnedtypecode")]
        [DataRow(26, "querytype")]
        public void AuditOrCandidateChangesCannotOverridePrimaryIdentity(int type, string field)
        {
            var a = Row(type); var b = Row(type, a.Id);
            a[field] = field == "ismanaged" ? (object)false : "before";
            b[field] = field == "ismanaged" ? (object)true : "after";
            var source = new Fixture(type, a); var target = new Fixture(type, b);
            var result = Compare(source, target).Single();
            Assert.AreEqual(MembershipPresence.PresentInBoth, result.Presence);
            Assert.AreEqual(IdentityResolutionStatus.Resolved, result.Source.Status);
            Assert.AreEqual(result.Source.ComparisonKey, result.Target.ComparisonKey);
            StringAssert.StartsWith(result.Source.ComparisonKey, type == 29 ? WorkflowSemanticPolicy.CloudFlowPrefix : "savedquery:v1:savedqueryid:");
        }

        [DataTestMethod, TestCategory(Category)]
        [DataRow(29)] [DataRow(26)]
        public void IdenticalContextAndHashesNeverMatchDifferentPrimaryIds(int type)
        {
            var a = Row(type); var b = Row(type);
            a["clientdata"] = b["clientdata"] = "same";
            a["fetchxml"] = b["fetchxml"] = "<fetch/>";
            var rows = Compare(new Fixture(type, a), new Fixture(type, b));
            Assert.AreEqual(2, rows.Count);
            Assert.AreEqual(1, rows.Count(r => r.Presence == MembershipPresence.OnlyInSource));
            Assert.AreEqual(1, rows.Count(r => r.Presence == MembershipPresence.OnlyInTarget));
        }

        [DataTestMethod, TestCategory(Category)]
        [DataRow(29, true)] [DataRow(29, false)] [DataRow(26, true)] [DataRow(26, false)]
        public void OneSidedIdentityRequiresCompleteOppositeKindCoverage(int type, bool sourceSide)
        {
            var present = new Fixture(type, Row(type)); var empty = new Fixture(type);
            Assert.AreEqual(sourceSide ? MembershipPresence.OnlyInSource : MembershipPresence.OnlyInTarget,
                (sourceSide ? Compare(present, empty) : Compare(empty, present)).Single().Presence);
            var missing = new Fixture(type); missing.Resolve(new[] { Guid.NewGuid() });
            var result = sourceSide ? Compare(present, missing) : Compare(missing, present);
            Assert.IsTrue(result.All(r => r.Presence == MembershipPresence.Indeterminate));
        }

        [DataTestMethod, TestCategory(Category)]
        [DataRow(29)] [DataRow(26)]
        public void RepeatedRawPrimaryIdentityIsAmbiguousAndQueryIsDeduplicated(int type)
        {
            var row = Row(type); var fixture = new Fixture(type, row);
            fixture.Resolve(new[] { row.Id, row.Id });
            Assert.AreEqual(2, fixture.Membership.Components.Count);
            Assert.IsTrue(fixture.Membership.Components.All(i => i.Status == IdentityResolutionStatus.Ambiguous && i.ComparisonKey == null));
            Assert.AreEqual(1, fixture.Queries.Single().Criteria.Conditions.Single().Values.Count);
            Assert.IsTrue(Compare(fixture, new Fixture(type)).All(r => r.Presence == MembershipPresence.Indeterminate));
        }

        [DataTestMethod, TestCategory(Category)]
        [DataRow(29)] [DataRow(26)]
        public void DuplicateBackingRowsAreAmbiguous(int type)
        {
            var row = Row(type); var fixture = new Fixture(type, row, Row(type, row.Id));
            fixture.Resolve(new[] { row.Id });
            Assert.AreEqual(IdentityResolutionStatus.Ambiguous, fixture.Membership.Components.Single().Status);
        }

        [DataTestMethod, TestCategory(Category)]
        [DataRow(29, false)] [DataRow(29, true)] [DataRow(26, false)] [DataRow(26, true)]
        public void BlankOrAbsentPrimaryFieldCannotResolve(int type, bool absent)
        {
            var row = Row(type); var field = Primary(type);
            if (absent) row.Attributes.Remove(field); else row[field] = Guid.Empty;
            Assert.AreEqual(IdentityResolutionStatus.Unresolved, new Fixture(type, row).Membership.Components.Single().Status);
        }

        [DataTestMethod, TestCategory(Category)]
        [DataRow(29, "missing")] [DataRow(26, "missing")]
        [DataRow(29, "paging")] [DataRow(26, "paging")]
        [DataRow(29, "conflicting")] [DataRow(26, "conflicting")]
        [DataRow(29, "fault")] [DataRow(26, "fault")]
        public void DefectiveRetrievalCannotResolveOrEstablishAbsence(int type, string defect)
        {
            var row = Row(type); var fixture = new Fixture(type);
            fixture.Reply = query =>
            {
                if (defect == "fault") throw new FaultException("Test fault");
                if (defect == "missing") return Rows();
                if (defect == "conflicting") row[Primary(type)] = Guid.NewGuid();
                return new EntityCollection(new[] { row }.ToList()) { MoreRecords = defect == "paging" };
            };
            fixture.Resolve(new[] { row.Id });
            Assert.AreEqual(IdentityResolutionStatus.Unresolved, fixture.Membership.Components.Single().Status);
            Assert.IsTrue(Compare(fixture, new Fixture(type, Row(type))).All(r => r.Presence == MembershipPresence.Indeterminate));
        }

        [DataTestMethod, TestCategory(Category)]
        [DataRow(29)] [DataRow(26)]
        public void CancellationPropagates(int type)
        {
            var fixture = new Fixture(type); var cancellation = new CancellationTokenSource();
            fixture.Reply = query => { cancellation.Cancel(); return Rows(Row(type)); };
            Assert.ThrowsException<OperationCanceledException>(() => fixture.Resolve(new[] { Guid.NewGuid() }, cancellation.Token));
        }

        [DataTestMethod, TestCategory(Category)]
        [DataRow(29, 200, 1)] [DataRow(29, 201, 2)]
        [DataRow(26, 200, 1)] [DataRow(26, 201, 2)]
        public void ExistingBatchShapeCountsAndGuidFiltersArePreserved(int type, int count, int batches)
        {
            var rows = Enumerable.Range(0, count).Select(_ => Row(type)).ToArray();
            var fixture = new Fixture(type, rows);
            Assert.AreEqual(batches, fixture.Counter.GetQueryCount(Table(type)));
            Assert.AreEqual(1, fixture.Counter.GetExecuteCount("WhoAmI"));
            Assert.AreEqual(batches + 1, fixture.Counter.TotalRequests);
            Assert.AreEqual(0, fixture.Service.WriteCalls);
            Assert.IsTrue(fixture.Membership.Components.All(i => i.Status == IdentityResolutionStatus.Resolved));
            var expected = type == 29 ? WorkflowSemanticPolicy.Columns : new[] { "savedqueryid", "name", "returnedtypecode", "querytype", "savedqueryidunique", "componentstate", "ismanaged" };
            foreach (var query in fixture.Queries)
            {
                Assert.IsFalse(query.ColumnSet.AllColumns);
                CollectionAssert.AreEquivalent(expected, query.ColumnSet.Columns.ToArray());
                var condition = query.Criteria.Conditions.Single();
                Assert.AreEqual(Primary(type), condition.AttributeName);
                Assert.AreEqual(ConditionOperator.In, condition.Operator);
                Assert.IsTrue(condition.Values.All(v => v is Guid));
                Assert.IsTrue(condition.Values.Count <= 200);
            }
            CollectionAssert.AreEqual(rows.Select(r => r.Id).OrderBy(id => id).ToArray(),
                fixture.Queries.SelectMany(q => q.Criteria.Conditions.Single().Values.Cast<Guid>()).ToArray());
        }

        [TestMethod, TestCategory(Category)]
        public void CloudFlowPrecedesBothUniqueNameAndSemanticFallbackWithoutRequiringSubtype()
        {
            var a = Row(29); var b = Row(29, a.Id);
            a["uniquename"] = "source_name"; b["uniquename"] = "target_name";
            a.Attributes.Remove("modernflowtype"); b.Attributes.Remove("primaryentity");
            Assert.AreEqual(MembershipPresence.PresentInBoth, Compare(new Fixture(29, a), new Fixture(29, b)).Single().Presence);
        }

        [TestMethod, TestCategory(Category)]
        public void BusinessProcessUniqueNameAndFallbackAreUnchanged()
        {
            foreach (var named in new[] { true, false })
            {
                var a = Row(29); var b = Row(29);
                foreach (var row in new[] { a, b })
                {
                    row["category"] = new OptionSetValue(4); row["businessprocesstype"] = new OptionSetValue(0);
                    if (named) row["uniquename"] = "ava_caseprocess";
                }
                var result = Compare(new Fixture(29, a), new Fixture(29, b)).Single();
                Assert.AreEqual(MembershipPresence.PresentInBoth, result.Presence);
                Assert.AreEqual(named ? "ava_caseprocess" : WorkflowSemanticPolicy.Candidate(a, out _), result.Source.ComparisonKey);
            }
        }

        [TestMethod, TestCategory(Category)]
        public void UnknownCategoryKeepsLegacyResolutionButCannotProveCloudFlowAbsence()
        {
            var unknown = Row(29); unknown.Attributes.Remove("category"); unknown["uniquename"] = "legacy_unique";
            var source = new Fixture(29, unknown);
            Assert.AreEqual("legacy_unique", source.Membership.Components.Single().ComparisonKey);
            var target = new Fixture(29, Row(29));
            var flowResult = Compare(source, target).Single(r => r.Target != null);
            Assert.AreEqual(MembershipPresence.Indeterminate, flowResult.Presence);
        }

        [TestMethod, TestCategory(Category)]
        public void ScopedWorkflowAmbiguityBlocksCloudAbsenceButCannotSuppressProvenCloudMatch()
        {
            var flow = Row(29); var a = Row(29); var b = Row(29);
            foreach (var row in new[] { a, b }) row["category"] = new OptionSetValue(2);
            var source = new Fixture(29, flow, a, b);
            var target = new Fixture(29, Row(29, flow.Id), Row(29));
            var results = Compare(source, target);
            Assert.AreEqual(1, results.Count(r => r.Presence == MembershipPresence.PresentInBoth));
            Assert.AreEqual(MembershipPresence.Indeterminate, results.Single(r => r.Target?.Record.ObjectId == target.Backing[1].Id).Presence);
        }

        [TestMethod, TestCategory(Category)]
        public void UnrelatedLegacyFallbackDoesNotSuppressCloudIdMatchOrCreateFalsePair()
        {
            var flow = Row(29); var fallback = Row(29); fallback["category"] = new OptionSetValue(2);
            var source = new Fixture(29, flow); var target = new Fixture(29, Row(29, flow.Id), fallback);
            Assert.AreEqual(1, Compare(source, target).Count(r => r.Presence == MembershipPresence.PresentInBoth));
        }

        [DataTestMethod, TestCategory(Category)]
        [DataRow(true)] [DataRow(false)]
        public void SavedQueryNotComparedPresentationPreservesOneSidedMembership(bool sourceSide)
        {
            var a = new Fixture(26, Row(26)); var b = new Fixture(26);
            var source = sourceSide ? a : b; var target = sourceSide ? b : a;
            var membership = Present(source, target);
            var mapped = new ComponentDefinitionResultPresenter().Apply(membership, source.Definitions, target.Definitions);
            Assert.AreEqual("Not Compared", mapped.Rows.Single().DefinitionStatus);
            Assert.AreEqual(sourceSide ? "Source Only" : "Target Only", mapped.Rows.Single().MembershipStatus);
            Assert.AreEqual(IdentityResolutionStatus.Resolved, (mapped.Rows.Single().Comparison.Source ?? mapped.Rows.Single().Comparison.Target).Status);
            Assert.IsNull(ComponentDefinitionContractCatalog.For(ComponentSemanticKinds.SavedQuery));
        }

        [TestMethod, TestCategory(Category)]
        public void SavedQueryMatchedDefinitionIsNotComparedAndDiagnosticsDoNotFragment()
        {
            var a = Row(26); var b = Row(26, a.Id); b["name"] = "changed";
            var source = new Fixture(26, a); var target = new Fixture(26, b);
            var mapped = new ComponentDefinitionResultPresenter().Apply(Present(source, target), source.Definitions, target.Definitions);
            Assert.AreEqual("Not Compared", mapped.Rows.Single().DefinitionStatus);
            Assert.AreEqual("Present in Both", mapped.Rows.Single().MembershipStatus);
            var coverage = new MembershipCoverageDiagnosticsBuilder().Build(source.Membership);
            Assert.AreEqual(1, coverage.SemanticKinds.Single(k => k.SemanticKind == ComponentSemanticKinds.SavedQuery).Resolved);
            StringAssert.Contains(string.Join(";", source.Membership.Components.Single().DiagnosticEvidence), "diagnosticContext");
            var second = Row(26); second["name"] = "another";
            var grouped = new MembershipCoverageDiagnosticsBuilder().Build(new Fixture(26, a, second).Membership);
            Assert.AreEqual(0, grouped.SemanticKinds.Single(k => k.SemanticKind == ComponentSemanticKinds.SavedQuery).DiagnosticGroups.Count);
            Assert.AreEqual(source.Membership.Components.Single().Diagnostic,
                new Fixture(26, second).Membership.Components.Single().Diagnostic);
        }

        [TestMethod, TestCategory(Category)]
        public void NormalCoordinatedOperationReusesMembershipAndBackingReadsWithoutWrites()
        {
            var solution = Solution(); var flow = Row(29); var view = Row(26);
            var flowMember = ComponentRow(solution, 29); flowMember["objectid"] = flow.Id;
            var viewMember = ComponentRow(solution, 26); viewMember["objectid"] = view.Id;
            var counter = new DataverseRequestCounter();
            var service = Service(solution, query => query.EntityName == "solution" ? Rows(SolutionRow(solution)) :
                query.EntityName == "solutioncomponent" ? Rows(flowMember, viewMember) :
                query.EntityName == "workflow" ? Rows(flow) : query.EntityName == "savedquery" ? Rows(view) : throw new AssertFailedException(query.EntityName));
            var result = new DataverseComponentDefinitionOperation().ReadAndResolve(service, "Source", solution.UniqueName,
                CancellationToken.None, _ => { }, counter);
            Assert.IsTrue(result.Membership.Components.All(i => i.Status == IdentityResolutionStatus.Resolved));
            Assert.AreEqual(1, counter.GetExecuteCount("WhoAmI"));
            Assert.AreEqual(1, counter.GetQueryCount("solutioncomponent"));
            Assert.AreEqual(1, counter.GetQueryCount("workflow"));
            Assert.AreEqual(1, counter.GetQueryCount("savedquery"));
            Assert.AreEqual(5, counter.TotalRequests);
            Assert.AreEqual(0, service.WriteCalls);
            Assert.AreEqual(ComponentDefinitionReadStatus.Unsupported, result.Definitions.Single(d => d.Identity.SemanticKind == ComponentSemanticKinds.SavedQuery).Status);
        }

        private static MembershipComparisonPresentation Present(Fixture source, Fixture target) => new MembershipResultPresenter().Create(
            MembershipEnvironmentResult.FromSnapshot("Source", source.Membership, source.Counter.TotalRequests, TimeSpan.Zero),
            MembershipEnvironmentResult.FromSnapshot("Target", target.Membership, target.Counter.TotalRequests, TimeSpan.Zero));
        private static IReadOnlyList<MembershipCompareResult> Compare(Fixture source, Fixture target) =>
            new SolutionMembershipComparer().Compare(source.Membership, target.Membership);
        private static string Table(int type) => type == 29 ? "workflow" : "savedquery";
        private static string Primary(int type) => type == 29 ? "workflowid" : "savedqueryid";
        private static Entity Row(int type, Guid? id = null)
        {
            var key = id ?? Guid.NewGuid();
            var row = new Entity(Table(type), key) { [Primary(type)] = key, ["name"] = "Same context", ["ismanaged"] = false, ["componentstate"] = new OptionSetValue(0) };
            if (type == 29)
            {
                row["category"] = new OptionSetValue(5); row["type"] = new OptionSetValue(1);
                row["mode"] = new OptionSetValue(0); row["modernflowtype"] = new OptionSetValue(0);
                row["primaryentity"] = "account"; row["subprocess"] = false;
            }
            else { row["returnedtypecode"] = "account"; row["querytype"] = 0; }
            return row;
        }

        private sealed class Fixture
        {
            internal readonly int Type;
            internal readonly SolutionIdentity SolutionIdentity = Solution();
            internal readonly List<Entity> Backing;
            internal FakeOrganizationService Service;
            internal DataverseRequestCounter Counter;
            internal readonly List<QueryExpression> Queries = new List<QueryExpression>();
            internal MembershipSnapshot Membership;
            internal ComponentDefinitionSnapshot Definitions;
            internal Func<QueryExpression, EntityCollection> Reply;
            internal Fixture(int type, params Entity[] rows)
            {
                Type = type; Backing = rows.ToList();
                Resolve(rows.Select(r => r.Id).Distinct().ToArray());
            }
            internal void Resolve(Guid[] ids, CancellationToken token = default(CancellationToken))
            {
                Queries.Clear(); Counter = new DataverseRequestCounter();
                Service = ServiceFactory();
                var input = MembershipSnapshot.Complete(SolutionIdentity, ids.Select(id =>
                    new ComponentIdentity(new SolutionComponentRecord(Guid.NewGuid(), Type, id), IdentityResolutionStatus.Unresolved)), DateTimeOffset.UtcNow);
                var context = new DataverseReadContext(Service, SolutionIdentity.Environment, token, Counter);
                Membership = new DataverseComponentIdentityResolver().ResolveSnapshot(context, input, token);
                var requests = Counter.TotalRequests;
                Definitions = new DataverseComponentDefinitionReader().Read(context, Membership, token);
                Assert.AreEqual(requests, Counter.TotalRequests, "Definition reading must reuse rows and add no requests.");
            }
            private FakeOrganizationService ServiceFactory() => MembershipTestData.Service(SolutionIdentity, query =>
            {
                Queries.Add(query); Assert.AreEqual(Table(Type), query.EntityName);
                if (Reply != null) return Reply(query);
                var ids = query.Criteria.Conditions.Single().Values.Cast<Guid>().ToArray();
                return Rows(Backing.Where(r => ids.Contains(r.Id)).ToArray());
            });
        }
    }
}
