using System;
using System.Collections.Generic;
using System.Linq;
using System.ServiceModel;
using System.Threading;
using D365SolutionComparer.Infrastructure;
using D365SolutionComparer.Models.ComponentDetails;
using D365SolutionComparer.Models.Identity;
using D365SolutionComparer.Models.Membership;
using D365SolutionComparer.Services.ComponentDetails;
using D365SolutionComparer.Services.Membership;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using static D365SolutionComparer.Tests.MembershipTestData;

namespace D365SolutionComparer.Tests
{
    [TestClass]
    public class Type59MembershipResolutionTests
    {
        private const string Category = "Phase2GType59Production";
        private const string EntityName = "savedqueryvisualization";
        private const string PrimaryId = "savedqueryvisualizationid";
        private const string Prefix = "savedqueryvisualization:v1:savedqueryvisualizationid:36:";
        private static readonly string[] Columns = { PrimaryId, "name", "primaryentitytypecode", "type",
            "charttype", "savedqueryvisualizationidunique", "componentstate", "ismanaged" };

        [DataTestMethod, TestCategory(Category)]
        [DataRow("savedqueryvisualizationidunique")] [DataRow("ismanaged")] [DataRow("componentstate")]
        [DataRow("name")] [DataRow("primaryentitytypecode")] [DataRow("type")] [DataRow("charttype")]
        [DataRow("datadescription")] [DataRow("presentationdescription")]
        public void AuditAndDefinitionDifferencesCannotChangePrimaryIdMatch(string field)
        {
            var a = Chart(); var b = Chart(a.Id);
            a[field] = field == "ismanaged" ? (object)false : "Source evidence";
            b[field] = field == "ismanaged" ? (object)true : "Target evidence";
            var source = new Fixture(a); var target = new Fixture(b);
            var result = Compare(source, target).Single();
            Assert.AreEqual(MembershipPresence.PresentInBoth, result.Presence);
            Assert.AreEqual(Prefix + a.Id.ToString("D"), result.Source.ComparisonKey);
            Assert.AreEqual(result.Source.ComparisonKey, result.Target.ComparisonKey);
            Assert.AreEqual("Not Compared", Present(source, target).Rows.Single().DefinitionStatus);
            Assert.IsNull(ComponentDefinitionContractCatalog.For(ComponentSemanticKinds.SavedQueryVisualization));
        }

        [TestMethod, TestCategory(Category)]
        public void SameSemanticCandidatesAndDefinitionHashesNeverMatchDifferentPrimaryIds()
        {
            var a = Chart(); var b = Chart();
            a["datadescription"] = b["datadescription"] = "<data/>";
            a["presentationdescription"] = b["presentationdescription"] = "<presentation/>";
            var results = Compare(new Fixture(a), new Fixture(b));
            Assert.AreEqual(2, results.Count);
            Assert.AreEqual(1, results.Count(r => r.Presence == MembershipPresence.OnlyInSource));
            Assert.AreEqual(1, results.Count(r => r.Presence == MembershipPresence.OnlyInTarget));
            Assert.IsFalse(results.Any(r => r.Presence == MembershipPresence.PresentInBoth));
        }

        [TestMethod, TestCategory(Category)]
        public void CandidateAAndBCollisionsDoNotCreateProductionAmbiguity()
        {
            var a = Chart(); var b = Chart();
            var source = new Fixture(a, b); var target = new Fixture(Chart(a.Id), Chart(b.Id));
            var results = Compare(source, target);
            Assert.IsTrue(source.Membership.Components.All(i => i.Status == IdentityResolutionStatus.Resolved));
            Assert.AreEqual(2, results.Count(r => r.Presence == MembershipPresence.PresentInBoth));
            Assert.IsTrue(results.All(r => r.Source.ComparisonKey.StartsWith(Prefix, StringComparison.Ordinal)));
        }

        [DataTestMethod, TestCategory(Category)]
        [DataRow(true)] [DataRow(false)]
        public void OneSidedChartNeedsCompleteOppositeCoverageAndShowsNotCompared(bool sourceSide)
        {
            var present = new Fixture(Chart()); var empty = new Fixture();
            var source = sourceSide ? present : empty; var target = sourceSide ? empty : present;
            var result = Compare(source, target).Single();
            Assert.AreEqual(sourceSide ? MembershipPresence.OnlyInSource : MembershipPresence.OnlyInTarget, result.Presence);
            Assert.AreEqual(MembershipAbsenceEvidence.CompleteResolvedInventory, result.AbsenceEvidence);
            var row = Present(source, target).Rows.Single();
            Assert.AreEqual("Not Compared", row.DefinitionStatus);
            Assert.AreEqual("Saved Query Visualization", row.ComponentKind);
            var bucket = Coverage(empty);
            Assert.AreEqual(0, bucket.TotalCandidates); Assert.AreEqual(MembershipCoverageStatus.Complete, bucket.CoverageStatus);
        }

        [DataTestMethod, TestCategory(Category)]
        [DataRow(true, "missing")] [DataRow(false, "missing")]
        [DataRow(true, "duplicate")] [DataRow(false, "duplicate")]
        [DataRow(true, "incomplete")] [DataRow(false, "incomplete")]
        public void IncompleteOppositeCoverageCannotProduceOneSidedFinding(bool sourceSide, string defect)
        {
            var present = new Fixture(Chart()); var row = Chart(); var other = new Fixture();
            other.Reply = query => defect == "missing" ? Rows() : defect == "duplicate"
                ? Rows(row, Chart(row.Id)) : new EntityCollection(new[] { row }.ToList()) { MoreRecords = true };
            other.Resolve(new[] { row.Id });
            var results = sourceSide ? Compare(present, other) : Compare(other, present);
            Assert.IsTrue(results.All(r => r.Presence == MembershipPresence.Indeterminate && r.AbsenceEvidence == MembershipAbsenceEvidence.None));
            Assert.AreEqual(MembershipCoverageStatus.Incomplete, Coverage(other).CoverageStatus);
        }

        [DataTestMethod, TestCategory(Category)]
        [DataRow("blank")] [DataRow("absent")] [DataRow("string")] [DataRow("conflicting")]
        [DataRow("emptyEntityId")] [DataRow("missing")] [DataRow("paged")] [DataRow("fault")]
        public void InvalidPrimaryOrBackingEvidenceStaysUnresolved(string defect)
        {
            var row = Chart(); var requestedId = row.Id; var fixture = new Fixture();
            if (defect == "blank") row[PrimaryId] = Guid.Empty;
            if (defect == "absent") row.Attributes.Remove(PrimaryId);
            if (defect == "string") row[PrimaryId] = row.Id.ToString("D");
            if (defect == "conflicting") row[PrimaryId] = Guid.NewGuid();
            if (defect == "emptyEntityId") row.Id = Guid.Empty;
            fixture.Reply = query =>
            {
                if (defect == "fault") throw new FaultException("Test fault");
                if (defect == "missing") return Rows();
                return new EntityCollection(new[] { row }.ToList()) { MoreRecords = defect == "paged" };
            };
            fixture.Resolve(new[] { requestedId });
            var identity = fixture.Membership.Components.Single();
            Assert.AreEqual(IdentityResolutionStatus.Unresolved, identity.Status); Assert.IsNull(identity.ComparisonKey);
            Assert.AreEqual(ComponentSemanticKinds.SavedQueryVisualization, identity.SemanticKind);
            Assert.AreEqual("Unresolved", Present(fixture, new Fixture()).Rows.Single().DefinitionStatus);
            Assert.IsTrue(Compare(fixture, new Fixture(Chart())).All(r => r.Presence == MembershipPresence.Indeterminate));
        }

        [DataTestMethod, TestCategory(Category)]
        [DataRow(false)] [DataRow(true)]
        public void MissingOrEmptyObjectIdDoesNotQueryChartTable(bool empty)
        {
            var fixture = new Fixture(); fixture.Resolve(new Guid?[] { empty ? (Guid?)Guid.Empty : null });
            Assert.AreEqual(IdentityResolutionStatus.Unresolved, fixture.Membership.Components.Single().Status);
            Assert.AreEqual(0, fixture.Counter.GetQueryCount(EntityName));
        }

        [TestMethod, TestCategory(Category)]
        public void DuplicateBackingRowsStayAmbiguousEvenWithDifferentAuditValues()
        {
            var a = Chart(); var b = Chart(a.Id); b["name"] = "Other chart";
            var fixture = new Fixture(a, b);
            Assert.AreEqual(IdentityResolutionStatus.Ambiguous, fixture.Membership.Components.Single().Status);
            Assert.IsNull(fixture.Membership.Components.Single().ComparisonKey);
            Assert.AreEqual("Ambiguous", Present(fixture, new Fixture()).Rows.Single().DefinitionStatus);
            Assert.AreEqual(1, Coverage(fixture).Ambiguous);
        }

        [TestMethod, TestCategory(Category)]
        public void RepeatedRawPrimaryIdentityIsAmbiguousNotCollapsedAndQueryIsDeduplicated()
        {
            var row = Chart(); var fixture = new Fixture(row);
            fixture.Resolve(new[] { row.Id, row.Id });
            Assert.AreEqual(2, fixture.Membership.Components.Count);
            Assert.IsTrue(fixture.Membership.Components.All(i => i.Status == IdentityResolutionStatus.Ambiguous && i.ComparisonKey == null));
            Assert.AreEqual(1, fixture.Queries.Single().Criteria.Conditions.Single().Values.Count);
            Assert.AreEqual(2, Coverage(fixture).Ambiguous);
            Assert.IsTrue(Compare(fixture, new Fixture()).All(r => r.Presence == MembershipPresence.Indeterminate));
        }

        [TestMethod, TestCategory(Category)]
        public void PrimaryIdAloneCanResolveWithoutOptionalSemanticOrAuditFields()
        {
            var row = new Entity(EntityName, Guid.NewGuid()); row[PrimaryId] = row.Id;
            var fixture = new Fixture(row);
            Assert.AreEqual(IdentityResolutionStatus.Resolved, fixture.Membership.Components.Single().Status);
            var coverage = Coverage(fixture);
            Assert.AreEqual(1, coverage.Resolved); Assert.AreEqual(0, coverage.Unsupported);
            Assert.AreEqual(0, coverage.DiagnosticGroups.Count);
            StringAssert.Contains(fixture.Membership.Components.Single().Diagnostic, "primary-ID identity");
            StringAssert.Contains(fixture.Membership.Components.Single().DiagnosticEvidence.First(), "diagnostic context only");
        }

        [DataTestMethod, TestCategory(Category)]
        [DataRow(200, 1)] [DataRow(201, 2)]
        public void ProductionUsesExistingBatchesColumnsAndNoDefinitionOrXmlRequests(int count, int batches)
        {
            var rows = Enumerable.Range(0, count).Select(_ => Chart()).ToArray(); var fixture = new Fixture(rows);
            Assert.AreEqual(batches, fixture.Counter.GetQueryCount(EntityName));
            Assert.AreEqual(1, fixture.Counter.GetExecuteCount("WhoAmI"));
            Assert.AreEqual(batches + 1, fixture.Counter.TotalRequests); Assert.AreEqual(0, fixture.Service.WriteCalls);
            foreach (var query in fixture.Queries)
            {
                CollectionAssert.AreEquivalent(Columns, query.ColumnSet.Columns.ToArray());
                Assert.IsFalse(query.ColumnSet.AllColumns);
                Assert.IsFalse(query.ColumnSet.Columns.Contains("datadescription"));
                Assert.IsFalse(query.ColumnSet.Columns.Contains("presentationdescription"));
                var filter = query.Criteria.Conditions.Single();
                Assert.AreEqual(PrimaryId, filter.AttributeName); Assert.AreEqual(ConditionOperator.In, filter.Operator);
                Assert.IsTrue(filter.Values.Count <= 200 && filter.Values.All(v => v is Guid));
            }
            CollectionAssert.AreEqual(rows.Select(r => r.Id).OrderBy(id => id).ToArray(),
                fixture.Queries.SelectMany(q => q.Criteria.Conditions.Single().Values.Cast<Guid>()).ToArray());
            Assert.IsTrue(fixture.Definitions.Definitions.All(d => d.Status == ComponentDefinitionReadStatus.Unsupported));
        }

        [TestMethod, TestCategory(Category)]
        public void CoordinatedOperationReusesExistingChartDiagnosticRead()
        {
            var solution = Solution(); var chart = Chart(); var member = ComponentRow(solution, 59); member["objectid"] = chart.Id;
            var counter = new DataverseRequestCounter();
            var service = Service(solution, query => query.EntityName == "solution" ? Rows(SolutionRow(solution)) :
                query.EntityName == "solutioncomponent" ? Rows(member) : query.EntityName == EntityName ? Rows(chart) :
                throw new AssertFailedException(query.EntityName));
            var result = new DataverseComponentDefinitionOperation().ReadAndResolve(service, "Source", solution.UniqueName,
                CancellationToken.None, _ => { }, counter);
            Assert.AreEqual(IdentityResolutionStatus.Resolved, result.Membership.Components.Single().Status);
            Assert.AreEqual(1, counter.GetExecuteCount("WhoAmI")); Assert.AreEqual(1, counter.GetQueryCount("solutioncomponent"));
            Assert.AreEqual(1, counter.GetQueryCount(EntityName)); Assert.AreEqual(4, counter.TotalRequests);
            Assert.AreEqual(0, service.WriteCalls);
            Assert.AreEqual(ComponentDefinitionReadStatus.Unsupported, result.Definitions.Single().Status);
        }

        [TestMethod, TestCategory(Category)]
        public void CancellationPropagates()
        {
            var fixture = new Fixture(); using (var cancellation = new CancellationTokenSource())
            {
                fixture.Reply = query => { cancellation.Cancel(); return Rows(); };
                Assert.ThrowsException<OperationCanceledException>(() => fixture.Resolve(new[] { Guid.NewGuid() }, cancellation.Token));
            }
        }

#if DEBUG
        [TestMethod, TestCategory(Category)]
        public void LifecycleEvidenceCannotOverrideResolvedProductionMembership()
        {
            var a = Chart(); var b = Chart(a.Id); b["name"] = "Renamed chart";
            a["datadescription"] = "<data/>"; b["datadescription"] = "<different/>";
            var source = new Fixture(a); var target = new Fixture(b);
            var before = Compare(source, target).Single();
            var report = new Type59EvidenceCollector().Capture(source.Service, source.Membership, "1",
                target.Service, target.Membership, "2", CancellationToken.None);
            Assert.AreEqual(0, report.Pairs.Count, "Semantic evidence differs while production primary-ID membership still matches.");
            var after = Compare(source, target).Single();
            Assert.AreEqual(MembershipPresence.PresentInBoth, after.Presence);
            Assert.AreSame(before.Source, after.Source); Assert.AreSame(before.Target, after.Target);
            Assert.AreEqual(before.Source.ComparisonKey, after.Source.ComparisonKey);
            Assert.AreEqual(0, source.Service.WriteCalls + target.Service.WriteCalls);
        }
#endif

        private static MembershipCoverageBucket Coverage(Fixture fixture) => new MembershipCoverageDiagnosticsBuilder()
            .Build(fixture.Membership).SemanticKinds.Single(k => k.SemanticKind == ComponentSemanticKinds.SavedQueryVisualization);
        private static IReadOnlyList<MembershipCompareResult> Compare(Fixture source, Fixture target) =>
            new SolutionMembershipComparer().Compare(source.Membership, target.Membership);
        private static MembershipComparisonPresentation Present(Fixture source, Fixture target)
        {
            var membership = new MembershipResultPresenter().Create(
                MembershipEnvironmentResult.FromSnapshot("Source", source.Membership, source.Counter.TotalRequests, TimeSpan.Zero),
                MembershipEnvironmentResult.FromSnapshot("Target", target.Membership, target.Counter.TotalRequests, TimeSpan.Zero));
            return new ComponentDefinitionResultPresenter().Apply(membership, source.Definitions, target.Definitions);
        }
        private static Entity Chart(Guid? id = null)
        {
            var key = id ?? Guid.NewGuid();
            return new Entity(EntityName, key) { [PrimaryId] = key, ["savedqueryvisualizationidunique"] = Guid.NewGuid(),
                ["name"] = "Same chart", ["primaryentitytypecode"] = "account", ["type"] = new OptionSetValue(0),
                ["charttype"] = new OptionSetValue(0), ["componentstate"] = new OptionSetValue(0), ["ismanaged"] = false };
        }
        private sealed class Fixture
        {
            internal readonly SolutionIdentity SolutionIdentity = Solution();
            internal readonly List<Entity> Backing;
            internal readonly List<QueryExpression> Queries = new List<QueryExpression>();
            internal FakeOrganizationService Service;
            internal DataverseRequestCounter Counter;
            internal MembershipSnapshot Membership;
            internal ComponentDefinitionSnapshot Definitions;
            internal Func<QueryExpression, EntityCollection> Reply;
            internal Fixture(params Entity[] rows)
            {
                Backing = rows.ToList(); Resolve(rows.Select(r => r.Id).Distinct().ToArray());
            }
            internal void Resolve(Guid[] ids, CancellationToken token = default(CancellationToken)) =>
                Resolve(ids.Select(id => (Guid?)id).ToArray(), token);
            internal void Resolve(Guid?[] ids, CancellationToken token = default(CancellationToken))
            {
                Queries.Clear(); Counter = new DataverseRequestCounter();
                Service = MembershipTestData.Service(SolutionIdentity, query =>
                {
                    Queries.Add(query); Assert.AreEqual(EntityName, query.EntityName);
                    if (Reply != null) return Reply(query);
                    var keys = query.Criteria.Conditions.Single().Values.Cast<Guid>().ToArray();
                    return Rows(Backing.Where(r => keys.Contains(r.Id)).ToArray());
                });
                var input = MembershipSnapshot.Complete(SolutionIdentity, ids.Select(id => new ComponentIdentity(
                    new SolutionComponentRecord(Guid.NewGuid(), 59, id), IdentityResolutionStatus.Unresolved)), DateTimeOffset.UtcNow);
                var context = new DataverseReadContext(Service, SolutionIdentity.Environment, token, Counter);
                Membership = new DataverseComponentIdentityResolver().ResolveSnapshot(context, input, token);
                var requests = Counter.TotalRequests;
                Definitions = new DataverseComponentDefinitionReader().Read(context, Membership, token);
                Assert.AreEqual(requests, Counter.TotalRequests, "No chart definition query is permitted.");
            }
        }
    }
}
