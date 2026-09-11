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
    public class WorkflowComponentComparisonTests
    {
        [DataTestMethod, TestCategory("Phase2G4")]
        [DataRow("new_Process", "new_Process")]
        [DataRow("new_Process", "NEW_process")]
        public void UniqueNameMatchesAndRemainsStrongestIdentity(string first, string second)
        {
            var a = Workflow("Source display"); a["uniquename"] = first;
            var b = Workflow("Target display"); b["uniquename"] = second;
            var source = new Fixture(a); var target = new Fixture(b);
            Assert.AreEqual(first, source.Membership.Components.Single().ComparisonKey);
            Assert.AreEqual(ComponentDetailComparisonStatus.Match, Compare(source, target).Single().Status);
        }

        [DataTestMethod, TestCategory("Phase2G4")]
        [DataRow("Set Inspection Date", "Set Inspection Date")]
        [DataRow("Set Inspection Date", "SET INSPECTION DATE")]
        public void BlankUniqueNameUsesCaseInsensitiveFramedSemanticCandidate(string first, string second)
        {
            var a = new Fixture(Workflow(first)); var b = new Fixture(Workflow(second));
            var result = Compare(a, b).Single();
            Assert.AreEqual(MembershipPresence.PresentInBoth, result.Membership.Presence);
            Assert.AreEqual(ComponentDetailComparisonStatus.Match, result.Status);
            StringAssert.StartsWith(a.Membership.Components.Single().ComparisonKey, WorkflowSemanticPolicy.Prefix);
        }

        [DataTestMethod, TestCategory("Phase2G4")]
        [DataRow("category")]
        [DataRow("primaryentity")]
        [DataRow("name")]
        public void SemanticScopeDifferencesDoNotPair(string field)
        {
            var other = Workflow("Same name");
            other[field] = field == "category" ? (object)new OptionSetValue(0) : field == "primaryentity" ? "contact" : "Same-name";
            var results = Compare(new Fixture(Workflow("Same name")), new Fixture(other));
            Assert.AreEqual(2, results.Count);
            Assert.IsTrue(results.Any(r => r.Status == ComponentDetailComparisonStatus.SourceOnly));
            Assert.IsTrue(results.Any(r => r.Status == ComponentDetailComparisonStatus.TargetOnly));
        }

        [DataTestMethod, TestCategory("Phase2G4")]
        [DataRow(4, "businessprocesstype")]
        [DataRow(5, "modernflowtype")]
        [DataRow(6, "uiflowtype")]
        public void ApplicableSubtypeSeparatesCandidates(int category, string subtype)
        {
            var a = Workflow("Same", category); a[subtype] = new OptionSetValue(0);
            var b = Workflow("Same", category); b[subtype] = new OptionSetValue(1);
            var results = Compare(new Fixture(a), new Fixture(b));
            Assert.AreEqual(2, results.Count);
            Assert.IsFalse(results.Any(r => r.Membership.Presence == MembershipPresence.PresentInBoth));
        }

        [TestMethod, TestCategory("Phase2G4")]
        public void DuplicateFallbackIgnoresModeAndDeploymentAsDisambiguators()
        {
            var a = Workflow("Rule"); var b = Workflow("RULE");
            b["mode"] = new OptionSetValue(1); b["ismanaged"] = true;
            var source = new Fixture(a, b); var target = new Fixture(Workflow("Rule"));
            Assert.IsTrue(source.Membership.Components.All(i => i.Status == IdentityResolutionStatus.Ambiguous));
            Assert.IsTrue(source.Membership.Components.All(i => i.ComparisonKey == null));
            Assert.IsTrue(source.Membership.Components.All(i => i.BlockerScope == ResolutionBlockerScope.PortableIdentity));
            Assert.IsTrue(source.Membership.Components.All(i => !string.IsNullOrWhiteSpace(i.BlockerPortableIdentity)));
            var coverage = new MembershipCoverageDiagnosticsBuilder().Build(source.Membership).SemanticKinds.Single(k => k.SemanticKind == ComponentSemanticKinds.Process);
            Assert.AreEqual(2, coverage.Ambiguous);
            Assert.AreEqual(MembershipCoverageStatus.Incomplete, coverage.CoverageStatus);
            Assert.IsTrue(Compare(source, target).All(r => r.Membership.Presence == MembershipPresence.Indeterminate));
        }

        [TestMethod, TestCategory("Phase2G4")]
        public void RepeatedRawObjectDoesNotFabricateSemanticDuplicateButMembershipSafeguardRemains()
        {
            var row = Workflow("Rule"); var fixture = new Fixture(new[] { row }, new[] { row.Id, row.Id });
            Assert.AreEqual(1, fixture.Counter.GetQueryCount("workflow"));
            Assert.IsTrue(fixture.Membership.Components.All(i => i.Status == IdentityResolutionStatus.Resolved));
            var comparison = Compare(fixture, new Fixture(Workflow("Rule")));
            Assert.IsTrue(comparison.All(r => r.Membership.Presence == MembershipPresence.Indeterminate));
            Assert.IsTrue(comparison.Where(r => r.Membership.Source != null).All(r =>
                r.Membership.Source.Status == IdentityResolutionStatus.Ambiguous));
        }

        [DataTestMethod, TestCategory("Phase2G4")]
        [DataRow("name")]
        [DataRow("category")]
        [DataRow("primaryentity")]
        [DataRow("type")]
        public void MissingSemanticEvidenceRemainsUnresolved(string field)
        {
            var row = Workflow("Rule"); row.Attributes.Remove(field);
            var fixture = new Fixture(row);
            Assert.AreEqual(IdentityResolutionStatus.Unresolved, fixture.Membership.Components.Single().Status);
            Assert.IsNull(fixture.Membership.Components.Single().ComparisonKey);
            Assert.AreEqual(ComponentDefinitionReadStatus.Unresolved, fixture.Definitions.Definitions.Single().Status);
        }

        [DataTestMethod, TestCategory("Phase2G4")]
        [DataRow(true)]
        [DataRow(false)]
        public void OneSidedWorkflowRequiresCompleteOppositeCoverage(bool sourceSide)
        {
            var present = new Fixture(Workflow("One sided")); var empty = new Fixture();
            var result = Compare(sourceSide ? present : empty, sourceSide ? empty : present).Single();
            Assert.AreEqual(sourceSide ? ComponentDetailComparisonStatus.SourceOnly : ComponentDetailComparisonStatus.TargetOnly, result.Status);
            var incomplete = Workflow("Incomplete"); incomplete.Attributes.Remove("primaryentity");
            var blocked = Compare(sourceSide ? present : new Fixture(incomplete), sourceSide ? new Fixture(incomplete) : present);
            Assert.IsTrue(blocked.All(r => r.Membership.Presence == MembershipPresence.Indeterminate));
        }

        [DataTestMethod, TestCategory("Phase2G4")]
        [DataRow("workflowidunique")]
        [DataRow("ismanaged")]
        [DataRow("componentstate")]
        [DataRow("statecode")]
        [DataRow("statuscode")]
        [DataRow("ownerid")]
        public void LocalIdsAndStateDifferencesDoNotChangeIdentityOrDefinition(string field)
        {
            var a = Workflow("Rule"); var b = Workflow("Rule");
            b[field] = field == "workflowidunique" ? (object)Guid.NewGuid() : field == "ismanaged" ? true :
                field == "ownerid" ? new EntityReference("systemuser", Guid.NewGuid()) : (object)new OptionSetValue(99);
            Assert.AreNotEqual(a.Id, b.Id);
            Assert.AreEqual(ComponentDetailComparisonStatus.Match, Compare(new Fixture(a), new Fixture(b)).Single().Status);
        }

        [DataTestMethod, TestCategory("Phase2G4")]
        [DataRow("mode")]
        [DataRow("subprocess")]
        public void ConfigurationDifferenceProducesDifferentWithoutChangingMembership(string field)
        {
            var row = Workflow("Rule"); row[field] = field == "mode" ? (object)new OptionSetValue(1) : true;
            var result = Compare(new Fixture(Workflow("Rule")), new Fixture(row)).Single();
            Assert.AreEqual(MembershipPresence.PresentInBoth, result.Membership.Presence);
            Assert.AreEqual(ComponentDetailComparisonStatus.Different, result.Status);
            Assert.AreEqual(field, result.Differences.Single().PropertyName);
        }

        [DataTestMethod, TestCategory("Phase2G4")]
        [DataRow("mode")]
        [DataRow("subprocess")]
        public void IncompleteConfigurationNeverProducesMatch(string field)
        {
            var a = Workflow("Rule"); a.Attributes.Remove(field);
            Assert.AreEqual(ComponentDetailComparisonStatus.Unresolved,
                Compare(new Fixture(a), new Fixture(Workflow("Rule"))).Single().Status);
        }

        [TestMethod, TestCategory("Phase2G4")]
        public void MissingAndDuplicateCorrelationsRemainConservative()
        {
            var id = Guid.NewGuid();
            var missing = new Fixture(new Entity[0], new[] { id });
            Assert.AreEqual(IdentityResolutionStatus.Unresolved, missing.Membership.Components.Single().Status);
            var row = Workflow("Duplicate");
            var duplicate = new Fixture(new[] { row, row }, new[] { row.Id });
            Assert.AreEqual(IdentityResolutionStatus.Ambiguous, duplicate.Membership.Components.Single().Status);
        }

        [DataTestMethod, TestCategory("Phase2G4")]
        [DataRow("paging")]
        [DataRow("primarykey")]
        [DataRow("entity")]
        [DataRow("unexpected")]
        public void MalformedRetrievalCannotProducePartialResult(string condition)
        {
            var row = Workflow("Rule"); var solution = Solution();
            var service = Service(solution, query =>
            {
                if (condition == "primarykey") row["workflowid"] = Guid.NewGuid();
                if (condition == "entity") row.LogicalName = "account";
                if (condition == "unexpected") row.Id = Guid.NewGuid();
                var result = Rows(row); result.MoreRecords = condition == "paging"; return result;
            });
            var snapshot = Input(solution, new[] { row.Id });
            Assert.ThrowsException<InvalidOperationException>(() => new DataverseComponentIdentityResolver()
                .ResolveSnapshot(service, snapshot, CancellationToken.None));
        }

        [TestMethod, TestCategory("Phase2G4")]
        public void FaultsKeepFallbackConservativeAcrossBatches()
        {
            var rows = Enumerable.Range(0, 201).Select(i => Workflow("Rule" + i)).ToArray();
            var solution = Solution(); int calls = 0;
            var service = Service(solution, query =>
            {
                if (++calls == 2) throw new FaultException("Denied");
                var ids = Ids(query); return Rows(rows.Where(r => ids.Contains(r.Id)).ToArray());
            });
            var snapshot = new DataverseComponentIdentityResolver().ResolveSnapshot(service,
                Input(solution, rows.Select(r => r.Id)), CancellationToken.None);
            Assert.AreEqual(201, snapshot.Components.Count);
            Assert.IsTrue(snapshot.Components.All(i => i.Status == IdentityResolutionStatus.Unresolved));
        }

        [TestMethod, TestCategory("Phase2G4")]
        public void CancellationPropagatesWithoutSnapshot()
        {
            var solution = Solution(); var row = Workflow("Rule");
            using (var token = new CancellationTokenSource())
            {
                var service = Service(solution, query => { token.Cancel(); return Rows(row); });
                Assert.ThrowsException<OperationCanceledException>(() => new DataverseComponentIdentityResolver()
                    .ResolveSnapshot(service, Input(solution, new[] { row.Id }), token.Token));
            }
        }

        [DataTestMethod, TestCategory("Phase2G4")]
        [DataRow(19, 1)]
        [DataRow(25, 1)]
        [DataRow(201, 2)]
        public void BatchedRequestShapeAndCacheReuseArePreserved(int count, int expected)
        {
            var fixture = new Fixture(Enumerable.Range(0, count).Select(i => Workflow("Rule" + i)).ToArray());
            Assert.AreEqual(expected, fixture.Counter.GetQueryCount("workflow"));
            Assert.AreEqual(1, fixture.Counter.GetExecuteCount("WhoAmI"));
            Assert.AreEqual(0, fixture.Service.WriteCalls);
            Assert.AreEqual(count, fixture.Definitions.Definitions.Count);
            Assert.IsTrue(fixture.Definitions.Definitions.All(d => d.Status == ComponentDefinitionReadStatus.Available));
            var before = fixture.Counter.TotalRequests;
            new DataverseComponentDefinitionReader().Read(fixture.Context, fixture.Membership, CancellationToken.None);
            Assert.AreEqual(before, fixture.Counter.TotalRequests);
        }

        [TestMethod, TestCategory("Phase2G4")]
        public void ActivationReusesParentDefinitionAlreadyInMembership()
        {
            var parent = Workflow("Rule"); var activation = Workflow("Activation");
            activation["type"] = new OptionSetValue(2);
            activation["parentworkflowid"] = new EntityReference("workflow", parent.Id);
            var fixture = new Fixture(parent, activation);
            Assert.AreEqual(1, fixture.Counter.GetQueryCount("workflow"));
            Assert.IsTrue(fixture.Definitions.Definitions.All(d => d.Status == ComponentDefinitionReadStatus.Available));
            Assert.AreEqual(fixture.Membership.Components[0].ComparisonKey, fixture.Membership.Components[1].ComparisonKey);
        }

        [TestMethod, TestCategory("Phase2G4")]
        public void SemanticEvidenceCannotOverrideUniqueNameAcrossEnvironmentsOrProveAbsence()
        {
            var named = Workflow("Rule"); named["uniquename"] = "new_Strong";
            var results = Compare(new Fixture(Workflow("Rule")), new Fixture(named));
            Assert.AreEqual(2, results.Count);
            Assert.IsTrue(results.All(r => r.Membership.Presence == MembershipPresence.Indeterminate));
            Assert.IsTrue(results.All(r => r.Status == ComponentDetailComparisonStatus.Unresolved));
        }

        [TestMethod, TestCategory("Phase2G4")]
        public void StrongNamesDoNotMatchBySemanticCandidate()
        {
            var a = Workflow("Rule"); a["uniquename"] = "new_First";
            var b = Workflow("Rule"); b["uniquename"] = "new_Second";
            var results = Compare(new Fixture(a), new Fixture(b));
            Assert.AreEqual(2, results.Count);
            Assert.IsTrue(results.All(r => r.Membership.Presence != MembershipPresence.PresentInBoth));
        }

        [TestMethod, TestCategory("Phase2G4")]
        public void StableDiagnosticGroupingRetainsPerRecordAuditEvidence()
        {
            var a = Workflow("One"); var b = Workflow("Two");
            a.Attributes.Remove("primaryentity"); b.Attributes.Remove("primaryentity");
            var fixture = new Fixture(a, b);
            var coverage = new MembershipCoverageDiagnosticsBuilder().Build(fixture.Membership);
            Assert.AreEqual(1, coverage.SemanticKinds.Single(k => k.SemanticKind == ComponentSemanticKinds.Process).DiagnosticGroups.Count);
            Assert.AreEqual(2, coverage.SemanticKinds.Single(k => k.SemanticKind == ComponentSemanticKinds.Process).Unresolved);
            Assert.IsTrue(fixture.Membership.Components[0].DiagnosticEvidence.Any(e => e.Contains(a.Id.ToString("D"))));
            Assert.IsTrue(fixture.Membership.Components[1].DiagnosticEvidence.Any(e => e.Contains(b.Id.ToString("D"))));
        }

        [TestMethod, TestCategory("Phase2G4")]
        public void DetailPresentationShowsChangedConfigurationAndLimitedCoverage()
        {
            var row = Workflow("Rule"); row["mode"] = new OptionSetValue(1);
            var a = new Fixture(Workflow("Rule")); var b = new Fixture(row);
            var membership = new MembershipResultPresenter().Create(
                MembershipEnvironmentResult.FromSnapshot("DEV", a.Membership, a.Counter.TotalRequests, TimeSpan.Zero),
                MembershipEnvironmentResult.FromSnapshot("UAT", b.Membership, b.Counter.TotalRequests, TimeSpan.Zero));
            var presentation = new ComponentDefinitionResultPresenter().Apply(membership, a.Definitions, b.Definitions);
            var result = presentation.Rows.Single();
            Assert.AreEqual("Different", result.DefinitionStatus);
            Assert.AreEqual("mode", result.ChangedProperties);
            Assert.AreEqual(membership.Summary.PresentInBoth, presentation.Summary.PresentInBoth);
            Assert.IsTrue(result.DefinitionDetail.Properties.Single(p => p.PropertyName == "mode").Changed);
            StringAssert.Contains(result.DefinitionDetail.Diagnostic, "Configuration-only");
        }

        [TestMethod, TestCategory("Phase2G4")]
        public void OutputOrderDoesNotDependOnMembershipOrRetrievalOrder()
        {
            var rows = new[] { Workflow("Zulu"), Workflow("Alpha"), Workflow("Middle") };
            var a = new Fixture(rows); var b = new Fixture(rows.Reverse().ToArray());
            var first = Compare(a, b).Select(r => r.Membership.Source.ComparisonKey).ToArray();
            var second = Compare(b, a).Select(r => r.Membership.Source.ComparisonKey).ToArray();
            CollectionAssert.AreEqual(first, second);
        }

        [TestMethod, TestCategory("Phase2G4")]
        public void CoordinatedLiveOperationUsesOneWorkflowQueryAndNoWrites()
        {
            var solution = Solution(); var workflow = Workflow("Rule");
            var member = ComponentRow(solution, 29); member["objectid"] = workflow.Id;
            var counter = new DataverseRequestCounter();
            var service = Service(solution, query => query.EntityName == "solution" ? Rows(SolutionRow(solution)) :
                query.EntityName == "solutioncomponent" ? Rows(member) : Rows(workflow));
            var result = new DataverseComponentDefinitionOperation().ReadAndResolve(service, "DEV", solution.UniqueName,
                CancellationToken.None, progress => { }, counter);
            Assert.AreEqual(ComponentDefinitionReadStatus.Available, result.Definitions.Single().Status);
            Assert.AreEqual(4, counter.TotalRequests);
            Assert.AreEqual(1, counter.GetExecuteCount("WhoAmI"));
            Assert.AreEqual(1, counter.GetQueryCount("workflow"));
            Assert.AreEqual(0, service.WriteCalls);
        }

        [DataTestMethod, TestCategory("Phase2G4")]
        [DataRow("category")]
        [DataRow("primaryentity")]
        [DataRow("businessprocesstype")]
        [DataRow("modernflowtype")]
        [DataRow("uiflowtype")]
        public void StrongIdentityExposesOtherConfigurationDifferences(string field)
        {
            int category = field == "businessprocesstype" ? 4 : field == "modernflowtype" ? 5 : field == "uiflowtype" ? 6 : 2;
            var a = Workflow("Rule", category); var b = Workflow("Rule", category);
            a["uniquename"] = "new_Logical"; b["uniquename"] = "NEW_LOGICAL";
            if (category >= 4) { a[field] = new OptionSetValue(0); b[field] = new OptionSetValue(1); }
            else b[field] = field == "category" ? (object)new OptionSetValue(0) : "contact";
            var result = Compare(new Fixture(a), new Fixture(b)).Single();
            Assert.AreEqual(ComponentDetailComparisonStatus.Different, result.Status);
            Assert.AreEqual(field, result.Differences.Single().PropertyName);
        }

        [TestMethod, TestCategory("Phase2G4")]
        public void PrimaryEntityCaseFollowsLogicalNameSemantics()
        {
            var other = Workflow("Rule"); other["primaryentity"] = "ACCOUNT";
            Assert.AreEqual(ComponentDetailComparisonStatus.Match,
                Compare(new Fixture(Workflow("Rule")), new Fixture(other)).Single().Status);
        }

        [TestMethod, TestCategory("Phase2G4")]
        public void BlankNameAndUnknownSubtypeNeverAcquireFallbackKey()
        {
            var blank = Workflow("   "); var unknown = Workflow("Unknown", 99);
            var missingSubtype = Workflow("Modern", 5);
            var unknownSubtype = Workflow("Modern", 5); unknownSubtype["modernflowtype"] = new OptionSetValue(99);
            var fixture = new Fixture(blank, unknown, missingSubtype, unknownSubtype);
            Assert.IsTrue(fixture.Membership.Components.All(i => i.Status == IdentityResolutionStatus.Unresolved && i.ComparisonKey == null));
        }

        [TestMethod, TestCategory("Phase2G4")]
        public void NamedSiblingDoesNotMakeDuplicateFallbackSafe()
        {
            var named = Workflow("Rule"); named["uniquename"] = "new_Strong";
            var fixture = new Fixture(named, Workflow("RULE"));
            Assert.AreEqual(IdentityResolutionStatus.Resolved, fixture.Membership.Components[0].Status);
            Assert.AreEqual(IdentityResolutionStatus.Ambiguous, fixture.Membership.Components[1].Status);
        }

        [TestMethod, TestCategory("Phase2G4")]
        public void IncompleteSameNameCandidatePreventsFalseMatch()
        {
            var incomplete = Workflow("Rule"); incomplete.Attributes.Remove("category");
            var fixture = new Fixture(Workflow("Rule"), incomplete);
            Assert.IsTrue(fixture.Membership.Components.All(i => i.Status == IdentityResolutionStatus.Unresolved));
            Assert.IsTrue(Compare(fixture, new Fixture(Workflow("Rule"))).All(r => r.Membership.Presence == MembershipPresence.Indeterminate));
        }

        [TestMethod, TestCategory("Phase2G4")]
        public void FramingPreservesMeaningfulPunctuationAndWhitespace()
        {
            string reason;
            var first = Workflow("A:B"); first["primaryentity"] = "C";
            var second = Workflow("A"); second["primaryentity"] = "B:C";
            Assert.AreNotEqual(WorkflowSemanticPolicy.Candidate(first, out reason), WorkflowSemanticPolicy.Candidate(second, out reason));
            Assert.AreNotEqual(WorkflowSemanticPolicy.Candidate(Workflow(" Rule "), out reason),
                WorkflowSemanticPolicy.Candidate(Workflow("Rule"), out reason));
        }

        [TestMethod, TestCategory("Phase2G4")]
        public void PayloadAndIdentityFieldsRemainOutsideConfigurationContract()
        {
            var contract = ComponentDefinitionContractCatalog.For(ComponentSemanticKinds.Process);
            CollectionAssert.AreEquivalent(new[] { "type", "category", "primaryentity", "mode", "subprocess",
                "businessprocesstype", "modernflowtype", "uiflowtype" }, contract.ComparableProperties.ToArray());
            foreach (var field in new[] { "xaml", "clientdata", "definition", "workflowid", "workflowidunique", "name", "uniquename", "ismanaged", "componentstate" })
                Assert.IsFalse(contract.ComparableProperties.Contains(field));
            Assert.AreSame(StringComparer.OrdinalIgnoreCase, contract.IdentityComparer);
        }

        private static Entity Workflow(string name, int category = 2) => new Entity("workflow", Guid.NewGuid())
        {
            ["name"] = name, ["type"] = new OptionSetValue(1), ["category"] = new OptionSetValue(category),
            ["primaryentity"] = "account", ["mode"] = new OptionSetValue(0), ["subprocess"] = false,
            ["workflowidunique"] = Guid.NewGuid(), ["ismanaged"] = false,
            ["statecode"] = new OptionSetValue(0), ["statuscode"] = new OptionSetValue(1),
            ["componentstate"] = new OptionSetValue(0)
        };

        private static MembershipSnapshot Input(SolutionIdentity solution, IEnumerable<Guid> ids) =>
            MembershipSnapshot.Complete(solution, ids.Select(id => new ComponentIdentity(
                new SolutionComponentRecord(Guid.NewGuid(), 29, id), IdentityResolutionStatus.Unresolved)), DateTimeOffset.UtcNow);
        private static Guid[] Ids(QueryExpression query) => query.Criteria.Conditions.Single().Values.Cast<Guid>().ToArray();
        private static IReadOnlyList<ComponentDetailCompareResult> Compare(Fixture a, Fixture b) =>
            new ComponentDetailComparer().Compare(new SolutionMembershipComparer().Compare(a.Membership, b.Membership), a.Definitions, b.Definitions);

        private sealed class Fixture
        {
            internal readonly DataverseRequestCounter Counter = new DataverseRequestCounter();
            internal readonly FakeOrganizationService Service;
            internal readonly DataverseReadContext Context;
            internal readonly MembershipSnapshot Membership;
            internal readonly ComponentDefinitionSnapshot Definitions;
            internal Fixture(params Entity[] rows) : this(rows, rows.Select(r => r.Id).Distinct().ToArray()) { }
            internal Fixture(Entity[] rows, Guid[] membershipIds)
            {
                var solution = Solution();
                Service = MembershipTestData.Service(solution, query =>
                {
                    Assert.AreEqual("workflow", query.EntityName);
                    CollectionAssert.AreEquivalent(WorkflowSemanticPolicy.Columns, query.ColumnSet.Columns.ToArray());
                    Assert.AreEqual("workflowid", query.Criteria.Conditions.Single().AttributeName);
                    Assert.AreEqual(ConditionOperator.In, query.Criteria.Conditions.Single().Operator);
                    var ids = Ids(query);
                    Assert.IsTrue(ids.Length <= 200);
                    Assert.AreEqual(ids.Distinct().Count(), ids.Length);
                    CollectionAssert.AreEqual(ids.OrderBy(id => id).ToArray(), ids);
                    return Rows(rows.Where(r => ids.Contains(r.Id)).ToArray());
                });
                Context = new DataverseReadContext(Service, solution.Environment, CancellationToken.None, Counter);
                Membership = new DataverseComponentIdentityResolver().ResolveSnapshot(Context, Input(solution, membershipIds), CancellationToken.None);
                Definitions = new DataverseComponentDefinitionReader().Read(Context, Membership, CancellationToken.None);
            }
        }
    }
}
