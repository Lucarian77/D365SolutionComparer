using System;
using System.Collections.Generic;
using System.Linq;
using System.ServiceModel;
using System.Threading;
using D365SolutionComparer.Infrastructure;
using D365SolutionComparer.Models.Membership;
using D365SolutionComparer.Services.Membership;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Query;

namespace D365SolutionComparer.Tests
{
    [TestClass]
    public class AppSettingEvidenceTests
    {
        [TestMethod]
        [TestCategory("Phase2G6A")]
        public void CandidateCompositeIdentityComparisonIsCaseInsensitive()
        {
            var source = Report("DEV", Candidate("App.One+Setting.A"));
            var target = Report("UAT", Candidate("app.one+setting.a"));
            var result = AppSettingEvidenceComparison.Create(source, target);
            Assert.AreEqual(1, result.Candidates.Count);
            Assert.AreEqual("CandidateStable", result.Candidates[0].Outcome);
        }

        [TestMethod]
        [TestCategory("Phase2G6A")]
        public void DuplicateCompositeCandidatesRemainAmbiguous()
        {
            var source = Report("DEV", Candidate("app.one+setting.a"), Candidate("APP.ONE+SETTING.A"));
            var target = Report("UAT", Candidate("app.one+setting.a"));
            var result = AppSettingEvidenceComparison.Create(source, target);
            Assert.AreEqual(1, result.Candidates.Count);
            Assert.AreEqual("AmbiguousCandidate", result.Candidates[0].Outcome);
            Assert.AreEqual(2, result.Candidates[0].SourceCount);
        }

        [TestMethod]
        [TestCategory("Phase2G6A")]
        public void OneSidedCandidatesRemainDiagnosticOnly()
        {
            var result = AppSettingEvidenceComparison.Create(Report("DEV", Candidate("app.one+setting.a")),
                Report("UAT"));
            Assert.AreEqual("DEVOnlyCandidate", result.Candidates[0].Outcome);
        }

        [TestMethod]
        [TestCategory("Phase2G6A")]
        public void IncompleteCandidatesAreExcludedFromCompositeCorrelation()
        {
            var incomplete = new AppSettingCandidateEvidence("DEV", "edu", Guid.NewGuid(), Guid.NewGuid(),
                10094, "App Setting", Guid.NewGuid(), null, null, null, "test", null,
                AppSettingEvidenceState.Unresolved, string.Empty, "", string.Empty, "", "", "", "missing");
            var result = AppSettingEvidenceComparison.Create(Report("DEV", incomplete), Report("UAT"));
            Assert.AreEqual(0, result.Candidates.Count);
        }

        [TestMethod]
        [TestCategory("Phase2G6A")]
        public void EvidenceModelRetainsRawComponentAndObjectIds()
        {
            var solutionComponentId = Guid.NewGuid();
            var objectId = Guid.NewGuid();
            var candidate = new AppSettingCandidateEvidence("DEV", "edu", Guid.NewGuid(), solutionComponentId,
                10094, "App Setting", objectId, 0, Guid.NewGuid(), false, "label", null,
                AppSettingEvidenceState.Unresolved, "", "", "", "", "", "", "evidence");
            Assert.AreEqual(solutionComponentId, candidate.SolutionComponentId);
            Assert.AreEqual(objectId, candidate.ObjectId.Value);
            Assert.AreEqual(10094, candidate.ComponentType);
        }

        [TestMethod]
        [TestCategory("Phase2G6A")]
        public void EvidenceComparisonDoesNotCreateMembershipStatuses()
        {
            var result = AppSettingEvidenceComparison.Create(Report("DEV", Candidate("app.one+setting.a")),
                Report("UAT", Candidate("app.one+setting.a")));
            Assert.AreEqual("CandidateStable", result.Candidates[0].Outcome);
            Assert.AreEqual("Evidence Only", "Evidence Only");
        }

        [TestMethod]
        [TestCategory("Phase2G6A")]
        public void SolutionComponentEvidenceQueryUsesOnlySupportedAttributes()
        {
            var query = DataverseSolutionMembershipEvidenceQuery();
            CollectionAssert.AreEqual(new[] { "solutioncomponentid", "solutionid", "componenttype", "objectid",
                "rootcomponentbehavior", "rootsolutioncomponentid", "ismetadata" }, query.ColumnSet.Columns.ToArray());
            CollectionAssert.DoesNotContain(query.ColumnSet.Columns, "componentidunique");
            CollectionAssert.DoesNotContain(query.ColumnSet.Columns, "componentstate");
            CollectionAssert.DoesNotContain(query.ColumnSet.Columns, "ismanaged");
            CollectionAssert.Contains(query.ColumnSet.Columns, "solutioncomponentid");
            CollectionAssert.Contains(query.ColumnSet.Columns, "componenttype");
            CollectionAssert.Contains(query.ColumnSet.Columns, "objectid");
            CollectionAssert.Contains(query.ColumnSet.Columns, "solutionid");
            CollectionAssert.Contains(query.ColumnSet.Columns, "ismetadata");
            CollectionAssert.Contains(query.ColumnSet.Columns, "rootsolutioncomponentid");
            CollectionAssert.Contains(query.ColumnSet.Columns, "rootcomponentbehavior");
        }

        [DataTestMethod]
        [TestCategory("Phase2G6A")]
        [DataRow(10075, null, "AppSetting", "appsetting")]
        [DataRow(10075, "", "AppSetting", "appsetting")]
        [DataRow(54321, "Unrelated label", "AppSetting", "appsetting")]
        [DataRow(54321, "", "aPpSeTtInG", "APPSETTING")]
        public void RegisteredFamilySelectsLiveRowsAndExecutesFullCorrelationWithoutLabelEvidence(
            int type, string label, string definitionName, string primaryEntity)
        {
            var fixture = new CaptureFixture(type, label,
                new SolutionComponentDefinitionIdentity(type, definitionName, primaryEntity));
            var report = fixture.Capture();
            var candidate = report.Candidates.Single();
            Assert.AreEqual(type, candidate.ComponentType);
            Assert.AreEqual(label ?? string.Empty, candidate.FormattedLabel);
            Assert.AreEqual(fixture.ObjectId, candidate.ObjectId);
            Assert.AreEqual(fixture.MembershipRow.Id, candidate.SolutionComponentId);
            Assert.AreEqual(AppSettingEvidenceState.Confirmed, candidate.State);
            Assert.AreEqual(fixture.ObjectId, candidate.Correlations.Single(item =>
                item.EntityLogicalName == "appsetting").Records.Single().RecordId);
            Assert.AreEqual(fixture.DefinitionId.ToString("D"), candidate.SettingDefinitionId);
            Assert.AreEqual("new_setting", candidate.SettingDefinitionName);
            Assert.AreEqual(fixture.AppId.ToString("D"), candidate.ParentAppModuleId);
            Assert.AreEqual("new_app", candidate.ParentAppModuleUniqueName);
            Assert.AreEqual("new_app+new_setting", candidate.CandidateCompositeIdentity);
            StringAssert.Contains(candidate.CandidateReason, "Registered solutioncomponentdefinition");
            Assert.AreEqual(1, fixture.Counter.GetQueryCount("appsetting"));
            Assert.AreEqual(0, fixture.Counter.GetQueryCount("solutioncomponentdefinition"));
            Assert.AreEqual(10, report.Requests.Total);
            Assert.AreEqual(fixture.Service.Calls + fixture.Service.ExecuteCalls, report.Requests.Total);
            Assert.AreEqual(0, fixture.Service.WriteCalls);
            Assert.IsTrue(fixture.Snapshot.Components.All(item =>
                item.Status == IdentityResolutionStatus.Unsupported && item.ComparisonKey == null));
        }

        [DataTestMethod]
        [TestCategory("Phase2G6A")]
        [DataRow(10075, null, null, null)]
        [DataRow(10075, "AppSetting", null, null)]
        [DataRow(54321, "App Setting", null, null)]
        [DataRow(10075, "AppSetting", "AppSetting", "otherentity")]
        [DataRow(10075, "AppSetting", "OtherDefinition", "appsetting")]
        [DataRow(10075, "AppSetting", "AppSetting", "")]
        public void AbsentOrNonqualifyingRegistrationNeverSelectsEvenWithSuggestiveLabels(
            int type, string label, string name, string primaryEntity)
        {
            var fixture = new CaptureFixture(type, label, name == null ? null :
                new SolutionComponentDefinitionIdentity(type, name, primaryEntity));
            var report = fixture.Capture();
            Assert.AreEqual(0, report.Candidates.Count);
            Assert.AreEqual(0, fixture.Counter.GetQueryCount("appsetting"));
            Assert.AreEqual(0, fixture.Counter.GetQueryCount("solutioncomponentdefinition"));
            Assert.AreEqual(0, fixture.Counter.GetExecuteCount("RetrieveEntity"));
            Assert.AreEqual(2, report.Requests.Total); // environment verification and membership page only
            Assert.AreEqual(label ?? string.Empty, report.ComponentTypes.Single().FormattedLabel);
            Assert.AreEqual(0, fixture.Service.WriteCalls);
        }

        [TestMethod]
        [TestCategory("Phase2G6A")]
        public void DuplicateEquivalentRegistrationsSelectTypeOnceAndDoNotRepeatQueries()
        {
            var fixture = new CaptureFixture(54321, null,
                new SolutionComponentDefinitionIdentity(54321, "AppSetting", "appsetting"),
                new SolutionComponentDefinitionIdentity(54321, "AppSetting", "appsetting"),
                new SolutionComponentDefinitionIdentity(54321, "appsetting", "APPSETTING"));
            var report = fixture.Capture();
            Assert.AreEqual(1, report.Candidates.Count);
            Assert.AreEqual(1, fixture.Counter.GetQueryCount("appsetting"));
            Assert.AreEqual(10, report.Requests.Total);
        }

        [DataTestMethod]
        [TestCategory("Phase2G6A")]
        [DataRow("OtherDefinition", "appsetting", false)]
        [DataRow("AppSetting", "otherentity", false)]
        [DataRow("OtherDefinition", "appsetting", true)]
        public void ConflictingRegistrationsAreDiagnosedAndNeverSelectAnArbitraryFamily(
            string name, string primaryEntity, bool reverse)
        {
            var registrations = new[] { new SolutionComponentDefinitionIdentity(54321, "AppSetting", "appsetting"),
                new SolutionComponentDefinitionIdentity(54321, name, primaryEntity) };
            var fixture = new CaptureFixture(54321, "AppSetting",
                reverse ? registrations.Reverse().ToArray() : registrations);
            var report = fixture.Capture();
            Assert.AreEqual(0, report.Candidates.Count);
            Assert.AreEqual(0, fixture.Counter.GetQueryCount("appsetting"));
            Assert.IsTrue(report.Diagnostics.Any(item => item.Contains("Ambiguous registered definitions") &&
                item.Contains("54321")));
            Assert.AreEqual(2, report.Requests.Total);
        }

        [TestMethod]
        [TestCategory("Phase2G6A")]
        public void RegisteredTypeSelectionIsDistinctSortedAndExactWithoutTrimming()
        {
            var registrations = new[] { new SolutionComponentDefinitionIdentity(54321, "AppSetting", "appsetting"),
                new SolutionComponentDefinitionIdentity(10075, "AppSetting", "appsetting"),
                new SolutionComponentDefinitionIdentity(54321, "AppSetting", "appsetting"),
                new SolutionComponentDefinitionIdentity(60000, " AppSetting", "appsetting"),
                new SolutionComponentDefinitionIdentity(60001, "AppSetting", "appsetting "), null };
            CollectionAssert.AreEqual(new[] { 10075, 54321 },
                DataverseAppSettingEvidenceOperation.FindCandidateComponentTypes(registrations).ToArray());
        }

        [DataTestMethod]
        [TestCategory("Phase2G6A")]
        [DataRow("missing", AppSettingEvidenceState.Unresolved)]
        [DataRow("duplicate", AppSettingEvidenceState.Ambiguous)]
        [DataRow("fault", AppSettingEvidenceState.Faulted)]
        [DataRow("incomplete", AppSettingEvidenceState.Faulted)]
        [DataRow("unavailable", AppSettingEvidenceState.Unavailable)]
        public void RegisteredDiscoveryPreservesConservativeBackingReadOutcomes(
            string outcome, AppSettingEvidenceState expected)
        {
            var fixture = new CaptureFixture(54321, null,
                new SolutionComponentDefinitionIdentity(54321, "AppSetting", "appsetting"));
            fixture.Outcome = outcome;
            var candidate = fixture.Capture().Candidates.Single();
            Assert.AreEqual(expected, candidate.State);
            Assert.AreEqual(string.Empty, candidate.CandidateCompositeIdentity);
            Assert.AreEqual(0, fixture.Service.WriteCalls);
        }

        [TestMethod]
        [TestCategory("Phase2G6A")]
        public void RegisteredDiscoveryPreservesCancellationWithoutReturningPartialReport()
        {
            var fixture = new CaptureFixture(54321, null,
                new SolutionComponentDefinitionIdentity(54321, "AppSetting", "appsetting"));
            using (var cancellation = new CancellationTokenSource())
            {
                fixture.OnBackingRead = cancellation.Cancel;
                Assert.ThrowsException<OperationCanceledException>(() => fixture.Capture(cancellation.Token));
            }
            Assert.AreEqual(0, fixture.Service.WriteCalls);
        }

        [TestMethod]
        [TestCategory("Phase2G6A")]
        public void SolutionComponentEvidenceFieldsRemainAvailableWithoutUnsupportedAuditFields()
        {
            var query = DataverseSolutionMembershipEvidenceQuery();
            Assert.IsFalse(query.ColumnSet.Columns.Any(item =>
                string.Equals(item, "componentidunique", StringComparison.OrdinalIgnoreCase)));
            Assert.AreEqual("solutioncomponent", query.EntityName);
            Assert.AreEqual(ConditionOperator.Equal, query.Criteria.Conditions[0].Operator);
        }

        private static QueryExpression DataverseSolutionMembershipEvidenceQuery()
        {
            return D365SolutionComparer.Services.Membership.DataverseAppSettingEvidenceOperation
                .CreateSolutionComponentEvidenceQuery(Guid.NewGuid());
        }

        private sealed class CaptureFixture
        {
            internal readonly Guid ObjectId = Guid.NewGuid();
            internal readonly Guid DefinitionId = Guid.NewGuid();
            internal readonly Guid AppId = Guid.NewGuid();
            internal readonly Entity MembershipRow;
            internal readonly MembershipSnapshot Snapshot;
            internal readonly DataverseRequestCounter Counter = new DataverseRequestCounter();
            internal readonly FakeOrganizationService Service;
            internal string Outcome;
            internal Action OnBackingRead;

            internal CaptureFixture(int type, string label,
                params SolutionComponentDefinitionIdentity[] registrations)
            {
                var solution = MembershipTestData.Solution("EDU");
                MembershipRow = MembershipTestData.ComponentRow(solution, type);
                MembershipRow["objectid"] = ObjectId;
                if (label != null) MembershipRow.FormattedValues["componenttype"] = label;
                Snapshot = MembershipSnapshot.Complete(solution,
                    registrations.Select(registration => new ComponentIdentity(
                        new SolutionComponentRecord(MembershipRow.Id, type, ObjectId),
                        IdentityResolutionStatus.Unsupported, registeredDefinition: registration)), DateTimeOffset.UtcNow);
                Service = MembershipTestData.Service(solution, Query);
                Service.ExecuteRequest = request =>
                {
                    if (request is WhoAmIRequest) return MembershipTestData.WhoAmI(solution.Environment.OrganizationId);
                    var metadataRequest = request as RetrieveEntityRequest;
                    Assert.IsNotNull(metadataRequest, "Only WhoAmI and RetrieveEntity read requests are allowed.");
                    if (Outcome == "unavailable" && metadataRequest.LogicalName == "appsetting")
                        throw new FaultException("Metadata unavailable.");
                    return Metadata(metadataRequest.LogicalName);
                };
            }

            internal AppSettingEvidenceReport Capture(CancellationToken token = default(CancellationToken)) =>
                new DataverseAppSettingEvidenceOperation().Capture(Service, Snapshot, token, requestCounter: Counter);

            private EntityCollection Query(QueryExpression query)
            {
                if (query.EntityName == "solutioncomponent")
                {
                    CollectionAssert.AreEqual(new[] { "solutioncomponentid", "solutionid", "componenttype", "objectid",
                        "rootcomponentbehavior", "rootsolutioncomponentid", "ismetadata" }, query.ColumnSet.Columns.ToArray());
                    Assert.AreEqual(1, query.Criteria.Conditions.Count);
                    Assert.AreEqual(Snapshot.Solution.SolutionId, query.Criteria.Conditions.Single().Values.Single());
                    return MembershipTestData.Rows(MembershipRow);
                }
                CollectionAssert.Contains(new[] { "appsetting", "settingdefinition", "appmodule" }, query.EntityName);
                var filter = query.Criteria.Conditions.Single();
                Assert.AreEqual(query.EntityName + "id", filter.AttributeName);
                Assert.AreEqual(ConditionOperator.In, filter.Operator);
                Assert.IsTrue(filter.Values.All(item => item is Guid));
                if (query.EntityName == "appsetting")
                {
                    Assert.AreEqual(ObjectId, filter.Values.Single());
                    OnBackingRead?.Invoke();
                    if (Outcome == "fault") throw new FaultException("Backing read failed.");
                    if (Outcome == "missing") return MembershipTestData.Rows();
                    var row = new Entity("appsetting", ObjectId)
                    {
                        ["appsettingid"] = ObjectId,
                        ["settingdefinitionid"] = new EntityReference("settingdefinition", DefinitionId),
                        ["parentappmoduleid"] = new EntityReference("appmodule", AppId)
                    };
                    if (Outcome == "duplicate") return MembershipTestData.Rows(row, row);
                    var rows = MembershipTestData.Rows(row);
                    if (Outcome == "incomplete") rows.MoreRecords = true; // no cookie: never a completed inventory
                    return rows;
                }
                var expectedId = query.EntityName == "settingdefinition" ? DefinitionId : AppId;
                if (!filter.Values.Contains(expectedId)) return MembershipTestData.Rows();
                var record = new Entity(query.EntityName, expectedId)
                {
                    [query.EntityName + "id"] = expectedId,
                    ["name"] = query.EntityName == "settingdefinition" ? "new_setting" : "App"
                };
                if (query.EntityName == "appmodule") record["uniquename"] = "new_app";
                return MembershipTestData.Rows(record);
            }

            private static RetrieveEntityResponse Metadata(string entity)
            {
                var fields = entity == "appsetting"
                    ? new[] { "appsettingid", "settingdefinitionid", "parentappmoduleid" }
                    : entity == "settingdefinition" ? new[] { "settingdefinitionid", "name" }
                    : new[] { "appmoduleid", "name", "uniquename" };
                var metadata = new EntityMetadata { LogicalName = entity };
                typeof(EntityMetadata).GetProperty("PrimaryIdAttribute").SetValue(metadata, entity + "id", null);
                typeof(EntityMetadata).GetProperty("PrimaryNameAttribute").SetValue(metadata,
                    entity == "appsetting" ? null : "name", null);
                typeof(EntityMetadata).GetProperty("Attributes").SetValue(metadata,
                    fields.Select(field => (AttributeMetadata)new StringAttributeMetadata { LogicalName = field }).ToArray(), null);
                var response = new RetrieveEntityResponse();
                response.Results["EntityMetadata"] = metadata;
                return response;
            }
        }

        private static AppSettingCandidateEvidence Candidate(string identity)
        {
            return new AppSettingCandidateEvidence("DEV", "edu", Guid.NewGuid(), Guid.NewGuid(), 10094,
                "App Setting", Guid.NewGuid(), null, null, null, "test", null,
                AppSettingEvidenceState.Confirmed, "setting", "setting", "app", "app.one", "App", identity, "");
        }

        private static AppSettingEvidenceReport Report(string environment,
            params AppSettingCandidateEvidence[] candidates)
        {
            return new AppSettingEvidenceReport(environment, "edu", Guid.NewGuid(), null, null, candidates,
                new AppSettingRequestSummary(0, 0, 0, 0, 0, 0, 0, 0), new List<string>());
        }
    }
}
