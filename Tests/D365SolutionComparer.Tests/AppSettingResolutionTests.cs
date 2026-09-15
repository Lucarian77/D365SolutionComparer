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
using Microsoft.Crm.Sdk.Messages;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Query;

namespace D365SolutionComparer.Tests
{
    [TestClass]
    public class AppSettingResolutionTests
    {
        [DataTestMethod, TestCategory("Phase2G6B"), TestCategory("Phase2G6B2")]
        [DataRow(10075, "AppSetting", "appsetting", "")]
        [DataRow(54321, "AppSetting", "appsetting", "Unrelated label")]
        [DataRow(54322, "aPpSeTtInG", "AppSetting", "")]
        public void OnlyRegisteredDefinitionSelectsAppSetting(int type, string name, string entity, string label)
        {
            var fixture = new Fixture(type) { DefinitionName = name, PrimaryEntity = entity, Label = label };
            fixture.Add();
            var result = fixture.Resolve().Components.Single();
            Assert.AreEqual(IdentityResolutionStatus.Resolved, result.Status);
            Assert.AreEqual(ComponentSemanticKinds.AppSetting, result.SemanticKind);
            Assert.AreEqual(type, result.RegisteredDefinition.ObjectTypeCode);
            Assert.AreEqual(1, fixture.Counter.GetQueryCount("appsetting"));
            Assert.AreEqual(0, fixture.Service.WriteCalls);
        }

        [DataTestMethod, TestCategory("Phase2G6B"), TestCategory("Phase2G6B2")]
        [DataRow("none", "")]
        [DataRow("none", "AppSetting")]
        [DataRow("wrongentity", "AppSetting")]
        [DataRow("wrongname", "AppSetting")]
        [DataRow("conflict", "")]
        public void NumericTypeOrLabelOrConflictingRegistrationCannotSelectResolver(string registration, string label)
        {
            var fixture = new Fixture { Registration = registration, Label = label };
            fixture.Add();
            var result = fixture.Resolve().Components.Single();
            Assert.AreNotEqual(IdentityResolutionStatus.Resolved, result.Status);
            Assert.IsNull(result.ComparisonKey);
            Assert.AreEqual(0, fixture.Counter.GetQueryCount("appsetting"));
        }

        [TestMethod, TestCategory("Phase2G6B")]
        public void EquivalentRepeatedRegistrationCollapsesWithoutExtraBackingReads()
        {
            var fixture = new Fixture { Registration = "equivalent" }; fixture.Add();
            Assert.AreEqual(IdentityResolutionStatus.Resolved, fixture.Resolve().Components.Single().Status);
            Assert.AreEqual(1, fixture.Counter.GetQueryCount("appsetting"));
        }

        [TestMethod, TestCategory("Phase2G6B")]
        public void CompletedSnapshotRegisteredEvidenceIsReusedWithoutDefinitionRediscovery()
        {
            var fixture = new Fixture { UseSnapshotRegistration = true }; fixture.Add();
            Assert.AreEqual(IdentityResolutionStatus.Resolved, fixture.Resolve().Components.Single().Status);
            // Existing connection-reference discovery remains once; no raw-type IN discovery.
            Assert.AreEqual(1, fixture.Queries.Count(q => q.EntityName == "solutioncomponentdefinition"));
            Assert.IsFalse(fixture.Queries.Any(q => q.EntityName == "solutioncomponentdefinition" &&
                q.Criteria.Conditions.Any(c => c.AttributeName == "objecttypecode")));
        }

        [TestMethod, TestCategory("Phase2G6B")]
        public void ConflictingCompletedSnapshotDefinitionsCannotSelectAppSetting()
        {
            var fixture = new Fixture(); fixture.Add(); fixture.Add();
            var snapshot = MembershipSnapshot.Complete(new SolutionIdentity(fixture.Environment, Guid.NewGuid(), "edu"),
                fixture.Records.Select((record, index) => new ComponentIdentity(record, IdentityResolutionStatus.Unsupported,
                    registeredDefinition: new SolutionComponentDefinitionIdentity(fixture.Type,
                        index == 0 ? "AppSetting" : "Other", "appsetting"))), DateTimeOffset.UtcNow);
            var resolved = new DataverseComponentIdentityResolver().ResolveSnapshot(fixture.Service, snapshot, CancellationToken.None, fixture.Counter);
            Assert.IsTrue(resolved.Components.All(i => i.Status == IdentityResolutionStatus.Ambiguous && i.ComparisonKey == null));
            Assert.AreEqual(0, fixture.Counter.GetQueryCount("appsetting"));
        }

        [DataTestMethod, TestCategory("Phase2G6B"), TestCategory("Phase2G6B2")]
        [DataRow("missing", IdentityResolutionStatus.Unresolved)]
        [DataRow("duplicate", IdentityResolutionStatus.Ambiguous)]
        [DataRow("conflicting", IdentityResolutionStatus.Unresolved)]
        [DataRow("missingpk", IdentityResolutionStatus.Unresolved)]
        [DataRow("wrongentity", IdentityResolutionStatus.Unresolved)]
        [DataRow("paged", IdentityResolutionStatus.Unresolved)]
        [DataRow("fault", IdentityResolutionStatus.Unresolved)]
        [DataRow("null", IdentityResolutionStatus.Unresolved)]
        public void AppSettingCorrelationFailuresRemainConservative(string failure, IdentityResolutionStatus expected)
        {
            var fixture = new Fixture { FailureTable = "appsetting", Failure = failure }; fixture.Add();
            var result = fixture.Resolve().Components.Single();
            Assert.AreEqual(expected, result.Status);
            Assert.IsNull(result.ComparisonKey);
            Assert.AreEqual(ComponentSemanticKinds.AppSetting, result.SemanticKind);
            Assert.AreEqual(0, fixture.Counter.GetQueryCount("settingdefinition"));
            AssertSafe(result.Diagnostic + string.Join("\n", result.DiagnosticEvidence));
        }

        [DataTestMethod, TestCategory("Phase2G6B"), TestCategory("Phase2G6B2")]
        [DataRow("objectid")]
        [DataRow("emptyobjectid")]
        [DataRow("settingdefinitionid")]
        [DataRow("parentappmoduleid")]
        [DataRow("wrongdefinitionreference")]
        [DataRow("wrongparentreference")]
        [DataRow("definitionuniquename")]
        [DataRow("parentname")]
        public void MissingOrInvalidIdentityEvidenceNeverResolves(string missing)
        {
            var fixture = new Fixture(); var record = fixture.Add();
            var setting = fixture.Settings.Single();
            if (missing == "objectid") fixture.Records[0] = new SolutionComponentRecord(record.SolutionComponentId, fixture.Type, null);
            else if (missing == "emptyobjectid") fixture.Records[0] = new SolutionComponentRecord(record.SolutionComponentId, fixture.Type, Guid.Empty);
            else if (missing == "settingdefinitionid" || missing == "parentappmoduleid") setting.Attributes.Remove(missing);
            else if (missing == "wrongdefinitionreference") setting["settingdefinitionid"] = new EntityReference("appmodule", fixture.Definitions.Single().Id);
            else if (missing == "wrongparentreference") setting["parentappmoduleid"] = Guid.NewGuid();
            else if (missing == "definitionuniquename") fixture.Definitions.Single()["uniquename"] = " ";
            else fixture.Apps.Single()["uniquename"] = "";
            var identity = fixture.Resolve().Components.Single();
            Assert.AreEqual(IdentityResolutionStatus.Unresolved, identity.Status);
            Assert.IsNull(identity.ComparisonKey);
            if (missing == "objectid") Assert.AreEqual(0, fixture.Counter.GetQueryCount("appsetting"));
        }

        [DataTestMethod, TestCategory("Phase2G6B"), TestCategory("Phase2G6B2")]
        [DataRow("settingdefinition", "missing", IdentityResolutionStatus.Unresolved)]
        [DataRow("settingdefinition", "duplicate", IdentityResolutionStatus.Ambiguous)]
        [DataRow("settingdefinition", "conflicting", IdentityResolutionStatus.Ambiguous)]
        [DataRow("settingdefinition", "paged", IdentityResolutionStatus.Unresolved)]
        [DataRow("settingdefinition", "fault", IdentityResolutionStatus.Unresolved)]
        [DataRow("appmodule", "missing", IdentityResolutionStatus.Unresolved)]
        [DataRow("appmodule", "duplicate", IdentityResolutionStatus.Ambiguous)]
        [DataRow("appmodule", "conflicting", IdentityResolutionStatus.Ambiguous)]
        [DataRow("appmodule", "paged", IdentityResolutionStatus.Ambiguous)]
        [DataRow("appmodule", "fault", IdentityResolutionStatus.Unresolved)]
        [DataRow("appmodule", "null", IdentityResolutionStatus.Unresolved)]
        public void RelatedCorrelationFailuresDoNotBecomeMissing(string table, string failure, IdentityResolutionStatus expected)
        {
            var fixture = new Fixture { FailureTable = table, Failure = failure }; fixture.Add();
            var snapshot = fixture.Resolve(); var identity = snapshot.Components.Single();
            Assert.AreEqual(expected, identity.Status);
            Assert.IsNull(identity.ComparisonKey);
            Assert.AreEqual(MembershipPresence.Indeterminate, new SolutionMembershipComparer().Compare(snapshot,
                Empty(snapshot)).Single().Presence);
            AssertSafe(identity.Diagnostic + string.Join("\n", identity.DiagnosticEvidence));
        }

        [TestMethod, TestCategory("Phase2G6B"), TestCategory("Phase2G6B2")]
        public void PortableIdentityUsesCaseInsensitiveNamesAndIgnoresEveryLocalId()
        {
            var source = new Fixture(); source.Add("App.One", "Setting.One");
            var target = new Fixture(54321); target.Add("APP.ONE", "SETTING.ONE");
            var left = source.Resolve(); var right = target.Resolve();
            Assert.AreNotEqual(left.Components[0].Record.ObjectId, right.Components[0].Record.ObjectId);
            Assert.IsTrue(StringComparer.OrdinalIgnoreCase.Equals(left.Components[0].ComparisonKey, right.Components[0].ComparisonKey));
            Assert.AreEqual(MembershipPresence.PresentInBoth, new SolutionMembershipComparer().Compare(left, right).Single().Presence);
        }

        [TestMethod, TestCategory("Phase2G6B"), TestCategory("Phase2G6B2")]
        public void LiveShapedSettingDefinitionWithoutNameResolvesFromUniqueNameOnly()
        {
            var fixture = new Fixture(); fixture.Add("parent_app", "publisher_setting");
            Assert.IsFalse(fixture.Definitions.Single().Attributes.ContainsKey("name"));
            var identity = fixture.Resolve().Components.Single();
            Assert.AreEqual(IdentityResolutionStatus.Resolved, identity.Status);
            Assert.AreEqual(AppSettingResolutionOperation.PortableKey("parent_app", "publisher_setting"),
                identity.ComparisonKey);
            var query = fixture.Queries.Single(item => item.EntityName == "settingdefinition");
            CollectionAssert.AreEqual(new[] { "settingdefinitionid", "uniquename" },
                query.ColumnSet.Columns.ToArray());
            CollectionAssert.DoesNotContain(query.ColumnSet.Columns, "name");
        }

        [TestMethod, TestCategory("Phase2G6B")]
        public void FramingPreservesSeparatorsWhitespaceAndDifferentParentsWithoutCollisions()
        {
            Assert.AreNotEqual(AppSettingResolutionOperation.PortableKey("a+b", "c"), AppSettingResolutionOperation.PortableKey("a", "b+c"));
            Assert.AreNotEqual(AppSettingResolutionOperation.PortableKey("a:1", "b"), AppSettingResolutionOperation.PortableKey("a", "1:b"));
            Assert.AreNotEqual(AppSettingResolutionOperation.PortableKey("app1", "setting"), AppSettingResolutionOperation.PortableKey("app2", "setting"));
            Assert.AreNotEqual(AppSettingResolutionOperation.PortableKey(" app", "setting"), AppSettingResolutionOperation.PortableKey("app", "setting"));
            Assert.AreEqual("appsetting:v1:3:app:7:setting", AppSettingResolutionOperation.PortableKey("app", "setting"));
        }

        [TestMethod, TestCategory("Phase2G6B")]
        public void DuplicatePortableKeysAreScopedAndDoNotBlockUnrelatedAppSettings()
        {
            var source = new Fixture(); source.Add("app", "collision"); source.Add("app", "onlysource");
            var target = new Fixture(); target.Add("app", "collision"); target.Add("APP", "COLLISION"); target.Add("app", "onlytarget");
            var left = source.Resolve(); var right = target.Resolve();
            Assert.AreEqual(2, right.Components.Count(i => i.Status == IdentityResolutionStatus.Ambiguous));
            Assert.IsTrue(right.Components.Where(i => i.Status == IdentityResolutionStatus.Ambiguous).All(i =>
                i.BlockerScope == ResolutionBlockerScope.PortableIdentity && i.ComparisonKey == null));
            var comparison = new SolutionMembershipComparer().Compare(left, right);
            Assert.AreEqual(1, comparison.Count(i => i.Presence == MembershipPresence.OnlyInSource));
            Assert.AreEqual(1, comparison.Count(i => i.Presence == MembershipPresence.OnlyInTarget));
            Assert.AreEqual(3, comparison.Count(i => i.Presence == MembershipPresence.Indeterminate));
            Assert.AreEqual(0, comparison.Count(i => i.Presence == MembershipPresence.PresentInBoth));
        }

        [TestMethod, TestCategory("Phase2G6B")]
        public void RepeatedObjectIdsAreReadOnceButDifferentMembershipRowsRemainAmbiguous()
        {
            var fixture = new Fixture(); var record = fixture.Add();
            fixture.Records.Add(new SolutionComponentRecord(Guid.NewGuid(), fixture.Type, record.ObjectId));
            var result = fixture.Resolve();
            Assert.AreEqual(2, result.Components.Count(i => i.Status == IdentityResolutionStatus.Ambiguous));
            Assert.AreEqual(1, fixture.Queries.Single(q => q.EntityName == "appsetting").Criteria.Conditions.Single().Values.Count);
            Assert.AreEqual(2, result.Components.Select(i => i.DiagnosticEvidence[1]).Distinct().Count());
        }

        [DataTestMethod, TestCategory("Phase2G6B")]
        [DataRow("incomplete")]
        [DataRow("unavailable")]
        [DataRow("absent")]
        public void OppositeCoverageUsesExistingAbsenceRules(string state)
        {
            var source = new Fixture(); source.Add(); var left = source.Resolve();
            var target = new Fixture(); target.Add(); target.Settings.Clear(); var right = target.Resolve();
            if (state == "unavailable") right = MembershipSnapshot.Unavailable(right.Environment, "edu", DateTimeOffset.UtcNow, "unavailable");
            if (state == "absent") right = MembershipSnapshot.Absent(right.Environment, "edu", DateTimeOffset.UtcNow);
            var match = new SolutionMembershipComparer().Compare(left, right).Single(i => i.Source != null);
            Assert.AreEqual(state == "absent" ? MembershipPresence.OnlyInSource : MembershipPresence.Indeterminate, match.Presence);
        }

        [DataTestMethod, TestCategory("Phase2G6B")]
        [DataRow("datatype")]
        [DataRow("isoverridable")]
        [DataRow("overridablelevel")]
        [DataRow("releaselevel")]
        public void SafeStructuralDifferencesAppearInStandardDetailPresentation(string property)
        {
            var source = new Fixture(); source.Add(); var left = source.ReadDefinitions();
            var target = new Fixture(); target.Add();
            target.Definitions.Single()[property] = property == "isoverridable" ? (object)false : new OptionSetValue(2);
            var right = target.ReadDefinitions();
            var membership = new SolutionMembershipComparer().Compare(left.Membership, right.Membership);
            var detail = new ComponentDetailComparer().Compare(membership, left, right).Single();
            Assert.AreEqual(ComponentDetailComparisonStatus.Different, detail.Status);
            Assert.AreEqual(property, detail.Differences.Single().PropertyName);
            var view = Present(left, right).Rows.Single();
            Assert.AreEqual("App Setting", view.ComponentKind);
            Assert.AreEqual("Present in Both", view.MembershipStatus);
            Assert.AreEqual(property, view.ChangedProperties);
            Assert.IsTrue(view.DefinitionDetail.Properties.Single(p => p.PropertyName == property).Changed);
        }

        [DataTestMethod, TestCategory("Phase2G6B")]
        [DataRow("ismanaged")]
        [DataRow("componentstate")]
        [DataRow("componentidunique")]
        [DataRow("value")]
        [DataRow("defaultvalue")]
        public void AuditAndSensitiveFieldsCannotCreateDefinitionDifferences(string field)
        {
            var source = new Fixture(); source.Add(); var left = source.ReadDefinitions();
            var target = new Fixture(); target.Add();
            target.Settings.Single()[field] = Fixture.Secret;
            target.Definitions.Single()[field] = Fixture.Secret;
            var right = target.ReadDefinitions(); var view = Present(left, right);
            Assert.AreEqual("Match", view.Rows.Single().DefinitionStatus);
            Assert.IsFalse(right.Definitions[0].ComparableProperties.ContainsKey(field));
            AssertSafe(string.Join("\n", right.Definitions.SelectMany(d => d.DiagnosticEvidence)));
            AssertSafe(string.Join("\n", right.Membership.Components.SelectMany(i => i.DiagnosticEvidence)));
            AssertSafe(string.Join("\n", view.Rows[0].DefinitionDetail.Properties.Select(p => p.SourceValue + p.TargetValue)));
            int requests = target.Counter.TotalRequests;
            AssertSafe(new MembershipCoverageCsvExporter().CreateCsv(view));
            Assert.AreEqual(requests, target.Counter.TotalRequests);
            Assert.IsTrue(target.Queries.Where(q => q.EntityName == "appsetting" || q.EntityName == "settingdefinition")
                .All(q => !q.ColumnSet.AllColumns && !q.ColumnSet.Columns.Contains("value") && !q.ColumnSet.Columns.Contains("defaultvalue")));
        }

        [DataTestMethod, TestCategory("Phase2G6B"), TestCategory("Phase2G6B1")]
        [DataRow("metadatafault")]
        [DataRow("absentschema")]
        [DataRow("wrongschematype")]
        [DataRow("duplicatemetadata")]
        [DataRow("notreadable")]
        [DataRow("missingvalue")]
        [DataRow("wrongvaluetype")]
        public void UnverifiedOrIncompleteStructuralEvidenceNeverProducesMatch(string mode)
        {
            var fixture = new Fixture { SchemaMode = mode }; fixture.Add();
            if (mode == "missingvalue") fixture.Definitions.Single().Attributes.Remove("releaselevel");
            if (mode == "wrongvaluetype") fixture.Definitions.Single()["releaselevel"] = Fixture.Secret;
            var result = fixture.ReadDefinitions();
            Assert.AreEqual(IdentityResolutionStatus.Resolved, result.Membership.Components[0].Status);
            Assert.AreEqual(ComponentDefinitionReadStatus.Unresolved, result.Definitions[0].Status);
            Assert.AreEqual("Unresolved", Present(result, result).Rows.Single().DefinitionStatus);
            Assert.AreEqual(0, result.Definitions[0].ComparableProperties.Count);
            AssertSafe(result.Definitions[0].Diagnostic);
            if (mode == "absentschema" || mode == "wrongschematype" || mode == "duplicatemetadata" || mode == "notreadable")
                Assert.AreEqual(1, fixture.Queries.Count(q => q.EntityName == "settingdefinition"));
        }

        [TestMethod, TestCategory("Phase2G6B"), TestCategory("Phase2G6B1"),
            TestCategory("Phase2G6B2")]
        public void PortableIdentityDoesNotRetrieveStructuralMetadataOrOptionalProperties()
        {
            var fixture = new Fixture { SchemaMode = "metadatafault" }; fixture.Add();
            var snapshot = fixture.Resolve();
            Assert.AreEqual(IdentityResolutionStatus.Resolved, snapshot.Components.Single().Status);
            Assert.AreEqual(0, fixture.Counter.GetExecuteCount("RetrieveEntity"));
            var definitionQuery = fixture.Queries.Single(q => q.EntityName == "settingdefinition");
            CollectionAssert.AreEqual(new[] { "settingdefinitionid", "uniquename" },
                definitionQuery.ColumnSet.Columns.ToArray());
            Assert.IsFalse(definitionQuery.ColumnSet.Columns.Any(AppSettingResolutionOperation.StructuralFields.Contains));
        }

        [DataTestMethod, TestCategory("Phase2G6B"), TestCategory("Phase2G6B1"),
            TestCategory("Phase2G6B2")]
        [DataRow("missing", "Unresolved")]
        [DataRow("duplicate", "Ambiguous")]
        [DataRow("conflicting", "Unresolved")]
        [DataRow("missingpk", "Unresolved")]
        [DataRow("wrongentity", "Unresolved")]
        [DataRow("paged", "Unresolved")]
        [DataRow("fault", "Unresolved")]
        [DataRow("null", "Unresolved")]
        public void OptionalStructuralReadFailuresNeverInvalidatePortableIdentity(string failure,
            string expectedDefinition)
        {
            var fixture = new Fixture { FailureTable = "settingdefinitionstructure", Failure = failure };
            fixture.Add();
            var result = fixture.ReadDefinitions();
            Assert.AreEqual(IdentityResolutionStatus.Resolved, result.Membership.Components.Single().Status);
            Assert.IsNotNull(result.Membership.Components.Single().ComparisonKey);
            Assert.AreEqual(expectedDefinition, result.Definitions.Single().Status.ToString());
            Assert.AreNotEqual(ComponentDefinitionReadStatus.Available, result.Definitions.Single().Status);
            var presented = Present(result, result).Rows.Single();
            Assert.AreEqual("Present in Both", presented.MembershipStatus);
            Assert.AreNotEqual("Match", presented.DefinitionStatus);
            AssertSafe(result.Definitions.Single().Diagnostic);
        }

        [TestMethod, TestCategory("Phase2G6B"), TestCategory("Phase2G6B1")]
        public void StructuralPropertiesUseASeparateAllowlistedDefinitionQuery()
        {
            var fixture = new Fixture(); fixture.Add();
            var result = fixture.ReadDefinitions();
            Assert.AreEqual(IdentityResolutionStatus.Resolved, result.Membership.Components.Single().Status);
            Assert.AreEqual(ComponentDefinitionReadStatus.Available, result.Definitions.Single().Status);
            var queries = fixture.Queries.Where(q => q.EntityName == "settingdefinition").ToList();
            Assert.AreEqual(2, queries.Count);
            CollectionAssert.AreEqual(new[] { "settingdefinitionid", "uniquename" },
                queries[0].ColumnSet.Columns.ToArray());
            CollectionAssert.AreEqual(new[] { "settingdefinitionid", "datatype", "isoverridable",
                "overridablelevel", "releaselevel" }, queries[1].ColumnSet.Columns.ToArray());
            Assert.IsFalse(queries.SelectMany(q => q.ColumnSet.Columns).Any(field =>
                field == "value" || field == "defaultvalue"));
        }

        [DataTestMethod, TestCategory("Phase2G6B"), TestCategory("Phase2G6B1")]
        [DataRow(15, 1)]
        [DataRow(201, 2)]
        [DataRow(500, 3)]
        public void DeterministicBatchesAndDefinitionInventory(int count, int batches)
        {
            var fixture = new Fixture(); var first = fixture.Add();
            for (int i = 1; i < count; i++)
                fixture.Add("app", "setting" + i, fixture.Apps[0].Id);
            // Include the parent as normal Type 80 membership: its identity/definition read
            // must satisfy every AppSetting parent without another appmodule query.
            fixture.Records.Add(new SolutionComponentRecord(Guid.NewGuid(), 80, fixture.Apps[0].Id));
            var result = fixture.ReadDefinitions();
            Assert.AreEqual(batches, fixture.Counter.GetQueryCount("appsetting"));
            Assert.AreEqual(batches * 2, fixture.Counter.GetQueryCount("settingdefinition"));
            Assert.AreEqual(1, fixture.Counter.GetQueryCount("appmodule"));
            Assert.AreEqual(1, fixture.Counter.GetExecuteCount("WhoAmI"));
            Assert.AreEqual(1, fixture.Counter.GetExecuteCount("RetrieveEntity"));
            Assert.AreEqual(count + 1, result.Membership.Components.Count);
            Assert.AreEqual(0, fixture.Service.WriteCalls);
            var requests = fixture.Queries.Where(q => q.EntityName == "appsetting").ToList();
            var requestedIds = requests.SelectMany(q => q.Criteria.Conditions.Single().Values.Cast<Guid>()).ToArray();
            CollectionAssert.AreEqual(requestedIds.OrderBy(id => id).ToArray(), requestedIds);
            foreach (var query in requests)
            {
                Assert.AreEqual(ConditionOperator.In, query.Criteria.Conditions.Single().Operator);
                Assert.IsTrue(query.Criteria.Conditions.Single().Values.Count <= 200);
                Assert.IsTrue(query.Criteria.Conditions.Single().Values.All(v => v is Guid));
                CollectionAssert.AreEqual(new[] { "appsettingid", "settingdefinitionid", "parentappmoduleid" }, query.ColumnSet.Columns.ToArray());
            }
            int before = fixture.Counter.TotalRequests;
            new DataverseComponentDefinitionReader().Read(fixture.Context, result.Membership, CancellationToken.None);
            Assert.AreEqual(before, fixture.Counter.TotalRequests);
        }

        [DataTestMethod, TestCategory("Phase2G6B"), TestCategory("Phase2G6B1")]
        [DataRow("appsetting")]
        [DataRow("settingdefinition")]
        [DataRow("appmodule")]
        [DataRow("metadata")]
        [DataRow("settingdefinitionstructure")]
        public void CancellationPropagatesAtEveryNewStage(string stage)
        {
            using (var cancellation = new CancellationTokenSource())
            {
                var fixture = new Fixture { Cancel = cancellation, CancelStage = stage }; fixture.Add();
                Assert.ThrowsException<OperationCanceledException>(() =>
                    stage == "metadata" || stage == "settingdefinitionstructure"
                        ? (object)fixture.ReadDefinitions(cancellation.Token)
                        : fixture.Resolve(cancellation.Token));
                Assert.AreEqual(0, fixture.Service.WriteCalls);
            }
        }

        [TestMethod, TestCategory("Phase2G6B")]
        public void KnownStaticAndUnrelatedRegisteredKindsRetainTheirBehavior()
        {
            var fixture = new Fixture { DefinitionName = "AnotherFamily" }; fixture.Add();
            fixture.Records.Add(new SolutionComponentRecord(Guid.NewGuid(), 26, null));
            var result = fixture.Resolve();
            Assert.IsTrue(result.Components.All(i => i.Status == IdentityResolutionStatus.Unsupported));
            Assert.IsNull(result.Components[0].ComparisonKey);
            Assert.AreEqual("unsupported:componenttype:26", result.Components[1].SemanticKind);
            Assert.AreEqual(0, fixture.Counter.GetQueryCount("appsetting"));
        }

        [TestMethod, TestCategory("Phase2G6B")]
        public void IndependentRelatedIdsAlsoUseDeterministicBatchesOfAtMost200()
        {
            var fixture = new Fixture();
            for (int i = 0; i < 401; i++) fixture.Add("app" + i, "setting" + i);
            var result = fixture.ReadDefinitions();
            Assert.IsTrue(result.Membership.Components.All(i => i.Status == IdentityResolutionStatus.Resolved));
            foreach (var table in new[] { "appsetting", "appmodule" })
            {
                Assert.AreEqual(3, fixture.Counter.GetQueryCount(table));
                var queries = fixture.Queries.Where(q => q.EntityName == table).ToList();
                Assert.IsTrue(queries.All(q => q.Criteria.Conditions.Single().Values.Count <= 200));
                var ids = queries.SelectMany(q => q.Criteria.Conditions.Single().Values.Cast<Guid>()).ToArray();
                CollectionAssert.AreEqual(ids.OrderBy(id => id).ToArray(), ids);
                Assert.AreEqual(401, ids.Distinct().Count());
            }
            Assert.AreEqual(6, fixture.Counter.GetQueryCount("settingdefinition"));
            var definitionQueries = fixture.Queries.Where(q => q.EntityName == "settingdefinition").ToList();
            Assert.IsTrue(definitionQueries.All(q => q.Criteria.Conditions.Single().Values.Count <= 200));
            Assert.AreEqual(2, definitionQueries.SelectMany(q => q.Criteria.Conditions.Single().Values.Cast<Guid>())
                .GroupBy(id => id).Select(group => group.Count()).Distinct().Single());
        }

        [TestMethod, TestCategory("Phase2G6B"), TestCategory("Phase2G6B2")]
        public void NormalCoordinatedOperationIncludesAppSettingWithoutAutomaticEvidenceCapture()
        {
            var fixture = new Fixture();
            for (int i = 0; i < 15; i++) fixture.Add("app" + (i % 3), "setting" + i,
                i < 3 ? (Guid?)null : fixture.Apps[i % 3].Id);
            foreach (var app in fixture.Apps)
                fixture.Records.Add(new SolutionComponentRecord(Guid.NewGuid(), 80, app.Id));
            var result = new DataverseComponentDefinitionOperation().ReadAndResolve(fixture.Service, "test", "edu",
                CancellationToken.None, null, fixture.Counter);
            Assert.AreEqual(18, result.Membership.Components.Count);
            Assert.AreEqual(15, result.Membership.Components.Count(i => i.SemanticKind == ComponentSemanticKinds.AppSetting &&
                i.Status == IdentityResolutionStatus.Resolved));
            Assert.IsTrue(result.Definitions.All(d => d.Status == ComponentDefinitionReadStatus.Available));
            Assert.AreEqual(10, fixture.Counter.TotalRequests);
            foreach (var table in new[] { "solution", "solutioncomponent", "appsetting", "appmodule" })
                Assert.AreEqual(1, fixture.Counter.GetQueryCount(table));
            Assert.AreEqual(2, fixture.Counter.GetQueryCount("settingdefinition"));
            Assert.AreEqual(1, fixture.Counter.GetExecuteCount("WhoAmI"));
            Assert.AreEqual(1, fixture.Counter.GetExecuteCount("RetrieveEntity"));
            Assert.AreEqual(0, fixture.Service.WriteCalls);
        }

        [TestMethod, TestCategory("Phase2G6B")]
        public void FailedExistingAppModuleLookupIsNotRepeatedForAppSettingParent()
        {
            var fixture = new Fixture { FailureTable = "appmodule", Failure = "missing" }; fixture.Add();
            fixture.Records.Add(new SolutionComponentRecord(Guid.NewGuid(), 80, fixture.Apps[0].Id));
            Assert.IsTrue(fixture.Resolve().Components.All(i => i.Status == IdentityResolutionStatus.Unresolved));
            Assert.AreEqual(1, fixture.Counter.GetQueryCount("appmodule"));
        }

        [TestMethod, TestCategory("Phase2G6B")]
        public void StandaloneIdentityApiResolvesUsingTheSameProductionRules()
        {
            var fixture = new Fixture(); var record = fixture.Add();
            var identity = new DataverseComponentIdentityResolver().Resolve(fixture.Service, fixture.Environment, record, CancellationToken.None);
            Assert.AreEqual(IdentityResolutionStatus.Resolved, identity.Status);
            Assert.AreEqual(AppSettingResolutionOperation.PortableKey("app", "setting"), identity.ComparisonKey);
            Assert.AreEqual(1, fixture.Queries.Count(q => q.EntityName == "appsetting"));
        }

        [DataTestMethod, TestCategory("Phase2G6B")]
        [DataRow(true, "SourceOnly")]
        [DataRow(false, "TargetOnly")]
        public void OneSidedAppSettingDetailsUseTheAvailableDefinition(bool sourceSide, string expected)
        {
            var fixture = new Fixture(); fixture.Add(); var present = fixture.ReadDefinitions();
            var empty = new ComponentDefinitionSnapshot(Empty(present.Membership), new ComponentDefinition[0]);
            var row = (sourceSide ? Present(present, empty) : Present(empty, present)).Rows.Single();
            Assert.AreEqual(expected, row.DefinitionStatus);
            Assert.AreEqual(4, row.DefinitionDetail.Properties.Count);
            // The unchanged grid's NullValue displays this as "(not available)".
            Assert.IsTrue(row.DefinitionDetail.Properties.All(p =>
                (sourceSide ? p.TargetValue : p.SourceValue) == null));
        }

        [TestMethod, TestCategory("Phase2G6B")]
        public void SupportedCoverageReconcilesAndRetainsRegisteredEvidenceForCapture()
        {
            var fixture = new Fixture(); fixture.Add(); var snapshot = fixture.Resolve();
            var coverage = new MembershipCoverageDiagnosticsBuilder().Build(snapshot);
            var bucket = coverage.SemanticKinds.Single(b => b.SemanticKind == ComponentSemanticKinds.AppSetting);
            Assert.AreEqual(MembershipCoverageBucketType.SemanticKind, bucket.BucketType);
            Assert.AreEqual(MembershipCoverageStatus.Complete, bucket.CoverageStatus);
            Assert.AreEqual(1, bucket.Resolved);
            Assert.AreEqual(1, coverage.TotalCandidates);
            Assert.AreEqual(0, coverage.DynamicComponentTypes.Count);
            CollectionAssert.AreEqual(new[] { fixture.Type }, DataverseAppSettingEvidenceOperation.FindCandidateComponentTypes(
                snapshot.Components.Select(i => i.RegisteredDefinition)).ToArray());
        }

        [TestMethod, TestCategory("Phase2G6B")]
        public void EvidenceWordingUsesSuccessfulRelatedCorrelationWithoutChangingEvidenceStates()
        {
            var correlations = new[] { "settingdefinition", "appmodule" }.Select(table =>
                new AppSettingEntityCorrelation(table, AppSettingEvidenceState.Unresolved, table + "id", 0, null,
                    "No backing record matched the solutioncomponent.objectid.")).ToList();
            var candidate = new AppSettingCandidateEvidence("DEV", "edu", Guid.NewGuid(), Guid.NewGuid(), 54321, "",
                Guid.NewGuid(), null, null, null, "registered", correlations, AppSettingEvidenceState.Confirmed,
                Guid.NewGuid().ToString(), "setting", Guid.NewGuid().ToString(), "app", "App", "app+setting", "");
            var report = new AppSettingEvidenceReport("DEV", "edu", Guid.NewGuid(), null, null, new[] { candidate },
                new AppSettingRequestSummary(0, 0, 0, 0, 0, 0, 0, 0), new string[0]);
            var formatter = typeof(AppSettingEvidenceResultsForm).GetMethod("FormatCandidates",
                System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
            var text = (string)formatter.Invoke(null, new object[] { report, report });
            StringAssert.Contains(text, "Resolved through appsetting.settingdefinitionid.");
            StringAssert.Contains(text, "Resolved through appsetting.parentappmoduleid.");
            Assert.IsFalse(text.Contains("settingdefinition => Unresolved"));
            Assert.IsFalse(text.Contains("appmodule => Unresolved"));
            Assert.IsTrue(correlations.All(c => c.State == AppSettingEvidenceState.Unresolved));
        }

        [TestMethod, TestCategory("Phase2G6B")]
        public void StableReasonsGroupWhileRawEvidenceAndSupportedCoverageArePreserved()
        {
            var fixture = new Fixture { FailureTable = "appsetting", Failure = "missing" }; fixture.Add(); fixture.Add();
            var snapshot = fixture.Resolve();
            Assert.AreEqual(1, snapshot.Components.Select(i => i.Diagnostic).Distinct().Count());
            Assert.AreEqual(2, snapshot.Components.Select(i => i.DiagnosticEvidence[1]).Distinct().Count());
            var coverage = new MembershipCoverageDiagnosticsBuilder().Build(snapshot);
            var bucket = coverage.SemanticKinds.Single(b => b.SemanticKind == ComponentSemanticKinds.AppSetting);
            Assert.AreEqual(2, bucket.TotalCandidates);
            Assert.AreEqual(2, bucket.Unresolved);
        }

        private static MembershipComparisonPresentation Present(ComponentDefinitionSnapshot source, ComponentDefinitionSnapshot target)
        {
            var presenter = new MembershipResultPresenter();
            var membership = presenter.Create(MembershipEnvironmentResult.FromSnapshot("Source", source.Membership, 0, TimeSpan.Zero),
                MembershipEnvironmentResult.FromSnapshot("Target", target.Membership, 0, TimeSpan.Zero));
            return new ComponentDefinitionResultPresenter().Apply(membership, source, target);
        }

        private static MembershipSnapshot Empty(MembershipSnapshot snapshot) =>
            MembershipSnapshot.Complete(snapshot.Solution, new ComponentIdentity[0], snapshot.CapturedAt);
        private static void AssertSafe(string text) => Assert.IsFalse(text.Contains(Fixture.Secret), "Sensitive data escaped.");

        private sealed class Fixture
        {
            internal const string Secret = "NEVER-EMIT-THIS-AppSetting-SECRET";
            internal readonly int Type;
            internal readonly EnvironmentIdentity Environment = new EnvironmentIdentity(Guid.NewGuid(), "test");
            internal readonly List<SolutionComponentRecord> Records = new List<SolutionComponentRecord>();
            internal readonly List<Entity> Settings = new List<Entity>();
            internal readonly List<Entity> Definitions = new List<Entity>();
            internal readonly List<Entity> Apps = new List<Entity>();
            internal readonly List<QueryExpression> Queries = new List<QueryExpression>();
            internal readonly DataverseRequestCounter Counter = new DataverseRequestCounter();
            internal readonly FakeOrganizationService Service;
            internal DataverseReadContext Context;
            internal string DefinitionName = "AppSetting", PrimaryEntity = "appsetting", Registration = "valid", Label = "",
                FailureTable, Failure, SchemaMode, CancelStage;
            internal bool UseSnapshotRegistration;
            internal CancellationTokenSource Cancel;
            internal Fixture(int type = 10075)
            {
                Type = type;
                Service = new FakeOrganizationService { RetrievePage = Query, ExecuteRequest = Execute };
            }
            internal SolutionComponentRecord Add(string app = "app", string name = "setting", Guid? appId = null, Guid? definitionId = null)
            {
                var id = Guid.NewGuid(); var parent = appId ?? Guid.NewGuid(); var definition = definitionId ?? Guid.NewGuid();
                var record = new SolutionComponentRecord(Guid.NewGuid(), Type, id); Records.Add(record);
                Settings.Add(new Entity("appsetting", id) { ["appsettingid"] = id,
                    ["parentappmoduleid"] = new EntityReference("appmodule", parent),
                    ["settingdefinitionid"] = new EntityReference("settingdefinition", definition), ["value"] = Secret });
                if (!Definitions.Any(e => e.Id == definition)) Definitions.Add(new Entity("settingdefinition", definition)
                { ["settingdefinitionid"] = definition, ["uniquename"] = name, ["datatype"] = new OptionSetValue(1),
                    ["isoverridable"] = true, ["overridablelevel"] = new OptionSetValue(1),
                    ["releaselevel"] = new OptionSetValue(1), ["defaultvalue"] = Secret });
                if (!Apps.Any(e => e.Id == parent)) Apps.Add(new Entity("appmodule", parent)
                { ["appmoduleid"] = parent, ["uniquename"] = app, ["name"] = "display", ["description"] = "description",
                    ["clienttype"] = new OptionSetValue(1), ["formfactor"] = new OptionSetValue(1), ["navigationtype"] = new OptionSetValue(1),
                    ["componentstate"] = new OptionSetValue(0), ["ismanaged"] = false, ["appmoduleidunique"] = Guid.NewGuid() });
                return record;
            }
            internal MembershipSnapshot Resolve(CancellationToken token = default(CancellationToken))
            {
                Context = new DataverseReadContext(Service, Environment, token, Counter);
                var snapshot = MembershipSnapshot.Complete(new SolutionIdentity(Environment, Guid.NewGuid(), "edu"), Records.Select(r =>
                    new ComponentIdentity(r, IdentityResolutionStatus.Unresolved, registeredDefinition:
                        UseSnapshotRegistration && r.ComponentType == Type ? new SolutionComponentDefinitionIdentity(Type, DefinitionName, PrimaryEntity) : null)), DateTimeOffset.UtcNow);
                return new DataverseComponentIdentityResolver().ResolveSnapshot(Context, snapshot, token);
            }
            internal ComponentDefinitionSnapshot ReadDefinitions(
                CancellationToken token = default(CancellationToken))
            {
                var membership = Resolve(token);
                return new DataverseComponentDefinitionReader().Read(Context, membership, token);
            }
            private EntityCollection Query(QueryExpression query)
            {
                Queries.Add(query);
                var queryStage = query.EntityName == "settingdefinition" &&
                    query.ColumnSet.Columns.Any(column => AppSettingResolutionOperation.StructuralFields.Contains(column))
                    ? "settingdefinitionstructure" : query.EntityName;
                if (CancelStage == queryStage) Cancel.Cancel();
                if (query.EntityName == "solution") return Rows(new Entity("solution", Environment.OrganizationId)
                    { ["uniquename"] = "edu" });
                if (query.EntityName == "solutioncomponent") return Rows(Records.Select(r =>
                    new Entity("solutioncomponent", r.SolutionComponentId) { ["componenttype"] = new OptionSetValue(r.ComponentType),
                        ["objectid"] = r.ObjectId, ["solutionid"] = new EntityReference("solution", Environment.OrganizationId) }).ToArray());
                if (query.EntityName == "solutioncomponentdefinition")
                {
                    if (query.Criteria.Conditions.Any(c => c.AttributeName == "primaryentityname"))
                        return Rows(new Entity("solutioncomponentdefinition") { ["objecttypecode"] = 19999 });
                    if (Registration == "none") return Rows();
                    var definition = new Entity("solutioncomponentdefinition") { ["objecttypecode"] = Type,
                        ["name"] = Registration == "wrongname" ? "other" : DefinitionName,
                        ["primaryentityname"] = Registration == "wrongentity" ? "other" : PrimaryEntity };
                    definition.FormattedValues["objecttypecode"] = Label;
                    if (Registration == "conflict" || Registration == "equivalent") return Rows(definition,
                        new Entity("solutioncomponentdefinition") { ["objecttypecode"] = Type,
                            ["name"] = Registration == "conflict" ? "other" : DefinitionName.ToUpperInvariant(), ["primaryentityname"] = PrimaryEntity });
                    return Rows(definition);
                }
                var source = query.EntityName == "appsetting" ? Settings : query.EntityName == "settingdefinition" ? Definitions : Apps;
                var ids = query.Criteria.Conditions.Single().Values.Cast<Guid>().ToList();
                var rows = source.Where(r => ids.Contains(r.Id)).ToList();
                if (FailureTable == queryStage)
                {
                    if (Failure == "fault") throw new FaultException(Secret);
                    if (Failure == "null") return null;
                    if (Failure == "missing") rows.Clear();
                    if (Failure == "duplicate") rows.AddRange(rows.ToArray());
                    if (Failure == "conflicting") rows[0][query.EntityName + "id"] = Guid.NewGuid();
                    if (Failure == "missingpk") rows[0].Attributes.Remove(query.EntityName + "id");
                    if (Failure == "wrongentity") rows[0].LogicalName = "other";
                }
                var result = Rows(rows.ToArray());
                if (FailureTable == queryStage && Failure == "paged") result.MoreRecords = true;
                return result;
            }
            private OrganizationResponse Execute(OrganizationRequest request)
            {
                if (request is WhoAmIRequest) return new WhoAmIResponse { Results = new ParameterCollection { ["OrganizationId"] = Environment.OrganizationId } };
                if (request is RetrieveEntityRequest)
                {
                    if (CancelStage == "metadata") Cancel.Cancel();
                    if (SchemaMode == "metadatafault") throw new FaultException(Secret);
                    var query = (RetrieveEntityRequest)request;
                    Assert.AreEqual("settingdefinition", query.LogicalName);
                    Assert.AreEqual(EntityFilters.Entity | EntityFilters.Attributes, query.EntityFilters);
                    Assert.IsFalse(query.RetrieveAsIfPublished);
                    var fields = AppSettingResolutionOperation.StructuralFields.Select(field =>
                        field == "isoverridable" ? (AttributeMetadata)new BooleanAttributeMetadata { LogicalName = field } :
                        new PicklistAttributeMetadata { LogicalName = field }).ToList();
                    if (SchemaMode == "absentschema") fields.RemoveAll(f => f.LogicalName == "releaselevel");
                    if (SchemaMode == "wrongschematype") { fields.RemoveAll(f => f.LogicalName == "releaselevel"); fields.Add(new StringAttributeMetadata { LogicalName = "releaselevel" }); }
                    if (SchemaMode == "duplicatemetadata") fields.Add(new PicklistAttributeMetadata { LogicalName = "releaselevel" });
                    foreach (var field in fields) typeof(AttributeMetadata).GetProperty("IsValidForRead").SetValue(field,
                        !(SchemaMode == "notreadable" && field.LogicalName == "releaselevel"), null);
                    var metadata = new EntityMetadata { LogicalName = "settingdefinition" };
                    typeof(EntityMetadata).GetProperty("PrimaryIdAttribute").SetValue(metadata, "settingdefinitionid", null);
                    typeof(EntityMetadata).GetProperty("Attributes").SetValue(metadata, fields.ToArray(), null);
                    return new RetrieveEntityResponse { Results = new ParameterCollection { ["EntityMetadata"] = metadata } };
                }
                // Unknown raw types can still run existing diagnostic-only metadata probes.
                throw new FaultException("Metadata unavailable for diagnostic-only probe.");
            }
            private static EntityCollection Rows(params Entity[] rows) => new EntityCollection(rows.ToList());
        }
    }
}
