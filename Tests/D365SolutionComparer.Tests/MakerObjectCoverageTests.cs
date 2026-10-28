using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.ServiceModel;
using System.Threading;
using D365SolutionComparer.Models.Membership;
using D365SolutionComparer.Models.Presentation;
using D365SolutionComparer.Services.Membership;
using D365SolutionComparer.Services.ComponentDetails;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace D365SolutionComparer.Tests
{
    [TestClass, TestCategory("R2EMakerCoverage")]
    public class MakerObjectCoverageTests
    {
        public TestContext TestContext { get; set; }
        [DataTestMethod, DataRow("Dashboard"), DataRow("Process"), DataRow("Library"), DataRow("AI"), DataRow("Mask")]
        public void VerifiedNamesAndKindsPreserveAuthoritativeRowsAndCsv(string kind)
        {
            var f = ExplorerFixture.Make(new[] { Identity(kind) }); var csv = ExplorerFixture.Csv(f);
            var original = f.Presentation.Rows[0]; var cache = Prepare(f.Presentation.Source.Snapshot, out var service);
            new MakerObjectDisplayReader().Read(service.Context, f.Presentation.Source.Snapshot, CancellationToken.None);
            var projection = Project(f, cache); var item = new SolutionObjectNavigation(projection).Items(ObjectsWorkspaceScope.Source).Single();
            Assert.AreEqual(Category(kind), item.Category); Assert.AreEqual("Same retained name", item.Value(false));
            Assert.AreSame(original, projection.Rows[0].Original.AuthoritativeRow); Assert.AreEqual(csv, ExplorerFixture.Csv(f));
            Assert.AreEqual(IdentityResolutionStatus.Unsupported, original.SourceIdentity.Status); Assert.IsNull(original.SourceIdentity.ComparisonKey);
            Assert.AreEqual("Needs Review", item.Status(projection)); Assert.AreEqual("Not established", item.Presence(DisplayEnvironmentSide.Target, projection));
            Assert.AreEqual(kind == "AI" || kind == "Mask" ? 1 : 0, service.Queries.Count);
            Assert.AreEqual(0, service.Writes);
        }

        [DataTestMethod, DataRow("Dashboard"), DataRow("Process"), DataRow("Library"), DataRow("AI"), DataRow("Mask")]
        public void LocalAliasesDeduplicateButDistinctEqualNamesRemainSeparate(string kind)
        {
            var id = Guid.NewGuid(); var f = ExplorerFixture.Make(new[] { Identity(kind, id), Identity(kind, id), Identity(kind) });
            var cache = Prepare(f.Presentation.Source.Snapshot, out var service); Read(service, f.Presentation.Source.Snapshot);
            var p = Project(f, cache); var items = new SolutionObjectNavigation(p).Items(ObjectsWorkspaceScope.Source);
            Assert.AreEqual(2, items.Count); Assert.AreEqual(3, p.Rows.Count); Assert.AreEqual(3, p.Occurrences.Count);
            Assert.AreEqual(2, items.Single(i => i.Locals[0].Id == id).Rows.Count); Assert.AreEqual(2, items.Select(i => i.Key).Distinct().Count());
            Assert.IsTrue(items.All(i => i.Value(false) == "Same retained name"));
        }

        [DataTestMethod, DataRow("Dashboard"), DataRow("Process"), DataRow("Library"), DataRow("AI"), DataRow("Mask")]
        public void EachSideOwnsItsObjectsAndUnsupportedCounterpartsNeverBecomeMatch(string kind)
        {
            var id = Guid.NewGuid(); var f = ExplorerFixture.Make(new[] { Identity(kind, id) }, new[] { Identity(kind, id) });
            var left = Prepare(f.Presentation.Source.Snapshot, out var s); var right = Prepare(f.Presentation.Target.Snapshot, out var t);
            Read(s, f.Presentation.Source.Snapshot); Read(t, f.Presentation.Target.Snapshot);
            var p = Project(f, left, right); var model = new SolutionObjectNavigation(p);
            Assert.AreEqual(1, model.Items(ObjectsWorkspaceScope.Source).Count); Assert.AreEqual(1, model.Items(ObjectsWorkspaceScope.Target).Count);
            Assert.AreEqual(2, model.Items(ObjectsWorkspaceScope.Compare).Count);
            Assert.IsTrue(model.Items(ObjectsWorkspaceScope.Source).Single().Has(DisplayEnvironmentSide.Source));
            Assert.IsFalse(model.Items(ObjectsWorkspaceScope.Source).Single().Has(DisplayEnvironmentSide.Target));
            Assert.IsTrue(p.Rows.All(r => r.Original.AuthoritativeRow.Presence == MembershipPresence.Indeterminate));
        }

        [DataTestMethod, DataRow("Dashboard"), DataRow("Process"), DataRow("Library"), DataRow("AI"), DataRow("Mask")]
        public void OwnerRecordAndSideMismatchesRejectTheWholeObservation(string kind)
        {
            var identity = Identity(kind); var f = ExplorerFixture.Make(new[] { identity }); var owner = f.Presentation.Source.Snapshot;
            var cache = Prepare(owner, out var s); Read(s, owner); MakerObjectDisplayEvidence value;
            Assert.IsTrue(cache.TryGetMakerObjectDisplay(owner, identity, out value));
            var copy = MembershipSnapshot.Complete(owner.Solution, owner.Components, owner.CapturedAt);
            Assert.IsFalse(cache.TryGetMakerObjectDisplay(copy, identity, out value));
            Assert.IsFalse(cache.TryGetMakerObjectDisplay(owner, Identity(kind, identity.Record.ObjectId), out value));
            var good = ComponentDisplayEvidenceProjector.Project(f.Presentation, cache, null)[0].SourceDisplayEvidence;
            var foreign = new ComponentSideDisplayEvidence(DisplayEnvironmentSide.Target, good.SolutionComponentId, good.LocalObjectId,
                good.RawComponentType, good.ComponentCategory, good.FriendlyName, good.InternalName, good.BackingCorrelationStatus,
                good.ParentDisplayContext, makerObject: good.MakerObject);
            f.Completed = new CompletedMembershipComparison(f.Presentation, new[] { new ComponentDisplayEvidence(0, foreign, null) }, DisplayProjectionOutcome.Completed);
            Assert.AreEqual(0, new SolutionObjectNavigation(ExplorerFixture.Project(f)).Items(ObjectsWorkspaceScope.Source).Count);
        }

        [DataTestMethod, DataRow(0, true), DataRow(10, true), DataRow(13, true), DataRow(103, true)]
        [DataRow(2, false), DataRow(7, false), DataRow(999, false), DataRow(-1, false)]
        public void DashboardSubtypeIsRequiredAndOtherFormsAreNotDashboardObjects(int type, bool accepted)
        {
            var f = ExplorerFixture.Make(new[] { Identity("Dashboard") }); var cache = Prepare(f.Presentation.Source.Snapshot, out var s);
            var id = f.Presentation.Source.Snapshot.Components[0].Record.ObjectId.Value; Entity row;
            Assert.IsTrue(cache.TryGetEntityRow("systemform", id, out row));
            if (type < 0) row.Attributes.Remove("type"); else row["type"] = new OptionSetValue(type);
            Read(s, f.Presentation.Source.Snapshot);
            Assert.AreEqual(accepted ? 1 : 0, new SolutionObjectNavigation(Project(f, cache)).Items(ObjectsWorkspaceScope.Source).Count);
            Assert.AreEqual(0, s.Queries.Count); Assert.AreEqual(0, s.SchemaCalls);
        }

        [DataTestMethod, DataRow(1, true), DataRow(0, false), DataRow(2, false), DataRow(3, false)]
        [DataRow(4, false), DataRow(5, false), DataRow(999, false), DataRow(-1, false)]
        public void OnlyTypedCanvasLibraryDiscriminatorClassifiesAComponentLibrary(int type, bool accepted)
        {
            var f = ExplorerFixture.Make(new[] { Identity("Library") }); var cache = Prepare(f.Presentation.Source.Snapshot, out var s, canvasType: type);
            Read(s, f.Presentation.Source.Snapshot); var p = Project(f, cache);
            Assert.AreEqual(accepted ? 1 : 0, new SolutionObjectNavigation(p).Items(ObjectsWorkspaceScope.Source).Count);
            Assert.AreEqual("NormalMembershipDiagnostic.canvasapp.displayname", p.Rows[0].SourceDisplay.FriendlyName.DisplayEvidenceSource);
            Assert.AreEqual(0, s.Queries.Count); Assert.AreEqual(0, s.SchemaCalls);
        }

        [TestMethod]
        public void MalformedCanvasDiscriminatorDoesNotDiscardRetainedFriendlyNames()
        {
            var f = ExplorerFixture.Make(new[] { Identity("Library") }); var cache = Prepare(f.Presentation.Source.Snapshot, out var s, malformedCanvas: true);
            Read(s, f.Presentation.Source.Snapshot); var p = Project(f, cache);
            Assert.AreEqual(0, new SolutionObjectNavigation(p).Items(ObjectsWorkspaceScope.Source).Count);
            Assert.AreEqual("Same retained name", p.Rows[0].SourceDisplay.FriendlyName.Value);
        }

        [DataTestMethod, DataRow(0, true), DataRow(1, true), DataRow(3, true), DataRow(4, true)]
        [DataRow(2, false), DataRow(5, false), DataRow(6, false), DataRow(7, false), DataRow(999, false), DataRow(-1, false)]
        public void ProcessCategoriesRemainDistinctFromRulesAndFlows(int category, bool accepted)
        {
            var f = ExplorerFixture.Make(new[] { Identity("Process") }); var cache = Prepare(f.Presentation.Source.Snapshot, out var s);
            var row = cache.WorkflowDefinitions.Values.Single(); if (category < 0) row.Attributes.Remove("category"); else row["category"] = new OptionSetValue(category);
            Read(s, f.Presentation.Source.Snapshot);
            Assert.AreEqual(accepted ? 1 : 0, new SolutionObjectNavigation(Project(f, cache)).Items(ObjectsWorkspaceScope.Source).Count);
            Assert.AreEqual(0, s.Queries.Count); Assert.AreEqual(0, s.SchemaCalls);
        }

        [DataTestMethod, DataRow(true), DataRow(false)]
        public void ProcessAliasesRequireTheExistingVerifiedDefinitionRelationship(bool verified)
        {
            var definition = Identity("Process"); var activation = Identity("Process"); var f = ExplorerFixture.Make(new[] { definition, activation });
            var cache = Prepare(f.Presentation.Source.Snapshot, out var s); var raw = activation.Record.ObjectId.Value;
            cache.WorkflowDefinitions[raw] = cache.WorkflowDefinitions[definition.Record.ObjectId.Value];
            if (verified) cache.WorkflowDisplayAliases[raw] = definition.Record.ObjectId.Value;
            Read(s, f.Presentation.Source.Snapshot); var p = Project(f, cache); var items = new SolutionObjectNavigation(p).Items(ObjectsWorkspaceScope.Source);
            Assert.AreEqual(1, items.Count); Assert.AreEqual(verified ? 2 : 1, items[0].Rows.Count);
            Assert.AreEqual(2, p.Rows.Count); Assert.AreEqual(2, p.Occurrences.Count);
            Assert.AreEqual(definition.Record.ObjectId.Value, items[0].Locals[0].Id);
        }

        [DataTestMethod, DataRow("Dashboard"), DataRow("Process")]
        public void ExistingAuthoritativeSharedMatchesRemainTheOnlyCrossEnvironmentJoin(string kind)
        {
            var f = ExplorerFixture.Make(new[] { Identity(kind, key: "shared") }, new[] { Identity(kind, key: "shared") });
            var a = Prepare(f.Presentation.Source.Snapshot, out var s); var b = Prepare(f.Presentation.Target.Snapshot, out var t);
            Read(s, f.Presentation.Source.Snapshot); Read(t, f.Presentation.Target.Snapshot); var p = Project(f, a, b);
            var shared = new SolutionObjectNavigation(p).Items(ObjectsWorkspaceScope.Compare).Single();
            Assert.IsTrue(shared.Has(DisplayEnvironmentSide.Source) && shared.Has(DisplayEnvironmentSide.Target));
            Assert.AreEqual(MembershipPresence.PresentInBoth, p.Rows[0].Original.AuthoritativeRow.Presence);
        }

        [DataTestMethod]
        [DataRow("AI", "missing"), DataRow("AI", "duplicate"), DataRow("AI", "wrong-id"), DataRow("AI", "wrong-table")]
        [DataRow("AI", "incomplete"), DataRow("AI", "fault"), DataRow("AI", "blank"), DataRow("AI", "unrequested")]
        [DataRow("Mask", "missing"), DataRow("Mask", "duplicate"), DataRow("Mask", "wrong-id"), DataRow("Mask", "wrong-table")]
        [DataRow("Mask", "incomplete"), DataRow("Mask", "fault"), DataRow("Mask", "blank"), DataRow("Mask", "unrequested")]
        public void RegisteredBackingFailuresCannotFabricateObjectsOrAbsence(string kind, string failure)
        {
            var f = ExplorerFixture.Make(new[] { Identity(kind) }); var csv = ExplorerFixture.Csv(f); var cache = Prepare(f.Presentation.Source.Snapshot, out var s);
            s.DataFailure = failure; Read(s, f.Presentation.Source.Snapshot); var p = Project(f, cache);
            Assert.AreEqual(0, new SolutionObjectNavigation(p).Items(ObjectsWorkspaceScope.Source).Count);
            Assert.AreEqual(1, p.Rows.Count); Assert.AreEqual(csv, ExplorerFixture.Csv(f)); Assert.IsTrue(cache.MakerDisplayLimitations.Count > 0);
            Assert.AreEqual(MembershipPresence.Indeterminate, p.Rows[0].Original.AuthoritativeRow.Presence);
        }

        [DataTestMethod]
        [DataRow("AI", "wrong-table"), DataRow("AI", "wrong-id-type"), DataRow("AI", "name-memo"), DataRow("AI", "name-unreadable")]
        [DataRow("AI", "id-unreadable"), DataRow("AI", "secured-name"), DataRow("AI", "no-primary")]
        [DataRow("Mask", "wrong-table"), DataRow("Mask", "wrong-id-type"), DataRow("Mask", "name-memo"), DataRow("Mask", "name-unreadable")]
        [DataRow("Mask", "id-unreadable"), DataRow("Mask", "secured-name"), DataRow("Mask", "no-primary")]
        public void RegisteredSchemaMustVerifyReadableIdentityAndNameBeforeAnyDataQuery(string kind, string failure)
        {
            var f = ExplorerFixture.Make(new[] { Identity(kind) }); var cache = Prepare(f.Presentation.Source.Snapshot, out var s);
            s.SchemaFailure = failure; Read(s, f.Presentation.Source.Snapshot);
            Assert.AreEqual(0, s.Queries.Count); Assert.AreEqual(1, s.SchemaCalls);
            Assert.AreEqual(0, new SolutionObjectNavigation(Project(f, cache)).Items(ObjectsWorkspaceScope.Source).Count);
        }

        [TestMethod]
        public void OneCategoryFailureDoesNotSuppressIndependentVerifiedCategories()
        {
            var f = ExplorerFixture.Make(new[] { Identity("AI"), Identity("Mask"), Identity("Dashboard"), Identity("Process"), Identity("Library") });
            var cache = Prepare(f.Presentation.Source.Snapshot, out var s); s.FailAiSchemaOnly = true; Read(s, f.Presentation.Source.Snapshot);
            var items = new SolutionObjectNavigation(Project(f, cache)).Items(ObjectsWorkspaceScope.Source);
            Assert.AreEqual(4, items.Count); Assert.IsFalse(items.Any(i => i.Category == "AI Skill Config"));
        }

        [DataTestMethod, DataRow("before"), DataRow("schema"), DataRow("query")]
        public void CancellationDiscardsOptionalPublicationWithoutFurtherRequests(string checkpoint)
        {
            using (var cancel = new CancellationTokenSource())
            {
                var f = ExplorerFixture.Make(new[] { Identity("AI") }); var cache = Prepare(f.Presentation.Source.Snapshot, out var s, token: cancel.Token);
                if (checkpoint == "before") cancel.Cancel();
                else if (checkpoint == "schema") s.AfterSchema = cancel.Cancel; else s.AfterQuery = cancel.Cancel;
                Assert.ThrowsException<OperationCanceledException>(() => new MakerObjectDisplayReader().Read(s.Context, f.Presentation.Source.Snapshot, cancel.Token));
                MakerObjectDisplayEvidence value; Assert.IsFalse(cache.TryGetMakerObjectDisplay(f.Presentation.Source.Snapshot, f.Presentation.Source.Snapshot.Components[0], out value));
                Assert.AreEqual(checkpoint == "query" ? 1 : 0, s.Queries.Count);
            }
        }

        [DataTestMethod, DataRow("AI"), DataRow("Mask")]
        public void RegisteredReadsUseDiscoveredColumnsExactIdsAndTwoHundredIdBatches(string kind)
        {
            var identities = Enumerable.Range(0, 201).Select(i => Identity(kind)).ToList(); identities.Add(Identity(kind, identities[0].Record.ObjectId));
            var f = ExplorerFixture.Make(identities); var cache = Prepare(f.Presentation.Source.Snapshot, out var s); Read(s, f.Presentation.Source.Snapshot);
            Assert.AreEqual(1, s.SchemaCalls); Assert.AreEqual(2, s.Queries.Count);
            CollectionAssert.AreEquivalent(identities.Select(i => i.Record.ObjectId.Value).Distinct().ToArray(), s.Queries.SelectMany(q => q.Criteria.Conditions.Single().Values.Cast<Guid>()).ToArray());
            Assert.IsTrue(s.Queries.All(q => q.Criteria.Conditions.Count == 1 && q.Criteria.Conditions[0].Operator == ConditionOperator.In && q.Criteria.Conditions[0].Values.Count <= 200));
            Assert.IsTrue(s.Queries.All(q => q.ColumnSet.Columns.Count == 2 && q.ColumnSet.Columns.Contains("recordkey") && q.ColumnSet.Columns.Contains("caption")));
            Assert.AreEqual(201, new SolutionObjectNavigation(Project(f, cache)).Items(ObjectsWorkspaceScope.Source).Count);
        }

        [DataTestMethod, DataRow("complete"), DataRow("partial"), DataRow("conflicting")]
        public void RegisteredSchemaReuseRequiresCompleteUnambiguousReadableMetadata(string state)
        {
            var f = ExplorerFixture.Make(new[] { Identity("AI") }); var cache = Prepare(f.Presentation.Source.Snapshot, out var service);
            var schema = state == "partial" ? new EntityMetadata { LogicalName = "fixture_ai" } : Backend.Schema("fixture_ai", null);
            schema.MetadataId = Guid.NewGuid(); cache.Store(schema);
            if (state == "conflicting") { var second = Backend.Schema("fixture_ai", null); second.MetadataId = Guid.NewGuid(); cache.Store(second); }
            Read(service, f.Presentation.Source.Snapshot);
            Assert.AreEqual(state == "partial" ? 1 : 0, service.SchemaCalls);
            Assert.AreEqual(state == "conflicting" ? 0 : 1, service.Queries.Count);
            Assert.AreEqual(state == "conflicting" ? 0 : 1, new SolutionObjectNavigation(Project(f, cache)).Items(ObjectsWorkspaceScope.Source).Count);
        }

        [TestMethod]
        public void TenThousandCachedDashboardsRequireNoSchemaOrDataRequests()
        {
            var f = ExplorerFixture.Make(Enumerable.Range(0, 10000).Select(i => Identity("Dashboard"))); var cache = Prepare(f.Presentation.Source.Snapshot, out var s);
            var watch = Stopwatch.StartNew(); Read(s, f.Presentation.Source.Snapshot);
            Assert.AreEqual(10000, new SolutionObjectNavigation(Project(f, cache)).Items(ObjectsWorkspaceScope.Source).Count);
            Assert.AreEqual(0, s.Queries.Count); Assert.AreEqual(0, s.SchemaCalls); TestContext.WriteLine("10000 cached dashboard objects elapsed(ms)=" + watch.ElapsedMilliseconds);
        }

        [TestMethod]
        public void ControlledHistoricalCoverageFixtureKeepsEightHundredFortyTechnicalRows()
        {
            var source = Enumerable.Range(0, 133).Select(i => ExplorerFixture.Identity(61, "baseline" + i)).ToList();
            source.AddRange(new[] { Identity("AI"), Identity("AI"), Identity("Library"), Identity("Process"), Identity("Mask") });
            source.AddRange(Enumerable.Range(0, 8).Select(i => Identity("Dashboard")));
            source.AddRange(Enumerable.Range(0, 694).Select(i => ExplorerFixture.Identity(2, "technical-column" + i)));
            var f = ExplorerFixture.Make(source); var csv = ExplorerFixture.Csv(f); var rows = f.Presentation.Rows.ToArray();
            var cache = Prepare(f.Presentation.Source.Snapshot, out var s); Read(s, f.Presentation.Source.Snapshot); var p = Project(f, cache);
            var items = new SolutionObjectNavigation(p).Items(ObjectsWorkspaceScope.Source);
            Assert.AreEqual(146, items.Count); Assert.AreEqual(840, p.Rows.Count); Assert.AreEqual(840, p.Occurrences.Count);
            Assert.AreEqual(8, items.Count(i => i.Category == "Dashboards")); Assert.AreEqual(2, items.Count(i => i.Category == "AI Skill Config"));
            CollectionAssert.AreEqual(rows, p.Rows.Select(r => r.Original.AuthoritativeRow).ToArray()); Assert.AreEqual(csv, ExplorerFixture.Csv(f));
            TestContext.WriteLine("CONTROLLED FIXTURE only: 133 baseline objects + 13 verified objects = 146; 840 technical rows unchanged; no current SAA inventory claim.");
        }

        [DataTestMethod, DataRow(false), DataRow(true)]
        public void ProductionCompositionOptsIntoBoundedDisplayReadWithoutChangingIdentity(bool production)
        {
            var identity = Identity("AI"); var s = new Backend { PipelineIdentity = identity };
            var operation = production ? SolutionComparerControl.CreateMembershipComparisonOperation() : new DataverseComponentDefinitionOperation();
            var result = operation.ReadAndResolveWithCache(s, "Offline", "SAA", CancellationToken.None, null);
            MakerObjectDisplayEvidence value; var actual = result.Definitions.Membership.Components.Single();
            Assert.AreEqual(production, result.Cache.TryGetMakerObjectDisplay(result.Definitions.Membership, actual, out value));
            Assert.AreEqual(IdentityResolutionStatus.Unsupported, actual.Status); Assert.IsNull(actual.ComparisonKey);
            Assert.AreEqual(production ? 1 : 0, s.SchemaCalls); Assert.AreEqual(production ? 1 : 0, s.Queries.Count); Assert.AreEqual(0, s.Writes);
        }

        private static ComponentIdentity Identity(string kind, Guid? id = null, string key = null)
        {
            int raw = kind == "Dashboard" ? 60 : kind == "Process" ? 29 : kind == "Library" ? 300 : kind == "AI" ? 10276 : 74;
            var definition = kind == "AI" || kind == "Mask" ? new SolutionComponentDefinitionIdentity(raw, kind == "AI" ? "AiSkillConfig" : "MaskingRule", kind == "AI" ? "fixture_ai" : "fixture_mask") : null;
            return new ComponentIdentity(new SolutionComponentRecord(Guid.NewGuid(), raw, id ?? Guid.NewGuid()),
                key == null ? IdentityResolutionStatus.Unsupported : IdentityResolutionStatus.Resolved, key,
                semanticKind: kind == "Dashboard" ? ComponentSemanticKinds.SystemForm : kind == "Process" ? ComponentSemanticKinds.Process : null,
                registeredDefinition: definition);
        }
        private static string Category(string kind) => kind == "Dashboard" ? "Dashboards" : kind == "Process" ? "Processes" : kind == "Library" ? "Component libraries" : kind == "AI" ? "AI Skill Config" : "Secured Masking Rule";
        private static void Read(Backend service, MembershipSnapshot owner) => new MakerObjectDisplayReader().Read(service.Context, owner, CancellationToken.None);
        private static SolutionObjectsExplorerProjection Project(ExplorerFixture f, DataverseComponentMetadataCache source, DataverseComponentMetadataCache target = null)
        {
            f.Completed = new CompletedMembershipComparison(f.Presentation, ComponentDisplayEvidenceProjector.Project(f.Presentation, source, target), DisplayProjectionOutcome.Completed);
            return ExplorerFixture.Project(f);
        }
        private static DataverseComponentMetadataCache Prepare(MembershipSnapshot owner, out Backend service, int canvasType = 1, bool malformedCanvas = false, CancellationToken token = default(CancellationToken))
        {
            service = new Backend { Organization = owner.Environment.OrganizationId };
            service.Context = new DataverseReadContext(service, owner.Environment, token); var cache = service.Context.MetadataCache; cache.BeginNormalDisplayNames(owner);
            var canvas = new Dictionary<Guid, List<Entity>>();
            foreach (var identity in owner.Components)
            {
                Guid id = identity.Record.ObjectId.Value; int type = identity.Record.ComponentType;
                var row = new Entity(type == 60 ? "systemform" : type == 29 ? "workflow" : type == 300 ? "canvasapp" : type == 61 ? "webresource" : type == 10276 ? "fixture_ai" : "fixture_mask", id)
                { [type == 60 ? "formid" : type == 29 ? "workflowid" : type == 300 ? "canvasappid" : "recordkey"] = id, ["name"] = "Same retained name", ["uniquename"] = "verified_unique", ["type"] = new OptionSetValue(type == 60 ? 0 : 1), ["category"] = new OptionSetValue(0), ["caption"] = "Same retained name", ["displayname"] = "Same retained name" };
                if (type == 61) { row["webresourceid"] = id; row["displayname"] = "Offline baseline object " + id; cache.StoreEntityRow(row, id); }
                else if (type == 60) cache.StoreEntityRow(row, id);
                else if (type == 29) cache.WorkflowDefinitions[id] = row;
                else if (type == 300)
                {
                    if (canvasType >= 0) row["canvasapptype"] = malformedCanvas ? (object)"1" : new OptionSetValue(canvasType);
                    canvas[id] = new List<Entity> { row };
                }
                else if (type == 10276 || type == 74) service.Rows[row.LogicalName + id] = row;
            }
            if (canvas.Count > 0) cache.ObserveCanvasAppDisplayNames(new DiagnosticQueryResult(canvas.Keys.ToList(), canvas,
                new Dictionary<Guid, string>(), canvas.Keys.ToDictionary(id => id, id => new List<Entity>()), canvas.Count, true));
            cache.SealNormalDisplayNames(owner); return cache;
        }
        private static void Set(object instance, string field, object value) => instance.GetType().GetProperty(field).SetValue(instance, value, null);
        private sealed class Backend : IOrganizationService
        {
            internal Guid Organization = Guid.NewGuid(), Solution = Guid.NewGuid();
            internal DataverseReadContext Context;
            internal readonly Dictionary<string, Entity> Rows = new Dictionary<string, Entity>();
            internal readonly List<QueryExpression> Queries = new List<QueryExpression>();
            internal string DataFailure, SchemaFailure;
            internal int SchemaCalls, Writes;
            internal bool FailAiSchemaOnly;
            internal Action AfterSchema, AfterQuery;
            internal ComponentIdentity PipelineIdentity;
            public OrganizationResponse Execute(OrganizationRequest request)
            {
                if (request is WhoAmIRequest) { var who = new WhoAmIResponse(); who.Results["OrganizationId"] = Organization; return who; }
                var metadata = request as RetrieveEntityRequest;
                if (metadata == null) throw new InvalidOperationException("Unexpected offline request " + request.RequestName);
                SchemaCalls++; Assert.IsFalse(metadata.RetrieveAsIfPublished);
                Assert.AreEqual(EntityFilters.Entity | EntityFilters.Attributes, metadata.EntityFilters);
                var schema = Schema(metadata.LogicalName, FailAiSchemaOnly && metadata.LogicalName == "fixture_ai" ? "wrong-table" : SchemaFailure);
                var response = new RetrieveEntityResponse(); response.Results["EntityMetadata"] = schema; AfterSchema?.Invoke(); return response;
            }
            public EntityCollection RetrieveMultiple(QueryBase query)
            {
                var q = (QueryExpression)query;
                if (q.EntityName == "solution") return Collection(new Entity("solution", Solution) { ["solutionid"] = Solution, ["uniquename"] = "SAA" });
                if (q.EntityName == "solutioncomponent") return Collection(new Entity("solutioncomponent", PipelineIdentity.Record.SolutionComponentId)
                { ["solutioncomponentid"] = PipelineIdentity.Record.SolutionComponentId, ["componenttype"] = new OptionSetValue(10276), ["objectid"] = PipelineIdentity.Record.ObjectId.Value });
                if (q.EntityName == "solutioncomponentdefinition") return q.Criteria.Conditions.Any(c => c.AttributeName == "primaryentityname") ? Collection() :
                    Collection(new Entity("solutioncomponentdefinition", Guid.NewGuid()) { ["objecttypecode"] = 10276, ["name"] = "AiSkillConfig", ["primaryentityname"] = "fixture_ai" });
                Assert.IsTrue(q.EntityName == "fixture_ai" || q.EntityName == "fixture_mask"); Queries.Add(q);
                CollectionAssert.AreEquivalent(new[] { "recordkey", "caption" }, q.ColumnSet.Columns.ToArray());
                var ids = q.Criteria.Conditions.Single().Values.Cast<Guid>().ToArray(); Assert.IsTrue(ids.Length > 0 && ids.Length <= 200);
                var data = ids.Select(id => Rows.TryGetValue(q.EntityName + id, out var r) ? r : new Entity(q.EntityName, id) { ["recordkey"] = id, ["caption"] = "Same retained name" }).ToList();
                if (DataFailure == "fault") throw new FaultException("Synthetic failure; not retained");
                if (DataFailure == "missing") data.Clear();
                if (DataFailure == "duplicate") data.Add(data[0]);
                if (DataFailure == "wrong-id") data[0]["recordkey"] = Guid.NewGuid();
                if (DataFailure == "wrong-table") data[0].LogicalName = "foreign_table";
                if (DataFailure == "blank") data[0]["caption"] = "";
                if (DataFailure == "unrequested") data.Add(new Entity(q.EntityName, Guid.NewGuid()) { ["recordkey"] = Guid.NewGuid(), ["caption"] = "foreign" });
                var result = new EntityCollection(data) { MoreRecords = DataFailure == "incomplete" }; AfterQuery?.Invoke(); return result;
            }
            internal static EntityMetadata Schema(string table, string failure)
            {
                var primary = failure == "wrong-id-type" ? (AttributeMetadata)new StringAttributeMetadata() : new UniqueIdentifierAttributeMetadata(); primary.LogicalName = "recordkey";
                var name = failure == "name-memo" ? (AttributeMetadata)new MemoAttributeMetadata() : new StringAttributeMetadata(); name.LogicalName = "caption";
                Set(primary, "IsValidForRead", failure != "id-unreadable"); Set(name, "IsValidForRead", failure != "name-unreadable"); Set(name, "IsSecured", failure == "secured-name");
                var schema = new EntityMetadata { LogicalName = failure == "wrong-table" ? "foreign_table" : table };
                Set(schema, "PrimaryIdAttribute", failure == "no-primary" ? null : "recordkey"); Set(schema, "PrimaryNameAttribute", "caption");
                Set(schema, "Attributes", new[] { primary, name }); return schema;
            }
            private static EntityCollection Collection(params Entity[] rows) => new EntityCollection(rows.ToList());
            public Guid Create(Entity entity) { Writes++; throw new InvalidOperationException(); }
            public void Update(Entity entity) { Writes++; throw new InvalidOperationException(); }
            public void Delete(string entityName, Guid id) { Writes++; throw new InvalidOperationException(); }
            public Entity Retrieve(string entityName, Guid id, ColumnSet columns) => throw new InvalidOperationException();
            public void Associate(string entityName, Guid id, Relationship relation, EntityReferenceCollection rows) { Writes++; throw new InvalidOperationException(); }
            public void Disassociate(string entityName, Guid id, Relationship relation, EntityReferenceCollection rows) { Writes++; throw new InvalidOperationException(); }
        }
    }
}
