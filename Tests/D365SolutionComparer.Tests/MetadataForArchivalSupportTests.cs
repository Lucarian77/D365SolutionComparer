using System;
using System.Collections.Generic;
using System.Linq;
using System.ServiceModel;
using System.Threading;
using D365SolutionComparer.Models.Identity;
using D365SolutionComparer.Models.Membership;
using D365SolutionComparer.Models.Presentation;
using D365SolutionComparer.Models.ComponentDetails;
using D365SolutionComparer.Services.Membership;
using D365SolutionComparer.Services.ComponentDetails;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Metadata.Query;
using Microsoft.Xrm.Sdk.Query;

namespace D365SolutionComparer.Tests
{
    [TestClass]
    public class MetadataForArchivalSupportTests
    {
        [TestMethod]
        public void OwnerUatIdsResolveIndependentlyFromTheirTables()
        {
            var f = new Fixture(); f.Add(0, "mnp_applicationparticipant"); f.Add(1, "mnp_registration"); f.AddTables(); f.Resolve();
            Assert.AreEqual(4, f.Snapshot.Components.Count);
            Assert.AreEqual(2, f.Snapshot.Components.Count(x => x.Record.ComponentType == 10139));
            Assert.IsTrue(f.Snapshot.Components.All(x => x.Status == IdentityResolutionStatus.Resolved));
            Assert.AreEqual(4, f.Snapshot.Components.Select(x => x.ComponentTypeKey + x.ComparisonKey).Distinct().Count());
            var projection = Project(f); var items = new SolutionObjectNavigation(projection).Items(ObjectsWorkspaceScope.Source);
            Assert.AreEqual(2, items.Count(x => x.Category == "MetadataForArchival")); Assert.AreEqual(2, items.Count(x => x.Category == "Tables"));
            Assert.IsTrue(items.Where(x => x.Category == "MetadataForArchival").All(x => x.Type == "MetadataForArchival"));
            Assert.AreEqual(Guid.Parse("e7c04ff2-4d43-4234-83b9-9af755a5bf24"), f.Rows[0].Id);
            Assert.AreEqual(Guid.Parse("85f5fa0f-8997-4574-b524-63112271bba7"), f.Rows[1].Id);
            Assert.AreEqual(0, f.Service.WriteCalls);
        }

        [DataTestMethod, DataRow(false), DataRow(true)]
        public void EachSideOwnsItsNamesIdsAndInformationalValues(bool target)
        {
            var f = new Fixture(target); f.Add(0, "mnp_registration"); f.Resolve(); var p = Project(f, target);
            var info = target ? p.Rows[0].TargetDisplay.ArchivalInformation : p.Rows[0].SourceDisplay.ArchivalInformation;
            Assert.IsNotNull(info); Assert.AreEqual(target ? DisplayEnvironmentSide.Target : DisplayEnvironmentSide.Source, info.Side);
            Assert.AreEqual(f.Rows[0].Id, info.ObjectId); Assert.AreEqual("False", info.Field("Archival readiness").Value);
            StringAssert.Contains(p.Rows[0].ComparedPropertiesExplanation, "not compared");
            StringAssert.Contains(p.Rows[0].ComparedPropertiesExplanation, "mnp_registration");
            Assert.IsNull(target ? p.Rows[0].SourceDisplay : p.Rows[0].TargetDisplay);
            Assert.IsNull(target ? p.Rows[0].SourceDisplay?.ArchivalInformation : p.Rows[0].TargetDisplay?.ArchivalInformation);
            Assert.AreEqual("MetadataForArchival", new SolutionObjectNavigation(p).Items(target ? ObjectsWorkspaceScope.Target : ObjectsWorkspaceScope.Source).Single().Category);
        }

        [TestMethod]
        public void DifferentEnvironmentGuidsMatchOnlyVerifiedLogicalIdentity()
        {
            var source = new Fixture(); var target = new Fixture(true); source.Add(0, "mnp_registration"); target.Add(0, "mnp_registration"); source.Resolve(); target.Resolve();
            Assert.AreNotEqual(source.Rows[0].Id, target.Rows[0].Id); Assert.AreNotEqual(source.TableRows.Keys.Single(), target.TableRows.Keys.Single());
            var row = new SolutionMembershipComparer().Compare(source.Snapshot, target.Snapshot).Single();
            Assert.AreEqual(MembershipPresence.PresentInBoth, row.Presence);
            Assert.AreEqual(source.Snapshot.Components[0].ComparisonKey, target.Snapshot.Components[0].ComparisonKey);
            Assert.AreSame(source.Snapshot.Components[0], row.Source); Assert.AreSame(target.Snapshot.Components[0], row.Target);
        }

        [DataTestMethod, DataRow(false), DataRow(true)]
        public void EmptyCompleteOppositeInventoryAloneEstablishesMissingSide(bool targetOnly)
        {
            var f = new Fixture(targetOnly); f.Add(0, "mnp_registration"); f.Resolve(); var empty = Empty();
            var row = new SolutionMembershipComparer().Compare(targetOnly ? empty : f.Snapshot, targetOnly ? f.Snapshot : empty).Single();
            Assert.AreEqual(targetOnly ? MembershipPresence.OnlyInTarget : MembershipPresence.OnlyInSource, row.Presence);
        }

        [DataTestMethod, DataRow("query-failure"), DataRow("incomplete"), DataRow("missing-row"), DataRow("duplicate-row"), DataRow("wrong-primary")]
        [DataRow("schema-failure"), DataRow("bad-schema"), DataRow("missing-name"), DataRow("wrong-relation"), DataRow("missing-relation"), DataRow("table-failure")]
        public void FailedOrUnverifiedCounterpartNeverBecomesAbsent(string failure)
        {
            var source = new Fixture(); var target = new Fixture(true) { Failure = failure };
            source.Add(0, "mnp_registration"); target.Add(0, "mnp_registration"); source.Resolve(); target.Resolve();
            Assert.AreNotEqual(IdentityResolutionStatus.Resolved, target.Snapshot.Components[0].Status, failure);
            var compared = new SolutionMembershipComparer().Compare(source.Snapshot, target.Snapshot);
            Assert.IsTrue(compared.All(x => x.Presence == MembershipPresence.Indeterminate), failure);
            Assert.AreEqual(0, target.Service.WriteCalls);
        }

        [DataTestMethod, DataRow("ismanaged"), DataRow("isavailableforarchival"), DataRow("isreadyforarchival")]
        public void OptionalBooleanOmissionIsNotExplicitFalse(string field)
        {
            var f = new Fixture(); f.Add(0, "mnp_registration"); f.Rows[0].Attributes.Remove(field); f.Resolve();
            var info = Project(f).Rows[0].SourceDisplay.ArchivalInformation;
            string label = field == "ismanaged" ? "Managed" : field == "isavailableforarchival" ? "Archival availability" : "Archival readiness";
            Assert.AreEqual(DisplayAvailability.AttributeOmitted, info.Field(label).Availability); Assert.IsNull(info.Field(label).Value);
            Assert.AreEqual(IdentityResolutionStatus.Resolved, f.Snapshot.Components[0].Status);
        }

        [TestMethod]
        public void SchemaAbsentOptionalFieldsAreNotRequestedOrDefaulted()
        {
            var f = new Fixture { OmitOptionalSchema = true }; f.Add(0, "mnp_registration"); f.Resolve();
            var info = Project(f).Rows[0].SourceDisplay.ArchivalInformation;
            Assert.AreEqual(DisplayAvailability.NotRetrieved, info.Field("Managed").Availability);
            Assert.IsFalse(f.DataQueries.Single().ColumnSet.Columns.Contains("ismanaged"));
            Assert.AreEqual("False", info.Field("Archival readiness").Value);
        }

        [TestMethod]
        public void DuplicateOwnNamesStaySeparateAmbiguousObjects()
        {
            var f = new Fixture(); f.Add(0, "mnp_registration"); f.Add(1, "mnp_registration"); f.Resolve();
            Assert.IsTrue(f.Snapshot.Components.All(x => x.Status == IdentityResolutionStatus.Ambiguous && x.ComparisonKey == null));
            var p = Project(f); var items = new SolutionObjectNavigation(p).Items(ObjectsWorkspaceScope.Source);
            Assert.AreEqual(2, items.Count); Assert.AreEqual(2, items.Select(x => x.Key).Distinct().Count());
            Assert.IsTrue(items.All(x => x.Status(p) == "Needs Review"));
        }

        [TestMethod]
        public void RelatedTableNameAloneDoesNotSupplyArchivalIdentity()
        {
            var f = new Fixture(); f.Add(0, "mnp_registration"); f.Rows[0]["name"] = "OtherOwnName"; f.Resolve();
            Assert.AreEqual(IdentityResolutionStatus.Unresolved, f.Snapshot.Components[0].Status);
            Assert.IsNull(f.Snapshot.Components[0].ComparisonKey);
        }

        [TestMethod]
        public void BackingSolutionIdDoesNotOverrideMembershipSolution()
        {
            var f = new Fixture(); f.Add(0, "mnp_registration"); f.Rows[0]["solutionid"] = Id(9000); f.Resolve();
            Assert.AreEqual(Guid.Parse("6d64348f-9694-43c1-830f-a8c35703cd14"), f.Snapshot.Solution.SolutionId);
            Assert.AreEqual(IdentityResolutionStatus.Resolved, f.Snapshot.Components[0].Status);
            Assert.IsTrue(f.DataQueries.All(q => q.Criteria.Conditions.Single().AttributeName == "metadataforarchivalid"));
        }

        [TestMethod]
        public void MatchedMembershipNeverCreatesArchivalConfigurationMatch()
        {
            var a = new Fixture(); var b = new Fixture(true); a.Add(0, "mnp_registration"); b.Add(0, "mnp_registration"); a.Resolve(); b.Resolve();
            var reader = new DataverseComponentDefinitionReader(); var left = reader.Read(a.Context, a.Snapshot, CancellationToken.None); var right = reader.Read(b.Context, b.Snapshot, CancellationToken.None);
            Assert.AreEqual(ComponentDefinitionReadStatus.Unsupported, left.Definitions.Single().Status);
            var membership = new SolutionMembershipComparer().Compare(a.Snapshot, b.Snapshot);
            Assert.AreEqual(MembershipPresence.PresentInBoth, membership.Single().Presence);
            var detail = new ComponentDetailComparer().Compare(membership, left, right).Single();
            Assert.AreEqual(ComponentDetailComparisonStatus.Unsupported, detail.Status);
        }

        [TestMethod]
        public void ProjectionRejectsAnotherSnapshotAndForeignSideInformation()
        {
            var f = new Fixture(); f.Add(0, "mnp_registration"); f.Resolve(); var copy = MembershipSnapshot.Complete(f.Snapshot.Solution, f.Snapshot.Components, f.Snapshot.CapturedAt);
            MetadataForArchivalInformation info;
            Assert.IsFalse(f.Context.MetadataCache.MetadataForArchival.TryInformation(copy, copy.Components[0], DisplayEnvironmentSide.Source, out info));
            Assert.IsTrue(f.Context.MetadataCache.MetadataForArchival.TryInformation(f.Snapshot, f.Snapshot.Components[0], DisplayEnvironmentSide.Target, out info));
            Assert.ThrowsException<ArgumentException>(() => new ComponentSideDisplayEvidence(DisplayEnvironmentSide.Source,
                info.SolutionComponentId, info.ObjectId, 10139, new ComponentDisplayValue(DisplayAvailability.NotRetrieved),
                info.Field("Name"), info.Field("Name"), DisplayBackingCorrelationStatus.Unique, info.Field("Related Table"), archivalInformation: info));
        }

        [TestMethod]
        public void MetadataIsBatchedAndDoesNotRunForExistingKinds()
        {
            var f = new Fixture(); for (int i = 0; i < 205; i++) f.Add(i, "mnp_table" + i); f.Resolve();
            Assert.AreEqual(2, f.DataQueries.Count); Assert.AreEqual(2, f.TableRequests);
            Assert.IsTrue(f.DataQueries.All(q => q.Criteria.Conditions.Single().Values.Count <= 200));
            Assert.AreEqual(205, f.Snapshot.Components.Count); Assert.IsTrue(f.Snapshot.Components.All(x => x.Status == IdentityResolutionStatus.Resolved));
            var empty = new Fixture(); empty.Resolve(); Assert.AreEqual(0, empty.SchemaRequests); Assert.AreEqual(0, empty.DataQueries.Count);
        }

        [TestMethod]
        public void FilteringCategoryAndDetailsKeepOriginalRowsAndExportSchema()
        {
            var f = new Fixture(); f.Add(0, "mnp_registration"); f.AddTables(); f.Resolve(); var p = Project(f); var model = new SolutionObjectNavigation(p);
            var csv = new MembershipCoverageCsvExporter().CreateCsv(p.Owner); var originals = p.Rows.Select(x => x.Original.AuthoritativeRow).ToArray();
            var archival = model.Items(ObjectsWorkspaceScope.Source).Single(x => x.Category == "MetadataForArchival");
            var chosen = model.Items(ObjectsWorkspaceScope.Source).Where(x => x.Rows.Contains(archival.Rows.Single())).Single();
            Assert.AreEqual(archival.Key, chosen.Key); Assert.AreSame(originals[chosen.Rows.Single()], p.Rows[chosen.Rows.Single()].Original.AuthoritativeRow);
            Assert.AreEqual(2, model.Items(ObjectsWorkspaceScope.Source).Count); Assert.AreEqual(0, model.Items(ObjectsWorkspaceScope.Target).Count);
            Assert.AreEqual(csv, new MembershipCoverageCsvExporter().CreateCsv(p.Owner));
            Assert.AreEqual(2, p.Rows.Count); Assert.AreEqual(2, p.Occurrences.Count);
            StringAssert.Contains(p.Rows[archival.Rows.Single()].ComparedPropertiesExplanation, "Archival readiness=False");
        }

        private static Guid Id(int value) => new Guid(value, 0, 0, new byte[] { 0, 0, 0, 0, 0, 0, 0, 1 });
        private static MembershipSnapshot Empty() => MembershipSnapshot.Complete(new SolutionIdentity(new EnvironmentIdentity(Id(9998), "Empty offline side"), Id(9999), "SynthetiCsyntaf254"), new ComponentIdentity[0], DateTimeOffset.MinValue);
        private static SolutionObjectsExplorerProjection Project(Fixture f, bool target = false)
        {
            var present = MembershipEnvironmentResult.FromSnapshot("Offline", f.Snapshot, 0, TimeSpan.Zero); var empty = MembershipEnvironmentResult.FromSnapshot("Empty", Empty(), 0, TimeSpan.Zero);
            var p = new MembershipResultPresenter().Create(target ? empty : present, target ? present : empty);
            var completed = new CompletedMembershipComparison(p, ComponentDisplayEvidenceProjector.Project(p, target ? null : f.Context.MetadataCache, target ? f.Context.MetadataCache : null), DisplayProjectionOutcome.Completed);
            return ExplorerFixture.Project(new ExplorerFixture { Presentation = p, Completed = completed });
        }
        private sealed class Fixture
        {
            private readonly bool target;
            internal string Failure;
            internal bool OmitOptionalSchema;
            internal readonly List<Entity> Rows = new List<Entity>();
            internal readonly Dictionary<Guid, EntityMetadata> TableRows = new Dictionary<Guid, EntityMetadata>();
            internal readonly List<ComponentIdentity> Raw = new List<ComponentIdentity>();
            internal readonly List<QueryExpression> DataQueries = new List<QueryExpression>();
            internal FakeOrganizationService Service;
            internal DataverseReadContext Context;
            internal MembershipSnapshot Snapshot;
            internal int TableRequests, SchemaRequests;
            internal Fixture(bool isTarget = false) { target = isTarget; }
            internal void Add(int i, string name)
            {
                Guid own = !target && i < 2 ? Guid.Parse(i == 0 ? "e7c04ff2-4d43-4234-83b9-9af755a5bf24" : "85f5fa0f-8997-4574-b524-63112271bba7") : Id(10000 + (target ? 1000 : 0) + i);
                Guid table = !target && i < 2 ? Guid.Parse(i == 0 ? "75b85b19-2df7-429a-bfd2-20ddef9318c9" : "49891303-fe48-4896-9f23-d344775b3e18") : Id(20000 + (target ? 1000 : 0) + i);
                Raw.Add(new ComponentIdentity(new SolutionComponentRecord(Id(30000 + (target ? 1000 : 0) + i), 10139, own), IdentityResolutionStatus.Unresolved));
                Rows.Add(new Entity("metadataforarchival", own) { ["metadataforarchivalid"] = own, ["name"] = name, ["extensionofrecordid"] = table,
                    ["ismanaged"] = true, ["isavailableforarchival"] = true, ["isreadyforarchival"] = false });
                TableRows[table] = ParentMetadataTestData.Root(table, name);
            }
            internal void AddTables()
            {
                foreach (var table in TableRows.Values) Raw.Add(new ComponentIdentity(new SolutionComponentRecord(Id(40000 + Raw.Count), 1, table.MetadataId), IdentityResolutionStatus.Unresolved));
            }
            internal void Resolve()
            {
                var environment = new EnvironmentIdentity(Id(target ? 7001 : 7000), target ? "Offline Target" : "SYN-THETIC-SAG fixture");
                Snapshot = MembershipSnapshot.Complete(new SolutionIdentity(environment, target ? Id(8001) : Guid.Parse("6d64348f-9694-43c1-830f-a8c35703cd14"), "SynthetiCsyntaf254"), Raw, DateTimeOffset.MinValue);
                Service = new FakeOrganizationService { ExecuteRequest = request => {
                    if (request is WhoAmIRequest) { var who = new WhoAmIResponse(); who.Results["OrganizationId"] = environment.OrganizationId; return who; }
                    if (request is RetrieveEntityRequest) { SchemaRequests++; if (Failure == "schema-failure") throw new FaultException("Synthetic failure"); var response = new RetrieveEntityResponse(); response.Results["EntityMetadata"] = Schema(); return response; }
                    var change = request as RetrieveMetadataChangesRequest; Assert.IsNotNull(change); TableRequests++;
                    if (Failure == "table-failure") throw new FaultException("Synthetic failure");
                    var ids = change.Query.Criteria.Conditions.Select(x => (Guid)x.Value).ToArray();
                    return ParentMetadataTestData.Response(ids.Where(TableRows.ContainsKey).Select(x => TableRows[x]).ToArray());
                }, RetrievePage = query => {
                    Assert.AreEqual("metadataforarchival", query.EntityName); DataQueries.Add(query);
                    if (Failure == "query-failure") throw new FaultException("Synthetic failure");
                    var ids = query.Criteria.Conditions.Single().Values.Cast<Guid>().ToArray();
                    var rows = Rows.Where(x => ids.Contains(x.Id)).Select(x => { var copy = new Entity(x.LogicalName, x.Id); foreach (var column in query.ColumnSet.Columns) if (x.Contains(column)) copy[column] = x[column]; return copy; }).ToList();
                    if (Failure == "missing-row") rows.Clear();
                    if (Failure == "duplicate-row") rows.AddRange(rows.ToArray());
                    if (rows.Count > 0 && Failure == "wrong-primary") rows[0]["metadataforarchivalid"] = Id(777);
                    if (rows.Count > 0 && Failure == "missing-name") rows[0].Attributes.Remove("name");
                    if (rows.Count > 0 && Failure == "wrong-relation") rows[0]["extensionofrecordid"] = Id(778);
                    if (rows.Count > 0 && Failure == "missing-relation") rows[0].Attributes.Remove("extensionofrecordid");
                    return new EntityCollection(rows) { MoreRecords = Failure == "incomplete" };
                }};
                Context = new DataverseReadContext(Service, environment, CancellationToken.None);
                Snapshot = new DataverseComponentIdentityResolver().ResolveSnapshot(Context, Snapshot, CancellationToken.None);
            }
            private EntityMetadata Schema()
            {
                var fields = new List<AttributeMetadata> { new UniqueIdentifierAttributeMetadata { LogicalName = "metadataforarchivalid" }, new StringAttributeMetadata { LogicalName = "name" }, new UniqueIdentifierAttributeMetadata { LogicalName = "extensionofrecordid" },
                    new BooleanAttributeMetadata { LogicalName = "isavailableforarchival" }, new BooleanAttributeMetadata { LogicalName = "isreadyforarchival" } };
                if (!OmitOptionalSchema) fields.Add(new BooleanAttributeMetadata { LogicalName = "ismanaged" });
                fields.Add(new UniqueIdentifierAttributeMetadata { LogicalName = "solutionid" });
                foreach (var field in fields) { field.GetType().GetProperty("IsValidForRead").SetValue(field, true); field.GetType().GetProperty("IsSecured").SetValue(field, false); }
                var schema = new EntityMetadata { LogicalName = Failure == "bad-schema" ? "wrong_table" : "metadataforarchival" };
                typeof(EntityMetadata).GetProperty("PrimaryIdAttribute").SetValue(schema, "metadataforarchivalid"); typeof(EntityMetadata).GetProperty("PrimaryNameAttribute").SetValue(schema, "name"); typeof(EntityMetadata).GetProperty("Attributes").SetValue(schema, fields.ToArray()); return schema;
            }
        }
    }
}
