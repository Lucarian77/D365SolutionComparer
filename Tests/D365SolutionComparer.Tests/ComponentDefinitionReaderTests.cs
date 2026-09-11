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
    public class ComponentDefinitionReaderTests
    {
        [DataTestMethod]
        [DataRow("webresource", 61, "webresource", "webresourceid", "name")]
        [DataRow("environmentvariabledefinition", 380, "environmentvariabledefinition", "environmentvariabledefinitionid", "schemaname")]
        [DataRow("connectionreference", 10003, "connectionreference", "connectionreferenceid", "connectionreferencelogicalname")]
        [DataRow("appmodule", 80, "appmodule", "appmoduleid", "uniquename")]
        [DataRow("sitemap", 62, "sitemap", "sitemapid", "sitemapnameunique")]
        public void EntityBackedSupportedKindsUseOneReadOnlyBatchedQuery(string kind, int type,
            string entityName, string primaryId, string identityAttribute)
        {
            var fixture = Fixture(Identity(kind, type, "portable"));
            QueryExpression captured = null;
            fixture.Service.RetrievePage = query =>
            {
                captured = query;
                var row = new Entity(entityName, fixture.Identity.Record.ObjectId.Value)
                {
                    [primaryId] = fixture.Identity.Record.ObjectId.Value,
                    [identityAttribute] = "portable"
                };
                return Rows(row);
            };
            var counter = new DataverseRequestCounter();

            var result = new DataverseComponentDefinitionReader().Read(fixture.Service,
                fixture.Snapshot, CancellationToken.None, counter).Definitions.Single();

            Assert.AreEqual(ComponentDefinitionReadStatus.Available, result.Status);
            Assert.AreEqual(entityName, captured.EntityName);
            Assert.AreEqual(ConditionOperator.In, captured.Criteria.Conditions.Single().Operator);
            Assert.IsTrue(captured.ColumnSet.Columns.Contains(primaryId));
            Assert.IsTrue(captured.ColumnSet.Columns.Contains(identityAttribute));
            Assert.IsTrue(ComponentDefinitionContractCatalog.For(kind).ComparableProperties
                .All(item => captured.ColumnSet.Columns.Contains(item)));
            Assert.AreEqual(1, counter.GetQueryCount(entityName));
            Assert.AreEqual(1, counter.GetExecuteCount("WhoAmI"));
            Assert.AreEqual(2, counter.TotalRequests);
        }

        [TestMethod]
        public void RepeatedObjectIdsAreDeduplicatedWithinEntityBatch()
        {
            var objectId = Guid.NewGuid();
            var first = Identity("webresource", 61, "new_/script.js", objectId);
            var second = Identity("webresource", 61, "new_/script.js", objectId);
            var fixture = Fixture(first, second);
            fixture.Service.RetrievePage = query => Rows(new Entity("webresource", objectId)
            {
                ["webresourceid"] = objectId,
                ["name"] = "new_/script.js",
                ["content"] = "YWJj"
            });
            var counter = new DataverseRequestCounter();

            var definitions = new DataverseComponentDefinitionReader().Read(fixture.Service,
                fixture.Snapshot, CancellationToken.None, counter).Definitions;

            Assert.AreEqual(2, definitions.Count);
            Assert.IsTrue(definitions.All(item => item.Status == ComponentDefinitionReadStatus.Available));
            Assert.AreEqual(1, counter.GetQueryCount("webresource"));
        }

        [TestMethod]
        public void EntityBackedReadsUseDeterministicBatchesOfAtMostTwoHundredIds()
        {
            var identities = Enumerable.Range(0, 201)
                .Select(index => Identity("webresource", 61, "new_/script" + index + ".js"))
                .ToArray();
            var fixture = Fixture(identities);
            var batchSizes = new List<int>();
            fixture.Service.RetrievePage = query =>
            {
                var ids = query.Criteria.Conditions.Single().Values.Cast<Guid>().ToList();
                batchSizes.Add(ids.Count);
                return Rows(ids.Select(id => new Entity("webresource", id)
                {
                    ["webresourceid"] = id,
                    ["name"] = "portable"
                }).ToArray());
            };
            var counter = new DataverseRequestCounter();

            var definitions = new DataverseComponentDefinitionReader().Read(fixture.Service,
                fixture.Snapshot, CancellationToken.None, counter).Definitions;

            Assert.AreEqual(201, definitions.Count);
            CollectionAssert.AreEqual(new[] { 200, 1 }, batchSizes);
            Assert.AreEqual(2, counter.GetQueryCount("webresource"));
        }

        [TestMethod]
        public void TableDefinitionsUseGroupedMetadataIdEqualsConditionsWithGuidValues()
        {
            var identity = Identity("table", 1, "account");
            var fixture = Fixture(identity);
            RetrieveMetadataChangesRequest captured = null;
            fixture.Service.ExecuteRequest = request =>
            {
                if (request is WhoAmIRequest) return MembershipTestData.WhoAmI(fixture.Solution.Environment.OrganizationId);
                captured = (RetrieveMetadataChangesRequest)request;
                return MetadataRows(new EntityMetadata
                {
                    MetadataId = identity.Record.ObjectId,
                    LogicalName = "account",
                    SchemaName = "Account",
                    OwnershipType = OwnershipTypes.UserOwned,
                    IsActivity = false
                });
            };
            var counter = new DataverseRequestCounter();

            var definition = new DataverseComponentDefinitionReader().Read(fixture.Service,
                fixture.Snapshot, CancellationToken.None, counter).Definitions.Single();

            Assert.AreEqual(ComponentDefinitionReadStatus.Available, definition.Status);
            Assert.AreEqual(LogicalOperator.Or, captured.Query.Criteria.FilterOperator);
            Assert.AreEqual(MetadataConditionOperator.Equals,
                captured.Query.Criteria.Conditions.Single().ConditionOperator);
            Assert.IsInstanceOfType(captured.Query.Criteria.Conditions.Single().Value, typeof(Guid));
            Assert.AreEqual("Account", definition.ComparableProperties["SchemaName"]);
            Assert.AreEqual(1, counter.GetExecuteCount("RetrieveMetadataChanges"));
        }

        [TestMethod, TestCategory("Phase2G3")]
        public void ColumnAndRelationshipDefinitionsUseReadOnlyMetadataRequests()
        {
            var column = Identity("column", 2, "account.new_code");
            var relationship = Identity("relationship", 10, "new_account_contact");
            var fixture = Fixture(column, relationship);
            fixture.Service.ExecuteRequest = request =>
            {
                if (request is WhoAmIRequest) return MembershipTestData.WhoAmI(fixture.Solution.Environment.OrganizationId);
                var metadata = (RetrieveMetadataChangesRequest)request;
                ParentMetadataTestData.AssertQuery(metadata);
                Assert.AreEqual("LogicalName", metadata.Query.Criteria.Conditions.Single().PropertyName);
                Assert.AreEqual("account", metadata.Query.Criteria.Conditions.Single().Value);
                return ParentMetadataTestData.Response(ParentMetadataTestData.Root(Guid.NewGuid(), "account",
                    new[] { ParentMetadataTestData.Column(column.Record.ObjectId.Value) },
                    new[] { ParentMetadataTestData.Relationship(relationship.Record.ObjectId.Value) }));
            };
            var counter = new DataverseRequestCounter();
            var definitions = new DataverseComponentDefinitionReader().Read(fixture.Service,
                fixture.Snapshot, CancellationToken.None, counter).Definitions;
            Assert.IsTrue(definitions.All(item => item.Status == ComponentDefinitionReadStatus.Available));
            Assert.AreEqual("100", definitions[0].ComparableProperties["MaxLength"]);
            Assert.AreEqual("contact", definitions[1].ComparableProperties["ReferencingEntity"]);
            Assert.AreEqual(1, counter.GetExecuteCount("RetrieveMetadataChanges"));
            Assert.AreEqual(0, counter.GetExecuteCount("RetrieveAttribute"));
            Assert.AreEqual(0, counter.GetExecuteCount("RetrieveRelationship"));
        }

        [TestMethod]
        public void RepeatedMetadataObjectIdsDoNotRepeatDirectMetadataRequests()
        {
            var id = Guid.NewGuid();
            var fixture = Fixture(Identity("column", 2, "account.new_code", id),
                Identity("column", 2, "account.new_code", id));
            fixture.Service.ExecuteRequest = request => request is WhoAmIRequest
                ? (OrganizationResponse)MembershipTestData.WhoAmI(fixture.Solution.Environment.OrganizationId)
                : ParentMetadataTestData.Response(ParentMetadataTestData.Root(Guid.NewGuid(), "account",
                    new[] { ParentMetadataTestData.Column(id) }));
            var counter = new DataverseRequestCounter();
            var definitions = new DataverseComponentDefinitionReader().Read(fixture.Service,
                fixture.Snapshot, CancellationToken.None, counter).Definitions;
            Assert.AreEqual(2, definitions.Count);
            Assert.IsTrue(definitions.All(item => item.Status == ComponentDefinitionReadStatus.Available));
            Assert.AreEqual(1, counter.GetExecuteCount("RetrieveMetadataChanges"));
            Assert.AreEqual(0, counter.GetExecuteCount("RetrieveAttribute"));
        }

        [TestMethod]
        public void EduSizedColumnAndRelationshipPopulationUsesOneParentBatchNotIndividualRequests()
        {
            var columns = Enumerable.Range(0, 438).Select(i => Identity("column", 2, "account.new_column" + i)).ToArray();
            var relationships = Enumerable.Range(0, 59).Select(i => Identity("relationship", 10, "new_relationship" + i)).ToArray();
            var fixture = Fixture(columns.Concat(relationships).ToArray());
            fixture.Service.ExecuteRequest = request =>
            {
                if (request is WhoAmIRequest) return MembershipTestData.WhoAmI(fixture.Solution.Environment.OrganizationId);
                ParentMetadataTestData.AssertQuery((RetrieveMetadataChangesRequest)request);
                return ParentMetadataTestData.Response(ParentMetadataTestData.Root(Guid.NewGuid(), "account",
                    columns.Select((item, i) => ParentMetadataTestData.Column(item.Record.ObjectId.Value, "new_column" + i)).ToArray(),
                    relationships.Select((item, i) => ParentMetadataTestData.Relationship(item.Record.ObjectId.Value, "new_relationship" + i)).ToArray()));
            };
            var counter = new DataverseRequestCounter();
            var definitions = new DataverseComponentDefinitionReader().Read(fixture.Service,
                fixture.Snapshot, CancellationToken.None, counter).Definitions;
            Assert.AreEqual(497, definitions.Count);
            Assert.IsTrue(definitions.All(item => item.Status == ComponentDefinitionReadStatus.Available));
            Assert.AreEqual(1, counter.GetExecuteCount("RetrieveMetadataChanges"));
            Assert.AreEqual(0, counter.GetExecuteCount("RetrieveAttribute"));
            Assert.AreEqual(0, counter.GetExecuteCount("RetrieveRelationship"));
        }

        [TestMethod]
        public void GlobalChoicesReuseOneCatalogReadAndIndexByMetadataId()
        {
            var first = Identity("globalchoice", 9, "new_choice_a");
            var second = Identity("globalchoice", 9, "new_choice_b");
            var fixture = Fixture(first, second);
            fixture.Service.ExecuteRequest = request =>
            {
                if (request is WhoAmIRequest) return MembershipTestData.WhoAmI(fixture.Solution.Environment.OrganizationId);
                var response = new RetrieveAllOptionSetsResponse();
                response.Results["OptionSetMetadata"] = new OptionSetMetadataBase[]
                {
                    new OptionSetMetadata { MetadataId = first.Record.ObjectId, Name = "new_choice_a", IsGlobal = true },
                    new OptionSetMetadata { MetadataId = second.Record.ObjectId, Name = "new_choice_b", IsGlobal = true },
                    new OptionSetMetadata { MetadataId = Guid.NewGuid(), Name = "unrelated", IsGlobal = true }
                };
                return response;
            };
            var counter = new DataverseRequestCounter();

            var definitions = new DataverseComponentDefinitionReader().Read(fixture.Service,
                fixture.Snapshot, CancellationToken.None, counter).Definitions;

            Assert.IsTrue(definitions.All(item => item.Status == ComponentDefinitionReadStatus.Available));
            Assert.AreEqual(1, counter.GetExecuteCount("RetrieveAllOptionSets"));
            Assert.AreEqual(2, counter.TotalRequests);
        }

        [TestMethod]
        public void OperationScopedMetadataCacheAvoidsRepeatedIdentityMetadataAndWhoAmIReads()
        {
            var column = Identity("column", 2, "account.new_code");
            var relationship = Identity("relationship", 10, "new_account_contact");
            var choice = Identity("globalchoice", 9, "new_choice");
            var fixture = Fixture(column, relationship, choice);
            var counter = new DataverseRequestCounter();
            var context = new D365SolutionComparer.Services.Membership.DataverseReadContext(
                fixture.Service, fixture.Solution.Environment, CancellationToken.None, counter);
            var inventory = new D365SolutionComparer.Services.Membership.ParentEntityMetadataInventory();
            inventory.Load(new[] { ParentMetadataTestData.Root(Guid.NewGuid(), "account",
                new[] { ParentMetadataTestData.Column(column.Record.ObjectId.Value) },
                new[] { ParentMetadataTestData.Relationship(relationship.Record.ObjectId.Value) }) });
            context.MetadataCache.ParentMetadata = inventory;
            context.MetadataCache.StoreOptionSetCatalog(new OptionSetMetadataBase[]
            {
                new OptionSetMetadata
                {
                    MetadataId = choice.Record.ObjectId,
                    Name = "new_choice",
                    IsGlobal = true
                }
            });

            var definitions = new DataverseComponentDefinitionReader().Read(context,
                fixture.Snapshot, CancellationToken.None).Definitions;

            Assert.IsTrue(definitions.All(item => item.Status == ComponentDefinitionReadStatus.Available));
            Assert.AreEqual(1, counter.TotalRequests, "Only the context's initial WhoAmI should execute.");
            Assert.AreEqual(1, counter.GetExecuteCount("WhoAmI"));
            Assert.AreEqual(0, counter.GetExecuteCount("RetrieveMetadataChanges"));
            Assert.AreEqual(0, counter.GetExecuteCount("RetrieveAllOptionSets"));
        }

        [TestMethod]
        public void BulkIdentityResolutionMetadataIsReusedByDefinitionRetrieval()
        {
            var parentId = Guid.NewGuid();
            var rawColumn = new SolutionComponentRecord(Guid.NewGuid(), 2, Guid.NewGuid());
            var rawRelationship = new SolutionComponentRecord(Guid.NewGuid(), 10, Guid.NewGuid());
            var rawChoice = new SolutionComponentRecord(Guid.NewGuid(), 9, Guid.NewGuid());
            var rawWebResource = new SolutionComponentRecord(Guid.NewGuid(), 61, Guid.NewGuid());
            var fixture = Fixture(ParentMetadataTestData.Raw(1, parentId),
                new ComponentIdentity(rawColumn, IdentityResolutionStatus.Unresolved),
                new ComponentIdentity(rawRelationship, IdentityResolutionStatus.Unresolved),
                new ComponentIdentity(rawChoice, IdentityResolutionStatus.Unresolved),
                new ComponentIdentity(rawWebResource, IdentityResolutionStatus.Unresolved));
            fixture.Service.RetrievePage = query => WebResourceRow(rawWebResource.ObjectId.Value,
                "new_/script.js", "YWJj");
            fixture.Service.ExecuteRequest = request =>
            {
                if (request is WhoAmIRequest)
                    return MembershipTestData.WhoAmI(fixture.Solution.Environment.OrganizationId);
                if (request is RetrieveAllOptionSetsRequest)
                {
                    var optionResponse = new RetrieveAllOptionSetsResponse();
                    optionResponse.Results["OptionSetMetadata"] = new OptionSetMetadataBase[]
                    {
                        new OptionSetMetadata
                        {
                            MetadataId = rawChoice.ObjectId,
                            Name = "new_choice",
                            IsGlobal = true
                        }
                    };
                    return optionResponse;
                }
                ParentMetadataTestData.AssertQuery((RetrieveMetadataChangesRequest)request);
                return ParentMetadataTestData.Response(ParentMetadataTestData.Root(parentId, "account",
                    new[] { ParentMetadataTestData.Column(rawColumn.ObjectId.Value) },
                    new[] { ParentMetadataTestData.Relationship(rawRelationship.ObjectId.Value) }));
            };
            var counter = new DataverseRequestCounter();
            var context = new D365SolutionComparer.Services.Membership.DataverseReadContext(
                fixture.Service, fixture.Solution.Environment, CancellationToken.None, counter);
            var resolved = new D365SolutionComparer.Services.Membership.DataverseComponentIdentityResolver()
                .ResolveSnapshot(context, fixture.Snapshot, CancellationToken.None);
            var beforeDefinitionRead = counter.TotalRequests;

            var definitions = new DataverseComponentDefinitionReader().Read(context,
                resolved, CancellationToken.None).Definitions;

            Assert.IsTrue(definitions.All(item => item.Status == ComponentDefinitionReadStatus.Available));
            Assert.AreEqual(beforeDefinitionRead, counter.TotalRequests);
            Assert.AreEqual(1, counter.GetExecuteCount("WhoAmI"));
            Assert.AreEqual(1, counter.GetExecuteCount("RetrieveMetadataChanges"));
            Assert.AreEqual(1, counter.GetExecuteCount("RetrieveAllOptionSets"));
            Assert.AreEqual(0, counter.GetExecuteCount("RetrieveAttribute"));
            Assert.AreEqual(0, counter.GetExecuteCount("RetrieveRelationship"));
            Assert.AreEqual(1, counter.GetQueryCount("webresource"));
        }

        [TestMethod]
        public void WebResourceContentIsRetrievedAndAContentChangeIsDifferent()
        {
            var source = Fixture(Identity("webresource", 61, "new_/script.js"));
            var target = Fixture(Identity("webresource", 61, "NEW_/SCRIPT.JS"));
            source.Service.RetrievePage = query => WebResourceRow(
                source.Identity.Record.ObjectId.Value, "new_/script.js", "YmVmb3Jl");
            target.Service.RetrievePage = query => WebResourceRow(
                target.Identity.Record.ObjectId.Value, "NEW_/SCRIPT.JS", "YWZ0ZXI=");
            var sourceDefinitions = new DataverseComponentDefinitionReader().Read(source.Service,
                source.Snapshot, CancellationToken.None);
            var targetDefinitions = new DataverseComponentDefinitionReader().Read(target.Service,
                target.Snapshot, CancellationToken.None);
            var memberships = new D365SolutionComparer.Services.Membership.SolutionMembershipComparer()
                .Compare(source.Snapshot, target.Snapshot);

            var result = new ComponentDetailComparer().Compare(memberships,
                sourceDefinitions, targetDefinitions).Single();

            Assert.AreEqual(ComponentDetailComparisonStatus.Different, result.Status);
            Assert.AreEqual("content", result.Differences.Single().PropertyName);
        }

        [TestMethod]
        public void IncompleteEntityBatchIsConservative()
        {
            var fixture = Fixture(Identity("webresource", 61, "new_/script.js"));
            fixture.Service.RetrievePage = query => new EntityCollection { MoreRecords = true };
            var definition = new DataverseComponentDefinitionReader().Read(fixture.Service,
                fixture.Snapshot, CancellationToken.None).Definitions.Single();
            Assert.AreEqual(ComponentDefinitionReadStatus.Unresolved, definition.Status);
        }

        [TestMethod]
        public void RetrievalFaultIsConservativeAndDoesNotBecomeAnEmptyMatch()
        {
            var fixture = Fixture(Identity("webresource", 61, "new_/script.js"));
            fixture.Service.RetrievePage = query => throw new FaultException("read failed");
            var definition = new DataverseComponentDefinitionReader().Read(fixture.Service,
                fixture.Snapshot, CancellationToken.None).Definitions.Single();
            Assert.AreEqual(ComponentDefinitionReadStatus.Unresolved, definition.Status);
            StringAssert.Contains(definition.Diagnostic, "read failed");
        }

        [TestMethod]
        public void CancellationPropagatesBeforeAnyDataverseRequest()
        {
            var fixture = Fixture(Identity("webresource", 61, "new_/script.js"));
            var source = new CancellationTokenSource();
            source.Cancel();
            Assert.ThrowsException<OperationCanceledException>(() =>
                new DataverseComponentDefinitionReader().Read(fixture.Service,
                    fixture.Snapshot, source.Token));
            Assert.AreEqual(0, fixture.Service.ExecuteCalls);
            Assert.AreEqual(0, fixture.Service.Calls);
        }

        [TestMethod]
        public void UnsupportedResolvedKindDoesNotIssueDefinitionRequest()
        {
            var fixture = Fixture(Identity("securityrole", 20, "new_role"));
            var counter = new DataverseRequestCounter();
            var definition = new DataverseComponentDefinitionReader().Read(fixture.Service,
                fixture.Snapshot, CancellationToken.None, counter).Definitions.Single();
            Assert.AreEqual(ComponentDefinitionReadStatus.Unsupported, definition.Status);
            Assert.AreEqual(1, counter.TotalRequests);
        }

        [TestMethod]
        public void DefinitionReaderIssuesNoDataverseWriteRequests()
        {
            var fixture = Fixture(Identity("webresource", 61, "new_/script.js"));
            fixture.Service.RetrievePage = query => WebResourceRow(
                fixture.Identity.Record.ObjectId.Value, "new_/script.js", "YWJj");
            var definition = new DataverseComponentDefinitionReader().Read(fixture.Service,
                fixture.Snapshot, CancellationToken.None).Definitions.Single();
            Assert.AreEqual(ComponentDefinitionReadStatus.Available, definition.Status);
            Assert.AreEqual(1, fixture.Service.Calls);
            Assert.AreEqual(1, fixture.Service.ExecuteCalls);
        }

        private static ReaderFixture Fixture(params ComponentIdentity[] identities)
        {
            var solution = new SolutionIdentity(new EnvironmentIdentity(Guid.NewGuid(), "Test"),
                Guid.NewGuid(), "sample");
            var service = MembershipTestData.Service(solution, query => Rows());
            return new ReaderFixture(solution, identities.First(), service,
                MembershipSnapshot.Complete(solution, identities, DateTimeOffset.UtcNow));
        }

        private static ComponentIdentity Identity(string kind, int type, string key,
            Guid? objectId = null) => new ComponentIdentity(
                new SolutionComponentRecord(Guid.NewGuid(), type, objectId ?? Guid.NewGuid()),
                IdentityResolutionStatus.Resolved, key, componentTypeKey: kind, semanticKind: kind);

        private static EntityCollection Rows(params Entity[] rows)
        {
            var collection = new EntityCollection();
            collection.Entities.AddRange(rows);
            return collection;
        }

        private static EntityCollection WebResourceRow(Guid id, string name, string content) =>
            Rows(new Entity("webresource", id)
            {
                ["webresourceid"] = id,
                ["name"] = name,
                ["webresourcetype"] = new OptionSetValue(3),
                ["content"] = content,
                ["displayname"] = "Script",
                ["description"] = "Test resource"
            });

        private static RetrieveMetadataChangesResponse MetadataRows(params EntityMetadata[] rows)
        {
            var response = new RetrieveMetadataChangesResponse();
            var collection = new EntityMetadataCollection();
            collection.AddRange(rows);
            response.Results["EntityMetadata"] = collection;
            return response;
        }

        private static EntityMetadata EntityWithAttributes(string logicalName,
            AttributeMetadata[] attributes)
        {
            var metadata = new EntityMetadata { LogicalName = logicalName };
            typeof(EntityMetadata).GetProperty("Attributes").SetValue(metadata, attributes, null);
            return metadata;
        }

        private static EntityMetadata EntityWithRelationships(string logicalName,
            OneToManyRelationshipMetadata[] relationships)
        {
            var metadata = new EntityMetadata { LogicalName = logicalName };
            typeof(EntityMetadata).GetProperty("OneToManyRelationships")
                .SetValue(metadata, relationships, null);
            return metadata;
        }

        private sealed class ReaderFixture
        {
            public ReaderFixture(SolutionIdentity solution, ComponentIdentity identity,
                FakeOrganizationService service, MembershipSnapshot snapshot)
            {
                Solution = solution;
                Identity = identity;
                Service = service;
                Snapshot = snapshot;
            }

            public SolutionIdentity Solution { get; }
            public ComponentIdentity Identity { get; }
            public FakeOrganizationService Service { get; }
            public MembershipSnapshot Snapshot { get; }
        }
    }
}
