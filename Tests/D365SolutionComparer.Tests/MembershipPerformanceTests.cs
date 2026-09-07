using System;
using System.Collections.Generic;
using System.Linq;
using System.ServiceModel;
using System.Threading;
using D365SolutionComparer.Infrastructure;
using D365SolutionComparer.Models.Membership;
using D365SolutionComparer.Services.Membership;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Metadata.Query;
using Microsoft.Xrm.Sdk.Query;
using static D365SolutionComparer.Tests.MembershipTestData;

namespace D365SolutionComparer.Tests
{
    [TestClass]
    public class MembershipPerformanceTests
    {
        [DataTestMethod]
        [DataRow(100, 4)]
        [DataRow(500, 6)]
        public void BulkOperationBatchesRepresentativeEntitySnapshots(int count, int expectedRequests)
        {
            var solution = Solution();
            var components = Enumerable.Range(0, count).Select(index => ComponentRow(solution)).ToArray();
            var names = components.ToDictionary(item => item.GetAttributeValue<Guid>("objectid"),
                item => "new_/resource_" + item.Id.ToString("N") + ".js");
            var service = Service(solution, query =>
            {
                if (query.EntityName == "solution") return Rows(SolutionRow(solution));
                if (query.EntityName == "solutioncomponent") return Rows(components);
                Assert.AreEqual("webresource", query.EntityName);
                var ids = QueryIds(query).ToArray();
                Assert.IsTrue(ids.Length <= 200);
                return Rows(ids.Select(id => new Entity("webresource", id) { ["name"] = names[id] }).ToArray());
            });
            var counter = new DataverseRequestCounter();
            var result = new DataverseSolutionMembershipOperation().ReadAndResolve(service, solution,
                CancellationToken.None, requestCounter: counter);
            Assert.AreEqual(count, result.Components.Count);
            Assert.IsTrue(result.Components.All(item => item.Status == IdentityResolutionStatus.Resolved));
            Assert.AreEqual(expectedRequests, counter.TotalRequests);
            Assert.AreEqual(1, counter.GetExecuteCount("WhoAmI"));
            Assert.AreEqual(1, counter.GetQueryCount("solution"));
            Assert.AreEqual(1, counter.GetQueryCount("solutioncomponent"));
            Assert.AreEqual(count == 100 ? 1 : 3, counter.GetQueryCount("webresource"));
        }

        [DataTestMethod]
        [DataRow(100, 36)]
        [DataRow(500, 136)]
        public void MixedSupportedSnapshotsHaveBoundedRequestCounts(int count, int expectedRequests)
        {
            var solution = Solution(); const int connectionType = 10027;
            var types = new[] { 1, 2, 10, 61, 29, 20, 380, connectionType };
            var components = Enumerable.Range(0, count).Select(index =>
            {
                var component = ComponentRow(solution, types[index % types.Length]);
                return component;
            }).ToArray();
            var service = Service(solution, query =>
            {
                if (query.EntityName == "solution") return Rows(SolutionRow(solution));
                if (query.EntityName == "solutioncomponent") return Rows(components);
                if (query.EntityName == "solutioncomponentdefinition")
                    return Rows(new Entity("solutioncomponentdefinition", Guid.NewGuid()) { ["objecttypecode"] = connectionType });
                return IdentityRows(query);
            });
            service.ExecuteRequest = request => MetadataResponse(solution, request);
            var counter = new DataverseRequestCounter();
            var result = new DataverseSolutionMembershipOperation().ReadAndResolve(service, solution,
                CancellationToken.None, requestCounter: counter);
            Assert.AreEqual(count, result.Components.Count);
            Assert.IsTrue(result.Components.All(item => item.Status == IdentityResolutionStatus.Resolved));
            Assert.AreEqual(expectedRequests, counter.TotalRequests);
            Assert.AreEqual(1, counter.GetExecuteCount("WhoAmI"));
            Assert.AreEqual(1, counter.GetExecuteCount("RetrieveMetadataChanges"));
            Assert.AreEqual((count + 6) / 8, counter.GetExecuteCount("RetrieveAttribute"));
            Assert.AreEqual((count + 5) / 8, counter.GetExecuteCount("RetrieveRelationship"));
            Assert.AreEqual(1, counter.GetQueryCount("solutioncomponentdefinition"));
        }

        [TestMethod]
        public void RepeatedObjectIdsUseOneLookupPerResolverFamilyAndKeepEveryRawRecord()
        {
            var solution = Solution(); const int connectionType = 10027;
            var pairs = new[]
            {
                Pair(1), Pair(2), Pair(10), Pair(61), Pair(29), Pair(20), Pair(380), Pair(connectionType)
            };
            var identities = pairs.SelectMany(pair => pair).Select(record =>
                new ComponentIdentity(record, IdentityResolutionStatus.Unresolved)).ToArray();
            var service = Service(solution, query =>
            {
                if (query.EntityName == "solutioncomponentdefinition")
                    return Rows(new Entity("solutioncomponentdefinition", Guid.NewGuid()) { ["objecttypecode"] = connectionType });
                return IdentityRows(query);
            });
            service.ExecuteRequest = request => MetadataResponse(solution, request);
            var counter = new DataverseRequestCounter();
            var result = new DataverseComponentIdentityResolver().ResolveSnapshot(
                service, MembershipSnapshot.Complete(solution, identities, DateTimeOffset.UtcNow),
                CancellationToken.None, counter);
            Assert.AreEqual(16, result.Components.Count);
            Assert.IsTrue(result.Components.All(item => item.Status == IdentityResolutionStatus.Resolved));
            Assert.AreEqual(4, counter.ExecuteRequests); // WhoAmI plus one request for each metadata family.
            Assert.AreEqual(6, counter.QueryRequests);   // Five entity families plus connection-type discovery.
            foreach (var pair in result.Components.Select((item, index) => new { item, index }).GroupBy(x => x.index / 2))
            {
                Assert.AreEqual(pair.First().item.ComparisonKey, pair.Last().item.ComparisonKey);
                Assert.AreNotSame(pair.First().item.Record, pair.Last().item.Record);
            }
        }

        [TestMethod]
        public void GroupedTableMetadataUsesOneRequestForDuplicateAndDistinctIds()
        {
            var solution = Solution(); var firstId = Guid.NewGuid(); var secondId = Guid.NewGuid();
            var records = new[] { Record(1, firstId), Record(1, firstId), Record(1, secondId) };
            var service = Service(solution);
            service.ExecuteRequest = request =>
            {
                if (request is WhoAmIRequest) return WhoAmI(solution.Environment.OrganizationId);
                var metadataRequest = (RetrieveMetadataChangesRequest)request;
                var condition = metadataRequest.Query.Criteria.Conditions.Single();
                Assert.AreEqual("MetadataId", condition.PropertyName);
                Assert.AreEqual(MetadataConditionOperator.In, condition.ConditionOperator);
                CollectionAssert.AreEquivalent(new object[] { firstId, secondId }, (object[])condition.Value);
                var response = new RetrieveMetadataChangesResponse();
                response.Results["EntityMetadata"] = new EntityMetadataCollection
                {
                    new EntityMetadata { MetadataId = firstId, LogicalName = "account" },
                    new EntityMetadata { MetadataId = secondId, LogicalName = "contact" }
                };
                return response;
            };
            var snapshot = MembershipSnapshot.Complete(solution, records.Select(record =>
                new ComponentIdentity(record, IdentityResolutionStatus.Unresolved)), DateTimeOffset.UtcNow);
            var result = new DataverseComponentIdentityResolver().ResolveSnapshot(service, snapshot, CancellationToken.None);
            Assert.AreEqual(2, service.ExecuteCalls);
            CollectionAssert.AreEqual(new[] { "account", "account", "contact" },
                result.Components.Select(item => item.ComparisonKey).ToArray());
        }

        [TestMethod]
        public void BatchFaultMarksEveryAffectedIdentityUnresolvedWithoutFalseMissing()
        {
            var solution = Solution();
            var records = Enumerable.Range(0, 10).Select(index => Record(61, Guid.NewGuid())).ToArray();
            var service = Service(solution, query => throw new FaultException("Denied"));
            var snapshot = MembershipSnapshot.Complete(solution, records.Select(record =>
                new ComponentIdentity(record, IdentityResolutionStatus.Unresolved)), DateTimeOffset.UtcNow);
            var result = new DataverseComponentIdentityResolver().ResolveSnapshot(service, snapshot, CancellationToken.None);
            Assert.AreEqual(1, service.Calls);
            Assert.IsTrue(result.Components.All(item => item.Status == IdentityResolutionStatus.Unresolved));
            var comparison = new SolutionMembershipComparer().Compare(result, Snapshot()).ToArray();
            Assert.IsTrue(comparison.All(item => item.Presence == MembershipPresence.Indeterminate));
        }

        [TestMethod]
        public void CancellationAfterGroupedQueryNeverReturnsCompletedResolution()
        {
            var solution = Solution(); var cancellation = new CancellationTokenSource();
            var records = Enumerable.Range(0, 10).Select(index => Record(61, Guid.NewGuid())).ToArray();
            var service = Service(solution, query =>
            {
                cancellation.Cancel();
                return Rows(QueryIds(query).Select(id => new Entity("webresource", id) { ["name"] = id.ToString("N") }).ToArray());
            });
            var snapshot = MembershipSnapshot.Complete(solution, records.Select(record =>
                new ComponentIdentity(record, IdentityResolutionStatus.Unresolved)), DateTimeOffset.UtcNow);
            Assert.ThrowsException<OperationCanceledException>(() => new DataverseComponentIdentityResolver()
                .ResolveSnapshot(service, snapshot, cancellation.Token));
        }

        [TestMethod]
        public void AbsentBulkReadUsesOneWhoAmIAndNoResolverRequests()
        {
            var solution = Solution(); var service = Service(solution, query => Rows());
            var counter = new DataverseRequestCounter();
            var result = new DataverseSolutionMembershipOperation().ReadAndResolve(service, solution.Environment,
                solution.UniqueName, CancellationToken.None, requestCounter: counter);
            Assert.AreEqual(MembershipSnapshotState.SolutionAbsent, result.State);
            Assert.AreEqual(2, counter.TotalRequests);
            Assert.AreEqual(1, counter.GetExecuteCount("WhoAmI"));
            Assert.AreEqual(1, counter.GetQueryCount("solution"));
        }

        private static SolutionComponentRecord[] Pair(int type)
        {
            var objectId = Guid.NewGuid();
            return new[] { Record(type, objectId), Record(type, objectId) };
        }

        private static SolutionComponentRecord Record(int type, Guid objectId) =>
            new SolutionComponentRecord(Guid.NewGuid(), type, objectId);

        private static IEnumerable<Guid> QueryIds(QueryExpression query) =>
            query.Criteria.Conditions.Single().Values.Cast<Guid>();

        private static EntityCollection IdentityRows(QueryExpression query)
        {
            string attribute;
            switch (query.EntityName)
            {
                case "webresource": attribute = "name"; break;
                case "workflow": attribute = "uniquename"; break;
                case "role": attribute = "roletemplateid"; break;
                case "environmentvariabledefinition": attribute = "schemaname"; break;
                case "connectionreference": attribute = "connectionreferencelogicalname"; break;
                default: throw new AssertFailedException("Unexpected query: " + query.EntityName);
            }
            return Rows(QueryIds(query).Select(id =>
            {
                var entity = new Entity(query.EntityName, id);
                entity[attribute] = query.EntityName == "role"
                    ? (object)new EntityReference("roletemplate", Guid.NewGuid())
                    : query.EntityName + "_key";
                return entity;
            }).ToArray());
        }

        private static OrganizationResponse MetadataResponse(Models.Identity.SolutionIdentity solution,
            OrganizationRequest request)
        {
            if (request is WhoAmIRequest) return WhoAmI(solution.Environment.OrganizationId);
            if (request is RetrieveMetadataChangesRequest)
            {
                var metadataRequest = (RetrieveMetadataChangesRequest)request;
                var ids = (object[])metadataRequest.Query.Criteria.Conditions.Single().Value;
                var response = new RetrieveMetadataChangesResponse();
                var metadata = new EntityMetadataCollection();
                metadata.AddRange(ids.Cast<Guid>().Select(id => new EntityMetadata
                {
                    MetadataId = id,
                    LogicalName = "table_" + id.ToString("N")
                }));
                response.Results["EntityMetadata"] = metadata;
                return response;
            }
            if (request is RetrieveAttributeRequest)
            {
                var response = new RetrieveAttributeResponse();
                var metadata = new StringAttributeMetadata
                {
                    MetadataId = ((RetrieveAttributeRequest)request).MetadataId,
                    LogicalName = "column"
                };
                typeof(AttributeMetadata).GetProperty("EntityLogicalName").SetValue(metadata, "table");
                response.Results["AttributeMetadata"] = metadata;
                return response;
            }
            var relationship = (RetrieveRelationshipRequest)request;
            var relationshipResponse = new RetrieveRelationshipResponse();
            relationshipResponse.Results["RelationshipMetadata"] = new OneToManyRelationshipMetadata
            {
                MetadataId = relationship.MetadataId,
                SchemaName = "relationship_" + relationship.MetadataId.ToString("N")
            };
            return relationshipResponse;
        }

    }
}
