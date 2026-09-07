using System;
using System.Linq;
using System.ServiceModel;
using System.Threading;
using D365SolutionComparer.Models.Membership;
using D365SolutionComparer.Services.Membership;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using static D365SolutionComparer.Tests.MembershipTestData;

namespace D365SolutionComparer.Tests
{
    [TestClass]
    public class ComponentIdentityResolverTests
    {
        [DataTestMethod]
        [DataRow(1, "table", "account")]
        [DataRow(2, "column", "account.name")]
        [DataRow(10, "relationship", "new_account_contact")]
        public void MetadataResolversUseObjectIdsAndPublishedStrongNames(int type, string kind, string expectedKey)
        {
            var solution = Solution(); var record = Identity("unused", type).Record;
            var service = Service(solution);
            service.ExecuteRequest = request =>
            {
                if (request is WhoAmIRequest) return WhoAmI(solution.Environment.OrganizationId);
                Assert.AreEqual(record.ObjectId, request.Parameters["MetadataId"]);
                Assert.AreEqual(false, request.Parameters["RetrieveAsIfPublished"]);
                if (type == 1)
                {
                    Assert.IsInstanceOfType(request, typeof(RetrieveEntityRequest));
                    Assert.AreEqual(EntityFilters.Entity, ((RetrieveEntityRequest)request).EntityFilters);
                    var response = new RetrieveEntityResponse();
                    response.Results["EntityMetadata"] = new EntityMetadata { MetadataId = record.ObjectId, LogicalName = "account" };
                    return response;
                }
                if (type == 2)
                {
                    Assert.IsInstanceOfType(request, typeof(RetrieveAttributeRequest));
                    var response = new RetrieveAttributeResponse();
                    var metadata = new StringAttributeMetadata { MetadataId = record.ObjectId, LogicalName = "name" };
                    // EntityLogicalName is read-only to ordinary SDK consumers; populate a server response fixture.
                    typeof(AttributeMetadata).GetProperty("EntityLogicalName").SetValue(metadata, "account");
                    response.Results["AttributeMetadata"] = metadata;
                    return response;
                }
                Assert.IsInstanceOfType(request, typeof(RetrieveRelationshipRequest));
                var relationshipResponse = new RetrieveRelationshipResponse();
                relationshipResponse.Results["RelationshipMetadata"] = new OneToManyRelationshipMetadata
                {
                    MetadataId = record.ObjectId, SchemaName = "new_account_contact"
                };
                return relationshipResponse;
            };
            var identity = new DataverseComponentIdentityResolver().Resolve(service, solution.Environment, record, CancellationToken.None);
            Assert.AreEqual(IdentityResolutionStatus.Resolved, identity.Status);
            Assert.AreEqual(kind, identity.ComponentTypeKey);
            Assert.AreEqual(expectedKey, identity.ComparisonKey);
            Assert.AreSame(record, identity.Record);
            Assert.AreEqual(0, service.Calls);
        }

        [DataTestMethod]
        [DataRow(61, "webresource", "webresourceid", "name", "new_/script.js")]
        [DataRow(29, "workflow", "workflowid", "uniquename", "new_Process")]
        [DataRow(380, "environmentvariabledefinition", "environmentvariabledefinitionid", "schemaname", "new_Setting")]
        public void RecordResolversSelectOnlyIdAndPortableName(int type, string table, string primaryId, string attribute, string key)
        {
            var solution = Solution(); var record = Identity("unused", type).Record;
            var service = Service(solution, query =>
            {
                Assert.AreEqual(table, query.EntityName);
                CollectionAssert.AreEquivalent(new[] { primaryId, attribute }, query.ColumnSet.Columns.ToArray());
                Assert.AreEqual(primaryId, query.Criteria.Conditions.Single().AttributeName);
                Assert.AreEqual(record.ObjectId, query.Criteria.Conditions.Single().Values.Single());
                Assert.AreEqual(2, query.TopCount);
                return Rows(new Entity(table, record.ObjectId.Value) { [attribute] = key });
            });
            var result = new DataverseComponentIdentityResolver().Resolve(service, solution.Environment, record, CancellationToken.None);
            Assert.AreEqual(IdentityResolutionStatus.Resolved, result.Status);
            Assert.AreEqual(key, result.ComparisonKey);
            Assert.AreSame(record, result.Record);
        }

        [TestMethod]
        public void SecurityRolesUseTemplateIdAndNeverRoleNameOrLocalRoleId()
        {
            var solution = Solution(); var record = Identity("unused", 20).Record; var templateId = Guid.NewGuid();
            var service = Service(solution, query =>
            {
                Assert.AreEqual("role", query.EntityName);
                CollectionAssert.AreEquivalent(new[] { "roleid", "roletemplateid" }, query.ColumnSet.Columns.ToArray());
                return Rows(new Entity("role", record.ObjectId.Value) { ["roletemplateid"] = new EntityReference("roletemplate", templateId) });
            });
            var resolver = new DataverseComponentIdentityResolver();
            var resolved = resolver.Resolve(service, solution.Environment, record, CancellationToken.None);
            Assert.AreEqual(templateId.ToString("D"), resolved.ComparisonKey);
            service.RetrievePage = q => Rows(new Entity("role", record.ObjectId.Value) { ["name"] = "Same display name" });
            var customRole = resolver.Resolve(service, solution.Environment, record, CancellationToken.None);
            Assert.AreEqual(IdentityResolutionStatus.Unresolved, customRole.Status);
            Assert.IsNull(customRole.ComparisonKey);
        }

        [DataTestMethod]
        [DataRow(true)]
        [DataRow(false)]
        public void ProcessWithoutUniqueNameOrDeletedObjectRemainsUnresolved(bool recordExists)
        {
            var solution = Solution(); var record = Identity("unused", 29).Record;
            var service = Service(solution, q => recordExists ? Rows(new Entity("workflow", record.ObjectId.Value) { ["name"] = "Same display name" }) : Rows());
            var result = new DataverseComponentIdentityResolver().Resolve(service, solution.Environment, record, CancellationToken.None);
            Assert.AreEqual(IdentityResolutionStatus.Unresolved, result.Status);
            Assert.IsNull(result.ComparisonKey);
        }

        [TestMethod]
        public void ConnectionReferencesDiscoverEnvironmentSpecificCodesAndMatchCanonically()
        {
            var resolver = new DataverseComponentIdentityResolver();
            var snapshots = new MembershipSnapshot[2];
            for (int side = 0; side < 2; side++)
            {
                var solution = Solution(); int localTypeCode = side == 0 ? 10027 : 10150;
                var record = Identity("unused", localTypeCode).Record;
                int mappingCalls = 0;
                var service = Service(solution, query =>
                {
                    if (query.EntityName == "solutioncomponentdefinition")
                    {
                        mappingCalls++;
                        Assert.AreEqual("primaryentityname", query.Criteria.Conditions.Single().AttributeName);
                        Assert.AreEqual("connectionreference", query.Criteria.Conditions.Single().Values.Single());
                        return Rows(new Entity("solutioncomponentdefinition", Guid.NewGuid()) { ["objecttypecode"] = localTypeCode });
                    }
                    Assert.AreEqual("connectionreference", query.EntityName);
                    CollectionAssert.AreEquivalent(new[] { "connectionreferenceid", "connectionreferencelogicalname" }, query.ColumnSet.Columns.ToArray());
                    return Rows(new Entity("connectionreference", (Guid)query.Criteria.Conditions.Single().Values.Single())
                    {
                        ["connectionreferencelogicalname"] = "new_shared"
                    });
                });
                var input = MembershipSnapshot.Complete(solution, new[] { new ComponentIdentity(record, IdentityResolutionStatus.Unresolved) }, DateTimeOffset.UtcNow);
                snapshots[side] = resolver.ResolveSnapshot(service, input, CancellationToken.None);
                Assert.AreEqual(1, mappingCalls);
                Assert.AreEqual(localTypeCode, snapshots[side].Components.Single().Record.ComponentType);
                Assert.AreEqual("connectionreference", snapshots[side].Components.Single().ComponentTypeKey);
            }
            var result = new SolutionMembershipComparer().Compare(snapshots[0], snapshots[1]).Single();
            Assert.AreEqual(MembershipPresence.PresentInBoth, result.Presence);
        }

        [TestMethod]
        public void UnsupportedTypesKeepRawRecordsAndDiscoverMappingOnlyOncePerSnapshot()
        {
            var solution = Solution(); var first = Identity("unused", 99999).Record; var second = Identity("unused", 99998).Record;
            var service = Service(solution, query =>
            {
                Assert.AreEqual("solutioncomponentdefinition", query.EntityName);
                return Rows();
            });
            var input = MembershipSnapshot.Complete(solution, new[] { new ComponentIdentity(first, IdentityResolutionStatus.Unresolved),
                new ComponentIdentity(second, IdentityResolutionStatus.Unresolved) }, DateTimeOffset.UtcNow);
            var result = new DataverseComponentIdentityResolver().ResolveSnapshot(service, input, CancellationToken.None);
            Assert.IsTrue(result.Components.All(c => c.Status == IdentityResolutionStatus.Unsupported));
            Assert.AreSame(first, result.Components[0].Record);
            Assert.AreSame(second, result.Components[1].Record);
            Assert.AreEqual(1, service.Calls);
            Assert.AreEqual(input.CapturedAt, result.CapturedAt);
        }

        [DataTestMethod]
        [DataRow(3)]
        [DataRow(11)]
        [DataRow(12)]
        [DataRow(9)]
        [DataRow(91)]
        [DataRow(381)]
        public void KnownUnsupportedTypesNeverDiscoverOrLookupConnectionReferences(int type)
        {
            var solution = Solution(); var record = Identity("unused", type).Record;
            var service = Service(solution, query =>
            {
                Assert.Fail("Known unsupported types must not query " + query.EntityName);
                return Rows();
            });
            var result = new DataverseComponentIdentityResolver().Resolve(service, solution.Environment, record, CancellationToken.None);
            Assert.AreEqual(IdentityResolutionStatus.Unsupported, result.Status);
            Assert.AreSame(record, result.Record);
            Assert.IsNull(result.ComparisonKey);
            Assert.AreEqual(0, service.Calls);
        }

        [DataTestMethod]
        [DataRow(3)]
        [DataRow(11)]
        [DataRow(12)]
        [DataRow(1)]
        [DataRow(61)]
        [DataRow(91)]
        [DataRow(381)]
        [DataRow(432)]
        public void DiscoveryCollidingWithKnownComponentKindIsAmbiguousAndCached(int discoveredCode)
        {
            var solution = Solution();
            var records = new[] { Identity(null, 10027, IdentityResolutionStatus.Unresolved),
                Identity(null, 10150, IdentityResolutionStatus.Unresolved) };
            var service = Service(solution, query =>
            {
                Assert.AreEqual("solutioncomponentdefinition", query.EntityName, "An invalid mapping must not issue an identity lookup.");
                return Rows(new Entity("solutioncomponentdefinition", Guid.NewGuid()) { ["objecttypecode"] = discoveredCode });
            });
            var snapshot = MembershipSnapshot.Complete(solution, records, DateTimeOffset.UtcNow);
            var result = new DataverseComponentIdentityResolver().ResolveSnapshot(service, snapshot, CancellationToken.None);
            Assert.AreEqual(1, service.Calls);
            for (int index = 0; index < records.Length; index++)
            {
                Assert.AreEqual(IdentityResolutionStatus.Ambiguous, result.Components[index].Status);
                Assert.AreSame(records[index].Record, result.Components[index].Record);
                Assert.IsNull(result.Components[index].ComparisonKey);
                StringAssert.Contains(result.Components[index].Diagnostic, "conflicts");
            }
        }

        [DataTestMethod]
        [DataRow(10027)]
        [DataRow(10150)]
        public void ValidDynamicConnectionCodeStillResolvesWithOneDiscoveryPerSnapshot(int discoveredCode)
        {
            var solution = Solution();
            var records = new[] { Identity(null, discoveredCode, IdentityResolutionStatus.Unresolved),
                Identity(null, discoveredCode, IdentityResolutionStatus.Unresolved) };
            int discoveryCalls = 0; int identityCalls = 0;
            var service = Service(solution, query =>
            {
                if (query.EntityName == "solutioncomponentdefinition")
                {
                    discoveryCalls++;
                    return Rows(new Entity("solutioncomponentdefinition", Guid.NewGuid()) { ["objecttypecode"] = discoveredCode });
                }
                Assert.AreEqual("connectionreference", query.EntityName);
                identityCalls++;
                return Rows(query.Criteria.Conditions.Single().Values.Cast<Guid>().Select((id, index) =>
                    new Entity("connectionreference", id)
                    {
                        ["connectionreferencelogicalname"] = "new_reference" + (index + 1)
                    }).ToArray());
            });
            var snapshot = MembershipSnapshot.Complete(solution, records, DateTimeOffset.UtcNow);
            var result = new DataverseComponentIdentityResolver().ResolveSnapshot(service, snapshot, CancellationToken.None);
            Assert.AreEqual(1, discoveryCalls);
            Assert.AreEqual(1, identityCalls);
            for (int index = 0; index < records.Length; index++)
            {
                Assert.AreEqual(IdentityResolutionStatus.Resolved, result.Components[index].Status);
                Assert.AreEqual("connectionreference", result.Components[index].ComponentTypeKey);
                Assert.AreEqual("new_reference" + (index + 1), result.Components[index].ComparisonKey);
                Assert.AreSame(records[index].Record, result.Components[index].Record);
            }
        }

        [TestMethod]
        public void AmbiguousObjectLookupRemainsAmbiguous()
        {
            var solution = Solution(); var record = Identity("unused", 29).Record;
            var service = Service(solution, q => Rows(new Entity("workflow", record.ObjectId.Value), new Entity("workflow", record.ObjectId.Value)));
            var result = new DataverseComponentIdentityResolver().Resolve(service, solution.Environment, record, CancellationToken.None);
            Assert.AreEqual(IdentityResolutionStatus.Ambiguous, result.Status);
        }

        [TestMethod]
        public void AmbiguousConnectionMappingIsNotGuessed()
        {
            var solution = Solution(); var record = Identity("unused", 10150).Record;
            var service = Service(solution, q => Rows(new Entity("solutioncomponentdefinition"), new Entity("solutioncomponentdefinition")));
            var result = new DataverseComponentIdentityResolver().Resolve(service, solution.Environment, record, CancellationToken.None);
            Assert.AreEqual(IdentityResolutionStatus.Ambiguous, result.Status);
            Assert.IsNull(result.ComparisonKey);
        }

        [TestMethod]
        public void MissingObjectIdRemainsRawAndDoesNotIssueObjectQuery()
        {
            var solution = Solution(); var record = new SolutionComponentRecord(Guid.NewGuid(), 61, null);
            var service = Service(solution);
            var result = new DataverseComponentIdentityResolver().Resolve(service, solution.Environment, record, CancellationToken.None);
            Assert.AreSame(record, result.Record);
            Assert.AreEqual(IdentityResolutionStatus.Unresolved, result.Status);
            Assert.AreEqual(0, service.Calls);
        }

        [TestMethod]
        public void IdentityReadFaultIsUnresolvedWithDiagnostic()
        {
            var solution = Solution(); var service = Service(solution, q => throw new FaultException("Denied"));
            var result = new DataverseComponentIdentityResolver().Resolve(service, solution.Environment, Identity("unused", 61).Record, CancellationToken.None);
            Assert.AreEqual(IdentityResolutionStatus.Unresolved, result.Status);
            StringAssert.Contains(result.Diagnostic, "Denied");
        }

        [TestMethod]
        public void MappingReadFaultIsUnresolvedRatherThanUnsupportedOrMissing()
        {
            var solution = Solution(); var service = Service(solution, q => throw new FaultException("Mapping denied"));
            var result = new DataverseComponentIdentityResolver().Resolve(service, solution.Environment, Identity("unused", 10150).Record, CancellationToken.None);
            Assert.AreEqual(IdentityResolutionStatus.Unresolved, result.Status);
        }

        [TestMethod]
        public void CancellationDuringIdentityReadPropagates()
        {
            var solution = Solution();
            using (var cancellation = new CancellationTokenSource())
            {
                var service = Service(solution, q => { cancellation.Cancel(); return Rows(); });
                Assert.ThrowsException<OperationCanceledException>(() => new DataverseComponentIdentityResolver().Resolve(
                    service, solution.Environment, Identity("unused", 61).Record, cancellation.Token));
            }
        }

        [TestMethod]
        public void AbsentAndUnavailableSnapshotsNeedNoResolutionQueries()
        {
            var solution = Solution(); var resolver = new DataverseComponentIdentityResolver();
            var absent = MembershipSnapshot.Absent(solution.Environment, solution.UniqueName, DateTimeOffset.UtcNow);
            var unavailable = MembershipSnapshot.Unavailable(solution.Environment, solution.UniqueName, DateTimeOffset.UtcNow, "Disconnected");
            Assert.AreSame(absent, resolver.ResolveSnapshot(null, absent, CancellationToken.None));
            Assert.AreSame(unavailable, resolver.ResolveSnapshot(null, unavailable, CancellationToken.None));
        }
    }
}
