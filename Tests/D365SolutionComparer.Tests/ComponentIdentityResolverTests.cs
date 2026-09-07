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
using Microsoft.Xrm.Sdk.Metadata.Query;
using Microsoft.Xrm.Sdk.Query;
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
            StringAssert.Contains(result.Diagnostic, recordExists
                ? "Unsupported workflow record type (missing)"
                : "Raw workflow row was not found");
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
            int connectionDiscoveryCalls = 0; int familyDiscoveryCalls = 0;
            var service = Service(solution, query =>
            {
                Assert.AreEqual("solutioncomponentdefinition", query.EntityName);
                if (query.Criteria.Conditions.Single().AttributeName == "primaryentityname")
                    connectionDiscoveryCalls++;
                else
                {
                    familyDiscoveryCalls++;
                    Assert.AreEqual("objecttypecode", query.Criteria.Conditions.Single().AttributeName);
                    CollectionAssert.AreEquivalent(new[] { 99998, 99999 }, query.Criteria.Conditions.Single()
                        .Values.Cast<int>().ToArray());
                    CollectionAssert.AreEquivalent(new[] { "objecttypecode", "name", "primaryentityname" },
                        query.ColumnSet.Columns.ToArray());
                }
                return Rows();
            });
            service.ExecuteRequest = request => request is WhoAmIRequest
                ? (OrganizationResponse)WhoAmI(solution.Environment.OrganizationId) : MetadataRows();
            var input = MembershipSnapshot.Complete(solution, new[] { new ComponentIdentity(first, IdentityResolutionStatus.Unresolved),
                new ComponentIdentity(second, IdentityResolutionStatus.Unresolved) }, DateTimeOffset.UtcNow);
            var result = new DataverseComponentIdentityResolver().ResolveSnapshot(service, input, CancellationToken.None);
            Assert.IsTrue(result.Components.All(c => c.Status == IdentityResolutionStatus.Unsupported));
            Assert.IsTrue(result.Components.All(c => c.SemanticKind == null));
            Assert.AreSame(first, result.Components[0].Record);
            Assert.AreSame(second, result.Components[1].Record);
            Assert.AreEqual(1, connectionDiscoveryCalls);
            Assert.AreEqual(1, familyDiscoveryCalls);
            Assert.AreEqual(2, service.Calls);
            Assert.AreEqual(input.CapturedAt, result.CapturedAt);
        }

        [TestMethod]
        public void RegisteredFamilyUsesSameSemanticBucketAcrossEnvironmentLocalTypeCodes()
        {
            var results = new ComponentIdentity[2];
            var solutions = new D365SolutionComparer.Models.Identity.SolutionIdentity[2];
            for (int side = 0; side < 2; side++)
            {
                var solution = Solution(); int type = side == 0 ? 10266 : 10267;
                solutions[side] = solution;
                var record = Identity(null, type, IdentityResolutionStatus.Unresolved).Record;
                var service = Service(solution, query =>
                {
                    if (query.Criteria.Conditions.Single().AttributeName == "primaryentityname") return Rows();
                    return Rows(Definition(type, "Contoso.Education.Assessment", "contoso_assessment"));
                });
                results[side] = new DataverseComponentIdentityResolver().Resolve(service,
                    solution.Environment, record, CancellationToken.None);
                Assert.AreEqual(2, service.Calls);
                Assert.AreEqual(IdentityResolutionStatus.Unsupported, results[side].Status);
                Assert.IsNull(results[side].ComparisonKey);
                Assert.AreEqual("componenttype:" + type, results[side].ComponentTypeKey);
                Assert.AreEqual("Contoso.Education.Assessment", results[side].RegisteredDefinition.Name);
                Assert.AreEqual("contoso_assessment", results[side].RegisteredDefinition.PrimaryEntityName);
            }

            Assert.AreEqual(results[0].SemanticKind, results[1].SemanticKind);
            Assert.AreNotEqual(results[0].ComponentTypeKey, results[1].ComponentTypeKey);
            var comparison = new SolutionMembershipComparer().Compare(
                MembershipSnapshot.Complete(solutions[0], new[] { results[0] }, DateTimeOffset.UtcNow),
                MembershipSnapshot.Complete(solutions[1], new[] { results[1] }, DateTimeOffset.UtcNow));
            Assert.AreEqual(2, comparison.Count);
            Assert.IsTrue(comparison.All(item => item.Presence == MembershipPresence.Indeterminate));
        }

        [TestMethod]
        public void RepeatedRawTypeUsesOneGroupedDefinitionLookupAndRetainsEveryRow()
        {
            var solution = Solution(); const int type = 10075; int familyQueries = 0;
            var records = Enumerable.Range(0, 15)
                .Select(index => Identity(null, type, IdentityResolutionStatus.Unresolved)).ToArray();
            var service = Service(solution, query =>
            {
                if (query.Criteria.Conditions.Single().AttributeName == "primaryentityname") return Rows();
                familyQueries++;
                return Rows(Definition(type, "Microsoft.Education.Component", "msdyn_educationcomponent"));
            });

            var result = new DataverseComponentIdentityResolver().ResolveSnapshot(service,
                MembershipSnapshot.Complete(solution, records, DateTimeOffset.UtcNow), CancellationToken.None);

            Assert.AreEqual(1, familyQueries);
            Assert.AreEqual(15, result.Components.Count);
            Assert.IsTrue(result.Components.All(item => item.Status == IdentityResolutionStatus.Unsupported));
            Assert.AreEqual(1, result.Components.Select(item => item.SemanticKind).Distinct().Count());
            Assert.IsTrue(result.Components.All(item => item.RegisteredDefinition != null));
        }

        [TestMethod]
        public void MultipleDefinitionsForOneRawTypeRemainBroadAndAmbiguous()
        {
            var solution = Solution(); const int type = 10072;
            var service = Service(solution, query => query.Criteria.Conditions.Single().AttributeName ==
                "primaryentityname" ? Rows() : Rows(
                    Definition(type, "Family.One", "family_one"),
                    Definition(type, "Family.Two", "family_two")));

            var result = new DataverseComponentIdentityResolver().Resolve(service, solution.Environment,
                Identity(null, type, IdentityResolutionStatus.Unresolved).Record, CancellationToken.None);

            Assert.AreEqual(IdentityResolutionStatus.Ambiguous, result.Status);
            Assert.IsNull(result.SemanticKind);
            Assert.IsNull(result.RegisteredDefinition);
        }

        [TestMethod]
        public void IncompleteRegisteredDefinitionRemainsBroad()
        {
            var solution = Solution(); const int type = 511;
            var service = Service(solution, query => query.Criteria.Conditions.Single().AttributeName ==
                "primaryentityname" ? Rows() : Rows(new Entity("solutioncomponentdefinition", Guid.NewGuid())
                {
                    ["objecttypecode"] = type,
                    ["primaryentityname"] = "msdyn_component"
                }));

            var result = new DataverseComponentIdentityResolver().Resolve(service, solution.Environment,
                Identity(null, type, IdentityResolutionStatus.Unresolved).Record, CancellationToken.None);

            Assert.AreEqual(IdentityResolutionStatus.Unresolved, result.Status);
            Assert.IsNull(result.SemanticKind);
            Assert.IsNull(result.RegisteredDefinition);
            StringAssert.Contains(result.Diagnostic, "stable name");
        }

        [TestMethod]
        public void DefinitionDiscoveryFaultLeavesCandidateBroadAndUnresolved()
        {
            var solution = Solution(); const int type = 10276;
            var service = Service(solution, query =>
            {
                if (query.Criteria.Conditions.Single().AttributeName == "primaryentityname") return Rows();
                throw new FaultException("Definition access denied");
            });

            var result = new DataverseComponentIdentityResolver().Resolve(service, solution.Environment,
                Identity(null, type, IdentityResolutionStatus.Unresolved).Record, CancellationToken.None);

            Assert.AreEqual(IdentityResolutionStatus.Unresolved, result.Status);
            Assert.IsNull(result.SemanticKind);
            StringAssert.Contains(result.Diagnostic, "Definition access denied");
        }

        [TestMethod]
        public void CancellationDuringDefinitionDiscoveryPropagates()
        {
            var solution = Solution(); const int type = 10276;
            using (var cancellation = new CancellationTokenSource())
            {
                var service = Service(solution, query =>
                {
                    if (query.Criteria.Conditions.Single().AttributeName == "primaryentityname") return Rows();
                    cancellation.Cancel();
                    return Rows(Definition(type, "Family.Cancelled", "cancelled_component"));
                });
                Assert.ThrowsException<OperationCanceledException>(() =>
                    new DataverseComponentIdentityResolver().Resolve(service, solution.Environment,
                        Identity(null, type, IdentityResolutionStatus.Unresolved).Record, cancellation.Token));
            }
        }

        [TestMethod]
        public void BroadTypesCollectGroupedEntityMetadataDiagnosticsWithoutChangingClassification()
        {
            var solution = Solution();
            var records = new[]
            {
                Identity(null, 80, IdentityResolutionStatus.Unresolved),
                Identity(null, 80, IdentityResolutionStatus.Unresolved),
                Identity(null, 511, IdentityResolutionStatus.Unresolved)
            };
            int metadataQueries = 0;
            var service = Service(solution, query => Rows());
            service.ExecuteRequest = request =>
            {
                if (request is WhoAmIRequest) return WhoAmI(solution.Environment.OrganizationId);
                metadataQueries++;
                var metadataRequest = (RetrieveMetadataChangesRequest)request;
                CollectionAssert.AreEquivalent(new[] { "ObjectTypeCode", "LogicalName", "SchemaName" },
                    metadataRequest.Query.Properties.PropertyNames.ToArray());
                Assert.AreEqual(LogicalOperator.Or, metadataRequest.Query.Criteria.FilterOperator);
                Assert.AreEqual(2, metadataRequest.Query.Criteria.Conditions.Count);
                Assert.IsTrue(metadataRequest.Query.Criteria.Conditions.All(condition =>
                    condition.PropertyName == "ObjectTypeCode" &&
                    condition.ConditionOperator == MetadataConditionOperator.Equals &&
                    condition.Value != null && condition.Value.GetType() == typeof(int)));
                CollectionAssert.AreEquivalent(new[] { 80, 511 }, metadataRequest.Query.Criteria.Conditions
                    .Select(condition => (int)condition.Value).ToArray());
                return MetadataRows(EntityMetadata(80, "sample_type_80", "SampleType80"),
                    EntityMetadata(511, "sample_type_511", "SampleType511"));
            };

            var result = new DataverseComponentIdentityResolver().ResolveSnapshot(service,
                MembershipSnapshot.Complete(solution, records, DateTimeOffset.UtcNow), CancellationToken.None);

            Assert.AreEqual(1, metadataQueries);
            Assert.AreEqual(3, result.Components.Count);
            Assert.IsTrue(result.Components.All(item => item.Status == IdentityResolutionStatus.Unsupported));
            Assert.IsTrue(result.Components.All(item => item.SemanticKind == null));
            Assert.IsTrue(result.Components.All(item => item.ComparisonKey == null));
            Assert.IsTrue(result.Components.Where(item => item.Record.ComponentType == 80).All(item =>
                item.Diagnostic.Contains("sample_type_80") && item.Diagnostic.Contains("SampleType80")));
            StringAssert.Contains(result.Components.Single(item => item.Record.ComponentType == 511).Diagnostic,
                "sample_type_511");

            var sourceColumn = Identity("account.name", 2, kind: ComponentSemanticKinds.Column);
            var comparison = new SolutionMembershipComparer().Compare(
                MembershipSnapshot.Complete(Solution(), new[] { sourceColumn }, DateTimeOffset.UtcNow), result);
            Assert.AreEqual(MembershipPresence.Indeterminate,
                comparison.Single(item => item.Source == sourceColumn).Presence);
        }

        [TestMethod]
        public void MissingEntityMetadataCandidateIsReportedAndRemainsBroad()
        {
            var solution = Solution(); var service = Service(solution, query => Rows());
            service.ExecuteRequest = request =>
            {
                if (request is WhoAmIRequest) return WhoAmI(solution.Environment.OrganizationId);
                return MetadataRows();
            };

            var result = new DataverseComponentIdentityResolver().Resolve(service, solution.Environment,
                Identity(null, 80, IdentityResolutionStatus.Unresolved).Record, CancellationToken.None);

            Assert.AreEqual(IdentityResolutionStatus.Unsupported, result.Status);
            Assert.IsNull(result.SemanticKind);
            StringAssert.Contains(result.Diagnostic, "No entity metadata candidate");
            StringAssert.Contains(result.Diagnostic, "ObjectTypeCode 80");
        }

        [TestMethod]
        public void MultipleEntityMetadataCandidatesAreReportedWithoutClassification()
        {
            var solution = Solution(); var service = Service(solution, query => Rows());
            service.ExecuteRequest = request =>
            {
                if (request is WhoAmIRequest) return WhoAmI(solution.Environment.OrganizationId);
                return MetadataRows(EntityMetadata(511, "first", "First"),
                    EntityMetadata(511, "second", "Second"));
            };

            var result = new DataverseComponentIdentityResolver().Resolve(service, solution.Environment,
                Identity(null, 511, IdentityResolutionStatus.Unresolved).Record, CancellationToken.None);

            Assert.AreEqual(IdentityResolutionStatus.Unsupported, result.Status);
            Assert.IsNull(result.SemanticKind);
            StringAssert.Contains(result.Diagnostic, "Multiple entity metadata candidates");
            StringAssert.Contains(result.Diagnostic, "first");
            StringAssert.Contains(result.Diagnostic, "second");
        }

        [TestMethod]
        public void EntityMetadataDiagnosticFaultAndCancellationRemainConservative()
        {
            var solution = Solution();
            var faulted = Service(solution, query => Rows());
            faulted.ExecuteRequest = request => request is WhoAmIRequest
                ? WhoAmI(solution.Environment.OrganizationId) : throw new FaultException("Metadata denied");
            var faultedResult = new DataverseComponentIdentityResolver().Resolve(faulted,
                solution.Environment, Identity(null, 80, IdentityResolutionStatus.Unresolved).Record,
                CancellationToken.None);
            Assert.AreEqual(IdentityResolutionStatus.Unsupported, faultedResult.Status);
            Assert.IsNull(faultedResult.SemanticKind);
            StringAssert.Contains(faultedResult.Diagnostic, "Metadata denied");

            using (var cancellation = new CancellationTokenSource())
            {
                var cancelled = Service(solution, query => Rows());
                cancelled.ExecuteRequest = request =>
                {
                    if (request is WhoAmIRequest) return WhoAmI(solution.Environment.OrganizationId);
                    cancellation.Cancel();
                    return MetadataRows();
                };
                Assert.ThrowsException<OperationCanceledException>(() =>
                    new DataverseComponentIdentityResolver().Resolve(cancelled, solution.Environment,
                        Identity(null, 511, IdentityResolutionStatus.Unresolved).Record, cancellation.Token));
            }
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
            Assert.AreEqual(2, service.Calls);
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

        private static Entity Definition(int objectTypeCode, string name, string primaryEntityName)
        {
            return new Entity("solutioncomponentdefinition", Guid.NewGuid())
            {
                ["objecttypecode"] = objectTypeCode,
                ["name"] = name,
                ["primaryentityname"] = primaryEntityName
            };
        }

        private static EntityMetadata EntityMetadata(int objectTypeCode, string logicalName,
            string schemaName)
        {
            var metadata = new EntityMetadata { LogicalName = logicalName, SchemaName = schemaName };
            typeof(EntityMetadata).GetProperty("ObjectTypeCode").SetValue(metadata, (int?)objectTypeCode);
            return metadata;
        }

        private static RetrieveMetadataChangesResponse MetadataRows(params EntityMetadata[] items)
        {
            var response = new RetrieveMetadataChangesResponse();
            var metadata = new EntityMetadataCollection();
            metadata.AddRange(items);
            response.Results["EntityMetadata"] = metadata;
            return response;
        }
    }
}
