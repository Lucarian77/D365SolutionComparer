using System;
using System.Collections.Generic;
using System.Linq;
using System.ServiceModel;
using System.Threading;
using D365SolutionComparer.Models.Identity;
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
                Identity(null, 98765, IdentityResolutionStatus.Unresolved),
                Identity(null, 98765, IdentityResolutionStatus.Unresolved),
                Identity(null, 511, IdentityResolutionStatus.Unresolved)
            };
            int metadataQueries = 0; int choiceQueries = 0;
            var service = Service(solution, query => Rows());
            service.ExecuteRequest = request =>
            {
                if (request is WhoAmIRequest) return WhoAmI(solution.Environment.OrganizationId);
                if (request is RetrieveAttributeRequest)
                {
                    choiceQueries++;
                    return ComponentTypeChoices(
                        new OptionMetadata(new Label("Unclassified Test Component", 1033), 98765),
                        new OptionMetadata(new Label("Team Template", 1033), 511));
                }
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
                CollectionAssert.AreEquivalent(new[] { 98765, 511 }, metadataRequest.Query.Criteria.Conditions
                    .Select(condition => (int)condition.Value).ToArray());
                return MetadataRows(EntityMetadata(98765, "sample_type_98765", "SampleType98765"),
                    EntityMetadata(511, "sample_type_511", "SampleType511"));
            };

            var result = new DataverseComponentIdentityResolver().ResolveSnapshot(service,
                MembershipSnapshot.Complete(solution, records, DateTimeOffset.UtcNow), CancellationToken.None);

            Assert.AreEqual(1, metadataQueries);
            Assert.AreEqual(1, choiceQueries);
            Assert.AreEqual(3, result.Components.Count);
            Assert.IsTrue(result.Components.All(item => item.Status == IdentityResolutionStatus.Unsupported));
            Assert.IsTrue(result.Components.All(item => item.SemanticKind == null));
            Assert.IsTrue(result.Components.All(item => item.ComparisonKey == null));
            Assert.IsTrue(result.Components.Where(item => item.Record.ComponentType == 98765).All(item =>
                item.Diagnostic.Contains("sample_type_98765") && item.Diagnostic.Contains("SampleType98765")));
            StringAssert.Contains(result.Components.Single(item => item.Record.ComponentType == 511).Diagnostic,
                "sample_type_511");

            var sourceColumn = Identity("account.name", 2, kind: ComponentSemanticKinds.Column);
            var comparison = new SolutionMembershipComparer().Compare(
                MembershipSnapshot.Complete(Solution(), new[] { sourceColumn }, DateTimeOffset.UtcNow), result);
            Assert.AreEqual(MembershipPresence.Indeterminate,
                comparison.Single(item => item.Source == sourceColumn).Presence);
        }

        [TestMethod]
        public void BroadTypesKeepStableChoiceDiagnosticAndExposeRawIdsAsCoverageEvidence()
        {
            var solution = Solution();
            var objectIds = new[] { Guid.NewGuid(), Guid.NewGuid(), Guid.NewGuid() };
            var componentIds = new[] { Guid.NewGuid(), Guid.NewGuid(), Guid.NewGuid() };
            var records = new[]
            {
                new ComponentIdentity(new SolutionComponentRecord(componentIds[0], 98765, objectIds[0]),
                    IdentityResolutionStatus.Unresolved),
                new ComponentIdentity(new SolutionComponentRecord(componentIds[1], 98765, objectIds[1]),
                    IdentityResolutionStatus.Unresolved),
                new ComponentIdentity(new SolutionComponentRecord(componentIds[2], 511, objectIds[2]),
                    IdentityResolutionStatus.Unresolved)
            };
            int choiceRequests = 0;
            var service = Service(solution, query => Rows());
            service.ExecuteRequest = request =>
            {
                if (request is WhoAmIRequest) return WhoAmI(solution.Environment.OrganizationId);
                if (request is RetrieveMetadataChangesRequest) return MetadataRows();
                var attributeRequest = request as RetrieveAttributeRequest;
                Assert.IsNotNull(attributeRequest);
                Assert.AreEqual("solutioncomponent", attributeRequest.EntityLogicalName);
                Assert.AreEqual("componenttype", attributeRequest.LogicalName);
                Assert.IsFalse(attributeRequest.RetrieveAsIfPublished);
                choiceRequests++;
                return ComponentTypeChoices(
                    new OptionMetadata(new Label("Unclassified Test Component", 1033), 98765),
                    new OptionMetadata(new Label("Team Template", 1033), 511));
            };

            var result = new DataverseComponentIdentityResolver().ResolveSnapshot(service,
                MembershipSnapshot.Complete(solution, records, DateTimeOffset.UtcNow), CancellationToken.None);

            Assert.AreEqual(1, choiceRequests);
            Assert.AreEqual(3, result.Components.Count);
            Assert.IsTrue(result.Components.All(item => item.Status == IdentityResolutionStatus.Unsupported));
            Assert.IsTrue(result.Components.All(item => item.SemanticKind == null));
            Assert.IsTrue(result.Components.All(item => item.ComparisonKey == null));
            Assert.IsTrue(result.Components.Where(item => item.Record.ComponentType == 98765)
                .All(item => item.Diagnostic.Contains("1033:'Unclassified Test Component'")));
            StringAssert.Contains(result.Components.Single(item => item.Record.ComponentType == 511).Diagnostic,
                "1033:'Team Template'");

            var coverage = new MembershipCoverageDiagnosticsBuilder().Build(result);
            var type98765 = coverage.BroadRawComponentTypes.Single(item => item.ComponentType == 98765);
            Assert.AreEqual(2, type98765.DiagnosticGroups.Single().Count);
            Assert.AreEqual(2, type98765.Evidence.Count);
            Assert.AreEqual(1, coverage.BroadRawComponentTypes.Single(item => item.ComponentType == 511)
                .DiagnosticGroups.Single().Count);
            CollectionAssert.AreEquivalent(componentIds, coverage.BroadRawComponentTypes
                .SelectMany(item => item.Evidence).Select(item => item.SolutionComponentId).ToArray());
            CollectionAssert.AreEquivalent(objectIds, coverage.BroadRawComponentTypes
                .SelectMany(item => item.Evidence).Select(item => item.ObjectId.Value).ToArray());
            Assert.IsTrue(result.Components.All(item => !item.Diagnostic.Contains(item.Record.ObjectId.Value
                .ToString("D")) && !item.Diagnostic.Contains(item.Record.SolutionComponentId.ToString("D"))));
        }

        [TestMethod]
        public void Type80AppModuleLookupUsesUniqueNameAndRetainsOtherFieldsAsAuditEvidence()
        {
            var solution = Solution(); var firstId = Guid.NewGuid(); var secondId = Guid.NewGuid();
            var records = new[]
            {
                new ComponentIdentity(new SolutionComponentRecord(Guid.NewGuid(), 80, firstId),
                    IdentityResolutionStatus.Unresolved),
                new ComponentIdentity(new SolutionComponentRecord(Guid.NewGuid(), 80, secondId),
                    IdentityResolutionStatus.Unresolved)
            };
            int appModuleQueries = 0;
            var service = BroadTypeService(solution, query =>
            {
                appModuleQueries++;
                AssertAppModuleQuery(query, firstId, secondId);
                return Rows(AppModule(firstId, "new_First", "First App", false),
                    AppModule(secondId, "new_Second", "Second App", true));
            });

            var result = new DataverseComponentIdentityResolver().ResolveSnapshot(service,
                MembershipSnapshot.Complete(solution, records, DateTimeOffset.UtcNow), CancellationToken.None);

            Assert.AreEqual(1, appModuleQueries);
            Assert.IsTrue(result.Components.All(item => item.Status == IdentityResolutionStatus.Resolved));
            Assert.IsTrue(result.Components.All(item => item.SemanticKind == ComponentSemanticKinds.AppModule));
            Assert.AreEqual("new_First", result.Components.Single(item => item.Record.ObjectId == firstId)
                .ComparisonKey);
            Assert.AreEqual("new_Second", result.Components.Single(item => item.Record.ObjectId == secondId)
                .ComparisonKey);
            var coverage = new MembershipCoverageDiagnosticsBuilder().Build(result);
            Assert.IsFalse(coverage.BroadRawComponentTypes.Any(item => item.ComponentType == 80));
            var appModules = coverage.SemanticKinds.Single(item =>
                item.SemanticKind == ComponentSemanticKinds.AppModule);
            Assert.AreEqual(MembershipCoverageStatus.Complete, appModules.CoverageStatus);
            Assert.AreEqual(2, appModules.Resolved);
            Assert.AreEqual(2, appModules.AuditEvidence.Count);
            Assert.IsTrue(appModules.AuditEvidence.Single(item => item.ObjectId == firstId).DiagnosticEvidence
                .Single().Contains("uniquename='new_First'"));
            var firstEvidence = appModules.AuditEvidence.Single(item => item.ObjectId == firstId)
                .DiagnosticEvidence.Single();
            StringAssert.Contains(firstEvidence, "appmoduleid=" + firstId.ToString("D"));
            StringAssert.Contains(firstEvidence, "name='First App'");
            StringAssert.Contains(firstEvidence, "appmoduleidunique=");
            StringAssert.Contains(firstEvidence, "componentstate=0 ('Published')");
            StringAssert.Contains(firstEvidence, "ismanaged=False");
            Assert.IsTrue(appModules.AuditEvidence.Single(item => item.ObjectId == secondId).DiagnosticEvidence
                .Single().Contains("uniquename='new_Second'"));
            Assert.IsTrue(appModules.AuditEvidence.Single(item => item.ObjectId == secondId).DiagnosticEvidence
                .Single().Contains("ismanaged=True"));
        }

        [TestMethod]
        public void NonType80BroadCandidateDoesNotTriggerAppModuleLookup()
        {
            var solution = Solution(); int appModuleQueries = 0;
            var service = BroadTypeService(solution, query =>
            {
                appModuleQueries++;
                return Rows();
            });

            var result = new DataverseComponentIdentityResolver().Resolve(service, solution.Environment,
                new SolutionComponentRecord(Guid.NewGuid(), 511, Guid.NewGuid()), CancellationToken.None);

            Assert.AreEqual(0, appModuleQueries);
            Assert.AreEqual(IdentityResolutionStatus.Unsupported, result.Status);
            Assert.IsNull(result.SemanticKind);
            Assert.AreEqual(1, result.DiagnosticEvidence.Count);
            StringAssert.Contains(result.DiagnosticEvidence.Single(), "No teamtemplate row matched");
        }

        [TestMethod]
        public void Type80UsesStaticAppModuleClassificationWithoutDynamicDiscovery()
        {
            var solution = Solution(); var objectId = Guid.NewGuid(); int discoveryRequests = 0;
            var service = Service(solution, query => Rows(
                AppModule(objectId, "new_StaticApp", "Static App", false)));
            service.ExecuteRequest = request =>
            {
                if (request is WhoAmIRequest) return WhoAmI(solution.Environment.OrganizationId);
                discoveryRequests++;
                throw new NotSupportedException(request.RequestName);
            };

            var result = new DataverseComponentIdentityResolver().Resolve(service, solution.Environment,
                new SolutionComponentRecord(Guid.NewGuid(), 80, objectId), CancellationToken.None);

            Assert.AreEqual(0, discoveryRequests);
            Assert.AreEqual(IdentityResolutionStatus.Resolved, result.Status);
            Assert.AreEqual(ComponentSemanticKinds.AppModule, result.SemanticKind);
            Assert.AreEqual("new_StaticApp", result.ComparisonKey);
        }

        [TestMethod]
        public void Type80AppModuleLookupZeroResultIsUnresolvedWithinAppModuleKind()
        {
            var solution = Solution(); var objectId = Guid.NewGuid();
            var result = new DataverseComponentIdentityResolver().Resolve(
                BroadTypeService(solution, query => Rows()), solution.Environment,
                new SolutionComponentRecord(Guid.NewGuid(), 80, objectId), CancellationToken.None);

            Assert.AreEqual(IdentityResolutionStatus.Unresolved, result.Status);
            Assert.AreEqual(ComponentSemanticKinds.AppModule, result.SemanticKind);
            Assert.IsNull(result.ComparisonKey);
            StringAssert.Contains(result.Diagnostic, "No appmodule row matched");
            Assert.AreEqual(0, result.DiagnosticEvidence.Count);
        }

        [TestMethod]
        public void Type80AppModuleLookupDuplicateResultIsAmbiguousAndKeepsEvidence()
        {
            var solution = Solution(); var objectId = Guid.NewGuid();
            var service = BroadTypeService(solution, query => Rows(
                AppModule(objectId, "new_First", "First", false),
                AppModule(objectId, "new_Second", "Second", false)));

            var result = new DataverseComponentIdentityResolver().Resolve(service, solution.Environment,
                new SolutionComponentRecord(Guid.NewGuid(), 80, objectId), CancellationToken.None);

            Assert.AreEqual(IdentityResolutionStatus.Ambiguous, result.Status);
            Assert.AreEqual(ComponentSemanticKinds.AppModule, result.SemanticKind);
            Assert.IsNull(result.ComparisonKey);
            StringAssert.Contains(result.Diagnostic, "multiple records");
            Assert.AreEqual(2, result.DiagnosticEvidence.Count);
            Assert.IsTrue(result.DiagnosticEvidence.Any(item => item.Contains("uniquename='new_First'")));
            Assert.IsTrue(result.DiagnosticEvidence.Any(item => item.Contains("uniquename='new_Second'")));
        }

        [TestMethod]
        public void Type80AppModuleLookupBatchesAndDeduplicatesObjectIds()
        {
            var solution = Solution();
            var objectIds = Enumerable.Range(0, 201).Select(index => Guid.NewGuid()).ToList();
            var records = objectIds.Concat(new[] { objectIds[0] }).Select(objectId =>
                new ComponentIdentity(new SolutionComponentRecord(Guid.NewGuid(), 80, objectId),
                    IdentityResolutionStatus.Unresolved)).ToArray();
            var queriedIds = new List<Guid>(); int appModuleQueries = 0;
            var service = BroadTypeService(solution, query =>
            {
                appModuleQueries++;
                var ids = query.Criteria.Conditions.Single().Values.Cast<Guid>().ToList();
                Assert.AreEqual(ids.Count, ids.Distinct().Count());
                Assert.IsTrue(ids.Count <= 200);
                queriedIds.AddRange(ids);
                return Rows(ids.Select(id => AppModule(id, "app_" + id.ToString("N"), "App", false)).ToArray());
            });
            var counter = new D365SolutionComparer.Infrastructure.DataverseRequestCounter();

            var result = new DataverseComponentIdentityResolver().ResolveSnapshot(service,
                MembershipSnapshot.Complete(solution, records, DateTimeOffset.UtcNow),
                CancellationToken.None, counter);

            Assert.AreEqual(2, appModuleQueries);
            Assert.AreEqual(2, counter.GetQueryCount("appmodule"));
            Assert.AreEqual(201, queriedIds.Count);
            CollectionAssert.AreEquivalent(objectIds, queriedIds);
            Assert.AreEqual(202, result.Components.Count);
            Assert.IsTrue(result.Components.All(item => item.Status == IdentityResolutionStatus.Resolved &&
                item.SemanticKind == ComponentSemanticKinds.AppModule && item.ComparisonKey != null));
            Assert.AreEqual(result.Components[0].ComparisonKey, result.Components[201].ComparisonKey);
        }

        [TestMethod]
        public void Type80AppModuleLookupFaultIsUnresolvedWithinAppModuleKind()
        {
            var solution = Solution();
            var result = new DataverseComponentIdentityResolver().Resolve(
                BroadTypeService(solution, query => throw new FaultException("AppModule denied")),
                solution.Environment, new SolutionComponentRecord(Guid.NewGuid(), 80, Guid.NewGuid()),
                CancellationToken.None);

            Assert.AreEqual(IdentityResolutionStatus.Unresolved, result.Status);
            Assert.AreEqual(ComponentSemanticKinds.AppModule, result.SemanticKind);
            Assert.IsNull(result.ComparisonKey);
            StringAssert.Contains(result.Diagnostic, "AppModule denied");
            Assert.AreEqual(0, result.DiagnosticEvidence.Count);
        }

        [TestMethod]
        public void Type80AppModuleLookupUsesUniqueNameWhenOtherDiagnosticFieldsAreMissing()
        {
            var solution = Solution(); var objectId = Guid.NewGuid();
            var incomplete = AppModule(objectId, "new_Incomplete", "Incomplete", false);
            incomplete.Attributes.Remove("appmoduleidunique");
            var result = new DataverseComponentIdentityResolver().Resolve(
                BroadTypeService(solution, query => Rows(incomplete)), solution.Environment,
                new SolutionComponentRecord(Guid.NewGuid(), 80, objectId), CancellationToken.None);

            Assert.AreEqual(IdentityResolutionStatus.Resolved, result.Status);
            Assert.AreEqual(ComponentSemanticKinds.AppModule, result.SemanticKind);
            Assert.AreEqual("new_Incomplete", result.ComparisonKey);
            StringAssert.Contains(result.DiagnosticEvidence.Single(), "incomplete diagnostic data");
            StringAssert.Contains(result.DiagnosticEvidence.Single(), "appmoduleidunique=(not supplied)");
        }

        [TestMethod]
        public void Type80AppModuleLookupBlankUniqueNameIsUnresolvedAndKeepsAuditEvidence()
        {
            var solution = Solution(); var objectId = Guid.NewGuid();
            var result = new DataverseComponentIdentityResolver().Resolve(
                BroadTypeService(solution, query => Rows(AppModule(objectId, "  ", "Display only", false))),
                solution.Environment, new SolutionComponentRecord(Guid.NewGuid(), 80, objectId),
                CancellationToken.None);

            Assert.AreEqual(IdentityResolutionStatus.Unresolved, result.Status);
            Assert.AreEqual(ComponentSemanticKinds.AppModule, result.SemanticKind);
            Assert.IsNull(result.ComparisonKey);
            StringAssert.Contains(result.Diagnostic, "no nonblank uniquename");
            StringAssert.Contains(result.DiagnosticEvidence.Single(), "name='Display only'");
        }

        [TestMethod]
        public void BlankAppModuleUniqueNamesShareStableDiagnosticGroupAndKeepSeparateEvidence()
        {
            var solution = Solution(); var firstId = Guid.NewGuid(); var secondId = Guid.NewGuid();
            var records = new[]
            {
                new ComponentIdentity(new SolutionComponentRecord(Guid.NewGuid(), 80, firstId),
                    IdentityResolutionStatus.Unresolved),
                new ComponentIdentity(new SolutionComponentRecord(Guid.NewGuid(), 80, secondId),
                    IdentityResolutionStatus.Unresolved)
            };
            var result = new DataverseComponentIdentityResolver().ResolveSnapshot(
                BroadTypeService(solution, query => Rows(
                    AppModule(firstId, null, "First display", false),
                    AppModule(secondId, " ", "Second display", true))),
                MembershipSnapshot.Complete(solution, records, DateTimeOffset.UtcNow), CancellationToken.None);

            Assert.IsTrue(result.Components.All(item => item.Status == IdentityResolutionStatus.Unresolved));
            Assert.AreEqual(1, result.Components.Select(item => item.Diagnostic).Distinct().Count());
            Assert.IsTrue(result.Components.All(item => !item.Diagnostic.Contains(item.Record.ObjectId.Value
                .ToString("D"))));
            var bucket = new MembershipCoverageDiagnosticsBuilder().Build(result).SemanticKinds.Single(item =>
                item.SemanticKind == ComponentSemanticKinds.AppModule);
            Assert.AreEqual(2, bucket.DiagnosticGroups.Single().Count);
            Assert.AreEqual(2, bucket.AuditEvidence.Count);
            CollectionAssert.AreEquivalent(new[] { firstId, secondId }, bucket.AuditEvidence
                .Select(item => item.ObjectId.Value).ToArray());
            Assert.IsTrue(bucket.AuditEvidence.All(item => item.DiagnosticEvidence.Count == 1));
        }

        [TestMethod]
        public void DuplicateAppModulePortableKeysRemainAmbiguousDuringComparison()
        {
            var solution = Solution(); var firstId = Guid.NewGuid(); var secondId = Guid.NewGuid();
            var records = new[]
            {
                new ComponentIdentity(new SolutionComponentRecord(Guid.NewGuid(), 80, firstId),
                    IdentityResolutionStatus.Unresolved),
                new ComponentIdentity(new SolutionComponentRecord(Guid.NewGuid(), 80, secondId),
                    IdentityResolutionStatus.Unresolved)
            };
            var resolved = new DataverseComponentIdentityResolver().ResolveSnapshot(
                BroadTypeService(solution, query => Rows(
                    AppModule(firstId, "new_Duplicate", "First", false),
                    AppModule(secondId, "NEW_DUPLICATE", "Second", false))),
                MembershipSnapshot.Complete(solution, records, DateTimeOffset.UtcNow), CancellationToken.None);

            Assert.IsTrue(resolved.Components.All(item => item.Status == IdentityResolutionStatus.Resolved));
            var compared = new SolutionMembershipComparer().Compare(resolved,
                MembershipSnapshot.Complete(solution, new ComponentIdentity[0], DateTimeOffset.UtcNow));
            Assert.AreEqual(2, compared.Count);
            Assert.IsTrue(compared.All(item => item.Presence == MembershipPresence.Indeterminate));
            Assert.IsTrue(compared.All(item => item.Source.Status == IdentityResolutionStatus.Ambiguous));
        }

        [TestMethod]
        public void AppModulesMatchAcrossEnvironmentLocalIdsByUniqueNameOnly()
        {
            var sourceSolution = Solution(); var targetSolution = new SolutionIdentity(
                new EnvironmentIdentity(Guid.NewGuid(), "Target"), Guid.NewGuid(), sourceSolution.UniqueName);
            var sourceId = Guid.NewGuid(); var targetId = Guid.NewGuid();
            var resolver = new DataverseComponentIdentityResolver();
            var source = resolver.Resolve(BroadTypeService(sourceSolution, query => Rows(
                    AppModule(sourceId, "new_PortableApp", "DEV name", false))),
                sourceSolution.Environment, new SolutionComponentRecord(Guid.NewGuid(), 80, sourceId),
                CancellationToken.None);
            var target = resolver.Resolve(BroadTypeService(targetSolution, query => Rows(
                    AppModule(targetId, "NEW_PORTABLEAPP", "UAT name", true))),
                targetSolution.Environment, new SolutionComponentRecord(Guid.NewGuid(), 80, targetId),
                CancellationToken.None);

            Assert.AreNotEqual(sourceId, targetId);
            Assert.AreEqual(IdentityResolutionStatus.Resolved, source.Status);
            Assert.AreEqual(IdentityResolutionStatus.Resolved, target.Status);
            Assert.AreNotEqual(source.DiagnosticEvidence.Single(), target.DiagnosticEvidence.Single());
            var comparison = new SolutionMembershipComparer().Compare(
                MembershipSnapshot.Complete(sourceSolution, new[] { source }, DateTimeOffset.UtcNow),
                MembershipSnapshot.Complete(targetSolution, new[] { target }, DateTimeOffset.UtcNow));
            Assert.AreEqual(1, comparison.Count);
            Assert.AreEqual(MembershipPresence.PresentInBoth, comparison.Single().Presence);
        }

        [TestMethod]
        public void Type80MissingObjectIdIsUnresolvedWithoutAppModuleQuery()
        {
            var solution = Solution(); int queryCount = 0;
            var result = new DataverseComponentIdentityResolver().Resolve(
                BroadTypeService(solution, query => { queryCount++; return Rows(); }), solution.Environment,
                new SolutionComponentRecord(Guid.NewGuid(), 80, null), CancellationToken.None);

            Assert.AreEqual(0, queryCount);
            Assert.AreEqual(IdentityResolutionStatus.Unresolved, result.Status);
            Assert.AreEqual(ComponentSemanticKinds.AppModule, result.SemanticKind);
            Assert.IsNull(result.ComparisonKey);
            StringAssert.Contains(result.Diagnostic, "no usable object ID");
        }

        [TestMethod]
        public void CancellationDuringType80AppModuleLookupPropagates()
        {
            var solution = Solution();
            using (var cancellation = new CancellationTokenSource())
            {
                var service = BroadTypeService(solution, query =>
                {
                    cancellation.Cancel();
                    return Rows();
                });
                Assert.ThrowsException<OperationCanceledException>(() =>
                    new DataverseComponentIdentityResolver().Resolve(service, solution.Environment,
                        new SolutionComponentRecord(Guid.NewGuid(), 80, Guid.NewGuid()), cancellation.Token));
            }
        }

        [TestMethod]
        public void ComponentTypeChoiceDiagnosticFaultRemainsBroadWithSeparateRawEvidence()
        {
            var solution = Solution(); var objectId = Guid.NewGuid(); var componentId = Guid.NewGuid();
            var record = new SolutionComponentRecord(componentId, 511, objectId);
            var service = Service(solution, query => Rows());
            service.ExecuteRequest = request =>
            {
                if (request is WhoAmIRequest) return WhoAmI(solution.Environment.OrganizationId);
                if (request is RetrieveMetadataChangesRequest) return MetadataRows();
                throw new FaultException("Choice metadata denied");
            };

            var result = new DataverseComponentIdentityResolver().Resolve(service, solution.Environment,
                record, CancellationToken.None);

            Assert.AreEqual(IdentityResolutionStatus.Unsupported, result.Status);
            Assert.IsNull(result.SemanticKind);
            Assert.IsNull(result.ComparisonKey);
            StringAssert.Contains(result.Diagnostic, "Choice metadata denied");
            Assert.IsFalse(result.Diagnostic.Contains(objectId.ToString("D")));
            Assert.IsFalse(result.Diagnostic.Contains(componentId.ToString("D")));
            var evidence = new MembershipCoverageDiagnosticsBuilder().Build(
                MembershipSnapshot.Complete(solution, new[] { result }, DateTimeOffset.UtcNow))
                .BroadRawComponentTypes.Single().Evidence.Single();
            Assert.AreEqual(objectId, evidence.ObjectId);
            Assert.AreEqual(componentId, evidence.SolutionComponentId);
            Assert.AreEqual(result.Diagnostic, evidence.Diagnostic);
        }

        [TestMethod]
        public void CancellationDuringComponentTypeChoiceDiagnosticPropagates()
        {
            var solution = Solution();
            using (var cancellation = new CancellationTokenSource())
            {
                var service = Service(solution, query => Rows());
                service.ExecuteRequest = request =>
                {
                    if (request is WhoAmIRequest) return WhoAmI(solution.Environment.OrganizationId);
                    if (request is RetrieveMetadataChangesRequest) return MetadataRows();
                    cancellation.Cancel();
                    return ComponentTypeChoices(new OptionMetadata(new Label("Team Template", 1033), 511));
                };

                Assert.ThrowsException<OperationCanceledException>(() =>
                    new DataverseComponentIdentityResolver().Resolve(service, solution.Environment,
                        Identity(null, 511, IdentityResolutionStatus.Unresolved).Record, cancellation.Token));
            }
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
                Identity(null, 511, IdentityResolutionStatus.Unresolved).Record, CancellationToken.None);

            Assert.AreEqual(IdentityResolutionStatus.Unsupported, result.Status);
            Assert.IsNull(result.SemanticKind);
            StringAssert.Contains(result.Diagnostic, "No entity metadata candidate");
            StringAssert.Contains(result.Diagnostic, "ObjectTypeCode 511");
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
                solution.Environment, Identity(null, 511, IdentityResolutionStatus.Unresolved).Record,
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

        [TestMethod]
        public void Type9OptionSetLookupCapturesCompleteEvidenceWithoutChangingIdentity()
        {
            var solution = Solution(); var objectId = Guid.NewGuid(); var componentId = Guid.NewGuid();
            var service = OptionSetService(solution, request =>
            {
                Assert.AreEqual(false, request.RetrieveAsIfPublished);
                return AllOptionSetsResponse(OptionSet(objectId, "new_Priority", true,
                    OptionSetType.Picklist, true, false));
            });

            var result = new DataverseComponentIdentityResolver().Resolve(service, solution.Environment,
                new SolutionComponentRecord(componentId, 9, objectId), CancellationToken.None);

            Assert.AreEqual(IdentityResolutionStatus.Unsupported, result.Status);
            Assert.AreEqual("unsupported:componenttype:9", result.SemanticKind);
            Assert.IsNull(result.ComparisonKey);
            Assert.AreEqual("No identity resolver supports this known component type.", result.Diagnostic);
            var evidence = result.DiagnosticEvidence.First();
            StringAssert.Contains(evidence, "matched uniquely");
            StringAssert.Contains(evidence, "MetadataId=" + objectId.ToString("D"));
            StringAssert.Contains(evidence, "Name='new_Priority'");
            StringAssert.Contains(evidence, "IsGlobal=True");
            StringAssert.Contains(evidence, "OptionSetType=0 ('Picklist')");
            StringAssert.Contains(evidence, "IsManaged=True");
            StringAssert.Contains(evidence, "IsCustomOptionSet=False");
            var summary = result.DiagnosticEvidence.Single(item =>
                item.StartsWith("OptionSet metadata diagnostic summary:", StringComparison.Ordinal));
            StringAssert.Contains(summary, "RawType9MembershipCount=1");
            StringAssert.Contains(summary, "DistinctNonemptyObjectIdCount=1");
            StringAssert.Contains(summary, "ReturnedOptionSetCount=1");
            StringAssert.Contains(summary, "CorrelatedMetadataIdCount=1");
            StringAssert.Contains(summary, "MissingRequestedMetadataIdCount=0");
            StringAssert.Contains(summary, "BlankNameCount=0");
            StringAssert.Contains(summary, "IsGlobalNotTrueCount=0");
            StringAssert.Contains(summary, "DistinctCandidateNames=['new_Priority']");

            var coverage = new MembershipCoverageDiagnosticsBuilder().Build(
                MembershipSnapshot.Complete(solution, new[] { result }, DateTimeOffset.UtcNow));
            var bucket = coverage.SemanticKinds.Single(item => item.SemanticKind ==
                "unsupported:componenttype:9");
            Assert.AreEqual(MembershipCoverageBucketType.KnownUnsupportedIsolatedType, bucket.BucketType);
            Assert.AreEqual(1, bucket.DiagnosticGroups.Count);
            Assert.AreEqual(result.Diagnostic, bucket.DiagnosticGroups.Single().Diagnostic);
            Assert.AreEqual(componentId, bucket.AuditEvidence.Single().SolutionComponentId);
            Assert.AreEqual(objectId, bucket.AuditEvidence.Single().ObjectId);
            CollectionAssert.AreEqual(result.DiagnosticEvidence.ToArray(),
                bucket.AuditEvidence.Single().DiagnosticEvidence.ToArray());
        }

        [TestMethod]
        public void Type9DistinctAndRepeatedIdsUseOneCatalogRequestAndIgnoreUnrelatedMetadata()
        {
            var solution = Solution();
            var firstId = Guid.NewGuid(); var secondId = Guid.NewGuid(); var unrelatedId = Guid.NewGuid();
            int lookupCount = 0;
            var service = OptionSetService(solution, request =>
            {
                lookupCount++;
                Assert.AreEqual(false, request.RetrieveAsIfPublished);
                return AllOptionSetsResponse(
                    OptionSet(firstId, "new_First", true, OptionSetType.Picklist, false, true),
                    OptionSet(secondId, "new_Second", true, OptionSetType.Picklist, false, true),
                    OptionSet(unrelatedId, "new_Unrelated", true, OptionSetType.Picklist, false, true));
            });
            var records = new[] { firstId, firstId, secondId }.Select(objectId =>
                new ComponentIdentity(new SolutionComponentRecord(Guid.NewGuid(), 9, objectId),
                    IdentityResolutionStatus.Unresolved)).ToArray();
            var counter = new D365SolutionComparer.Infrastructure.DataverseRequestCounter();

            var result = new DataverseComponentIdentityResolver().ResolveSnapshot(service,
                MembershipSnapshot.Complete(solution, records, DateTimeOffset.UtcNow),
                CancellationToken.None, counter);

            Assert.AreEqual(1, lookupCount);
            Assert.AreEqual(1, counter.GetExecuteCount("RetrieveAllOptionSets"));
            Assert.AreEqual(0, counter.GetExecuteCount("RetrieveOptionSet"));
            Assert.AreEqual(1, counter.GetExecuteCount("WhoAmI"));
            Assert.AreEqual(2, counter.TotalRequests);
            Assert.IsTrue(result.Components.All(item =>
                item.Status == IdentityResolutionStatus.Unsupported &&
                item.SemanticKind == "unsupported:componenttype:9" && item.ComparisonKey == null));
            Assert.AreEqual(result.Components[0].DiagnosticEvidence.First(),
                result.Components[1].DiagnosticEvidence.First());
            Assert.AreNotEqual(result.Components[0].DiagnosticEvidence.First(),
                result.Components[2].DiagnosticEvidence.First());
            Assert.IsFalse(result.Components.SelectMany(item => item.DiagnosticEvidence)
                .Any(item => item.Contains("new_Unrelated") && !item.StartsWith(
                    "OptionSet metadata diagnostic summary:", StringComparison.Ordinal)));
            var summary = result.Components.SelectMany(item => item.DiagnosticEvidence).Single(item =>
                item.StartsWith("OptionSet metadata diagnostic summary:", StringComparison.Ordinal));
            StringAssert.Contains(summary, "RawType9MembershipCount=3");
            StringAssert.Contains(summary, "DistinctNonemptyObjectIdCount=2");
            StringAssert.Contains(summary, "ReturnedOptionSetCount=3");
            StringAssert.Contains(summary, "CorrelatedMetadataIdCount=2");
            StringAssert.Contains(summary, "DistinctCandidateNames=['new_First', 'new_Second']");
            var coverage = new MembershipCoverageDiagnosticsBuilder().Build(result);
            var bucket = coverage.SemanticKinds.Single(item => item.SemanticKind ==
                "unsupported:componenttype:9");
            Assert.AreEqual(1, bucket.DiagnosticGroups.Count);
            Assert.AreEqual(3, bucket.DiagnosticGroups.Single().Count);
            Assert.AreEqual(3, bucket.AuditEvidence.Count);
            Assert.AreEqual(2, bucket.AuditEvidence.Select(item => item.ObjectId).Distinct().Count());
            Assert.AreEqual(3, bucket.AuditEvidence.SelectMany(item => item.DiagnosticEvidence).Distinct().Count());
        }

        [TestMethod]
        public void Type9EduScaleUsesOneCatalogRequestForThirtyEightDistinctIds()
        {
            var solution = Solution();
            var objectIds = Enumerable.Range(0, 38).Select(index => Guid.NewGuid()).ToList();
            var records = objectIds.Concat(new[] { objectIds[0], objectIds[1] }).Select(objectId =>
                new ComponentIdentity(new SolutionComponentRecord(Guid.NewGuid(), 9, objectId),
                    IdentityResolutionStatus.Unresolved)).ToArray();
            var service = OptionSetService(solution, request => AllOptionSetsResponse(objectIds.Select(
                (objectId, index) => OptionSet(objectId, "new_Choice" + index, true,
                    OptionSetType.Picklist, false, true)).ToArray()));
            var counter = new D365SolutionComparer.Infrastructure.DataverseRequestCounter();

            var result = new DataverseComponentIdentityResolver().ResolveSnapshot(service,
                MembershipSnapshot.Complete(solution, records, DateTimeOffset.UtcNow),
                CancellationToken.None, counter);

            Assert.AreEqual(1, counter.GetExecuteCount("RetrieveAllOptionSets"));
            Assert.AreEqual(2, counter.TotalRequests);
            var summary = result.Components.SelectMany(item => item.DiagnosticEvidence).Single(item =>
                item.StartsWith("OptionSet metadata diagnostic summary:", StringComparison.Ordinal));
            StringAssert.Contains(summary, "RawType9MembershipCount=40");
            StringAssert.Contains(summary, "DistinctNonemptyObjectIdCount=38");
            StringAssert.Contains(summary, "ReturnedOptionSetCount=38");
            StringAssert.Contains(summary, "CorrelatedMetadataIdCount=38");
            StringAssert.Contains(summary, "MissingRequestedMetadataIdCount=0");
            Assert.IsTrue(result.Components.All(item =>
                item.Status == IdentityResolutionStatus.Unsupported && item.ComparisonKey == null));
        }

        [TestMethod]
        public void Type9MissingAndQuestionableMetadataRemainExplicitDiagnosticEvidence()
        {
            var solution = Solution(); var missingId = Guid.NewGuid(); var blankId = Guid.NewGuid();
            var nonGlobalId = Guid.NewGuid(); var conflictingId = Guid.NewGuid(); var unrelatedId = Guid.NewGuid();
            var service = OptionSetService(solution, request =>
            {
                return AllOptionSetsResponse(
                    OptionSet(blankId, " ", true, OptionSetType.Picklist, false, true),
                    OptionSet(nonGlobalId, "new_Local", false, OptionSetType.Picklist, false, true),
                    OptionSet(conflictingId, "new_ConflictA", true, OptionSetType.Picklist, false, true),
                    OptionSet(conflictingId, "new_ConflictB", true, OptionSetType.Picklist, false, true),
                    OptionSet(unrelatedId, "new_Unrelated", true, OptionSetType.Picklist, false, true));
            });
            var records = new[] { missingId, blankId, nonGlobalId, conflictingId }.Select(objectId =>
                new ComponentIdentity(new SolutionComponentRecord(Guid.NewGuid(), 9, objectId),
                    IdentityResolutionStatus.Unresolved)).ToArray();

            var result = new DataverseComponentIdentityResolver().ResolveSnapshot(service,
                MembershipSnapshot.Complete(solution, records, DateTimeOffset.UtcNow), CancellationToken.None);

            StringAssert.Contains(result.Components[0].DiagnosticEvidence.First(), "No option-set metadata");
            StringAssert.Contains(result.Components[1].DiagnosticEvidence.First(), "Name is blank");
            StringAssert.Contains(result.Components[2].DiagnosticEvidence.First(), "IsGlobal is not true");
            StringAssert.Contains(result.Components[3].DiagnosticEvidence.First(),
                "Multiple option-set metadata definitions");
            var summary = result.Components.SelectMany(item => item.DiagnosticEvidence).Single(item =>
                item.StartsWith("OptionSet metadata diagnostic summary:", StringComparison.Ordinal));
            StringAssert.Contains(summary, "ReturnedOptionSetCount=5");
            StringAssert.Contains(summary, "CorrelatedMetadataIdCount=2");
            StringAssert.Contains(summary, "MissingRequestedMetadataIdCount=1");
            StringAssert.Contains(summary, "BlankNameCount=1");
            StringAssert.Contains(summary, "IsGlobalNotTrueCount=1");
            StringAssert.Contains(summary, "NonUniqueMetadataIdCount=1");
            Assert.IsTrue(result.Components.All(item =>
                item.Status == IdentityResolutionStatus.Unsupported && item.ComparisonKey == null));
        }

        [TestMethod]
        public void Type9FaultRemainsDiagnosticOnlyAndCancellationPropagates()
        {
            var solution = Solution(); var objectId = Guid.NewGuid();
            var faulted = new DataverseComponentIdentityResolver().Resolve(
                OptionSetService(solution, request => throw new FaultException("OptionSet metadata denied")),
                solution.Environment, new SolutionComponentRecord(Guid.NewGuid(), 9, objectId),
                CancellationToken.None);
            Assert.AreEqual(IdentityResolutionStatus.Unsupported, faulted.Status);
            Assert.AreEqual("unsupported:componenttype:9", faulted.SemanticKind);
            Assert.IsNull(faulted.ComparisonKey);
            StringAssert.Contains(faulted.DiagnosticEvidence.First(), "OptionSet metadata denied");
            StringAssert.Contains(faulted.DiagnosticEvidence.Single(item =>
                item.StartsWith("OptionSet metadata diagnostic summary:", StringComparison.Ordinal)),
                "ReturnedOptionSetCount=(unavailable)");

            using (var cancellation = new CancellationTokenSource())
            {
                var service = OptionSetService(solution, request =>
                {
                    cancellation.Cancel();
                    return AllOptionSetsResponse(OptionSet(objectId, "new_Priority", true,
                        OptionSetType.Picklist, false, true));
                });
                Assert.ThrowsException<OperationCanceledException>(() =>
                    new DataverseComponentIdentityResolver().Resolve(service, solution.Environment,
                        new SolutionComponentRecord(Guid.NewGuid(), 9, objectId), cancellation.Token));
            }
        }

        [TestMethod]
        public void Type9MissingObjectIdStillUsesOneCatalogRequestAndComparisonRemainsIndeterminate()
        {
            var solution = Solution(); int optionSetQueries = 0;
            var service = OptionSetService(solution, request =>
            {
                optionSetQueries++;
                return AllOptionSetsResponse();
            });
            var result = new DataverseComponentIdentityResolver().Resolve(service, solution.Environment,
                new SolutionComponentRecord(Guid.NewGuid(), 9, null), CancellationToken.None);

            Assert.AreEqual(1, optionSetQueries);
            Assert.AreEqual(IdentityResolutionStatus.Unsupported, result.Status);
            Assert.AreEqual("unsupported:componenttype:9", result.SemanticKind);
            StringAssert.Contains(result.DiagnosticEvidence.First(), "objectid is unavailable");
            var summary = result.DiagnosticEvidence.Single(item =>
                item.StartsWith("OptionSet metadata diagnostic summary:", StringComparison.Ordinal));
            StringAssert.Contains(summary, "RawType9MembershipCount=1");
            StringAssert.Contains(summary, "DistinctNonemptyObjectIdCount=0");
            StringAssert.Contains(summary, "MissingObjectIdRecordCount=1");
            var compared = new SolutionMembershipComparer().Compare(
                MembershipSnapshot.Complete(solution, new[] { result }, DateTimeOffset.UtcNow),
                MembershipSnapshot.Complete(solution, new ComponentIdentity[0], DateTimeOffset.UtcNow));
            Assert.AreEqual(MembershipPresence.Indeterminate, compared.Single().Presence);
        }

        [TestMethod]
        public void Type31ReportLookupCapturesSignedCandidateEvidenceWithoutCreatingIdentity()
        {
            var solution = Solution(); var objectId = Guid.NewGuid(); var componentId = Guid.NewGuid();
            var signatureId = Guid.NewGuid(); var reportIdUnique = Guid.NewGuid();
            var service = ReportService(solution, query =>
            {
                AssertReportQuery(query, objectId);
                return Rows(Report(objectId, "Account Summary", "AccountSummary.rdl", 1, signatureId, 1033,
                    reportIdUnique, false));
            });

            var result = new DataverseComponentIdentityResolver().Resolve(service, solution.Environment,
                new SolutionComponentRecord(componentId, 31, objectId), CancellationToken.None);

            Assert.AreEqual(IdentityResolutionStatus.Unsupported, result.Status);
            Assert.AreEqual("unsupported:componenttype:31", result.SemanticKind);
            Assert.IsNull(result.ComparisonKey);
            Assert.AreEqual("No identity resolver supports this known component type.", result.Diagnostic);
            var evidence = result.DiagnosticEvidence.First();
            StringAssert.Contains(evidence, "reportid=" + objectId.ToString("D"));
            StringAssert.Contains(evidence, "name='Account Summary'");
            StringAssert.Contains(evidence, "filename='AccountSummary.rdl'");
            StringAssert.Contains(evidence, "reporttypecode=1 ('Reporting Services Report')");
            StringAssert.Contains(evidence, "signatureid=" + signatureId.ToString("D"));
            StringAssert.Contains(evidence, "signaturelcid=1033");
            StringAssert.Contains(evidence, "reportidunique=" + reportIdUnique.ToString("D"));
            StringAssert.Contains(evidence, "componentstate=0 ('Published')");
            StringAssert.Contains(evidence, "ismanaged=False");
            StringAssert.Contains(evidence, "candidateSignatureId='" + signatureId.ToString("D") + "'");
            StringAssert.Contains(evidence, "signatureLcid=1033");
            var summary = result.DiagnosticEvidence.Single(item =>
                item.StartsWith("Signed Report diagnostic summary:", StringComparison.Ordinal));
            StringAssert.Contains(summary, "RawType31Count=1");
            StringAssert.Contains(summary, "DistinctObjectIdCount=1");
            StringAssert.Contains(summary, "ReturnedRowCount=1");
            StringAssert.Contains(summary, "CorrelatedCount=1");
            StringAssert.Contains(summary, "BlankSignatureIdCount=0");
            StringAssert.Contains(summary, "NonblankSignatureIdCount=1");
            StringAssert.Contains(summary, "DistinctSignatureIdCount=1");
            StringAssert.Contains(summary, "DuplicateSignatureIdCount=0");
        }

        [TestMethod]
        public void Type31SignedUnsignedAndDuplicateSignaturesAreSummarizedConservatively()
        {
            var solution = Solution(); var firstId = Guid.NewGuid(); var secondId = Guid.NewGuid();
            var unsignedId = Guid.NewGuid(); var signatureId = Guid.NewGuid();
            var service = ReportService(solution, query => Rows(
                Report(firstId, "Signed EN", "SignedEn.rdl", 1, signatureId, 1033, Guid.NewGuid(), false),
                Report(secondId, "Signed FR", "SignedFr.rdl", 1, signatureId, 1036, Guid.NewGuid(), true),
                Report(unsignedId, "Custom", "Custom.rdl", 1, null, null, Guid.NewGuid(), false)));
            var records = new[] { firstId, secondId, unsignedId }.Select(objectId =>
                new ComponentIdentity(new SolutionComponentRecord(Guid.NewGuid(), 31, objectId),
                    IdentityResolutionStatus.Unresolved)).ToArray();

            var result = new DataverseComponentIdentityResolver().ResolveSnapshot(service,
                MembershipSnapshot.Complete(solution, records, DateTimeOffset.UtcNow), CancellationToken.None);

            Assert.IsTrue(result.Components[0].DiagnosticEvidence.Any(item => item.Contains("occurs on 2")));
            Assert.IsTrue(result.Components[1].DiagnosticEvidence.Any(item => item.Contains("occurs on 2")));
            StringAssert.Contains(result.Components[0].DiagnosticEvidence.First(),
                "candidateSignatureId='" + signatureId.ToString("D") + "'; signatureLcid=1033");
            StringAssert.Contains(result.Components[1].DiagnosticEvidence.First(),
                "candidateSignatureId='" + signatureId.ToString("D") + "'; signatureLcid=1036");
            StringAssert.Contains(result.Components[2].DiagnosticEvidence.First(), "signatureid=(not supplied)");
            StringAssert.Contains(result.Components[2].DiagnosticEvidence.First(),
                "candidateSignatureId=(unavailable)");
            var summary = result.Components.SelectMany(item => item.DiagnosticEvidence).Single(item =>
                item.StartsWith("Signed Report diagnostic summary:", StringComparison.Ordinal));
            StringAssert.Contains(summary, "BlankSignatureIdCount=1");
            StringAssert.Contains(summary, "NonblankSignatureIdCount=2");
            StringAssert.Contains(summary, "DistinctSignatureIdCount=1");
            StringAssert.Contains(summary, "DuplicateSignatureIdCount=1");
            StringAssert.Contains(summary,
                "DistinctCandidateSignatureIds=['" + signatureId.ToString("D") + "']");
            Assert.IsFalse(summary.Contains("|lcid="));
            Assert.IsTrue(result.Components.All(item => item.Status == IdentityResolutionStatus.Unsupported &&
                item.ComparisonKey == null));
        }

        [TestMethod]
        public void Type31ReportLookupBatchesDeduplicatesAndKeepsStableGrouping()
        {
            var solution = Solution();
            var objectIds = Enumerable.Range(0, 201).Select(index => Guid.NewGuid()).ToList();
            var records = objectIds.Concat(new[] { objectIds[0] }).Select(objectId =>
                new ComponentIdentity(new SolutionComponentRecord(Guid.NewGuid(), 31, objectId),
                    IdentityResolutionStatus.Unresolved)).ToArray();
            var queriedIds = new List<Guid>();
            var service = ReportService(solution, query =>
            {
                var ids = query.Criteria.Conditions.Single().Values.Cast<Guid>().ToList();
                Assert.IsTrue(ids.Count <= 200);
                Assert.AreEqual(ids.Count, ids.Distinct().Count());
                queriedIds.AddRange(ids);
                return Rows(ids.Select((id, index) => Report(id, "Report " + index,
                    "Report" + index + ".rdl", 1, Guid.NewGuid(), 1033, Guid.NewGuid(), false)).ToArray());
            });
            var counter = new D365SolutionComparer.Infrastructure.DataverseRequestCounter();

            var result = new DataverseComponentIdentityResolver().ResolveSnapshot(service,
                MembershipSnapshot.Complete(solution, records, DateTimeOffset.UtcNow),
                CancellationToken.None, counter);

            CollectionAssert.AreEquivalent(objectIds, queriedIds);
            Assert.AreEqual(2, counter.GetQueryCount("report"));
            Assert.AreEqual(1, counter.GetExecuteCount("WhoAmI"));
            Assert.AreEqual(3, counter.TotalRequests);
            var bucket = new MembershipCoverageDiagnosticsBuilder().Build(result).SemanticKinds.Single(item =>
                item.SemanticKind == "unsupported:componenttype:31");
            Assert.AreEqual(MembershipCoverageBucketType.KnownUnsupportedIsolatedType, bucket.BucketType);
            Assert.AreEqual(1, bucket.DiagnosticGroups.Count);
            Assert.AreEqual(202, bucket.DiagnosticGroups.Single().Count);
            Assert.AreEqual(202, bucket.AuditEvidence.Count);
        }

        [TestMethod]
        public void Type31MissingDuplicateAndIncompleteRowsRemainDiagnosticOnly()
        {
            var solution = Solution(); var missingId = Guid.NewGuid(); var duplicateId = Guid.NewGuid();
            var incompleteId = Guid.NewGuid();
            var incomplete = new Entity("report", incompleteId)
            {
                ["reportid"] = incompleteId,
                ["name"] = "Partial",
                ["signatureid"] = Guid.NewGuid()
            };
            var service = ReportService(solution, query => Rows(
                Report(duplicateId, "First", "First.rdl", 1, Guid.NewGuid(), 1033,
                    Guid.NewGuid(), false),
                Report(duplicateId, "Second", "Second.rdl", 1, Guid.NewGuid(), 1033,
                    Guid.NewGuid(), true), incomplete));
            var records = new[] { missingId, duplicateId, incompleteId }.Select(objectId =>
                new ComponentIdentity(new SolutionComponentRecord(Guid.NewGuid(), 31, objectId),
                    IdentityResolutionStatus.Unresolved)).ToArray();

            var result = new DataverseComponentIdentityResolver().ResolveSnapshot(service,
                MembershipSnapshot.Complete(solution, records, DateTimeOffset.UtcNow), CancellationToken.None);

            StringAssert.Contains(result.Components[0].DiagnosticEvidence.First(), "No report row matched");
            StringAssert.Contains(result.Components[1].DiagnosticEvidence.First(), "Multiple report rows matched");
            Assert.IsTrue(result.Components[1].DiagnosticEvidence.Any(item => item.Contains("name='First'")));
            Assert.IsTrue(result.Components[1].DiagnosticEvidence.Any(item => item.Contains("name='Second'")));
            var incompleteEvidence = result.Components[2].DiagnosticEvidence.First();
            StringAssert.Contains(incompleteEvidence, "matched but returned incomplete data");
            StringAssert.Contains(incompleteEvidence, "filename=(not supplied)");
            StringAssert.Contains(incompleteEvidence, "signaturelcid=(not supplied)");
            var summary = result.Components.SelectMany(item => item.DiagnosticEvidence).Single(item =>
                item.StartsWith("Signed Report diagnostic summary:", StringComparison.Ordinal));
            StringAssert.Contains(summary, "MissingCount=1");
            StringAssert.Contains(summary, "NonUniqueObjectIdCount=1");
            Assert.IsTrue(result.Components.All(item => item.Status == IdentityResolutionStatus.Unsupported &&
                item.ComparisonKey == null));
        }

        [TestMethod]
        public void Type31ConflictingAndIncompleteResponsesRemainConservative()
        {
            var solution = Solution(); var conflictId = Guid.NewGuid();
            var conflicting = Report(conflictId, "Conflict", "Conflict.rdl", 1, Guid.NewGuid(), 1033,
                Guid.NewGuid(), false);
            conflicting.Id = Guid.NewGuid();
            var conflict = new DataverseComponentIdentityResolver().Resolve(
                ReportService(solution, query => Rows(conflicting)), solution.Environment,
                new SolutionComponentRecord(Guid.NewGuid(), 31, conflictId), CancellationToken.None);

            Assert.AreEqual(IdentityResolutionStatus.Unsupported, conflict.Status);
            Assert.IsNull(conflict.ComparisonKey);
            Assert.IsTrue(conflict.DiagnosticEvidence.Any(item => item.Contains("conflicting or incomplete")));
            Assert.IsTrue(conflict.DiagnosticEvidence.Any(item => item.Contains("name='Conflict'")));
            StringAssert.Contains(conflict.DiagnosticEvidence.Single(item =>
                item.StartsWith("Signed Report diagnostic summary:", StringComparison.Ordinal)),
                "ReturnedRowCount=(unavailable)");

            var incompleteId = Guid.NewGuid();
            var incomplete = new DataverseComponentIdentityResolver().Resolve(
                ReportService(solution, query =>
                {
                    var rows = Rows(Report(incompleteId, "Incomplete", "Incomplete.rdl", 1,
                        Guid.NewGuid(), 1033, Guid.NewGuid(), false));
                    rows.MoreRecords = true;
                    return rows;
                }), solution.Environment, new SolutionComponentRecord(Guid.NewGuid(), 31, incompleteId),
                CancellationToken.None);
            Assert.AreEqual(IdentityResolutionStatus.Unsupported, incomplete.Status);
            Assert.IsNull(incomplete.ComparisonKey);
            Assert.IsTrue(incomplete.DiagnosticEvidence.Any(item => item.Contains("incomplete result set")));
        }

        [TestMethod]
        public void Type31FaultIsDiagnosticAndCancellationPropagates()
        {
            var solution = Solution(); var objectId = Guid.NewGuid();
            var faulted = new DataverseComponentIdentityResolver().Resolve(
                ReportService(solution, query => throw new FaultException("Report denied")),
                solution.Environment, new SolutionComponentRecord(Guid.NewGuid(), 31, objectId),
                CancellationToken.None);
            Assert.AreEqual(IdentityResolutionStatus.Unsupported, faulted.Status);
            Assert.AreEqual("unsupported:componenttype:31", faulted.SemanticKind);
            Assert.IsNull(faulted.ComparisonKey);
            StringAssert.Contains(faulted.DiagnosticEvidence.First(), "Report denied");

            using (var cancellation = new CancellationTokenSource())
            {
                var service = ReportService(solution, query =>
                {
                    cancellation.Cancel();
                    return Rows(Report(objectId, "Report", "Report.rdl", 1, Guid.NewGuid(), 1033,
                        Guid.NewGuid(), false));
                });
                Assert.ThrowsException<OperationCanceledException>(() =>
                    new DataverseComponentIdentityResolver().Resolve(service, solution.Environment,
                        new SolutionComponentRecord(Guid.NewGuid(), 31, objectId), cancellation.Token));
            }
        }

        [TestMethod]
        public void Type31MatchingSignatureIdsIgnoreLcidAndCannotCreateMembershipMatches()
        {
            var sourceSolution = Solution();
            var targetSolution = new SolutionIdentity(new EnvironmentIdentity(Guid.NewGuid(), "Target"),
                Guid.NewGuid(), sourceSolution.UniqueName);
            var signatureId = Guid.NewGuid(); var resolver = new DataverseComponentIdentityResolver();
            var sourceId = Guid.NewGuid(); var targetId = Guid.NewGuid();
            var source = resolver.Resolve(ReportService(sourceSolution, query => Rows(
                    Report(sourceId, "DEV report", "Dev.rdl", 1, signatureId, 1033,
                        Guid.NewGuid(), false))), sourceSolution.Environment,
                new SolutionComponentRecord(Guid.NewGuid(), 31, sourceId), CancellationToken.None);
            var target = resolver.Resolve(ReportService(targetSolution, query => Rows(
                    Report(targetId, "UAT report", "Uat.rdl", 1, signatureId, 1036,
                        Guid.NewGuid(), true))), targetSolution.Environment,
                new SolutionComponentRecord(Guid.NewGuid(), 31, targetId), CancellationToken.None);

            var compared = new SolutionMembershipComparer().Compare(
                MembershipSnapshot.Complete(sourceSolution, new[] { source }, DateTimeOffset.UtcNow),
                MembershipSnapshot.Complete(targetSolution, new[] { target }, DateTimeOffset.UtcNow));

            Assert.AreEqual(2, compared.Count);
            Assert.IsTrue(compared.All(item => item.Presence == MembershipPresence.Indeterminate));
            Assert.IsNull(source.ComparisonKey);
            Assert.IsNull(target.ComparisonKey);
            StringAssert.Contains(source.DiagnosticEvidence.First(),
                "candidateSignatureId='" + signatureId.ToString("D") + "'; signatureLcid=1033");
            StringAssert.Contains(target.DiagnosticEvidence.First(),
                "candidateSignatureId='" + signatureId.ToString("D") + "'; signatureLcid=1036");
        }

        [TestMethod]
        public void Type31MissingObjectIdDoesNotQueryReport()
        {
            var solution = Solution(); int queryCount = 0;
            var result = new DataverseComponentIdentityResolver().Resolve(
                ReportService(solution, query => { queryCount++; return Rows(); }), solution.Environment,
                new SolutionComponentRecord(Guid.NewGuid(), 31, null), CancellationToken.None);

            Assert.AreEqual(0, queryCount);
            Assert.AreEqual(IdentityResolutionStatus.Unsupported, result.Status);
            Assert.AreEqual("unsupported:componenttype:31", result.SemanticKind);
            Assert.IsNull(result.ComparisonKey);
            StringAssert.Contains(result.DiagnosticEvidence.First(), "objectid is unavailable");
        }

        [TestMethod]
        public void Type60SystemFormLookupCapturesCandidateEvidenceWithoutCreatingIdentity()
        {
            var solution = Solution(); var objectId = Guid.NewGuid(); var componentId = Guid.NewGuid();
            int metadataRequests = 0;
            var service = SystemFormService(solution, query =>
            {
                AssertSystemFormQuery(query, objectId);
                return Rows(SystemForm(objectId, "new_AccountMain", "Account main", "account", 2,
                    Guid.NewGuid(), false));
            }, request => { metadataRequests++; return MetadataRows(); });

            var result = new DataverseComponentIdentityResolver().Resolve(service, solution.Environment,
                new SolutionComponentRecord(componentId, 60, objectId), CancellationToken.None);

            Assert.AreEqual(0, metadataRequests);
            Assert.AreEqual(IdentityResolutionStatus.Unsupported, result.Status);
            Assert.AreEqual("unsupported:componenttype:60", result.SemanticKind);
            Assert.IsNull(result.ComparisonKey);
            Assert.AreEqual("No identity resolver supports this known component type.", result.Diagnostic);
            var evidence = result.DiagnosticEvidence.First();
            StringAssert.Contains(evidence, "formid=" + objectId.ToString("D"));
            StringAssert.Contains(evidence, "uniquename='new_AccountMain'");
            StringAssert.Contains(evidence, "name='Account main'");
            StringAssert.Contains(evidence, "objecttypecode='account'");
            StringAssert.Contains(evidence, "entitylogicalname=account");
            StringAssert.Contains(evidence, "type=2 ('Main')");
            StringAssert.Contains(evidence, "formidunique=");
            StringAssert.Contains(evidence, "componentstate=0 ('Published')");
            StringAssert.Contains(evidence, "ismanaged=False");
            StringAssert.Contains(evidence, "candidateportableidentity='account.new_AccountMain'");
            var summary = result.DiagnosticEvidence.Single(item =>
                item.StartsWith("System Form diagnostic summary:", StringComparison.Ordinal));
            StringAssert.Contains(summary, "RawType60MembershipCount=1");
            StringAssert.Contains(summary, "UniqueObjectIdCorrelationCount=1");
            StringAssert.Contains(summary,
                "DistinctCandidatePortableIdentities=['account.new_AccountMain']");
        }

        [TestMethod]
        public void Type60NumericObjectTypeCodesUseOneGroupedMetadataLookup()
        {
            var solution = Solution(); var firstId = Guid.NewGuid(); var secondId = Guid.NewGuid();
            int metadataRequests = 0;
            var service = SystemFormService(solution, query => Rows(
                SystemForm(firstId, "new_First", "First", 1, 2, Guid.NewGuid(), false),
                SystemForm(secondId, "new_Second", "Second", 1, 7, Guid.NewGuid(), false)), request =>
            {
                metadataRequests++;
                CollectionAssert.AreEquivalent(new[] { 1 }, AssertEntityLogicalNameMetadataQuery(request));
                return MetadataRows(EntityMetadata(1, "account", "Account"));
            });
            var records = new[] { firstId, secondId }.Select(objectId =>
                new ComponentIdentity(new SolutionComponentRecord(Guid.NewGuid(), 60, objectId),
                    IdentityResolutionStatus.Unresolved)).ToArray();
            var counter = new D365SolutionComparer.Infrastructure.DataverseRequestCounter();

            var result = new DataverseComponentIdentityResolver().ResolveSnapshot(service,
                MembershipSnapshot.Complete(solution, records, DateTimeOffset.UtcNow),
                CancellationToken.None, counter);

            Assert.AreEqual(1, metadataRequests);
            Assert.AreEqual(1, counter.GetQueryCount("systemform"));
            Assert.AreEqual(1, counter.GetExecuteCount("RetrieveMetadataChanges"));
            Assert.AreEqual(1, counter.GetExecuteCount("WhoAmI"));
            Assert.AreEqual(3, counter.TotalRequests);
            Assert.IsTrue(result.Components.All(item => item.DiagnosticEvidence.First()
                .Contains("entitylogicalname=account")));
            Assert.IsTrue(result.Components.All(item => item.Status == IdentityResolutionStatus.Unsupported &&
                item.ComparisonKey == null));
        }

        [TestMethod]
        public void Type60SystemFormLookupBatchesDeduplicatesAndKeepsStableGrouping()
        {
            var solution = Solution();
            var objectIds = Enumerable.Range(0, 201).Select(index => Guid.NewGuid()).ToList();
            var records = objectIds.Concat(new[] { objectIds[0] }).Select(objectId =>
                new ComponentIdentity(new SolutionComponentRecord(Guid.NewGuid(), 60, objectId),
                    IdentityResolutionStatus.Unresolved)).ToArray();
            var queriedIds = new List<Guid>();
            var service = SystemFormService(solution, query =>
            {
                var ids = query.Criteria.Conditions.Single().Values.Cast<Guid>().ToList();
                Assert.IsTrue(ids.Count <= 200);
                Assert.AreEqual(ids.Count, ids.Distinct().Count());
                queriedIds.AddRange(ids);
                return Rows(ids.Select((id, index) => SystemForm(id, "new_Form" + index,
                    "Form " + index, "account", 2, Guid.NewGuid(), false)).ToArray());
            });
            var counter = new D365SolutionComparer.Infrastructure.DataverseRequestCounter();

            var result = new DataverseComponentIdentityResolver().ResolveSnapshot(service,
                MembershipSnapshot.Complete(solution, records, DateTimeOffset.UtcNow),
                CancellationToken.None, counter);

            CollectionAssert.AreEquivalent(objectIds, queriedIds);
            Assert.AreEqual(2, counter.GetQueryCount("systemform"));
            Assert.AreEqual(1, counter.GetExecuteCount("WhoAmI"));
            Assert.AreEqual(3, counter.TotalRequests);
            var bucket = new MembershipCoverageDiagnosticsBuilder().Build(result).SemanticKinds.Single(item =>
                item.SemanticKind == "unsupported:componenttype:60");
            Assert.AreEqual(MembershipCoverageBucketType.KnownUnsupportedIsolatedType, bucket.BucketType);
            Assert.AreEqual(1, bucket.DiagnosticGroups.Count);
            Assert.AreEqual(202, bucket.DiagnosticGroups.Single().Count);
            Assert.AreEqual(202, bucket.AuditEvidence.Count);
        }

        [TestMethod]
        public void Type60ZeroDuplicateAndBlankIdentityInputsRemainDiagnosticOnly()
        {
            var solution = Solution(); var missingId = Guid.NewGuid(); var duplicateId = Guid.NewGuid();
            var blankId = Guid.NewGuid();
            var service = SystemFormService(solution, query => Rows(
                SystemForm(duplicateId, "new_First", "First", "account", 2, Guid.NewGuid(), false),
                SystemForm(duplicateId, "new_Second", "Second", "account", 2, Guid.NewGuid(), true),
                SystemForm(blankId, " ", "Blank unique name", " ", 2, Guid.NewGuid(), false)));
            var records = new[] { missingId, duplicateId, blankId }.Select(objectId =>
                new ComponentIdentity(new SolutionComponentRecord(Guid.NewGuid(), 60, objectId),
                    IdentityResolutionStatus.Unresolved)).ToArray();

            var result = new DataverseComponentIdentityResolver().ResolveSnapshot(service,
                MembershipSnapshot.Complete(solution, records, DateTimeOffset.UtcNow), CancellationToken.None);

            StringAssert.Contains(result.Components[0].DiagnosticEvidence.First(), "No systemform row matched");
            StringAssert.Contains(result.Components[1].DiagnosticEvidence.First(), "Multiple systemform rows matched");
            Assert.IsTrue(result.Components[1].DiagnosticEvidence.Any(item => item.Contains("name='First'")));
            Assert.IsTrue(result.Components[1].DiagnosticEvidence.Any(item => item.Contains("name='Second'")));
            StringAssert.Contains(result.Components[2].DiagnosticEvidence.First(), "uniquename=' '");
            StringAssert.Contains(result.Components[2].DiagnosticEvidence.First(),
                "entity logical name is blank");
            StringAssert.Contains(result.Components[2].DiagnosticEvidence.First(),
                "candidateportableidentity=(unavailable)");
            var summary = result.Components.SelectMany(item => item.DiagnosticEvidence).Single(item =>
                item.StartsWith("System Form diagnostic summary:", StringComparison.Ordinal));
            StringAssert.Contains(summary, "MissingRequestedObjectIdCount=1");
            StringAssert.Contains(summary, "BlankUniqueNameCount=1");
            StringAssert.Contains(summary, "UnresolvedEntityLogicalNameCount=1");
            StringAssert.Contains(summary, "NonUniqueObjectIdCount=1");
            Assert.IsTrue(result.Components.All(item =>
                item.Status == IdentityResolutionStatus.Unsupported && item.ComparisonKey == null));
        }

        [TestMethod]
        public void Type60ConflictingResultPreservesEvidenceAndRemainsUnsupported()
        {
            var solution = Solution(); var objectId = Guid.NewGuid();
            var conflicting = SystemForm(objectId, "new_Conflict", "Conflict", "account", 2,
                Guid.NewGuid(), false);
            conflicting.Id = Guid.NewGuid();

            var result = new DataverseComponentIdentityResolver().Resolve(
                SystemFormService(solution, query => Rows(conflicting)), solution.Environment,
                new SolutionComponentRecord(Guid.NewGuid(), 60, objectId), CancellationToken.None);

            Assert.AreEqual(IdentityResolutionStatus.Unsupported, result.Status);
            Assert.AreEqual("unsupported:componenttype:60", result.SemanticKind);
            Assert.IsNull(result.ComparisonKey);
            Assert.IsTrue(result.DiagnosticEvidence.Any(item => item.Contains("conflicting or incomplete")));
            Assert.IsTrue(result.DiagnosticEvidence.Any(item => item.Contains("name='Conflict'")));
            StringAssert.Contains(result.DiagnosticEvidence.Single(item =>
                item.StartsWith("System Form diagnostic summary:", StringComparison.Ordinal)),
                "ReturnedSystemFormRowCount=(unavailable)");
        }

        [TestMethod]
        public void Type60MetadataFaultIsConservativeAndQueryCancellationPropagates()
        {
            var solution = Solution(); var objectId = Guid.NewGuid();
            var metadataFault = new DataverseComponentIdentityResolver().Resolve(
                SystemFormService(solution, query => Rows(SystemForm(objectId, "new_Form", "Form", 1, 2,
                    Guid.NewGuid(), false)), request => throw new FaultException("Entity metadata denied")),
                solution.Environment, new SolutionComponentRecord(Guid.NewGuid(), 60, objectId),
                CancellationToken.None);
            Assert.AreEqual(IdentityResolutionStatus.Unsupported, metadataFault.Status);
            Assert.AreEqual("unsupported:componenttype:60", metadataFault.SemanticKind);
            Assert.IsNull(metadataFault.ComparisonKey);
            StringAssert.Contains(metadataFault.DiagnosticEvidence.First(), "Entity metadata denied");
            StringAssert.Contains(metadataFault.DiagnosticEvidence.First(),
                "candidateportableidentity=(unavailable)");

            using (var cancellation = new CancellationTokenSource())
            {
                var service = SystemFormService(solution, query =>
                {
                    cancellation.Cancel();
                    return Rows(SystemForm(objectId, "new_Form", "Form", "account", 2,
                        Guid.NewGuid(), false));
                });
                Assert.ThrowsException<OperationCanceledException>(() =>
                    new DataverseComponentIdentityResolver().Resolve(service, solution.Environment,
                        new SolutionComponentRecord(Guid.NewGuid(), 60, objectId), cancellation.Token));
            }
        }

        [TestMethod]
        public void Type60QueryFaultRemainsDiagnosticAndCandidateNamesCannotCreateMatches()
        {
            var sourceSolution = Solution();
            var targetSolution = new SolutionIdentity(new EnvironmentIdentity(Guid.NewGuid(), "Target"),
                Guid.NewGuid(), sourceSolution.UniqueName);
            var sourceId = Guid.NewGuid(); var targetId = Guid.NewGuid();
            var resolver = new DataverseComponentIdentityResolver();
            var source = resolver.Resolve(SystemFormService(sourceSolution, query => Rows(
                    SystemForm(sourceId, "new_Main", "DEV", "account", 2, Guid.NewGuid(), false))),
                sourceSolution.Environment, new SolutionComponentRecord(Guid.NewGuid(), 60, sourceId),
                CancellationToken.None);
            var target = resolver.Resolve(SystemFormService(targetSolution, query => Rows(
                    SystemForm(targetId, "NEW_MAIN", "UAT", "ACCOUNT", 2, Guid.NewGuid(), true))),
                targetSolution.Environment, new SolutionComponentRecord(Guid.NewGuid(), 60, targetId),
                CancellationToken.None);
            var compared = new SolutionMembershipComparer().Compare(
                MembershipSnapshot.Complete(sourceSolution, new[] { source }, DateTimeOffset.UtcNow),
                MembershipSnapshot.Complete(targetSolution, new[] { target }, DateTimeOffset.UtcNow));
            Assert.AreEqual(2, compared.Count);
            Assert.IsTrue(compared.All(item => item.Presence == MembershipPresence.Indeterminate));
            Assert.IsNull(source.ComparisonKey);
            Assert.IsNull(target.ComparisonKey);

            var faulted = resolver.Resolve(SystemFormService(sourceSolution,
                    query => throw new FaultException("System Form denied")), sourceSolution.Environment,
                new SolutionComponentRecord(Guid.NewGuid(), 60, Guid.NewGuid()), CancellationToken.None);
            Assert.AreEqual(IdentityResolutionStatus.Unsupported, faulted.Status);
            Assert.IsNull(faulted.ComparisonKey);
            StringAssert.Contains(faulted.DiagnosticEvidence.First(), "System Form denied");
        }

        [TestMethod]
        public void Type60MissingObjectIdDoesNotQuerySystemForm()
        {
            var solution = Solution(); int queryCount = 0;
            var result = new DataverseComponentIdentityResolver().Resolve(
                SystemFormService(solution, query => { queryCount++; return Rows(); }), solution.Environment,
                new SolutionComponentRecord(Guid.NewGuid(), 60, null), CancellationToken.None);

            Assert.AreEqual(0, queryCount);
            Assert.AreEqual(IdentityResolutionStatus.Unsupported, result.Status);
            Assert.AreEqual("unsupported:componenttype:60", result.SemanticKind);
            Assert.IsNull(result.ComparisonKey);
            StringAssert.Contains(result.DiagnosticEvidence.First(), "objectid is unavailable");
        }

        [TestMethod]
        public void Type62SiteMapLookupCapturesCandidateEvidenceWithoutCreatingIdentity()
        {
            var solution = Solution(); var objectId = Guid.NewGuid(); var componentId = Guid.NewGuid();
            var siteMapIdUnique = Guid.NewGuid();
            var service = SiteMapService(solution, query =>
            {
                AssertSiteMapQuery(query, objectId);
                return Rows(SiteMap(objectId, "new_EduNavigation", "EDU navigation", siteMapIdUnique,
                    true, false));
            });

            var result = new DataverseComponentIdentityResolver().Resolve(service, solution.Environment,
                new SolutionComponentRecord(componentId, 62, objectId), CancellationToken.None);

            Assert.AreEqual(IdentityResolutionStatus.Unsupported, result.Status);
            Assert.AreEqual("unsupported:componenttype:62", result.SemanticKind);
            Assert.IsNull(result.ComparisonKey);
            Assert.AreEqual("No identity resolver supports this known component type.", result.Diagnostic);
            var evidence = result.DiagnosticEvidence.First();
            StringAssert.Contains(evidence, "sitemapid=" + objectId.ToString("D"));
            StringAssert.Contains(evidence, "sitemapnameunique='new_EduNavigation'");
            StringAssert.Contains(evidence, "sitemapname='EDU navigation'");
            StringAssert.Contains(evidence, "sitemapidunique=" + siteMapIdUnique.ToString("D"));
            StringAssert.Contains(evidence, "isappaware=True");
            StringAssert.Contains(evidence, "componentstate=0 ('Published')");
            StringAssert.Contains(evidence, "ismanaged=False");
            StringAssert.Contains(evidence, "candidatesitemapname='new_EduNavigation'");
            var summary = result.DiagnosticEvidence.Single(item =>
                item.StartsWith("Site Map diagnostic summary:", StringComparison.Ordinal));
            StringAssert.Contains(summary, "RawType62MembershipCount=1");
            StringAssert.Contains(summary, "DistinctNonemptyObjectIdCount=1");
            StringAssert.Contains(summary, "ReturnedSiteMapRowCount=1");
            StringAssert.Contains(summary, "UniqueObjectIdCorrelationCount=1");
            StringAssert.Contains(summary, "AppAwareCount=1");
            StringAssert.Contains(summary, "DistinctCandidateSiteMapNames=['new_EduNavigation']");
        }

        [TestMethod]
        public void Type62SiteMapLookupBatchesDeduplicatesAndKeepsStableGrouping()
        {
            var solution = Solution();
            var objectIds = Enumerable.Range(0, 201).Select(index => Guid.NewGuid()).ToList();
            var records = objectIds.Concat(new[] { objectIds[0] }).Select(objectId =>
                new ComponentIdentity(new SolutionComponentRecord(Guid.NewGuid(), 62, objectId),
                    IdentityResolutionStatus.Unresolved)).ToArray();
            var queriedIds = new List<Guid>();
            var service = SiteMapService(solution, query =>
            {
                var ids = query.Criteria.Conditions.Single().Values.Cast<Guid>().ToList();
                Assert.IsTrue(ids.Count <= 200);
                Assert.AreEqual(ids.Count, ids.Distinct().Count());
                queriedIds.AddRange(ids);
                return Rows(ids.Select((id, index) => SiteMap(id, "new_SiteMap" + index,
                    "Site map " + index, Guid.NewGuid(), true, false)).ToArray());
            });
            var counter = new D365SolutionComparer.Infrastructure.DataverseRequestCounter();

            var result = new DataverseComponentIdentityResolver().ResolveSnapshot(service,
                MembershipSnapshot.Complete(solution, records, DateTimeOffset.UtcNow),
                CancellationToken.None, counter);

            CollectionAssert.AreEquivalent(objectIds, queriedIds);
            Assert.AreEqual(2, counter.GetQueryCount("sitemap"));
            Assert.AreEqual(1, counter.GetExecuteCount("WhoAmI"));
            Assert.AreEqual(3, counter.TotalRequests);
            var bucket = new MembershipCoverageDiagnosticsBuilder().Build(result).SemanticKinds.Single(item =>
                item.SemanticKind == "unsupported:componenttype:62");
            Assert.AreEqual(MembershipCoverageBucketType.KnownUnsupportedIsolatedType, bucket.BucketType);
            Assert.AreEqual(1, bucket.DiagnosticGroups.Count);
            Assert.AreEqual(202, bucket.DiagnosticGroups.Single().Count);
            Assert.AreEqual(202, bucket.AuditEvidence.Count);
        }

        [TestMethod]
        public void Type62MissingDuplicateAndBlankCandidatesRemainDiagnosticOnly()
        {
            var solution = Solution(); var missingId = Guid.NewGuid(); var duplicateId = Guid.NewGuid();
            var blankId = Guid.NewGuid();
            var service = SiteMapService(solution, query => Rows(
                SiteMap(duplicateId, "new_First", "First", Guid.NewGuid(), true, false),
                SiteMap(duplicateId, "new_Second", "Second", Guid.NewGuid(), true, true),
                SiteMap(blankId, " ", "Blank unique name", Guid.NewGuid(), false, false)));
            var records = new[] { missingId, duplicateId, blankId }.Select(objectId =>
                new ComponentIdentity(new SolutionComponentRecord(Guid.NewGuid(), 62, objectId),
                    IdentityResolutionStatus.Unresolved)).ToArray();

            var result = new DataverseComponentIdentityResolver().ResolveSnapshot(service,
                MembershipSnapshot.Complete(solution, records, DateTimeOffset.UtcNow), CancellationToken.None);

            StringAssert.Contains(result.Components[0].DiagnosticEvidence.First(), "No sitemap row matched");
            StringAssert.Contains(result.Components[1].DiagnosticEvidence.First(), "Multiple sitemap rows matched");
            Assert.IsTrue(result.Components[1].DiagnosticEvidence.Any(item => item.Contains("sitemapname='First'")));
            Assert.IsTrue(result.Components[1].DiagnosticEvidence.Any(item => item.Contains("sitemapname='Second'")));
            StringAssert.Contains(result.Components[2].DiagnosticEvidence.First(), "sitemapnameunique=' '");
            StringAssert.Contains(result.Components[2].DiagnosticEvidence.First(),
                "candidatesitemapname=(unavailable)");
            var summary = result.Components.SelectMany(item => item.DiagnosticEvidence).Single(item =>
                item.StartsWith("Site Map diagnostic summary:", StringComparison.Ordinal));
            StringAssert.Contains(summary, "MissingRequestedObjectIdCount=1");
            StringAssert.Contains(summary, "BlankSiteMapNameUniqueCount=1");
            StringAssert.Contains(summary, "NonUniqueObjectIdCount=1");
            Assert.IsTrue(result.Components.All(item => item.Status == IdentityResolutionStatus.Unsupported &&
                item.ComparisonKey == null));
        }

        [TestMethod]
        public void Type62ConflictingAndIncompleteResponsesRemainConservative()
        {
            var solution = Solution(); var conflictId = Guid.NewGuid();
            var conflicting = SiteMap(conflictId, "new_Conflict", "Conflict", Guid.NewGuid(), true, false);
            conflicting.Id = Guid.NewGuid();
            var conflict = new DataverseComponentIdentityResolver().Resolve(
                SiteMapService(solution, query => Rows(conflicting)), solution.Environment,
                new SolutionComponentRecord(Guid.NewGuid(), 62, conflictId), CancellationToken.None);

            Assert.AreEqual(IdentityResolutionStatus.Unsupported, conflict.Status);
            Assert.IsNull(conflict.ComparisonKey);
            Assert.IsTrue(conflict.DiagnosticEvidence.Any(item => item.Contains("conflicting or incomplete")));
            Assert.IsTrue(conflict.DiagnosticEvidence.Any(item => item.Contains("sitemapname='Conflict'")));
            StringAssert.Contains(conflict.DiagnosticEvidence.Single(item =>
                item.StartsWith("Site Map diagnostic summary:", StringComparison.Ordinal)),
                "ReturnedSiteMapRowCount=(unavailable)");

            var incompleteId = Guid.NewGuid();
            var incomplete = new DataverseComponentIdentityResolver().Resolve(
                SiteMapService(solution, query =>
                {
                    var rows = Rows(SiteMap(incompleteId, "new_Incomplete", "Incomplete", Guid.NewGuid(),
                        true, false));
                    rows.MoreRecords = true;
                    return rows;
                }), solution.Environment, new SolutionComponentRecord(Guid.NewGuid(), 62, incompleteId),
                CancellationToken.None);
            Assert.AreEqual(IdentityResolutionStatus.Unsupported, incomplete.Status);
            Assert.IsNull(incomplete.ComparisonKey);
            Assert.IsTrue(incomplete.DiagnosticEvidence.Any(item => item.Contains("incomplete result set")));
        }

        [TestMethod]
        public void Type62IncompleteRowPreservesEveryAvailableFieldWithoutCreatingIdentity()
        {
            var solution = Solution(); var objectId = Guid.NewGuid();
            var row = new Entity("sitemap", objectId)
            {
                ["sitemapid"] = objectId,
                ["sitemapnameunique"] = "new_Partial",
                ["isappaware"] = true
            };

            var result = new DataverseComponentIdentityResolver().Resolve(
                SiteMapService(solution, query => Rows(row)), solution.Environment,
                new SolutionComponentRecord(Guid.NewGuid(), 62, objectId), CancellationToken.None);

            Assert.AreEqual(IdentityResolutionStatus.Unsupported, result.Status);
            Assert.IsNull(result.ComparisonKey);
            var evidence = result.DiagnosticEvidence.First();
            StringAssert.Contains(evidence, "matched but returned incomplete data");
            StringAssert.Contains(evidence, "sitemapnameunique='new_Partial'");
            StringAssert.Contains(evidence, "sitemapname=(not supplied)");
            StringAssert.Contains(evidence, "sitemapidunique=(not supplied)");
            StringAssert.Contains(evidence, "componentstate=(not supplied)");
            StringAssert.Contains(evidence, "ismanaged=(not supplied)");
        }

        [TestMethod]
        public void Type62FaultIsDiagnosticAndCancellationPropagates()
        {
            var solution = Solution(); var objectId = Guid.NewGuid();
            var faulted = new DataverseComponentIdentityResolver().Resolve(
                SiteMapService(solution, query => throw new FaultException("Site Map denied")),
                solution.Environment, new SolutionComponentRecord(Guid.NewGuid(), 62, objectId),
                CancellationToken.None);
            Assert.AreEqual(IdentityResolutionStatus.Unsupported, faulted.Status);
            Assert.AreEqual("unsupported:componenttype:62", faulted.SemanticKind);
            Assert.IsNull(faulted.ComparisonKey);
            StringAssert.Contains(faulted.DiagnosticEvidence.First(), "Site Map denied");

            using (var cancellation = new CancellationTokenSource())
            {
                var service = SiteMapService(solution, query =>
                {
                    cancellation.Cancel();
                    return Rows(SiteMap(objectId, "new_SiteMap", "Site Map", Guid.NewGuid(), true, false));
                });
                Assert.ThrowsException<OperationCanceledException>(() =>
                    new DataverseComponentIdentityResolver().Resolve(service, solution.Environment,
                        new SolutionComponentRecord(Guid.NewGuid(), 62, objectId), cancellation.Token));
            }
        }

        [TestMethod]
        public void Type62CandidateNamesCannotCreateMembershipMatches()
        {
            var sourceSolution = Solution();
            var targetSolution = new SolutionIdentity(new EnvironmentIdentity(Guid.NewGuid(), "Target"),
                Guid.NewGuid(), sourceSolution.UniqueName);
            var resolver = new DataverseComponentIdentityResolver();
            var sourceId = Guid.NewGuid(); var targetId = Guid.NewGuid();
            var source = resolver.Resolve(SiteMapService(sourceSolution, query => Rows(
                    SiteMap(sourceId, "new_EduNavigation", "DEV", Guid.NewGuid(), true, false))),
                sourceSolution.Environment, new SolutionComponentRecord(Guid.NewGuid(), 62, sourceId),
                CancellationToken.None);
            var target = resolver.Resolve(SiteMapService(targetSolution, query => Rows(
                    SiteMap(targetId, "NEW_EDUNAVIGATION", "UAT", Guid.NewGuid(), true, true))),
                targetSolution.Environment, new SolutionComponentRecord(Guid.NewGuid(), 62, targetId),
                CancellationToken.None);

            var compared = new SolutionMembershipComparer().Compare(
                MembershipSnapshot.Complete(sourceSolution, new[] { source }, DateTimeOffset.UtcNow),
                MembershipSnapshot.Complete(targetSolution, new[] { target }, DateTimeOffset.UtcNow));

            Assert.AreEqual(2, compared.Count);
            Assert.IsTrue(compared.All(item => item.Presence == MembershipPresence.Indeterminate));
            Assert.IsNull(source.ComparisonKey);
            Assert.IsNull(target.ComparisonKey);
        }

        [TestMethod]
        public void Type62MissingObjectIdDoesNotQuerySiteMap()
        {
            var solution = Solution(); int queryCount = 0;
            var result = new DataverseComponentIdentityResolver().Resolve(
                SiteMapService(solution, query => { queryCount++; return Rows(); }), solution.Environment,
                new SolutionComponentRecord(Guid.NewGuid(), 62, null), CancellationToken.None);

            Assert.AreEqual(0, queryCount);
            Assert.AreEqual(IdentityResolutionStatus.Unsupported, result.Status);
            Assert.AreEqual("unsupported:componenttype:62", result.SemanticKind);
            Assert.IsNull(result.ComparisonKey);
            StringAssert.Contains(result.DiagnosticEvidence.First(), "objectid is unavailable");
        }

        [TestMethod]
        public void Type300CanvasAppLookupCorrelatesByIdAndPreservesEveryRequestedField()
        {
            var solution = Solution(); var objectId = Guid.NewGuid(); var componentId = Guid.NewGuid();
            var service = CanvasAppService(solution, query =>
            {
                AssertCanvasAppQuery(query, objectId);
                return Rows(CanvasApp(objectId, "new_InspectionApp", "Inspection App",
                    "a52d4d32-54af-42f6-a4cd-6d494b072cb1", false));
            });

            var result = new DataverseComponentIdentityResolver().Resolve(service, solution.Environment,
                new SolutionComponentRecord(componentId, 300, objectId), CancellationToken.None);

            Assert.AreEqual(IdentityResolutionStatus.Unsupported, result.Status);
            Assert.AreEqual("unsupported:componenttype:300", result.SemanticKind);
            Assert.IsNull(result.ComparisonKey);
            Assert.AreEqual("No identity resolver supports this known component type.", result.Diagnostic);
            var evidence = result.DiagnosticEvidence.First();
            StringAssert.Contains(evidence, "canvasappid=" + objectId.ToString("D"));
            StringAssert.Contains(evidence, "name='new_InspectionApp'");
            StringAssert.Contains(evidence, "displayname='Inspection App'");
            StringAssert.Contains(evidence, "uniquecanvasappid='a52d4d32-54af-42f6-a4cd-6d494b072cb1'");
            StringAssert.Contains(evidence, "componentstate=0 ('Published')");
            StringAssert.Contains(evidence, "ismanaged=False");
            var summary = result.DiagnosticEvidence.Single(item =>
                item.StartsWith("Canvas App diagnostic summary:", StringComparison.Ordinal));
            StringAssert.Contains(summary, "RawType300MembershipCount=1");
            StringAssert.Contains(summary, "DistinctNonemptyObjectIdCount=1");
            StringAssert.Contains(summary, "ReturnedCanvasAppRowCount=1");
            StringAssert.Contains(summary, "UniqueObjectIdCorrelationCount=1");
            StringAssert.Contains(summary, "BlankNameCount=0");
            StringAssert.Contains(summary, "DistinctCandidateNames=['new_InspectionApp']");
        }

        [TestMethod]
        public void Type300CanvasAppLookupBatchesDeduplicatesAndCountsRequests()
        {
            var solution = Solution();
            var objectIds = Enumerable.Range(0, 201).Select(index => Guid.NewGuid()).ToList();
            var records = objectIds.Concat(new[] { objectIds[0] }).Select(objectId =>
                new ComponentIdentity(new SolutionComponentRecord(Guid.NewGuid(), 300, objectId),
                    IdentityResolutionStatus.Unresolved)).ToArray();
            var queriedIds = new List<Guid>();
            var service = CanvasAppService(solution, query =>
            {
                var ids = query.Criteria.Conditions.Single().Values.Cast<Guid>().ToList();
                Assert.IsTrue(ids.Count <= 200);
                Assert.AreEqual(ids.Count, ids.Distinct().Count());
                queriedIds.AddRange(ids);
                return Rows(ids.Select((id, index) => CanvasApp(id, "new_App" + index,
                    "App " + index, Guid.NewGuid().ToString("D"), false)).ToArray());
            });
            var counter = new D365SolutionComparer.Infrastructure.DataverseRequestCounter();

            var result = new DataverseComponentIdentityResolver().ResolveSnapshot(service,
                MembershipSnapshot.Complete(solution, records, DateTimeOffset.UtcNow),
                CancellationToken.None, counter);

            CollectionAssert.AreEquivalent(objectIds, queriedIds);
            Assert.AreEqual(2, counter.GetQueryCount("canvasapp"));
            Assert.AreEqual(1, counter.GetExecuteCount("WhoAmI"));
            Assert.AreEqual(3, counter.TotalRequests);
            Assert.IsTrue(result.Components.All(item =>
                item.Status == IdentityResolutionStatus.Unsupported &&
                item.SemanticKind == "unsupported:componenttype:300" && item.ComparisonKey == null));
            var summary = result.Components.SelectMany(item => item.DiagnosticEvidence).Single(item =>
                item.StartsWith("Canvas App diagnostic summary:", StringComparison.Ordinal));
            StringAssert.Contains(summary, "RawType300MembershipCount=202");
            StringAssert.Contains(summary, "DistinctNonemptyObjectIdCount=201");
            StringAssert.Contains(summary, "UniqueObjectIdCorrelationCount=201");
            var coverage = new MembershipCoverageDiagnosticsBuilder().Build(result);
            var bucket = coverage.SemanticKinds.Single(item => item.SemanticKind ==
                "unsupported:componenttype:300");
            Assert.AreEqual(1, bucket.DiagnosticGroups.Count);
            Assert.AreEqual(202, bucket.DiagnosticGroups.Single().Count);
            Assert.AreEqual(202, bucket.AuditEvidence.Count);
        }

        [TestMethod]
        public void Type300BlankNameIsExplicitAndRemainsDiagnosticOnly()
        {
            var solution = Solution(); var objectId = Guid.NewGuid();
            var result = new DataverseComponentIdentityResolver().Resolve(
                CanvasAppService(solution, query => Rows(
                    CanvasApp(objectId, " ", "Display only", "unique-canvas-app", false))),
                solution.Environment, new SolutionComponentRecord(Guid.NewGuid(), 300, objectId),
                CancellationToken.None);

            Assert.AreEqual(IdentityResolutionStatus.Unsupported, result.Status);
            Assert.AreEqual("unsupported:componenttype:300", result.SemanticKind);
            Assert.IsNull(result.ComparisonKey);
            StringAssert.Contains(result.DiagnosticEvidence.First(), "name=' '");
            var summary = result.DiagnosticEvidence.Single(item =>
                item.StartsWith("Canvas App diagnostic summary:", StringComparison.Ordinal));
            StringAssert.Contains(summary, "UniqueObjectIdCorrelationCount=1");
            StringAssert.Contains(summary, "BlankNameCount=1");
            StringAssert.Contains(summary, "DistinctCandidateNames=[]");
        }

        [TestMethod]
        public void Type300ZeroAndDuplicateResultsRemainDiagnosticOnly()
        {
            var solution = Solution(); var missingId = Guid.NewGuid(); var duplicateId = Guid.NewGuid();
            var service = CanvasAppService(solution, query => Rows(
                CanvasApp(duplicateId, "new_First", "First", "unique-first", false),
                CanvasApp(duplicateId, "new_Second", "Second", "unique-second", true)));
            var records = new[] { missingId, duplicateId }.Select(objectId =>
                new ComponentIdentity(new SolutionComponentRecord(Guid.NewGuid(), 300, objectId),
                    IdentityResolutionStatus.Unresolved)).ToArray();

            var result = new DataverseComponentIdentityResolver().ResolveSnapshot(service,
                MembershipSnapshot.Complete(solution, records, DateTimeOffset.UtcNow), CancellationToken.None);

            StringAssert.Contains(result.Components[0].DiagnosticEvidence.First(), "No canvasapp row matched");
            StringAssert.Contains(result.Components[1].DiagnosticEvidence.First(), "Multiple canvasapp rows matched");
            Assert.IsTrue(result.Components[1].DiagnosticEvidence.Any(item => item.Contains("name='new_First'")));
            Assert.IsTrue(result.Components[1].DiagnosticEvidence.Any(item => item.Contains("name='new_Second'")));
            var summary = result.Components.SelectMany(item => item.DiagnosticEvidence).Single(item =>
                item.StartsWith("Canvas App diagnostic summary:", StringComparison.Ordinal));
            StringAssert.Contains(summary, "ReturnedCanvasAppRowCount=2");
            StringAssert.Contains(summary, "UniqueObjectIdCorrelationCount=0");
            StringAssert.Contains(summary, "MissingRequestedObjectIdCount=1");
            StringAssert.Contains(summary, "NonUniqueObjectIdCount=1");
            Assert.IsTrue(result.Components.All(item =>
                item.Status == IdentityResolutionStatus.Unsupported && item.ComparisonKey == null));
        }

        [TestMethod]
        public void Type300ConflictingResultPreservesEvidenceAndRemainsUnsupported()
        {
            var solution = Solution(); var objectId = Guid.NewGuid();
            var conflicting = CanvasApp(objectId, "new_Conflict", "Conflict", "unique-conflict", false);
            conflicting.Id = Guid.NewGuid();

            var result = new DataverseComponentIdentityResolver().Resolve(
                CanvasAppService(solution, query => Rows(conflicting)), solution.Environment,
                new SolutionComponentRecord(Guid.NewGuid(), 300, objectId), CancellationToken.None);

            Assert.AreEqual(IdentityResolutionStatus.Unsupported, result.Status);
            Assert.AreEqual("unsupported:componenttype:300", result.SemanticKind);
            Assert.IsNull(result.ComparisonKey);
            Assert.IsTrue(result.DiagnosticEvidence.Any(item => item.Contains("conflicting or incomplete")));
            Assert.IsTrue(result.DiagnosticEvidence.Any(item => item.Contains("name='new_Conflict'")));
            StringAssert.Contains(result.DiagnosticEvidence.Single(item =>
                item.StartsWith("Canvas App diagnostic summary:", StringComparison.Ordinal)),
                "ReturnedCanvasAppRowCount=(unavailable)");
        }

        [TestMethod]
        public void Type300FaultRemainsDiagnosticOnlyAndCancellationPropagates()
        {
            var solution = Solution(); var objectId = Guid.NewGuid();
            var faulted = new DataverseComponentIdentityResolver().Resolve(
                CanvasAppService(solution, query => throw new FaultException("Canvas App denied")),
                solution.Environment, new SolutionComponentRecord(Guid.NewGuid(), 300, objectId),
                CancellationToken.None);
            Assert.AreEqual(IdentityResolutionStatus.Unsupported, faulted.Status);
            Assert.AreEqual("unsupported:componenttype:300", faulted.SemanticKind);
            Assert.IsNull(faulted.ComparisonKey);
            StringAssert.Contains(faulted.DiagnosticEvidence.First(), "Canvas App denied");

            using (var cancellation = new CancellationTokenSource())
            {
                var service = CanvasAppService(solution, query =>
                {
                    cancellation.Cancel();
                    return Rows(CanvasApp(objectId, "new_App", "App", "unique-app", false));
                });
                Assert.ThrowsException<OperationCanceledException>(() =>
                    new DataverseComponentIdentityResolver().Resolve(service, solution.Environment,
                        new SolutionComponentRecord(Guid.NewGuid(), 300, objectId), cancellation.Token));
            }
        }

        [TestMethod]
        public void Type300NamesRemainEvidenceAndCannotCreatePortableMatches()
        {
            var sourceSolution = Solution();
            var targetSolution = new SolutionIdentity(new EnvironmentIdentity(Guid.NewGuid(), "Target"),
                Guid.NewGuid(), sourceSolution.UniqueName);
            var resolver = new DataverseComponentIdentityResolver();
            var sourceId = Guid.NewGuid(); var targetId = Guid.NewGuid();
            var source = resolver.Resolve(CanvasAppService(sourceSolution, query => Rows(
                    CanvasApp(sourceId, "new_PortableApp", "DEV App", "dev-unique", false))),
                sourceSolution.Environment, new SolutionComponentRecord(Guid.NewGuid(), 300, sourceId),
                CancellationToken.None);
            var target = resolver.Resolve(CanvasAppService(targetSolution, query => Rows(
                    CanvasApp(targetId, "NEW_PORTABLEAPP", "UAT App", "uat-unique", true))),
                targetSolution.Environment, new SolutionComponentRecord(Guid.NewGuid(), 300, targetId),
                CancellationToken.None);

            Assert.IsTrue(source.DiagnosticEvidence.Any(item => item.Contains("name='new_PortableApp'")));
            Assert.IsTrue(target.DiagnosticEvidence.Any(item => item.Contains("name='NEW_PORTABLEAPP'")));
            var compared = new SolutionMembershipComparer().Compare(
                MembershipSnapshot.Complete(sourceSolution, new[] { source }, DateTimeOffset.UtcNow),
                MembershipSnapshot.Complete(targetSolution, new[] { target }, DateTimeOffset.UtcNow));
            Assert.AreEqual(2, compared.Count);
            Assert.IsTrue(compared.All(item => item.Presence == MembershipPresence.Indeterminate));
            Assert.IsTrue(compared.All(item => item.Source == null || item.Source.ComparisonKey == null));
            Assert.IsTrue(compared.All(item => item.Target == null || item.Target.ComparisonKey == null));

            var coverage = new MembershipCoverageDiagnosticsBuilder().Build(
                MembershipSnapshot.Complete(sourceSolution, new[] { source }, DateTimeOffset.UtcNow));
            var bucket = coverage.SemanticKinds.Single(item => item.SemanticKind ==
                "unsupported:componenttype:300");
            Assert.AreEqual(MembershipCoverageBucketType.KnownUnsupportedIsolatedType, bucket.BucketType);
            Assert.AreEqual(1, bucket.DiagnosticGroups.Count);
            Assert.AreEqual(source.Diagnostic, bucket.DiagnosticGroups.Single().Diagnostic);
        }

        [TestMethod]
        public void Type300MissingObjectIdDoesNotQueryCanvasApp()
        {
            var solution = Solution(); int queryCount = 0;
            var result = new DataverseComponentIdentityResolver().Resolve(
                CanvasAppService(solution, query => { queryCount++; return Rows(); }), solution.Environment,
                new SolutionComponentRecord(Guid.NewGuid(), 300, null), CancellationToken.None);

            Assert.AreEqual(0, queryCount);
            Assert.AreEqual(IdentityResolutionStatus.Unsupported, result.Status);
            Assert.AreEqual("unsupported:componenttype:300", result.SemanticKind);
            Assert.IsNull(result.ComparisonKey);
            StringAssert.Contains(result.DiagnosticEvidence.First(), "objectid is unavailable");
        }

        [TestMethod]
        public void Type511SingleCompleteTeamTemplateIsClassifiedWithoutPortableIdentity()
        {
            var solution = Solution(); var objectId = Guid.NewGuid(); var componentId = Guid.NewGuid();
            int teamTemplateQueries = 0; int entityNameMetadataQueries = 0;
            var service = TeamTemplateService(solution, query =>
            {
                teamTemplateQueries++;
                AssertTeamTemplateQuery(query, objectId);
                return Rows(TeamTemplate(objectId, "Account access team", 1, 3, false));
            }, request =>
            {
                var codes = AssertEntityLogicalNameMetadataQuery(request);
                entityNameMetadataQueries++;
                CollectionAssert.AreEquivalent(new[] { 1 }, codes);
                return MetadataRows(EntityMetadata(1, "account", "Account"));
            });

            var result = new DataverseComponentIdentityResolver().Resolve(service, solution.Environment,
                new SolutionComponentRecord(componentId, 511, objectId), CancellationToken.None);

            Assert.AreEqual(1, teamTemplateQueries);
            Assert.AreEqual(1, entityNameMetadataQueries);
            Assert.AreEqual(IdentityResolutionStatus.Unsupported, result.Status);
            Assert.AreEqual(ComponentSemanticKinds.TeamTemplate, result.SemanticKind);
            Assert.AreEqual(ComponentSemanticKinds.TeamTemplate, result.ComponentTypeKey);
            Assert.IsNull(result.ComparisonKey);
            var evidence = result.DiagnosticEvidence.Single();
            StringAssert.Contains(evidence, "teamtemplateid=" + objectId.ToString("D"));
            StringAssert.Contains(evidence, "teamtemplatename='Account access team'");
            StringAssert.Contains(evidence, "objecttypecode=1");
            StringAssert.Contains(evidence, "entitylogicalname=account");
            StringAssert.Contains(evidence, "defaultaccessrightsmask=3");
            StringAssert.Contains(evidence, "componentidunique=");
            StringAssert.Contains(evidence, "componentstate=0 ('Published')");
            StringAssert.Contains(evidence, "ismanaged=False");

            var coverage = new MembershipCoverageDiagnosticsBuilder().Build(
                MembershipSnapshot.Complete(solution, new[] { result }, DateTimeOffset.UtcNow));
            Assert.AreEqual(0, coverage.BroadUnclassifiable.TotalCandidates);
            Assert.IsFalse(coverage.BroadRawComponentTypes.Any(item => item.ComponentType == 511));
            var bucket = coverage.SemanticKinds.Single(item =>
                item.SemanticKind == ComponentSemanticKinds.TeamTemplate);
            Assert.AreEqual("Team Template", bucket.DisplayName);
            Assert.AreEqual(1, bucket.Unsupported);
            Assert.AreEqual(MembershipCoverageStatus.Incomplete, bucket.CoverageStatus);
            var audit = bucket.AuditEvidence.Single();
            Assert.AreEqual(componentId, audit.SolutionComponentId);
            Assert.AreEqual(objectId, audit.ObjectId);
            Assert.AreEqual(evidence, audit.DiagnosticEvidence.Single());
        }

        [TestMethod]
        public void Type511TeamTemplateLookupZeroResultRemainsBroadWithStableEvidence()
        {
            var solution = Solution(); var objectId = Guid.NewGuid();
            var result = new DataverseComponentIdentityResolver().Resolve(
                TeamTemplateService(solution, query => Rows()), solution.Environment,
                new SolutionComponentRecord(Guid.NewGuid(), 511, objectId), CancellationToken.None);

            Assert.AreEqual(IdentityResolutionStatus.Unsupported, result.Status);
            Assert.IsNull(result.SemanticKind);
            Assert.IsNull(result.ComparisonKey);
            Assert.AreEqual(1, result.DiagnosticEvidence.Count);
            StringAssert.Contains(result.DiagnosticEvidence.Single(), "No teamtemplate row matched");
            Assert.IsFalse(result.DiagnosticEvidence.Single().Contains(objectId.ToString("D")));
        }

        [TestMethod]
        public void Type511DuplicateTeamTemplateResultsRemainBroadAndPreserveEveryReturnedRow()
        {
            var solution = Solution(); var objectId = Guid.NewGuid();
            var result = new DataverseComponentIdentityResolver().Resolve(
                TeamTemplateService(solution, query => Rows(
                    TeamTemplate(objectId, "First template", 1, 1, false),
                    TeamTemplate(objectId, "Second template", 1, 2, true)),
                    request => MetadataRows(EntityMetadata(1, "account", "Account"))),
                solution.Environment, new SolutionComponentRecord(Guid.NewGuid(), 511, objectId),
                CancellationToken.None);

            Assert.AreEqual(IdentityResolutionStatus.Unsupported, result.Status);
            Assert.IsNull(result.SemanticKind);
            Assert.IsNull(result.ComparisonKey);
            Assert.AreEqual(3, result.DiagnosticEvidence.Count);
            StringAssert.Contains(result.DiagnosticEvidence[0], "Multiple teamtemplate rows matched");
            Assert.IsTrue(result.DiagnosticEvidence.Any(item => item.Contains("teamtemplatename='First template'")));
            Assert.IsTrue(result.DiagnosticEvidence.Any(item => item.Contains("teamtemplatename='Second template'")));
        }

        [TestMethod]
        public void Type511ConflictingTeamTemplateResultRemainsBroadAndPreservesReturnedValues()
        {
            var solution = Solution(); var objectId = Guid.NewGuid();
            var conflicting = TeamTemplate(objectId, "Conflicting template", 1, 7, false);
            conflicting.Id = Guid.NewGuid();
            var result = new DataverseComponentIdentityResolver().Resolve(
                TeamTemplateService(solution, query => Rows(conflicting)), solution.Environment,
                new SolutionComponentRecord(Guid.NewGuid(), 511, objectId), CancellationToken.None);

            Assert.AreEqual(IdentityResolutionStatus.Unsupported, result.Status);
            Assert.IsNull(result.SemanticKind);
            Assert.IsNull(result.ComparisonKey);
            Assert.IsTrue(result.DiagnosticEvidence.Any(item => item.Contains("conflicting or incomplete")));
            Assert.IsTrue(result.DiagnosticEvidence.Any(item => item.Contains("Conflicting template")));
        }

        [TestMethod]
        public void Type511IncompleteTeamTemplateRowRemainsBroad()
        {
            var solution = Solution(); var objectId = Guid.NewGuid();
            var incomplete = TeamTemplate(objectId, "Incomplete template", 1, 1, false);
            incomplete.Attributes.Remove("componentidunique");
            var result = new DataverseComponentIdentityResolver().Resolve(
                TeamTemplateService(solution, query => Rows(incomplete),
                    request => MetadataRows(EntityMetadata(1, "account", "Account"))),
                solution.Environment, new SolutionComponentRecord(Guid.NewGuid(), 511, objectId),
                CancellationToken.None);

            Assert.AreEqual(IdentityResolutionStatus.Unsupported, result.Status);
            Assert.IsNull(result.SemanticKind);
            Assert.IsNull(result.ComparisonKey);
            StringAssert.Contains(result.DiagnosticEvidence.Single(), "returned incomplete data");
            StringAssert.Contains(result.DiagnosticEvidence.Single(), "componentidunique=(not supplied)");
        }

        [TestMethod]
        public void Type511IncompleteTeamTemplateResultSetRemainsBroad()
        {
            var solution = Solution(); var objectId = Guid.NewGuid();
            var result = new DataverseComponentIdentityResolver().Resolve(
                TeamTemplateService(solution, query =>
                {
                    var rows = Rows(TeamTemplate(objectId, "Partial template", 1, 1, false));
                    rows.MoreRecords = true;
                    return rows;
                }, request => MetadataRows(EntityMetadata(1, "account", "Account"))),
                solution.Environment, new SolutionComponentRecord(Guid.NewGuid(), 511, objectId),
                CancellationToken.None);

            Assert.AreEqual(IdentityResolutionStatus.Unsupported, result.Status);
            Assert.IsNull(result.SemanticKind);
            Assert.IsNull(result.ComparisonKey);
            Assert.IsTrue(result.DiagnosticEvidence.Any(item => item.Contains("incomplete result set")));
            Assert.IsTrue(result.DiagnosticEvidence.Any(item => item.Contains("Partial template")));
        }

        [TestMethod]
        public void Type511MissingEntityMetadataMatchRemainsBroad()
        {
            var solution = Solution(); var objectId = Guid.NewGuid();
            var result = new DataverseComponentIdentityResolver().Resolve(
                TeamTemplateService(solution,
                    query => Rows(TeamTemplate(objectId, "Template", 1, 1, false)),
                    request => MetadataRows()),
                solution.Environment, new SolutionComponentRecord(Guid.NewGuid(), 511, objectId),
                CancellationToken.None);

            Assert.AreEqual(IdentityResolutionStatus.Unsupported, result.Status);
            Assert.IsNull(result.SemanticKind);
            Assert.IsNull(result.ComparisonKey);
            StringAssert.Contains(result.DiagnosticEvidence.Single(), "no entity metadata match");
        }

        [TestMethod]
        public void Type511TeamTemplateLookupBatchesDeduplicatesAndCountsRequests()
        {
            var solution = Solution();
            var objectIds = Enumerable.Range(0, 201).Select(index => Guid.NewGuid()).ToList();
            var records = objectIds.Concat(new[] { objectIds[0] }).Select(objectId =>
                new ComponentIdentity(new SolutionComponentRecord(Guid.NewGuid(), 511, objectId),
                    IdentityResolutionStatus.Unresolved)).ToArray();
            var queriedIds = new List<Guid>(); int entityNameMetadataQueries = 0;
            var service = TeamTemplateService(solution, query =>
            {
                var ids = query.Criteria.Conditions.Single().Values.Cast<Guid>().ToList();
                Assert.IsTrue(ids.Count <= 200);
                Assert.AreEqual(ids.Count, ids.Distinct().Count());
                queriedIds.AddRange(ids);
                return Rows(ids.Select(id => TeamTemplate(id, "Template", 1, 1, false)).ToArray());
            }, request =>
            {
                entityNameMetadataQueries++;
                CollectionAssert.AreEquivalent(new[] { 1 }, AssertEntityLogicalNameMetadataQuery(request));
                return MetadataRows(EntityMetadata(1, "account", "Account"));
            });
            var counter = new D365SolutionComparer.Infrastructure.DataverseRequestCounter();

            var result = new DataverseComponentIdentityResolver().ResolveSnapshot(service,
                MembershipSnapshot.Complete(solution, records, DateTimeOffset.UtcNow),
                CancellationToken.None, counter);

            Assert.AreEqual(201, queriedIds.Count);
            CollectionAssert.AreEquivalent(objectIds, queriedIds);
            Assert.AreEqual(1, entityNameMetadataQueries);
            Assert.AreEqual(2, counter.GetQueryCount("teamtemplate"));
            Assert.AreEqual(2, counter.GetQueryCount("solutioncomponentdefinition"));
            Assert.AreEqual(2, counter.GetExecuteCount("RetrieveMetadataChanges"));
            Assert.AreEqual(1, counter.GetExecuteCount("RetrieveAttribute"));
            Assert.AreEqual(1, counter.GetExecuteCount("WhoAmI"));
            Assert.AreEqual(8, counter.TotalRequests);
            Assert.AreEqual(202, result.Components.Count);
            Assert.IsTrue(result.Components.All(item => item.Status == IdentityResolutionStatus.Unsupported &&
                item.SemanticKind == ComponentSemanticKinds.TeamTemplate && item.ComparisonKey == null &&
                item.DiagnosticEvidence.Single().Contains("entitylogicalname=account")));
        }

        [TestMethod]
        public void Type511TeamTemplateFaultAndMetadataFaultRemainDiagnosticOnly()
        {
            var solution = Solution(); var objectId = Guid.NewGuid();
            var lookupFault = new DataverseComponentIdentityResolver().Resolve(
                TeamTemplateService(solution, query => throw new FaultException("TeamTemplate denied")),
                solution.Environment, new SolutionComponentRecord(Guid.NewGuid(), 511, objectId),
                CancellationToken.None);
            Assert.AreEqual(IdentityResolutionStatus.Unsupported, lookupFault.Status);
            Assert.IsNull(lookupFault.SemanticKind);
            Assert.IsNull(lookupFault.ComparisonKey);
            StringAssert.Contains(lookupFault.DiagnosticEvidence.Single(), "TeamTemplate denied");

            var metadataFault = new DataverseComponentIdentityResolver().Resolve(
                TeamTemplateService(solution,
                    query => Rows(TeamTemplate(objectId, "Template", 1, 1, false)),
                    request => throw new FaultException("Entity metadata denied")),
                solution.Environment, new SolutionComponentRecord(Guid.NewGuid(), 511, objectId),
                CancellationToken.None);
            Assert.AreEqual(IdentityResolutionStatus.Unsupported, metadataFault.Status);
            Assert.IsNull(metadataFault.SemanticKind);
            Assert.IsNull(metadataFault.ComparisonKey);
            StringAssert.Contains(metadataFault.DiagnosticEvidence.Single(), "Entity metadata denied");
        }

        [TestMethod]
        public void CancellationDuringType511TeamTemplateLookupPropagates()
        {
            var solution = Solution();
            using (var cancellation = new CancellationTokenSource())
            {
                var service = TeamTemplateService(solution, query =>
                {
                    cancellation.Cancel();
                    return Rows();
                });
                Assert.ThrowsException<OperationCanceledException>(() =>
                    new DataverseComponentIdentityResolver().Resolve(service, solution.Environment,
                        new SolutionComponentRecord(Guid.NewGuid(), 511, Guid.NewGuid()), cancellation.Token));
            }
        }

        [TestMethod]
        public void CancellationDuringType511EntityLogicalNameLookupPropagates()
        {
            var solution = Solution(); var objectId = Guid.NewGuid();
            using (var cancellation = new CancellationTokenSource())
            {
                var service = TeamTemplateService(solution,
                    query => Rows(TeamTemplate(objectId, "Template", 1, 1, false)),
                    request =>
                    {
                        cancellation.Cancel();
                        return MetadataRows(EntityMetadata(1, "account", "Account"));
                    });
                Assert.ThrowsException<OperationCanceledException>(() =>
                    new DataverseComponentIdentityResolver().Resolve(service, solution.Environment,
                        new SolutionComponentRecord(Guid.NewGuid(), 511, objectId), cancellation.Token));
            }
        }

        [TestMethod]
        public void VerifiedType511IsolatedCoverageAllowsColumnAbsenceButNeverTeamTemplateAbsence()
        {
            var solution = Solution(); var objectId = Guid.NewGuid();
            var type511 = new DataverseComponentIdentityResolver().Resolve(
                TeamTemplateService(solution,
                    query => Rows(TeamTemplate(objectId, "Account template", 1, 1, false)),
                    request => MetadataRows(EntityMetadata(1, "account", "Account"))),
                solution.Environment, new SolutionComponentRecord(Guid.NewGuid(), 511, objectId),
                CancellationToken.None);
            var sourceColumn = Identity("account.name", 2, kind: ComponentSemanticKinds.Column);

            var comparison = new SolutionMembershipComparer().Compare(
                MembershipSnapshot.Complete(solution, new[] { sourceColumn }, DateTimeOffset.UtcNow),
                MembershipSnapshot.Complete(solution, new[] { type511 }, DateTimeOffset.UtcNow));

            Assert.AreEqual(MembershipPresence.OnlyInSource,
                comparison.Single(item => item.Source == sourceColumn).Presence);
            Assert.AreEqual(MembershipPresence.Indeterminate,
                comparison.Single(item => item.Target != null).Presence);
            Assert.AreEqual(ComponentSemanticKinds.TeamTemplate, type511.SemanticKind);
            Assert.IsNull(type511.ComparisonKey);

            var coverage = new MembershipCoverageDiagnosticsBuilder().Build(
                MembershipSnapshot.Complete(solution, new[] { type511 }, DateTimeOffset.UtcNow));
            Assert.AreEqual(0, coverage.BroadUnclassifiable.TotalCandidates);
            Assert.AreEqual(MembershipCoverageStatus.Incomplete, coverage.SemanticKinds.Single(item =>
                item.SemanticKind == ComponentSemanticKinds.TeamTemplate).CoverageStatus);
            Assert.AreEqual(MembershipCoverageStatus.Complete, coverage.SemanticKinds.Single(item =>
                item.SemanticKind == ComponentSemanticKinds.Column).CoverageStatus);
        }

        [DataTestMethod]
        [DataRow(3)]
        [DataRow(11)]
        [DataRow(12)]
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

        [DataTestMethod]
        [DataRow(26)]
        [DataRow(36)]
        [DataRow(59)]
        public void WeakIdentityLookupUsesExactRequestShapeAndRemainsUnsupported(int componentType)
        {
            var solution = Solution(); var objectId = Guid.NewGuid();
            var service = WeakIdentityService(solution, componentType, query =>
            {
                AssertWeakIdentityQuery(query, componentType, objectId);
                return Rows(WeakIdentityRow(componentType, objectId, "Diagnostic name"));
            });

            var result = new DataverseComponentIdentityResolver().Resolve(service, solution.Environment,
                new SolutionComponentRecord(Guid.NewGuid(), componentType, objectId), CancellationToken.None);

            Assert.AreEqual(IdentityResolutionStatus.Unsupported, result.Status);
            Assert.AreEqual("unsupported:componenttype:" + componentType, result.SemanticKind);
            Assert.IsNull(result.ComparisonKey);
            Assert.AreEqual("No identity resolver supports this known component type.", result.Diagnostic);
            var evidence = result.DiagnosticEvidence.First();
            foreach (var column in WeakIdentityColumns(componentType))
                StringAssert.Contains(evidence, column + "=");
            StringAssert.Contains(evidence, "Diagnostic evidence only; no value is used for membership comparison.");
            Assert.IsTrue(result.DiagnosticEvidence.Any(item => item.Contains("RawComponentCount=1") &&
                item.Contains("CorrelatedCount=1") && item.Contains("IncompleteCorrelatedRowCount=0")));
        }

        [DataTestMethod]
        [DataRow(26)]
        [DataRow(36)]
        [DataRow(59)]
        public void WeakIdentityLookupBatchesDeduplicatesCountsAndKeepsStableGrouping(int componentType)
        {
            var solution = Solution();
            var objectIds = Enumerable.Range(0, 201).Select(index => Guid.NewGuid()).ToList();
            var records = objectIds.Concat(new[] { objectIds[0] }).Select(objectId =>
                new ComponentIdentity(new SolutionComponentRecord(Guid.NewGuid(), componentType, objectId),
                    IdentityResolutionStatus.Unresolved)).ToArray();
            var queriedIds = new List<Guid>();
            var service = WeakIdentityService(solution, componentType, query =>
            {
                var ids = query.Criteria.Conditions.Single().Values.Cast<Guid>().ToList();
                Assert.IsTrue(ids.Count <= 200);
                Assert.AreEqual(ids.Count, ids.Distinct().Count());
                queriedIds.AddRange(ids);
                return Rows(ids.Select(id => WeakIdentityRow(componentType, id,
                    "Different " + id.ToString("D"))).ToArray());
            });
            var counter = new D365SolutionComparer.Infrastructure.DataverseRequestCounter();

            var result = new DataverseComponentIdentityResolver().ResolveSnapshot(service,
                MembershipSnapshot.Complete(solution, records, DateTimeOffset.UtcNow),
                CancellationToken.None, counter);

            CollectionAssert.AreEquivalent(objectIds, queriedIds);
            Assert.AreEqual(2, counter.GetQueryCount(WeakIdentityEntity(componentType)));
            Assert.AreEqual(1, counter.GetExecuteCount("WhoAmI"));
            Assert.AreEqual(3, counter.TotalRequests);
            var bucket = new MembershipCoverageDiagnosticsBuilder().Build(result).SemanticKinds.Single(item =>
                item.SemanticKind == "unsupported:componenttype:" + componentType);
            Assert.AreEqual(MembershipCoverageBucketType.KnownUnsupportedIsolatedType, bucket.BucketType);
            Assert.AreEqual(1, bucket.DiagnosticGroups.Count);
            Assert.AreEqual(202, bucket.DiagnosticGroups.Single().Count);
            Assert.AreEqual(202, bucket.AuditEvidence.Count);
        }

        [DataTestMethod]
        [DataRow(26)]
        [DataRow(36)]
        [DataRow(59)]
        public void WeakIdentityMissingDuplicateAndBlankRowsRemainDiagnosticOnly(int componentType)
        {
            var solution = Solution(); var missingId = Guid.NewGuid(); var duplicateId = Guid.NewGuid();
            var blankId = Guid.NewGuid();
            var service = WeakIdentityService(solution, componentType, query => Rows(
                WeakIdentityRow(componentType, duplicateId, "First"),
                WeakIdentityRow(componentType, duplicateId, "Second"),
                WeakIdentityRow(componentType, blankId, "")));
            var records = new[] { missingId, duplicateId, blankId }.Select(objectId =>
                new ComponentIdentity(new SolutionComponentRecord(Guid.NewGuid(), componentType, objectId),
                    IdentityResolutionStatus.Unresolved)).ToArray();

            var result = new DataverseComponentIdentityResolver().ResolveSnapshot(service,
                MembershipSnapshot.Complete(solution, records, DateTimeOffset.UtcNow), CancellationToken.None);

            Assert.IsTrue(result.Components[0].DiagnosticEvidence.Any(item => item.StartsWith("No ")));
            Assert.IsTrue(result.Components[1].DiagnosticEvidence.Any(item => item.StartsWith("Multiple ")));
            Assert.IsTrue(result.Components[2].DiagnosticEvidence.Any(item =>
                item.Contains("matched but returned incomplete data")));
            var summary = result.Components.SelectMany(item => item.DiagnosticEvidence).Single(item =>
                item.Contains(" diagnostic summary:"));
            StringAssert.Contains(summary, "MissingCount=1");
            StringAssert.Contains(summary, "NonUniqueObjectIdCount=1");
            StringAssert.Contains(summary, "BlankNameCount=1");
            Assert.IsTrue(result.Components.All(item => item.Status == IdentityResolutionStatus.Unsupported &&
                item.ComparisonKey == null));
        }

        [DataTestMethod]
        [DataRow(26)]
        [DataRow(36)]
        [DataRow(59)]
        public void WeakIdentityConflictsAndPagingAnomaliesRemainConservative(int componentType)
        {
            var solution = Solution(); var objectId = Guid.NewGuid();
            var conflicting = WeakIdentityRow(componentType, objectId, "Conflict");
            conflicting.Id = Guid.NewGuid();
            var conflict = new DataverseComponentIdentityResolver().Resolve(
                WeakIdentityService(solution, componentType, query => Rows(conflicting)), solution.Environment,
                new SolutionComponentRecord(Guid.NewGuid(), componentType, objectId), CancellationToken.None);
            Assert.IsTrue(conflict.DiagnosticEvidence.Any(item => item.Contains("conflicting or incomplete")));
            Assert.IsTrue(conflict.DiagnosticEvidence.Any(item => item.Contains("Conflict")));

            var paged = new DataverseComponentIdentityResolver().Resolve(
                WeakIdentityService(solution, componentType, query =>
                {
                    var rows = Rows(WeakIdentityRow(componentType, objectId, "Partial"));
                    rows.MoreRecords = true;
                    return rows;
                }), solution.Environment, new SolutionComponentRecord(Guid.NewGuid(), componentType, objectId),
                CancellationToken.None);
            Assert.IsTrue(paged.DiagnosticEvidence.Any(item => item.Contains("incomplete result set")));
            Assert.IsTrue(new[] { conflict, paged }.All(item => item.Status == IdentityResolutionStatus.Unsupported &&
                item.ComparisonKey == null));
        }

        [DataTestMethod]
        [DataRow(26)]
        [DataRow(36)]
        [DataRow(59)]
        public void WeakIdentityFaultIsDiagnosticAndCancellationPropagates(int componentType)
        {
            var solution = Solution(); var objectId = Guid.NewGuid();
            var faulted = new DataverseComponentIdentityResolver().Resolve(
                WeakIdentityService(solution, componentType, query => throw new FaultException("Denied")),
                solution.Environment, new SolutionComponentRecord(Guid.NewGuid(), componentType, objectId),
                CancellationToken.None);
            Assert.AreEqual(IdentityResolutionStatus.Unsupported, faulted.Status);
            StringAssert.Contains(faulted.DiagnosticEvidence.First(), "Denied");

            using (var cancellation = new CancellationTokenSource())
            {
                var service = WeakIdentityService(solution, componentType, query =>
                {
                    cancellation.Cancel();
                    return Rows();
                });
                Assert.ThrowsException<OperationCanceledException>(() =>
                    new DataverseComponentIdentityResolver().Resolve(service, solution.Environment,
                        new SolutionComponentRecord(Guid.NewGuid(), componentType, objectId), cancellation.Token));
            }
        }

        [DataTestMethod]
        [DataRow(26)]
        [DataRow(36)]
        [DataRow(59)]
        public void WeakIdentityDiagnosticContextCannotCreateMembershipMatches(int componentType)
        {
            var sourceSolution = Solution();
            var targetSolution = new SolutionIdentity(new EnvironmentIdentity(Guid.NewGuid(), "Target"),
                Guid.NewGuid(), sourceSolution.UniqueName);
            var sourceId = Guid.NewGuid(); var targetId = Guid.NewGuid();
            var resolver = new DataverseComponentIdentityResolver();
            var source = resolver.Resolve(WeakIdentityService(sourceSolution, componentType, query =>
                    Rows(WeakIdentityRow(componentType, sourceId, "Same name"))), sourceSolution.Environment,
                new SolutionComponentRecord(Guid.NewGuid(), componentType, sourceId), CancellationToken.None);
            var target = resolver.Resolve(WeakIdentityService(targetSolution, componentType, query =>
                    Rows(WeakIdentityRow(componentType, targetId, "Same name"))), targetSolution.Environment,
                new SolutionComponentRecord(Guid.NewGuid(), componentType, targetId), CancellationToken.None);

            var compared = new SolutionMembershipComparer().Compare(
                MembershipSnapshot.Complete(sourceSolution, new[] { source }, DateTimeOffset.UtcNow),
                MembershipSnapshot.Complete(targetSolution, new[] { target }, DateTimeOffset.UtcNow));

            Assert.AreEqual(2, compared.Count);
            Assert.IsTrue(compared.All(item => item.Presence == MembershipPresence.Indeterminate));
            Assert.IsNull(source.ComparisonKey);
            Assert.IsNull(target.ComparisonKey);
        }

        [DataTestMethod]
        [DataRow(26)]
        [DataRow(36)]
        [DataRow(59)]
        public void WeakIdentityMissingObjectIdDoesNotQueryBackingTable(int componentType)
        {
            var solution = Solution(); int queryCount = 0;
            var result = new DataverseComponentIdentityResolver().Resolve(
                WeakIdentityService(solution, componentType, query => { queryCount++; return Rows(); }),
                solution.Environment, new SolutionComponentRecord(Guid.NewGuid(), componentType, null),
                CancellationToken.None);
            Assert.AreEqual(0, queryCount);
            Assert.AreEqual(IdentityResolutionStatus.Unsupported, result.Status);
            Assert.IsNull(result.ComparisonKey);
            StringAssert.Contains(result.DiagnosticEvidence.First(), "objectid is unavailable");
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

        private static RetrieveAttributeResponse ComponentTypeChoices(params OptionMetadata[] options)
        {
            var optionSet = new OptionSetMetadata();
            foreach (var option in options) optionSet.Options.Add(option);
            var response = new RetrieveAttributeResponse();
            response.Results["AttributeMetadata"] = new PicklistAttributeMetadata { OptionSet = optionSet };
            return response;
        }

        private static FakeOrganizationService OptionSetService(SolutionIdentity solution,
            Func<RetrieveAllOptionSetsRequest, RetrieveAllOptionSetsResponse> optionSetLookup)
        {
            var service = Service(solution, query =>
            {
                Assert.Fail("Type 9 diagnostics must not query table " + query.EntityName);
                return Rows();
            });
            service.ExecuteRequest = request =>
            {
                if (request is WhoAmIRequest) return WhoAmI(solution.Environment.OrganizationId);
                var optionSetRequest = request as RetrieveAllOptionSetsRequest;
                if (optionSetRequest != null) return optionSetLookup(optionSetRequest);
                throw new NotSupportedException(request.RequestName);
            };
            return service;
        }

        private static RetrieveAllOptionSetsResponse AllOptionSetsResponse(
            params OptionSetMetadataBase[] metadata)
        {
            var response = new RetrieveAllOptionSetsResponse();
            response.Results["OptionSetMetadata"] = metadata;
            return response;
        }

        private static OptionSetMetadata OptionSet(Guid metadataId, string name, bool? isGlobal,
            OptionSetType? optionSetType, bool? isManaged, bool? isCustomOptionSet)
        {
            var metadata = new OptionSetMetadata
            {
                MetadataId = metadataId,
                Name = name,
                IsGlobal = isGlobal,
                OptionSetType = optionSetType,
                IsCustomOptionSet = isCustomOptionSet
            };
            typeof(OptionSetMetadataBase).GetProperty("IsManaged").SetValue(metadata, isManaged);
            return metadata;
        }

        private static FakeOrganizationService WeakIdentityService(SolutionIdentity solution,
            int componentType, Func<QueryExpression, EntityCollection> queryHandler)
        {
            var entityName = WeakIdentityEntity(componentType);
            var service = Service(solution, query =>
            {
                if (query.EntityName == entityName) return queryHandler(query);
                Assert.Fail("Type " + componentType + " diagnostics must not query table " + query.EntityName);
                return Rows();
            });
            service.ExecuteRequest = request => request is WhoAmIRequest
                ? (OrganizationResponse)WhoAmI(solution.Environment.OrganizationId)
                : throw new NotSupportedException(request.RequestName);
            return service;
        }

        private static void AssertWeakIdentityQuery(QueryExpression query, int componentType,
            params Guid[] objectIds)
        {
            Assert.AreEqual(WeakIdentityEntity(componentType), query.EntityName);
            CollectionAssert.AreEquivalent(WeakIdentityColumns(componentType),
                query.ColumnSet.Columns.ToArray());
            Assert.AreEqual(1, query.Criteria.Conditions.Count);
            var condition = query.Criteria.Conditions.Single();
            Assert.AreEqual(WeakIdentityPrimaryId(componentType), condition.AttributeName);
            Assert.AreEqual(ConditionOperator.In, condition.Operator);
            Assert.IsTrue(condition.Values.All(value => value != null && value.GetType() == typeof(Guid)));
            CollectionAssert.AreEquivalent(objectIds, condition.Values.Cast<Guid>().ToArray());
        }

        private static Entity WeakIdentityRow(int componentType, Guid id, string name)
        {
            var entityName = WeakIdentityEntity(componentType);
            var primaryId = WeakIdentityPrimaryId(componentType);
            var row = new Entity(entityName, id)
            {
                [primaryId] = id,
                [componentType == 36 ? "title" : "name"] = name,
                [componentType == 26 ? "savedqueryidunique" : componentType == 36
                    ? "templateidunique" : "savedqueryvisualizationidunique"] = Guid.NewGuid(),
                ["componentstate"] = new OptionSetValue(0),
                ["ismanaged"] = false
            };
            if (componentType == 26)
            {
                row["returnedtypecode"] = "account";
                row["querytype"] = 0;
            }
            else if (componentType == 36)
            {
                row["templatetypecode"] = "account";
                row["ispersonal"] = false;
                row["languagecode"] = 1033;
            }
            else
            {
                row["primaryentitytypecode"] = "account";
                row["type"] = new OptionSetValue(0);
                row["charttype"] = new OptionSetValue(0);
            }
            row.FormattedValues["componentstate"] = "Published";
            return row;
        }

        private static string WeakIdentityEntity(int componentType) => componentType == 26
            ? "savedquery" : componentType == 36 ? "template" : "savedqueryvisualization";

        private static string WeakIdentityPrimaryId(int componentType) => componentType == 26
            ? "savedqueryid" : componentType == 36 ? "templateid" : "savedqueryvisualizationid";

        private static string[] WeakIdentityColumns(int componentType)
        {
            if (componentType == 26)
                return new[] { "savedqueryid", "name", "returnedtypecode", "querytype",
                    "savedqueryidunique", "componentstate", "ismanaged" };
            if (componentType == 36)
                return new[] { "templateid", "title", "templatetypecode", "templateidunique",
                    "ispersonal", "languagecode", "componentstate", "ismanaged" };
            return new[] { "savedqueryvisualizationid", "name", "primaryentitytypecode", "type",
                "charttype", "savedqueryvisualizationidunique", "componentstate", "ismanaged" };
        }

        private static FakeOrganizationService ReportService(SolutionIdentity solution,
            Func<QueryExpression, EntityCollection> reportQuery)
        {
            var service = Service(solution, query =>
            {
                if (query.EntityName == "report") return reportQuery(query);
                Assert.Fail("Type 31 diagnostics must not query table " + query.EntityName);
                return Rows();
            });
            service.ExecuteRequest = request => request is WhoAmIRequest
                ? (OrganizationResponse)WhoAmI(solution.Environment.OrganizationId)
                : throw new NotSupportedException(request.RequestName);
            return service;
        }

        private static void AssertReportQuery(QueryExpression query, params Guid[] objectIds)
        {
            Assert.AreEqual("report", query.EntityName);
            CollectionAssert.AreEquivalent(new[] { "reportid", "name", "filename", "reporttypecode",
                "signatureid", "signaturelcid", "reportidunique", "componentstate", "ismanaged" },
                query.ColumnSet.Columns.ToArray());
            Assert.AreEqual(1, query.Criteria.Conditions.Count);
            var condition = query.Criteria.Conditions.Single();
            Assert.AreEqual("reportid", condition.AttributeName);
            Assert.AreEqual(ConditionOperator.In, condition.Operator);
            Assert.IsTrue(condition.Values.All(value => value != null && value.GetType() == typeof(Guid)));
            CollectionAssert.AreEquivalent(objectIds, condition.Values.Cast<Guid>().ToArray());
        }

        private static Entity Report(Guid id, string name, string filename, int reportTypeCode,
            Guid? signatureId, int? signatureLcid, Guid reportIdUnique, bool isManaged)
        {
            var row = new Entity("report", id)
            {
                ["reportid"] = id,
                ["name"] = name,
                ["filename"] = filename,
                ["reporttypecode"] = new OptionSetValue(reportTypeCode),
                ["reportidunique"] = reportIdUnique,
                ["componentstate"] = new OptionSetValue(0),
                ["ismanaged"] = isManaged
            };
            if (signatureId.HasValue) row["signatureid"] = signatureId.Value;
            if (signatureLcid.HasValue) row["signaturelcid"] = signatureLcid.Value;
            row.FormattedValues["reporttypecode"] = reportTypeCode == 1
                ? "Reporting Services Report" : "Other Report";
            row.FormattedValues["componentstate"] = "Published";
            return row;
        }

        private static FakeOrganizationService SystemFormService(SolutionIdentity solution,
            Func<QueryExpression, EntityCollection> systemFormQuery,
            Func<RetrieveMetadataChangesRequest, RetrieveMetadataChangesResponse> metadataQuery = null)
        {
            var service = Service(solution, query =>
            {
                if (query.EntityName == "systemform") return systemFormQuery(query);
                Assert.Fail("Type 60 diagnostics must not query table " + query.EntityName);
                return Rows();
            });
            service.ExecuteRequest = request =>
            {
                if (request is WhoAmIRequest) return WhoAmI(solution.Environment.OrganizationId);
                var metadataRequest = request as RetrieveMetadataChangesRequest;
                if (metadataRequest != null && metadataQuery != null) return metadataQuery(metadataRequest);
                throw new NotSupportedException(request.RequestName);
            };
            return service;
        }

        private static void AssertSystemFormQuery(QueryExpression query, params Guid[] objectIds)
        {
            Assert.AreEqual("systemform", query.EntityName);
            CollectionAssert.AreEquivalent(new[] { "formid", "uniquename", "name", "objecttypecode",
                "type", "formidunique", "componentstate", "ismanaged" }, query.ColumnSet.Columns.ToArray());
            Assert.AreEqual(1, query.Criteria.Conditions.Count);
            var condition = query.Criteria.Conditions.Single();
            Assert.AreEqual("formid", condition.AttributeName);
            Assert.AreEqual(ConditionOperator.In, condition.Operator);
            Assert.IsTrue(condition.Values.All(value => value != null && value.GetType() == typeof(Guid)));
            CollectionAssert.AreEquivalent(objectIds, condition.Values.Cast<Guid>().ToArray());
        }

        private static Entity SystemForm(Guid id, string uniqueName, string name, object objectTypeCode,
            int formType, Guid formIdUnique, bool isManaged)
        {
            var row = new Entity("systemform", id)
            {
                ["formid"] = id,
                ["uniquename"] = uniqueName,
                ["name"] = name,
                ["objecttypecode"] = objectTypeCode,
                ["type"] = new OptionSetValue(formType),
                ["formidunique"] = formIdUnique,
                ["componentstate"] = new OptionSetValue(0),
                ["ismanaged"] = isManaged
            };
            row.FormattedValues["type"] = formType == 2 ? "Main" : "Quick Create";
            row.FormattedValues["componentstate"] = "Published";
            return row;
        }

        private static FakeOrganizationService SiteMapService(SolutionIdentity solution,
            Func<QueryExpression, EntityCollection> siteMapQuery)
        {
            var service = Service(solution, query =>
            {
                if (query.EntityName == "sitemap") return siteMapQuery(query);
                Assert.Fail("Type 62 diagnostics must not query table " + query.EntityName);
                return Rows();
            });
            service.ExecuteRequest = request => request is WhoAmIRequest
                ? (OrganizationResponse)WhoAmI(solution.Environment.OrganizationId)
                : throw new NotSupportedException(request.RequestName);
            return service;
        }

        private static void AssertSiteMapQuery(QueryExpression query, params Guid[] objectIds)
        {
            Assert.AreEqual("sitemap", query.EntityName);
            CollectionAssert.AreEquivalent(new[] { "sitemapid", "sitemapnameunique", "sitemapname",
                "sitemapidunique", "isappaware", "componentstate", "ismanaged" },
                query.ColumnSet.Columns.ToArray());
            Assert.AreEqual(1, query.Criteria.Conditions.Count);
            var condition = query.Criteria.Conditions.Single();
            Assert.AreEqual("sitemapid", condition.AttributeName);
            Assert.AreEqual(ConditionOperator.In, condition.Operator);
            Assert.IsTrue(condition.Values.All(value => value != null && value.GetType() == typeof(Guid)));
            CollectionAssert.AreEquivalent(objectIds, condition.Values.Cast<Guid>().ToArray());
        }

        private static Entity SiteMap(Guid id, string uniqueName, string name, Guid siteMapIdUnique,
            bool isAppAware, bool isManaged)
        {
            var row = new Entity("sitemap", id)
            {
                ["sitemapid"] = id,
                ["sitemapnameunique"] = uniqueName,
                ["sitemapname"] = name,
                ["sitemapidunique"] = siteMapIdUnique,
                ["isappaware"] = isAppAware,
                ["componentstate"] = new OptionSetValue(0),
                ["ismanaged"] = isManaged
            };
            row.FormattedValues["componentstate"] = "Published";
            return row;
        }

        private static FakeOrganizationService CanvasAppService(SolutionIdentity solution,
            Func<QueryExpression, EntityCollection> canvasAppQuery)
        {
            var service = Service(solution, query =>
            {
                if (query.EntityName == "canvasapp") return canvasAppQuery(query);
                Assert.Fail("Type 300 diagnostics must not query table " + query.EntityName);
                return Rows();
            });
            service.ExecuteRequest = request => request is WhoAmIRequest
                ? (OrganizationResponse)WhoAmI(solution.Environment.OrganizationId)
                : throw new NotSupportedException(request.RequestName);
            return service;
        }

        private static void AssertCanvasAppQuery(QueryExpression query, params Guid[] objectIds)
        {
            Assert.AreEqual("canvasapp", query.EntityName);
            CollectionAssert.AreEquivalent(new[] { "canvasappid", "name", "displayname",
                "uniquecanvasappid", "componentstate", "ismanaged" }, query.ColumnSet.Columns.ToArray());
            Assert.AreEqual(1, query.Criteria.Conditions.Count);
            var condition = query.Criteria.Conditions.Single();
            Assert.AreEqual("canvasappid", condition.AttributeName);
            Assert.AreEqual(ConditionOperator.In, condition.Operator);
            Assert.IsTrue(condition.Values.All(value => value != null && value.GetType() == typeof(Guid)));
            CollectionAssert.AreEquivalent(objectIds, condition.Values.Cast<Guid>().ToArray());
        }

        private static Entity CanvasApp(Guid id, string name, string displayName, string uniqueCanvasAppId,
            bool isManaged)
        {
            var row = new Entity("canvasapp", id)
            {
                ["canvasappid"] = id,
                ["name"] = name,
                ["displayname"] = displayName,
                ["uniquecanvasappid"] = uniqueCanvasAppId,
                ["componentstate"] = new OptionSetValue(0),
                ["ismanaged"] = isManaged
            };
            row.FormattedValues["componentstate"] = "Published";
            return row;
        }

        private static FakeOrganizationService TeamTemplateService(SolutionIdentity solution,
            Func<QueryExpression, EntityCollection> teamTemplateQuery,
            Func<RetrieveMetadataChangesRequest, RetrieveMetadataChangesResponse> entityNameMetadataQuery = null)
        {
            var service = Service(solution, query => query.EntityName == "teamtemplate"
                ? teamTemplateQuery(query) : Rows());
            service.ExecuteRequest = request =>
            {
                if (request is WhoAmIRequest) return WhoAmI(solution.Environment.OrganizationId);
                var metadataRequest = request as RetrieveMetadataChangesRequest;
                if (metadataRequest != null)
                {
                    bool isEntityNameLookup = metadataRequest.Query.Properties.PropertyNames.Count == 2 &&
                        metadataRequest.Query.Properties.PropertyNames.Contains("ObjectTypeCode") &&
                        metadataRequest.Query.Properties.PropertyNames.Contains("LogicalName");
                    return isEntityNameLookup && entityNameMetadataQuery != null
                        ? entityNameMetadataQuery(metadataRequest) : MetadataRows();
                }
                if (request is RetrieveAttributeRequest)
                    return ComponentTypeChoices(new OptionMetadata(new Label("Team Template", 1033), 511));
                throw new NotSupportedException(request.RequestName);
            };
            return service;
        }

        private static int[] AssertEntityLogicalNameMetadataQuery(RetrieveMetadataChangesRequest request)
        {
            CollectionAssert.AreEquivalent(new[] { "ObjectTypeCode", "LogicalName" },
                request.Query.Properties.PropertyNames.ToArray());
            Assert.AreEqual(LogicalOperator.Or, request.Query.Criteria.FilterOperator);
            Assert.IsTrue(request.Query.Criteria.Conditions.All(condition =>
                condition.PropertyName == "ObjectTypeCode" &&
                condition.ConditionOperator == MetadataConditionOperator.Equals &&
                condition.Value != null && condition.Value.GetType() == typeof(int)));
            return request.Query.Criteria.Conditions.Select(condition => (int)condition.Value).ToArray();
        }

        private static void AssertTeamTemplateQuery(QueryExpression query, params Guid[] objectIds)
        {
            Assert.AreEqual("teamtemplate", query.EntityName);
            CollectionAssert.AreEquivalent(new[] { "teamtemplateid", "teamtemplatename", "objecttypecode",
                "defaultaccessrightsmask", "componentidunique", "componentstate", "ismanaged" },
                query.ColumnSet.Columns.ToArray());
            Assert.AreEqual(1, query.Criteria.Conditions.Count);
            var condition = query.Criteria.Conditions.Single();
            Assert.AreEqual("teamtemplateid", condition.AttributeName);
            Assert.AreEqual(ConditionOperator.In, condition.Operator);
            Assert.IsTrue(condition.Values.All(value => value != null && value.GetType() == typeof(Guid)));
            CollectionAssert.AreEquivalent(objectIds, condition.Values.Cast<Guid>().ToArray());
        }

        private static Entity TeamTemplate(Guid id, string name, int objectTypeCode, int accessMask,
            bool isManaged)
        {
            var row = new Entity("teamtemplate", id)
            {
                ["teamtemplateid"] = id,
                ["teamtemplatename"] = name,
                ["objecttypecode"] = objectTypeCode,
                ["defaultaccessrightsmask"] = accessMask,
                ["componentidunique"] = Guid.NewGuid(),
                ["componentstate"] = new OptionSetValue(0),
                ["ismanaged"] = isManaged
            };
            row.FormattedValues["componentstate"] = "Published";
            return row;
        }

        private static FakeOrganizationService BroadTypeService(SolutionIdentity solution,
            Func<QueryExpression, EntityCollection> appModuleQuery)
        {
            var service = Service(solution, query => query.EntityName == "appmodule"
                ? appModuleQuery(query) : Rows());
            service.ExecuteRequest = request =>
            {
                if (request is WhoAmIRequest) return WhoAmI(solution.Environment.OrganizationId);
                if (request is RetrieveMetadataChangesRequest) return MetadataRows();
                if (request is RetrieveAttributeRequest)
                    return ComponentTypeChoices(new OptionMetadata(new Label("App Module", 1033), 80));
                throw new NotSupportedException(request.RequestName);
            };
            return service;
        }

        private static void AssertAppModuleQuery(QueryExpression query, params Guid[] objectIds)
        {
            Assert.AreEqual("appmodule", query.EntityName);
            CollectionAssert.AreEquivalent(new[] { "appmoduleid", "uniquename", "name",
                "appmoduleidunique", "componentstate", "ismanaged" }, query.ColumnSet.Columns.ToArray());
            Assert.AreEqual(1, query.Criteria.Conditions.Count);
            var condition = query.Criteria.Conditions.Single();
            Assert.AreEqual("appmoduleid", condition.AttributeName);
            Assert.AreEqual(ConditionOperator.In, condition.Operator);
            Assert.IsTrue(condition.Values.All(value => value != null && value.GetType() == typeof(Guid)));
            CollectionAssert.AreEquivalent(objectIds, condition.Values.Cast<Guid>().ToArray());
        }

        private static Entity AppModule(Guid id, string uniqueName, string name, bool isManaged)
        {
            var row = new Entity("appmodule", id)
            {
                ["appmoduleid"] = id,
                ["uniquename"] = uniqueName,
                ["name"] = name,
                ["appmoduleidunique"] = Guid.NewGuid(),
                ["componentstate"] = new OptionSetValue(0),
                ["ismanaged"] = isManaged
            };
            row.FormattedValues["componentstate"] = "Published";
            return row;
        }
    }
}
