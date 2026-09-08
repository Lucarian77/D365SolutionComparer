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
        public void Type511TeamTemplateLookupAddsCompleteDiagnosticEvidenceWithoutClassification()
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
            Assert.IsNull(result.SemanticKind);
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

            var audit = new MembershipCoverageDiagnosticsBuilder().Build(
                MembershipSnapshot.Complete(solution, new[] { result }, DateTimeOffset.UtcNow))
                .BroadRawComponentTypes.Single(item => item.ComponentType == 511).Evidence.Single();
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
                item.SemanticKind == null && item.ComparisonKey == null &&
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
        public void Type511DiagnosticEvidenceDoesNotChangeBroadAbsenceBlocking()
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

            Assert.AreEqual(MembershipPresence.Indeterminate,
                comparison.Single(item => item.Source == sourceColumn).Presence);
            Assert.AreEqual(MembershipPresence.Indeterminate,
                comparison.Single(item => item.Target != null).Presence);
            Assert.IsNull(type511.SemanticKind);
            Assert.IsNull(type511.ComparisonKey);
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

        private static RetrieveAttributeResponse ComponentTypeChoices(params OptionMetadata[] options)
        {
            var optionSet = new OptionSetMetadata();
            foreach (var option in options) optionSet.Options.Add(option);
            var response = new RetrieveAttributeResponse();
            response.Results["AttributeMetadata"] = new PicklistAttributeMetadata { OptionSet = optionSet };
            return response;
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
