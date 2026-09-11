using System;
using System.Collections.Generic;
using System.Linq;
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
    public class ComponentDefinitionPresentationTests
    {
        [TestMethod, TestCategory("Phase2G2")]
        public void MatchIsMappedWithoutChangingMembershipResult()
        {
            var fixture = Fixture(ComponentSemanticKinds.WebResource, "new_/script.js");
            var presentation = Apply(fixture, Available(fixture.Source, "content", "same"),
                Available(fixture.Target, "content", "same"));

            var row = presentation.Rows.Single();
            Assert.AreEqual("Present in Both", row.MembershipStatus);
            Assert.AreEqual("Match", row.DefinitionStatus);
            Assert.AreEqual(string.Empty, row.ChangedProperties);
            Assert.AreEqual(1, presentation.Summary.PresentInBoth);
        }

        [TestMethod, TestCategory("Phase2G2")]
        public void DifferentListsChangedPropertiesAndDetailValues()
        {
            var fixture = Fixture(ComponentSemanticKinds.WebResource, "new_/script.js");
            var presentation = Apply(fixture, Available(fixture.Source,
                    new KeyValuePair<string, string>("content", "before"),
                    new KeyValuePair<string, string>("description", "source")),
                Available(fixture.Target,
                    new KeyValuePair<string, string>("content", "after"),
                    new KeyValuePair<string, string>("description", "target")));

            var row = presentation.Rows.Single();
            Assert.AreEqual("Different", row.DefinitionStatus);
            Assert.AreEqual("content, description", row.ChangedProperties);
            var content = row.DefinitionDetail.Properties.Single(item => item.PropertyName == "content");
            Assert.AreEqual("before", content.SourceValue);
            Assert.AreEqual("after", content.TargetValue);
            Assert.IsTrue(content.Changed);
            Assert.AreEqual("Different", content.Comparison);
        }

        [DataTestMethod, TestCategory("Phase2G2")]
        [DataRow(true, "SourceOnly")]
        [DataRow(false, "TargetOnly")]
        public void OneSidedDefinitionStatusPreservesAuthoritativeMembership(bool sourceSide,
            string expectedDefinitionStatus)
        {
            var fixture = Fixture(ComponentSemanticKinds.Table, "account");
            var present = sourceSide ? fixture.SourceSnapshot : fixture.TargetSnapshot;
            var empty = EmptySnapshot(sourceSide ? fixture.TargetSnapshot : fixture.SourceSnapshot);
            var sourceMembership = sourceSide ? present : empty;
            var targetMembership = sourceSide ? empty : present;
            var membership = Present(sourceMembership, targetMembership);
            var sourceDefinitions = Definitions(sourceMembership,
                sourceSide ? Available(fixture.Source, "SchemaName", "Account") : null);
            var targetDefinitions = Definitions(targetMembership,
                sourceSide ? null : Available(fixture.Target, "SchemaName", "Account"));

            var result = new ComponentDefinitionResultPresenter().Apply(membership,
                sourceDefinitions, targetDefinitions);

            Assert.AreEqual(sourceSide ? "Source Only" : "Target Only",
                result.Rows.Single().MembershipStatus);
            Assert.AreEqual(expectedDefinitionStatus, result.Rows.Single().DefinitionStatus);
            Assert.IsTrue(result.Rows.Single().DefinitionDetail.Properties.All(item =>
                item.Comparison == (sourceSide ? "Source only" : "Target only")));
        }

        [DataTestMethod, TestCategory("Phase2G2")]
        [DataRow(IdentityResolutionStatus.Unsupported, "Unsupported")]
        [DataRow(IdentityResolutionStatus.Unresolved, "Unresolved")]
        [DataRow(IdentityResolutionStatus.Ambiguous, "Ambiguous")]
        public void ConservativeIdentityStatusesRemainConservative(
            IdentityResolutionStatus identityStatus, string expectedDefinitionStatus)
        {
            var source = Identity(ComponentSemanticKinds.Process, null, identityStatus, 29);
            var sourceSnapshot = Snapshot("Source", source);
            var targetSnapshot = EmptySnapshot(sourceSnapshot, "Target");
            var membership = Present(sourceSnapshot, targetSnapshot);
            var sourceDefinition = new ComponentDefinition(source,
                identityStatus == IdentityResolutionStatus.Unsupported
                    ? ComponentDefinitionReadStatus.Unsupported
                    : identityStatus == IdentityResolutionStatus.Ambiguous
                        ? ComponentDefinitionReadStatus.Ambiguous
                        : ComponentDefinitionReadStatus.Unresolved);

            var result = new ComponentDefinitionResultPresenter().Apply(membership,
                Definitions(sourceSnapshot, sourceDefinition), Definitions(targetSnapshot));

            Assert.AreEqual(expectedDefinitionStatus, result.Rows.Single().DefinitionStatus);
            StringAssert.StartsWith(result.Rows.Single().MembershipStatus, "Indeterminate");
            Assert.AreEqual(0, result.Summary.SourceOnly);
        }

        [TestMethod, TestCategory("Phase2G2")]
        public void IncompleteDefinitionEvidenceNeverPresentsMatch()
        {
            var fixture = Fixture(ComponentSemanticKinds.Column, "account.new_code");
            var source = Available(fixture.Source, "MaxLength", "100");
            var target = new ComponentDefinition(fixture.Target,
                ComponentDefinitionReadStatus.Unresolved,
                diagnostic: "Column metadata was incomplete.");

            var row = Apply(fixture, source, target).Rows.Single();

            Assert.AreEqual("Unresolved", row.DefinitionStatus);
            Assert.AreNotEqual("Match", row.DefinitionStatus);
            StringAssert.Contains(row.DefinitionDetail.Diagnostic, "incomplete");
        }

        [TestMethod, TestCategory("Phase2G2")]
        public void UnavailableEnvironmentLeavesMembershipAndDefinitionIndeterminate()
        {
            var fixture = Fixture(ComponentSemanticKinds.AppModule, "new_app");
            var source = MembershipEnvironmentResult.FromSnapshot("Source", fixture.SourceSnapshot,
                4, TimeSpan.Zero);
            var target = MembershipEnvironmentResult.Unavailable("Target", "sample", 1,
                TimeSpan.Zero, "Access denied");
            var membership = new MembershipResultPresenter().Create(source, target);

            var result = new ComponentDefinitionResultPresenter().Apply(membership,
                Definitions(fixture.SourceSnapshot, Available(fixture.Source, "name", "App")), null);

            Assert.AreEqual("Indeterminate - Environment Unavailable",
                result.Rows.Single().MembershipStatus);
            Assert.AreEqual("Unresolved", result.Rows.Single().DefinitionStatus);
            StringAssert.Contains(result.Rows.Single().DefinitionDetail.Diagnostic,
                "environment retrieval is unavailable");
        }

        [TestMethod, TestCategory("Phase2G2")]
        public void DetailModelContainsOnlyContractPropertiesAndKeepsAuditEvidenceSeparate()
        {
            var fixture = Fixture(ComponentSemanticKinds.AppModule, "new_app");
            var source = new ComponentDefinition(fixture.Source,
                ComponentDefinitionReadStatus.Available, Complete(fixture.Source,
                    new KeyValuePair<string, string>("name", "Application")),
                diagnosticEvidence: new[] { "appmoduleid=" + Guid.NewGuid() });
            var target = new ComponentDefinition(fixture.Target,
                ComponentDefinitionReadStatus.Available, Complete(fixture.Target,
                    new KeyValuePair<string, string>("name", "Application")),
                diagnosticEvidence: new[] { "appmoduleid=" + Guid.NewGuid() });

            var detail = Apply(fixture, source, target).Rows.Single().DefinitionDetail;

            CollectionAssert.AreEquivalent(
                ComponentDefinitionContractCatalog.For(ComponentSemanticKinds.AppModule)
                    .ComparableProperties.ToArray(),
                detail.Properties.Select(item => item.PropertyName).ToArray());
            Assert.IsFalse(detail.Properties.Any(item => item.PropertyName == "appmoduleid"));
            StringAssert.StartsWith(detail.SourceEvidence.Single(), "appmoduleid=");
        }

        [TestMethod, TestCategory("Phase2G2")]
        public void LiveDefinitionOperationReusesWhoAmIAndWebResourceBackingRead()
        {
            var solution = MembershipTestData.Solution();
            var objectId = Guid.NewGuid();
            var componentId = Guid.NewGuid();
            var service = MembershipTestData.Service(solution, query =>
            {
                if (query.EntityName == "solution")
                    return MembershipTestData.Rows(MembershipTestData.SolutionRow(solution));
                if (query.EntityName == "solutioncomponent")
                    return MembershipTestData.Rows(ComponentRow(solution, componentId, 61, objectId));
                if (query.EntityName == "webresource")
                    return MembershipTestData.Rows(new Entity("webresource", objectId)
                    {
                        ["webresourceid"] = objectId,
                        ["name"] = "new_/script.js",
                        ["webresourcetype"] = new OptionSetValue(3),
                        ["content"] = "YWJj",
                        ["displayname"] = "Script",
                        ["description"] = "Description"
                    });
                throw new AssertFailedException("Unexpected query: " + query.EntityName);
            });
            var counter = new DataverseRequestCounter();
            var stages = new List<MembershipOperationStage>();

            var result = new DataverseComponentDefinitionOperation().ReadAndResolve(service,
                "Live", solution.UniqueName, CancellationToken.None,
                progress => stages.Add(progress.Stage), counter);

            Assert.AreEqual(ComponentDefinitionReadStatus.Available,
                result.Definitions.Single().Status);
            Assert.AreEqual(1, counter.GetExecuteCount("WhoAmI"));
            Assert.AreEqual(1, counter.GetQueryCount("webresource"));
            Assert.AreEqual(4, counter.TotalRequests);
            CollectionAssert.Contains(stages, MembershipOperationStage.ReadingDefinitions);
        }

        [TestMethod, TestCategory("Phase2G2")]
        public void LiveDefinitionOperationReusesGlobalChoiceCatalog()
        {
            var solution = MembershipTestData.Solution();
            var objectId = Guid.NewGuid();
            var componentId = Guid.NewGuid();
            var service = MembershipTestData.Service(solution, query =>
            {
                if (query.EntityName == "solution")
                    return MembershipTestData.Rows(MembershipTestData.SolutionRow(solution));
                if (query.EntityName == "solutioncomponent")
                    return MembershipTestData.Rows(ComponentRow(solution, componentId, 9, objectId));
                throw new AssertFailedException("Unexpected query: " + query.EntityName);
            });
            service.ExecuteRequest = request =>
            {
                if (request is WhoAmIRequest)
                    return MembershipTestData.WhoAmI(solution.Environment.OrganizationId);
                if (request is RetrieveAllOptionSetsRequest)
                {
                    var response = new RetrieveAllOptionSetsResponse();
                    response.Results["OptionSetMetadata"] = new OptionSetMetadataBase[]
                    {
                        new OptionSetMetadata
                        {
                            MetadataId = objectId,
                            Name = "new_choice",
                            IsGlobal = true
                        }
                    };
                    return response;
                }
                throw new AssertFailedException("Unexpected request: " + request.RequestName);
            };
            var counter = new DataverseRequestCounter();

            var result = new DataverseComponentDefinitionOperation().ReadAndResolve(service,
                "Live", solution.UniqueName, CancellationToken.None, progress => { }, counter);

            Assert.AreEqual(ComponentDefinitionReadStatus.Available,
                result.Definitions.Single().Status);
            Assert.AreEqual(1, counter.GetExecuteCount("WhoAmI"));
            Assert.AreEqual(1, counter.GetExecuteCount("RetrieveAllOptionSets"));
            Assert.AreEqual(4, counter.TotalRequests);
        }

        [TestMethod, TestCategory("Phase2G2"), TestCategory("Phase2G3")]
        public void LiveDefinitionOperationReusesBulkColumnAndRelationshipMetadata()
        {
            var solution = MembershipTestData.Solution();
            var parentId = Guid.NewGuid(); var columnId = Guid.NewGuid(); var relationshipId = Guid.NewGuid();
            var service = MembershipTestData.Service(solution, query =>
            {
                if (query.EntityName == "solution") return MembershipTestData.Rows(MembershipTestData.SolutionRow(solution));
                Assert.AreEqual("solutioncomponent", query.EntityName);
                return MembershipTestData.Rows(ComponentRow(solution, Guid.NewGuid(), 1, parentId),
                    ComponentRow(solution, Guid.NewGuid(), 2, columnId), ComponentRow(solution, Guid.NewGuid(), 10, relationshipId));
            });
            service.ExecuteRequest = request =>
            {
                if (request is WhoAmIRequest) return MembershipTestData.WhoAmI(solution.Environment.OrganizationId);
                ParentMetadataTestData.AssertQuery((RetrieveMetadataChangesRequest)request);
                return ParentMetadataTestData.Response(ParentMetadataTestData.Root(parentId, "account",
                    new[] { ParentMetadataTestData.Column(columnId) },
                    new[] { ParentMetadataTestData.Relationship(relationshipId) }));
            };
            var counter = new DataverseRequestCounter();
            var result = new DataverseComponentDefinitionOperation().ReadAndResolve(service,
                "Live", solution.UniqueName, CancellationToken.None, progress => { }, counter);
            Assert.IsTrue(result.Definitions.All(item => item.Status == ComponentDefinitionReadStatus.Available));
            Assert.AreEqual(1, counter.GetExecuteCount("WhoAmI"));
            Assert.AreEqual(1, counter.GetExecuteCount("RetrieveMetadataChanges"));
            Assert.AreEqual(0, counter.GetExecuteCount("RetrieveAttribute"));
            Assert.AreEqual(0, counter.GetExecuteCount("RetrieveRelationship"));
            Assert.AreEqual(4, counter.TotalRequests);
            Assert.AreEqual(0, service.WriteCalls);
        }

        [TestMethod, TestCategory("Phase2G3")]
        public void LiveDefinitionOperationReusesSiteMapLookupAndIssuesNoWrites()
        {
            var solution = MembershipTestData.Solution();
            var objectId = Guid.NewGuid();
            var componentId = Guid.NewGuid();
            var siteMapQueryCount = 0;
            var service = MembershipTestData.Service(solution, query =>
            {
                if (query.EntityName == "solution")
                    return MembershipTestData.Rows(MembershipTestData.SolutionRow(solution));
                if (query.EntityName == "solutioncomponent")
                    return MembershipTestData.Rows(ComponentRow(solution, componentId, 62, objectId));
                if (query.EntityName == "sitemap")
                {
                    siteMapQueryCount++;
                    CollectionAssert.AreEquivalent(new[] { "sitemapid", "sitemapnameunique",
                        "sitemapname", "sitemapidunique", "isappaware", "sitemapxml",
                        "componentstate", "ismanaged" }, query.ColumnSet.Columns.ToArray());
                    return MembershipTestData.Rows(new Entity("sitemap", objectId)
                    {
                        ["sitemapid"] = objectId,
                        ["sitemapnameunique"] = "new_EDU",
                        ["sitemapname"] = "EDU",
                        ["sitemapidunique"] = Guid.NewGuid(),
                        ["isappaware"] = true,
                        ["sitemapxml"] = "<SiteMap><Area Id='edu' /></SiteMap>",
                        ["componentstate"] = new OptionSetValue(0),
                        ["ismanaged"] = false
                    });
                }
                throw new AssertFailedException("Unexpected query: " + query.EntityName);
            });
            var counter = new DataverseRequestCounter();

            var result = new DataverseComponentDefinitionOperation().ReadAndResolve(service,
                "Live", solution.UniqueName, CancellationToken.None, progress => { }, counter);

            var definition = result.Definitions.Single();
            Assert.AreEqual(IdentityResolutionStatus.Resolved, definition.Identity.Status);
            Assert.AreEqual(ComponentSemanticKinds.SiteMap, definition.Identity.SemanticKind);
            Assert.AreEqual("new_EDU", definition.Identity.ComparisonKey);
            Assert.AreEqual(ComponentDefinitionReadStatus.Available, definition.Status);
            Assert.AreEqual("<SiteMap><Area Id='edu' /></SiteMap>",
                definition.ComparableProperties["sitemapxml"]);
            Assert.IsTrue(definition.DiagnosticEvidence.Any(item => item.Contains("sitemapid=")));
            Assert.AreEqual(1, siteMapQueryCount);
            Assert.AreEqual(1, counter.GetQueryCount("sitemap"));
            Assert.AreEqual(1, counter.GetExecuteCount("WhoAmI"));
            Assert.AreEqual(4, counter.TotalRequests);
            Assert.AreEqual(0, service.WriteCalls);
        }

        [TestMethod, TestCategory("Phase2G3")]
        public void SiteMapDefinitionMatchAndDifferenceFlowThroughExistingUiPresentation()
        {
            var fixture = Fixture(ComponentSemanticKinds.SiteMap, "new_EDU");
            var source = new ComponentDefinition(fixture.Source,
                ComponentDefinitionReadStatus.Available, Complete(fixture.Source,
                    new KeyValuePair<string, string>("sitemapname", "EDU"),
                    new KeyValuePair<string, string>("isappaware", "True"),
                    new KeyValuePair<string, string>("sitemapxml", "<SiteMap><Area Id='source' /></SiteMap>")),
                diagnosticEvidence: new[] { "sitemapid=" + Guid.NewGuid(), "ismanaged=False" });
            var equalTarget = new ComponentDefinition(fixture.Target,
                ComponentDefinitionReadStatus.Available, Complete(fixture.Target,
                    new KeyValuePair<string, string>("sitemapname", "EDU"),
                    new KeyValuePair<string, string>("isappaware", "True"),
                    new KeyValuePair<string, string>("sitemapxml", "<SiteMap><Area Id='source' /></SiteMap>")),
                diagnosticEvidence: new[] { "sitemapid=" + Guid.NewGuid(), "ismanaged=True" });

            var matched = Apply(fixture, source, equalTarget).Rows.Single();
            Assert.AreEqual("Site Map", matched.ComponentKind);
            Assert.AreEqual("Present in Both", matched.MembershipStatus);
            Assert.AreEqual("Match", matched.DefinitionStatus);
            Assert.AreEqual(string.Empty, matched.ChangedProperties);

            var changedTarget = new ComponentDefinition(fixture.Target,
                ComponentDefinitionReadStatus.Available, Complete(fixture.Target,
                    new KeyValuePair<string, string>("sitemapname", "EDU"),
                    new KeyValuePair<string, string>("isappaware", "True"),
                    new KeyValuePair<string, string>("sitemapxml", "<SiteMap><Area Id='target' /></SiteMap>")),
                diagnosticEvidence: new[] { "sitemapid=" + Guid.NewGuid(), "ismanaged=True" });
            var different = Apply(fixture, source, changedTarget).Rows.Single();
            Assert.AreEqual("Different", different.DefinitionStatus);
            Assert.AreEqual("sitemapxml", different.ChangedProperties);
            var property = different.DefinitionDetail.Properties.Single(item =>
                item.PropertyName == "sitemapxml");
            Assert.IsTrue(property.Changed);
            Assert.AreEqual("<SiteMap><Area Id='source' /></SiteMap>", property.SourceValue);
            Assert.AreEqual("<SiteMap><Area Id='target' /></SiteMap>", property.TargetValue);
            Assert.IsTrue(different.DefinitionDetail.SourceEvidence.Any(item => item.Contains("ismanaged=False")));
            Assert.IsTrue(different.DefinitionDetail.TargetEvidence.Any(item => item.Contains("ismanaged=True")));
        }

        [TestMethod, TestCategory("Phase2G2")]
        public void LiveDefinitionOperationCancellationReturnsNoPartialResult()
        {
            var solution = MembershipTestData.Solution();
            var service = MembershipTestData.Service(solution, query => MembershipTestData.Rows());
            var counter = new DataverseRequestCounter();
            var cancellation = new CancellationTokenSource();
            cancellation.Cancel();

            Assert.ThrowsException<OperationCanceledException>(() =>
                new DataverseComponentDefinitionOperation().ReadAndResolve(service, "Live",
                    solution.UniqueName, cancellation.Token, progress => { }, counter));
            Assert.AreEqual(0, counter.TotalRequests);
        }

        private static MembershipComparisonPresentation Apply(FixtureData fixture,
            ComponentDefinition source, ComponentDefinition target)
        {
            var membership = Present(fixture.SourceSnapshot, fixture.TargetSnapshot);
            return new ComponentDefinitionResultPresenter().Apply(membership,
                Definitions(fixture.SourceSnapshot, source),
                Definitions(fixture.TargetSnapshot, target));
        }

        private static MembershipComparisonPresentation Present(MembershipSnapshot source,
            MembershipSnapshot target) => new MembershipResultPresenter().Create(
                MembershipEnvironmentResult.FromSnapshot("Source", source, 1, TimeSpan.Zero),
                MembershipEnvironmentResult.FromSnapshot("Target", target, 1, TimeSpan.Zero));

        private static ComponentDefinition Available(ComponentIdentity identity,
            string property, string value) => Available(identity,
                new KeyValuePair<string, string>(property, value));

        private static ComponentDefinition Available(ComponentIdentity identity,
            params KeyValuePair<string, string>[] overrides) => new ComponentDefinition(identity,
                ComponentDefinitionReadStatus.Available, Complete(identity, overrides));

        private static IEnumerable<KeyValuePair<string, string>> Complete(ComponentIdentity identity,
            params KeyValuePair<string, string>[] overrides)
        {
            var values = overrides.ToDictionary(item => item.Key, item => item.Value,
                StringComparer.OrdinalIgnoreCase);
            return ComponentDefinitionContractCatalog.For(identity.SemanticKind).ComparableProperties
                .Select(item => new KeyValuePair<string, string>(item,
                    values.ContainsKey(item) ? values[item] : null));
        }

        private static ComponentDefinitionSnapshot Definitions(MembershipSnapshot snapshot,
            params ComponentDefinition[] definitions) => new ComponentDefinitionSnapshot(snapshot,
                definitions.Where(item => item != null));

        private static FixtureData Fixture(string kind, string key)
        {
            var type = kind == ComponentSemanticKinds.Table ? 1 :
                kind == ComponentSemanticKinds.Column ? 2 :
                kind == ComponentSemanticKinds.AppModule ? 80 :
                kind == ComponentSemanticKinds.SiteMap ? 62 : 61;
            var source = Identity(kind, key, IdentityResolutionStatus.Resolved, type);
            var target = Identity(kind, key.ToUpperInvariant(), IdentityResolutionStatus.Resolved, type);
            return new FixtureData(source, target, Snapshot("Source", source), Snapshot("Target", target));
        }

        private static ComponentIdentity Identity(string kind, string key,
            IdentityResolutionStatus status, int type) => new ComponentIdentity(
                new SolutionComponentRecord(Guid.NewGuid(), type, Guid.NewGuid()), status,
                status == IdentityResolutionStatus.Resolved ? key : null,
                diagnostic: status == IdentityResolutionStatus.Resolved ? null : "Identity unavailable.",
                componentTypeKey: kind, semanticKind: kind);

        private static MembershipSnapshot Snapshot(string environment,
            params ComponentIdentity[] identities)
        {
            var solution = new SolutionIdentity(new EnvironmentIdentity(Guid.NewGuid(), environment),
                Guid.NewGuid(), "sample");
            return MembershipSnapshot.Complete(solution, identities, DateTimeOffset.UtcNow);
        }

        private static MembershipSnapshot EmptySnapshot(MembershipSnapshot template,
            string environment = null)
        {
            var identity = new EnvironmentIdentity(Guid.NewGuid(), environment ??
                template.Environment.DisplayName);
            return MembershipSnapshot.Complete(new SolutionIdentity(identity, Guid.NewGuid(),
                template.SolutionUniqueName), new ComponentIdentity[0], DateTimeOffset.UtcNow);
        }

        private static Entity ComponentRow(SolutionIdentity solution, Guid componentId,
            int type, Guid objectId) => new Entity("solutioncomponent", componentId)
            {
                ["solutionid"] = new EntityReference("solution", solution.SolutionId),
                ["componenttype"] = new OptionSetValue(type),
                ["objectid"] = objectId,
                ["rootcomponentbehavior"] = new OptionSetValue(0),
                ["ismetadata"] = false
            };

        private static RetrieveMetadataChangesResponse MetadataRows(
            params EntityMetadata[] rows)
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

        private sealed class FixtureData
        {
            public FixtureData(ComponentIdentity source, ComponentIdentity target,
                MembershipSnapshot sourceSnapshot, MembershipSnapshot targetSnapshot)
            {
                Source = source;
                Target = target;
                SourceSnapshot = sourceSnapshot;
                TargetSnapshot = targetSnapshot;
            }

            public ComponentIdentity Source { get; }
            public ComponentIdentity Target { get; }
            public MembershipSnapshot SourceSnapshot { get; }
            public MembershipSnapshot TargetSnapshot { get; }
        }
    }
}
