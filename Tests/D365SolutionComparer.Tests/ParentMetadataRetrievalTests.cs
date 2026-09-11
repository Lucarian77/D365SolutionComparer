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
using Microsoft.Xrm.Sdk.Metadata.Query;
using Microsoft.Xrm.Sdk.Query;

namespace D365SolutionComparer.Tests
{
    internal static class ParentMetadataTestData
    {
        internal static EntityMetadata Root(Guid id, string name, AttributeMetadata[] columns = null,
            OneToManyRelationshipMetadata[] one = null, OneToManyRelationshipMetadata[] many = null,
            ManyToManyRelationshipMetadata[] multiple = null)
        {
            var root = new EntityMetadata { MetadataId = id, LogicalName = name, SchemaName = name };
            typeof(EntityMetadata).GetProperty("Attributes").SetValue(root, columns ?? new AttributeMetadata[0]);
            typeof(EntityMetadata).GetProperty("OneToManyRelationships").SetValue(root, one ?? new OneToManyRelationshipMetadata[0]);
            typeof(EntityMetadata).GetProperty("ManyToOneRelationships").SetValue(root, many ?? new OneToManyRelationshipMetadata[0]);
            typeof(EntityMetadata).GetProperty("ManyToManyRelationships").SetValue(root, multiple ?? new ManyToManyRelationshipMetadata[0]);
            return root;
        }
        internal static AttributeMetadata Column(Guid id, string name = "new_code") =>
            new StringAttributeMetadata { MetadataId = id, LogicalName = name, SchemaName = name, MaxLength = 100 };
        internal static OneToManyRelationshipMetadata Relationship(Guid id, string name = "new_account_contact") =>
            new OneToManyRelationshipMetadata { MetadataId = id, SchemaName = name,
                ReferencedEntity = "account", ReferencedAttribute = "accountid",
                ReferencingEntity = "contact", ReferencingAttribute = "new_accountid" };
        internal static RetrieveMetadataChangesResponse Response(params EntityMetadata[] entities)
        {
            var response = new RetrieveMetadataChangesResponse();
            var rows = new EntityMetadataCollection(); rows.AddRange(entities);
            response.Results["EntityMetadata"] = rows;
            return response;
        }
        internal static ComponentIdentity Raw(int type, Guid id) => new ComponentIdentity(
            new SolutionComponentRecord(Guid.NewGuid(), type, id), IdentityResolutionStatus.Unresolved);
        internal static void AssertQuery(RetrieveMetadataChangesRequest request)
        {
            Assert.IsNull(request.ClientVersionStamp);
            var query = request.Query;
            Assert.IsTrue(query.Criteria.Conditions.Count > 0 && query.Criteria.Conditions.Count <= 200);
            Assert.AreEqual(LogicalOperator.Or, query.Criteria.FilterOperator);
            Assert.IsTrue(query.Criteria.Conditions.All(item =>
                item.ConditionOperator == MetadataConditionOperator.Equals &&
                (item.PropertyName == "MetadataId" && item.Value.GetType() == typeof(Guid) ||
                 item.PropertyName == "LogicalName" && item.Value.GetType() == typeof(string))));
            Assert.IsNotNull(query.AttributeQuery); Assert.IsNotNull(query.RelationshipQuery);
            Assert.AreEqual(0, query.AttributeQuery.Criteria.Conditions.Count);
            Assert.AreEqual(0, query.AttributeQuery.Criteria.Filters.Count);
            Assert.AreEqual(0, query.RelationshipQuery.Criteria.Conditions.Count);
            Assert.AreEqual(0, query.RelationshipQuery.Criteria.Filters.Count);
            foreach (var name in ParentEntityMetadataReader.Collections)
                CollectionAssert.Contains(query.Properties.PropertyNames, name);
            // Flattened definition labels are not legal SDK metadata property names.
            CollectionAssert.Contains(query.AttributeQuery.Properties.PropertyNames, "Format");
            CollectionAssert.Contains(query.AttributeQuery.Properties.PropertyNames, "OptionSet");
            CollectionAssert.Contains(query.RelationshipQuery.Properties.PropertyNames, "CascadeConfiguration");
            Assert.IsFalse(query.AttributeQuery.Properties.PropertyNames.Contains("IntegerFormat"));
            Assert.IsFalse(query.RelationshipQuery.Properties.PropertyNames.Contains("CascadeDelete"));
        }
    }

    [TestClass]
    public class ParentMetadataRetrievalTests
    {
        [DataTestMethod, TestCategory("Phase2G3")]
        [DataRow(438, 59, 1, 1)]
        [DataRow(382, 55, 1, 1)]
        [DataRow(438, 59, 201, 2)]
        public void EduScaleDependsOnParentBatchesAndDefinitionsNeverReadAgain(
            int columnCount, int relationshipCount, int parentCount, int expectedBatches)
        {
            var parents = Enumerable.Range(1, parentCount).Select(i => new Guid(i, 0, 0, new byte[8])).ToArray();
            var columns = Enumerable.Range(0, columnCount).Select(i => ParentMetadataTestData.Column(Guid.NewGuid(), "new_column" + i)).ToArray();
            var relationships = Enumerable.Range(0, relationshipCount).Select(i => ParentMetadataTestData.Relationship(Guid.NewGuid(), "new_relationship" + i)).ToArray();
            var roots = parents.Select((id, i) => ParentMetadataTestData.Root(id, "table" + i,
                i == 0 ? columns : null, i == 0 ? relationships : null)).ToDictionary(item => item.MetadataId.Value);
            var records = parents.Select(id => ParentMetadataTestData.Raw(1, id))
                .Concat(columns.Select(item => ParentMetadataTestData.Raw(2, item.MetadataId.Value)))
                .Concat(relationships.Select(item => ParentMetadataTestData.Raw(10, item.MetadataId.Value))).ToArray();
            var f = new Fixture(records);
            var queried = new List<Guid>();
            f.Service.ExecuteRequest = request =>
            {
                if (request is WhoAmIRequest) return MembershipTestData.WhoAmI(f.Solution.Environment.OrganizationId);
                var metadata = (RetrieveMetadataChangesRequest)request;
                ParentMetadataTestData.AssertQuery(metadata);
                var ids = metadata.Query.Criteria.Conditions.Select(item => (Guid)item.Value).ToArray();
                queried.AddRange(ids);
                return ParentMetadataTestData.Response(ids.Select(id => roots[id]).ToArray());
            };
            var resolved = f.Resolve();
            Assert.AreEqual(records.Length, resolved.Components.Count);
            Assert.IsTrue(resolved.Components.All(item => item.Status == IdentityResolutionStatus.Resolved));
            Assert.AreEqual(expectedBatches, f.Counter.GetExecuteCount("RetrieveMetadataChanges"));
            CollectionAssert.AreEqual(parents.OrderBy(id => id).ToArray(), queried.ToArray());
            var before = f.Counter.TotalRequests;
            var definitions = f.Definitions(resolved);
            Assert.IsTrue(definitions.Definitions.All(item => item.Status == ComponentDefinitionReadStatus.Available));
            Assert.AreEqual(before, f.Counter.TotalRequests);
            Assert.AreEqual(1, f.Counter.GetExecuteCount("WhoAmI"));
            Assert.AreEqual(0, f.Counter.GetExecuteCount("RetrieveAttribute"));
            Assert.AreEqual(0, f.Counter.GetExecuteCount("RetrieveRelationship"));
            Assert.AreEqual(0, f.Service.WriteCalls);
        }

        [DataTestMethod, TestCategory("Phase2G3")]
        [DataRow("OneToMany")]
        [DataRow("ManyToOne")]
        [DataRow("ManyToMany")]
        public void EachRelationshipCollectionCorrelates(string collection)
        {
            var f = new Fixture();
            var id = f.RelationshipId;
            var root = ParentMetadataTestData.Root(f.ParentId, "account",
                new[] { ParentMetadataTestData.Column(f.ColumnId) },
                collection == "OneToMany" ? new[] { ParentMetadataTestData.Relationship(id) } : null,
                collection == "ManyToOne" ? new[] { ParentMetadataTestData.Relationship(id) } : null,
                collection == "ManyToMany" ? new[] { new ManyToManyRelationshipMetadata
                    { MetadataId = id, SchemaName = "new_account_contact", Entity1LogicalName = "account",
                      Entity2LogicalName = "contact", IntersectEntityName = "new_intersection" } } : null);
            f.Return(root);
            var resolved = f.Resolve();
            Assert.AreEqual("account.new_code", resolved.Components.Single(item => item.Record.ComponentType == 2).ComparisonKey);
            Assert.AreEqual("new_account_contact", resolved.Components.Single(item => item.Record.ComponentType == 10).ComparisonKey);
            Assert.IsTrue(f.Definitions(resolved).Definitions.All(item => item.Status == ComponentDefinitionReadStatus.Available));
        }

        [DataTestMethod, TestCategory("Phase2G3")]
        [DataRow(2)] [DataRow(10)]
        public void MissingCorrelationIsUnresolvedAndCannotMatch(int type)
        {
            var f = new Fixture();
            f.Return(ParentMetadataTestData.Root(f.ParentId, "account"));
            var snapshot = f.Resolve();
            var identity = snapshot.Components.Single(item => item.Record.ComponentType == type);
            Assert.AreEqual(IdentityResolutionStatus.Unresolved, identity.Status);
            Assert.IsNull(identity.ComparisonKey);
            Assert.IsTrue(f.Definitions(snapshot).Definitions.Where(item => item.Identity.Record.ComponentType == type)
                .All(item => item.Status == ComponentDefinitionReadStatus.Unresolved));
            Assert.AreEqual(MembershipPresence.Indeterminate, new SolutionMembershipComparer().Compare(snapshot,
                MembershipSnapshot.Complete(f.Solution, new ComponentIdentity[0], DateTimeOffset.UtcNow))
                .Single(item => item.Source == identity).Presence);
        }

        [DataTestMethod, TestCategory("Phase2G3")]
        [DataRow(2)] [DataRow(10)]
        public void DuplicateOrConflictingCorrelationIsAmbiguous(int type)
        {
            var f = new Fixture();
            var root = f.Root();
            if (type == 2) typeof(EntityMetadata).GetProperty("Attributes").SetValue(root,
                new[] { ParentMetadataTestData.Column(f.ColumnId), ParentMetadataTestData.Column(f.ColumnId) });
            else typeof(EntityMetadata).GetProperty("ManyToOneRelationships").SetValue(root,
                new[] { ParentMetadataTestData.Relationship(f.RelationshipId, "different_name") });
            f.Return(root);
            var resolved = f.Resolve();
            var identity = resolved.Components.Single(item => item.Record.ComponentType == type);
            Assert.AreEqual(IdentityResolutionStatus.Ambiguous, identity.Status);
            Assert.IsNull(identity.ComparisonKey);
            Assert.AreEqual(ComponentDefinitionReadStatus.Ambiguous,
                f.Definitions(resolved).Definitions.Single(item => item.Identity == identity).Status);
        }

        [TestMethod, TestCategory("Phase2G3")]
        public void EquivalentRelationshipFromBothEndpointsCollapsesButConfigurationConflictDoesNot()
        {
            foreach (bool conflict in new[] { false, true })
            {
                var f = new Fixture(); var root = f.Root();
                var copy = ParentMetadataTestData.Relationship(f.RelationshipId);
                if (conflict) copy.ReferencingAttribute = "different_configuration";
                typeof(EntityMetadata).GetProperty("ManyToOneRelationships").SetValue(root, new[] { copy });
                f.Return(root); var resolved = f.Resolve();
                Assert.AreEqual(conflict ? IdentityResolutionStatus.Ambiguous : IdentityResolutionStatus.Resolved,
                    resolved.Components.Single(item => item.Record.ComponentType == 10).Status);
            }
        }

        [DataTestMethod, TestCategory("Phase2G3")]
        [DataRow("Attributes")] [DataRow("OneToManyRelationships")]
        [DataRow("ManyToOneRelationships")] [DataRow("ManyToManyRelationships")]
        public void MissingRequestedCollectionCannotProduceMatch(string property)
        {
            var f = new Fixture(); var root = f.Root();
            typeof(EntityMetadata).GetProperty(property).SetValue(root, null);
            f.Return(root); var resolved = f.Resolve();
            Assert.IsTrue(resolved.Components.Where(item => item.Record.ComponentType != 1)
                .All(item => item.Status == IdentityResolutionStatus.Unresolved));
        }

        [TestMethod, TestCategory("Phase2G3")]
        public void NoParentScopeDoesNotGuessOrIssueDiscoveryRequests()
        {
            var f = new Fixture(new[] { ParentMetadataTestData.Raw(2, Guid.NewGuid()),
                ParentMetadataTestData.Raw(10, Guid.NewGuid()) });
            var resolved = f.Resolve();
            Assert.IsTrue(resolved.Components.All(item => item.Status == IdentityResolutionStatus.Unresolved));
            Assert.AreEqual(1, f.Counter.TotalRequests);
            Assert.IsTrue(resolved.Components.All(item => item.DiagnosticEvidence.Single().Contains("No verified parent")));
        }

        [TestMethod, TestCategory("Phase2G3")]
        public void FaultOnLaterParentBatchDiscardsEarlierChildResultsAndDoesNotRetryDefinitions()
        {
            var parentIds = Enumerable.Range(1, 201).Select(i => new Guid(i, 0, 0, new byte[8])).ToArray();
            var columnId = Guid.NewGuid();
            var f = new Fixture(parentIds.Select(id => ParentMetadataTestData.Raw(1, id))
                .Concat(new[] { ParentMetadataTestData.Raw(2, columnId) }).ToArray());
            f.Service.ExecuteRequest = request =>
            {
                if (request is WhoAmIRequest) return MembershipTestData.WhoAmI(f.Solution.Environment.OrganizationId);
                var query = ((RetrieveMetadataChangesRequest)request).Query;
                if (query.AttributeQuery == null) return ParentMetadataTestData.Response(query.Criteria.Conditions
                    .Select(c => ParentMetadataTestData.Root((Guid)c.Value, "table" + c.Value)).ToArray());
                if (f.Counter.GetExecuteCount("RetrieveMetadataChanges") == 2) throw new FaultException("EDU denied");
                return ParentMetadataTestData.Response(query.Criteria.Conditions.Select((c, i) =>
                    ParentMetadataTestData.Root((Guid)c.Value, "table" + c.Value,
                        i == 0 ? new[] { ParentMetadataTestData.Column(columnId) } : null)).ToArray());
            };
            var resolved = f.Resolve();
            var column = resolved.Components.Single(item => item.Record.ComponentType == 2);
            Assert.AreEqual(IdentityResolutionStatus.Unresolved, column.Status);
            StringAssert.Contains(column.DiagnosticEvidence.Single(), "batch 2; parent entities=1");
            StringAssert.Contains(column.DiagnosticEvidence.Single(), "EDU denied");
            StringAssert.Contains(column.DiagnosticEvidence.Single(), "Attributes, OneToManyRelationships, ManyToOneRelationships, ManyToManyRelationships");
            var before = f.Counter.TotalRequests;
            Assert.IsFalse(f.Definitions(resolved).Definitions.Where(item => item.Identity == column)
                .Any(item => item.Status == ComponentDefinitionReadStatus.Available));
            Assert.AreEqual(before, f.Counter.TotalRequests);
        }

        [TestMethod, TestCategory("Phase2G3")]
        public void CancellationDuringParentReadProducesNoSnapshot()
        {
            var f = new Fixture();
            using (var cancellation = new CancellationTokenSource())
            {
                f.Service.ExecuteRequest = request =>
                {
                    if (request is WhoAmIRequest) return MembershipTestData.WhoAmI(f.Solution.Environment.OrganizationId);
                    cancellation.Cancel(); return ParentMetadataTestData.Response(f.Root());
                };
                Assert.ThrowsException<OperationCanceledException>(() =>
                    new DataverseComponentIdentityResolver().ResolveSnapshot(f.Service, f.Snapshot,
                        cancellation.Token, f.Counter));
            }
        }

        [TestMethod, TestCategory("Phase2G3")]
        public void MetadataGuidCollisionAcrossFamiliesDoesNotMixCacheEntries()
        {
            var id = Guid.NewGuid(); var parent = Guid.NewGuid();
            var f = new Fixture(new[] { ParentMetadataTestData.Raw(1, parent),
                ParentMetadataTestData.Raw(2, id), ParentMetadataTestData.Raw(10, id) });
            f.Return(ParentMetadataTestData.Root(parent, "account", new[] { ParentMetadataTestData.Column(id) },
                new[] { ParentMetadataTestData.Relationship(id) }));
            var resolved = f.Resolve();
            var definitions = f.Definitions(resolved);
            Assert.AreEqual("100", definitions.Definitions.Single(item => item.Identity.Record.ComponentType == 2)
                .ComparableProperties["MaxLength"]);
            Assert.AreEqual("contact", definitions.Definitions.Single(item => item.Identity.Record.ComponentType == 10)
                .ComparableProperties["ReferencingEntity"]);
            Assert.AreEqual(2, f.Counter.TotalRequests);
        }

        [DataTestMethod, TestCategory("Phase2G3")]
        [DataRow(true)] [DataRow(false)]
        public void FakeRejectsTheLiveIncompatibleChildIdFilter(bool column)
        {
            var service = new FakeOrganizationService { ExecuteRequest = request => new OrganizationResponse() };
            var criteria = new MetadataFilterExpression(LogicalOperator.Or);
            criteria.Conditions.Add(new MetadataConditionExpression("MetadataId", MetadataConditionOperator.Equals, Guid.NewGuid()));
            var query = new EntityQueryExpression();
            if (column) query.AttributeQuery = new AttributeQueryExpression { Criteria = criteria };
            else query.RelationshipQuery = new RelationshipQueryExpression { Criteria = criteria };
            var error = Assert.ThrowsException<FaultException>(() => service.Execute(new RetrieveMetadataChangesRequest { Query = query }));
            StringAssert.Contains(error.Message, "Unable to evaluate query");
        }

        [DataTestMethod, TestCategory("Phase2G3")]
        [DataRow("MissingParent")] [DataRow("DuplicateParent")] [DataRow("UnexpectedParent")]
        [DataRow("NullCollection")] [DataRow("NullChild")] [DataRow("MissingChildId")]
        public void MalformedParentInventoryNeverProducesAvailableChildDefinitions(string condition)
        {
            var f = new Fixture(); var root = f.Root();
            f.Service.ExecuteRequest = request =>
            {
                if (request is WhoAmIRequest) return MembershipTestData.WhoAmI(f.Solution.Environment.OrganizationId);
                ParentMetadataTestData.AssertQuery((RetrieveMetadataChangesRequest)request);
                if (condition == "MissingParent") return ParentMetadataTestData.Response();
                if (condition == "DuplicateParent") return ParentMetadataTestData.Response(root, root);
                if (condition == "UnexpectedParent") root.MetadataId = Guid.NewGuid();
                if (condition == "NullCollection")
                {
                    var response = new RetrieveMetadataChangesResponse();
                    response.Results["EntityMetadata"] = null;
                    return response;
                }
                if (condition == "NullChild") typeof(EntityMetadata).GetProperty("Attributes")
                    .SetValue(root, new AttributeMetadata[] { null });
                if (condition == "MissingChildId") root.Attributes[0].MetadataId = null;
                return ParentMetadataTestData.Response(root);
            };
            var resolved = f.Resolve();
            Assert.IsTrue(resolved.Components.Where(item => item.Record.ComponentType != 1)
                .All(item => item.Status == IdentityResolutionStatus.Unresolved));
            Assert.IsTrue(f.Definitions(resolved).Definitions.Where(item => item.Identity.Record.ComponentType != 1)
                .All(item => item.Status == ComponentDefinitionReadStatus.Unresolved));
            Assert.AreEqual(1, f.Counter.GetExecuteCount("RetrieveMetadataChanges"));
        }

        [TestMethod, TestCategory("Phase2G3")]
        public void ConflictingColumnParentAndBlankNamesRemainUnresolved()
        {
            foreach (bool blank in new[] { false, true })
            {
                var f = new Fixture(); var root = f.Root();
                if (blank) root.Attributes[0].LogicalName = " ";
                else typeof(AttributeMetadata).GetProperty("EntityLogicalName")
                    .SetValue(root.Attributes[0], "conflicting_table");
                root.OneToManyRelationships[0].SchemaName = " ";
                f.Return(root); var resolved = f.Resolve();
                Assert.IsTrue(resolved.Components.Where(item => item.Record.ComponentType != 1)
                    .All(item => item.Status == IdentityResolutionStatus.Unresolved));
            }
        }

        [TestMethod, TestCategory("Phase2G3")]
        public void DefinitionFailureEvidenceIsPreservedWithoutRetryOrMembershipMutation()
        {
            var f = new Fixture();
            f.Service.ExecuteRequest = request => request is WhoAmIRequest
                ? (OrganizationResponse)MembershipTestData.WhoAmI(f.Solution.Environment.OrganizationId)
                : throw new FaultException("metadata permission denied");
            var resolved = f.Resolve();
            var before = new SolutionMembershipComparer().Compare(resolved, resolved);
            var count = f.Counter.TotalRequests;
            var definitions = f.Definitions(resolved);
            var comparisons = new ComponentDetailComparer().Compare(before, definitions, definitions);
            Assert.IsTrue(comparisons.All(item => item.Status == ComponentDetailComparisonStatus.Unresolved));
            Assert.IsTrue(before.All(item => item.Presence == MembershipPresence.Indeterminate));
            Assert.IsTrue(definitions.Definitions.All(item => item.DiagnosticEvidence
                .Any(evidence => evidence.Contains("metadata permission denied"))));
            Assert.AreEqual(count, f.Counter.TotalRequests);
            Assert.AreEqual(3, resolved.Components.Count);
            Assert.AreEqual(0, f.Service.WriteCalls);
        }

        [TestMethod, TestCategory("Phase2G3")]
        public void ChildProjectionNamesAreActualSdkPropertiesAndEqualityContractsAreUnchanged()
        {
            var attributeNames = new[] { typeof(AttributeMetadata), typeof(StringAttributeMetadata),
                typeof(MemoAttributeMetadata), typeof(IntegerAttributeMetadata), typeof(DecimalAttributeMetadata),
                typeof(MoneyAttributeMetadata), typeof(DateTimeAttributeMetadata), typeof(LookupAttributeMetadata),
                typeof(EnumAttributeMetadata) }.SelectMany(type => type.GetProperties()).Select(item => item.Name).ToArray();
            var relationshipNames = new[] { typeof(RelationshipMetadataBase), typeof(OneToManyRelationshipMetadata),
                typeof(ManyToManyRelationshipMetadata) }.SelectMany(type => type.GetProperties()).Select(item => item.Name).ToArray();
            Assert.IsTrue(ParentEntityMetadataReader.AttributeProperties.All(name => attributeNames.Contains(name)));
            Assert.IsTrue(ParentEntityMetadataReader.RelationshipProperties.All(name => relationshipNames.Contains(name)));
            CollectionAssert.Contains(ComponentDefinitionContractCatalog.For(ComponentSemanticKinds.Column)
                .ComparableProperties.ToArray(), "IntegerFormat");
            CollectionAssert.Contains(ComponentDefinitionContractCatalog.For(ComponentSemanticKinds.Relationship)
                .ComparableProperties.ToArray(), "CascadeDelete");
        }

        private sealed class Fixture
        {
            internal readonly Guid ParentId = Guid.NewGuid(), ColumnId = Guid.NewGuid(), RelationshipId = Guid.NewGuid();
            internal readonly SolutionIdentity Solution = MembershipTestData.Solution();
            internal readonly DataverseRequestCounter Counter = new DataverseRequestCounter();
            internal readonly MembershipSnapshot Snapshot;
            internal readonly FakeOrganizationService Service;
            private DataverseReadContext context;
            internal Fixture(ComponentIdentity[] records = null)
            {
                Snapshot = MembershipSnapshot.Complete(Solution, records ?? new[] {
                    ParentMetadataTestData.Raw(1, ParentId), ParentMetadataTestData.Raw(2, ColumnId),
                    ParentMetadataTestData.Raw(10, RelationshipId) }, DateTimeOffset.UtcNow);
                Service = MembershipTestData.Service(Solution);
                Return(Root());
            }
            internal EntityMetadata Root() => ParentMetadataTestData.Root(ParentId, "account",
                new[] { ParentMetadataTestData.Column(ColumnId) },
                new[] { ParentMetadataTestData.Relationship(RelationshipId) });
            internal void Return(EntityMetadata root) => Service.ExecuteRequest = request =>
            {
                if (request is WhoAmIRequest) return MembershipTestData.WhoAmI(Solution.Environment.OrganizationId);
                ParentMetadataTestData.AssertQuery((RetrieveMetadataChangesRequest)request);
                return ParentMetadataTestData.Response(root);
            };
            internal MembershipSnapshot Resolve()
            {
                context = new DataverseReadContext(Service, Solution.Environment, CancellationToken.None, Counter);
                return new DataverseComponentIdentityResolver().ResolveSnapshot(context, Snapshot, CancellationToken.None);
            }
            internal ComponentDefinitionSnapshot Definitions(MembershipSnapshot resolved) =>
                new DataverseComponentDefinitionReader().Read(context, resolved, CancellationToken.None);
        }
    }
}
