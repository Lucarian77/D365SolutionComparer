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
using Microsoft.Xrm.Sdk.Metadata.Query;

namespace D365SolutionComparer.Tests
{
    [TestClass]
    public class EntityKeyResolutionTests
    {
        [TestMethod, TestCategory("Phase2G7B1")]
        public void GroupedMetadataQueryRequestsKeysWithRequiredPropertiesAndNoPerKeyRequest()
        {
            var fixture = Fixture.Create();
            var seen = new List<RetrieveMetadataChangesRequest>();
            fixture.Service.ExecuteRequest = request =>
            {
                if (request is WhoAmIRequest) return MembershipTestData.WhoAmI(fixture.Solution.Environment.OrganizationId);
                var metadata = request as RetrieveMetadataChangesRequest;
                Assert.IsNotNull(metadata);
                seen.Add(metadata);
                return ParentMetadataTestData.Response(fixture.Root);
            };
            var result = fixture.Resolve(true);
            Assert.AreEqual(IdentityResolutionStatus.Resolved, result.Components[1].Status,
                result.Components[1].Diagnostic);
            Assert.AreEqual("entitykey:v1:29:environmentvariabledefinition:13:definitionkey",
                result.Components[1].ComparisonKey);
            Assert.AreEqual(1, seen.Count);
            var query = seen[0].Query;
            CollectionAssert.Contains(query.Properties.PropertyNames, "Keys");
            Assert.IsNotNull(query.KeyQuery);
            CollectionAssert.AreEquivalent(new[] { "MetadataId", "EntityLogicalName", "LogicalName",
                "SchemaName", "KeyAttributes" }, query.KeyQuery.Properties.PropertyNames.ToArray());
            Assert.AreEqual(0, fixture.Counter.GetExecuteCount("RetrieveEntityKey"));
        }

        [TestMethod, TestCategory("Phase2G7B1")]
        public void CrossEnvironmentGuidsDifferButPortableIdentityMatches()
        {
            var first = Fixture.Create(Guid.NewGuid(), Guid.NewGuid());
            var second = Fixture.Create(Guid.NewGuid(), Guid.NewGuid());
            var one = first.Resolve(true); var two = second.Resolve(true);
            Assert.AreNotEqual(first.KeyId, second.KeyId);
            Assert.IsTrue(StringComparer.OrdinalIgnoreCase.Equals(one.Components[1].ComparisonKey,
                two.Components[1].ComparisonKey));
        }

        [TestMethod, TestCategory("Phase2G7B1")]
        public void BlankNamesAndMissingOrMultipleMetadataRemainConservative()
        {
            var blank = Fixture.Create(); blank.Key.LogicalName = " ";
            Assert.AreEqual(IdentityResolutionStatus.Unresolved, blank.Resolve(true).Components[1].Status);
            var parentBlank = Fixture.Create(); parentBlank.Root.LogicalName = " ";
            Assert.AreEqual(IdentityResolutionStatus.Unresolved, parentBlank.Resolve(true).Components[1].Status);
            var missing = Fixture.Create();
            typeof(EntityMetadata).GetProperty("Keys").SetValue(missing.Root, new EntityKeyMetadata[0]);
            Assert.AreEqual(IdentityResolutionStatus.Unresolved, missing.Resolve(true).Components[1].Status);
            var multiple = Fixture.Create();
            typeof(EntityMetadata).GetProperty("Keys").SetValue(multiple.Root, new[] { multiple.Key,
                Fixture.CreateKey(multiple.KeyId, "environmentvariabledefinition", "definitionkey", "DefinitionKey2",
                    new[] { "schemaname" }) });
            Assert.AreEqual(IdentityResolutionStatus.Ambiguous, multiple.Resolve(true).Components[1].Status);
        }

        [TestMethod, TestCategory("Phase2G7B1")]
        public void DuplicateCanonicalIdentityIsAmbiguousButRepeatedRawObjectIsNot()
        {
            var duplicate = Fixture.Create();
            var secondKey = Fixture.CreateKey(Guid.NewGuid(), "environmentvariabledefinition", "definitionkey",
                "DefinitionKey2", new[] { "schemaname" });
            typeof(EntityMetadata).GetProperty("Keys").SetValue(duplicate.Root,
                new[] { duplicate.Key, secondKey });
            duplicate.Records = duplicate.Records.Concat(new[] { Raw(14, secondKey.MetadataId.Value,
                duplicate.RootRecordId) }).ToArray();
            var ambiguous = duplicate.Resolve(true).Components.Where(item => item.Record.ComponentType == 14).ToList();
            Assert.IsTrue(ambiguous.All(item => item.Status == IdentityResolutionStatus.Ambiguous));
            Assert.AreEqual(1, duplicate.Counter.GetExecuteCount("RetrieveMetadataChanges"));

            var repeated = Fixture.Create();
            repeated.Records = repeated.Records.Concat(new[] { Raw(14, repeated.KeyId,
                repeated.RootRecordId) }).ToArray();
            var resolved = repeated.Resolve(true).Components.Where(item => item.Record.ComponentType == 14).ToList();
            Assert.IsTrue(resolved.All(item => item.Status == IdentityResolutionStatus.Resolved));
            Assert.AreEqual(1, repeated.Counter.GetExecuteCount("RetrieveMetadataChanges"));
        }

        [TestMethod, TestCategory("Phase2G7B1")]
        public void DifferentParentsWithSameKeyNameRemainDifferentPortableIdentities()
        {
            var fixture = Fixture.Create();
            var otherParent = Guid.NewGuid(); var otherKey = Guid.NewGuid();
            var otherRootRecordId = Guid.NewGuid();
            var root2 = ParentMetadataTestData.Root(otherParent, "account", keys: new[] {
                Fixture.CreateKey(otherKey, "account", "definitionkey", "DefinitionKey", new[] { "name" }) });
            fixture.Records = fixture.Records.Concat(new[] {
                new ComponentIdentity(new SolutionComponentRecord(otherRootRecordId, 1, otherParent), IdentityResolutionStatus.Unresolved),
                Raw(14, otherKey, otherRootRecordId)
            }).ToArray();
            fixture.ExtraRoots = new[] { root2 };
            var result = fixture.Resolve(true);
            var keys = result.Components.Where(item => item.Record.ComponentType == 14)
                .Select(item => item.ComparisonKey).ToList();
            Assert.AreEqual(2, keys.Distinct(StringComparer.OrdinalIgnoreCase).Count());
        }

        [TestMethod, TestCategory("Phase2G7B1")]
        public void KeyAttributesAreOrderAndCaseInsensitiveAndDefinitionDifferencesRemainVisible()
        {
            var a = Fixture.Create(); var b = Fixture.Create();
            a.Key.KeyAttributes = new[] { "Name", "Code", "Code" };
            b.Key.KeyAttributes = new[] { "code", "name" };
            var da = a.Definitions(a.Resolve(true)); var db = b.Definitions(b.Resolve(true));
            Assert.AreEqual(da.Definitions.Single(item => item.Identity.Record.ComponentType == 14)
                .ComparableProperties["KeyAttributes"],
                db.Definitions.Single(item => item.Identity.Record.ComponentType == 14)
                .ComparableProperties["KeyAttributes"]);
            b.Key.KeyAttributes = new[] { "other" };
            var changed = b.Definitions(b.Resolve(true)).Definitions.Single(item => item.Identity.Record.ComponentType == 14);
            Assert.AreNotEqual("code,name", changed.ComparableProperties["KeyAttributes"]);
        }

        [TestMethod, TestCategory("Phase2G7B1")]
        public void EmptyKeyAttributesAndMetadataFailureRemainUnresolved()
        {
            var empty = Fixture.Create(); empty.Key.KeyAttributes = new string[0];
            var resolved = empty.Resolve(true);
            Assert.AreEqual(IdentityResolutionStatus.Resolved, resolved.Components[1].Status);
            Assert.AreEqual(ComponentDefinitionReadStatus.Unresolved,
                empty.Definitions(resolved).Definitions.Single(item => item.Identity.Record.ComponentType == 14).Status);
            var failed = Fixture.Create();
            failed.Service.ExecuteRequest = request => request is WhoAmIRequest
                ? (OrganizationResponse)MembershipTestData.WhoAmI(failed.Solution.Environment.OrganizationId)
                : throw new InvalidOperationException("metadata unavailable");
            Assert.AreEqual(IdentityResolutionStatus.Unresolved, failed.Resolve(true).Components[1].Status);
        }

        [TestMethod, TestCategory("Phase2G7B1")]
        public void DefinitionReadReusesTheOperationScopedEntityKeyMetadataInventory()
        {
            var fixture = Fixture.Create();
            var resolved = fixture.ResolveAndDefinitions(true);
            Assert.AreEqual(ComponentDefinitionReadStatus.Available,
                resolved.Definitions.Single(item => item.Identity.Record.ComponentType == 14).Status);
            Assert.AreEqual(1, fixture.Counter.GetExecuteCount("RetrieveMetadataChanges"));
            Assert.AreEqual(0, fixture.Counter.GetExecuteCount("RetrieveEntityKey"));
        }

        [TestMethod, TestCategory("Phase2G7B2")]
        public void PublicResolverSupportsType14AfterPromotion()
        {
            var fixture = Fixture.Create();
            var snapshot = MembershipSnapshot.Complete(fixture.Solution, fixture.Records, DateTimeOffset.UtcNow);
            var result = new DataverseComponentIdentityResolver().ResolveSnapshot(fixture.Service, snapshot,
                CancellationToken.None, fixture.Counter);
            Assert.AreEqual(IdentityResolutionStatus.Resolved, result.Components[1].Status);
            Assert.AreEqual("entitykey:v1:29:environmentvariabledefinition:13:definitionkey",
                result.Components[1].ComparisonKey);
        }

        [TestMethod, TestCategory("Phase2G7B1")]
        public void MetadataCancellationPropagatesWithoutReturningAResolvedKey()
        {
            var fixture = Fixture.Create();
            using (var cancellation = new CancellationTokenSource())
            {
                fixture.Service.ExecuteRequest = request =>
                {
                    if (request is WhoAmIRequest)
                        return MembershipTestData.WhoAmI(fixture.Solution.Environment.OrganizationId);
                    cancellation.Cancel();
                    return ParentMetadataTestData.Response(fixture.Root);
                };
                Assert.ThrowsException<OperationCanceledException>(() => fixture.ResolveWithToken(true,
                    cancellation.Token));
            }
        }

        [TestMethod, TestCategory("Phase2G7B1")]
        public void MissingKeyCollectionDoesNotInvalidateExistingColumnResolution()
        {
            var fixture = Fixture.Create();
            var columnId = Guid.NewGuid();
            typeof(EntityMetadata).GetProperty("Attributes").SetValue(fixture.Root,
                new[] { ParentMetadataTestData.Column(columnId, "definition") });
            typeof(EntityMetadata).GetProperty("Keys").SetValue(fixture.Root, null);
            fixture.Records = fixture.Records.Concat(new[] {
                new ComponentIdentity(new SolutionComponentRecord(Guid.NewGuid(), 2, columnId,
                    rootSolutionComponentId: fixture.RootRecordId), IdentityResolutionStatus.Unresolved)
            }).ToArray();
            var result = fixture.Resolve(true);
            Assert.AreEqual(IdentityResolutionStatus.Unresolved,
                result.Components.Single(item => item.Record.ComponentType == 14).Status);
            Assert.AreEqual(IdentityResolutionStatus.Resolved,
                result.Components.Single(item => item.Record.ComponentType == 2).Status);
        }

        private static ComponentIdentity Raw(int type, Guid objectId, Guid rootId) =>
            new ComponentIdentity(new SolutionComponentRecord(Guid.NewGuid(), type, objectId,
                rootSolutionComponentId: rootId), IdentityResolutionStatus.Unresolved);

        private sealed class Fixture
        {
            internal readonly SolutionIdentity Solution = MembershipTestData.Solution();
            internal readonly Guid ParentId;
            internal readonly Guid KeyId;
            internal readonly Guid RootRecordId;
            internal readonly EntityMetadata Root;
            internal readonly EntityKeyMetadata Key;
            internal readonly DataverseRequestCounter Counter = new DataverseRequestCounter();
            internal readonly FakeOrganizationService Service;
            internal ComponentIdentity[] Records;
            internal EntityMetadata[] ExtraRoots = new EntityMetadata[0];

            private Fixture(Guid parentId, Guid keyId)
            {
                ParentId = parentId; KeyId = keyId; RootRecordId = Guid.NewGuid();
                Key = CreateKey(keyId, "environmentvariabledefinition", "definitionkey", "DefinitionKey",
                    new[] { "schemaname", "environmentvariabledefinitionid" });
                Root = ParentMetadataTestData.Root(parentId, "environmentvariabledefinition", keys: new[] { Key });
                Records = new[] {
                    new ComponentIdentity(new SolutionComponentRecord(RootRecordId, 1, parentId), IdentityResolutionStatus.Unresolved),
                    Raw(14, keyId, RootRecordId)
                };
                Service = MembershipTestData.Service(Solution);
                Service.ExecuteRequest = request =>
                {
                    if (request is WhoAmIRequest) return MembershipTestData.WhoAmI(Solution.Environment.OrganizationId);
                    var metadata = request as RetrieveMetadataChangesRequest;
                    if (metadata == null) throw new InvalidOperationException("unexpected request");
                    var ids = metadata.Query.Criteria.Conditions.Select(item => (Guid)item.Value).ToList();
                    return ParentMetadataTestData.Response(ids.Select(id => id == ParentId ? Root : null)
                        .Where(item => item != null).Concat(ExtraRoots).ToArray());
                };
            }

            internal static EntityKeyMetadata CreateKey(Guid id, string entity, string logical, string schema,
                string[] attributes)
            {
                var key = new EntityKeyMetadata { MetadataId = id, LogicalName = logical,
                    SchemaName = schema, KeyAttributes = attributes };
                typeof(EntityKeyMetadata).GetProperty("EntityLogicalName").SetValue(key, entity);
                return key;
            }

            internal static Fixture Create(Guid? parent = null, Guid? key = null) =>
                new Fixture(parent ?? Guid.NewGuid(), key ?? Guid.NewGuid());

            internal MembershipSnapshot Resolve(bool enabled)
            {
                var snapshot = MembershipSnapshot.Complete(Solution, Records, DateTimeOffset.UtcNow);
                return new DataverseComponentIdentityResolver(enabled).ResolveSnapshot(Service, snapshot,
                    CancellationToken.None, Counter);
            }

            internal MembershipSnapshot ResolveWithToken(bool enabled, CancellationToken token)
            {
                var snapshot = MembershipSnapshot.Complete(Solution, Records, DateTimeOffset.UtcNow);
                return new DataverseComponentIdentityResolver(enabled).ResolveSnapshot(Service, snapshot,
                    token, Counter);
            }

            internal ComponentDefinitionSnapshot Definitions(MembershipSnapshot snapshot)
            {
                var context = new DataverseReadContext(Service, Solution.Environment, CancellationToken.None, Counter);
                return new DataverseComponentDefinitionReader().Read(context, snapshot, CancellationToken.None);
            }

            internal ComponentDefinitionSnapshot ResolveAndDefinitions(bool enabled)
            {
                var snapshot = MembershipSnapshot.Complete(Solution, Records, DateTimeOffset.UtcNow);
                var context = new DataverseReadContext(Service, Solution.Environment,
                    CancellationToken.None, Counter);
                var resolved = new DataverseComponentIdentityResolver(enabled)
                    .ResolveSnapshot(context, snapshot, CancellationToken.None);
                return new DataverseComponentDefinitionReader().Read(context, resolved,
                    CancellationToken.None);
            }
        }
    }
}
