using System;
using System.Collections.Generic;
using System.Linq;
using D365SolutionComparer.Models;
using D365SolutionComparer.Models.ComponentDetails;
using D365SolutionComparer.Models.Identity;
using D365SolutionComparer.Models.Membership;
using D365SolutionComparer.Services;
using D365SolutionComparer.Services.ComponentDetails;
using D365SolutionComparer.Services.Membership;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace D365SolutionComparer.Tests
{
    [TestClass]
    public class ComponentDetailComparisonTests
    {
        [TestMethod]
        public void IdenticalDefinitionsAreMatch()
        {
            var fixture = Fixture("webresource", "new_/script.js");
            var result = Compare(fixture, Available(fixture.Source, "Content", "YWJj"),
                Available(fixture.Target, "Content", "YWJj"));
            Assert.AreEqual(ComponentDetailComparisonStatus.Match, result.Status);
            Assert.AreEqual(0, result.Differences.Count);
        }

        [TestMethod]
        public void OneDifferingPropertyIsDifferentAndIdentified()
        {
            var fixture = Fixture("webresource", "new_/script.js");
            var result = Compare(fixture, Available(fixture.Source, "Content", "before"),
                Available(fixture.Target, "Content", "after"));
            Assert.AreEqual(ComponentDetailComparisonStatus.Different, result.Status);
            Assert.AreEqual("content", result.Differences.Single().PropertyName);
        }

        [TestMethod]
        public void SourceOnlyAndTargetOnlyComeFromMembershipEvidence()
        {
            var sourceFixture = Fixture("table", "account");
            var sourceMembership = new MembershipCompareResult(sourceFixture.Source, null,
                MembershipPresence.OnlyInSource, MembershipAbsenceEvidence.CompleteResolvedInventory);
            var sourceResult = Compare(sourceFixture, Available(sourceFixture.Source, "SchemaName", "Account"),
                null, sourceMembership);
            Assert.AreEqual(ComponentDetailComparisonStatus.SourceOnly, sourceResult.Status);

            var targetFixture = Fixture("table", "contact");
            var targetMembership = new MembershipCompareResult(null, targetFixture.Target,
                MembershipPresence.OnlyInTarget, MembershipAbsenceEvidence.OppositeSolutionAbsent);
            var targetResult = Compare(targetFixture, null,
                Available(targetFixture.Target, "SchemaName", "Contact"), targetMembership);
            Assert.AreEqual(ComponentDetailComparisonStatus.TargetOnly, targetResult.Status);
        }

        [TestMethod]
        public void EnvironmentLocalGuidDifferencesAreDiagnosticOnly()
        {
            var fixture = Fixture("appmodule", "new_app");
            var source = new ComponentDefinition(fixture.Source, ComponentDefinitionReadStatus.Available,
                Complete(fixture.Source, Pair("navigationtype", "0")),
                diagnosticEvidence: new[] { "appmoduleid=" + Guid.NewGuid() });
            var target = new ComponentDefinition(fixture.Target, ComponentDefinitionReadStatus.Available,
                Complete(fixture.Target, Pair("navigationtype", "0")),
                diagnosticEvidence: new[] { "appmoduleid=" + Guid.NewGuid() });
            Assert.AreNotEqual(fixture.Source.Record.ObjectId, fixture.Target.Record.ObjectId);
            Assert.AreEqual(ComponentDetailComparisonStatus.Match, Compare(fixture, source, target).Status);
        }

        [TestMethod]
        public void EstablishedPortableIdentityMatchingRemainsCaseInsensitive()
        {
            var fixture = Fixture("globalchoice", "new_StatusChoice", "NEW_statuschoice");
            var memberships = new SolutionMembershipComparer().Compare(fixture.SourceSnapshot,
                fixture.TargetSnapshot);
            Assert.AreEqual(MembershipPresence.PresentInBoth, memberships.Single().Presence);
            var result = new ComponentDetailComparer().Compare(memberships,
                Definitions(fixture.SourceSnapshot, Available(fixture.Source, "Options", "1=Active")),
                Definitions(fixture.TargetSnapshot, Available(fixture.Target, "Options", "1=Active"))).Single();
            Assert.AreEqual(ComponentDetailComparisonStatus.Match, result.Status);
        }

        [TestMethod]
        public void IncompleteRetrievalRemainsUnresolved()
        {
            var fixture = Fixture("column", "account.new_code");
            var source = Available(fixture.Source, "RequiredLevel", "None");
            var target = new ComponentDefinition(fixture.Target,
                ComponentDefinitionReadStatus.Unresolved,
                diagnostic: "Column metadata retrieval failed.");
            Assert.AreEqual(ComponentDetailComparisonStatus.Unresolved,
                Compare(fixture, source, target).Status);
        }

        [TestMethod]
        public void UnsupportedAndAmbiguousDefinitionsRemainConservative()
        {
            var fixture = Fixture("relationship", "new_account_contact");
            Assert.AreEqual(ComponentDetailComparisonStatus.Unsupported,
                Compare(fixture, new ComponentDefinition(fixture.Source,
                    ComponentDefinitionReadStatus.Unsupported),
                    Available(fixture.Target, "RelationshipType", "OneToManyRelationship")).Status);
            Assert.AreEqual(ComponentDetailComparisonStatus.Ambiguous,
                Compare(fixture, new ComponentDefinition(fixture.Source,
                    ComponentDefinitionReadStatus.Ambiguous),
                    Available(fixture.Target, "RelationshipType", "OneToManyRelationship")).Status);
        }

        [TestMethod]
        public void DefinitionComparisonDoesNotChangeMembershipCounts()
        {
            var fixture = Fixture("webresource", "new_/script.js");
            var memberships = new SolutionMembershipComparer().Compare(fixture.SourceSnapshot,
                fixture.TargetSnapshot);
            var before = memberships.GroupBy(item => item.Presence)
                .ToDictionary(group => group.Key, group => group.Count());
            new ComponentDetailComparer().Compare(memberships,
                Definitions(fixture.SourceSnapshot, Available(fixture.Source, "Content", "a")),
                Definitions(fixture.TargetSnapshot, Available(fixture.Target, "Content", "b")));
            CollectionAssert.AreEquivalent(before.ToArray(), memberships.GroupBy(item => item.Presence)
                .ToDictionary(group => group.Key, group => group.Count()).ToArray());
        }

        [TestMethod]
        public void ExistingSolutionLevelComparisonRemainsIndependent()
        {
            var source = new SolutionInfo { UniqueName = "Sample", DisplayName = "Name", Version = "1", Publisher = "P", IsManaged = false };
            var target = new SolutionInfo { UniqueName = "sample", DisplayName = "Name", Version = "1", Publisher = "P", IsManaged = false };
            Assert.AreEqual("Match", new SolutionComparisonService().Compare(
                new List<SolutionInfo> { source }, new List<SolutionInfo> { target }).Single().Status);
        }

        [DataTestMethod]
        [DataRow("table", "OwnershipType", "UserOwned", "OrganizationOwned")]
        [DataRow("column", "MaxLength", "100", "200")]
        [DataRow("relationship", "CascadeDelete", "RemoveLink", "Cascade")]
        [DataRow("webresource", "content", "YWJj", "ZGVm")]
        [DataRow("globalchoice", "Options", "1:Active", "1:Enabled")]
        [DataRow("environmentvariabledefinition", "defaultvalue", "one", "two")]
        [DataRow("connectionreference", "connectorid", "/providers/a", "/providers/b")]
        [DataRow("appmodule", "navigationtype", "0", "1")]
        [DataRow("sitemap", "sitemapxml", "<SiteMap><Area Id='before' /></SiteMap>", "<SiteMap><Area Id='after' /></SiteMap>")]
        public void EverySupportedFamilyDetectsOneComparablePropertyDifference(string kind,
            string property, string before, string after)
        {
            var fixture = Fixture(kind, "portable-key");
            var result = Compare(fixture, Available(fixture.Source, property, before),
                Available(fixture.Target, property, after));
            Assert.AreEqual(ComponentDetailComparisonStatus.Different, result.Status);
            Assert.AreEqual(property, result.Differences.Single().PropertyName,
                true, System.Globalization.CultureInfo.InvariantCulture);
        }

        [DataTestMethod]
        [DataRow("table")]
        [DataRow("column")]
        [DataRow("relationship")]
        [DataRow("webresource")]
        [DataRow("globalchoice")]
        [DataRow("environmentvariabledefinition")]
        [DataRow("connectionreference")]
        [DataRow("appmodule")]
        [DataRow("sitemap")]
        public void EverySupportedFamilyUsesItsEstablishedCaseInsensitivePortableKeyRule(string kind)
        {
            var fixture = Fixture(kind, "Publisher_Component", "publisher_component");
            Assert.AreSame(StringComparer.OrdinalIgnoreCase,
                ComponentDefinitionContractCatalog.For(kind).IdentityComparer);
            var result = Compare(fixture, Available(fixture.Source,
                    ComponentDefinitionContractCatalog.For(kind).ComparableProperties[0], "same"),
                Available(fixture.Target,
                    ComponentDefinitionContractCatalog.For(kind).ComparableProperties[0], "same"));
            Assert.AreEqual(ComponentDetailComparisonStatus.Match, result.Status);
        }

        [TestMethod]
        public void IncompleteAvailableDefinitionCannotBeConstructedOrBecomeMatch()
        {
            var fixture = Fixture("webresource", "new_/script.js");
            Assert.ThrowsException<ArgumentException>(() => new ComponentDefinition(fixture.Source,
                ComponentDefinitionReadStatus.Available,
                new[] { Pair("content", "YWJj") }));
            var incomplete = new ComponentDefinition(fixture.Source,
                ComponentDefinitionReadStatus.Unresolved,
                diagnostic: "Required Web Resource definition fields were not obtained.");
            Assert.AreEqual(ComponentDetailComparisonStatus.Unresolved,
                Compare(fixture, incomplete, Available(fixture.Target, "content", "YWJj")).Status);
        }

        private static ComponentDetailCompareResult Compare(ComparisonFixture fixture,
            ComponentDefinition source, ComponentDefinition target,
            MembershipCompareResult membership = null)
        {
            var memberships = membership == null
                ? new SolutionMembershipComparer().Compare(fixture.SourceSnapshot,
                    fixture.TargetSnapshot)
                : new[] { membership };
            return new ComponentDetailComparer().Compare(memberships,
                Definitions(fixture.SourceSnapshot, source),
                Definitions(fixture.TargetSnapshot, target)).Single();
        }

        private static ComponentDefinition Available(ComponentIdentity identity,
            string property, string value) => new ComponentDefinition(identity,
                ComponentDefinitionReadStatus.Available, Complete(identity, Pair(property, value)));

        private static IEnumerable<KeyValuePair<string, string>> Complete(ComponentIdentity identity,
            params KeyValuePair<string, string>[] overrides)
        {
            var values = overrides.ToDictionary(item => item.Key, item => item.Value,
                StringComparer.OrdinalIgnoreCase);
            return ComponentDefinitionContractCatalog.For(identity.SemanticKind).ComparableProperties
                .Select(item => Pair(item, values.ContainsKey(item) ? values[item] : null));
        }

        private static KeyValuePair<string, string> Pair(string key, string value) =>
            new KeyValuePair<string, string>(key, value);

        private static ComponentDefinitionSnapshot Definitions(MembershipSnapshot snapshot,
            params ComponentDefinition[] definitions) =>
            new ComponentDefinitionSnapshot(snapshot, definitions.Where(item => item != null));

        private static ComparisonFixture Fixture(string kind, string sourceKey,
            string targetKey = null)
        {
            var sourceSolution = new SolutionIdentity(new EnvironmentIdentity(Guid.NewGuid(), "Source"),
                Guid.NewGuid(), "sample");
            var targetSolution = new SolutionIdentity(new EnvironmentIdentity(Guid.NewGuid(), "Target"),
                Guid.NewGuid(), "sample");
            var source = Identity(kind, sourceKey);
            var target = Identity(kind, targetKey ?? sourceKey);
            return new ComparisonFixture(source, target,
                MembershipSnapshot.Complete(sourceSolution, new[] { source }, DateTimeOffset.UtcNow),
                MembershipSnapshot.Complete(targetSolution, new[] { target }, DateTimeOffset.UtcNow));
        }

        private static ComponentIdentity Identity(string kind, string key)
        {
            int type = kind == "table" ? 1 : kind == "column" ? 2 :
                kind == "relationship" ? 10 : kind == "globalchoice" ? 9 :
                kind == "appmodule" ? 80 : kind == "sitemap" ? 62 : 61;
            return new ComponentIdentity(new SolutionComponentRecord(Guid.NewGuid(), type, Guid.NewGuid()),
                IdentityResolutionStatus.Resolved, key, componentTypeKey: kind, semanticKind: kind);
        }

        private sealed class ComparisonFixture
        {
            public ComparisonFixture(ComponentIdentity source, ComponentIdentity target,
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
