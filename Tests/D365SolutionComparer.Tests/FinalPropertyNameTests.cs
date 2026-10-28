using System;
using System.Collections.Generic;
using System.Linq;
using D365SolutionComparer.Models.ComponentDetails;
using D365SolutionComparer.Models.Membership;
using D365SolutionComparer.Models.Presentation;
using D365SolutionComparer.Services.ComponentDetails;
using D365SolutionComparer.Services.Membership;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Microsoft.Xrm.Sdk;

namespace D365SolutionComparer.Tests
{
    [TestClass]
    public class FinalPropertyNameTests
    {
        [DataTestMethod, DataRow("Team", true), DataRow("Team", false), DataRow("Mask", true), DataRow("Mask", false)]
        public void VerifiedSingleSideNamePreservesItsOwnerAndNeedsReview(string family, bool source)
        {
            var f = Fixture(family, source ? "source" : "target"); string csv = ExplorerFixture.Csv(f);
            var p = Project(f, source ? Cache(f.Presentation.Source.Snapshot, family, "Synthetic maker name") : null,
                source ? null : Cache(f.Presentation.Target.Snapshot, family, "Synthetic maker name"));
            var item = new SolutionObjectNavigation(p).Items(source ? ObjectsWorkspaceScope.Source : ObjectsWorkspaceScope.Target).Single();
            Assert.AreEqual(source ? "Source: Synthetic maker name \u00b7 Target: [Not available]" : "Source: [Not available] \u00b7 Target: Synthetic maker name", item.Value(true));
            Assert.AreEqual("Needs Review", item.Status(p)); Assert.AreEqual(1, item.Locals.Count);
            Assert.AreEqual(source ? DisplayEnvironmentSide.Source : DisplayEnvironmentSide.Target, item.Locals[0].Side);
            Assert.AreEqual(csv, ExplorerFixture.Csv(f)); Assert.AreSame(f.Presentation.Rows[0], p.Rows[0].Original.AuthoritativeRow);
        }

        [DataTestMethod, DataRow("Team", true), DataRow("Team", false), DataRow("Mask", true), DataRow("Mask", false)]
        public void EqualOrDifferentNamesNeverMergeUnconfirmedCounterparts(string family, bool equal)
        {
            var f = Fixture(family); var p = Project(f, Cache(f.Presentation.Source.Snapshot, family, "Synthetic left"),
                Cache(f.Presentation.Target.Snapshot, family, equal ? "Synthetic left" : "Synthetic right"));
            var items = new SolutionObjectNavigation(p).Items(ObjectsWorkspaceScope.Compare);
            Assert.AreEqual(2, items.Count); Assert.IsTrue(items.All(i => i.Locals.Count == 1 && i.Status(p) == "Needs Review"));
            Assert.IsTrue(items.Any(i => i.Value(true) == "Source: Synthetic left \u00b7 Target: [Not available]"));
            Assert.IsTrue(items.Any(i => i.Value(true) == "Source: [Not available] \u00b7 Target: " + (equal ? "Synthetic left" : "Synthetic right")));
            Assert.IsTrue(p.Rows.All(r => r.Original.AuthoritativeRow.Presence == MembershipPresence.Indeterminate));
        }

        [DataTestMethod, DataRow("Team", true), DataRow("Team", false), DataRow("Mask", true), DataRow("Mask", false)]
        public void MissingOrEmptyNameEvidenceDoesNotInventAName(string family, bool missing)
        {
            var f = Fixture(family, "source"); var p = Project(f, Cache(f.Presentation.Source.Snapshot, family, missing ? null : ""), null);
            Assert.IsNull(p.Rows[0].PrimaryName(DisplayEnvironmentSide.Source, false));
            Assert.IsNull(p.Rows[0].PrimaryName(DisplayEnvironmentSide.Target, false));
            Assert.AreEqual(IdentityResolutionStatus.Unsupported, p.Rows[0].Original.AuthoritativeRow.SourceIdentity.Status);
            Assert.IsNull(p.Rows[0].Original.AuthoritativeRow.SourceIdentity.ComparisonKey);
        }

        [DataTestMethod, DataRow("Team", true), DataRow("Team", false), DataRow("Mask", true), DataRow("Mask", false)]
        public void ReplayedOrOppositeSideCachesCannotSupplyNames(string family, bool replay)
        {
            var f = Fixture(family); var owner = f.Presentation.Source.Snapshot;
            var cache = Cache(owner, family, "Source owned metadata");
            if (replay)
            {
                var copy = MembershipSnapshot.Complete(owner.Solution, owner.Components, owner.CapturedAt);
                if (family == "Team") { DataverseComponentMetadataCache.NormalDisplayNameObservation unused; Assert.IsFalse(cache.TryGetNormalDisplayName(copy, copy.Components[0], out unused)); }
                else { MakerObjectDisplayEvidence unused; Assert.IsFalse(cache.TryGetMakerObjectDisplay(copy, copy.Components[0], out unused)); }
            }
            var p = Project(f, cache, cache);
            Assert.AreEqual("Source owned metadata", p.Rows.Single(r => r.SourceDisplay != null).PrimaryName(DisplayEnvironmentSide.Source, false));
            Assert.IsNull(p.Rows.Single(r => r.TargetDisplay != null).PrimaryName(DisplayEnvironmentSide.Target, false));
        }

        [DataTestMethod, DataRow("Team"), DataRow("Mask")]
        public void AmbiguousNameCorrelationIsNotVerifiedNameEvidence(string family)
        {
            var f = Fixture(family, "source"); var identity = f.Presentation.Rows[0].SourceIdentity;
            var maker = family == "Mask" ? new MakerObjectDisplayEvidence(identity, MakerObjectKind.SecuredMaskingRule, identity.Record.ObjectId.Value,
                new ComponentDisplayValue(DisplayAvailability.Available, "Unverified name", "syntheticmask.caption"),
                new ComponentDisplayValue(DisplayAvailability.NotRetrieved), DisplayBackingCorrelationStatus.Ambiguous) : null;
            var evidence = new ComponentSideDisplayEvidence(DisplayEnvironmentSide.Source, identity.Record.SolutionComponentId,
                identity.Record.ObjectId, identity.Record.ComponentType, new ComponentDisplayValue(DisplayAvailability.NotRetrieved),
                new ComponentDisplayValue(DisplayAvailability.Available, "Unverified name", family == "Team" ? "NormalMembershipDiagnostic.teamtemplate.teamtemplatename" : "syntheticmask.caption"),
                new ComponentDisplayValue(DisplayAvailability.NotRetrieved), DisplayBackingCorrelationStatus.Ambiguous,
                new ComponentDisplayValue(DisplayAvailability.NotRetrieved), makerObject: maker);
            f.Completed = new CompletedMembershipComparison(f.Presentation, new[] { new ComponentDisplayEvidence(0, evidence,
                ExplorerFixture.Display(null, DisplayEnvironmentSide.Target)) }, DisplayProjectionOutcome.Completed);
            Assert.IsNull(ExplorerFixture.Project(f).Rows[0].PrimaryName(DisplayEnvironmentSide.Source, false));
        }

        [TestMethod]
        public void RetainedEnvironmentDifferenceHasAllSevenRowsAndExactSideValues()
        {
            var f = Environment(); var row = ExplorerFixture.Project(f).Rows[0];
            Assert.AreEqual("Different", row.Original.DefinitionStatus);
            CollectionAssert.AreEqual(new[] { "defaultvalue", "description", "displayname", "isrequired", "secretstore", "type", "valueschema" }, row.OrderedComparedProperties.Select(p => p.PropertyName).ToArray());
            Assert.AreEqual("defaultvalue", row.OrderedComparedProperties[0].PropertyName);
            Assert.AreEqual("synthetic-left", row.OrderedComparedProperties[0].SourceValue);
            Assert.AreEqual("synthetic-right", row.OrderedComparedProperties[0].TargetValue);
            Assert.AreSame(row.Original.AuthoritativeRow.DefinitionDetail.Properties.Single(p => p.PropertyName == "defaultvalue"), row.OrderedComparedProperties[0]);
            StringAssert.Contains(row.ComparedPropertiesExplanation, "Changed properties shown first: defaultvalue");
        }

        [DataTestMethod, DataRow(true), DataRow(false)]
        public void UnavailableChangedValueRetainsDifferentAndItsAvailabilityLimitation(bool sourceMissing)
        {
            var property = new ComponentPropertyPresentation("defaultvalue", sourceMissing ? null : "synthetic-left", sourceMissing ? "synthetic-right" : null, true, "Different");
            Assert.AreEqual("Different", SolutionObjectsExplorerRow.PropertyComparison(property));
            StringAssert.Contains(SolutionObjectsExplorerRow.PropertyExplanation("Different", new[] { property }, true), "value evidence is unavailable");
            Assert.IsTrue(property.Changed); Assert.AreEqual("Different", property.Comparison);
        }

        [TestMethod]
        public void FreshPropertyViewPreservesTheCurrentSelectedOwner()
        {
            var first = ExplorerFixture.Project(Environment()).Rows[0]; var next = ExplorerFixture.Project(Environment()).Rows[0];
            Assert.AreNotSame(first.Original.AuthoritativeRow, next.Original.AuthoritativeRow);
            Assert.AreSame(next.Original.AuthoritativeRow.DefinitionDetail.Properties.Single(p => p.PropertyName == "defaultvalue"), next.OrderedComparedProperties[0]);
            Assert.AreNotSame(first.OrderedComparedProperties[0], next.OrderedComparedProperties[0]);
            Assert.AreEqual("defaultvalue", next.OrderedComparedProperties[0].PropertyName);
        }

        [TestMethod]
        public void MissingValuesAreNotConfirmedEqualByThePresentation()
        {
            var property = new ComponentPropertyPresentation("field", null, null, false, "Match");
            Assert.AreEqual("Unavailable", SolutionObjectsExplorerRow.PropertyComparison(property));
            Assert.AreEqual("Match", property.Comparison); Assert.IsFalse(property.Changed);
        }

        private static ExplorerFixture Fixture(string family, string side = "both")
        {
            Func<ComponentIdentity> identity = () => new ComponentIdentity(new SolutionComponentRecord(Guid.NewGuid(), family == "Team" ? 511 : 74, Guid.NewGuid()),
                IdentityResolutionStatus.Unsupported, semanticKind: family == "Team" ? ComponentSemanticKinds.TeamTemplate : null,
                registeredDefinition: family == "Team" ? null : new SolutionComponentDefinitionIdentity(74, "MaskingRule", "syntheticmask"));
            return ExplorerFixture.Make(side == "target" ? new ComponentIdentity[0] : new[] { identity() }, side == "source" ? new ComponentIdentity[0] : new[] { identity() });
        }
        private static DataverseComponentMetadataCache Cache(MembershipSnapshot owner, string family, string name)
        {
            if (family == "Team")
                return ComponentDisplayEvidenceTests.A1NameCache(owner, 511, owner.Components.Select(i =>
                {
                    var row = new Entity("teamtemplate", i.Record.ObjectId.Value) { ["teamtemplateid"] = i.Record.ObjectId.Value };
                    if (name != null) row["teamtemplatename"] = name; return row;
                }).ToArray());
            var cache = new DataverseComponentMetadataCache(); cache.BeginNormalDisplayNames(owner); cache.SealNormalDisplayNames(owner);
            var values = new Dictionary<SolutionComponentRecord, MakerObjectDisplayEvidence>();
            if (!string.IsNullOrWhiteSpace(name)) foreach (var i in owner.Components)
                values[i.Record] = new MakerObjectDisplayEvidence(i, MakerObjectKind.SecuredMaskingRule, i.Record.ObjectId.Value,
                    new ComponentDisplayValue(DisplayAvailability.Available, name, "syntheticmask.caption"),
                    new ComponentDisplayValue(DisplayAvailability.NotRetrieved), DisplayBackingCorrelationStatus.Unique);
            cache.SealMakerObjectDisplay(owner, values, new string[0]); return cache;
        }
        private static SolutionObjectsExplorerProjection Project(ExplorerFixture f, DataverseComponentMetadataCache source, DataverseComponentMetadataCache target)
        {
            f.Completed = new CompletedMembershipComparison(f.Presentation, ComponentDisplayEvidenceProjector.Project(f.Presentation, source, target), DisplayProjectionOutcome.Completed);
            return ExplorerFixture.Project(f);
        }
        private static ExplorerFixture Environment()
        {
            Func<ComponentIdentity> identity = () => new ComponentIdentity(new SolutionComponentRecord(Guid.NewGuid(), 380, Guid.NewGuid()), IdentityResolutionStatus.Resolved,
                "synthetic_environment", semanticKind: ComponentSemanticKinds.EnvironmentVariableDefinition);
            var f = ExplorerFixture.Make(new[] { identity() }, new[] { identity() });
            Func<ComponentIdentity, bool, ComponentDefinition> definition = (i, source) => new ComponentDefinition(i, ComponentDefinitionReadStatus.Available,
                ComponentDefinitionContractCatalog.For(i.SemanticKind).ComparableProperties.Select(name => new KeyValuePair<string, string>(name,
                    name == "defaultvalue" ? source ? "synthetic-left" : "synthetic-right" : "synthetic-match")));
            var p = new ComponentDefinitionResultPresenter().Apply(f.Presentation,
                new ComponentDefinitionSnapshot(f.Presentation.Source.Snapshot, new[] { definition(f.Presentation.Rows[0].SourceIdentity, true) }),
                new ComponentDefinitionSnapshot(f.Presentation.Target.Snapshot, new[] { definition(f.Presentation.Rows[0].TargetIdentity, false) }));
            f.Presentation = p; f.Completed = new CompletedMembershipComparison(p, f.Completed.DisplayEvidence, DisplayProjectionOutcome.Completed); return f;
        }
    }
}
