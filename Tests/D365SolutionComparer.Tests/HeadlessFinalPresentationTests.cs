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
    public class HeadlessFinalPresentationTests
    {
        [DataTestMethod]
        [DataRow("equal"), DataRow("different"), DataRow("source-only"), DataRow("target-only")]
        [DataRow("source-unavailable"), DataRow("target-unavailable"), DataRow("both-unavailable")]
        public void ReportsUseIndependentlyRetainedNamesWithoutReplacingIdentity(string scenario)
        {
            var f = Reports(scenario == "target-only" ? "target" : scenario == "source-only" ? "source" : "both");
            var originals = f.Presentation.Rows.ToArray(); string csv = ExplorerFixture.Csv(f);
            var source = NameCache(f.Presentation.Source.Snapshot, "Offline report alpha", scenario == "source-unavailable" || scenario == "both-unavailable");
            var target = NameCache(f.Presentation.Target.Snapshot, scenario == "different" ? "Offline report beta" : "Offline report alpha",
                scenario == "target-unavailable" || scenario == "both-unavailable");
            var p = Project(f, source, target); var row = p.Rows[0];
            Assert.AreSame(originals[0], row.Original.AuthoritativeRow); Assert.AreEqual(csv, ExplorerFixture.Csv(f));
            foreach (var side in new[] { row.SourceDisplay, row.TargetDisplay }.Where(s => s != null))
                Assert.AreEqual(DisplayAvailability.NotRetrieved, side.InternalName.Availability);
            if (scenario == "both-unavailable")
            {
                Assert.IsNull(row.PrimaryName(DisplayEnvironmentSide.Source, false));
                Assert.IsNull(row.PrimaryName(DisplayEnvironmentSide.Target, false)); return;
            }
            var scope = scenario == "target-only" ? ObjectsWorkspaceScope.Target : scenario == "source-only" ? ObjectsWorkspaceScope.Source : ObjectsWorkspaceScope.Compare;
            var item = new SolutionObjectNavigation(p).Items(scope).Single();
            string expected = scenario == "equal" ? "Offline report alpha" : scenario == "different" ? "Source: Offline report alpha\u00b7Target: Offline report beta" :
                scenario == "source-only" || scenario == "target-unavailable" ? "Source: Offline report alpha\u00b7Target: [Not available]" :
                "Source: [Not available]\u00b7Target: Offline report alpha";
            Assert.AreEqual(expected, item.Value(true).Replace(" \u00b7 ", "\u00b7"));
            Assert.AreEqual("Reports", item.Category); Assert.AreEqual(originals.Length, p.Rows.Count);
            Assert.AreSame(originals[0].SourceIdentity, row.Original.AuthoritativeRow.SourceIdentity);
            Assert.AreSame(originals[0].TargetIdentity, row.Original.AuthoritativeRow.TargetIdentity);
        }

        [DataTestMethod, DataRow(false), DataRow(true)]
        public void ReportCacheAndSideOwnershipRemainExact(bool swap)
        {
            var f = Reports(); var cache = NameCache(f.Presentation.Source.Snapshot, "Owned Source name");
            var p = Project(f, swap ? null : cache, cache);
            Assert.IsNull(p.Rows[0].PrimaryName(DisplayEnvironmentSide.Target, false));
            if (swap) Assert.IsNull(p.Rows[0].PrimaryName(DisplayEnvironmentSide.Source, false));
            else Assert.AreEqual("Owned Source name", p.Rows[0].PrimaryName(DisplayEnvironmentSide.Source, false));
            Assert.AreEqual("Present in Both", p.Rows[0].Original.MembershipStatus);
        }

        [DataTestMethod, DataRow(IdentityResolutionStatus.Unsupported), DataRow(IdentityResolutionStatus.Unresolved)]
        public void EqualReportTitlesNeverMergeUnresolvedCounterparts(IdentityResolutionStatus status)
        {
            var f = Reports(status: status); string csv = ExplorerFixture.Csv(f);
            var p = Project(f, NameCache(f.Presentation.Source.Snapshot, "Equal title"), NameCache(f.Presentation.Target.Snapshot, "Equal title"));
            var items = new SolutionObjectNavigation(p).Items(ObjectsWorkspaceScope.Compare);
            Assert.AreEqual(2, items.Count); Assert.AreEqual(2, p.Rows.Count); Assert.AreEqual(csv, ExplorerFixture.Csv(f));
            Assert.IsTrue(items.All(i => i.Locals.Count == 1 && i.Status(p) == "Needs Review"));
            Assert.IsTrue(p.Rows.All(r => r.Original.AuthoritativeRow.Presence == MembershipPresence.Indeterminate));
        }

        [DataTestMethod, DataRow("defaultvalue"), DataRow("description"), DataRow("isrequired")]
        public void ActualRetainedDifferenceIsIdentifiedAndOrderedFirst(string property)
        {
            var f = EnvironmentFixture(property); var row = ExplorerFixture.Project(f).Rows.Single();
            Assert.AreEqual("Different", row.Original.DefinitionStatus); Assert.AreEqual(property, row.Original.ChangedProperties);
            Assert.AreEqual(property, row.OrderedComparedProperties[0].PropertyName);
            Assert.AreEqual("source difference", row.OrderedComparedProperties[0].SourceValue);
            Assert.AreEqual("target difference", row.OrderedComparedProperties[0].TargetValue);
            StringAssert.Contains(row.ComparedPropertiesExplanation, property);
            Assert.IsTrue(row.RetainedComparedProperties.Any(p => p.PropertyName == "defaultvalue"));
            Assert.AreSame(row.Original.AuthoritativeRow.DefinitionDetail.Properties.Single(p => p.PropertyName == property), row.OrderedComparedProperties[0]);
        }

        [DataTestMethod, DataRow("none"), DataRow("empty"), DataRow("matching-only"), DataRow("null-matches"), DataRow("reported-name")]
        public void DifferentWithoutChangedEvidenceGetsAnHonestExplanation(string scenario)
        {
            var f = EnvironmentFixture(); var before = f.Presentation.Rows.Single();
            var properties = scenario == "matching-only" || scenario == "reported-name" ? new[] { Property("displayname", "same", "same", "Match") } :
                scenario == "null-matches" ? new[] { Property("displayname", null, null, "Match") } : new ComponentPropertyPresentation[0];
            var detail = scenario == "none" ? null : new ComponentDefinitionDetailPresentation("Environment variable", before.PortableKey,
                before.MembershipStatus, "Different", properties, "Retained diagnostic, unchanged");
            var replacement = before.WithDefinition("Different", scenario == "reported-name" ? "recorded_property" : "", detail);
            f.Presentation = new MembershipComparisonPresentation("SAA", f.Presentation.Source, f.Presentation.Target, new[] { replacement });
            f.Completed = new CompletedMembershipComparison(f.Presentation, f.Completed.DisplayEvidence, DisplayProjectionOutcome.Completed);
            var row = ExplorerFixture.Project(f).Rows.Single();
            Assert.AreSame(replacement, row.Original.AuthoritativeRow); Assert.AreEqual("Different", row.Original.DefinitionStatus);
            StringAssert.Contains(row.ComparedPropertiesExplanation, "difference-producing property details are unavailable");
            if (scenario == "reported-name") StringAssert.Contains(row.ComparedPropertiesExplanation, "recorded_property");
            Assert.IsTrue(row.OrderedComparedProperties.All(p => !p.Changed));
        }

        [DataTestMethod]
        [DataRow(null, null, "Match", "Unavailable")]
        [DataRow(null, "known", "Match", "Unavailable")]
        [DataRow("known", null, "Match", "Unavailable")]
        [DataRow("", "", "Match", "Match")]
        [DataRow("same", "same", "Match", "Match")]
        [DataRow(null, "different", "Different", "Different")]
        [DataRow(null, null, "Unsupported", "Unsupported")]
        [DataRow(null, null, "Indeterminate", "Indeterminate")]
        [DataRow(null, null, "Missing in Source", "Missing in Source")]
        [DataRow(null, null, "Not compared", "Not compared")]
        public void AvailabilityPresentationNeverChangesRawPropertyClassification(string source, string target, string raw, string display)
        {
            var p = Property("field", source, target, raw);
            Assert.AreEqual(display, SolutionObjectsExplorerRow.PropertyComparison(p));
            Assert.AreEqual(raw, p.Comparison); Assert.AreEqual(source, p.SourceValue); Assert.AreEqual(target, p.TargetValue);
            if (source == null) StringAssert.Contains(SolutionObjectsExplorerRow.PropertyValue(source), "availability not established");
        }

        [TestMethod]
        public void SelectedRowViewsRefreshWithoutLeakingPreviousDetails()
        {
            var matching = EnvironmentFixture(); var different = EnvironmentFixture("valueschema");
            var first = ExplorerFixture.Project(matching).Rows.Single(); var second = ExplorerFixture.Project(different).Rows.Single();
            Assert.AreEqual("Match", first.Original.DefinitionStatus); Assert.AreEqual("Different", second.Original.DefinitionStatus);
            Assert.AreEqual("valueschema", second.OrderedComparedProperties[0].PropertyName);
            Assert.IsTrue(first.OrderedComparedProperties.All(p => !p.Changed));
            StringAssert.Contains(second.ComparedPropertiesExplanation, "valueschema");
            Assert.IsFalse(first.ComparedPropertiesExplanation.Contains("valueschema"));
            Assert.AreSame(matching.Presentation.Rows[0], first.Original.AuthoritativeRow);
            Assert.AreSame(different.Presentation.Rows[0], second.Original.AuthoritativeRow);
        }

        [DataTestMethod, DataRow("Unsupported", "Not Compared"), DataRow("Unresolved", "Unresolved"), DataRow("Ambiguous", "Ambiguous")]
        public void UnavailableDefinitionsDoNotProduceConfirmedEqualProperties(string state, string expected)
        {
            var status = (ComponentDefinitionReadStatus)Enum.Parse(typeof(ComponentDefinitionReadStatus), state);
            var f = EnvironmentFixture(); var original = f.Presentation.Rows.Single();
            var source = new ComponentDefinition(original.SourceIdentity, status); var target = new ComponentDefinition(original.TargetIdentity, status);
            var result = new ComponentDefinitionResultPresenter().Apply(f.Presentation,
                new ComponentDefinitionSnapshot(f.Presentation.Source.Snapshot, new[] { source }),
                new ComponentDefinitionSnapshot(f.Presentation.Target.Snapshot, new[] { target }));
            Assert.AreEqual(expected, result.Rows.Single().DefinitionStatus);
            Assert.IsTrue(result.Rows.Single().DefinitionDetail.Properties.All(p => p.Comparison == "Not compared" &&
                SolutionObjectsExplorerRow.PropertyComparison(p) == "Not compared"));
            Assert.AreEqual(original.MembershipStatus, result.Rows.Single().MembershipStatus);
        }

        private static ComponentPropertyPresentation Property(string name, string source, string target, string comparison) =>
            new ComponentPropertyPresentation(name, source, target, comparison == "Different", comparison);
        private static ExplorerFixture Reports(string side = "both", IdentityResolutionStatus status = IdentityResolutionStatus.Resolved)
        {
            string key = "report:v1:offline-established-signature";
            Func<ComponentIdentity> identity = () => new ComponentIdentity(new SolutionComponentRecord(Guid.NewGuid(), 31, Guid.NewGuid()),
                status, status == IdentityResolutionStatus.Resolved ? key : null, semanticKind: ComponentSemanticKinds.Report);
            return ExplorerFixture.Make(side == "target" ? new ComponentIdentity[0] : new[] { identity() },
                side == "source" ? new ComponentIdentity[0] : new[] { identity() });
        }
        private static DataverseComponentMetadataCache NameCache(MembershipSnapshot owner, string name, bool missing = false)
        {
            var rows = owner.Components.Select(i =>
            {
                Guid id = i.Record.ObjectId.Value; var row = new Entity("report", id) { ["reportid"] = id };
                if (!missing) row["name"] = name; return row;
            }).ToArray();
            return ComponentDisplayEvidenceTests.A1NameCache(owner, 31, rows);
        }
        private static SolutionObjectsExplorerProjection Project(ExplorerFixture f, DataverseComponentMetadataCache source, DataverseComponentMetadataCache target)
        {
            f.Completed = new CompletedMembershipComparison(f.Presentation, ComponentDisplayEvidenceProjector.Project(f.Presentation, source, target), DisplayProjectionOutcome.Completed);
            return ExplorerFixture.Project(f);
        }
        private static ExplorerFixture EnvironmentFixture(string changed = null)
        {
            Func<ComponentIdentity> identity = () => new ComponentIdentity(new SolutionComponentRecord(Guid.NewGuid(), 380, Guid.NewGuid()),
                IdentityResolutionStatus.Resolved, "offline_environment", semanticKind: ComponentSemanticKinds.EnvironmentVariableDefinition);
            var f = ExplorerFixture.Make(new[] { identity() }, new[] { identity() });
            Func<ComponentIdentity, bool, ComponentDefinition> definition = (i, source) => new ComponentDefinition(i, ComponentDefinitionReadStatus.Available,
                ComponentDefinitionContractCatalog.For(i.SemanticKind).ComparableProperties.Select(name =>
                    new KeyValuePair<string, string>(name, name == changed ? source ? "source difference" : "target difference" : "same")));
            var p = new ComponentDefinitionResultPresenter().Apply(f.Presentation,
                new ComponentDefinitionSnapshot(f.Presentation.Source.Snapshot, new[] { definition(f.Presentation.Rows[0].SourceIdentity, true) }),
                new ComponentDefinitionSnapshot(f.Presentation.Target.Snapshot, new[] { definition(f.Presentation.Rows[0].TargetIdentity, false) }));
            f.Presentation = p; f.Completed = new CompletedMembershipComparison(p, f.Completed.DisplayEvidence, DisplayProjectionOutcome.Completed); return f;
        }
    }
}
