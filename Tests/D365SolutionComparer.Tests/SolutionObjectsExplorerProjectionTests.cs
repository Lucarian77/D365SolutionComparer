using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using D365SolutionComparer.Models.Identity;
using D365SolutionComparer.Models.Membership;
using D365SolutionComparer.Models.Presentation;
using D365SolutionComparer.Services.Membership;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace D365SolutionComparer.Tests
{
    [TestClass, TestCategory("H2DExplorerProjection")]
    public class SolutionObjectsExplorerProjectionTests
    {
        [DataTestMethod]
        [DataRow(1, "Tables"), DataRow(80, "Model-driven Apps"), DataRow(61, "Web Resources"), DataRow(300, "Canvas Apps / Libraries (kind unclassified)")]
        public void CategoriesAreNavigationAndDoNotReclassify(int type, string category)
        {
            var status = type == 300 ? IdentityResolutionStatus.Unsupported : IdentityResolutionStatus.Resolved;
            var f = ExplorerFixture.Make(new[] { ExplorerFixture.Identity(type, "same", status: status) }, new[] { ExplorerFixture.Identity(type, "same", status: status) });
            var csv = ExplorerFixture.Csv(f); var p = ExplorerFixture.Project(f);
            Assert.AreEqual(type == 300 ? 2 : 1, p.Category(category).RowIndices.Count);
            Assert.AreEqual(1, p.Category(category).SourceOccurrences.Count); Assert.AreEqual(1, p.Category(category).TargetOccurrences.Count);
            Assert.AreSame(f.Presentation.Rows[0], p.Rows[0].Original.AuthoritativeRow);
            Assert.AreEqual(csv, ExplorerFixture.Csv(f));
        }
        [DataTestMethod]
        [DataRow(2, "Columns"), DataRow(10, "Relationships"), DataRow(60, "Forms"), DataRow(26, "Views"), DataRow(59, "Other Components")]
        public void RootLinkedChildrenRetainInclusionEvidence(int type, string group)
        {
            var parent = ExplorerFixture.Identity(1, "account"); var child = ExplorerFixture.Identity(type, "child", parent.Record.SolutionComponentId);
            var f = ExplorerFixture.Make(new[] { parent, child }); var p = ExplorerFixture.Project(f);
            var root = p.Occurrences.Single(o => ReferenceEquals(o.Original.Record, parent.Record)); var link = root.Children.Single();
            Assert.AreEqual(group, link.Group); Assert.AreEqual(HierarchyRelationshipKind.InclusionRoot, link.OriginalAssociation.Kind);
            Assert.AreSame(child.Record, p.Occurrences[link.ChildOccurrenceIndex].Original.Record);
            StringAssert.Contains(link.OriginalAssociation.Provenance, "rootsolutioncomponentid");
        }
        [TestMethod]
        public void StructuralTableScopeRetainsIndependentMetadataEvidence()
        {
            var f = ExplorerFixture.Make(new[] { ExplorerFixture.Identity(1, "account"), ExplorerFixture.Identity(2, "not-a-parsable-key") }, scope: "account");
            var p = ExplorerFixture.Project(f); var link = p.Occurrences.Single(o => o.Original.Record.ComponentType == 1).Children.Single();
            Assert.AreEqual(HierarchyRelationshipKind.StructuralParent, link.OriginalAssociation.Kind);
            Assert.AreEqual("AttributeMetadata.EntityLogicalName", link.OriginalAssociation.Provenance);
        }
        [TestMethod]
        public void AppSettingsRequireVerifiedParentEvidenceAndStaySideSeparated()
        {
            var f = ExplorerFixture.Apps(); var p = ExplorerFixture.Project(f);
            foreach (var side in new[] { DisplayEnvironmentSide.Source, DisplayEnvironmentSide.Target })
            {
                var parent = p.Occurrences.Single(o => o.Original.Side == side && o.Original.Record.ComponentType == 80);
                Assert.AreEqual(5, parent.Children.Count); Assert.IsTrue(parent.Children.All(l => l.Group == "AppSettings"));
                Assert.IsTrue(parent.Children.All(l => p.Occurrences[l.ChildOccurrenceIndex].Original.Side == side));
                Assert.IsTrue(parent.Children.All(l => l.OriginalAssociation.Kind == HierarchyRelationshipKind.StructuralParent));
            }
            Assert.AreEqual(6, p.Rows.Count); Assert.AreEqual(12, p.Occurrences.Count);
        }
        [TestMethod]
        public void DecodedAppSettingKeyDoesNotCreateParenthood()
        {
            var f = ExplorerFixture.Make(new[] { ExplorerFixture.Identity(80, "app"), ExplorerFixture.Identity(10075, ExplorerFixture.Key("app", "OnlineMode")) });
            var p = ExplorerFixture.Project(f); Assert.IsTrue(p.Occurrences.All(o => o.Children.Count == 0));
            Assert.IsTrue(p.Occurrences.Single(o => o.Original.Record.ComponentType == 10075).PlacementNotEstablished);
        }
        [DataTestMethod]
        [DataRow(IdentityResolutionStatus.Unsupported), DataRow(IdentityResolutionStatus.Unresolved), DataRow(IdentityResolutionStatus.Ambiguous)]
        public void UncertainIdentitiesRemainReviewRows(IdentityResolutionStatus status)
        {
            var f = ExplorerFixture.Make(new[] { ExplorerFixture.Identity(300, null, status: status) }); var p = ExplorerFixture.Project(f);
            Assert.AreEqual(MembershipPresence.Indeterminate, p.Rows[0].Original.AuthoritativeRow.Presence);
            Assert.AreEqual(status, p.Occurrences[0].Original.Identity.Status); Assert.AreEqual(1, p.Category("review").RowIndices.Count);
            Assert.AreEqual("Indeterminate", p.Rows[0].Target);
        }
        [TestMethod]
        public void RepeatedRawOccurrencesAndMultipleRowMappingsAreNeverCollapsed()
        {
            var shared = ExplorerFixture.Identity(300, null, status: IdentityResolutionStatus.Unsupported);
            var f = ExplorerFixture.Make(new[] { shared, shared }); var p = ExplorerFixture.Project(f);
            Assert.AreEqual(2, p.Occurrences.Count); Assert.AreEqual(2, p.Rows.Count);
            Assert.IsTrue(p.Occurrences.All(o => o.Original.ComparisonRowIndices.Count == 2));
            Assert.AreEqual(2, p.Category("all").SourceOccurrences.Count); Assert.AreEqual(2, p.Category("all").RowIndices.Count);
        }
        [TestMethod]
        public void UnmappedComparisonAndRawRecordsRemainReachable()
        {
            var f = ExplorerFixture.Make(new[] { ExplorerFixture.Identity(61, "raw") });
            var other = ExplorerFixture.Make(new[] { ExplorerFixture.Identity(61, "different") });
            var presentation = new MembershipComparisonPresentation("SAA", f.Presentation.Source, f.Presentation.Target, other.Presentation.Rows);
            var h = ComponentHierarchyProjector.Project(presentation); var p = SolutionObjectsExplorerProjector.Project(presentation, null, h);
            Assert.AreEqual(1, p.Category("unmapped").RowIndices.Count); Assert.AreEqual(1, p.Category("unmapped").SourceOccurrences.Count);
            Assert.AreEqual(1, p.Category("all").RowIndices.Count); Assert.AreEqual(1, p.Occurrences.Count);
        }
        [TestMethod]
        public void ExactEduShapeConservesEveryOriginalReferenceAndCsv()
        {
            var f = ExplorerFixture.Saa(); var csv = ExplorerFixture.Csv(f); var p = ExplorerFixture.Project(f);
            Assert.AreEqual(840, p.Rows.Count); Assert.AreEqual(1512, p.Occurrences.Count);
            Assert.AreEqual(792, p.SourceOccurrenceCount); Assert.AreEqual(720, p.TargetOccurrenceCount);
            Assert.AreEqual(840, p.Category("all").RowIndices.Count);
            for (int i = 0; i < p.Rows.Count; i++) { Assert.AreEqual(i, p.Rows[i].OriginalRowIndex); Assert.AreSame(f.Presentation.Rows[i], p.Rows[i].Original.AuthoritativeRow); }
            foreach (var o in p.Occurrences) Assert.AreSame(o.Original.Snapshot.Components[o.Original.OriginalRawIndex].Record, o.Original.Record);
            Assert.AreEqual(672, f.Presentation.Summary.PresentInBoth); Assert.AreEqual(103, f.Presentation.Summary.SourceOnly); Assert.AreEqual(32, f.Presentation.Summary.TargetOnly);
            Assert.AreEqual(26, f.Presentation.Summary.IndeterminateUnsupported); Assert.AreEqual(4, f.Presentation.Summary.IndeterminateAmbiguous); Assert.AreEqual(3, f.Presentation.Summary.IndeterminateResolved);
            Assert.AreEqual(csv, ExplorerFixture.Csv(f));
        }
        [TestMethod]
        public void RepeatedProjectionIsDeterministicAndHasNoServiceDependency()
        {
            var f = ExplorerFixture.Apps(); var a = ExplorerFixture.Project(f); var b = ExplorerFixture.Project(f);
            CollectionAssert.AreEqual(a.Occurrences.Select(o => o.Label).ToArray(), b.Occurrences.Select(o => o.Label).ToArray());
            CollectionAssert.AreEqual(a.Categories.Select(c => c.CountText).ToArray(), b.Categories.Select(c => c.CountText).ToArray());
            Assert.IsFalse(typeof(SolutionObjectsExplorerProjector).GetMethods(System.Reflection.BindingFlags.Static | System.Reflection.BindingFlags.NonPublic)
                .SelectMany(m => m.GetParameters()).Any(p => typeof(Microsoft.Xrm.Sdk.IOrganizationService).IsAssignableFrom(p.ParameterType)));
        }
        [TestMethod]
        public void CancelledAndWrongOwnerProjectionNeverReturnsPartialResults()
        {
            var f = ExplorerFixture.Apps(); var h = ComponentHierarchyProjector.Project(f.Completed, f.Completed.SourceHierarchyParents, f.Completed.TargetHierarchyParents);
            Assert.ThrowsException<OperationCanceledException>(() => SolutionObjectsExplorerProjector.Project(f.Presentation, f.Completed, h, new CancellationToken(true)));
            var other = ExplorerFixture.Apps(); Assert.ThrowsException<InvalidOperationException>(() => SolutionObjectsExplorerProjector.Project(other.Presentation, other.Completed, h));
        }
        [DataTestMethod, DataRow(10266), DataRow(10267), DataRow(10072), DataRow(74), DataRow(10276)]
        public void UnverifiedLocalRegistrationsCannotManufactureCategory(int type)
        {
            var f = ExplorerFixture.Make(new[] { ExplorerFixture.Identity(type, null, status: IdentityResolutionStatus.Unsupported) });
            Assert.AreEqual("Other / unclassified components", ExplorerFixture.Project(f).Occurrences[0].Category);
        }
        [TestMethod]
        public void SharedAppActionLabelDoesNotCrossMatchRawNamespaces()
        {
            Func<int, ComponentIdentity> action = raw => new ComponentIdentity(new SolutionComponentRecord(Guid.NewGuid(), raw, Guid.NewGuid()),
                IdentityResolutionStatus.Unsupported, componentTypeKey: "raw:" + raw, registeredDefinition: new SolutionComponentDefinitionIdentity(raw, "appaction", "appaction"));
            var f = ExplorerFixture.Make(new[] { action(10266) }, new[] { action(10267) }); var p = ExplorerFixture.Project(f);
            Assert.AreEqual(2, p.Category("App Actions (diagnostic family only)").RowIndices.Count);
            Assert.IsTrue(p.Rows.All(r => r.Original.AuthoritativeRow.Presence == MembershipPresence.Indeterminate));
            CollectionAssert.AreEquivalent(new[] { "raw:10266", "raw:10267" }, p.Occurrences.Select(o => o.Original.Identity.ComponentTypeKey).ToArray());
        }
        [TestMethod]
        public void ConflictingRootAndStructuralParentsRemainReviewWithoutChoosingParent()
        {
            var root = ExplorerFixture.Identity(1, "account"); var structure = ExplorerFixture.Identity(1, "contact");
            var f = ExplorerFixture.Make(new[] { root, structure, ExplorerFixture.Identity(2, "column", root.Record.SolutionComponentId) }, scope: "contact");
            var p = ExplorerFixture.Project(f); var child = p.Occurrences.Single(o => o.Original.Record.ComponentType == 2);
            Assert.IsTrue(child.PlacementNotEstablished); Assert.IsTrue(child.Original.HasDifferentParentAssociations);
            Assert.IsTrue(p.Occurrences.Where(o => o.Original.Record.ComponentType == 1).All(o => o.Children.Count == 0));
            Assert.AreEqual(2, child.Original.Associations.Count(a => a.Status == HierarchyAssociationStatus.Verified));
        }
        [TestMethod]
        public void FailedOptionalDisplayStillRetainsEveryAuthoritativeRowAndRawOccurrence()
        {
            var f = ExplorerFixture.Make(new[] { ExplorerFixture.Identity(1, "account") });
            f.Completed = new CompletedMembershipComparison(f.Presentation, new ComponentDisplayEvidence[0], DisplayProjectionOutcome.Failed);
            var p = ExplorerFixture.Project(f); Assert.AreEqual(1, p.Rows.Count); Assert.AreEqual(1, p.Occurrences.Count);
            Assert.IsNull(p.Rows[0].SourceDisplay); Assert.AreSame(f.Presentation.Rows[0], p.Rows[0].Original.AuthoritativeRow);
        }
        [TestMethod, TestCategory("H2E2B1")]
        public void WorkspaceSideUnitsNeverBorrowUnionCandidateCountOrMakerBenchmark()
        {
            var f = ExplorerFixture.Make(Enumerable.Range(0, 179).Select(i => ExplorerFixture.Identity(61, "resource" + i)),
                new[] { ExplorerFixture.Identity(61, "resource0") }); var p = ExplorerFixture.Project(f);
            Assert.AreEqual(179, p.Category("objects").RowIndices.Count);
            var target = p.WorkspaceCategories(ObjectsWorkspaceScope.Target, Enumerable.Range(0, p.Rows.Count), true).First();
            Assert.AreEqual(1, target.TargetOccurrences.Count); Assert.AreEqual(0, target.SourceOccurrences.Count); Assert.AreEqual(1, target.RowIndices.Count);
            Assert.AreNotEqual(150, target.TargetOccurrences.Count); Assert.AreNotEqual(179, target.TargetOccurrences.Count);
            var compare = p.WorkspaceCategories(ObjectsWorkspaceScope.Compare, Enumerable.Range(0, p.Rows.Count), true).First();
            Assert.AreEqual(179, compare.RowIndices.Count);
        }
        [DataTestMethod, TestCategory("H2E2B1"), DataRow(29), DataRow(60), DataRow(300)]
        public void IncompleteSubtypeContractNeverManufacturesMakerCategory(int raw)
        {
            var identity = ExplorerFixture.Identity(raw, "diagnostic", status: raw == 300 ? IdentityResolutionStatus.Unsupported : IdentityResolutionStatus.Resolved).WithWorkflowCategory(raw == 29 ? (int?)5 : null);
            var p = ExplorerFixture.Project(ExplorerFixture.Make(new[] { identity }, new[] { identity }));
            Assert.IsTrue(p.Occurrences.All(o => o.PrimaryCategory == null));
            var cats = p.WorkspaceCategories(ObjectsWorkspaceScope.Target, Enumerable.Range(0, p.Rows.Count), true);
            Assert.AreEqual(0, cats.First().TargetOccurrences.Count); Assert.AreEqual(1, cats.Last().TargetOccurrences.Count);
            Assert.IsFalse(cats.Any(c => c.Name == "Cloud flows" || c.Name == "Dashboards" || c.Name == "Component libraries"));
        }
        [TestMethod, TestCategory("H2E2B1")]
        public void WorkspaceConservesEduRowsOccurrencesAndExcludesStandaloneSettings()
        {
            var f = ExplorerFixture.Saa(); var csv = ExplorerFixture.Csv(f); var p = ExplorerFixture.Project(f);
            Assert.AreEqual(840, p.Rows.Count); Assert.AreEqual(1512, p.Occurrences.Count); Assert.AreEqual(792, p.SourceOccurrenceCount); Assert.AreEqual(720, p.TargetOccurrenceCount);
            var compare = p.WorkspaceCategories(ObjectsWorkspaceScope.Compare, Enumerable.Range(0, 840), true).First();
            CollectionAssert.AreEqual(p.Occurrences.Where(o => o.PrimaryCategory != null).SelectMany(o => o.Original.ComparisonRowIndices).Distinct().OrderBy(i => i).ToArray(), compare.RowIndices.ToArray());
            Assert.AreEqual(840, p.Category("all").RowIndices.Count, "Technical All remains complete.");
            for (int i = 0; i < 840; i++) Assert.AreSame(f.Presentation.Rows[i], p.Rows[i].Original.AuthoritativeRow);
            var apps = ExplorerFixture.Project(ExplorerFixture.Apps());
            Assert.IsFalse(apps.WorkspaceCategories(ObjectsWorkspaceScope.Target, Enumerable.Range(0, apps.Rows.Count), true).Any(c => c.Name == "AppSettings"));
            Assert.AreEqual(10, apps.Occurrences.Count(o => o.Original.Record.ComponentType == 10075));
            Assert.AreEqual(10, apps.Occurrences.Sum(o => o.Children.Count(l => l.Group == "AppSettings" && l.OriginalAssociation.Kind == HierarchyRelationshipKind.StructuralParent)));
            Assert.AreEqual(csv, ExplorerFixture.Csv(f));
        }
        [TestMethod, TestCategory("H2E2B1")]
        public void RepeatedRawOccurrenceIsNotDeduplicatedIntoVerifiedLocalObjects()
        {
            var identity = ExplorerFixture.Identity(61, "repeated"); var p = ExplorerFixture.Project(ExplorerFixture.Make(new[] { identity, identity }));
            var local = p.WorkspaceCategories(ObjectsWorkspaceScope.Source, Enumerable.Range(0, p.Rows.Count), true).First();
            Assert.AreEqual(2, local.SourceOccurrences.Count); Assert.AreEqual(2, p.Occurrences.Count);
            Assert.AreEqual(p.Rows.Count, local.RowIndices.Count);
        }

        [TestMethod, TestCategory("H2E2C2A")]
        public void CandidateUnionPreservesRepeatedAliasesOneSidedAndExcludedRecords()
        {
            var shared = ExplorerFixture.Identity(61, "shared");
            var sourceOnly = ExplorerFixture.Identity(61, "source");
            var f = ExplorerFixture.Make(new[] { shared, sourceOnly, sourceOnly, ExplorerFixture.Identity(300, null, status: IdentityResolutionStatus.Unsupported) },
                new[] { ExplorerFixture.Identity(61, "shared"), ExplorerFixture.Identity(80, "target"), ExplorerFixture.Identity(29, null, status: IdentityResolutionStatus.Ambiguous) });
            var original = ExplorerFixture.Project(f); var csv = ExplorerFixture.Csv(f);
            // Synthetic navigation aliases exercise projection union, never production cross-family matching.
            var aliases = new SolutionObjectsExplorerProjection(original.Owner, original.Rows, original.Occurrences.Select(o =>
                new SolutionObjectsExplorerOccurrence(o.Original, o.Category, o.Label, o.ObjectCandidate, o.Children,
                    o.PrimaryCategory != null && o.Original.Side == DisplayEnvironmentSide.Target ? "Apps" : o.PrimaryCategory)), original.Categories, original.Diagnostics);
            var allowed = Enumerable.Range(0, aliases.Rows.Count).ToArray(); var categories = aliases.WorkspaceCategories(ObjectsWorkspaceScope.Compare, allowed, true);
            var expected = aliases.Occurrences.Where(o => o.PrimaryCategory != null).SelectMany(o => o.Original.ComparisonRowIndices).Distinct().OrderBy(i => i).ToArray();
            CollectionAssert.AreEqual(expected, categories.First().RowIndices.ToArray());
            Assert.IsTrue(categories.Skip(1).SelectMany(c => c.RowIndices).GroupBy(i => i).Any(g => g.Count() > 1));
            Assert.AreEqual(expected.Length, categories.First().RowIndices.Distinct().Count());
            Assert.AreEqual(7, aliases.Occurrences.Count); Assert.AreEqual(4, aliases.SourceOccurrenceCount); Assert.AreEqual(3, aliases.TargetOccurrenceCount);
            Assert.AreEqual(aliases.Rows.Count, aliases.Category("all").RowIndices.Count); Assert.IsTrue(expected.Length < aliases.Rows.Count);
            var source = aliases.WorkspaceCategories(ObjectsWorkspaceScope.Source, allowed, true).First(); Assert.AreEqual(3, source.SourceOccurrences.Count); Assert.AreEqual(0, source.TargetOccurrences.Count);
            var filtered = expected.Take(1).ToArray(); CollectionAssert.AreEqual(filtered, aliases.WorkspaceCategories(ObjectsWorkspaceScope.Compare, filtered, false).First().RowIndices.ToArray());
            Assert.AreEqual(csv, ExplorerFixture.Csv(f));
        }

        [TestMethod, TestCategory("H2E2C2A")]
        public void CandidateEligibilityRemainsUnchangedAndDifferencesUseOriginalRows()
        {
            var f = ExplorerFixture.Make(new[] { ExplorerFixture.Identity(36, null, status: IdentityResolutionStatus.Unsupported),
                ExplorerFixture.Identity(511, null, status: IdentityResolutionStatus.Unsupported), ExplorerFixture.Identity(29, "workflow"), ExplorerFixture.Identity(60, "form"),
                ExplorerFixture.Identity(300, null, status: IdentityResolutionStatus.Unsupported) }); var p = ExplorerFixture.Project(f);
            Assert.IsNotNull(p.Occurrences.Single(o => o.Original.Record.ComponentType == 36).PrimaryCategory);
            Assert.IsNotNull(p.Occurrences.Single(o => o.Original.Record.ComponentType == 511).PrimaryCategory);
            Assert.IsTrue(p.Occurrences.Where(o => new[] { 29, 60, 300 }.Contains(o.Original.Record.ComponentType)).All(o => o.PrimaryCategory == null));
            var differences = p.Rows.Where(row => row.Original.AuthoritativeRow.Presence != MembershipPresence.PresentInBoth || row.Original.DefinitionStatus == "Different").Select(row => row.OriginalRowIndex).ToArray();
            var eligible = p.Occurrences.Where(o => o.PrimaryCategory != null).SelectMany(o => o.Original.ComparisonRowIndices).Distinct().Where(differences.Contains).OrderBy(i => i).ToArray();
            CollectionAssert.AreEqual(eligible, p.WorkspaceCategories(ObjectsWorkspaceScope.Compare, differences, false).First().RowIndices.ToArray());
            Assert.AreEqual(5, p.Category("all").RowIndices.Count);
        }

    }

    internal sealed class ExplorerFixture
    {
        internal MembershipComparisonPresentation Presentation;
        internal CompletedMembershipComparison Completed;
        internal static ComponentIdentity Identity(int type, string key, Guid? root = null,
            IdentityResolutionStatus status = IdentityResolutionStatus.Resolved, bool matchOnly = false) =>
            new ComponentIdentity(new SolutionComponentRecord(Guid.NewGuid(), type, Guid.NewGuid(), rootSolutionComponentId: root), status,
                status == IdentityResolutionStatus.Resolved ? key : null, "Synthetic retained diagnostic",
                componentTypeKey: type == 1 ? ComponentSemanticKinds.Table : type == 80 ? ComponentSemanticKinds.AppModule : type == 10075 ? ComponentSemanticKinds.AppSetting : null,
                registeredDefinition: type == 10075 ? new SolutionComponentDefinitionIdentity(10075, "AppSetting", "appsetting") : null,
                inventoryAbsencePolicy: matchOnly ? InventoryAbsencePolicy.MatchOnly : InventoryAbsencePolicy.CompleteInventory);
        internal static string Key(string app, string setting) => "appsetting:v1:" + app.Length + ":" + app + ":" + setting.Length + ":" + setting;
        internal static ExplorerFixture Make(IEnumerable<ComponentIdentity> source, IEnumerable<ComponentIdentity> target = null, string scope = null)
        {
            Func<IEnumerable<ComponentIdentity>, bool, MembershipEnvironmentResult> env = (items, right) => MembershipEnvironmentResult.FromSnapshot(right ? "Target fixture" : "Source fixture",
                MembershipSnapshot.Complete(new SolutionIdentity(new EnvironmentIdentity(Guid.NewGuid(), right ? "Target fixture" : "Source fixture"), Guid.NewGuid(), "SAA"), items,
                    new DateTimeOffset(2026, 10, 13, 0, 0, 0, TimeSpan.Zero)), 27, TimeSpan.Zero);
            var p = new MembershipResultPresenter().Create(env(source, false), env(target ?? new ComponentIdentity[0], true));
            var f = new ExplorerFixture { Presentation = p };
            f.Completed = new CompletedMembershipComparison(p, p.Rows.Select((r, index) => new ComponentDisplayEvidence(index,
                Display(r.SourceIdentity, DisplayEnvironmentSide.Source, scope), Display(r.TargetIdentity, DisplayEnvironmentSide.Target, scope))),
                DisplayProjectionOutcome.Completed, DisplayCacheAssociation.Verified, DisplayCacheAssociation.Verified);
            return f;
        }
        internal static ComponentSideDisplayEvidence Display(ComponentIdentity i, DisplayEnvironmentSide side, string scope = null,
            ComponentDisplayValue friendly = null, ComponentDisplayValue name = null) => new ComponentSideDisplayEvidence(side,
                i?.Record.SolutionComponentId, i?.Record.ObjectId, i?.Record.ComponentType, new ComponentDisplayValue(DisplayAvailability.NotRetrieved),
                friendly ?? new ComponentDisplayValue(DisplayAvailability.NotRetrieved), name ?? new ComponentDisplayValue(DisplayAvailability.NotRetrieved),
                DisplayBackingCorrelationStatus.VerifiedMetadata, scope == null ? new ComponentDisplayValue(DisplayAvailability.NotRetrieved) :
                    new ComponentDisplayValue(DisplayAvailability.Available, scope, "AttributeMetadata.EntityLogicalName"));
        internal static SolutionObjectsExplorerProjection Project(ExplorerFixture f) => SolutionObjectsExplorerProjector.Project(f.Presentation, f.Completed,
            ComponentHierarchyProjector.Project(f.Completed, f.Completed.SourceHierarchyParents, f.Completed.TargetHierarchyParents));
        internal static string Csv(ExplorerFixture f) => new MembershipCoverageCsvExporter().CreateCsv(f.Presentation);
        internal static ExplorerFixture Apps()
        {
            var left = new List<ComponentIdentity> { Identity(80, "app") }; var right = new List<ComponentIdentity> { Identity(80, "app") };
            foreach (var setting in new[] { "OnlineMode", "ChannelAvailability", "EnableRoleBasedOfflineProfileAccess", "msdyn_enablertecopilot", "msdyn_rtecopilotdefaultonruncompleted" })
            { left.Add(Identity(10075, Key("app", setting))); right.Add(Identity(10075, Key("app", setting))); }
            var f = Make(left, right);
            Func<MembershipSnapshot, DisplayEnvironmentSide, ComponentHierarchyParentEvidence> parents = (snapshot, side) =>
            {
                var parent = snapshot.Components.Single(c => c.Record.ComponentType == 80);
                return new ComponentHierarchyParentEvidence(snapshot, side, snapshot.Components.Where(c => c.Record.ComponentType == 10075).Select(c =>
                    new HierarchyStructuralObservation(c.Record.SolutionComponentId, c.Record.ObjectId, parent.Record.ObjectId,
                        ComponentSemanticKinds.AppModule, "app", "Synthetic bound AppSetting parent", true)));
            };
            f.Completed = new CompletedMembershipComparison(f.Completed, parents(f.Presentation.Source.Snapshot, DisplayEnvironmentSide.Source),
                parents(f.Presentation.Target.Snapshot, DisplayEnvironmentSide.Target)); return f;
        }
        internal static ExplorerFixture Saa()
        {
            var source = Enumerable.Range(0, 672).Select(i => Identity(61, "paired" + i)).ToList();
            var target = Enumerable.Range(0, 672).Select(i => Identity(61, "paired" + i)).ToList();
            source.AddRange(Enumerable.Range(0, 103).Select(i => Identity(61, "source" + i)));
            target.AddRange(Enumerable.Range(0, 32).Select(i => Identity(61, "target" + i)));
            source.AddRange(Enumerable.Range(0, 14).Select(i => Identity(300, null, status: IdentityResolutionStatus.Unsupported)));
            target.AddRange(Enumerable.Range(0, 12).Select(i => Identity(300, null, status: IdentityResolutionStatus.Unsupported)));
            source.AddRange(Enumerable.Range(0, 2).Select(i => Identity(29, null, status: IdentityResolutionStatus.Ambiguous)));
            target.AddRange(Enumerable.Range(0, 2).Select(i => Identity(29, null, status: IdentityResolutionStatus.Ambiguous)));
            source.Add(Identity(60, "source-match-only", matchOnly: true));
            target.Add(Identity(60, "target-match-only1", matchOnly: true)); target.Add(Identity(60, "target-match-only2", matchOnly: true));
            return Make(source, target);
        }
    }
}
