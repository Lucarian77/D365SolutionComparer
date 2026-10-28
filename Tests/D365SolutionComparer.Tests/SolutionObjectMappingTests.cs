using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Runtime.ExceptionServices;
using System.Threading;
using System.Windows.Forms;
using D365SolutionComparer.Models.Membership;
using D365SolutionComparer.Models.Presentation;
using D365SolutionComparer.Services.Membership;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Microsoft.Xrm.Sdk;

namespace D365SolutionComparer.Tests
{
    [TestClass, TestCategory("R2EObjectMapping")]
    public class SolutionObjectMappingTests
    {
        public TestContext TestContext { get; set; }
        [TestMethod]
        public void RepeatedLocalTeamTemplateObservationsMakeOneNamedObjectWithoutLosingRows()
        {
            var id = Guid.NewGuid(); var f = Fixture(new[] { Team(id), Team(id) });
            var backing = new Entity("teamtemplate", id) { ["teamtemplateid"] = id, ["teamtemplatename"] = "Offline team " + id };
            var cache = ComponentDisplayEvidenceTests.A1NameCache(f.Presentation.Source.Snapshot, 511, new[] { backing });
            f.Completed = new CompletedMembershipComparison(f.Presentation, ComponentDisplayEvidenceProjector.Project(f.Presentation, cache, null), DisplayProjectionOutcome.Completed);
            var p = ExplorerFixture.Project(f); var original = p.Rows.Select(r => r.Original.AuthoritativeRow).ToArray(); var csv = ExplorerFixture.Csv(f);
            var objects = new SolutionObjectNavigation(p).Items(ObjectsWorkspaceScope.Source);
            Assert.AreEqual(2, p.Rows.Count); Assert.AreEqual(2, p.Occurrences.Count); Assert.AreEqual(1, objects.Count);
            Assert.AreEqual("Team template", objects[0].Category); Assert.AreEqual("Team Template", objects[0].Type);
            Assert.AreEqual("Offline team " + id, objects[0].Value(false)); Assert.AreEqual(2, objects[0].Rows.Count);
            CollectionAssert.AreEqual(new[] { 0, 1 }, objects[0].Rows.ToArray()); Assert.AreEqual("Needs Review", objects[0].Status(p));
            for (int i = 0; i < original.Length; i++) Assert.AreSame(original[i], p.Rows[i].Original.AuthoritativeRow);
            Assert.AreEqual(csv, ExplorerFixture.Csv(f));
            TestContext.WriteLine("FIXTURE TeamTemplate: Source raw=2; technical rows=2; Source objects=1; authoritative classifications and CSV unchanged.");
        }
        [TestMethod]
        public void DistinctTemplatesWithEqualNamesStayDistinct()
        {
            var f = Fixture(new[] { Team(Guid.NewGuid()), Team(Guid.NewGuid()) }, displayName: "Same retained name");
            var p = ExplorerFixture.Project(f); var objects = new SolutionObjectNavigation(p).Items(ObjectsWorkspaceScope.Source);
            Assert.AreEqual(2, objects.Count); Assert.AreEqual(2, objects.Select(i => i.Key).Distinct().Count());
            Assert.IsTrue(objects.All(i => i.Value(false) == "Same retained name"));
        }
        [TestMethod]
        public void EqualGuidUnsupportedTemplatesAreNotSilentlyGivenPortableIdentity()
        {
            var id = Guid.NewGuid(); var f = Fixture(new[] { Team(id) }, new[] { Team(id) }); var p = ExplorerFixture.Project(f);
            var model = new SolutionObjectNavigation(p);
            Assert.AreEqual(1, model.Items(ObjectsWorkspaceScope.Source).Count); Assert.AreEqual(1, model.Items(ObjectsWorkspaceScope.Target).Count);
            Assert.AreEqual(2, model.Items(ObjectsWorkspaceScope.Compare).Count);
            Assert.IsTrue(p.Rows.All(r => r.Original.AuthoritativeRow.Presence == MembershipPresence.Indeterminate));
            Assert.IsTrue(model.Items(ObjectsWorkspaceScope.Compare).All(i => i.Status(p) == "Needs Review"));
            TestContext.WriteLine("FIXTURE Unsupported equal GUID: Source objects=1; Target objects=1; Compare unreconciled local objects=2. No new cross-environment identity policy.");
        }
        [TestMethod]
        public void ExistingAuthoritativeSharedRowsProvideCompareUnionWithoutGuidMatching()
        {
            var f = Fixture(new[] { ExplorerFixture.Identity(61, "shared"), ExplorerFixture.Identity(61, "source") },
                new[] { ExplorerFixture.Identity(61, "shared"), ExplorerFixture.Identity(61, "target") });
            var p = ExplorerFixture.Project(f); var model = new SolutionObjectNavigation(p);
            Assert.AreEqual(2, model.Items(ObjectsWorkspaceScope.Source).Count); Assert.AreEqual(2, model.Items(ObjectsWorkspaceScope.Target).Count);
            Assert.AreEqual(3, model.Items(ObjectsWorkspaceScope.Compare).Count);
            Assert.IsTrue(model.Items(ObjectsWorkspaceScope.Source).All(i => i.Has(DisplayEnvironmentSide.Source)));
            Assert.IsFalse(model.Items(ObjectsWorkspaceScope.Source).Any(i => i.Value(false) == "target"));
            var shared = model.Items(ObjectsWorkspaceScope.Compare).Single(i => i.Value(false) == "shared");
            Assert.IsTrue(shared.Has(DisplayEnvironmentSide.Source) && shared.Has(DisplayEnvironmentSide.Target));
            Assert.AreEqual("Present", shared.Presence(DisplayEnvironmentSide.Source, p)); Assert.AreEqual("Present", shared.Presence(DisplayEnvironmentSide.Target, p));
            TestContext.WriteLine("FIXTURE Union: Source objects=2; Target objects=2; Compare objects=3; technical rows=3; raw observations=4.");
        }
        [DataTestMethod, DataRow(false), DataRow(true)]
        public void ChartsNeverBecomeAnIndependentMakerCategoryAndRequireTypedParent(bool verifiedParent)
        {
            var table = ExplorerFixture.Identity(1, "incident");
            var chart = new ComponentIdentity(new SolutionComponentRecord(Guid.NewGuid(), 59, Guid.NewGuid(), rootSolutionComponentId: table.Record.SolutionComponentId),
                IdentityResolutionStatus.Resolved, "chart-key");
            var f = Fixture(new[] { table, chart }, parent: verifiedParent ? "incident" : null); var p = ExplorerFixture.Project(f);
            var roots = new SolutionObjectNavigation(p).Items(ObjectsWorkspaceScope.Source);
            Assert.AreEqual(1, roots.Count); Assert.AreEqual("Tables", roots[0].Category); Assert.IsFalse(roots.Any(i => i.Category == "Charts"));
            Assert.AreEqual(verifiedParent ? 1 : 0, roots[0].Children.Count);
            Assert.AreEqual(2, p.Rows.Count); Assert.AreEqual(2, p.Occurrences.Count);
            if (verifiedParent) { Assert.AreEqual("Chart", roots[0].Children[0].Type); Assert.IsTrue(roots[0].Children[0].Rows.Contains(p.Rows.Single(r => r.Original.SourceRawComponentType == 59).OriginalRowIndex)); }
            TestContext.WriteLine("FIXTURE Charts: technical charts=1; top-level Charts categories=0; Table child charts=" + roots[0].Children.Count + "; inclusion root alone is not ownership.");
        }
        [TestMethod]
        public void ChartParentIsSideOwnedAndNoMissingParentTableIsFabricated()
        {
            var f = Fixture(new[] { ExplorerFixture.Identity(59, "chart") }, new[] { ExplorerFixture.Identity(1, "incident") }, parent: "incident");
            var model = new SolutionObjectNavigation(ExplorerFixture.Project(f));
            Assert.AreEqual(0, model.Items(ObjectsWorkspaceScope.Source).Count);
            Assert.AreEqual(1, model.Items(ObjectsWorkspaceScope.Target).Count);
            Assert.AreEqual(0, model.Items(ObjectsWorkspaceScope.Target)[0].Children.Count);
        }
        [DataTestMethod, DataRow("diagnostic"), DataRow("unavailable"), DataRow("ambiguous"), DataRow("wrong-side")]
        public void UnverifiedNamesAndWrongSideEvidenceNeverFabricateObjects(string state)
        {
            var identity = Team(Guid.NewGuid()); var f = ExplorerFixture.Make(new[] { identity });
            var evidence = Display(identity, state == "wrong-side" ? DisplayEnvironmentSide.Target : DisplayEnvironmentSide.Source,
                state == "unavailable" || state == "diagnostic" ? null : "Name", null,
                state == "ambiguous" ? DisplayBackingCorrelationStatus.Ambiguous : DisplayBackingCorrelationStatus.Unique);
            f.Completed = new CompletedMembershipComparison(f.Presentation, new[] { new ComponentDisplayEvidence(0, evidence, ExplorerFixture.Display(null, DisplayEnvironmentSide.Target)) }, DisplayProjectionOutcome.Completed);
            var p = ExplorerFixture.Project(f); var model = new SolutionObjectNavigation(p);
            Assert.AreEqual(0, model.Items(ObjectsWorkspaceScope.Source).Count); Assert.IsTrue(model.MappingLimited);
            Assert.AreEqual(1, p.Rows.Count); Assert.AreEqual(1, p.Occurrences.Count);
        }
        [DataTestMethod, DataRow("unique"), DataRow("omitted"), DataRow("null"), DataRow("duplicate"), DataRow("failed"), DataRow("conflict")]
        public void ExistingTeamLookupRetainsFriendlyNameWithoutParsingOrAnotherRequest(string state)
        {
            var identity = Team(Guid.NewGuid()); var f = ExplorerFixture.Make(new[] { identity }); var id = identity.Record.ObjectId.Value;
            var row = new Entity("teamtemplate", id) { ["teamtemplateid"] = id, ["teamtemplatename"] = "Runtime supplied name" };
            if (state == "omitted") row.Attributes.Remove("teamtemplatename"); if (state == "null") row["teamtemplatename"] = null;
            if (state == "conflict") row["teamtemplateid"] = Guid.NewGuid();
            var returned = state == "duplicate" ? new[] { row, row } : new[] { row };
            var cache = ComponentDisplayEvidenceTests.A1NameCache(f.Presentation.Source.Snapshot, 511, returned, state == "failed");
            var display = ComponentDisplayEvidenceProjector.Project(f.Presentation, cache, null)[0].SourceDisplayEvidence;
            Assert.AreEqual(IdentityResolutionStatus.Unsupported, identity.Status); Assert.IsNull(identity.ComparisonKey);
            if (state == "unique")
            { Assert.AreEqual("Runtime supplied name", display.FriendlyName.Value); Assert.AreEqual("NormalMembershipDiagnostic.teamtemplate.teamtemplatename", display.FriendlyName.DisplayEvidenceSource); }
            else if (state == "null") { Assert.AreEqual(DisplayAvailability.Available, display.FriendlyName.Availability); Assert.IsNull(display.FriendlyName.Value); }
            else Assert.AreNotEqual(DisplayAvailability.Available, display.FriendlyName.Availability);
            var copy = MembershipSnapshot.Complete(f.Presentation.Source.Snapshot.Solution, f.Presentation.Source.Snapshot.Components, f.Presentation.Source.Snapshot.CapturedAt);
            DataverseComponentMetadataCache.NormalDisplayNameObservation observation;
            Assert.IsFalse(cache.TryGetNormalDisplayName(copy, identity, out observation));
            Assert.IsFalse(cache.TryGetNormalDisplayName(f.Presentation.Source.Snapshot, Team(id), out observation));
        }
        [TestMethod]
        public void ExistingChartLookupRetainsTypedTableContextWithoutChangingIdentity()
        {
            var chart = ExplorerFixture.Identity(59, "chart"); var f = ExplorerFixture.Make(new[] { chart }); var id = chart.Record.ObjectId.Value;
            var row = new Entity("savedqueryvisualization", id) { ["savedqueryvisualizationid"] = id, ["name"] = "Retained chart", ["primaryentitytypecode"] = "incident" };
            var cache = ComponentDisplayEvidenceTests.A1NameCache(f.Presentation.Source.Snapshot, 59, new[] { row });
            var display = ComponentDisplayEvidenceProjector.Project(f.Presentation, cache, null)[0].SourceDisplayEvidence;
            Assert.AreEqual("Retained chart", display.FriendlyName.Value); Assert.AreEqual("incident", display.ParentDisplayContext.Value);
            Assert.AreEqual("NormalMembershipDiagnostic.savedqueryvisualization.primaryentitytypecode", display.ParentDisplayContext.DisplayEvidenceSource);
            Assert.AreEqual("chart", chart.ComparisonKey);
        }
        [TestMethod]
        public void ObjectUiPreservesSearchScopeSelectionAndNeverCallsFullFormatter()
        {
            Sta(() =>
            {
                var f = Fixture(new[] { ExplorerFixture.Identity(61, "source"), ExplorerFixture.Identity(61, "shared") },
                    new[] { ExplorerFixture.Identity(61, "target"), ExplorerFixture.Identity(61, "shared") }); int formats = 0;
                using (var host = new Form { Size = new Size(1250, 800) })
                using (var panel = new SimplifiedSolutionExplorerPanel(f.Presentation, row => { formats++; return "secret"; }))
                {
                    panel.UseObjectNavigation(); host.Controls.Add(panel); host.Show(); panel.Bind(ExplorerFixture.Project(f));
                    Assert.AreEqual(3, panel.Explorer.VisibleObjects.Count);
                    Item<ToolStripComboBox>(panel, "SolutionExplorerScope").SelectedIndex = 1;
                    Assert.AreEqual(2, panel.Explorer.VisibleObjects.Count); Assert.IsTrue(panel.Explorer.VisibleObjects.All(i => i.Has(DisplayEnvironmentSide.Source)));
                    var row = panel.Projection.Rows.Single(r => r.Original.AuthoritativeRow.SourceIdentity?.ComparisonKey == "source");
                    panel.Explorer.SelectComparison(row.OriginalRowIndex, DisplayEnvironmentSide.Source);
                    var original = panel.Explorer.SelectedComparison; Assert.IsNotNull(original);
                    Item<ToolStripComboBox>(panel, "SolutionExplorerScope").SelectedIndex = 2;
                    Assert.AreEqual(2, panel.Explorer.VisibleObjects.Count); Assert.AreSame(original, panel.Explorer.SelectedComparison);
                    Assert.AreEqual(DisplayEnvironmentSide.Source, panel.Explorer.SelectedSide);
                    Item<ToolStripComboBox>(panel, "SolutionExplorerScope").SelectedIndex = 1;
                    Assert.AreSame(original, panel.Explorer.SelectedComparison);
                    Item<ToolStripTextBox>(panel, "SolutionExplorerSearch").Text = "no results";
                    Assert.AreEqual(0, panel.Explorer.VisibleObjects.Count);
                    Item<ToolStripButton>(panel, "SolutionExplorerReset").PerformClick();
                    Assert.AreEqual(2, panel.Explorer.VisibleObjects.Count); Assert.IsNull(panel.Explorer.SelectedObject);
                    Assert.AreEqual(1, Item<ToolStripComboBox>(panel, "SolutionExplorerScope").SelectedIndex); Assert.AreEqual(0, formats);
                    Assert.IsFalse(panel.Explorer.Navigation.Nodes.Cast<TreeNode>().Any(n => n.Text.Contains("Unmapped") || n.Text.Contains("Needs Review") || n.Text.StartsWith("Charts")));
                }
            });
        }
        [TestMethod]
        public void ObjectUiTeamDetailsDoNotChooseAnArbitraryTechnicalRow()
        {
            Sta(() =>
            {
                var id = Guid.NewGuid(); var f = Fixture(new[] { Team(id), Team(id) });
                using (var host = new Form { Size = new Size(1250, 800) })
                using (var panel = new SimplifiedSolutionExplorerPanel(f.Presentation, row => { Assert.Fail("No technical formatting"); return ""; }))
                {
                    panel.UseObjectNavigation(); host.Controls.Add(panel); host.Show(); panel.Bind(ExplorerFixture.Project(f));
                    var category = panel.Explorer.Navigation.Nodes.Cast<TreeNode>().Single(n => n.Text.StartsWith("Team template")); category.Expand();
                    Assert.AreEqual(1, category.Nodes.Count); panel.Explorer.Navigation.SelectedNode = category.Nodes[0];
                    Assert.AreEqual(1, panel.Explorer.VisibleObjects.Count); Assert.IsNotNull(panel.Explorer.SelectedObject); Assert.IsNull(panel.Explorer.SelectedComparison);
                    var identity = Descendants(panel).OfType<TextBox>().Single(c => c.Name == "SolutionExplorerSelectedIdentity");
                    StringAssert.Contains(identity.Text, "Offline team " + id); StringAssert.Contains(identity.Text, "Team Template");
                    var properties = Descendants(panel).OfType<DataGridView>().Single(c => c.Name == "ExplorerComparedProperties");
                    Assert.AreEqual(0, properties.RowCount); Assert.IsNull(properties.Tag); Assert.AreEqual(2, panel.Projection.Rows.Count);
                }
            });
        }
        [TestMethod]
        public void TenThousandObjectsRemainVirtualWithBoundedNativePages()
        {
            Sta(() =>
            {
                var f = Fixture(Enumerable.Range(0, 10000).Select(i => ExplorerFixture.Identity(61, "resource" + i)).ToArray());
                using (var host = new Form { Size = new Size(1250, 800) })
                using (var panel = new SimplifiedSolutionExplorerPanel(f.Presentation, row => ""))
                {
                    panel.UseObjectNavigation(); host.Controls.Add(panel); host.Show(); panel.Bind(ExplorerFixture.Project(f));
                    Assert.AreEqual(10000, panel.Explorer.VisibleObjects.Count); Assert.IsTrue(panel.Explorer.Grid.VirtualMode);
                    var web = panel.Explorer.Navigation.Nodes.Cast<TreeNode>().Single(n => n.Text.StartsWith("Web resources")); web.Expand();
                    Assert.AreEqual(50, web.Nodes.Count); web.Nodes[0].Expand(); Assert.AreEqual(200, web.Nodes[0].Nodes.Count);
                    Assert.IsTrue(Nodes(panel.Explorer.Navigation.Nodes).Count(n => n.Name.Length != 0) < 300);
                    Assert.AreEqual(200, Nodes(panel.Explorer.Navigation.Nodes).Count(n => n.Tag is SolutionObjectNavigation.Item), "Only one page of object nodes is materialized; expand sentinels are not objects.");
                    Assert.AreEqual(10000, panel.Projection.Rows.Count);
                    TestContext.WriteLine("FIXTURE 10000: 10000 objects/technical rows; 50 native pages; first page=200; virtual grid; no service dependency.");
                }
            });
        }

        [TestMethod]
        public void ChartChildCountsAndSearchRespectActiveInventory()
        {
            Sta(() =>
            {
                var f = Fixture(new[] { ExplorerFixture.Identity(1, "incident"), ExplorerFixture.Identity(59, "shared-a"), ExplorerFixture.Identity(59, "shared-b"), ExplorerFixture.Identity(59, "source-child") },
                    new[] { ExplorerFixture.Identity(1, "incident"), ExplorerFixture.Identity(59, "shared-a"), ExplorerFixture.Identity(59, "shared-b") }, parent: "incident");
                var p = ExplorerFixture.Project(f); var model = new SolutionObjectNavigation(p);
                Assert.AreEqual(3, model.Items(ObjectsWorkspaceScope.Source).Single().Children.Count);
                Assert.AreEqual(2, model.Items(ObjectsWorkspaceScope.Target).Single().Children.Count);
                Assert.AreEqual(3, model.Items(ObjectsWorkspaceScope.Compare).Single().Children.Count);
                using (var host = new Form { Size = new Size(1250, 800) })
                using (var panel = new SimplifiedSolutionExplorerPanel(f.Presentation, row => ""))
                {
                    panel.UseObjectNavigation(); host.Controls.Add(panel); host.Show(); panel.Bind(p);
                    Item<ToolStripTextBox>(panel, "SolutionExplorerSearch").Text = "source-child";
                    var tables = panel.Explorer.Navigation.Nodes.Cast<TreeNode>().Single(n => n.Text.StartsWith("Tables")); tables.Expand();
                    Assert.AreEqual(1, tables.Nodes.Count); tables.Nodes[0].Expand();
                    var charts = tables.Nodes[0].Nodes.Cast<TreeNode>().Single(n => n.Text.StartsWith("Charts")); charts.Expand();
                    Assert.AreEqual(1, charts.Nodes.Count); panel.Explorer.Navigation.SelectedNode = charts.Nodes[0];
                    Assert.AreEqual("source-child", panel.Explorer.SelectedObject.Value(false));
                    Assert.AreEqual("Source Only", panel.Explorer.SelectedObject.Status(p));
                    Item<ToolStripComboBox>(panel, "SolutionExplorerScope").SelectedIndex = 2;
                    Assert.IsFalse(panel.Explorer.VisibleObjects.Any(i => i.Value(false) == "source-child"));
                    Assert.AreEqual(4, p.Rows.Count); Assert.AreEqual(7, p.Occurrences.Count);
                }
                TestContext.WriteLine("FIXTURE Charts: Source children=3; Target children=2; Compare children=3; top-level Tables objects=1; charts not included in top-level All.");
            });
        }
        [DataTestMethod, DataRow("Match"), DataRow("Different")]
        public void SelectedObjectPropertyPaneKeepsExactComparisonOwnerAndValues(string status)
        {
            Sta(() =>
            {
                var f = Fixture(new[] { ExplorerFixture.Identity(61, "resource") }, new[] { ExplorerFixture.Identity(61, "resource") });
                var previous = f.Completed; var original = f.Presentation.Rows[0];
                var definition = new D365SolutionComparer.Models.ComponentDetails.ComponentDefinitionDetailPresentation("Web resource", "resource", original.MembershipStatus, status,
                    new[] { new D365SolutionComparer.Models.ComponentDetails.ComponentPropertyPresentation("value", "left", status == "Match" ? "left" : "right", status == "Different", status) }, "Retained internal diagnostic");
                f.Presentation = new MembershipComparisonPresentation("SAA", f.Presentation.Source, f.Presentation.Target, new[] { original.WithDefinition(status, status == "Different" ? "value" : "", definition) });
                f.Completed = new CompletedMembershipComparison(f.Presentation, previous.DisplayEvidence, DisplayProjectionOutcome.Completed);
                using (var host = new Form { Size = new Size(1250, 800) })
                using (var panel = new SimplifiedSolutionExplorerPanel(f.Presentation, row => { Assert.Fail("No full evidence disclosure"); return ""; }))
                {
                    panel.UseObjectNavigation(); host.Controls.Add(panel); host.Show(); panel.Bind(ExplorerFixture.Project(f)); panel.Explorer.SelectComparison(0);
                    var properties = Descendants(panel).OfType<DataGridView>().Single(c => c.Name == "ExplorerComparedProperties");
                    Assert.AreSame(definition, properties.Tag); Assert.AreSame(f.Presentation.Rows[0], panel.Explorer.SelectedComparison.AuthoritativeRow);
                    Assert.AreEqual("left", properties.Rows[0].Cells[1].Value); Assert.AreEqual(status == "Match" ? "left" : "right", properties.Rows[0].Cells[2].Value);
                    Assert.AreEqual(status, panel.Explorer.SelectedObject.Status(panel.Projection));
                    Item<ToolStripTextBox>(panel, "SolutionExplorerSearch").Text = "no results";
                    Assert.AreEqual(0, panel.Explorer.VisibleObjects.Count); Assert.IsNull(panel.Explorer.SelectedObject); Assert.AreEqual(0, properties.RowCount); Assert.IsNull(properties.Tag);
                }
            });
        }
        [TestMethod]
        public void DirectEmptyObjectFilterClearsDisplayedPropertiesAndOriginalRowsRemainIntact()
        {
            Sta(() =>
            {
                var f = Fixture(new[] { ExplorerFixture.Identity(61, "retained") });
                using (var panel = new SimplifiedSolutionExplorerPanel(f.Presentation, row => ""))
                {
                    panel.UseObjectNavigation(); panel.Bind(ExplorerFixture.Project(f)); panel.Explorer.SelectComparison(0);
                    Assert.IsNotNull(panel.Explorer.SelectedObject); panel.Explorer.ApplyFilter(new int[0], false);
                    Assert.AreEqual(0, panel.Explorer.VisibleObjects.Count); Assert.IsNull(panel.Explorer.SelectedObject);
                    Assert.IsNull(panel.Explorer.SelectedComparison); Assert.AreEqual("", panel.Explorer.EvidenceText);
                    Assert.AreEqual(1, panel.Projection.Rows.Count); Assert.AreSame(f.Presentation.Rows[0], panel.Projection.Rows[0].Original.AuthoritativeRow);
                }
            });
        }

        [TestMethod]
        public void RebindingAnObjectControlClearsPriorOwnerSelectionAndDisplayedText()
        {
            Sta(() =>
            {
                var one = Fixture(new[] { ExplorerFixture.Identity(61, "old-owner") }); var two = Fixture(new[] { ExplorerFixture.Identity(61, "new-owner") });
                using (var panel = new SimplifiedSolutionExplorerPanel(one.Presentation, row => ""))
                {
                    panel.UseObjectNavigation(); panel.Bind(ExplorerFixture.Project(one)); panel.Explorer.SelectComparison(0);
                    Assert.IsNotNull(panel.Explorer.SelectedObject);
                    panel.Explorer.Bind(ExplorerFixture.Project(two), new[] { 0 }, true);
                    Assert.IsNull(panel.Explorer.SelectedObject); Assert.IsNull(panel.Explorer.SelectedComparison); Assert.AreEqual("", panel.Explorer.EvidenceText);
                    var identity = Descendants(panel).OfType<TextBox>().Single(c => c.Name == "SolutionExplorerSelectedIdentity");
                    Assert.IsFalse(identity.Text.Contains("old-owner"));
                }
            });
        }
        [DataTestMethod, DataRow(1000, 620, 9F), DataRow(1300, 850, 12F), DataRow(1500, 1000, 16F)]
        public void MakerDetailsKeepOneCompletePropertyRowOrExplicitExpand(int width, int height, float font)
        {
            Sta(() =>
            {
                var f = Fixture(new[] { ExplorerFixture.Identity(61, "resource") }, new[] { ExplorerFixture.Identity(61, "resource") });
                var old = f.Completed; var row = f.Presentation.Rows[0];
                var detail = new D365SolutionComparer.Models.ComponentDetails.ComponentDefinitionDetailPresentation("Web resource", "resource", row.MembershipStatus, "Different",
                    new[] { new D365SolutionComparer.Models.ComponentDetails.ComponentPropertyPresentation("value", "left", "right", true, "Different") }, "Retained diagnostic");
                f.Presentation = new MembershipComparisonPresentation("SAA", f.Presentation.Source, f.Presentation.Target, new[] { row.WithDefinition("Different", "value", detail) });
                f.Completed = new CompletedMembershipComparison(f.Presentation, old.DisplayEvidence, DisplayProjectionOutcome.Completed);
                using (var host = new Form { Size = new Size(width, height), Font = new Font("Segoe UI", font) })
                using (var panel = new SimplifiedSolutionExplorerPanel(f.Presentation, r => ""))
                {
                    panel.UseObjectNavigation(); host.Controls.Add(panel); host.Show(); panel.Bind(ExplorerFixture.Project(f)); panel.Explorer.SelectComparison(0); Application.DoEvents();
                    var properties = Descendants(panel).OfType<DataGridView>().Single(c => c.Name == "ExplorerComparedProperties");
                    int required = properties.ColumnHeadersHeight + properties.RowTemplate.Height;
                    var cue = Descendants(panel).OfType<System.Windows.Forms.Label>().Single(c => c.Name == "ExplorerDetailsScrollCue");
                    var expand = Descendants(panel).OfType<LinkLabel>().Single(c => c.Name == "ExplorerExpandDetails");
                    Assert.IsTrue(properties.ClientSize.Height >= required || cue.Visible && expand.Visible);
                    typeof(LinkLabel).GetMethod("OnLinkClicked", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic)
                        .Invoke(expand, new object[] { new LinkLabelLinkClickedEventArgs(expand.Links[0]) }); Application.DoEvents();
                    Assert.IsTrue(properties.ClientSize.Height >= required); Assert.AreSame(detail, properties.Tag);
                    Assert.AreEqual("value", properties.Rows[0].Cells[0].Value); Assert.AreEqual("right", properties.Rows[0].Cells[2].Value);
                    Assert.IsTrue(properties.ReadOnly && properties.VirtualMode);
                }
            });
        }

        [TestMethod]
        public void UnavailableInventoryDoesNotBecomeAZeroObjectClaimOrConfirmedAbsence()
        {
            Sta(() =>
            {
                var f = Fixture(new ComponentIdentity[0], new[] { ExplorerFixture.Identity(61, "target") }); var old = f.Completed;
                f.Presentation = new MembershipResultPresenter().Create(MembershipEnvironmentResult.Unavailable("Source", "SAA", 0, TimeSpan.Zero, "Offline synthetic unavailable"), f.Presentation.Target);
                f.Completed = new CompletedMembershipComparison(f.Presentation, old.DisplayEvidence, DisplayProjectionOutcome.Completed);
                using (var panel = new SimplifiedSolutionExplorerPanel(f.Presentation, row => ""))
                {
                    panel.UseObjectNavigation(); panel.Bind(ExplorerFixture.Project(f));
                    Assert.AreEqual(1, panel.Explorer.VisibleObjects.Count);
                    Assert.AreEqual("Needs Review", panel.Explorer.VisibleObjects[0].Status(panel.Projection));
                    Assert.AreEqual("Not established", panel.Explorer.VisibleObjects[0].Presence(DisplayEnvironmentSide.Source, panel.Projection));
                    Item<ToolStripComboBox>(panel, "SolutionExplorerScope").SelectedIndex = 1;
                    StringAssert.Contains(panel.Explorer.ObjectSummary, "unavailable");
                    StringAssert.Contains(panel.Explorer.Navigation.Nodes[0].Text, "not available");
                    Assert.IsFalse(panel.Explorer.ObjectSummary.Contains("0 objects"));
                }
            });
        }

        [TestMethod]
        public void ProductionEntryUsesObjectNavigationAndPreservesCompletedTechnicalInventory()
        {
            Sta(() =>
            {
                var f = Fixture(new[] { ExplorerFixture.Identity(1, "incident"), ExplorerFixture.Identity(59, "chart"),
                    ExplorerFixture.Identity(61, "resource"), ExplorerFixture.Identity(300, null, status: IdentityResolutionStatus.Unsupported) }, parent: "incident");
                var spy = new FakeOrganizationService(); int capture = 0;
                using (var form = new MembershipResultsForm(f.Completed, captureAppSettingEvidence: () => capture++, sourceService: spy, targetService: spy))
                {
                    form.UseSimplifiedExplorer(); form.Show(); var clock = Stopwatch.StartNew();
                    while (form.ExplorerProjection == null && clock.ElapsedMilliseconds < 10000) { Application.DoEvents(); Thread.Sleep(5); }
                    Assert.IsNotNull(form.ExplorerProjection); var panel = form.Controls.OfType<SimplifiedSolutionExplorerPanel>().Single();
                    Assert.IsTrue(panel.Explorer.UsesObjectNavigation); Assert.AreEqual(2, panel.Explorer.VisibleObjects.Count);
                    Assert.AreEqual(4, panel.Projection.Rows.Count); Assert.AreSame(f.Presentation, panel.Projection.Owner);
                    Assert.IsFalse(panel.Explorer.Navigation.Nodes.Cast<TreeNode>().Any(n => n.Text.StartsWith("Charts") || n.Text.Contains("Unmapped") || n.Text.Contains("relationship")));
                    Assert.AreEqual(0, capture + spy.Calls + spy.ExecuteCalls + spy.WriteCalls);
                }
            });
        }


        [DataTestMethod, DataRow(false), DataRow(true)]
        public void OptionAVerifiedCloudFlowUsesRetainedNameAndExactSide(bool target)
        {
            var identity = OptionAIdentity(29, "cloud-key", category: 5);
            var f = ExplorerFixture.Make(target ? new ComponentIdentity[0] : new[] { identity }, target ? new[] { identity } : null);
            var side = target ? DisplayEnvironmentSide.Target : DisplayEnvironmentSide.Source;
            var cache = new DataverseComponentMetadataCache(); Guid id = identity.Record.ObjectId.Value;
            cache.WorkflowDefinitions[id] = new Entity("workflow", id) { ["workflowid"] = id, ["name"] = "Retained cloud name", ["uniquename"] = "technical_flow", ["category"] = new OptionSetValue(5) };
            var csv = ExplorerFixture.Csv(f); var original = f.Presentation.Rows.ToArray();
            f.Completed = new CompletedMembershipComparison(f.Presentation,
                ComponentDisplayEvidenceProjector.Project(f.Presentation, target ? null : cache, target ? cache : null), DisplayProjectionOutcome.Completed);
            var p = ExplorerFixture.Project(f); var model = new SolutionObjectNavigation(p);
            var item = model.Items(target ? ObjectsWorkspaceScope.Target : ObjectsWorkspaceScope.Source).Single();
            Assert.AreEqual("Cloud flows", item.Category); Assert.AreEqual("Cloud flow", item.Type);
            Assert.AreEqual("Retained cloud name", item.Value(false)); Assert.AreEqual("technical_flow", item.Value(true));
            Assert.IsTrue(item.Has(side)); Assert.IsFalse(item.Has(target ? DisplayEnvironmentSide.Source : DisplayEnvironmentSide.Target));
            Assert.AreEqual(0, model.Items(target ? ObjectsWorkspaceScope.Source : ObjectsWorkspaceScope.Target).Count);
            Assert.AreSame(identity, target ? p.Rows[0].Original.AuthoritativeRow.TargetIdentity : p.Rows[0].Original.AuthoritativeRow.SourceIdentity);
            Assert.AreEqual("cloud-key", identity.ComparisonKey); Assert.AreEqual(5, identity.WorkflowCategory);
            OptionAPreserved(f, p, original, csv);
        }

        [DataTestMethod, DataRow("missing-category"), DataRow("workflow"), DataRow("bpf"), DataRow("unknown-category")]
        [DataRow("missing-name"), DataRow("empty-name"), DataRow("wrong-provenance"), DataRow("missing-backing")]
        [DataRow("ambiguous-backing"), DataRow("wrong-side"), DataRow("wrong-record"), DataRow("unsupported"), DataRow("unresolved"), DataRow("wrong-kind")]
        public void OptionACloudFlowLabelRequiresTypedCategoryAndOwnedName(string state)
        {
            int? category = state == "missing-category" ? (int?)null : state == "workflow" ? 0 : state == "bpf" ? 4 : state == "unknown-category" ? 999 : 5;
            var status = state == "unsupported" ? IdentityResolutionStatus.Unsupported : state == "unresolved" ? IdentityResolutionStatus.Unresolved : IdentityResolutionStatus.Resolved;
            var identity = OptionAIdentity(29, "cloud-key", status, category, kind: state == "wrong-kind" ? ComponentSemanticKinds.WebResource : null);
            var f = ExplorerFixture.Make(new[] { identity }); var csv = ExplorerFixture.Csv(f); var original = f.Presentation.Rows.ToArray();
            var evidence = OptionADisplay(identity, state == "wrong-side" ? DisplayEnvironmentSide.Target : DisplayEnvironmentSide.Source,
                state == "empty-name" ? "" : "Looks like Cloud flow Business Process", state == "missing-backing" ? DisplayBackingCorrelationStatus.Unavailable :
                state == "ambiguous-backing" ? DisplayBackingCorrelationStatus.Ambiguous : DisplayBackingCorrelationStatus.Unique,
                state == "wrong-provenance" ? "template.title" : "workflow.name", state == "missing-name");
            if (state == "wrong-record") evidence = OptionADisplay(OptionAIdentity(29, "foreign", category: 5), DisplayEnvironmentSide.Source, "Looks like Cloud flow Business Process");
            f.Completed = new CompletedMembershipComparison(f.Presentation, new[] { new ComponentDisplayEvidence(0, evidence, ExplorerFixture.Display(null, DisplayEnvironmentSide.Target)) }, DisplayProjectionOutcome.Completed);
            var p = ExplorerFixture.Project(f); var items = new SolutionObjectNavigation(p).Items(ObjectsWorkspaceScope.Compare);
            Assert.AreEqual(0, items.Count, "A display-name hint is not verified category/name/owner evidence: " + state);
            Assert.AreEqual(status, identity.Status); OptionAPreserved(f, p, original, csv);
        }

        [DataTestMethod, DataRow(91, "Plug-in assemblies", "Plug-in assembly"), DataRow(92, "Plug-in steps", "Plug-in step")]
        public void OptionAPluginCategoriesPreserveOriginalRowsAndTechnicalEvidence(int raw, string category, string type)
        {
            var identity = OptionAIdentity(raw, "original-portable-key"); var f = OptionAFixture(new[] { identity });
            var original = f.Presentation.Rows.ToArray(); var csv = ExplorerFixture.Csv(f); var display = f.Completed.DisplayEvidence[0];
            var p = ExplorerFixture.Project(f); var item = new SolutionObjectNavigation(p).Items(ObjectsWorkspaceScope.Source).Single();
            Assert.AreEqual(category, item.Category); Assert.AreEqual(type, item.Type); Assert.AreEqual("Same retained name", item.Value(false));
            Assert.AreSame(identity, p.Rows[0].Original.AuthoritativeRow.SourceIdentity); Assert.AreEqual("original-portable-key", identity.ComparisonKey);
            Assert.AreSame(display.SourceDisplayEvidence, p.Rows[0].SourceDisplay); CollectionAssert.AreEqual(new[] { 0 }, item.Rows.ToArray());
            Assert.AreEqual(1, item.Locals[0].Occurrences.Count); Assert.AreEqual("Present", item.Presence(DisplayEnvironmentSide.Source, p));
            OptionAPreserved(f, p, original, csv);
        }

        [DataTestMethod, DataRow(91), DataRow(92)]
        public void OptionAEqualPluginNamesAndLocalIdsDoNotMatchUnsupportedCounterparts(int raw)
        {
            Guid id = Guid.NewGuid(); var source = OptionAIdentity(raw, null, IdentityResolutionStatus.Unsupported, id: id);
            var target = OptionAIdentity(raw, null, IdentityResolutionStatus.Unsupported, id: id); var f = OptionAFixture(new[] { source }, new[] { target });
            var original = f.Presentation.Rows.ToArray(); var csv = ExplorerFixture.Csv(f); var p = ExplorerFixture.Project(f); var model = new SolutionObjectNavigation(p);
            Assert.AreEqual(2, model.Items(ObjectsWorkspaceScope.Compare).Count); Assert.AreEqual(1, model.Items(ObjectsWorkspaceScope.Source).Count);
            Assert.AreEqual(1, model.Items(ObjectsWorkspaceScope.Target).Count); Assert.AreEqual(2, p.Rows.Count);
            Assert.IsTrue(p.Rows.All(r => r.Original.AuthoritativeRow.Presence == MembershipPresence.Indeterminate));
            Assert.IsTrue(model.Items(ObjectsWorkspaceScope.Compare).All(i => i.Status(p) == "Needs Review"));
            Assert.IsTrue(model.Items(ObjectsWorkspaceScope.Source).All(i => i.Has(DisplayEnvironmentSide.Source) && !i.Has(DisplayEnvironmentSide.Target)));
            Assert.IsNull(source.ComparisonKey); Assert.IsNull(target.ComparisonKey); OptionAPreserved(f, p, original, csv);
        }

        [DataTestMethod, DataRow(91), DataRow(92)]
        public void OptionASharedPluginsJoinOnlyTheExistingAuthoritativeResolvedRow(int raw)
        {
            var source = OptionAIdentity(raw, "shared-key"); var target = OptionAIdentity(raw, "shared-key");
            Assert.AreNotEqual(source.Record.ObjectId, target.Record.ObjectId); var f = OptionAFixture(new[] { source }, new[] { target });
            var original = f.Presentation.Rows.ToArray(); var csv = ExplorerFixture.Csv(f); var p = ExplorerFixture.Project(f); var model = new SolutionObjectNavigation(p);
            Assert.AreEqual(1, p.Rows.Count); Assert.AreEqual(2, p.Occurrences.Count); var item = model.Items(ObjectsWorkspaceScope.Compare).Single();
            Assert.IsTrue(item.Has(DisplayEnvironmentSide.Source) && item.Has(DisplayEnvironmentSide.Target));
            Assert.AreSame(source, p.Rows[0].Original.AuthoritativeRow.SourceIdentity); Assert.AreSame(target, p.Rows[0].Original.AuthoritativeRow.TargetIdentity);
            Assert.AreEqual(MembershipPresence.PresentInBoth, original[0].Presence);
            Assert.AreEqual(1, model.Items(ObjectsWorkspaceScope.Source).Count); Assert.AreEqual(1, model.Items(ObjectsWorkspaceScope.Target).Count);
            OptionAPreserved(f, p, original, csv);
        }

        [DataTestMethod, DataRow(91, "no-name"), DataRow(92, "no-name"), DataRow(91, "ambiguous"), DataRow(92, "ambiguous")]
        [DataRow(91, "wrong-side"), DataRow(92, "wrong-side")]
        public void OptionAUnavailablePluginObservationsDoNotFabricateObjects(int raw, string state)
        {
            var identity = OptionAIdentity(raw, null, IdentityResolutionStatus.Unsupported); var f = ExplorerFixture.Make(new[] { identity });
            var original = f.Presentation.Rows.ToArray(); var csv = ExplorerFixture.Csv(f);
            var display = OptionADisplay(identity, state == "wrong-side" ? DisplayEnvironmentSide.Target : DisplayEnvironmentSide.Source,
                "Same retained name", state == "ambiguous" ? DisplayBackingCorrelationStatus.Ambiguous : DisplayBackingCorrelationStatus.Unique,
                omitName: state == "no-name");
            f.Completed = new CompletedMembershipComparison(f.Presentation, new[] { new ComponentDisplayEvidence(0, display, ExplorerFixture.Display(null, DisplayEnvironmentSide.Target)) }, DisplayProjectionOutcome.Completed);
            var p = ExplorerFixture.Project(f); var model = new SolutionObjectNavigation(p);
            Assert.AreEqual(0, model.Items(ObjectsWorkspaceScope.Compare).Count); Assert.IsTrue(model.MappingLimited);
            Assert.AreEqual(IdentityResolutionStatus.Unsupported, identity.Status); Assert.IsNull(identity.ComparisonKey); OptionAPreserved(f, p, original, csv);
        }

        [TestMethod]
        public void OptionACategoryAndScopeRefreshKeepExactRowAndOccurrenceReferences()
        {
            var cloud = OptionAIdentity(29, "flow", category: 5); var assembly = OptionAIdentity(91, "assembly"); var step = OptionAIdentity(92, "step");
            var target = OptionAIdentity(92, "target-step"); var f = OptionAFixture(new[] { cloud, assembly, step }, new[] { target });
            var original = f.Presentation.Rows.ToArray(); var csv = ExplorerFixture.Csv(f); var p = ExplorerFixture.Project(f); var model = new SolutionObjectNavigation(p);
            var before = model.Items(ObjectsWorkspaceScope.Compare).Select(i => i.Key).ToArray();
            Assert.AreEqual(4, before.Length); Assert.AreEqual(3, model.Items(ObjectsWorkspaceScope.Source).Count); Assert.AreEqual(1, model.Items(ObjectsWorkspaceScope.Target).Count);
            foreach (var category in new[] { "Cloud flows", "Plug-in assemblies", "Plug-in steps" })
            {
                var selected = model.Items(ObjectsWorkspaceScope.Source).Single(i => i.Category == category);
                int row = selected.Rows.Single(); var exact = p.Rows[row].Original.AuthoritativeRow;
                var filtered = model.Items(ObjectsWorkspaceScope.Source).Where(i => i.Rows.Contains(row)).Single();
                Assert.AreEqual(selected.Key, filtered.Key); Assert.AreSame(exact, p.Rows[filtered.Rows.Single()].Original.AuthoritativeRow);
                CollectionAssert.AreEqual(selected.Locals[0].Occurrences.ToArray(), filtered.Locals[0].Occurrences.ToArray());
                Assert.IsTrue(filtered.Has(DisplayEnvironmentSide.Source)); Assert.IsFalse(filtered.Has(DisplayEnvironmentSide.Target));
                Assert.AreEqual(0, model.Items(ObjectsWorkspaceScope.Source).Count(i => i.Value(false).Contains("no results")));
            }
            CollectionAssert.AreEqual(before, model.Items(ObjectsWorkspaceScope.Compare).Select(i => i.Key).ToArray()); OptionAPreserved(f, p, original, csv);
        }

        [DataTestMethod, DataRow(29), DataRow(91), DataRow(92)]
        public void OptionANewCategoryUsesUnchangedProfessionalSelectionMetadata(int raw)
        {
            var f = OptionAFixture(new[] { OptionAIdentity(raw, "key", category: raw == 29 ? 5 : (int?)null) }); var p = ExplorerFixture.Project(f);
            var item = new SolutionObjectNavigation(p).Items(ObjectsWorkspaceScope.Source).Single(); var status = item.Status(p);
            var style = new DataGridViewCellStyle(); SolutionObjectsExplorerControl.ProfessionalExplorerCellStyle(style, status, true, false);
            Assert.AreEqual(Color.FromArgb(0x1D, 0x4E, 0xD8), style.SelectionBackColor); Assert.AreEqual(Color.White, style.SelectionForeColor);
            Assert.AreEqual(status, item.Status(p)); Assert.IsTrue(item.Rows.All(i => ReferenceEquals(f.Presentation.Rows[i], p.Rows[i].Original.AuthoritativeRow)));
        }

        [DataTestMethod, DataRow(91), DataRow(92)]
        public void OptionAExistingPluginProjectionFeedsNamesWithoutNewRetrieval(int raw)
        {
            var identity = OptionAIdentity(raw, "plugin-key"); var f = ExplorerFixture.Make(new[] { identity }); var cache = new DataverseComponentMetadataCache();
            Guid id = identity.Record.ObjectId.Value;
            if (raw == 91) cache.StoreEntityRow(new Entity("pluginassembly", id) { ["name"] = "Retained assembly" }, id);
            else cache.Type92Evidence = new Type92EnvironmentEvidence(f.Presentation.Source.Snapshot, "", new[] { identity.Record },
                new[] { new Type92StepEvidence(id, 1, "Unique", "") { Name = "Retained step" } }, new Dictionary<string, int>());
            var original = f.Presentation.Rows.ToArray(); var csv = ExplorerFixture.Csv(f);
            f.Completed = new CompletedMembershipComparison(f.Presentation, ComponentDisplayEvidenceProjector.Project(f.Presentation, cache, null), DisplayProjectionOutcome.Completed);
            var p = ExplorerFixture.Project(f); var item = new SolutionObjectNavigation(p).Items(ObjectsWorkspaceScope.Source).Single();
            Assert.AreEqual(raw == 91 ? "Retained assembly" : "Retained step", item.Value(false));
            Assert.AreEqual(raw == 91 ? "Plug-in assembly" : "Plug-in step", item.Type); OptionAPreserved(f, p, original, csv);
        }

        private static ComponentIdentity OptionAIdentity(int raw, string key, IdentityResolutionStatus status = IdentityResolutionStatus.Resolved,
            int? category = null, Guid? id = null, string kind = null) => new ComponentIdentity(new SolutionComponentRecord(Guid.NewGuid(), raw, id ?? Guid.NewGuid()),
                status, status == IdentityResolutionStatus.Resolved ? key : null, componentTypeKey: kind ?? (raw == 29 ? ComponentSemanticKinds.Process :
                raw == 91 ? ComponentSemanticKinds.PluginAssembly : ComponentSemanticKinds.SdkMessageProcessingStep)).WithWorkflowCategory(category);
        private static ComponentSideDisplayEvidence OptionADisplay(ComponentIdentity identity, DisplayEnvironmentSide side, string name,
            DisplayBackingCorrelationStatus correlation = DisplayBackingCorrelationStatus.Unique, string provenance = null, bool omitName = false) =>
            new ComponentSideDisplayEvidence(side, identity.Record.SolutionComponentId, identity.Record.ObjectId, identity.Record.ComponentType,
                new ComponentDisplayValue(DisplayAvailability.NotRetrieved), omitName ? new ComponentDisplayValue(DisplayAvailability.NotRetrieved) :
                new ComponentDisplayValue(DisplayAvailability.Available, name, provenance ?? (identity.Record.ComponentType == 29 ? "workflow.name" : "Completed retained plugin name")),
                new ComponentDisplayValue(DisplayAvailability.NotRetrieved), correlation, new ComponentDisplayValue(DisplayAvailability.NotRetrieved));
        private static ExplorerFixture OptionAFixture(ComponentIdentity[] source, ComponentIdentity[] target = null)
        {
            var f = ExplorerFixture.Make(source, target); f.Completed = new CompletedMembershipComparison(f.Presentation,
                f.Presentation.Rows.Select((r, i) => new ComponentDisplayEvidence(i,
                    r.SourceIdentity == null ? ExplorerFixture.Display(null, DisplayEnvironmentSide.Source) : OptionADisplay(r.SourceIdentity, DisplayEnvironmentSide.Source, "Same retained name"),
                    r.TargetIdentity == null ? ExplorerFixture.Display(null, DisplayEnvironmentSide.Target) : OptionADisplay(r.TargetIdentity, DisplayEnvironmentSide.Target, "Same retained name"))), DisplayProjectionOutcome.Completed);
            return f;
        }
        private static void OptionAPreserved(ExplorerFixture f, SolutionObjectsExplorerProjection projection,
            MembershipResultRow[] rows, string csv)
        {
            Assert.AreEqual(rows.Length, projection.Rows.Count);
            Assert.AreEqual(f.Presentation.Source.Snapshot.Components.Count + f.Presentation.Target.Snapshot.Components.Count, projection.Occurrences.Count);
            for (int i = 0; i < rows.Length; i++) Assert.AreSame(rows[i], projection.Rows[i].Original.AuthoritativeRow);
            Assert.AreEqual(csv, ExplorerFixture.Csv(f));
        }

        private static ComponentIdentity Team(Guid id) => new ComponentIdentity(new SolutionComponentRecord(Guid.NewGuid(), 511, id),
            IdentityResolutionStatus.Unsupported, componentTypeKey: ComponentSemanticKinds.TeamTemplate, semanticKind: ComponentSemanticKinds.TeamTemplate);
        private static ExplorerFixture Fixture(ComponentIdentity[] source, ComponentIdentity[] target = null, string parent = null, string displayName = null)
        {
            var f = ExplorerFixture.Make(source, target); f.Completed = new CompletedMembershipComparison(f.Presentation,
                f.Presentation.Rows.Select((r, i) => new ComponentDisplayEvidence(i,
                    Display(r.SourceIdentity, DisplayEnvironmentSide.Source, displayName ?? Name(r.SourceIdentity), parent),
                    Display(r.TargetIdentity, DisplayEnvironmentSide.Target, displayName ?? Name(r.TargetIdentity), parent))), DisplayProjectionOutcome.Completed);
            return f;
        }
        private static string Name(ComponentIdentity i) => i == null ? null : i.Record.ComponentType == 511 ? "Offline team " + i.Record.ObjectId : i.ComparisonKey;
        private static ComponentSideDisplayEvidence Display(ComponentIdentity i, DisplayEnvironmentSide side, string name, string parent,
            DisplayBackingCorrelationStatus correlation = DisplayBackingCorrelationStatus.Unique) => i == null ? ExplorerFixture.Display(null, side) :
            new ComponentSideDisplayEvidence(side, i.Record.SolutionComponentId, i.Record.ObjectId, i.Record.ComponentType,
                new ComponentDisplayValue(DisplayAvailability.NotRetrieved), name == null ? new ComponentDisplayValue(DisplayAvailability.NotRetrieved) :
                    new ComponentDisplayValue(DisplayAvailability.Available, name, "Synthetic typed retained name"),
                i.Status == IdentityResolutionStatus.Resolved ? new ComponentDisplayValue(DisplayAvailability.Available, i.ComparisonKey, "Synthetic typed logical name") :
                    new ComponentDisplayValue(DisplayAvailability.NotRetrieved), correlation,
                i.Record.ComponentType == 59 && parent != null ? new ComponentDisplayValue(DisplayAvailability.Available, parent,
                    "NormalMembershipDiagnostic.savedqueryvisualization.primaryentitytypecode") : new ComponentDisplayValue(DisplayAvailability.NotRetrieved));
        private static IEnumerable<Control> Descendants(Control root) => root.Controls.Cast<Control>().SelectMany(c => new[] { c }.Concat(Descendants(c)));
        private static IEnumerable<TreeNode> Nodes(TreeNodeCollection nodes) => nodes.Cast<TreeNode>().SelectMany(n => new[] { n }.Concat(Nodes(n.Nodes)));
        private static T Item<T>(Control root, string name) where T : ToolStripItem => Descendants(root).OfType<ToolStrip>().SelectMany(t => t.Items.OfType<T>()).Single(i => i.Name == name);
        private static void Sta(Action action)
        {
            Exception error = null; var thread = new Thread(() => { try { action(); } catch (Exception e) { error = e; } });
            thread.SetApartmentState(ApartmentState.STA); thread.IsBackground = true; thread.Start(); Assert.IsTrue(thread.Join(60000));
            if (error != null) ExceptionDispatchInfo.Capture(error).Throw();
        }
    }
}
