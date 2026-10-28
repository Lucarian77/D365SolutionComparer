using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Runtime.ExceptionServices;
using System.Threading;
using System.Windows.Forms;
using D365SolutionComparer.Models.ComponentDetails;
using D365SolutionComparer.Models.Membership;
using D365SolutionComparer.Models.Presentation;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace D365SolutionComparer.Tests
{
    [TestClass, TestCategory("R2ESimplified")]
    public class SimplifiedSolutionExplorerTests
    {
        [TestMethod]
        public void CompleteComparisonCoverageKeepsAuthoritativeRowsRawUnitsAndCsv()
        {
            Sta(() =>
            {
                var fixture = ExplorerFixture.Saa(); var projection = ExplorerFixture.Project(fixture);
                var originals = fixture.Presentation.Rows.ToArray(); var csv = ExplorerFixture.Csv(fixture); int formats = 0;
                using (var host = new Form { Size = new Size(1300, 800) })
                using (var panel = new SimplifiedSolutionExplorerPanel(fixture.Presentation, row => { formats++; return "technical"; }))
                {
                    host.Controls.Add(panel); host.Show(); panel.Bind(projection); Application.DoEvents();
                    Assert.AreEqual(792, projection.SourceOccurrenceCount); Assert.AreEqual(720, projection.TargetOccurrenceCount);
                    Assert.AreEqual(840, panel.Explorer.VisibleRows.Count);
                    CollectionAssert.AreEqual(Enumerable.Range(0, 840).ToArray(), panel.Explorer.VisibleRows.ToArray());
                    for (int i = 0; i < originals.Length; i++) Assert.AreSame(originals[i], projection.Rows[i].Original.AuthoritativeRow);
                    foreach (var node in panel.Explorer.Navigation.Nodes.Cast<TreeNode>().ToArray()) panel.Explorer.Navigation.SelectedNode = node;
                    panel.Explorer.ResetNavigation(); Assert.AreEqual(840, panel.Explorer.VisibleRows.Count);
                    Assert.AreEqual(csv, ExplorerFixture.Csv(fixture)); Assert.AreEqual(0, formats);
                    Assert.IsTrue(panel.Explorer.Navigation.Nodes.Cast<TreeNode>().Any(n => n.Text.StartsWith("Unmapped comparison records")));
                }
            });
        }

        [TestMethod]
        public void RowsWithoutOccurrenceMappingsRemainAccessible()
        {
            Sta(() =>
            {
                var fixture = ExplorerFixture.Make(new[] { ExplorerFixture.Identity(300, null, status: IdentityResolutionStatus.Unsupported) });
                var original = ExplorerFixture.Project(fixture);
                var projection = new SolutionObjectsExplorerProjection(original.Owner, original.Rows, new SolutionObjectsExplorerOccurrence[0],
                    new[] { new SolutionObjectsExplorerCategory("all", "All Components", new[] { 0 }, new int[0], new int[0]),
                        new SolutionObjectsExplorerCategory("unmapped", "Raw / unmapped evidence", new[] { 0 }, new int[0], new int[0]) }, new string[0]);
                using (var panel = new SimplifiedSolutionExplorerPanel(fixture.Presentation, row => { Assert.Fail("No disclosure requested"); return ""; }))
                {
                    panel.Bind(projection); Assert.AreEqual(1, panel.Explorer.VisibleRows.Count);
                    panel.Explorer.Navigation.SelectedNode = panel.Explorer.Navigation.Nodes[1];
                    panel.Explorer.SelectComparison(0, DisplayEnvironmentSide.Source);
                    Assert.AreSame(fixture.Presentation.Rows[0], panel.Explorer.SelectedComparison.AuthoritativeRow);
                    Assert.AreEqual("Needs Review", SolutionObjectsExplorerControl.SimplifiedStatus(projection.Rows[0]));
                    Assert.AreEqual(0, projection.Occurrences.Count);
                }
            });
        }

        [DataTestMethod]
        [DataRow("Match", "Match")]
        [DataRow("Different", "Different")]
        [DataRow("Unsupported", "Needs Review")]
        [DataRow("Unresolved", "Needs Review")]
        [DataRow("Ambiguous", "Needs Review")]
        [DataRow("Not Compared", "Needs Review")]
        public void PublicStatusDoesNotOverwriteMembershipOrDefinition(string definition, string expected)
        {
            var fixture = DefinitionFixture(definition);
            var projection = ExplorerFixture.Project(fixture); var row = projection.Rows[0]; var original = row.Original.AuthoritativeRow;
            Assert.AreEqual(expected, SolutionObjectsExplorerControl.SimplifiedStatus(row));
            Assert.AreEqual(definition, original.DefinitionStatus); Assert.AreEqual(MembershipPresence.PresentInBoth, original.Presence);
            Assert.AreSame(original, fixture.Presentation.Rows[0]);
        }

        [TestMethod]
        public void OneSidedAndUnsupportedMembershipStayDistinct()
        {
            var fixture = ExplorerFixture.Make(new[] { ExplorerFixture.Identity(61, "left"), ExplorerFixture.Identity(300, null, status: IdentityResolutionStatus.Unsupported) },
                new[] { ExplorerFixture.Identity(61, "right") });
            var projection = ExplorerFixture.Project(fixture);
            CollectionAssert.AreEquivalent(new[] { "Source Only", "Target Only", "Needs Review" }, projection.Rows.Select(SolutionObjectsExplorerControl.SimplifiedStatus).ToArray());
            Assert.AreEqual(1, projection.Rows.Count(r => r.Original.AuthoritativeRow.Presence == MembershipPresence.Indeterminate));
        }

        [DataTestMethod, DataRow("Match"), DataRow("Different"), DataRow("Unsupported")]
        public void SelectingRowImmediatelyShowsIdentityPresenceAndExactProperties(string status)
        {
            Sta(() =>
            {
                var fixture = DefinitionFixture(status); int formats = 0;
                using (var host = new Form { Size = new Size(1300, 850) })
                using (var panel = new SimplifiedSolutionExplorerPanel(fixture.Presentation, row => { formats++; return "SECRET"; }))
                {
                    host.Controls.Add(panel); host.Show(); panel.Bind(ExplorerFixture.Project(fixture)); panel.Explorer.SelectComparison(0); Application.DoEvents();
                    var identity = Descendants(panel).OfType<TextBox>().Single(c => c.Name == "SolutionExplorerSelectedIdentity");
                    Assert.IsTrue(identity.Visible && identity.ReadOnly);
                    StringAssert.Contains(identity.Text, "resource"); StringAssert.Contains(identity.Text, "Source: Present"); StringAssert.Contains(identity.Text, "Target: Present");
                    StringAssert.Contains(identity.Text, "Definition: " + (status == "Match" ? "Match (compared properties)" : status));
                    var properties = Descendants(panel).OfType<DataGridView>().Single(c => c.Name == "ExplorerComparedProperties");
                    Assert.IsTrue(properties.Visible && properties.ReadOnly && properties.VirtualMode);
                    Assert.AreSame(fixture.Presentation.Rows[0].DefinitionDetail, properties.Tag);
                    Assert.AreEqual(2, properties.RowCount);
                    Assert.AreEqual("value", properties.Rows[0].Cells[0].Value);
                    Assert.AreEqual("source\r\nfull", properties.Rows[0].Cells[1].Value);
                    Assert.AreEqual("[No value supplied; availability not established]", properties.Rows[1].Cells[1].Value);
                    Assert.AreEqual(0, formats, "Normal selection must not invoke full technical formatting.");
                }
            });
        }

        [TestMethod]
        public void SearchStatusDifferencesAndResetRetainOwnerAndScope()
        {
            Sta(() =>
            {
                var fixture = WithRetainedNames(ExplorerFixture.Make(new[] { ExplorerFixture.Identity(61, "source-one"), ExplorerFixture.Identity(61, "paired") },
                    new[] { ExplorerFixture.Identity(61, "paired"), ExplorerFixture.Identity(61, "target-one") }));
                using (var panel = new SimplifiedSolutionExplorerPanel(fixture.Presentation, row => ""))
                {
                    var projection = ExplorerFixture.Project(fixture); panel.Bind(projection);
                    var scope = Item<ToolStripComboBox>(panel, "SolutionExplorerScope"); var search = Item<ToolStripTextBox>(panel, "SolutionExplorerSearch");
                    var status = Item<ToolStripComboBox>(panel, "SolutionExplorerStatus");
                    scope.SelectedIndex = 1; Assert.IsTrue(panel.Explorer.VisibleRows.All(i => fixture.Presentation.Rows[i].SourceIdentity != null));
                    search.Text = "source-one"; Assert.AreEqual(1, panel.Explorer.VisibleRows.Count);
                    status.SelectedIndex = 4; Assert.AreEqual(0, panel.Explorer.VisibleRows.Count); Assert.IsNull(panel.Explorer.SelectedComparison);
                    Item<ToolStripButton>(panel, "SolutionExplorerReset").PerformClick();
                    Assert.AreEqual(1, scope.SelectedIndex); Assert.AreEqual("", search.Text); Assert.AreEqual(0, status.SelectedIndex);
                    Assert.AreEqual(2, panel.Explorer.VisibleRows.Count); Assert.IsNull(panel.Explorer.SelectedComparison);
                    Assert.AreSame(projection, panel.Projection);
                    Item<ToolStripButton>(panel, "SolutionExplorerDifferences").PerformClick();
                    Assert.IsTrue(panel.Explorer.VisibleRows.All(i => SolutionObjectsExplorerControl.SimplifiedStatus(projection.Rows[i]) != "Match"));
                }
            });
        }

        [TestMethod]
        public void CategoryAndClearedSelectionNeverDisplayPreviousProperties()
        {
            Sta(() =>
            {
                var fixture = DefinitionFixture("Different");
                using (var panel = new SimplifiedSolutionExplorerPanel(fixture.Presentation, row => ""))
                {
                    panel.Bind(ExplorerFixture.Project(fixture)); panel.Explorer.SelectComparison(0);
                    var properties = Descendants(panel).OfType<DataGridView>().Single(c => c.Name == "ExplorerComparedProperties");
                    Assert.AreEqual(2, properties.RowCount);
                    panel.Explorer.ResetNavigation(); Assert.IsNull(panel.Explorer.SelectedComparison); Assert.AreEqual(0, properties.RowCount);
                    panel.Explorer.SelectComparison(0); Item<ToolStripTextBox>(panel, "SolutionExplorerSearch").Text = "no matching retained component";
                    Assert.AreEqual(0, panel.Explorer.VisibleRows.Count); Assert.IsNull(panel.Explorer.SelectedComparison); Assert.AreEqual(0, properties.RowCount);
                    StringAssert.Contains(Descendants(panel).OfType<TextBox>().Single(c => c.Name == "SolutionExplorerSelectedIdentity").Text, "Select a component");
                }
            });
        }

        [TestMethod]
        public void ExpandedEntriesKeepSideAndOnlyRetainedRelationships()
        {
            Sta(() =>
            {
                var fixture = ExplorerFixture.Apps();
                using (var host = new Form { Size = new Size(1300, 850) })
                using (var panel = new SimplifiedSolutionExplorerPanel(fixture.Presentation, row => ""))
                {
                    host.Controls.Add(panel); host.Show(); panel.Bind(ExplorerFixture.Project(fixture)); Application.DoEvents();
                    var apps = panel.Explorer.Navigation.Nodes.Cast<TreeNode>().Single(n => n.Text.StartsWith("Apps (")); apps.Expand();
                    Assert.AreEqual(2, apps.Nodes.Count);
                    foreach (TreeNode app in apps.Nodes)
                    {
                        StringAssert.Contains(app.Text, "raw occurrence"); app.Expand();
                        Assert.IsTrue(app.Nodes.Cast<TreeNode>().Any(n => n.Text.StartsWith("Settings — structural")));
                        panel.Explorer.Navigation.SelectedNode = app;
                        Assert.IsNotNull(panel.Explorer.SelectedComparison);
                        StringAssert.StartsWith(app.Text, panel.Explorer.SelectedSide + ":");
                    }
                }
            });
        }

        [DataTestMethod, DataRow(1000, 620, 9F), DataRow(1300, 850, 12F), DataRow(1500, 1000, 16F)]
        public void IntegratedLayoutHasSixColumnsReachableFiltersAndExpandableProperties(int width, int height, float font)
        {
            Sta(() =>
            {
                var fixture = DefinitionFixture("Different");
                using (var host = new Form { Size = new Size(width, height), Font = new Font("Segoe UI", font) })
                using (var panel = new SimplifiedSolutionExplorerPanel(fixture.Presentation, row => ""))
                {
                    host.Controls.Add(panel); host.Show(); panel.Bind(ExplorerFixture.Project(fixture)); panel.Explorer.SelectComparison(0); Application.DoEvents();
                    CollectionAssert.AreEqual(new[] { "Display Name", "Name", "Component Type", "Source", "Target", "Difference Status" },
                        panel.Explorer.Grid.Columns.Cast<DataGridViewColumn>().Where(c => c.Visible).Select(c => c.HeaderText).ToArray());
                    foreach (var item in Descendants(panel).OfType<ToolStrip>().SelectMany(b => b.Items.Cast<ToolStripItem>())) Assert.IsTrue(item.Available);
                    Assert.IsTrue(panel.Explorer.Grid.ReadOnly && panel.Explorer.Grid.VirtualMode);
                    foreach (var split in Descendants(panel).OfType<SplitContainer>()) Assert.IsFalse(split.IsSplitterFixed);
                    var expand = Descendants(panel).OfType<LinkLabel>().Single(c => c.Name == "ExplorerExpandDetails"); Assert.IsTrue(expand.Visible);
                    var properties = Descendants(panel).OfType<DataGridView>().Single(c => c.Name == "ExplorerComparedProperties");
                    Assert.IsTrue(properties.ClientSize.Height > properties.ColumnHeadersHeight + properties.RowTemplate.Height, "At least one complete property row is visible.");
                    Assert.AreEqual("Different", properties.Rows[0].Cells[3].Value);
                }
            });
        }

        [TestMethod]
        public void DifferentComparisonOwnerIsRejectedBeforePublishingRows()
        {
            Sta(() =>
            {
                var fixture = ExplorerFixture.Apps();
                using (var panel = new SimplifiedSolutionExplorerPanel(fixture.Presentation, row => ""))
                {
                    Assert.ThrowsException<InvalidOperationException>(() => panel.Bind(ExplorerFixture.Project(ExplorerFixture.Apps())));
                    Assert.IsNull(panel.Projection); Assert.IsNull(panel.Explorer.Projection);
                }
            });
        }

        [TestMethod]
        public void NormalResultEntryRemovesDiagnosticNavigationAndDoesNoRequests()
        {
            Sta(() =>
            {
                var fixture = ExplorerFixture.Apps(); var spy = new FakeOrganizationService(); int captures = 0;
                using (var form = new MembershipResultsForm(fixture.Completed, captureAppSettingEvidence: () => captures++, sourceService: spy, targetService: spy))
                {
                    form.UseSimplifiedExplorer(); form.Show();
                    Wait(() => form.ExplorerProjection != null);
                    var panel = form.Controls.OfType<SimplifiedSolutionExplorerPanel>().Single();
                    Assert.IsTrue(panel.Visible); Assert.IsTrue(form.Controls.Cast<Control>().Where(c => !ReferenceEquals(c, panel)).All(c => !c.Visible));
                    Assert.IsFalse(Descendants(form).OfType<ToolStrip>().Where(c => c.Visible).SelectMany(c => c.Items.Cast<ToolStripItem>()).Any(i => i.Text.Contains("Advanced") || i.Text.Contains("Coverage") || i.Text.Contains("Technical Membership")));
                    panel.Explorer.SelectComparison(0); panel.FocusSearch(); Assert.IsTrue(Item<ToolStripTextBox>(panel, "SolutionExplorerSearch").TextBox.Focused);
                    Assert.AreEqual(0, captures + spy.Calls + spy.ExecuteCalls + spy.WriteCalls);
                    Assert.AreSame(fixture.Presentation, panel.Projection.Owner);
                }
            });
        }

        [TestMethod]
        public void OriginalRowAndScopeHandoffPreserveExactSourceOccurrence()
        {
            Sta(() =>
            {
                var fixture = ExplorerFixture.Make(new[] { ExplorerFixture.Identity(61, "source-only") }, new[] { ExplorerFixture.Identity(61, "target-only") });
                using (var host = new Form { Size = new Size(1300, 850) })
                using (var panel = new SimplifiedSolutionExplorerPanel(fixture.Presentation, row => { Assert.Fail("No technical disclosure"); return ""; }))
                {
                    host.Controls.Add(panel); host.Show(); panel.Bind(ExplorerFixture.Project(fixture));
                    var scope = Item<ToolStripComboBox>(panel, "SolutionExplorerScope"); scope.SelectedIndex = 1;
                    int occurrence = panel.Projection.Occurrences.First(o => o.Original.Side == DisplayEnvironmentSide.Source && o.Original.Record.ComponentType == 61).Original.OccurrenceIndex;
                    panel.Explorer.SelectOccurrence(occurrence); var original = panel.Explorer.SelectedComparison;
                    Assert.IsNotNull(original); Assert.AreEqual(DisplayEnvironmentSide.Source, panel.Explorer.SelectedSide);
                    var locate = Descendants(panel).OfType<LinkLabel>().Single(l => l.Text == "Locate original comparison row");
                    typeof(LinkLabel).GetMethod("OnLinkClicked", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic)
                        .Invoke(locate, new object[] { new LinkLabelLinkClickedEventArgs(locate.Links[0]) });
                    Assert.AreSame(original, panel.Explorer.SelectedComparison); Assert.AreEqual(occurrence, panel.Explorer.SelectedOccurrenceIndex);
                    scope.SelectedIndex = 2;
                    Assert.AreSame(original, panel.Explorer.SelectedComparison); Assert.AreEqual(DisplayEnvironmentSide.Source, panel.Explorer.SelectedSide);
                    Assert.AreEqual(occurrence, panel.Explorer.SelectedOccurrenceIndex, "An inactive Source occurrence is retained without activating a Target route.");
                    scope.SelectedIndex = 1;
                    Assert.AreSame(original, panel.Explorer.SelectedComparison); Assert.AreEqual(DisplayEnvironmentSide.Source, panel.Explorer.SelectedSide);
                    Assert.AreEqual(occurrence, panel.Explorer.SelectedOccurrenceIndex);
                    Assert.AreSame(fixture.Presentation.Rows[original.OriginalRowIndex], original.AuthoritativeRow);
                }
            });
        }

        [TestMethod]
        public void NewerCategoryIntentWinsOverSavedScopeRow()
        {
            Sta(() =>
            {
                var fixture = ExplorerFixture.Apps();
                using (var host = new Form { Size = new Size(1300, 850) })
                using (var panel = new SimplifiedSolutionExplorerPanel(fixture.Presentation, row => ""))
                {
                    host.Controls.Add(panel); host.Show(); panel.Bind(ExplorerFixture.Project(fixture));
                    var scope = Item<ToolStripComboBox>(panel, "SolutionExplorerScope"); scope.SelectedIndex = 1;
                    panel.Explorer.SelectComparison(0, DisplayEnvironmentSide.Source);
                    // Synthetic explicit row intent uses the same callback as user grid/key navigation.
                    panel.Explorer.NavigationIntentChanged(); scope.SelectedIndex = 2;
                    panel.Explorer.Navigation.SelectedNode = panel.Explorer.Navigation.Nodes.Cast<TreeNode>().Single(n => n.Text.StartsWith("Apps ("));
                    Assert.IsNull(panel.Explorer.SelectedComparison); scope.SelectedIndex = 1;
                    Assert.IsNull(panel.Explorer.SelectedComparison); Assert.IsNull(panel.Explorer.SelectedOccurrenceIndex);
                    StringAssert.StartsWith(panel.Explorer.Navigation.SelectedNode.Text, "Apps (");
                    Assert.AreEqual(1, panel.Explorer.VisibleRows.Count);
                }
            });
        }

        [TestMethod]
        public void ClosedNormalWindowCannotPublishDelayedProjection()
        {
            Sta(() =>
            {
                var fixture = ExplorerFixture.Apps();
                using (var release = new ManualResetEventSlim())
                using (var started = new ManualResetEventSlim())
                using (var finished = new ManualResetEventSlim())
                using (var form = new MembershipResultsForm(fixture.Completed))
                {
                    var project = form.ProjectHierarchy;
                    form.ProjectHierarchy = token => { started.Set(); release.Wait(5000); try { return project(token); } finally { finished.Set(); } };
                    form.UseSimplifiedExplorer(); form.Show(); Wait(() => started.IsSet); form.Close(); release.Set(); Wait(() => finished.IsSet);
                    Assert.IsTrue(form.IsDisposed); Assert.IsNull(form.ExplorerProjection);
                }
            });
        }

        private static ExplorerFixture DefinitionFixture(string status)
        {
            var fixture = WithRetainedNames(ExplorerFixture.Make(new[] { ExplorerFixture.Identity(61, "resource") }, new[] { ExplorerFixture.Identity(61, "resource") }));
            var old = fixture.Completed; var rows = fixture.Presentation.Rows.ToArray(); var row = rows[0];
            var properties = new[] { new ComponentPropertyPresentation("value", "source\r\nfull", status == "Different" ? "target" : "source\r\nfull", status == "Different", status),
                new ComponentPropertyPresentation("unavailable", null, null, false, "Unsupported") };
            rows[0] = row.WithDefinition(status, status == "Different" ? "value" : "", new ComponentDefinitionDetailPresentation("Web resource", "resource", row.MembershipStatus, status, properties, "Retained maintenance diagnostic"));
            fixture.Presentation = new MembershipComparisonPresentation("SAA", fixture.Presentation.Source, fixture.Presentation.Target, rows);
            fixture.Completed = new CompletedMembershipComparison(fixture.Presentation, old.DisplayEvidence, DisplayProjectionOutcome.Completed); return fixture;
        }
        private static ExplorerFixture WithRetainedNames(ExplorerFixture fixture)
        {
            fixture.Completed = new CompletedMembershipComparison(fixture.Presentation, fixture.Presentation.Rows.Select((row, index) =>
                new ComponentDisplayEvidence(index, ExplorerFixture.Display(row.SourceIdentity, DisplayEnvironmentSide.Source,
                    friendly: new ComponentDisplayValue(DisplayAvailability.Available, row.SourceIdentity?.ComparisonKey, "Synthetic retained display name"),
                    name: new ComponentDisplayValue(DisplayAvailability.Available, row.SourceIdentity?.ComparisonKey, "Synthetic retained logical name")),
                    ExplorerFixture.Display(row.TargetIdentity, DisplayEnvironmentSide.Target,
                    friendly: new ComponentDisplayValue(DisplayAvailability.Available, row.TargetIdentity?.ComparisonKey, "Synthetic retained display name"),
                    name: new ComponentDisplayValue(DisplayAvailability.Available, row.TargetIdentity?.ComparisonKey, "Synthetic retained logical name")))), DisplayProjectionOutcome.Completed);
            return fixture;
        }
        private static IEnumerable<Control> Descendants(Control root) => root.Controls.Cast<Control>().SelectMany(c => new[] { c }.Concat(Descendants(c)));
        private static T Item<T>(Control root, string name) where T : ToolStripItem => Descendants(root).OfType<ToolStrip>().SelectMany(c => c.Items.Cast<ToolStripItem>()).OfType<T>().Single(c => c.Name == name);
        private static void Wait(Func<bool> ready)
        {
            var watch = Stopwatch.StartNew(); while (!ready() && watch.ElapsedMilliseconds < 10000) { Application.DoEvents(); Thread.Sleep(5); }
            Assert.IsTrue(ready(), "Local projection did not publish.");
        }
        private static void Sta(Action action)
        {
            Exception error = null; var thread = new Thread(() => { try { action(); } catch (Exception e) { error = e; } });
            thread.IsBackground = true; thread.SetApartmentState(ApartmentState.STA); thread.Start(); Assert.IsTrue(thread.Join(60000));
            if (error != null) ExceptionDispatchInfo.Capture(error).Throw();
        }
    }
}
