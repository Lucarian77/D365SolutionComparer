using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Runtime.ExceptionServices;
using System.Threading;
using System.Windows.Forms;
using D365SolutionComparer.Models.Membership;
using D365SolutionComparer.Models.Presentation;
using D365SolutionComparer.Services.Membership;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace D365SolutionComparer.Tests
{
    [TestClass, TestCategory("H2DExplorerUi")]
    public class SolutionObjectsExplorerUiTests
    {
        public TestContext TestContext { get; set; }
        [TestMethod]
        public void FlatRemainsDefaultAndExistingViewsAndDifferencesPredicateAreUnchanged()
        {
            Sta(() =>
            {
                var f = ExplorerFixture.Edu(); var spy = new FakeOrganizationService(); var csv = ExplorerFixture.Csv(f);
                using (var form = new MembershipResultsForm(f.Completed, sourceService: spy, targetService: spy))
                {
                    form.Show(); Assert.AreEqual(0, View(form).SelectedIndex);
                    CollectionAssert.AreEqual(new[] { "Flat View", "Hierarchy View", "Differences Only" }, View(form).Items.Cast<string>().ToArray());
                    Assert.AreEqual(840, Grid(form).Rows.Count); Assert.IsNull(form.ExplorerProjection);
                    Open(form); Assert.AreEqual(840, form.ExplorerProjection.Rows.Count); Assert.AreEqual(1512, form.ExplorerProjection.Occurrences.Count);
                    View(form).SelectedIndex = 2; Application.DoEvents();
                    CollectionAssert.AreEqual(f.Presentation.Rows.Select((r, i) => new { r, i }).Where(x => x.r.Presence != MembershipPresence.PresentInBoth || x.r.DefinitionStatus == "Different").Select(x => x.i).ToArray(),
                        form.FilteredRows().Select(r => r.OriginalRowIndex).ToArray());
                    Open(form); var explorer = Control(form);
                    explorer.Navigation.SelectedNode = explorer.Navigation.Nodes[1]; // All Components, rather than the object-candidate subset.
                    CollectionAssert.AreEqual(form.FilteredRows().Select(r => r.OriginalRowIndex).ToArray(), explorer.VisibleRows.ToArray());
                    explorer.SelectComparison(form.FilteredRows()[0].OriginalRowIndex);
                    var selected = explorer.SelectedComparison;
                    int excludedOccurrence = form.ExplorerProjection.Occurrences.First(o => o.Original.ComparisonRowIndices.Contains(0)).Original.OccurrenceIndex;
                    explorer.SelectOccurrence(excludedOccurrence);
                    Assert.AreSame(selected, explorer.SelectedComparison, "A filtered-out occurrence without a verified parent-context role must not enter Differences Only navigation.");
                    View(form).SelectedIndex = 0; Assert.AreEqual(840, Grid(form).Rows.Count);
                    Assert.AreEqual(csv, ExplorerFixture.Csv(f)); Assert.AreEqual(0, spy.Calls + spy.ExecuteCalls + spy.WriteCalls);
                }
            });
        }
        [TestMethod]
        public void ExplorerHasSixDefaultColumnsVirtualGridAndResizablePanels()
        {
            Sta(() =>
            {
                using (var form = new MembershipResultsForm(ExplorerFixture.Apps().Completed))
                {
                    form.Show(); Open(form); var control = Control(form);
                    CollectionAssert.AreEqual(new[] { "Display Name", "Name", "Type", "Source", "Target", "Comparison Status" },
                        control.Grid.Columns.Cast<DataGridViewColumn>().Where(c => c.Visible).Select(c => c.HeaderText).ToArray());
                    Assert.IsTrue(control.Grid.VirtualMode); Assert.IsTrue(control.Grid.ReadOnly);
                    Assert.IsFalse(string.IsNullOrWhiteSpace(control.Navigation.AccessibleName));
                    Assert.IsTrue(control.Navigation.TabStop && control.Grid.TabStop);
                    control.SelectComparison(3);
                    control.Grid.CurrentCell = control.Grid.Rows[4].Cells[0];
                    typeof(System.Windows.Forms.Control).GetMethod("OnKeyUp", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic)
                        .Invoke(control.Grid, new object[] { new KeyEventArgs(Keys.Down) });
                    Assert.AreEqual(4, control.SelectedComparison.OriginalRowIndex, "Keyboard selection must update the original-row evidence selection.");
                    foreach (var split in Controls(control).OfType<SplitContainer>()) { Assert.IsTrue(split.Panel1MinSize > 0 && split.Panel2MinSize > 0); Assert.IsFalse(split.IsSplitterFixed); }
                    Controls(form).OfType<CheckBox>().Single(c => c.Name == "OptionalMetadataToggle").Checked = true;
                    Assert.AreEqual(11, control.Grid.Columns.Cast<DataGridViewColumn>().Count(c => c.Visible));
                }
            });
        }
        [TestMethod]
        public void AppSettingDrillDownKeepsEverySideAndStructuralProvenance()
        {
            Sta(() =>
            {
                var f = ExplorerFixture.Apps(); using (var form = new MembershipResultsForm(f.Completed))
                {
                    form.Show(); Open(form); var c = Control(form);
                    foreach (var side in new[] { DisplayEnvironmentSide.Source, DisplayEnvironmentSide.Target })
                    {
                        int parent = form.ExplorerProjection.Occurrences.Single(o => o.Original.Side == side && o.Original.Record.ComponentType == 80).Original.OccurrenceIndex;
                        c.SelectOccurrence(parent); var node = c.Navigation.SelectedNode; node.Expand();
                        var group = node.Nodes.Cast<TreeNode>().Single(n => n.Text.StartsWith("AppSettings — structural")); group.Expand();
                        StringAssert.Contains(group.ToolTipText, "Verified structural parent");
                        Assert.AreEqual(5, group.Nodes.Count); c.Navigation.SelectedNode = group.Nodes[0];
                        StringAssert.Contains(c.EvidenceText, "StructuralParent: Verified"); StringAssert.Contains(c.EvidenceText, side + ":");
                        Assert.AreEqual(side, c.SelectedSide); Assert.IsNotNull(c.SelectedComparison);
                    }
                    Assert.AreEqual(12, form.ExplorerProjection.Occurrences.Count);
                }
            });
        }
        [TestMethod]
        public void TableDrillDownLabelsRootInclusionWithoutInventingOwnership()
        {
            Sta(() =>
            {
                var table = ExplorerFixture.Identity(1, "account"); var f = ExplorerFixture.Make(new[] { table, ExplorerFixture.Identity(26, "view", table.Record.SolutionComponentId) });
                using (var form = new MembershipResultsForm(f.Completed))
                {
                    form.Show(); Open(form); var c = Control(form); c.SelectOccurrence(form.ExplorerProjection.Occurrences.Single(o => o.Original.Record.ComponentType == 1).Original.OccurrenceIndex);
                    var nodes = c.Navigation.SelectedNode.Nodes.Cast<TreeNode>().ToArray();
                    Assert.IsTrue(nodes.Any(n => n.Text.StartsWith("Views — inclusion root")));
                    Assert.IsTrue(nodes.Any(n => n.Text.StartsWith("Columns — no retained verified links")));
                    var views = nodes.Single(n => n.Text.StartsWith("Views — inclusion root"));
                    StringAssert.Contains(views.ToolTipText, "not ownership"); views.Expand(); c.Navigation.SelectedNode = views.Nodes[0];
                    StringAssert.Contains(c.EvidenceText, "InclusionRoot: Verified"); StringAssert.Contains(c.EvidenceText, "rootsolutioncomponentid");
                }
            });
        }
        [TestMethod]
        public void MultipleMappingsRequireExplicitRowSelection()
        {
            Sta(() =>
            {
                var repeated = ExplorerFixture.Identity(300, null, status: IdentityResolutionStatus.Unsupported);
                using (var form = new MembershipResultsForm(ExplorerFixture.Make(new[] { repeated, repeated }).Completed))
                {
                    form.Show(); Open(form); var c = Control(form); c.SelectOccurrence(0);
                    Assert.AreEqual(2, c.VisibleRows.Count); Assert.IsNull(c.SelectedComparison);
                    StringAssert.Contains(c.EvidenceText, "no arbitrary first mapping");
                    c.SelectComparison(1, DisplayEnvironmentSide.Source); Assert.AreEqual(1, c.SelectedComparison.OriginalRowIndex);
                }
            });
        }
        [TestMethod]
        public void ExplicitOriginalRowAndSideSurviveExplorerHierarchyAndFlat()
        {
            Sta(() =>
            {
                using (var form = new MembershipResultsForm(ExplorerFixture.Apps().Completed))
                {
                    form.Show(); Open(form); var c = Control(form); c.SelectComparison(3, DisplayEnvironmentSide.Target);
                    View(form).SelectedIndex = 1; Pump(() => form.Controls.OfType<MembershipHierarchyControl>().Any(h => h.Visible));
                    var hierarchy = form.Controls.OfType<MembershipHierarchyControl>().Single();
                    Assert.AreEqual(3, hierarchy.SelectedComparison.OriginalRowIndex); Assert.AreEqual(DisplayEnvironmentSide.Target, hierarchy.SelectedSide);
                    View(form).SelectedIndex = 0; Assert.AreEqual(3, ((MembershipGridRow)Grid(form).CurrentRow.DataBoundItem).OriginalRowIndex);
                }
            });
        }
        [TestMethod]
        public void ProjectionAndNativeNavigationAreReusedOnSimpleViewSwitch()
        {
            Sta(() =>
            {
                using (var form = new MembershipResultsForm(ExplorerFixture.Apps().Completed))
                {
                    int count = 0; var original = form.ProjectExplorer;
                    form.ProjectExplorer = (h, t) => { count++; return original(h, t); };
                    form.Show(); Open(form); var projection = form.ExplorerProjection; var node = Control(form).Navigation.Nodes[0];
                    Button(form).PerformClick(); Assert.IsTrue(Grid(form).Visible); Open(form);
                    Assert.AreEqual(1, count); Assert.AreSame(projection, form.ExplorerProjection); Assert.AreSame(node, Control(form).Navigation.Nodes[0]);
                }
            });
        }
        [TestMethod]
        public void FailureRestoresFlatWithoutLeakingMessagesOrChangingResults()
        {
            Sta(() =>
            {
                var f = ExplorerFixture.Apps(); using (var form = new MembershipResultsForm(f.Completed))
                {
                    form.ProjectExplorer = (h, t) => throw new InvalidOperationException("PRIVATE-SECRET"); form.Show(); Button(form).PerformClick();
                    Pump(() => form.HierarchyDiagnostics.Any());
                    Assert.IsTrue(Grid(form).Visible); Assert.AreEqual(0, View(form).SelectedIndex); Assert.AreEqual(6, Grid(form).Rows.Count);
                    Assert.IsFalse(form.HierarchyDiagnostics.Any(d => d.Contains("PRIVATE-SECRET"))); Assert.AreEqual(6, f.Presentation.Rows.Count);
                }
            });
        }
        [TestMethod]
        public void StaleExplorerUpdateCannotOverrideAnExplicitFlatSelection()
        {
            Sta(() =>
            {
                using (var form = new MembershipResultsForm(ExplorerFixture.Apps().Completed))
                {
                    var release = new ManualResetEventSlim(); var started = new ManualResetEventSlim(); var original = form.ProjectExplorer;
                    form.ProjectExplorer = (h, t) => { started.Set(); release.Wait(t); return original(h, t); };
                    form.Show(); Button(form).PerformClick(); Pump(() => started.IsSet); View(form).SelectedIndex = 2; View(form).SelectedIndex = 0;
                    release.Set(); Pump(() => form.ExplorerProjection != null || !Button(form).Checked);
                    for (int i = 0; i < 30; i++) { Application.DoEvents(); Thread.Sleep(1); }
                    Assert.IsTrue(Grid(form).Visible); Assert.IsFalse(ControlOrNull(form)?.Visible ?? false);
                    release.Dispose(); started.Dispose();
                }
            });
        }
        [TestMethod]
        public void ClosingCancelsLocalProjectionWithoutPublishingPartialUi()
        {
            Sta(() =>
            {
                using (var started = new ManualResetEventSlim()) using (var form = new MembershipResultsForm(ExplorerFixture.Apps().Completed))
                {
                    form.ProjectExplorer = (h, t) => { started.Set(); t.WaitHandle.WaitOne(10000); t.ThrowIfCancellationRequested(); throw new InvalidOperationException(); };
                    form.Show(); Button(form).PerformClick(); Pump(() => started.IsSet); form.Close();
                    Assert.IsTrue(form.IsDisposed); Assert.IsNull(form.ExplorerProjection);
                }
            });
        }
        [DataTestMethod, DataRow(840), DataRow(10000)]
        public void VirtualGridAndFirstNativePageStayBoundedWithoutServiceCalls(int count)
        {
            Sta(() =>
            {
                var f = ExplorerFixture.Make(Enumerable.Range(0, count).Select(i => ExplorerFixture.Identity(61, "resource" + i)));
                var spy = new FakeOrganizationService(); using (var form = new MembershipResultsForm(f.Completed, sourceService: spy, targetService: spy))
                {
                    form.Show(); var timer = Stopwatch.StartNew(); Open(form); timer.Stop(); var c = Control(form); int initial = Walk(c.Navigation.Nodes).Count();
                    var side = c.Navigation.Nodes.Cast<TreeNode>().Single(n => n.Text.StartsWith("Source:"));
                    var category = side.Nodes.Cast<TreeNode>().Single(n => n.Text.StartsWith("Web Resources"));
                    var expand = Stopwatch.StartNew(); category.Expand(); category.Nodes[0].Expand(); expand.Stop();
                    int nodes = Walk(c.Navigation.Nodes).Count(); Assert.IsTrue(initial < 100); Assert.IsTrue(nodes < 600);
                    Assert.AreEqual(200, category.Nodes[0].Nodes.Count); Assert.AreEqual(count, c.Grid.RowCount);
                    Assert.AreEqual(0, spy.Calls + spy.ExecuteCalls + spy.WriteCalls); Assert.IsTrue(timer.Elapsed.TotalSeconds < 15);
                    c.SelectComparison(count - 1); var accessible = c.Grid.Rows[count - 1].AccessibilityObject;
                    Assert.IsFalse(string.IsNullOrWhiteSpace(accessible.Description)); Assert.AreEqual(AccessibleRole.Row, accessible.Role);
                    Assert.IsTrue(expand.Elapsed.TotalSeconds < 10);
                    TestContext.WriteLine("H2D native synthetic rows={0}; open_ms={1:F3}; expand_ms={2:F3}; initial_nodes={3}; expanded_nodes={4}; requests=0; writes=0", count, timer.Elapsed.TotalMilliseconds, expand.Elapsed.TotalMilliseconds, initial, nodes);
                }
            });
        }
        [TestMethod]
        public void NativeSyntheticScreenshotsRetainAppAndTableEvidence()
        {
            var path = Environment.GetEnvironmentVariable("H2D_EVIDENCE_DIRECTORY"); if (string.IsNullOrWhiteSpace(path)) return;
            Sta(() =>
            {
                Directory.CreateDirectory(path); var apps = SyntheticNames(ExplorerFixture.Apps()); using (var form = new MembershipResultsForm(apps.Completed))
                {
                    form.Size = new Size(1600, 900); form.Show(); Open(form); Save(form, path, "01-explorer.png");
                    var c = Control(form); int parent = form.ExplorerProjection.Occurrences.Single(o => o.Original.Side == DisplayEnvironmentSide.Source && o.Original.Record.ComponentType == 80).Original.OccurrenceIndex;
                    c.SelectOccurrence(parent); c.Navigation.SelectedNode.Expand();
                    var group = c.Navigation.SelectedNode.Nodes.Cast<TreeNode>().Single(n => n.Text.StartsWith("AppSettings — structural"));
                    group.Expand(); c.Navigation.SelectedNode = group.Nodes[0]; Save(form, path, "02-appsettings.png");
                    parent = form.ExplorerProjection.Occurrences.Single(o => o.Original.Side == DisplayEnvironmentSide.Target && o.Original.Record.ComponentType == 80).Original.OccurrenceIndex;
                    c.SelectOccurrence(parent); c.Navigation.SelectedNode.Expand();
                    group = c.Navigation.SelectedNode.Nodes.Cast<TreeNode>().Single(n => n.Text.StartsWith("AppSettings — structural"));
                    group.Expand(); c.Navigation.SelectedNode = group.Nodes[0]; Save(form, path, "03-target-app.png");
                }
                var table = ExplorerFixture.Identity(1, "account"); using (var form = new MembershipResultsForm(SyntheticNames(ExplorerFixture.Make(new[] { table, ExplorerFixture.Identity(2, "account.ava_regionlocation", table.Record.SolutionComponentId) })).Completed))
                {
                    form.Size = new Size(1600, 900); form.Show(); Open(form); var c = Control(form);
                    c.SelectOccurrence(form.ExplorerProjection.Occurrences.Single(o => o.Original.Record.ComponentType == 1).Original.OccurrenceIndex);
                    c.Navigation.SelectedNode.Expand(); Save(form, path, "04-table-links.png");
                }
            });
        }

        private static ExplorerFixture SyntheticNames(ExplorerFixture f)
        {
            Func<ComponentSideDisplayEvidence, ComponentSideDisplayEvidence> name = side =>
            {
                var row = f.Presentation.Rows.First(r => (side.EnvironmentSide == DisplayEnvironmentSide.Source ? r.SourceIdentity : r.TargetIdentity)?.Record.SolutionComponentId == side.SolutionComponentId);
                var identity = side.EnvironmentSide == DisplayEnvironmentSide.Source ? row.SourceIdentity : row.TargetIdentity;
                string app, setting; string text = MembershipGridRow.AppSettingIdentityNames(identity, out app, out setting) ? setting : identity.ComparisonKey;
                return ExplorerFixture.Display(identity, side.EnvironmentSide,
                    friendly: new ComponentDisplayValue(DisplayAvailability.Available, "Demo " + (identity.Record.ComponentType == 80 ? "Service App" : text), "Synthetic metadata only"),
                    name: new ComponentDisplayValue(DisplayAvailability.Available, text, "Synthetic metadata only"));
            };
            f.Completed = new CompletedMembershipComparison(new CompletedMembershipComparison(f.Presentation,
                f.Completed.DisplayEvidence.Select(e => new ComponentDisplayEvidence(e.RowIndex,
                    e.SourceDisplayEvidence.HasLocalObservation ? name(e.SourceDisplayEvidence) : e.SourceDisplayEvidence,
                    e.TargetDisplayEvidence.HasLocalObservation ? name(e.TargetDisplayEvidence) : e.TargetDisplayEvidence)), DisplayProjectionOutcome.Completed),
                f.Completed.SourceHierarchyParents, f.Completed.TargetHierarchyParents); return f;
        }

        private static void Open(MembershipResultsForm form) { if (!Button(form).Checked) Button(form).PerformClick(); Pump(() => form.ExplorerProjection != null && ControlOrNull(form)?.Visible == true); }
        private static SolutionObjectsExplorerControl Control(MembershipResultsForm form) => form.Controls.OfType<SolutionObjectsExplorerControl>().Single();
        private static SolutionObjectsExplorerControl ControlOrNull(MembershipResultsForm form) => form.Controls.OfType<SolutionObjectsExplorerControl>().SingleOrDefault();
        private static DataGridView Grid(MembershipResultsForm form) => form.Controls.OfType<DataGridView>().Single();
        private static ToolStripButton Button(MembershipResultsForm form) => Controls(form).OfType<ToolStrip>().SelectMany(t => t.Items.OfType<ToolStripButton>()).Single(b => b.Name == "SolutionObjectsExplorerAction");
        private static ToolStripComboBox View(MembershipResultsForm form) => Controls(form).OfType<ToolStrip>().SelectMany(t => t.Items.OfType<ToolStripComboBox>()).Single(c => c.Name == "MembershipViewSelector");
        private static IEnumerable<System.Windows.Forms.Control> Controls(System.Windows.Forms.Control root) => root.Controls.Cast<System.Windows.Forms.Control>().SelectMany(c => new[] { c }.Concat(Controls(c)));
        private static IEnumerable<TreeNode> Walk(TreeNodeCollection roots) => roots.Cast<TreeNode>().SelectMany(n => new[] { n }.Concat(Walk(n.Nodes)));
        private static void Pump(Func<bool> finished) { var timer = Stopwatch.StartNew(); while (!finished() && timer.Elapsed.TotalSeconds < 20) { Application.DoEvents(); Thread.Sleep(1); } Assert.IsTrue(finished(), "Local UI update did not complete."); }
        private static void Save(Form form, string path, string name) { Application.DoEvents(); using (var bitmap = new Bitmap(form.Width, form.Height)) { form.DrawToBitmap(bitmap, new Rectangle(Point.Empty, form.Size)); bitmap.Save(Path.Combine(path, name), ImageFormat.Png); } }
        private static void Sta(Action action) { Exception error = null; var thread = new Thread(() => { try { action(); } catch (Exception ex) { error = ex; } }); thread.IsBackground = true; thread.SetApartmentState(ApartmentState.STA); thread.Start(); Assert.IsTrue(thread.Join(60000)); if (error != null) ExceptionDispatchInfo.Capture(error).Throw(); }
    }
}
