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
using D365SolutionComparer.Models.ComponentDetails;
using D365SolutionComparer.Models.Presentation;
using D365SolutionComparer.Services.Membership;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace D365SolutionComparer.Tests
{
    [TestClass, TestCategory("H2DExplorerUi"), TestCategory("H2EInitialExplorer")]
    public class SolutionObjectsExplorerUiTests
    {
        public TestContext TestContext { get; set; }
        [TestMethod]
        public void ExplorerIsDefaultAndExistingViewsAndDifferencesPredicateAreUnchanged()
        {
            Sta(() =>
            {
                var f = ExplorerFixture.Saa(); var spy = new FakeOrganizationService(); var csv = ExplorerFixture.Csv(f);
                using (var form = new MembershipResultsForm(f.Completed, sourceService: spy, targetService: spy))
                {
                    Assert.AreEqual(3, View(form).SelectedIndex); Assert.IsNull(form.ExplorerProjection);
                    form.Show(); Assert.IsFalse(Grid(form).Visible);
                    CollectionAssert.AreEqual(new[] { "Flat View", "Hierarchy View", "Differences Only", "Solution Objects Explorer" }, View(form).Items.Cast<string>().ToArray());
                    Assert.AreEqual(840, Grid(form).Rows.Count); Assert.IsNull(form.ExplorerProjection);
                    Open(form); Assert.AreEqual(840, form.ExplorerProjection.Rows.Count); Assert.AreEqual(1512, form.ExplorerProjection.Occurrences.Count);
                    View(form).SelectedIndex = 2; Application.DoEvents();
                    CollectionAssert.AreEqual(f.Presentation.Rows.Select((r, i) => new { r, i }).Where(x => x.r.Presence != MembershipPresence.PresentInBoth || x.r.DefinitionStatus == "Different").Select(x => x.i).ToArray(),
                        form.FilteredRows().Select(r => r.OriginalRowIndex).ToArray());
                    Open(form); var explorer = Control(form);
                    explorer.Navigation.SelectedNode = explorer.Navigation.Nodes[1]; // All Components, rather than the object-candidate subset.
                    CollectionAssert.AreEqual(form.FilteredRows().Select(r => r.OriginalRowIndex).ToArray(), explorer.VisibleRows.ToArray());
                    Assert.AreEqual(3, View(form).SelectedIndex); Assert.IsTrue(Button(form).Checked);
                    Assert.AreEqual(840, explorer.VisibleRows.Count, "Explorer is an authoritative view, not a hidden Differences Only overlay.");
                    var filter = Controls(form).OfType<ToolStrip>().SelectMany(t => t.Items.OfType<ToolStripComboBox>()).Single(c => c.Name == "MembershipPresenceFilter");
                    filter.SelectedIndex = 2; Open(form);
                    CollectionAssert.AreEqual(f.Presentation.Rows.Select((r, i) => new { r, i }).Where(x => x.r.Presence == MembershipPresence.OnlyInSource).Select(x => x.i).ToArray(), explorer.VisibleRows.ToArray());
                    filter.SelectedIndex = 0; Open(form);
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
                    View(form).SelectedIndex = 0; Assert.IsTrue(Grid(form).Visible); Open(form);
                    Button(form).PerformClick(); Assert.AreEqual(3, View(form).SelectedIndex);
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
            var path = Environment.GetEnvironmentVariable("H2E_EVIDENCE_DIRECTORY") ?? Environment.GetEnvironmentVariable("H2D_EVIDENCE_DIRECTORY"); if (string.IsNullOrWhiteSpace(path)) return;
            Sta(() =>
            {
                Directory.CreateDirectory(path); var apps = SyntheticNames(ExplorerFixture.Apps()); using (var form = new MembershipResultsForm(apps.Completed))
                {
                    form.Size = new Size(1600, 900); form.Show(); Open(form); Save(form, path, "01-initial-explorer.png");
                    View(form).SelectedIndex = 0; Save(form, path, "05-flat.png");
                    View(form).SelectedIndex = 1; Pump(() => form.Controls.OfType<MembershipHierarchyControl>().Any(h => h.Visible)); Save(form, path, "06-hierarchy.png");
                    View(form).SelectedIndex = 2; Save(form, path, "07-differences-only.png"); Open(form);
                    var c = Control(form); int parent = form.ExplorerProjection.Occurrences.Single(o => o.Original.Side == DisplayEnvironmentSide.Source && o.Original.Record.ComponentType == 80).Original.OccurrenceIndex;
                    c.SelectOccurrence(parent); c.Navigation.SelectedNode.Expand();
                    var group = c.Navigation.SelectedNode.Nodes.Cast<TreeNode>().Single(n => n.Text.StartsWith("AppSettings — structural"));
                    group.Expand(); c.Navigation.SelectedNode = group.Nodes[0]; Save(form, path, "02-appsettings.png");
                    parent = form.ExplorerProjection.Occurrences.Single(o => o.Original.Side == DisplayEnvironmentSide.Target && o.Original.Record.ComponentType == 80).Original.OccurrenceIndex;
                    c.SelectOccurrence(parent); c.Navigation.SelectedNode.Expand();
                    group = c.Navigation.SelectedNode.Nodes.Cast<TreeNode>().Single(n => n.Text.StartsWith("AppSettings — structural"));
                    group.Expand(); c.Navigation.SelectedNode = group.Nodes[0]; Save(form, path, "03-target-app.png");
                }
                var table = ExplorerFixture.Identity(1, "account"); using (var form = new MembershipResultsForm(SyntheticNames(ExplorerFixture.Make(new[] { table, ExplorerFixture.Identity(2, "account.syn_regionlocation", table.Record.SolutionComponentId) })).Completed))
                {
                    form.Size = new Size(1600, 900); form.Show(); Open(form); var c = Control(form);
                    c.SelectOccurrence(form.ExplorerProjection.Occurrences.Single(o => o.Original.Record.ComponentType == 1).Original.OccurrenceIndex);
                    c.Navigation.SelectedNode.Expand(); Save(form, path, "04-table-links.png");
                }
            });
        }

        [DataTestMethod, TestCategory("H2EStartupReadiness"), DataRow(true), DataRow(false)]
        public void ShownViewAndFilterEventsDeferProjectionUntilAllCallbacksComplete(bool explorerSelected)
        {
            Sta(() =>
            {
                using (var form = new MembershipResultsForm(ExplorerFixture.Saa().Completed))
                {
                    int hierarchyCalls = 0, explorerCalls = 0; var hierarchy = form.ProjectHierarchy; var explorer = form.ProjectExplorer;
                    form.ProjectHierarchy = t => { Interlocked.Increment(ref hierarchyCalls); return hierarchy(t); };
                    form.ProjectExplorer = (h, t) => { Interlocked.Increment(ref explorerCalls); return explorer(h, t); };
                    form.Shown += (sender, args) =>
                    {
                        Scope(form).SelectedIndex = 2;
                        WorkspaceButton(form, "ObjectsDifferencesFilter").PerformClick();
                        WorkspaceButton(form, "TechnicalMembershipAction").PerformClick();
                        View(form).SelectedIndex = 1;
                        var filter = Controls(form).OfType<ToolStrip>().SelectMany(t => t.Items.OfType<ToolStripComboBox>()).Single(c => c.Name == "MembershipPresenceFilter");
                        filter.SelectedIndex = 2; filter.SelectedIndex = 0;
                        WorkspaceButton(form, "ObjectsBackAction").PerformClick();
                        if (!explorerSelected) { WorkspaceButton(form, "TechnicalMembershipAction").PerformClick(); View(form).SelectedIndex = 1; }
                        form.ShowExplorerAsync().GetAwaiter().GetResult();
                        form.ShowHierarchyAsync().GetAwaiter().GetResult();
                        Application.DoEvents();
                        Assert.AreEqual(0, hierarchyCalls + explorerCalls, "Shown selection/filter events must not enter either projector.");
                        Assert.IsNull(form.ExplorerProjection); Assert.IsNull(form.HierarchyProjection);
                    };
                    form.Shown += (sender, args) => Assert.AreEqual(0, hierarchyCalls + explorerCalls, "Later Shown subscribers must also finish before projection.");
                    form.Show();
                    if (explorerSelected) { OpenPrimary(form); Assert.AreEqual(1, explorerCalls); Assert.AreEqual(3, View(form).SelectedIndex); }
                    else { Pump(() => form.HierarchyProjection != null); Assert.AreEqual(0, explorerCalls); Assert.AreEqual(1, View(form).SelectedIndex); Assert.IsTrue(form.IsTechnicalWorkspace); }
                    Assert.AreEqual(1, hierarchyCalls, "Only the final selected view starts after Shown completes.");
                }
            });
        }

        [TestMethod]
        public void InitialExplorerStartsAfterShownAndNeverInvokesServicesOrCaptures()
        {
            Sta(() =>
            {
                var f = ExplorerFixture.Saa(); var spy = new FakeOrganizationService(); int captures = 0, hierarchyCalls = 0, explorerCalls = 0;
                var csv = ExplorerFixture.Csv(f); var rows = f.Presentation.Rows.ToArray();
                var definitions = rows.Select(r => r.DefinitionDetail).ToArray();
                using (var form = new MembershipResultsForm(f.Completed, sourceService: spy, targetService: spy, captureAppSettingEvidence: () => captures++))
                {
                    var hierarchy = form.ProjectHierarchy; var explorer = form.ProjectExplorer;
                    form.ProjectHierarchy = t => { Interlocked.Increment(ref hierarchyCalls); return hierarchy(t); };
                    form.ProjectExplorer = (h, t) => { Interlocked.Increment(ref explorerCalls); return explorer(h, t); };
#if DEBUG
                    form.CaptureType300Evidence = () => captures++; form.CaptureType10072Evidence = () => captures++;
                    form.CaptureAppActionEvidence = () => captures++; form.CaptureType92Evidence = () => captures++;
#endif
                    Assert.AreEqual(0, hierarchyCalls + explorerCalls); Assert.IsNull(form.ExplorerProjection);
                    bool shown = false;
                    form.Shown += (sender, args) =>
                    {
                        shown = true; Assert.IsTrue(form.IsHandleCreated);
                        Assert.AreEqual(3, View(form).SelectedIndex); Assert.IsFalse(Grid(form).Visible);
                        Assert.IsTrue(Controls(form).Single(c => c.Name == "MembershipViewPreparation").Visible);
                        Assert.AreEqual(0, hierarchyCalls + explorerCalls, "Construction/Shown callbacks precede queued initial work.");
                    };
                    form.Show(); Open(form); Assert.IsTrue(shown); Assert.IsFalse(Grid(form).Visible);
                    Assert.AreEqual(1, hierarchyCalls); Assert.AreEqual(1, explorerCalls);
                    Assert.AreEqual(840, form.ExplorerProjection.Rows.Count); Assert.AreEqual(1512, form.ExplorerProjection.Occurrences.Count);
                    for (int i = 0; i < rows.Length; i++)
                    { Assert.AreSame(rows[i], f.Presentation.Rows[i]); Assert.AreSame(definitions[i], f.Presentation.Rows[i].DefinitionDetail); }
                    Assert.AreEqual(csv, ExplorerFixture.Csv(f)); Assert.AreEqual(0, spy.Calls + spy.ExecuteCalls + spy.WriteCalls + captures);
                    var summary = f.Presentation.Summary;
                    Assert.AreEqual(672, summary.PresentInBoth); Assert.AreEqual(103, summary.SourceOnly); Assert.AreEqual(32, summary.TargetOnly);
                    Assert.AreEqual(26, summary.IndeterminateUnsupported); Assert.AreEqual(4, summary.IndeterminateAmbiguous); Assert.AreEqual(3, summary.IndeterminateResolved);
                }
            });
        }

        [DataTestMethod, DataRow(0), DataRow(1)]
        public void ExplicitLegacySelectionWinsDuringBlockedInitialExplorerAndReusesCache(int legacyIndex)
        {
            Sta(() =>
            {
                using (var gate = new ManualResetEventSlim()) using (var started = new ManualResetEventSlim()) using (var exited = new ManualResetEventSlim())
                using (var form = new MembershipResultsForm(ExplorerFixture.Apps().Completed))
                {
                    int calls = 0; var original = form.ProjectExplorer;
                    form.ProjectExplorer = (h, t) => { Interlocked.Increment(ref calls); started.Set(); gate.Wait(t); var result = original(h, t); exited.Set(); return result; };
                    form.Show(); Pump(() => started.IsSet);
                    Assert.IsFalse(Grid(form).Visible); Assert.IsTrue(Button(form).Checked);
                    View(form).SelectedIndex = legacyIndex;
                    if (legacyIndex == 1) Pump(() => form.Controls.OfType<MembershipHierarchyControl>().Any(h => h.Visible));
                    Assert.IsFalse(Button(form).Checked); gate.Set(); Pump(() => exited.IsSet);
                    // Drain the completion callback; a stale Explorer result must not repaint.
                    for (int i = 0; i < 100; i++) { Application.DoEvents(); Thread.Sleep(1); }
                    Assert.AreEqual(legacyIndex, View(form).SelectedIndex); Assert.IsFalse(ControlOrNull(form)?.Visible ?? false);
                    Assert.AreEqual(legacyIndex == 0, Grid(form).Visible); Assert.IsNull(form.ExplorerProjection);
                    Open(form); Assert.AreEqual(1, calls); Assert.AreEqual(3, View(form).SelectedIndex);
                    form.Hide(); form.Show(); Open(form); Assert.AreEqual(1, calls);
                }
            });
        }

        [TestMethod]
        public void PendingHierarchyAndFilterChangesShareOneInitialProjection()
        {
            Sta(() =>
            {
                using (var gate = new ManualResetEventSlim()) using (var started = new ManualResetEventSlim())
                using (var form = new MembershipResultsForm(ExplorerFixture.Saa().Completed))
                {
                    int hierarchyCalls = 0, explorerCalls = 0; var hierarchy = form.ProjectHierarchy; var explorer = form.ProjectExplorer;
                    form.ProjectHierarchy = t => { Interlocked.Increment(ref hierarchyCalls); started.Set(); gate.Wait(t); return hierarchy(t); };
                    form.ProjectExplorer = (h, t) => { Interlocked.Increment(ref explorerCalls); return explorer(h, t); };
                    form.Show(); Pump(() => started.IsSet);
                    var filter = Controls(form).OfType<ToolStrip>().SelectMany(t => t.Items.OfType<ToolStripComboBox>()).Single(c => c.Name == "MembershipPresenceFilter");
                    filter.SelectedIndex = 2; filter.SelectedIndex = 3; filter.SelectedIndex = 2;
                    gate.Set(); Open(form); var control = Control(form); control.Navigation.SelectedNode = control.Navigation.Nodes[1];
                    CollectionAssert.AreEqual(form.FilteredRows().Select(r => r.OriginalRowIndex).ToArray(), control.VisibleRows.ToArray());
                    Assert.AreEqual(103, control.VisibleRows.Count); Assert.AreEqual(1, hierarchyCalls); Assert.AreEqual(1, explorerCalls);
                }
            });
        }

        [TestMethod]
        public void InitialFailureFallsBackOnceAndDoesNotAutomaticallyRelaunch()
        {
            Sta(() =>
            {
                using (var form = new MembershipResultsForm(ExplorerFixture.Apps().Completed))
                {
                    int calls = 0; form.ProjectExplorer = (h, t) => { Interlocked.Increment(ref calls); throw new InvalidOperationException("PRIVATE TOKEN https://secret.invalid"); };
                    form.Show(); Pump(() => form.HierarchyDiagnostics.Any());
                    Assert.AreEqual(0, View(form).SelectedIndex); Assert.IsFalse(Button(form).Checked); Assert.IsTrue(Grid(form).Visible);
                    form.Hide(); form.Show(); Application.DoEvents(); Assert.AreEqual(1, calls);
                    Assert.IsFalse(Controls(form).Single(c => c.Name == "MembershipViewPreparation").Visible);
                    Assert.IsFalse(form.HierarchyDiagnostics.Any(d => d.Contains("PRIVATE") || d.Contains("https")));
                    Assert.AreEqual(6, Grid(form).Rows.Count);
                }
            });
        }

        [TestMethod]
        public void OriginalRowLinkSelectsFlatAndRetainsTargetSideOnReturn()
        {
            Sta(() =>
            {
                using (var form = new MembershipResultsForm(ExplorerFixture.Apps().Completed))
                {
                    form.Show(); Open(form); var control = Control(form); control.SelectComparison(3, DisplayEnvironmentSide.Target);
                    var link = Controls(control).OfType<LinkLabel>().Single(l => l.Text == "Open original row in Component Details");
                    Assert.IsTrue(link.Enabled);
                    typeof(LinkLabel).GetMethod("OnLinkClicked", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic)
                        .Invoke(link, new object[] { new LinkLabelLinkClickedEventArgs(link.Links[0]) });
                    Assert.AreEqual(0, View(form).SelectedIndex); Assert.IsFalse(Button(form).Checked);
                    Assert.AreEqual(3, ((MembershipGridRow)Grid(form).CurrentRow.DataBoundItem).OriginalRowIndex);
                    var metadata = Controls(form).OfType<CheckBox>().Single(c => c.Name == "OptionalMetadataToggle"); metadata.Checked = true;
                    Open(form); Assert.AreEqual(3, control.SelectedComparison.OriginalRowIndex); Assert.AreEqual(DisplayEnvironmentSide.Target, control.SelectedSide);
                    Assert.AreEqual(11, control.Grid.Columns.Cast<DataGridViewColumn>().Count(c => c.Visible));
                    View(form).SelectedIndex = 1; Pump(() => form.Controls.OfType<MembershipHierarchyControl>().Any(h => h.Visible));
                    Open(form); Assert.AreEqual(DisplayEnvironmentSide.Target, control.SelectedSide);
                    Assert.AreEqual(3, control.SelectedComparison.OriginalRowIndex);
                }
            });
        }

        [TestMethod]
        public void CloseDuringInitialWorkIsPromptAndCancelsWorker()
        {
            Sta(() =>
            {
                using (var started = new ManualResetEventSlim()) using (var cancelled = new ManualResetEventSlim())
                using (var form = new MembershipResultsForm(ExplorerFixture.Apps().Completed))
                {
                    form.ProjectExplorer = (h, t) => { started.Set(); t.WaitHandle.WaitOne(); cancelled.Set(); t.ThrowIfCancellationRequested(); return null; };
                    form.Show(); Pump(() => started.IsSet); var timer = Stopwatch.StartNew(); form.Close(); timer.Stop();
                    Assert.IsTrue(timer.ElapsedMilliseconds < 1000); Assert.IsTrue(cancelled.Wait(5000)); Assert.IsNull(form.ExplorerProjection);
                    TestContext.WriteLine("H2E initial-close_ms={0:F3}; cancellation observed; no projection published", timer.Elapsed.TotalMilliseconds);
                }
            });
        }

        [TestMethod]
        public void SyntheticInitialPreparationAndFailureScreenshots()
        {
            var path = Environment.GetEnvironmentVariable("H2E_EVIDENCE_DIRECTORY"); if (string.IsNullOrWhiteSpace(path)) return;
            Sta(() =>
            {
                Directory.CreateDirectory(path);
                using (var started = new ManualResetEventSlim()) using (var form = new MembershipResultsForm(SyntheticNames(ExplorerFixture.Apps()).Completed))
                {
                    form.ProjectExplorer = (h, t) => { started.Set(); t.WaitHandle.WaitOne(); t.ThrowIfCancellationRequested(); return null; };
                    form.Show(); Pump(() => started.IsSet); Save(form, path, "08-initial-preparation.png"); form.Close();
                }
                using (var form = new MembershipResultsForm(SyntheticNames(ExplorerFixture.Apps()).Completed))
                {
                    form.ProjectExplorer = (h, t) => throw new InvalidOperationException("private synthetic diagnostic");
                    form.Show(); Pump(() => form.HierarchyDiagnostics.Any()); Save(form, path, "09-flat-fallback.png");
                }
            });
        }

        [TestMethod, TestCategory("H2E2B1")]
        public void InitialTargetWorkspaceIsPartialAndTechnicalConservesAllDataWithoutReads()
        {
            Sta(() =>
            {
                var f = ExplorerFixture.Saa(); var spy = new FakeOrganizationService(); int captures = 0; var csv = ExplorerFixture.Csv(f);
                using (var form = new MembershipResultsForm(f.Completed, captureAppSettingEvidence: () => captures++, sourceService: spy, targetService: spy))
                {
                    form.Show(); OpenPrimary(form); var control = Control(form);
                    Assert.IsFalse(form.IsTechnicalWorkspace); Assert.IsTrue(control.IsPrimaryWorkspace); Assert.AreEqual(ObjectsWorkspaceScope.Target, control.WorkspaceScope);
                    Assert.AreEqual(0, Scope(form).SelectedIndex); StringAssert.Contains(Scope(form).SelectedItem.ToString(), f.Presentation.Target.Diagnostics.EnvironmentName);
                    Assert.IsFalse(Grid(form).Visible); Assert.IsFalse(Controls(form).OfType<ToolStrip>().Single(t => t.Name == "MembershipViewBar").Visible);
                    Assert.IsFalse(control.Navigation.Nodes.Cast<TreeNode>().Any(n => n.Text.StartsWith("Source:") || n.Text.StartsWith("Target:") || n.Text.StartsWith("AppSettings")));
                    StringAssert.Contains(Controls(control).OfType<Label>().Single(l => l.Name == "ExplorerAccounting").Text, "partial");
                    Assert.IsFalse(control.Navigation.Nodes.Cast<TreeNode>().Any(n => n.Text.Contains("150")));
                    Scope(form).SelectedIndex = 2; OpenPrimary(form);
                    var candidateRows = form.ExplorerProjection.Occurrences.Where(o => o.PrimaryCategory != null).SelectMany(o => o.Original.ComparisonRowIndices).Distinct().OrderBy(i => i).ToArray();
                    CollectionAssert.AreEqual(candidateRows, control.VisibleRows.ToArray());
                    var differences = WorkspaceButton(form, "ObjectsDifferencesFilter"); differences.PerformClick(); OpenPrimary(form);
                    var expected = f.Presentation.Rows.Select((row, i) => new { row, i }).Where(x => x.row.Presence != MembershipPresence.PresentInBoth || x.row.DefinitionStatus == "Different").Select(x => x.i).ToArray();
                    CollectionAssert.AreEqual(expected.Intersect(candidateRows).OrderBy(i => i).ToArray(), control.VisibleRows.ToArray());
                    Assert.IsTrue(expected.Any(i => f.Presentation.Rows[i].SourceIdentity?.Status == IdentityResolutionStatus.Unsupported || f.Presentation.Rows[i].TargetIdentity?.Status == IdentityResolutionStatus.Unsupported));
                    Open(form); Assert.AreEqual(840, form.ExplorerProjection.Rows.Count); Assert.AreEqual(1512, form.ExplorerProjection.Occurrences.Count);
                    View(form).SelectedIndex = 2; CollectionAssert.AreEqual(expected, form.FilteredRows().Select(r => r.OriginalRowIndex).ToArray());
                    Assert.AreEqual(840, form.ExplorerProjection.Category("all").RowIndices.Count);
                    for (int view = 0; view < 4; view++) { View(form).SelectedIndex = view; if (view == 1) Pump(() => form.Controls.OfType<MembershipHierarchyControl>().Any(h => h.Visible)); if (view == 3) Open(form); Assert.AreEqual(view, View(form).SelectedIndex); }
                    Assert.AreEqual(csv, ExplorerFixture.Csv(f)); Assert.AreEqual(0, captures + spy.Calls + spy.ExecuteCalls + spy.WriteCalls);
                }
            });
        }
        [TestMethod, TestCategory("H2E2B1")]
        public void PrimarySettingsRequireStructuralProofAndKeepTargetSide()
        {
            Sta(() =>
            {
                using (var form = new MembershipResultsForm(ExplorerFixture.Apps().Completed))
                {
                    form.Show(); OpenPrimary(form); var c = Control(form);
                    var apps = c.Navigation.Nodes.Cast<TreeNode>().Single(n => n.Text.StartsWith("Apps (")); apps.Expand(); var app = apps.Nodes[0]; app.Expand();
                    var settings = app.Nodes.Cast<TreeNode>().Single(n => n.Text.StartsWith("Settings — structural")); settings.Expand(); c.Navigation.SelectedNode = settings.Nodes[0];
                    Assert.AreEqual(DisplayEnvironmentSide.Target, c.SelectedSide); StringAssert.Contains(c.EvidenceText, "StructuralParent: Verified");
                    Assert.IsFalse(c.Navigation.Nodes.Cast<TreeNode>().Any(n => n.Text.StartsWith("AppSettings")));
                    Open(form); var target = c.Navigation.Nodes.Cast<TreeNode>().Single(n => n.Tag == null && n.Text.StartsWith("Target:"));
                    StringAssert.Contains(target.Nodes.Cast<TreeNode>().Single(n => n.Text.StartsWith("AppSettings")).Text, "5 raw");
                }
                var appIdentity = ExplorerFixture.Identity(80, "app"); var child = ExplorerFixture.Identity(10075, ExplorerFixture.Key("app", "setting"), appIdentity.Record.SolutionComponentId);
                using (var form = new MembershipResultsForm(ExplorerFixture.Make(new[] { appIdentity, child }, new[] { appIdentity, child }).Completed))
                {
                    form.Show(); OpenPrimary(form); var c = Control(form); var apps = c.Navigation.Nodes.Cast<TreeNode>().Single(n => n.Text.StartsWith("Apps (")); apps.Expand(); apps.Nodes[0].Expand();
                    Assert.IsFalse(apps.Nodes[0].Nodes.Cast<TreeNode>().Any(n => n.Text.StartsWith("Settings")));
                    Assert.AreEqual(4, form.ExplorerProjection.Occurrences.Count); Assert.AreEqual(2, form.ExplorerProjection.Rows.Count);
                }
            });
        }
        [TestMethod, TestCategory("H2E2B1")]
        public void OutsideScopeSelectionHasExplanationAndOriginalRowReturn()
        {
            Sta(() =>
            {
                using (var form = new MembershipResultsForm(ExplorerFixture.Saa().Completed))
                {
                    form.Show(); OpenPrimary(form); Scope(form).SelectedIndex = 1; OpenPrimary(form); var c = Control(form);
                    int original = form.CompletedComparison.Presentation.Rows.Select((row, i) => new { row, i }).First(x => x.row.Presence == MembershipPresence.OnlyInSource).i;
                    c.SelectComparison(original, DisplayEnvironmentSide.Source); Assert.AreEqual(original, c.SelectedComparison.OriginalRowIndex);
                    Scope(form).SelectedIndex = 0; OpenPrimary(form);
                    Assert.AreEqual(original, c.SelectedComparison.OriginalRowIndex); Assert.AreEqual(DisplayEnvironmentSide.Source, c.SelectedSide);
                    StringAssert.Contains(Controls(c).OfType<Label>().Single(l => l.Name == "ExplorerAccounting").Text, "outside");
                    WorkspaceButton(form, "ObjectsOriginalRowAction").PerformClick();
                    Assert.IsTrue(form.IsTechnicalWorkspace); Assert.AreEqual(0, View(form).SelectedIndex); Assert.AreEqual(original, ((MembershipGridRow)Grid(form).CurrentRow.DataBoundItem).OriginalRowIndex);
                }
            });
        }
        [DataTestMethod, TestCategory("H2E2B1"), DataRow(900), DataRow(1400), DataRow(1920)]
        public void PrimaryCriticalControlsFitAndOptionalMetadataDoesNotWidenGrid(int width)
        {
            Sta(() =>
            {
                using (var form = new MembershipResultsForm(ExplorerFixture.Apps().Completed))
                {
                    form.Size = new Size(width, 520); form.Show(); OpenPrimary(form); var c = Control(form);
                    var bar = Controls(form).OfType<ToolStrip>().Single(t => t.Name == "ObjectsWorkspaceBar");
                    foreach (var name in new[] { "ObjectsScopeSelector", "ObjectsDetailsAction", "ObjectsOriginalRowAction", "ObjectsMetadataAction" })
                    { var item = bar.Items.Cast<ToolStripItem>().Single(i => i.Name == name); Assert.IsTrue(item.Visible); Assert.IsTrue(item.Bounds.Right <= bar.ClientSize.Width); }
                    WorkspaceButton(form, "ObjectsDetailsAction").PerformClick();
                    Assert.AreEqual(6, c.Grid.Columns.Cast<DataGridViewColumn>().Count(column => column.Visible));
                    Assert.IsTrue(c.Grid.Columns["Source"].Visible); Assert.IsTrue(c.Grid.Columns["Target"].Visible);
                    Scope(form).SelectedIndex = 2; OpenPrimary(form);
                    Assert.AreEqual(6, c.Grid.Columns.Cast<DataGridViewColumn>().Count(column => column.Visible));
                    Assert.IsTrue(c.Grid.VirtualMode); Assert.IsTrue(c.Grid.ReadOnly);
                    Scope(form).SelectedIndex = 1; OpenPrimary(form); Assert.AreEqual(ObjectsWorkspaceScope.Source, c.WorkspaceScope);
                    Assert.IsFalse(c.Navigation.Nodes.Cast<TreeNode>().Any(n => n.Text.StartsWith("Target:")));
                }
            });
        }
        [DataTestMethod, TestCategory("H2E2B1"), DataRow(840), DataRow(10000)]
        public void PrimaryNativeNavigationRemainsLazyForLargeSolutions(int count)
        {
            Sta(() =>
            {
                var identities = Enumerable.Range(0, count).Select(i => ExplorerFixture.Identity(61, "synthetic" + i)).ToArray();
                using (var form = new MembershipResultsForm(ExplorerFixture.Make(identities, identities).Completed))
                {
                    var watch = Stopwatch.StartNew(); form.Show(); OpenPrimary(form); watch.Stop(); var c = Control(form);
                    Assert.AreEqual(count, c.Grid.RowCount); Assert.IsTrue(c.Grid.VirtualMode); Assert.IsTrue(Walk(c.Navigation.Nodes).Count() < 50);
                    var category = c.Navigation.Nodes.Cast<TreeNode>().Single(n => n.Text.StartsWith("Web resources")); category.Expand(); category.Nodes[0].Expand();
                    Assert.AreEqual(200, category.Nodes[0].Nodes.Count); Assert.IsTrue(Walk(c.Navigation.Nodes).Count() < 600);
                    Assert.IsTrue(watch.Elapsed.TotalSeconds < 15);
                    TestContext.WriteLine("H2E2B1 primary rows={0}; open_ms={1:F3}; native_nodes={2}; requests=0", count, watch.Elapsed.TotalMilliseconds, Walk(c.Navigation.Nodes).Count());
                }
            });
        }
        [DataTestMethod, TestCategory("H2E2B1"), DataRow(false), DataRow(true)]
        public void RealEmptyInventoryAndUnavailableInventoryHaveDifferentPresentation(bool unavailable)
        {
            Sta(() =>
            {
                var f = ExplorerFixture.Make(new[] { ExplorerFixture.Identity(61, "observed") });
                var presentation = unavailable ? new MembershipResultPresenter().Create(f.Presentation.Source,
                    MembershipEnvironmentResult.Unavailable("Target", "SAA", 0, TimeSpan.Zero, "Unavailable")) : f.Presentation;
                using (var form = new MembershipResultsForm(presentation))
                {
                    form.Show(); OpenPrimary(form); var c = Control(form); var text = Controls(c).OfType<Label>().Single(l => l.Name == "ExplorerAccounting").Text;
                    StringAssert.Contains(text, "partial");
                    if (unavailable) { StringAssert.Contains(text, "counts not established"); StringAssert.Contains(c.Navigation.Nodes[0].Text, "Not established"); }
                    else { StringAssert.Contains(text, "0 retained raw occurrences"); StringAssert.Contains(c.Navigation.Nodes[0].Text, "0 candidates"); StringAssert.Contains(c.Navigation.Nodes[0].ToolTipText, "raw occurrences"); }
                    Assert.AreEqual(0, c.VisibleRows.Count);
                }
            });
        }
        [TestMethod, TestCategory("H2E2B1")]
        public void FifteenSettingsPerSideRemainTechnicalWithoutInventedParents()
        {
            Sta(() =>
            {
                Func<ComponentIdentity[]> local = () => new[] { ExplorerFixture.Identity(80, "synthetic-app") }.Concat(
                    Enumerable.Range(0, 15).Select(i => ExplorerFixture.Identity(10075, ExplorerFixture.Key("synthetic-app", "setting" + i)))).ToArray();
                var f = ExplorerFixture.Make(local(), local()); var csv = ExplorerFixture.Csv(f);
                using (var form = new MembershipResultsForm(f.Completed))
                {
                    form.Show(); OpenPrimary(form); Assert.AreEqual(1, Control(form).VisibleRows.Count);
                    Assert.IsFalse(Control(form).Navigation.Nodes.Cast<TreeNode>().Any(n => n.Text.StartsWith("AppSettings")));
                    Open(form); var c = Control(form);
                    foreach (var side in new[] { "Source:", "Target:" })
                    {
                        var node = c.Navigation.Nodes.Cast<TreeNode>().Single(n => n.Text.StartsWith(side));
                        var settings = node.Nodes.Cast<TreeNode>().Single(n => n.Text.StartsWith("AppSettings")); StringAssert.Contains(settings.Text, "15 raw");
                        c.Navigation.SelectedNode = settings; Assert.AreEqual(15, c.VisibleRows.Count);
                    }
                    Assert.AreEqual(30, form.ExplorerProjection.Occurrences.Count(o => o.Original.Record.ComponentType == 10075));
                    Assert.IsFalse(form.ExplorerProjection.Occurrences.Any(o => o.Children.Any(l => l.Group == "AppSettings")));
                    Assert.AreEqual(csv, ExplorerFixture.Csv(f));
                }
            });
        }
        [TestMethod, TestCategory("H2E2B1")]
        public void DefinitionDifferenceIsIncludedWithoutChangingMembership()
        {
            Sta(() =>
            {
                var f = ExplorerFixture.Apps(); var old = f.Completed; var rows = f.Presentation.Rows.ToArray(); rows[0] = rows[0].WithDefinition("Different", "synthetic property", null);
                f.Presentation = new MembershipComparisonPresentation("SAA", f.Presentation.Source, f.Presentation.Target, rows);
                f.Completed = new CompletedMembershipComparison(f.Presentation, old.DisplayEvidence, DisplayProjectionOutcome.Completed);
                using (var form = new MembershipResultsForm(f.Completed))
                {
                    form.Show(); OpenPrimary(form); Scope(form).SelectedIndex = 2; WorkspaceButton(form, "ObjectsDifferencesFilter").PerformClick(); OpenPrimary(form);
                    CollectionAssert.AreEqual(new[] { 0 }, Control(form).VisibleRows.ToArray());
                    Assert.AreEqual(MembershipPresence.PresentInBoth, rows[0].Presence); Assert.AreEqual("Different", rows[0].DefinitionStatus);
                    Open(form); View(form).SelectedIndex = 2; CollectionAssert.AreEqual(new[] { 0 }, form.FilteredRows().Select(x => x.OriginalRowIndex).ToArray());
                }
            });
        }
        [DataTestMethod, TestCategory("H2E2B1"), DataRow(false), DataRow(true)]
        public void RapidScopeOrTechnicalSwitchCannotPublishStalePrimary(bool technical)
        {
            Sta(() =>
            {
                using (var gate = new ManualResetEventSlim()) using (var started = new ManualResetEventSlim()) using (var completed = new ManualResetEventSlim())
                using (var form = new MembershipResultsForm(ExplorerFixture.Saa().Completed))
                {
                    int calls = 0; var original = form.ProjectExplorer;
                    form.ProjectExplorer = (h, t) => { Interlocked.Increment(ref calls); started.Set(); gate.Wait(t); var result = original(h, t); completed.Set(); return result; };
                    form.Show(); Pump(() => started.IsSet); Scope(form).SelectedIndex = 1; Scope(form).SelectedIndex = 2;
                    WorkspaceButton(form, "ObjectsDifferencesFilter").PerformClick();
                    if (technical) WorkspaceButton(form, "TechnicalMembershipAction").PerformClick();
                    gate.Set(); Pump(() => completed.IsSet);
                    if (technical)
                    {
                        for (int i = 0; i < 100; i++) { Application.DoEvents(); Thread.Sleep(1); }
                        Assert.IsTrue(Grid(form).Visible); Assert.IsTrue(form.IsTechnicalWorkspace); Assert.IsFalse(ControlOrNull(form)?.Visible ?? false);
                    }
                    else { OpenPrimary(form); Assert.AreEqual(ObjectsWorkspaceScope.Compare, Control(form).WorkspaceScope);
                        var expected = form.ExplorerProjection.Occurrences.Where(o => o.PrimaryCategory != null).SelectMany(o => o.Original.ComparisonRowIndices).Distinct()
                            .Where(i => form.CompletedComparison.Presentation.Rows[i].Presence != MembershipPresence.PresentInBoth || form.CompletedComparison.Presentation.Rows[i].DefinitionStatus == "Different").OrderBy(i => i).ToArray();
                        CollectionAssert.AreEqual(expected, Control(form).VisibleRows.ToArray()); }
                    Assert.AreEqual(1, calls);
                }
            });
        }
        [TestMethod, TestCategory("H2E2B1")]
        public void PrimarySyntheticAcceptanceScreenshots()
        {
            var path = Environment.GetEnvironmentVariable("H2E2B1_EVIDENCE_DIRECTORY"); if (string.IsNullOrWhiteSpace(path)) return;
            Sta(() =>
            {
                Directory.CreateDirectory(path); using (var form = new MembershipResultsForm(SyntheticNames(ExplorerFixture.Apps()).Completed))
                {
                    form.Size = new Size(1400, 760); form.Text += " — SYNTHETIC H2E-2B1"; form.Show(); OpenPrimary(form); Save(form, path, "01-target.png");
                    Scope(form).SelectedIndex = 1; OpenPrimary(form); Save(form, path, "02-source.png");
                    Scope(form).SelectedIndex = 2; OpenPrimary(form); Save(form, path, "03-compare.png");
                    WorkspaceButton(form, "ObjectsDifferencesFilter").PerformClick(); OpenPrimary(form); Save(form, path, "04-differences.png");
                    Open(form); View(form).SelectedIndex = 0; Save(form, path, "05-technical-all.png"); Open(form);
                    var c = Control(form); var target = c.Navigation.Nodes.Cast<TreeNode>().Single(n => n.Text.StartsWith("Target:"));
                    c.Navigation.SelectedNode = target.Nodes.Cast<TreeNode>().Single(n => n.Text.StartsWith("AppSettings")); Save(form, path, "06-technical-settings.png");
                }
                using (var form = new MembershipResultsForm(ExplorerFixture.Make(new[] { ExplorerFixture.Identity(29, "workflow"), ExplorerFixture.Identity(60, "form"), ExplorerFixture.Identity(300, null, status: IdentityResolutionStatus.Unsupported) }).Completed))
                {
                    form.Text += " — SYNTHETIC uncertain subtype"; form.Show(); OpenPrimary(form); Scope(form).SelectedIndex = 2; OpenPrimary(form); Save(form, path, "07-uncertain-subtypes.png");
                }
                using (var form = new MembershipResultsForm(SyntheticNames(ExplorerFixture.Apps()).Completed))
                {
                    form.Text += " — SYNTHETIC deliberate failure"; form.ProjectExplorer = (h, t) => throw new InvalidOperationException("PRIVATE");
                    form.Show(); Pump(() => form.HierarchyDiagnostics.Any()); Save(form, path, "08-fallback.png"); Assert.IsTrue(form.IsTechnicalWorkspace);
                }
            });
        }

        [TestMethod, TestCategory("H2E2C2A")]
        public void CandidateOnlyCompareKeepsTechnicalAllAndOutsideCandidateReturn()
        {
            Sta(() =>
            {
                var f = ExplorerFixture.Saa(); var csv = ExplorerFixture.Csv(f); var spy = new FakeOrganizationService();
                using (var form = new MembershipResultsForm(f.Completed, sourceService: spy, targetService: spy))
                {
                    form.Show(); OpenPrimary(form); var c = Control(form);
                    Assert.IsFalse(c.Navigation.Nodes.Cast<TreeNode>().Any(n => n.Text.StartsWith("Technical Membership")));
                    StringAssert.Contains(c.Navigation.Nodes[0].ToolTipText, "side-local candidate raw occurrences");
                    Scope(form).SelectedIndex = 2; OpenPrimary(form);
                    var candidates = form.ExplorerProjection.Occurrences.Where(o => o.PrimaryCategory != null).SelectMany(o => o.Original.ComparisonRowIndices).Distinct().OrderBy(i => i).ToArray();
                    CollectionAssert.AreEqual(candidates, c.VisibleRows.ToArray());
                    StringAssert.Contains(c.Navigation.Nodes[0].ToolTipText, "distinct candidate original row indices");
                    StringAssert.Contains(Controls(c).Single(x => x.Name == "ExplorerAccounting").AccessibleDescription, "partial");
                    int excluded = f.Presentation.Rows.Select((row, i) => new { row, i }).First(x => x.row.SourceIdentity?.Record.ComponentType == 300).i;
                    c.SelectComparison(excluded, DisplayEnvironmentSide.Source);
                    Assert.AreEqual(excluded, c.SelectedComparison.OriginalRowIndex); Assert.AreEqual(DisplayEnvironmentSide.Source, c.SelectedSide);
                    StringAssert.Contains(Controls(c).Single(x => x.Name == "ExplorerAccounting").AccessibleDescription, "outside");
                    WorkspaceButton(form, "ObjectsOriginalRowAction").PerformClick();
                    Assert.IsTrue(form.IsTechnicalWorkspace); Assert.AreEqual(0, View(form).SelectedIndex); Assert.AreEqual(840, Grid(form).Rows.Count);
                    Assert.AreEqual(excluded, ((MembershipGridRow)Grid(form).CurrentRow.DataBoundItem).OriginalRowIndex);
                    Open(form); c.Navigation.SelectedNode = c.Navigation.Nodes.Cast<TreeNode>().Single(n => n.Text.StartsWith("All Components"));
                    Assert.AreEqual(840, c.VisibleRows.Count); Assert.AreEqual(1512, form.ExplorerProjection.Occurrences.Count);
                    Assert.AreEqual(csv, ExplorerFixture.Csv(f)); Assert.AreEqual(0, spy.Calls + spy.ExecuteCalls + spy.WriteCalls);
                }
            });
        }

        [DataTestMethod, TestCategory("H2E2C2A"), DataRow(900, 100), DataRow(900, 125), DataRow(900, 150), DataRow(1400, 150), DataRow(1920, 200)]
        public void MeasuredHeaderAndWrappedToolbarFitLongNamesAndFontScale(int width, int percent)
        {
            Sta(() =>
            {
                var f = LongHeaderFixture(); using (var form = new MembershipResultsForm(f.Completed))
                {
                    form.Font = new Font("Segoe UI", 9F * percent / 100F); form.Size = new Size(width, 520); form.Show(); OpenPrimary(form); form.PerformLayout(); Application.DoEvents();
                    var header = Controls(form).Single(x => x.Name == "ObjectsSnapshotHeader");
                    foreach (var name in new[] { "ObjectsSolutionTitle", "ObjectsSourceState", "ObjectsTargetState" })
                    {
                        var label = (Label)Controls(form).Single(x => x.Name == name); Assert.IsTrue(label.Visible); Assert.IsFalse(label.AutoEllipsis);
                        var rect = header.RectangleToClient(label.RectangleToScreen(label.ClientRectangle));
                        Assert.IsTrue(header.ClientRectangle.Contains(rect), name + " must fit inside measured header");
                        Assert.AreEqual(label.Text, label.AccessibleDescription);
                        var measured = TextRenderer.MeasureText(label.Text, label.Font, new Size(label.Width, int.MaxValue), TextFormatFlags.WordBreak | TextFormatFlags.TextBoxControl);
                        Assert.IsTrue(measured.Height <= label.Height, name + " text must wrap without clipping");
                    }
                    var bar = Controls(form).OfType<ToolStrip>().Single(t => t.Name == "ObjectsWorkspaceBar");
                    foreach (var item in bar.Items.Cast<ToolStripItem>().Where(i => i.Name != "" && i.Available))
                    { Assert.IsTrue(item.Visible); Assert.IsTrue(bar.ClientRectangle.Contains(item.Bounds), item.Name + " must remain on wrapped toolbar"); }
                    Assert.IsTrue(WorkspaceButton(form, "TechnicalMembershipAction").Text.Contains("&"));
                    TestContext.WriteLine("H2E2C2A synthetic width={0}; font_scale={1}%; header_height={2}; toolbar_height={3}", width, percent, header.Height, bar.Height);
                }
            });
        }

        [TestMethod, TestCategory("H2E2C2A")]
        public void CompleteDiagnosticsAreScrollableAndKeyboardAccessibleInBothWorkspaces()
        {
            Sta(() =>
            {
                using (var form = new MembershipResultsForm(LongHeaderFixture().Completed))
                {
                    form.Show(); OpenPrimary(form); Open(form); var expand = (CheckBox)Controls(form).Single(c => c.Name == "ObjectsAdvancedDiagnostics");
                    Assert.IsTrue(expand.TabStop); Assert.IsFalse(expand.Checked); expand.Checked = true; Application.DoEvents();
                    var diagnostics = (TextBox)Controls(form).Single(c => c.Name == "CompleteMembershipDiagnostics");
                    Assert.IsTrue(diagnostics.Visible); Assert.IsTrue(diagnostics.ReadOnly); Assert.AreEqual(ScrollBars.Both, diagnostics.ScrollBars);
                    StringAssert.Contains(diagnostics.Text, "Total Comparison Rows: 840"); StringAssert.Contains(diagnostics.Text, "Raw Source Components=792"); StringAssert.Contains(diagnostics.Text, "Raw Target Components=720");
                    var host = (UserControl)Controls(form).Single(c => c.Name == "ObjectsSnapshotHeader");
                    Assert.IsTrue(host.SelectNextControl(expand, true, true, true, false)); Assert.AreEqual(diagnostics, host.ActiveControl);
                    Open(form); Assert.IsTrue(diagnostics.Visible); Assert.IsTrue(WorkspaceButton(form, "AdvancedDiagnosticsAction").Checked);
                    WorkspaceButton(form, "AdvancedDiagnosticsAction").PerformClick(); Assert.IsFalse(diagnostics.Visible);
                }
            });
        }

        [TestMethod, TestCategory("H2E2C2A")]
        public void ControlledSyntheticScreenshotsCoverCandidateAndCompleteTechnicalViews()
        {
            var path = Environment.GetEnvironmentVariable("H2E2C2A_EVIDENCE_DIRECTORY"); if (string.IsNullOrWhiteSpace(path)) return;
            Sta(() =>
            {
                Directory.CreateDirectory(path); using (var form = new MembershipResultsForm(SyntheticNames(ExplorerFixture.Saa()).Completed))
                {
                    form.Text += " — SYNTHETIC H2E-2C2A"; form.Show(); OpenPrimary(form); Save(form, path, "01-target-all.png");
                    Scope(form).SelectedIndex = 1; OpenPrimary(form); Save(form, path, "02-source-all.png");
                    Scope(form).SelectedIndex = 2; OpenPrimary(form); Save(form, path, "03-compare-all.png");
                    WorkspaceButton(form, "ObjectsDifferencesFilter").PerformClick(); OpenPrimary(form); Save(form, path, "04-compare-differences.png");
                    Open(form); View(form).SelectedIndex = 0; Assert.AreEqual(840, Grid(form).Rows.Count); Save(form, path, "08-technical-all-840.png");
                }
                using (var form = new MembershipResultsForm(SyntheticNames(ExplorerFixture.Apps()).Completed))
                {
                    form.Text += " — SYNTHETIC verified Settings"; form.Show(); OpenPrimary(form); var c = Control(form);
                    var apps = c.Navigation.Nodes.Cast<TreeNode>().Single(n => n.Text.StartsWith("Apps (")); apps.Expand(); apps.Nodes[0].Expand();
                    c.Navigation.SelectedNode = apps.Nodes[0].Nodes.Cast<TreeNode>().Single(n => n.Text.StartsWith("Settings — structural")); Save(form, path, "05-apps-verified-settings.png");
                }
                var table = ExplorerFixture.Identity(1, "account"); var rows = new[] { table, ExplorerFixture.Identity(2, "account.column", table.Record.SolutionComponentId), ExplorerFixture.Identity(380, "synthetic_environment_variable") };
                using (var form = new MembershipResultsForm(SyntheticNames(ExplorerFixture.Make(rows, rows)).Completed))
                {
                    form.Text += " — SYNTHETIC Tables and Environment Variables"; form.Show(); OpenPrimary(form); var c = Control(form);
                    var tables = c.Navigation.Nodes.Cast<TreeNode>().Single(n => n.Text.StartsWith("Tables (")); tables.Expand(); tables.Nodes[0].Expand(); c.Navigation.SelectedNode = tables.Nodes[0]; Save(form, path, "06-tables.png");
                    c.Navigation.SelectedNode = c.Navigation.Nodes.Cast<TreeNode>().Single(n => n.Text.StartsWith("Environment variables")); Save(form, path, "07-environment-variables.png");
                }
                using (var form = new MembershipResultsForm(LongHeaderFixture().Completed))
                {
                    form.Text += " — SYNTHETIC long header"; form.Size = new Size(900, 520); form.Show(); OpenPrimary(form); Save(form, path, "09-long-header-minimum.png");
                    form.Font = new Font("Segoe UI", 13.5F); form.PerformLayout(); ((CheckBox)Controls(form).Single(c => c.Name == "ObjectsAdvancedDiagnostics")).Checked = true;
                    Save(form, path, "10-accessibility-overflow-font150.png");
                }
            });
        }
        [TestMethod, TestCategory("H2E2C3A")]
        public void OneDefaultBrowserHidesLegacyNavigationAndRetainsExactSupportReturn()
        {
            Sta(() =>
            {
                var f = ExplorerFixture.Apps(); var spy = new FakeOrganizationService(); var csv = ExplorerFixture.Csv(f);
                using (var form = new MembershipResultsForm(f.Completed, sourceService: spy, targetService: spy))
                {
                    form.Show(); OpenPrimary(form); var c = Control(form);
                    Assert.AreEqual(1, Controls(form).OfType<TreeView>().Count(t => t.Visible));
                    Assert.IsFalse(View(form).Visible); Assert.IsFalse(Button(form).Visible);
                    Assert.IsFalse(WorkspaceButton(form, "PrimaryObjectsAction").Visible);
                    Assert.AreEqual("&Technical Membership", WorkspaceButton(form, "TechnicalMembershipAction").Text);
                    Assert.IsTrue(WorkspaceButton(form, "ObjectsBackAction").Visible);
                    Assert.IsTrue(WorkspaceButton(form, "TechnicalMembershipAction").Visible);
                    Assert.AreEqual(ObjectsWorkspaceScope.Target, c.WorkspaceScope);
                    Assert.AreEqual(6, c.Grid.Columns.Cast<DataGridViewColumn>().Count(x => x.Visible));
                    int occurrence = form.ExplorerProjection.Occurrences.Single(o => o.Original.Side == DisplayEnvironmentSide.Target && o.Original.Record.ComponentType == 80).Original.OccurrenceIndex;
                    c.SelectOccurrence(occurrence); var original = c.SelectedComparison.OriginalRowIndex;
                    WorkspaceButton(form, "TechnicalMembershipAction").PerformClick(); Application.DoEvents();
                    Assert.IsTrue(form.IsTechnicalWorkspace); Assert.AreEqual(0, View(form).SelectedIndex);
                    Assert.AreEqual(f.Presentation.Rows.Count, Grid(form).Rows.Count);
                    Assert.AreEqual(original, ((MembershipGridRow)Grid(form).CurrentRow.DataBoundItem).OriginalRowIndex);
                    OpenPrimary(form); Assert.AreEqual(original, c.SelectedComparison.OriginalRowIndex);
                    Assert.AreEqual(DisplayEnvironmentSide.Target, c.SelectedSide); Assert.AreEqual(occurrence, c.SelectedOccurrenceIndex);
                    Assert.AreEqual(csv, ExplorerFixture.Csv(f)); Assert.AreEqual(0, spy.Calls + spy.ExecuteCalls + spy.WriteCalls);
                }
            });
        }
        [TestMethod, TestCategory("H2E2C3A")]
        public void LocalSearchUsesRetainedNamesAndClearingRestoresCandidatesWithoutRequests()
        {
            Sta(() =>
            {
                var f = SyntheticNames(ExplorerFixture.Saa()); var spy = new FakeOrganizationService(); var csv = ExplorerFixture.Csv(f);
                using (var form = new MembershipResultsForm(f.Completed, sourceService: spy, targetService: spy))
                {
                    form.Show(); OpenPrimary(form); var c = Control(form); var projection = form.ExplorerProjection; var before = c.VisibleRows.ToArray();
                    var search = Controls(form).OfType<ToolStrip>().SelectMany(t => t.Items.OfType<ToolStripTextBox>()).Single(t => t.Name == "ObjectsSearchFilter");
                    search.Text = " DEMO PAIRED670 "; Assert.AreEqual(1, c.VisibleRows.Count);
                    Assert.AreEqual("Demo paired670", c.Grid.Rows[0].Cells["DisplayName"].Value);
                    search.Text = "not-present-in-retained-names"; Assert.AreEqual(0, c.VisibleRows.Count);
                    search.Text = ""; CollectionAssert.AreEqual(before, c.VisibleRows.ToArray()); Assert.AreSame(projection, form.ExplorerProjection);
                    Scope(form).SelectedIndex = 2; OpenPrimary(form); WorkspaceButton(form, "ObjectsDifferencesFilter").PerformClick(); OpenPrimary(form);
                    var expected = form.ExplorerProjection.Occurrences.Where(o => o.PrimaryCategory != null).SelectMany(o => o.Original.ComparisonRowIndices).Distinct()
                        .Where(i => f.Presentation.Rows[i].Presence != MembershipPresence.PresentInBoth || f.Presentation.Rows[i].DefinitionStatus == "Different").OrderBy(i => i).ToArray();
                    CollectionAssert.AreEqual(expected, c.VisibleRows.ToArray());
                    Assert.AreEqual(csv, ExplorerFixture.Csv(f)); Assert.AreEqual(0, spy.Calls + spy.ExecuteCalls + spy.WriteCalls);
                    Assert.AreEqual(840, projection.Rows.Count); Assert.AreEqual(1512, projection.Occurrences.Count);
                }
            });
        }
        [DataTestMethod, TestCategory("H2E2C3A"), DataRow(false), DataRow(true)]
        public void CompactDefinitionDetailsSeparateComparedPropertiesFromSupportEvidence(bool changed)
        {
            Sta(() =>
            {
                var f = DefinitionFixture(changed); using (var form = new MembershipResultsForm(f.Completed))
                {
                    form.Show(); OpenPrimary(form); Scope(form).SelectedIndex = 2; OpenPrimary(form); var c = Control(form); c.SelectComparison(0, DisplayEnvironmentSide.Target);
                    WorkspaceButton(form, "ObjectsDetailsAction").PerformClick();
                    StringAssert.Contains(c.SimpleDetailsText, changed ? "Changed compared properties:" : "Compared properties match");
                    if (changed) { StringAssert.Contains(c.SimpleDetailsText, "sequence: Source 1; Target 2"); Assert.IsFalse(c.SimpleDetailsText.Contains("unchanged_audit_field")); }
                    foreach (var token in new[] { "Original comparison row:", "Portable identity:", "SolutionComponentId:", "H1 occurrence index:", "SECRET-DIAGNOSTIC" }) Assert.IsFalse(c.SimpleDetailsText.Contains(token), token);
                    var support = (CheckBox)Controls(c).Single(x => x.Name == "ExplorerSupportEvidenceToggle"); Assert.IsFalse(support.Checked); support.Checked = true;
                    StringAssert.Contains(c.SimpleDetailsText, "Original comparison row: 0"); StringAssert.Contains(c.SimpleDetailsText, "SECRET-DIAGNOSTIC");
                    support.Checked = false; Assert.IsFalse(c.SimpleDetailsText.Contains("SECRET-DIAGNOSTIC"));
                    Assert.AreEqual(changed ? "Different" : "Match", f.Presentation.Rows[0].DefinitionStatus);
                }
            });
        }
        [DataTestMethod, TestCategory("H2E2C3A"), DataRow(DisplayAvailability.NotRetrieved), DataRow(DisplayAvailability.AttributeOmitted), DataRow(DisplayAvailability.Unavailable), DataRow(DisplayAvailability.Available)]
        public void MissingNamesRemainDashWithExplicitAvailabilityAndNoGuidFallback(DisplayAvailability availability)
        {
            Sta(() =>
            {
                var identity = ExplorerFixture.Identity(61, Guid.NewGuid().ToString()); var f = ExplorerFixture.Make(new[] { identity });
                var field = new ComponentDisplayValue(availability, null, availability == DisplayAvailability.Available ? "Synthetic retrieved null" : null);
                f.Completed = new CompletedMembershipComparison(f.Presentation, new[] { new ComponentDisplayEvidence(0,
                    ExplorerFixture.Display(identity, DisplayEnvironmentSide.Source, friendly: field, name: field), ExplorerFixture.Display(null, DisplayEnvironmentSide.Target)) }, DisplayProjectionOutcome.Completed);
                using (var form = new MembershipResultsForm(f.Completed))
                {
                    form.Show(); OpenPrimary(form); Scope(form).SelectedIndex = 1; OpenPrimary(form); var c = Control(form); c.SelectComparison(0, DisplayEnvironmentSide.Source);
                    Assert.AreEqual("—", c.Grid.Rows[0].Cells["DisplayName"].Value); Assert.AreEqual("—", c.Grid.Rows[0].Cells["Name"].Value);
                    StringAssert.Contains(c.SimpleDetailsText, availability == DisplayAvailability.Available ? "Retrieved null" : availability.ToString());
                    Assert.IsFalse(c.SimpleDetailsText.Contains(identity.ComparisonKey)); Assert.IsFalse(c.Grid.Columns["Managed"].Visible);
                    Assert.AreEqual(MembershipPresence.OnlyInSource, f.Presentation.Rows[0].Presence);
                }
            });
        }
        [TestMethod, TestCategory("H2E2C3A")]
        public void OptionalMetadataOnlyDisplaysActualSameSideValues()
        {
            Sta(() =>
            {
                var f = ExplorerFixture.Make(new[] { ExplorerFixture.Identity(61, "resource") }, new[] { ExplorerFixture.Identity(61, "resource") });
                Func<ComponentIdentity, DisplayEnvironmentSide, bool, ComponentSideDisplayEvidence> display = (id, side, actual) => new ComponentSideDisplayEvidence(side,
                    id.Record.SolutionComponentId, id.Record.ObjectId, 61, new ComponentDisplayValue(DisplayAvailability.NotRetrieved),
                    new ComponentDisplayValue(DisplayAvailability.Available, "Resource", "Synthetic metadata"), new ComponentDisplayValue(DisplayAvailability.Available, "resource", "Synthetic metadata"),
                    DisplayBackingCorrelationStatus.Unique, new ComponentDisplayValue(DisplayAvailability.NotRetrieved),
                    new ComponentObjectMetadataEvidence(side, managed: actual ? new ComponentObjectMetadataValue<bool?>(DisplayAvailability.Available, false, "Synthetic ismanaged", true) : null));
                f.Completed = new CompletedMembershipComparison(f.Presentation, new[] { new ComponentDisplayEvidence(0,
                    display(f.Presentation.Rows[0].SourceIdentity, DisplayEnvironmentSide.Source, false), display(f.Presentation.Rows[0].TargetIdentity, DisplayEnvironmentSide.Target, true)) }, DisplayProjectionOutcome.Completed);
                using (var form = new MembershipResultsForm(f.Completed))
                {
                    form.Show(); OpenPrimary(form); var c = Control(form); Assert.IsFalse(c.Grid.Columns["Managed"].Visible);
                    WorkspaceButton(form, "ObjectsMetadataAction").PerformClick();
                    Assert.IsTrue(c.Grid.Columns["Managed"].Visible); Assert.AreEqual("No", c.Grid.Rows[0].Cells["Managed"].Value);
                    Scope(form).SelectedIndex = 1; OpenPrimary(form); Assert.IsTrue(c.Grid.Columns["Managed"].Visible);
                    Assert.IsFalse(string.Equals("No", c.Grid.Rows[0].Cells["Managed"].Value as string, StringComparison.Ordinal));
                    StringAssert.Contains(c.FullCellToolTip(0, "Managed"), "NotRetrieved");
                    WorkspaceButton(form, "ObjectsMetadataAction").PerformClick();
                    Assert.IsFalse(c.Grid.Columns["Managed"].Visible); Assert.IsFalse(c.Grid.Columns["Customized"].Visible);
                    Assert.IsFalse(c.Grid.Columns["Owner"].Visible); Assert.IsFalse(c.Grid.Columns["Status"].Visible);
                }
            });
        }
        [TestMethod, TestCategory("H2E2C3A")]
        public void DistinctRepeatedOccurrencesKeepExactOriginalRowThroughSupportDrillthrough()
        {
            Sta(() =>
            {
                var identity = ExplorerFixture.Identity(61, "repeated"); var f = ExplorerFixture.Make(new[] { identity, identity });
                using (var form = new MembershipResultsForm(f.Completed))
                {
                    form.Show(); OpenPrimary(form); Scope(form).SelectedIndex = 1; OpenPrimary(form); var c = Control(form); c.SelectOccurrence(1);
                    Assert.AreEqual(1, c.SelectedOccurrenceIndex); Assert.AreEqual(2, c.VisibleRows.Count); Assert.IsNull(c.SelectedComparison);
                    c.SelectComparison(1, DisplayEnvironmentSide.Source); c.SelectOccurrence(1); c.SelectComparison(1, DisplayEnvironmentSide.Source);
                    // The explicit comparison selection is required when a raw occurrence has multiple original mappings.
                    WorkspaceButton(form, "ObjectsOriginalRowAction").PerformClick(); Assert.AreEqual(1, ((MembershipGridRow)Grid(form).CurrentRow.DataBoundItem).OriginalRowIndex);
                    OpenPrimary(form); Assert.AreEqual(1, c.SelectedComparison.OriginalRowIndex); Assert.AreEqual(DisplayEnvironmentSide.Source, c.SelectedSide);
                    Assert.AreEqual(1, c.SelectedOccurrenceIndex);
                    Assert.AreEqual(2, form.ExplorerProjection.Occurrences.Count); Assert.AreEqual(2, form.ExplorerProjection.Rows.Count);
                }
            });
        }
        [TestMethod, TestCategory("H2E2C3A")]
        public void ConsolidatedBrowserSyntheticScreenshots()
        {
            var path = Environment.GetEnvironmentVariable("H2E2C3A_EVIDENCE_DIRECTORY"); if (string.IsNullOrWhiteSpace(path)) return;
            Directory.CreateDirectory(path);
            Sta(() =>
            {
                using (var form = new MembershipResultsForm(SyntheticNames(ExplorerFixture.Saa()).Completed))
                {
                    form.Text += " — SYNTHETIC H2E-2C3A"; form.Show(); OpenPrimary(form); Save(form, path, "01-target-all.png");
                    Scope(form).SelectedIndex = 1; OpenPrimary(form); Save(form, path, "02-source-all.png");
                    Scope(form).SelectedIndex = 2; OpenPrimary(form); Save(form, path, "03-compare-all.png");
                    WorkspaceButton(form, "ObjectsDifferencesFilter").PerformClick(); OpenPrimary(form); Save(form, path, "04-compare-differences.png");
                    WorkspaceButton(form, "TechnicalMembershipAction").PerformClick(); Save(form, path, "10-support-collapsed.png");
                    WorkspaceButton(form, "AdvancedDiagnosticsAction").PerformClick(); Save(form, path, "11-support-expanded.png");
                }
                using (var form = new MembershipResultsForm(SyntheticNames(ExplorerFixture.Apps()).Completed))
                {
                    form.Text += " — SYNTHETIC verified Settings"; form.Show(); OpenPrimary(form); var c = Control(form);
                    c.SelectOccurrence(form.ExplorerProjection.Occurrences.Single(o => o.Original.Side == DisplayEnvironmentSide.Target && o.Original.Record.ComponentType == 80).Original.OccurrenceIndex);
                    c.Navigation.SelectedNode.Expand(); c.Navigation.SelectedNode.Nodes[0].Expand(); c.Navigation.SelectedNode = c.Navigation.SelectedNode.Nodes[0]; Save(form, path, "05-apps-settings.png");
                }
                var table = ExplorerFixture.Identity(1, "account"); var child = ExplorerFixture.Identity(2, "account.column", table.Record.SolutionComponentId);
                using (var form = new MembershipResultsForm(SyntheticNames(ExplorerFixture.Make(new[] { table, child, ExplorerFixture.Identity(380, "variable") }, new[] { table, child, ExplorerFixture.Identity(380, "variable") })).Completed))
                {
                    form.Text += " — SYNTHETIC table / environment variables"; form.Show(); OpenPrimary(form); var c = Control(form);
                    c.Navigation.SelectedNode = c.Navigation.Nodes.Cast<TreeNode>().Single(n => n.Text.StartsWith("Tables (")); c.Navigation.SelectedNode.Expand(); c.Navigation.SelectedNode.Nodes[0].Expand(); Save(form, path, "06-table-child.png");
                    c.Navigation.SelectedNode = c.Navigation.Nodes.Cast<TreeNode>().Single(n => n.Text.StartsWith("Environment variables")); Save(form, path, "07-environment-variables.png");
                }
                foreach (bool changed in new[] { false, true }) using (var form = new MembershipResultsForm(DefinitionFixture(changed).Completed))
                {
                    form.Text += " — SYNTHETIC compared properties"; form.Show(); OpenPrimary(form); Scope(form).SelectedIndex = 2; OpenPrimary(form); Control(form).SelectComparison(0, DisplayEnvironmentSide.Target);
                    WorkspaceButton(form, "ObjectsDetailsAction").PerformClick(); Save(form, path, changed ? "09-details-different.png" : "08-details-match.png");
                }
                using (var form = new MembershipResultsForm(LongHeaderFixture().Completed))
                {
                    form.Text += " — SYNTHETIC minimum / scaled font"; form.Size = new Size(900, 520); form.Show(); OpenPrimary(form); Save(form, path, "12-minimum-long-name.png");
                    form.Font = new Font("Segoe UI", 13.5F); form.PerformLayout(); Save(form, path, "13-font150.png");
                }
            });
        }
        [DataTestMethod, TestCategory("H2E2C3B"), DataRow(900, 100), DataRow(900, 125), DataRow(900, 150), DataRow(900, 200), DataRow(1600, 100), DataRow(1600, 200)]
        public void SingleLineCellsFitActualBoundsAndExposeFullValuesAtFontScale(int width, int percent)
        {
            Sta(() =>
            {
                using (var form = new MembershipResultsForm(PolishFixture("Different", true).Completed))
                {
                    form.Size = new Size(width, 520); form.Show(); OpenPrimary(form); Scope(form).SelectedIndex = 2; OpenPrimary(form);
                    form.Font = new Font("Segoe UI", 9F * percent / 100F); form.PerformLayout(); Application.DoEvents(); var c = Control(form); var grid = c.Grid;
                    Assert.IsTrue(grid.VirtualMode); Assert.AreEqual(DataGridViewAutoSizeRowsMode.None, grid.AutoSizeRowsMode);
                    foreach (var name in new[] { "DisplayName", "Name", "Type", "ComparisonStatus" })
                    {
                        var column = grid.Columns[name]; grid.FirstDisplayedScrollingColumnIndex = column.Index; grid.Refresh(); Application.DoEvents();
                        var cell = grid.Rows[0].Cells[column.Index]; string full = Convert.ToString(cell.FormattedValue);
                        Assert.IsFalse(full.Contains("\r")); Assert.IsFalse(full.Contains("\n")); Assert.AreEqual(DataGridViewTriState.False, cell.InheritedStyle.WrapMode);
                        var bounds = grid.GetCellDisplayRectangle(column.Index, 0, true); Assert.IsTrue(bounds.Width > 0 && bounds.Height > 0);
                        int textHeight = TextRenderer.MeasureText(full, grid.Font, new Size(int.MaxValue, int.MaxValue), TextFormatFlags.SingleLine | TextFormatFlags.NoPrefix).Height;
                        Assert.IsTrue(textHeight <= bounds.Height - 4, name + " must fit vertically within actual painted cell bounds.");
                        Assert.IsTrue(cell.GetContentBounds(0).Height <= bounds.Height); Assert.AreEqual(c.PrimaryRowHeight, grid.Rows[0].Height);
                        var tip = ActualCellTooltip(grid, column.Index, 0); StringAssert.Contains(tip, full);
                        StringAssert.Contains(grid.Rows[0].AccessibilityObject.Description, full);
                        if (name == "ComparisonStatus") { StringAssert.Contains(full, "Present in Both"); StringAssert.Contains(full, "Different"); }
                    }
                    TestContext.WriteLine("SYNTHETIC width={0}; font={1}%; constant row height={2}; visible cells only; no bulk measurement", width, percent, c.PrimaryRowHeight);
                }
            });
        }
        [DataTestMethod, TestCategory("H2E2C3B"), DataRow(1600, 100), DataRow(900, 100), DataRow(900, 200)]
        public void ChangedPropertiesPrecedeCoverageAndRemainVisibleOrExpandable(int width, int percent)
        {
            Sta(() =>
            {
                using (var form = new MembershipResultsForm(PolishFixture("Different").Completed))
                {
                    form.Size = new Size(width, width == 1600 ? 900 : 520); form.Show(); OpenPrimary(form); Scope(form).SelectedIndex = 2; OpenPrimary(form);
                    form.Font = new Font("Segoe UI", 9F * percent / 100F); Application.DoEvents(); var c = Control(form); c.SelectComparison(0, DisplayEnvironmentSide.Target);
                    // R2E10 diagnostic-only reads: no layout, navigation, formatting or event pumping.
                    int diagSequence = 0;
                    string DiagQuote(object value) => "\"" + Convert.ToString(value, System.Globalization.CultureInfo.InvariantCulture).Replace("\\", "\\\\").Replace("\"", "\\\"").Replace("\r", "\\r").Replace("\n", "\\n").Replace("\t", "\\t") + "\"";
                    string DiagReference(object value) => value == null ? "null" : System.Runtime.CompilerServices.RuntimeHelpers.GetHashCode(value).ToString(System.Globalization.CultureInfo.InvariantCulture);
                    void DiagControl(System.Collections.Generic.Dictionary<string, string> values, string key, System.Windows.Forms.Control control)
                    {
                        values[key + ".type"] = control.GetType().Name; values[key + ".name"] = control.Name;
                        values[key + ".parent"] = control.Parent == null ? "null" : control.Parent.Name + ":" + control.Parent.GetType().Name;
                        values[key + ".bounds"] = control.Bounds.ToString(); values[key + ".client"] = control.ClientRectangle.ToString();
                        values[key + ".visible"] = control.Visible.ToString(); values[key + ".dock"] = control.Dock.ToString();
                        values[key + ".margin"] = control.Margin.ToString(); values[key + ".padding"] = control.Padding.ToString();
                        values[key + ".minimum"] = control.MinimumSize.ToString(); values[key + ".maximum"] = control.MaximumSize.ToString();
                        values[key + ".autoSize"] = control.AutoSize.ToString(); values[key + ".dpi"] = control.DeviceDpi.ToString();
                        values[key + ".font"] = control.Font.Name + ";size=" + control.Font.Size.ToString(System.Globalization.CultureInfo.InvariantCulture) + ";points=" + control.Font.SizeInPoints.ToString(System.Globalization.CultureInfo.InvariantCulture) + ";height=" + control.Font.Height + ";unit=" + control.Font.Unit + ";style=" + control.Font.Style;
                    }
                    void DiagMeasure(string checkpoint, bool? testedVisible = null)
                    {
                        var diagValues = new System.Collections.Generic.Dictionary<string, string>();
                        diagValues["checkpoint"] = checkpoint; diagValues["sequence"] = (++diagSequence).ToString();
                        diagValues["fixture.width"] = width.ToString(); diagValues["fixture.fontPercent"] = percent.ToString();
                        DiagControl(diagValues, "form", form); DiagControl(diagValues, "explorer", c);
                        diagValues["form.autoScaleMode"] = form.AutoScaleMode.ToString(); diagValues["form.autoScaleDimensions"] = form.AutoScaleDimensions.ToString();
                        foreach (string diagField in new[] { "headerPanel", "objectsHeader", "navigationPanel", "objectsNotice" })
                            DiagControl(diagValues, "header." + diagField, (System.Windows.Forms.Control)typeof(MembershipResultsForm).GetField(diagField, System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic).GetValue(form));
                        foreach (var diagBar in Controls(form).OfType<ToolStrip>())
                        {
                            DiagControl(diagValues, "toolbar." + diagBar.Name, diagBar);
                            foreach (ToolStripItem diagItem in diagBar.Items)
                                diagValues["toolbar." + diagBar.Name + ".item." + diagItem.Name] = "bounds=" + diagItem.Bounds + ";visible=" + diagItem.Visible + ";available=" + diagItem.Available;
                        }
                        foreach (var diagSplit in Controls(c).OfType<SplitContainer>())
                        {
                            string diagKey = "split." + diagSplit.Name; DiagControl(diagValues, diagKey, diagSplit);
                            diagValues[diagKey + ".orientation"] = diagSplit.Orientation.ToString(); diagValues[diagKey + ".distance"] = diagSplit.SplitterDistance.ToString();
                            diagValues[diagKey + ".width"] = diagSplit.SplitterWidth.ToString(); diagValues[diagKey + ".panel1Min"] = diagSplit.Panel1MinSize.ToString(); diagValues[diagKey + ".panel2Min"] = diagSplit.Panel2MinSize.ToString();
                            diagValues[diagKey + ".panel1Collapsed"] = diagSplit.Panel1Collapsed.ToString(); diagValues[diagKey + ".panel2Collapsed"] = diagSplit.Panel2Collapsed.ToString();
                            DiagControl(diagValues, diagKey + ".panel1", diagSplit.Panel1); DiagControl(diagValues, diagKey + ".panel2", diagSplit.Panel2);
                        }
                        foreach (string diagField in new[] { "detailsOptions", "detailsScrollCue", "detailsLinks", "counts", "comparedPropertiesPanel" })
                            DiagControl(diagValues, "details." + diagField, (System.Windows.Forms.Control)typeof(SolutionObjectsExplorerControl).GetField(diagField, System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic).GetValue(c));
                        var diagText = (TextBox)Controls(c).Single(control => control.Name == "ExplorerEvidence"); DiagControl(diagValues, "text", diagText);
                        diagValues["text.wordWrap"] = diagText.WordWrap.ToString(); diagValues["text.scrollBars"] = diagText.ScrollBars.ToString(); diagValues["text.multiline"] = diagText.Multiline.ToString();
                        diagValues["text.selectionStart"] = diagText.SelectionStart.ToString(); diagValues["text.selectionLength"] = diagText.SelectionLength.ToString(); diagValues["text.length"] = diagText.TextLength.ToString(); diagValues["text.handleCreated"] = diagText.IsHandleCreated.ToString();
                        int diagFirst = diagText.Text.IndexOf("sequence: Source 1; Target 2", StringComparison.Ordinal);
                        int diagHeading = diagText.Text.IndexOf("Changed compared properties:", StringComparison.Ordinal);
                        int diagAllocatorValue = diagHeading < 0 ? -1 : diagText.Text.IndexOf(Environment.NewLine, diagHeading, StringComparison.Ordinal) + Environment.NewLine.Length;
                        int diagY = diagFirst >= 0 && diagText.IsHandleCreated ? diagText.GetPositionFromCharIndex(diagFirst).Y : -1;
                        diagValues["change.index"] = diagFirst.ToString(); diagValues["change.y"] = diagY.ToString(); diagValues["change.fontHeight"] = diagText.Font.Height.ToString();
                        diagValues["change.clientHeight"] = diagText.ClientSize.Height.ToString(); diagValues["change.bottom"] = (diagY + diagText.Font.Height).ToString();
                        diagValues["change.visible"] = (diagFirst >= 0 && diagText.IsHandleCreated && diagY + diagText.Font.Height <= diagText.ClientSize.Height).ToString(); diagValues["change.testedVisible"] = testedVisible.HasValue ? testedVisible.Value.ToString() : "not yet calculated";
                        diagValues["allocator.valueIndex"] = diagAllocatorValue.ToString(); diagValues["allocator.valueY"] = (diagAllocatorValue >= 0 && diagText.IsHandleCreated ? diagText.GetPositionFromCharIndex(diagAllocatorValue).Y : -1).ToString();
                        diagValues["grid.columnHeadersHeight"] = c.Grid.ColumnHeadersHeight.ToString(); diagValues["grid.rowTemplateHeight"] = c.Grid.RowTemplate.Height.ToString();
                        var diagView = (ComboBox)Controls(c).Single(control => control.Name == "ExplorerDetailsView"); var diagSupport = (CheckBox)Controls(c).Single(control => control.Name == "ExplorerSupportEvidenceToggle");
                        diagValues["details.modeIndex"] = diagView.SelectedIndex.ToString(); diagValues["details.mode"] = Convert.ToString(diagView.SelectedItem, System.Globalization.CultureInfo.InvariantCulture); diagValues["details.support"] = diagSupport.Checked.ToString();
                        diagValues["row.index"] = c.SelectedComparison == null ? "null" : c.SelectedComparison.OriginalRowIndex.ToString(); diagValues["row.reference"] = DiagReference(c.SelectedComparison);
                        diagValues["row.authoritativeReference"] = DiagReference(c.SelectedComparison == null ? null : c.SelectedComparison.AuthoritativeRow); diagValues["row.exactProjectedOriginal"] = ReferenceEquals(c.SelectedComparison, c.Projection.Rows[0].Original).ToString();
                        diagValues["row.side"] = Convert.ToString(c.SelectedSide, System.Globalization.CultureInfo.InvariantCulture); diagValues["row.occurrence"] = Convert.ToString(c.SelectedOccurrenceIndex, System.Globalization.CultureInfo.InvariantCulture);
                        diagValues["row.projectionReference"] = DiagReference(c.Projection); diagValues["row.ownerReference"] = DiagReference(c.Projection.Owner); diagValues["form.ownerReference"] = DiagReference(form.CompletedComparison);
                        var diagOutput = new System.Text.StringBuilder("{");
                        foreach (var diagPair in diagValues)
                        {
                            if (diagOutput.Length > 1) diagOutput.Append(",");
                            diagOutput.Append(DiagQuote(diagPair.Key)).Append(":").Append(DiagQuote(diagPair.Value));
                        }
                        Console.WriteLine("R2E10_VIEWPORT " + diagOutput.Append("}"));
                    }
                    DiagMeasure("before.existing.details.click");
                    WorkspaceButton(form, "ObjectsDetailsAction").PerformClick(); Application.DoEvents();
                    var text = (TextBox)Controls(c).Single(x => x.Name == "ExplorerEvidence"); string body = text.Text;
                    DiagMeasure("after.existing.details.pump");
                    Assert.IsTrue(body.IndexOf("Display Name:", StringComparison.Ordinal) < body.IndexOf("Definition:", StringComparison.Ordinal));
                    Assert.IsTrue(body.IndexOf("Changed compared properties:", StringComparison.Ordinal) < body.IndexOf("Availability / coverage:", StringComparison.Ordinal));
                    int first = body.IndexOf("sequence: Source 1; Target 2", StringComparison.Ordinal); Assert.IsTrue(first >= 0);
                    bool visible = text.GetPositionFromCharIndex(first).Y + text.Font.Height <= text.ClientSize.Height;
                    var cue = Controls(c).Single(x => x.Name == "ExplorerDetailsScrollCue"); var expand = (LinkLabel)Controls(c).Single(x => x.Name == "ExplorerExpandDetails");
                    DiagMeasure("before.cue.visibility.assertion", visible);
                    Assert.IsTrue(visible || cue.Visible, "Constrained Details must expose a conspicuous cue rather than lose the first change.");
                    DiagMeasure("immediately.before.normal.visibility.assertion", visible);
                    if (width == 1600 || percent == 100) Assert.IsTrue(visible, "First changed value should fit at normal available size.");
                    Assert.IsTrue(expand.TabStop); Assert.IsFalse(string.IsNullOrWhiteSpace(expand.AccessibleName));
                    // Deliberately constrain the retained Details viewport without changing data.
                    var split = (SplitContainer)Controls(c).Single(x => x.Name == "ExplorerEvidenceSplit");
                    split.SplitterDistance = split.Height - split.Panel2MinSize - split.SplitterWidth; Application.DoEvents();
                    DiagMeasure("after.existing.constrained.pump");
                    Assert.IsTrue(cue.Visible); text.Focus(); Assert.IsTrue(text.Focused);
                    text.Select(text.TextLength, 0); text.ScrollToCaret(); Assert.AreEqual(body, text.Text);
                    DiagMeasure("after.existing.scroll.end");
                    text.Select(0, 0); text.ScrollToCaret();
                    DiagMeasure("after.existing.scroll.start");
                    typeof(LinkLabel).GetMethod("OnLinkClicked", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic)
                        .Invoke(expand, new object[] { new LinkLabelLinkClickedEventArgs(expand.Links[0]) }); Application.DoEvents();
                    DiagMeasure("after.existing.expand.pump");
                    Assert.IsTrue(split.Panel1Collapsed); Assert.IsFalse(cue.Visible); Assert.AreEqual(body, text.Text);
                    DiagMeasure("immediately.before.expanded.visibility.assertion");
                    Assert.IsTrue(text.GetPositionFromCharIndex(first).Y + text.Font.Height <= text.ClientSize.Height);
                    typeof(LinkLabel).GetMethod("OnLinkClicked", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic)
                        .Invoke(expand, new object[] { new LinkLabelLinkClickedEventArgs(expand.Links[0]) }); Assert.IsFalse(split.Panel1Collapsed);
                    DiagMeasure("after.existing.restore");
                    var support = (CheckBox)Controls(c).Single(x => x.Name == "ExplorerSupportEvidenceToggle"); string evidence = c.EvidenceText;
                    support.Checked = true; Assert.AreEqual(evidence, c.SimpleDetailsText); support.Checked = false; Assert.AreEqual(body, c.SimpleDetailsText);
                    DiagMeasure("after.existing.support.return");
                }
            });
        }
        [DataTestMethod, TestCategory("H2E2C3B"), DataRow("one-side"), DataRow("mixed"), DataRow("unverified-change"), DataRow("Match"), DataRow("Unsupported")]
        public void DetailsValuesPreserveUnavailableBlankNullAndUncertainComparisons(string scenario)
        {
            Sta(() =>
            {
                var properties = scenario == "one-side" ? new[] { new ComponentPropertyPresentation("field", null, "available target", true, "Different") } :
                    scenario == "mixed" ? new[] { new ComponentPropertyPresentation("omitted", "[AttributeOmitted]", "[NotRetrieved]", true, "Different"), new ComponentPropertyPresentation("blank", "", "[Retrieved null]", true, "Different") } :
                    scenario == "unverified-change" ? new[] { new ComponentPropertyPresentation("field", null, null, false, "Not compared") } : null;
                string status = scenario == "Match" ? "Match" : scenario == "Unsupported" ? "Unsupported" : "Different";
                var f = PolishFixture(status, properties: properties, unsupported: scenario == "Unsupported"); var csv = ExplorerFixture.Csv(f); int selectedRow = 0;
                using (var form = new MembershipResultsForm(f.Completed))
                {
                    form.Show(); OpenPrimary(form); var c = Control(form);
                    if (scenario == "Unsupported") { selectedRow = form.ExplorerProjection.Rows.Single(row => row.Original.AuthoritativeRow.SourceIdentity != null).OriginalRowIndex; c.SelectComparison(selectedRow, DisplayEnvironmentSide.Source); }
                    else { Scope(form).SelectedIndex = 2; OpenPrimary(form); c.SelectComparison(0, DisplayEnvironmentSide.Target); }
                    WorkspaceButton(form, "ObjectsDetailsAction").PerformClick(); string body = c.SimpleDetailsText;
                    if (scenario == "one-side") { StringAssert.Contains(body, "Source [No value supplied; availability not established]"); StringAssert.Contains(body, "Target available target"); }
                    if (scenario == "mixed") { StringAssert.Contains(body, "[AttributeOmitted]"); StringAssert.Contains(body, "[NotRetrieved]"); StringAssert.Contains(body, "[Blank value supplied]"); StringAssert.Contains(body, "[Retrieved null]"); }
                    if (scenario == "unverified-change" || scenario == "Match" || scenario == "Unsupported") Assert.IsFalse(body.Contains("Changed compared properties:"));
                    if (scenario == "Match") StringAssert.Contains(body, "Match (compared properties)");
                    if (scenario == "Unsupported") { StringAssert.Contains(body, "Unsupported"); StringAssert.Contains(body, "Indeterminate"); Assert.AreEqual(IdentityResolutionStatus.Unsupported, f.Presentation.Rows[selectedRow].SourceIdentity.Status); }
                    Assert.AreEqual(csv, ExplorerFixture.Csv(f)); Assert.AreEqual(status, f.Presentation.Rows[0].DefinitionStatus);
                }
            });
        }
        [TestMethod, TestCategory("H2E2C3B")]
        public void FontChangesScrollingAndSupportReturnRemainVirtualAndLosslessForTenThousandRows()
        {
            Sta(() =>
            {
                var identities = Enumerable.Range(0, 10000).Select(i => ExplorerFixture.Identity(61, "resource" + i)).ToArray();
                var fixture = ExplorerFixture.Make(identities, identities);
                var f = SyntheticNames(fixture);
                var csv = ExplorerFixture.Csv(f); var rows = f.Presentation.Rows.ToArray(); var spy = new FakeOrganizationService(); int captures = 0;
                using (var form = new MembershipResultsForm(f.Completed, sourceService: spy, targetService: spy, captureAppSettingEvidence: () => captures++))
                {
                    form.Show();
                    OpenPrimary(form);
                    var c = Control(form); var watch = Stopwatch.StartNew();
                    form.Font = new Font("Segoe UI", 18F); Application.DoEvents(); c.Grid.FirstDisplayedScrollingRowIndex = 9000; c.Grid.Refresh(); Application.DoEvents();
                    Assert.IsTrue(watch.ElapsedMilliseconds < 5000, "Constant metrics and visible-cell painting must remain bounded.");
                    Assert.IsTrue(c.Grid.VirtualMode); Assert.AreEqual(10000, c.Grid.RowCount); Assert.AreEqual(DataGridViewAutoSizeRowsMode.None, c.Grid.AutoSizeRowsMode);
                    c.Navigation.Nodes[1].Expand(); Assert.IsTrue(c.Navigation.Nodes[1].Nodes.Count <= 50); Assert.AreEqual(20000, form.ExplorerProjection.Occurrences.Count);
                    int occurrence = form.ExplorerProjection.Occurrences.Single(o => o.Original.Side == DisplayEnvironmentSide.Target && o.Original.ComparisonRowIndices.Contains(9000)).Original.OccurrenceIndex;
                    c.SelectOccurrence(occurrence); int original = c.SelectedComparison.OriginalRowIndex; var evidence = c.EvidenceText;
                    Scope(form).SelectedIndex = 1; OpenPrimary(form);
                    Scope(form).SelectedIndex = 0; OpenPrimary(form);
                    Scope(form).SelectedIndex = 2; OpenPrimary(form);
                    Assert.AreEqual(original, c.SelectedComparison.OriginalRowIndex); Assert.AreEqual(DisplayEnvironmentSide.Target, c.SelectedSide);
                    var search = Controls(form).OfType<ToolStrip>().SelectMany(t => t.Items.OfType<ToolStripTextBox>()).Single(t => t.Name == "ObjectsSearchFilter");
                    search.Text = "does not occur";
                    search.Text = "";
                    c.SelectOccurrence(occurrence);
                    WorkspaceButton(form, "ObjectsOriginalRowAction").PerformClick(); Assert.AreEqual(original, ((MembershipGridRow)Grid(form).CurrentRow.DataBoundItem).OriginalRowIndex);
                    OpenPrimary(form); Assert.AreEqual(original, c.SelectedComparison.OriginalRowIndex); Assert.AreEqual(occurrence, c.SelectedOccurrenceIndex); Assert.AreEqual(DisplayEnvironmentSide.Target, c.SelectedSide);
                    Assert.AreEqual(evidence, c.EvidenceText); Assert.AreEqual(csv, ExplorerFixture.Csv(f)); CollectionAssert.AreEqual(rows, f.Presentation.Rows.ToArray());
                    Assert.AreEqual(0, spy.Calls + spy.ExecuteCalls + spy.WriteCalls + captures);
                    TestContext.WriteLine("SYNTHETIC 10000-row font/scroll step={0}ms; virtual grid; no per-row measuring", watch.ElapsedMilliseconds);
                }
            });
        }
        [TestMethod, TestCategory("H2E2C3B")]
        public void ReadableGridAndDifferencesFirstDetailsSyntheticScreenshots()
        {
            var path = Environment.GetEnvironmentVariable("H2E2C3B_EVIDENCE_DIRECTORY"); if (string.IsNullOrWhiteSpace(path)) return; Directory.CreateDirectory(path);
            Sta(() =>
            {
                foreach (int percent in new[] { 100, 125, 150, 200 }) using (var form = new MembershipResultsForm(PolishFixture("Different", true).Completed))
                {
                    form.Text += " — SYNTHETIC single-line grid"; form.Size = new Size(900, 520); form.Show(); OpenPrimary(form); Scope(form).SelectedIndex = 2; OpenPrimary(form);
                    form.Font = new Font("Segoe UI", 9F * percent / 100F); Application.DoEvents(); Save(form, path, "grid-minimum-font" + percent + ".png");
                }
                foreach (int percent in new[] { 100, 200 }) using (var form = new MembershipResultsForm(PolishFixture("Different").Completed))
                {
                    form.Text += " — SYNTHETIC changed properties first"; form.Size = new Size(percent == 100 ? 1600 : 900, percent == 100 ? 900 : 520); form.Show(); OpenPrimary(form); Scope(form).SelectedIndex = 2; OpenPrimary(form);
                    form.Font = new Font("Segoe UI", 9F * percent / 100F); Application.DoEvents(); Control(form).SelectComparison(0, DisplayEnvironmentSide.Target); WorkspaceButton(form, "ObjectsDetailsAction").PerformClick();
                    Save(form, path, "details-different-font" + percent + ".png");
                    if (percent == 200) { var link = (LinkLabel)Controls(Control(form)).Single(x => x.Name == "ExplorerExpandDetails"); typeof(LinkLabel).GetMethod("OnLinkClicked", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic).Invoke(link, new object[] { new LinkLabelLinkClickedEventArgs(link.Links[0]) }); Save(form, path, "details-expanded-font200.png"); }
                }
                using (var form = new MembershipResultsForm(PolishFixture("Match").Completed))
                { form.Text += " — SYNTHETIC compared-property match"; form.Show(); OpenPrimary(form); Control(form).SelectComparison(0); WorkspaceButton(form, "ObjectsDetailsAction").PerformClick(); Save(form, path, "details-match.png"); }
            });
        }
        private static string ActualCellTooltip(DataGridView grid, int column, int row)
        {
            var args = (DataGridViewCellToolTipTextNeededEventArgs)Activator.CreateInstance(typeof(DataGridViewCellToolTipTextNeededEventArgs),
                System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic, null, new object[] { column, row, "" }, System.Globalization.CultureInfo.InvariantCulture);
            typeof(DataGridView).GetMethod("OnCellToolTipTextNeeded", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic,
                null, new[] { typeof(DataGridViewCellToolTipTextNeededEventArgs) }, null).Invoke(grid, new object[] { args });
            return args.ToolTipText;
        }
        [TestMethod, TestCategory("D4T8R2C")]
        public void R2COverviewPropertiesAndSupportRetainIndependentSidesAndOriginalRow()
        {
            Sta(() =>
            {
                var f = DefinitionFixture(true); var owner = f.Presentation.Rows[0];
                f.Completed = new CompletedMembershipComparison(f.Presentation, new[] { new ComponentDisplayEvidence(0,
                    ExplorerFixture.Display(owner.SourceIdentity, DisplayEnvironmentSide.Source,
                        friendly: new ComponentDisplayValue(DisplayAvailability.Available, "Source title", "Synthetic Source title"),
                        name: new ComponentDisplayValue(DisplayAvailability.Available, "source_logical", "Synthetic Source name")),
                    ExplorerFixture.Display(owner.TargetIdentity, DisplayEnvironmentSide.Target,
                        friendly: new ComponentDisplayValue(DisplayAvailability.Available, "Target title", "Synthetic Target title"),
                        name: new ComponentDisplayValue(DisplayAvailability.Available, "target_logical", "Synthetic Target name"))) }, DisplayProjectionOutcome.Completed);
                var spy = new FakeOrganizationService(); int captures = 0; var csv = ExplorerFixture.Csv(f);
                using (var form = new MembershipResultsForm(f.Completed, sourceService: spy, targetService: spy, captureAppSettingEvidence: () => captures++))
                {
                    form.Size = new Size(900, 520); form.Show(); OpenPrimary(form); var c = Control(form); var projection = form.ExplorerProjection;
                    var split = (SplitContainer)Controls(c).Single(x => x.Name == "ExplorerEvidenceSplit");
                    var view = (ComboBox)Controls(c).Single(x => x.Name == "ExplorerDetailsView");
                    var properties = (DataGridView)Controls(c).Single(x => x.Name == "ExplorerComparedProperties");
                    var text = (TextBox)Controls(c).Single(x => x.Name == "ExplorerEvidence");
                    var hiddenContent = text.Text; c.SelectComparison(0, DisplayEnvironmentSide.Target);
                    Assert.AreEqual(hiddenContent, text.Text, "Closed Details must defer the selected Overview until explicit disclosure.");
                    Assert.IsTrue(split.Panel2Collapsed); Assert.IsFalse(text.Focused); Assert.IsFalse(view.Focused);
                    Assert.AreEqual(0, properties.RowCount, "Selecting a row must not prepare the properties view.");
                    WorkspaceButton(form, "ObjectsDetailsAction").PerformClick(); Application.DoEvents();
                    Assert.AreEqual(0, view.SelectedIndex); Assert.IsTrue(text.Visible); Assert.IsFalse(properties.Visible);
                    StringAssert.Contains(c.SimpleDetailsText, "Display Name: Target title");
                    StringAssert.Contains(c.SimpleDetailsText, "Source display name: Source title");
                    StringAssert.Contains(c.SimpleDetailsText, "Target display name: Target title");
                    StringAssert.Contains(c.SimpleDetailsText, "Source internal name: source_logical");
                    StringAssert.Contains(c.SimpleDetailsText, "Target internal name: target_logical");
                    StringAssert.Contains(c.SimpleDetailsText, "Comparison row index: 0");
                    Assert.IsTrue(c.SimpleDetailsText.IndexOf("sequence: Source 1; Target 2", StringComparison.Ordinal) <
                        c.SimpleDetailsText.IndexOf("Availability / coverage:", StringComparison.Ordinal));
                    Scope(form).SelectedIndex = 1; OpenPrimary(form); c.SelectComparison(0, DisplayEnvironmentSide.Source);
                    Assert.AreEqual("Source title", c.Grid.Rows[0].Cells["DisplayName"].Value);
                    Scope(form).SelectedIndex = 2; OpenPrimary(form); c.SelectComparison(0, DisplayEnvironmentSide.Target);
                    var overview = c.SimpleDetailsText; var evidence = c.EvidenceText;
                    view.SelectedIndex = 1; Application.DoEvents();
                    Assert.IsTrue(properties.Visible); Assert.IsFalse(text.Visible); Assert.AreEqual(2, properties.RowCount);
                    Assert.IsTrue(properties.VirtualMode); Assert.AreEqual(DataGridViewAutoSizeRowsMode.None, properties.AutoSizeRowsMode);
                    Assert.IsTrue(properties.ClientSize.Height >= properties.ColumnHeadersHeight + properties.RowTemplate.Height + SystemInformation.HorizontalScrollBarHeight, "The initial properties table must fit a full row and its scrollbar.");
                    Assert.AreSame(owner.DefinitionDetail, properties.Tag);
                    Assert.AreEqual("sequence", properties.Rows[0].Cells[0].Value);
                    Assert.AreEqual("1", properties.Rows[0].Cells[1].Value); Assert.AreEqual("2", properties.Rows[0].Cells[2].Value);
                    Assert.AreEqual("Different", properties.Rows[0].Cells[3].Value);
                    Assert.AreEqual("unchanged_audit_field", properties.Rows[1].Cells[0].Value);
                    Assert.AreEqual("same", properties.Rows[1].Cells[1].Value); Assert.AreEqual("Match", properties.Rows[1].Cells[3].Value);
                    var support = (CheckBox)Controls(c).Single(x => x.Name == "ExplorerSupportEvidenceToggle"); support.Checked = true;
                    Assert.IsFalse(view.Enabled); Assert.IsFalse(properties.Visible); Assert.AreEqual(evidence, c.SimpleDetailsText);
                    support.Checked = false; Assert.IsTrue(view.Enabled); Assert.IsTrue(properties.Visible);
                    view.SelectedIndex = 0; Assert.AreEqual(overview, c.SimpleDetailsText);
                    WorkspaceButton(form, "ObjectsOriginalRowAction").PerformClick();
                    Assert.AreEqual(0, ((MembershipGridRow)Grid(form).CurrentRow.DataBoundItem).OriginalRowIndex);
                    OpenPrimary(form); Assert.AreSame(owner, c.SelectedComparison.AuthoritativeRow);
                    Assert.AreSame(projection, form.ExplorerProjection); Assert.AreEqual(evidence, c.EvidenceText);
                    Assert.AreEqual(csv, ExplorerFixture.Csv(f)); Assert.AreEqual(0, captures + spy.Calls + spy.ExecuteCalls + spy.WriteCalls);
                }
            });
        }
        [DataTestMethod, TestCategory("D4T8R2C"), DataRow("Not Compared"), DataRow("Unsupported"), DataRow("Unavailable")]
        public void R2CMissingDefinitionPropertiesExplainAvailabilityWithoutInventingEquality(string status)
        {
            Sta(() =>
            {
                var f = ExplorerFixture.Make(new[] { ExplorerFixture.Identity(61, "resource") });
                var rows = f.Presentation.Rows.Select(r => r.WithDefinition(status, "", null)).ToArray();
                f.Presentation = new MembershipComparisonPresentation("SAA", f.Presentation.Source, f.Presentation.Target, rows);
                f.Completed = new CompletedMembershipComparison(f.Presentation, f.Completed.DisplayEvidence, DisplayProjectionOutcome.Completed);
                var spy = new FakeOrganizationService(); int captures = 0; var csv = ExplorerFixture.Csv(f);
                using (var form = new MembershipResultsForm(f.Completed, sourceService: spy, targetService: spy, captureAppSettingEvidence: () => captures++))
                {
                    form.Show(); OpenPrimary(form); Scope(form).SelectedIndex = 1; OpenPrimary(form); var c = Control(form); c.SelectComparison(0);
                    WorkspaceButton(form, "ObjectsDetailsAction").PerformClick();
                    StringAssert.Contains(c.SimpleDetailsText, "Target information:"); StringAssert.Contains(c.SimpleDetailsText, "[No local observation]; environment absence is not established.");
                    ((ComboBox)Controls(c).Single(x => x.Name == "ExplorerDetailsView")).SelectedIndex = 1;
                    var properties = (DataGridView)Controls(c).Single(x => x.Name == "ExplorerComparedProperties");
                    Assert.AreEqual(0, properties.RowCount); Assert.IsNull(properties.Tag);
                    var context = Controls(c).Single(x => x.Name == "ExplorerComparedPropertiesContext").Text;
                    StringAssert.Contains(context, "not retained / not compared"); StringAssert.Contains(context, "Definition: " + status);
                    Assert.IsFalse(context.Contains("Objects match")); Assert.AreSame(rows[0], c.SelectedComparison.AuthoritativeRow);
                    Assert.AreEqual(csv, ExplorerFixture.Csv(f)); Assert.AreEqual(0, captures + spy.Calls + spy.ExecuteCalls + spy.WriteCalls);
                }
            });
        }
        [TestMethod, TestCategory("D4T8R2C")]
        public void R2CPropertiesShowChangesFirstKeepAllValuesAndDoNotMutateTheRetainedOrder()
        {
            Sta(() =>
            {
                string full = new string('x', 3000) + "\\\r\n\tEND";
                var fields = new[] { new ComponentPropertyPresentation("null", null, "[NotRetrieved]", false, "Not Compared"),
                    new ComponentPropertyPresentation("blank", "", "[Retrieved null]", false, "Unresolved"),
                    new ComponentPropertyPresentation("changed", "old", "new", true, "Different"),
                    new ComponentPropertyPresentation("long", full, full, false, "Match") };
                var f = PolishFixture("Different", properties: fields); var original = f.Presentation.Rows[0]; var spy = new FakeOrganizationService();
                using (var form = new MembershipResultsForm(f.Completed, sourceService: spy, targetService: spy))
                {
                    form.Show(); OpenPrimary(form); var c = Control(form); c.SelectComparison(0); WorkspaceButton(form, "ObjectsDetailsAction").PerformClick();
                    var properties = (DataGridView)Controls(c).Single(x => x.Name == "ExplorerComparedProperties");
                    Assert.AreEqual(0, properties.RowCount); ((ComboBox)Controls(c).Single(x => x.Name == "ExplorerDetailsView")).SelectedIndex = 1;
                    CollectionAssert.AreEqual(new[] { "changed", "null", "blank", "long" }, Enumerable.Range(0, properties.RowCount).Select(i => (string)properties.Rows[i].Cells[0].Value).ToArray());
                    Assert.AreEqual("[No value supplied; availability not established]", properties.Rows[1].Cells[1].Value);
                    Assert.AreEqual("[NotRetrieved]", properties.Rows[1].Cells[2].Value);
                    Assert.AreEqual("[Blank value supplied]", properties.Rows[2].Cells[1].Value); Assert.AreEqual("[Retrieved null]", properties.Rows[2].Cells[2].Value);
                    Assert.AreEqual(full, properties.Rows[3].Cells[1].Value); StringAssert.Contains(ActualCellTooltip(properties, 1, 3), full);
                    CollectionAssert.AreEqual(fields, original.DefinitionDetail.Properties.ToArray());
                    Assert.AreSame(original.DefinitionDetail, properties.Tag); Assert.AreEqual(ScrollBars.Both, properties.ScrollBars);
                    var row = c.SelectedComparison; var evidence = c.EvidenceText;
                    form.Font = new Font("Segoe UI", 12F); Application.DoEvents();
                    Assert.AreEqual(DataGridViewAutoSizeRowsMode.None, properties.AutoSizeRowsMode); Assert.IsTrue(properties.RowTemplate.Height >= properties.Font.Height + 10);
                    Assert.AreSame(row, c.SelectedComparison); Assert.AreEqual(evidence, c.EvidenceText);
                    Assert.AreEqual(0, spy.Calls + spy.ExecuteCalls + spy.WriteCalls);
                }
            });
        }


        private static ExplorerFixture PolishFixture(string status, bool longNames = false, ComponentPropertyPresentation[] properties = null, bool unsupported = false)
        {
            var left = ExplorerFixture.Identity(unsupported ? 36 : 61, unsupported ? null : "synthetic_resource", status: unsupported ? IdentityResolutionStatus.Unsupported : IdentityResolutionStatus.Resolved);
            var right = ExplorerFixture.Identity(unsupported ? 36 : 61, unsupported ? null : "synthetic_resource", status: unsupported ? IdentityResolutionStatus.Unsupported : IdentityResolutionStatus.Resolved);
            var f = ExplorerFixture.Make(new[] { left }, new[] { right });
            Func<ComponentIdentity, DisplayEnvironmentSide, ComponentSideDisplayEvidence> display = (id, side) => new ComponentSideDisplayEvidence(side, id.Record.SolutionComponentId, id.Record.ObjectId, id.Record.ComponentType,
                new ComponentDisplayValue(DisplayAvailability.Available, longNames ? "Synthetic extended component category " + string.Join(" ", Enumerable.Repeat("kind", 40)) : "Web resource", "Synthetic category only"),
                new ComponentDisplayValue(DisplayAvailability.Available, longNames ? side + " very long synthetic display name " + string.Join(" ", Enumerable.Repeat("value", 45)) : "Synthetic Resource", "Synthetic friendly name only"),
                new ComponentDisplayValue(DisplayAvailability.Available, longNames ? "synthetic_resource_" + new string('x', 220) : "synthetic_resource", "Synthetic internal name only"),
                DisplayBackingCorrelationStatus.Unique, new ComponentDisplayValue(DisplayAvailability.NotRetrieved));
            var rows = f.Presentation.Rows.Select(row => row.WithDefinition(status, status == "Different" ? "sequence" : "", new ComponentDefinitionDetailPresentation("Web resource", "synthetic_key", row.MembershipStatus, status,
                properties ?? new[] { new ComponentPropertyPresentation("sequence", "1", status == "Different" ? "2" : "1", status == "Different", status == "Different" ? "Different" : "Match") }, "Synthetic benign comparison evidence"))).ToArray();
            f.Presentation = new MembershipComparisonPresentation("SAA", f.Presentation.Source, f.Presentation.Target, rows);
            f.Completed = new CompletedMembershipComparison(f.Presentation, rows.Select((row, i) => new ComponentDisplayEvidence(i, row.SourceIdentity == null ? ExplorerFixture.Display(null, DisplayEnvironmentSide.Source) : display(row.SourceIdentity, DisplayEnvironmentSide.Source),
                row.TargetIdentity == null ? ExplorerFixture.Display(null, DisplayEnvironmentSide.Target) : display(row.TargetIdentity, DisplayEnvironmentSide.Target))), DisplayProjectionOutcome.Completed); return f;
        }

        [TestMethod, TestCategory("D4T8R2D")]
        public void R2DClosedPrimarySelectionDoesNotFormatTechnicalEvidence()
        {
            Sta(() =>
            {
                var f = ExplorerFixture.Make(new[] { ExplorerFixture.Identity(61, "a"), ExplorerFixture.Identity(61, "b") }); int calls = 0;
                using (var host = R2DHost(f, r => { calls++; return "retained row " + r.OriginalRowIndex; }, out var c))
                {
                    c.SelectComparison(1); c.Font = new Font("Segoe UI", 12F); host.PerformLayout(); Application.DoEvents();
                    Assert.AreEqual(0, calls); c.ShowDetails(true); var overview = c.SimpleDetailsText;
                    ((ComboBox)Controls(c).Single(x => x.Name == "ExplorerDetailsView")).SelectedIndex = 1;
                    Assert.AreEqual(0, calls); Assert.IsTrue(overview.Contains("Membership:"));
                    Assert.AreSame(f.Presentation.Rows[1], c.SelectedComparison.AuthoritativeRow);
                }
            });
        }
        [TestMethod, TestCategory("D4T8R2D")]
        public void R2DSupportDisclosureAndEvidenceGetterShareSelectedCache()
        {
            Sta(() =>
            {
                int calls = 0; var f = ExplorerFixture.Make(new[] { ExplorerFixture.Identity(61, "a") });
                using (var host = R2DHost(f, r => { calls++; return "EXACT retained\\value\r\n\t"; }, out var c))
                {
                    var split = (SplitContainer)Controls(c).Single(x => x.Name == "ExplorerEvidenceSplit");
                    Assert.IsTrue(split.Panel2Collapsed); string evidence = c.EvidenceText; Assert.AreEqual(1, calls);
                    Assert.AreEqual(evidence, c.EvidenceText); Assert.IsTrue(split.Panel2Collapsed); Assert.AreEqual(1, calls);
                    c.ShowDetails(true); string overview = c.SimpleDetailsText;
                    var support = (CheckBox)Controls(c).Single(x => x.Name == "ExplorerSupportEvidenceToggle");
                    support.Checked = true; Assert.AreEqual(evidence, c.SimpleDetailsText); Assert.AreEqual(1, calls);
                    support.Checked = false; Assert.AreEqual(overview, c.SimpleDetailsText); Assert.AreEqual(1, calls);
                }
            });
        }
        [TestMethod, TestCategory("D4T8R2D")]
        public void R2DCacheInvalidatesForRowSideOccurrenceAndProjection()
        {
            Sta(() =>
            {
                var ids = new[] { ExplorerFixture.Identity(61, "a"), ExplorerFixture.Identity(61, "b") };
                var f = ExplorerFixture.Make(ids, new[] { ExplorerFixture.Identity(61, "a"), ExplorerFixture.Identity(61, "b") }); int calls = 0;
                using (var host = R2DHost(f, r => { calls++; return r.FriendlyNameText; }, out var c))
                {
                    var original = f.Presentation.Rows[0]; var firstProjection = c.Projection;
                    var first = c.EvidenceText; Assert.AreEqual(1, calls); Assert.AreEqual(first, c.EvidenceText); Assert.AreEqual(1, calls);
                    c.SelectComparison(0, DisplayEnvironmentSide.Target); var otherSide = c.EvidenceText; Assert.AreEqual(2, calls);
                    int source = c.Projection.Rows[0].SourceOccurrences[0], target = c.Projection.Rows[0].TargetOccurrences[0];
                    c.SelectOccurrence(source); var sourceText = c.EvidenceText; Assert.AreEqual(3, calls); StringAssert.Contains(sourceText, "Source:");
                    c.SelectOccurrence(target); var targetText = c.EvidenceText; Assert.AreEqual(4, calls); StringAssert.Contains(targetText, "Target:");
                    c.SelectComparison(1); var otherRow = c.EvidenceText; Assert.AreEqual(5, calls);
                    f.Completed = new CompletedMembershipComparison(f.Presentation, f.Presentation.Rows.Select((r, i) => new ComponentDisplayEvidence(i,
                        ExplorerFixture.Display(r.SourceIdentity, DisplayEnvironmentSide.Source, friendly: new ComponentDisplayValue(DisplayAvailability.Available, "replacement Source", "Retained replacement")),
                        ExplorerFixture.Display(r.TargetIdentity, DisplayEnvironmentSide.Target, friendly: new ComponentDisplayValue(DisplayAvailability.Available, "replacement Target", "Retained replacement")))), DisplayProjectionOutcome.Completed);
                    var replacement = ExplorerFixture.Project(f); Assert.AreSame(firstProjection.Owner, replacement.Owner);
                    c.Bind(replacement, Enumerable.Range(0, replacement.Rows.Count), true); c.SelectComparison(0, DisplayEnvironmentSide.Source);
                    Assert.AreEqual("", R2DField<string>(c, "supportEvidence")); Assert.AreEqual(5, calls);
                    StringAssert.Contains(c.EvidenceText, "replacement Source"); Assert.AreEqual(6, calls);
                    Assert.AreSame(original, c.SelectedComparison.AuthoritativeRow);
                }
            });
        }
        [TestMethod, TestCategory("D4T8R2D")]
        public void R2DDisclosureFollowsSelectionScopeAndHiddenWorkspace()
        {
            Sta(() =>
            {
                int calls = 0; bool fail = false;
                var f = ExplorerFixture.Make(new[] { ExplorerFixture.Identity(1, "table"), ExplorerFixture.Identity(61, "resource") });
                using (var host = R2DHost(f, r => { calls++; if (fail) throw new InvalidOperationException("Synthetic formatting failure"); return "row " + r.OriginalRowIndex; }, out var c))
                {
                    var nodes = c.Navigation.Nodes.Cast<TreeNode>().ToArray(); Assert.IsTrue(nodes.Length >= 3);
                    c.Navigation.SelectedNode = nodes[1]; long revision = R2DField<long>(c, "technicalContextRevision"); var category = c.EvidenceText;
                    c.Navigation.SelectedNode = nodes[2]; Assert.IsTrue(R2DField<long>(c, "technicalContextRevision") > revision);
                    Assert.AreEqual("", R2DField<string>(c, "supportEvidence"));
                    Assert.AreEqual(category, c.EvidenceText, "Equal category captions still have independently replaced contexts."); Assert.AreEqual(0, calls);
                    c.ApplyFilter(new int[0], false); Assert.IsNull(c.SelectedComparison); Assert.AreEqual("", c.EvidenceText);
                    c.ApplyFilter(Enumerable.Range(0, c.Projection.Rows.Count), true); c.SelectComparison(0);
                    c.Visible = false; c.ConfigureWorkspace(false, ObjectsWorkspaceScope.Compare); c.SelectComparison(0);
                    c.ShowDetails(true);
                    Assert.AreEqual(0, calls, "Hidden ShowDetails(true) must retain the technical-disclosure visibility gate.");
                    Assert.AreEqual("", R2DField<string>(c, "supportEvidence"));
                    Assert.AreEqual("", ((TextBox)Controls(c).Single(x => x.Name == "ExplorerEvidence")).Text);
                    Assert.AreEqual(0, calls); c.Visible = true; Application.DoEvents(); Assert.AreEqual(1, calls);
                    string disclosed = c.EvidenceText;
                    Assert.AreEqual(disclosed, c.EvidenceText); Assert.AreEqual(disclosed, c.SimpleDetailsText);
                    Assert.AreEqual(1, calls, "Visible disclosure and repeated explicit reads must reuse the same cache entry.");
                    c.ConfigureWorkspace(true, ObjectsWorkspaceScope.Source); c.ApplyFilter(Enumerable.Range(0, c.Projection.Rows.Count), true);
                    c.SelectComparison(0, DisplayEnvironmentSide.Source); Assert.AreEqual(1, calls);
                    fail = true; Assert.ThrowsException<InvalidOperationException>(() => { var ignored = c.EvidenceText; });
                    Assert.AreEqual("", R2DField<string>(c, "supportEvidence")); Assert.AreEqual("", ((TextBox)Controls(c).Single(x => x.Name == "ExplorerEvidence")).Text);
                    int failedCalls = calls; Assert.AreEqual("", c.EvidenceText); Assert.AreEqual(failedCalls, calls, "A rejected context is not retried automatically.");
                }
            });
        }
        [TestMethod, TestCategory("D4T8R2D")]
        public void R2DExactRowTechnicalTextPreservesAvailabilityAndDefinitions()
        {
            Sta(() =>
            {
                foreach (int type in new[] { 31, 36, 300 })
                {
                    var f = ExplorerFixture.Make(new[] { ExplorerFixture.Identity(type, null, status: IdentityResolutionStatus.Unsupported) },
                        new[] { ExplorerFixture.Identity(type, null, status: IdentityResolutionStatus.Unsupported) });
                    var originalRows = f.Presentation.Rows.ToArray(); var csv = ExplorerFixture.Csv(f);
                    f.Completed = new CompletedMembershipComparison(f.Presentation, originalRows.Select((r, i) => new ComponentDisplayEvidence(i,
                        ExplorerFixture.Display(r.SourceIdentity, DisplayEnvironmentSide.Source,
                            friendly: new ComponentDisplayValue(DisplayAvailability.Available, "Same title\\exact\r\n\t", "Source retained provenance")),
                        ExplorerFixture.Display(r.TargetIdentity, DisplayEnvironmentSide.Target,
                            friendly: new ComponentDisplayValue(DisplayAvailability.AttributeOmitted)))), DisplayProjectionOutcome.Completed);
                    using (var form = new MembershipResultsForm(f.Completed))
                    {
                        form.Show(); Open(form); var c = Control(form);
                        for (int i = 0; i < originalRows.Length; i++)
                        {
                            c.SelectComparison(i, originalRows[i].SourceIdentity == null ? DisplayEnvironmentSide.Target : DisplayEnvironmentSide.Source);
                            Assert.AreEqual(R2DExpectedRow(form, c, i), c.EvidenceText);
                            Assert.AreEqual("", c.SelectedComparison.AuthoritativeRow.DefinitionStatus);
                            StringAssert.Contains(c.EvidenceText, "Membership status: Indeterminate - Unsupported");
                            Assert.AreSame(originalRows[i], c.SelectedComparison.AuthoritativeRow);
                        }
                        Assert.AreEqual(csv, ExplorerFixture.Csv(f)); CollectionAssert.AreEqual(originalRows, f.Presentation.Rows.ToArray());
                    }
                }
                var defined = DefinitionFixture(true);
                using (var form = new MembershipResultsForm(defined.Completed))
                {
                    form.Show(); Open(form); var c = Control(form); c.SelectComparison(0);
                    Assert.AreEqual(R2DExpectedRow(form, c, 0), c.EvidenceText); StringAssert.Contains(c.EvidenceText, "Definition diagnostic: SECRET-DIAGNOSTIC");
                    StringAssert.Contains(c.EvidenceText, "sequence: Source 1; Target 2; Different");
                }
            });
        }
        [TestMethod, TestCategory("D4T8R2D")]
        public void R2DExactOccurrenceAndAmbiguousMappingsRequireExplicitRow()
        {
            Sta(() =>
            {
                var repeated = ExplorerFixture.Identity(300, null, status: IdentityResolutionStatus.Unsupported);
                var f = ExplorerFixture.Make(new[] { repeated, repeated });
                using (var form = new MembershipResultsForm(f.Completed))
                {
                    form.Show(); Open(form); var c = Control(form); c.SelectOccurrence(0);
                    Assert.AreEqual(2, c.VisibleRows.Count); Assert.IsNull(c.SelectedComparison);
                    Assert.AreEqual(R2DExpectedOccurrence(c.Projection.Occurrences[0]) + Environment.NewLine + "Choose a comparison row explicitly; no arbitrary first mapping selected.", c.EvidenceText);
                    c.SelectComparison(1, DisplayEnvironmentSide.Source);
                    Assert.AreEqual(R2DExpectedOccurrence(c.Projection.Occurrences[0]) + Environment.NewLine + R2DExpectedRow(form, c, 1), c.EvidenceText);
                }
                var many = ExplorerFixture.Make(Enumerable.Range(0, 205).Select(i => ExplorerFixture.Identity(61, "resource" + i)));
                using (var host = R2DHost(many, r => "row " + r.OriginalRowIndex, out var c))
                {
                    var category = c.Navigation.Nodes.Cast<TreeNode>().Single(n => n.Text.StartsWith("Web resources", StringComparison.OrdinalIgnoreCase));
                    category.Expand(); var page = category.Nodes[0]; Assert.IsTrue(page.Text.StartsWith("Items ", StringComparison.Ordinal));
                    c.Navigation.SelectedNode = page; Assert.IsNull(c.SelectedComparison); Assert.AreEqual(page.Text, c.EvidenceText);
                    var next = category.Nodes[1]; long revision = R2DField<long>(c, "technicalContextRevision"); c.Navigation.SelectedNode = next;
                    Assert.IsTrue(R2DField<long>(c, "technicalContextRevision") > revision); Assert.AreEqual("", R2DField<string>(c, "supportEvidence"));
                    Assert.AreEqual(next.Text, c.EvidenceText); Assert.AreEqual(205, c.Projection.Rows.Count);
                }
            });
        }
        [TestMethod, TestCategory("D4T8R2D")]
        public void R2DRejectedWrapperAndIndependentSidesRemainExact()
        {
            Sta(() =>
            {
                var f = ExplorerFixture.Make(new[] { ExplorerFixture.Identity(61, "resource") }, new[] { ExplorerFixture.Identity(61, "resource") });
                var owner = f.Presentation.Rows[0];
                Func<DisplayEnvironmentSide, ComponentSideDisplayEvidence> target = side => ExplorerFixture.Display(owner.TargetIdentity, side,
                    friendly: new ComponentDisplayValue(DisplayAvailability.Available, "Target title", "Target retained"));
                f.Completed = new CompletedMembershipComparison(f.Presentation, new[] { new ComponentDisplayEvidence(0, target(DisplayEnvironmentSide.Target), target(DisplayEnvironmentSide.Target)) }, DisplayProjectionOutcome.Completed);
                using (var host = R2DHost(f, r => r.DisplayContextDiagnostic ?? r.FriendlyNameText, out var c))
                {
                    Assert.IsNull(c.Projection.Rows[0].SourceDisplay); Assert.IsNull(c.Projection.Rows[0].TargetDisplay);
                    StringAssert.Contains(c.EvidenceText, "environment side or field context is invalid");
                    f.Completed = new CompletedMembershipComparison(f.Presentation, new[] { new ComponentDisplayEvidence(0,
                        ExplorerFixture.Display(owner.SourceIdentity, DisplayEnvironmentSide.Source, friendly: new ComponentDisplayValue(DisplayAvailability.Available, "Source title", "Source retained")), target(DisplayEnvironmentSide.Target)) }, DisplayProjectionOutcome.Completed);
                    var valid = ExplorerFixture.Project(f); c.Bind(valid, new[] { 0 }, true); c.SelectComparison(0);
                    Assert.AreEqual("Source title", valid.Rows[0].PrimaryName(DisplayEnvironmentSide.Source, true));
                    Assert.AreEqual("Target title", valid.Rows[0].PrimaryName(DisplayEnvironmentSide.Target, true));
                    StringAssert.Contains(c.EvidenceText, "Source title"); StringAssert.Contains(c.EvidenceText, "Target title"); Assert.AreSame(owner, c.SelectedComparison.AuthoritativeRow);
                    f.Completed = new CompletedMembershipComparison(f.Presentation, f.Completed.DisplayEvidence, DisplayProjectionOutcome.Completed, DisplayCacheAssociation.Rejected);
                    var rejected = ExplorerFixture.Project(f); c.Bind(rejected, new[] { 0 }, true); c.SelectComparison(0);
                    Assert.IsNull(rejected.Rows[0].SourceDisplay); Assert.IsNull(rejected.Rows[0].TargetDisplay); StringAssert.Contains(c.EvidenceText, "cache association was rejected");
                }
            });
        }
        [TestMethod, TestCategory("D4T8R2D")]
        public void R2DSupportReturnPreservesOverviewPropertiesAndOriginalNavigation()
        {
            Sta(() =>
            {
                var f = DefinitionFixture(true); var csv = ExplorerFixture.Csv(f); var original = f.Presentation.Rows[0]; var spy = new FakeOrganizationService(); int captures = 0;
                using (var form = new MembershipResultsForm(f.Completed, sourceService: spy, targetService: spy, captureAppSettingEvidence: () => captures++))
                {
                    form.Show(); OpenPrimary(form); var c = Control(form); c.SelectComparison(0, DisplayEnvironmentSide.Target);
                    int occurrence = c.Projection.Rows[0].TargetOccurrences[0]; c.SelectOccurrence(occurrence); c.ShowDetails(true);
                    string overview = c.SimpleDetailsText, evidence = c.EvidenceText;
                    var view = (ComboBox)Controls(c).Single(x => x.Name == "ExplorerDetailsView"); view.SelectedIndex = 1;
                    var properties = (DataGridView)Controls(c).Single(x => x.Name == "ExplorerComparedProperties"); int count = properties.RowCount;
                    var support = (CheckBox)Controls(c).Single(x => x.Name == "ExplorerSupportEvidenceToggle"); support.Checked = true;
                    Assert.AreEqual(evidence, c.SimpleDetailsText); support.Checked = false; Assert.AreEqual(1, view.SelectedIndex); Assert.AreEqual(count, properties.RowCount);
                    Assert.IsTrue(properties.Visible); view.SelectedIndex = 0; Assert.AreEqual(overview, c.SimpleDetailsText);
                    WorkspaceButton(form, "ObjectsOriginalRowAction").PerformClick(); Assert.AreSame(original, ((MembershipGridRow)Grid(form).CurrentRow.DataBoundItem).AuthoritativeRow);
                    OpenPrimary(form); Assert.AreSame(original, c.SelectedComparison.AuthoritativeRow); Assert.AreEqual(occurrence, c.SelectedOccurrenceIndex);
                    Assert.AreEqual(DisplayEnvironmentSide.Target, c.SelectedSide); Assert.AreEqual(evidence, c.EvidenceText);
                    Assert.AreEqual(csv, ExplorerFixture.Csv(f)); Assert.AreEqual(0, captures + spy.Calls + spy.ExecuteCalls + spy.WriteCalls);
                }
            });
        }
        [TestMethod, TestCategory("D4T8R2D")]
        public void R2DTechnicalRoutesRemainExplicitAndZeroRequest()
        {
            Sta(() =>
            {
                var f = ExplorerFixture.Apps(); var spy = new FakeOrganizationService(); int captures = 0; var csv = ExplorerFixture.Csv(f);
                using (var form = new MembershipResultsForm(f.Completed, sourceService: spy, targetService: spy, captureAppSettingEvidence: () => captures++))
                {
                    form.Show(); Open(form); var c = Control(form); c.SelectComparison(0, DisplayEnvironmentSide.Source);
                    Assert.AreEqual(R2DExpectedRow(form, c, 0), c.SimpleDetailsText); var row = c.SelectedComparison;
                    using (var modal = form.CreateDisplayEvidenceDetails(row))
                    {
                        var text = modal.Controls.OfType<TextBox>().Single(); Assert.AreEqual(form.FormatDisplayEvidenceDetails(row), text.Text);
                        Assert.IsFalse(text.Text.StartsWith("Original comparison row: ", StringComparison.Ordinal));
                    }
                    View(form).SelectedIndex = 1; Pump(() => form.Controls.OfType<MembershipHierarchyControl>().Any(x => x.Visible));
                    var hierarchy = form.Controls.OfType<MembershipHierarchyControl>().Single(); hierarchy.SelectComparison(0, DisplayEnvironmentSide.Source);
                    Assert.AreSame(row.AuthoritativeRow, hierarchy.SelectedComparison.AuthoritativeRow);
                    StringAssert.Contains(((TextBox)Controls(hierarchy).Single(x => x.Name == "HierarchyEvidence")).Text, form.FormatDisplayEvidenceDetails(hierarchy.SelectedComparison));
                    View(form).SelectedIndex = 2; Application.DoEvents(); View(form).SelectedIndex = 0; Application.DoEvents();
                    Assert.AreEqual(f.Presentation.Rows.Count, Grid(form).Rows.Count); Assert.AreEqual(csv, ExplorerFixture.Csv(f));
                    Assert.AreEqual(0, captures + spy.Calls + spy.ExecuteCalls + spy.WriteCalls);
                }
            });
        }
        private static Form R2DHost(ExplorerFixture fixture, Func<MembershipGridRow, string> describe, out SolutionObjectsExplorerControl control)
        {
            int setupFormatterCalls = 0;
            var form = new Form { Size = new Size(1000, 700), StartPosition = FormStartPosition.Manual, Location = new Point(-4000, -4000), ShowInTaskbar = false };
            control = new SolutionObjectsExplorerControl(row => { setupFormatterCalls++; return describe(row); }, _ => { }, _ => { }); form.Controls.Add(control);
            control.ConfigureWorkspace(true, ObjectsWorkspaceScope.Compare); var projection = ExplorerFixture.Project(fixture);
            control.Bind(projection, Enumerable.Range(0, projection.Rows.Count), true);
            form.Show(); Application.DoEvents();
            if (projection.Rows.Count > 0)
            {
                control.SelectComparison(0, DisplayEnvironmentSide.Source); Application.DoEvents();
                Assert.AreSame(projection.Rows[0].Original, control.SelectedComparison, "Post-show setup must select the intended original comparison row.");
                Assert.AreEqual(0, control.SelectedComparison.OriginalRowIndex);
                Assert.AreEqual(DisplayEnvironmentSide.Source, control.SelectedSide);
                Assert.IsNull(control.SelectedOccurrenceIndex);
            }
            Assert.AreSame(projection, control.Projection);
            Assert.IsTrue(((SplitContainer)Controls(control).Single(x => x.Name == "ExplorerEvidenceSplit")).Panel2Collapsed);
            Assert.AreEqual(0, setupFormatterCalls, "Host setup must not format technical evidence before disclosure.");
            return form;
        }
        private static T R2DField<T>(object owner, string name) => (T)owner.GetType().GetField(name,
            System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic).GetValue(owner);
        private static string R2DExpectedRow(MembershipResultsForm form, SolutionObjectsExplorerControl c, int index)
        {
            var row = c.Projection.Rows[index]; var text = "Original comparison row: " + index + Environment.NewLine +
                "Source raw occurrence mappings: " + string.Join(",", row.SourceOccurrences) + Environment.NewLine +
                "Target raw occurrence mappings: " + string.Join(",", row.TargetOccurrences) + Environment.NewLine + form.FormatDisplayEvidenceDetails(row.Original);
            var detail = row.Original.AuthoritativeRow.DefinitionDetail;
            if (detail != null)
            {
                text += Environment.NewLine + "Definition diagnostic: " + detail.Diagnostic;
                foreach (var property in detail.Properties) text += Environment.NewLine + property.PropertyName + ": Source " + property.SourceValue + "; Target " + property.TargetValue + "; " + property.Comparison;
            }
            return text;
        }
        private static string R2DExpectedOccurrence(SolutionObjectsExplorerOccurrence item)
        {
            var o = item.Original; var r = o.Record;
            var lines = new System.Collections.Generic.List<string> { o.Side + ": " + o.Snapshot.Environment.DisplayName, "Solution: " + o.Snapshot.SolutionUniqueName,
                "Original raw index: " + o.OriginalRawIndex, "H1 occurrence index: " + o.OccurrenceIndex,
                "SolutionComponentId: " + r.SolutionComponentId, "ObjectId (local audit): " + r.ObjectId,
                "ComponentType: " + r.ComponentType, "RootSolutionComponentId: " + r.RootSolutionComponentId,
                "RootComponentBehavior: " + r.RootComponentBehavior, "IsMetadata: " + r.IsMetadata,
                "Identity resolution: " + o.Identity.Status, "Portable identity: " + o.Identity.ComparisonKey,
                "Placement: " + o.Placement + "; multiple parents: " + o.HasDifferentParentAssociations + "; not top-level proof" };
            foreach (var a in o.Associations) lines.Add(a.Kind + ": " + a.Status + "; parent occurrence=" + a.ParentOccurrenceIndex + "; provenance=" + a.Provenance + "; " + a.Diagnostic);
            lines.Add("Original identity diagnostic: " + o.Identity.Diagnostic);
            return string.Join(Environment.NewLine, lines) + Environment.NewLine;
        }

        private static ExplorerFixture DefinitionFixture(bool changed)
        {
            var f = SyntheticNames(ExplorerFixture.Make(new[] { ExplorerFixture.Identity(61, "resource") }, new[] { ExplorerFixture.Identity(61, "resource") }));
            var old = f.Completed; var rows = f.Presentation.Rows.ToArray(); string status = changed ? "Different" : "Match";
            var detail = new ComponentDefinitionDetailPresentation("Web resource", "not-a-display-name", "Present in Both", status,
                new[] { new ComponentPropertyPresentation("sequence", "1", changed ? "2" : "1", changed, changed ? "Different" : "Match"), new ComponentPropertyPresentation("unchanged_audit_field", "same", "same", false, "Match") }, "SECRET-DIAGNOSTIC");
            rows[0] = rows[0].WithDefinition(status, changed ? "sequence" : string.Empty, detail); f.Presentation = new MembershipComparisonPresentation("SAA", f.Presentation.Source, f.Presentation.Target, rows);
            f.Completed = new CompletedMembershipComparison(f.Presentation, old.DisplayEvidence, DisplayProjectionOutcome.Completed); return f;
        }

        [TestMethod, TestCategory("D4T8R2B")]
        public void R2BWorkspaceSelectorKeepsCachedCompleteRowsAndSixPrimaryColumns()
        {
            Sta(() =>
            {
                var f = ExplorerFixture.Saa(); var originals = f.Presentation.Rows.ToArray(); var csv = ExplorerFixture.Csv(f);
                var spy = new FakeOrganizationService(); int captures = 0, hierarchyBuilds = 0, explorerBuilds = 0;
                using (var form = new MembershipResultsForm(f.Completed, captureAppSettingEvidence: () => captures++, sourceService: spy, targetService: spy))
                {
                    var hierarchy = form.ProjectHierarchy; var explorer = form.ProjectExplorer;
                    form.ProjectHierarchy = token => { Interlocked.Increment(ref hierarchyBuilds); return hierarchy(token); };
                    form.ProjectExplorer = (model, token) => { Interlocked.Increment(ref explorerBuilds); return explorer(model, token); };
                    form.Show(); OpenPrimary(form); var c = Control(form); var projection = form.ExplorerProjection;
                    Assert.IsFalse(form.IsTechnicalWorkspace); Assert.AreEqual(ObjectsWorkspaceScope.Target, c.WorkspaceScope);
                    var summary = (Label)Controls(form).Single(x => x.Name == "ObjectsGlobalSummary"); var globalText = summary.Text;
                    StringAssert.Contains(globalText, "Compared components: 840"); StringAssert.Contains(globalText, "Present in Both: 672");
                    StringAssert.Contains(globalText, "Source Only: 103"); StringAssert.Contains(globalText, "Target Only: 32"); StringAssert.Contains(globalText, "Needs Review: 33");
                    Assert.AreEqual("Objects — partial view. Additional components are available in Technical Membership.", Controls(form).Single(x => x.Name == "ObjectsPartialNotice").Text);
                    for (int scope = 0; scope < 3; scope++)
                    {
                        Scope(form).SelectedIndex = scope; OpenPrimary(form);
                        CollectionAssert.AreEqual(new[] { "Display Name", "Name", "Type", "Source", "Target", "Comparison" }, c.Grid.Columns.Cast<DataGridViewColumn>().Where(x => x.Visible).Select(x => x.HeaderText).ToArray());
                    }
                    var search = Controls(form).OfType<ToolStrip>().SelectMany(t => t.Items.OfType<ToolStripTextBox>()).Single(t => t.Name == "ObjectsSearchFilter");
                    search.Text = "does not occur"; Assert.AreEqual(0, c.VisibleRows.Count); Assert.AreEqual(globalText, summary.Text);
                    WorkspaceButton(form, "ObjectsClearSearchAction").PerformClick(); Assert.AreEqual("", search.Text); Assert.IsTrue(c.VisibleRows.Count > 0);
                    for (int view = 0; view < 4; view++)
                    {
                        WorkspaceButton(form, "TechnicalMembershipAction").PerformClick(); View(form).SelectedIndex = view;
                        if (view == 1) Pump(() => form.Controls.OfType<MembershipHierarchyControl>().Any(x => x.Visible));
                        if (view == 3) Open(form);
                        Assert.IsTrue(WorkspaceButton(form, "ObjectsBackAction").Visible); Assert.IsTrue(WorkspaceButton(form, "TechnicalMembershipAction").Visible);
                        Assert.IsTrue(WorkspaceButton(form, "TechnicalMembershipAction").Checked);
                        OpenPrimary(form); WorkspaceButton(form, "TechnicalMembershipAction").PerformClick();
                        Assert.AreEqual(view, View(form).SelectedIndex, "Returning must retain the existing technical view.");
                        OpenPrimary(form);
                    }
                    Assert.AreSame(projection, form.ExplorerProjection); Assert.AreEqual(1, hierarchyBuilds); Assert.AreEqual(1, explorerBuilds);
                    Assert.AreEqual(840, projection.Rows.Count); Assert.AreEqual(792, projection.SourceOccurrenceCount); Assert.AreEqual(720, projection.TargetOccurrenceCount);
                    for (int i = 0; i < originals.Length; i++) { Assert.AreSame(originals[i], projection.Rows[i].Original.AuthoritativeRow); Assert.AreEqual(i, projection.Rows[i].OriginalRowIndex); }
                    CollectionAssert.AreEqual(originals, f.Presentation.Rows.ToArray()); Assert.AreEqual(csv, ExplorerFixture.Csv(f));
                    Assert.AreEqual(0, captures + spy.Calls + spy.ExecuteCalls + spy.WriteCalls);
                }
            });
        }

        private static ExplorerFixture LongHeaderFixture()
        {
            var f = ExplorerFixture.Saa(); var old = f.Completed;
            var source = MembershipEnvironmentResult.FromSnapshot("SYNTHETIC SOURCE — " + string.Join(" ", Enumerable.Repeat("Long environment name", 8)), f.Presentation.Source.Snapshot, 27, TimeSpan.Zero);
            var target = MembershipEnvironmentResult.FromSnapshot("SYNTHETIC TARGET — " + string.Join(" ", Enumerable.Repeat("Long target environment", 8)), f.Presentation.Target.Snapshot, 27, TimeSpan.Zero);
            f.Presentation = new MembershipComparisonPresentation("SAA", source, target, f.Presentation.Rows);
            f.Completed = new CompletedMembershipComparison(f.Presentation, old.DisplayEvidence, old.DisplayProjectionOutcome); return f;
        }

        private static void OpenPrimary(MembershipResultsForm form)
        {
            if (form.IsTechnicalWorkspace) WorkspaceButton(form, "ObjectsBackAction").PerformClick();
            Pump(() => form.ExplorerProjection != null && ControlOrNull(form)?.Visible == true && Control(form).IsPrimaryWorkspace);
        }
        private static ToolStripButton WorkspaceButton(MembershipResultsForm form, string name) => Controls(form).OfType<ToolStrip>().SelectMany(t => t.Items.OfType<ToolStripButton>()).Single(b => b.Name == name);
        private static ToolStripComboBox Scope(MembershipResultsForm form) => Controls(form).OfType<ToolStrip>().SelectMany(t => t.Items.OfType<ToolStripComboBox>()).Single(c => c.Name == "ObjectsScopeSelector");

        private static ExplorerFixture SyntheticNames(ExplorerFixture f)
        {
            var sourceRows = f.Presentation.Rows.ToLookup(r => r.SourceIdentity?.Record.SolutionComponentId);
            var targetRows = f.Presentation.Rows.ToLookup(r => r.TargetIdentity?.Record.SolutionComponentId);
            Func<ComponentSideDisplayEvidence, ComponentSideDisplayEvidence> name = side =>
            {
                var row = (side.EnvironmentSide == DisplayEnvironmentSide.Source ? sourceRows : targetRows)[side.SolutionComponentId].First(r => true);
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

        private static void Open(MembershipResultsForm form) { if (!form.IsTechnicalWorkspace) WorkspaceButton(form, "TechnicalMembershipAction").PerformClick(); if (!Button(form).Checked) Button(form).PerformClick(); Pump(() => form.ExplorerProjection != null && ControlOrNull(form)?.Visible == true); }
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
        #region R2E Explorer additions A
        [TestMethod, TestCategory("D4T8R2E")]
        public void R2EObjectsFilterStateIsIndependentFromTechnical()
        {
            Sta(() =>
            {
                using (var form = R2EOpen())
                {
                    var primary = R2ECombo(form, "ObjectsMembershipPresenceFilter"); primary.SelectedIndex = 2; OpenPrimary(form);
                    var expected = Control(form).VisibleRows.ToArray();
                    Open(form); View(form).SelectedIndex = 0; R2ECombo(form, "MembershipPresenceFilter").SelectedIndex = 3;
                    Assert.IsTrue(form.FilteredRows().All(r => r.AuthoritativeRow.Presence == MembershipPresence.OnlyInTarget));
                    OpenPrimary(form); Assert.AreEqual(2, primary.SelectedIndex); CollectionAssert.AreEqual(expected, Control(form).VisibleRows.ToArray());
                    Assert.IsTrue(Control(form).VisibleRows.All(i => Control(form).Projection.Rows[i].Original.AuthoritativeRow.Presence == MembershipPresence.OnlyInSource));
                    Open(form); Assert.AreEqual(3, R2ECombo(form, "MembershipPresenceFilter").SelectedIndex);
                }
            });
        }
        [TestMethod, TestCategory("D4T8R2E")]
        public void R2EScopeChangeRetainsSearchAndSuspendsCompareDifferences()
        {
            Sta(() =>
            {
                using (var form = R2EOpen())
                {
                    var search = R2ESearch(form); search.Text = " DEMO "; var difference = WorkspaceButton(form, "ObjectsDifferencesFilter"); difference.PerformClick(); OpenPrimary(form);
                    var projection = form.ExplorerProjection;
                    for (int scope = 0; scope < 3; scope++)
                    {
                        Scope(form).SelectedIndex = scope; OpenPrimary(form);
                        Assert.AreEqual(" DEMO ", search.Text); Assert.IsTrue(difference.Checked); Assert.AreEqual(scope == 2, difference.Enabled);
                        foreach (int index in Control(form).VisibleRows)
                        {
                            var row = projection.Rows[index];
                            if (scope == 0) Assert.IsTrue(row.TargetOccurrences.Count > 0);
                            if (scope == 1) Assert.IsTrue(row.SourceOccurrences.Count > 0);
                            if (scope == 2) Assert.IsTrue(row.Original.AuthoritativeRow.Presence != MembershipPresence.PresentInBoth || row.Original.AuthoritativeRow.DefinitionStatus == "Different");
                        }
                    }
                    Assert.AreSame(projection, form.ExplorerProjection);
                }
            });
        }
        [TestMethod, TestCategory("D4T8R2E")]
        public void R2ECategorySelectionSurvivesSearchAndClearByStableKey()
        {
            Sta(() =>
            {
                using (var form = R2EOpen())
                {
                    var c = Control(form); var category = c.Navigation.Nodes.Cast<TreeNode>().Single(n => n.Name == "workspace-Web resources");
                    c.Navigation.SelectedNode = category; category.Expand(); var expected = c.VisibleRows.ToArray(); Assert.IsNull(c.SelectedComparison);
                    R2ESearch(form).Text = "no retained name matches this";
                    Assert.AreEqual(0, c.VisibleRows.Count); Assert.AreEqual(category.Name, c.Navigation.SelectedNode.Name); Assert.IsNull(c.SelectedComparison);
                    WorkspaceButton(form, "ObjectsClearSearchAction").PerformClick();
                    Assert.AreEqual(category.Name, c.Navigation.SelectedNode.Name); CollectionAssert.AreEqual(expected, c.VisibleRows.ToArray()); Assert.IsNull(c.SelectedComparison);
                }
            });
        }
        [TestMethod, TestCategory("D4T8R2E")]
        public void R2EObjectsResetClearsFiltersWithoutChangingScopeOrSnapshot()
        {
            Sta(() =>
            {
                using (var form = R2EOpen())
                {
                    Scope(form).SelectedIndex = 1; OpenPrimary(form); var c = Control(form); var owner = form.CompletedComparison;
                    R2ECombo(form, "ObjectsMembershipPresenceFilter").SelectedIndex = 2; R2ESearch(form).Text = "source";
                    WorkspaceButton(form, "ObjectsDifferencesFilter").Checked = true; WorkspaceButton(form, "ObjectsDetailsAction").Checked = true;
                    WorkspaceButton(form, "ObjectsMetadataAction").Checked = true; WorkspaceButton(form, "SharedAdvancedDiagnosticsAction").PerformClick();
                    WorkspaceButton(form, "MembershipResetFiltersAction").PerformClick(); OpenPrimary(form);
                    Assert.AreEqual(1, Scope(form).SelectedIndex); Assert.AreEqual("", R2ESearch(form).Text); Assert.AreEqual(0, R2ECombo(form, "ObjectsMembershipPresenceFilter").SelectedIndex);
                    Assert.IsFalse(WorkspaceButton(form, "ObjectsDifferencesFilter").Checked); Assert.AreEqual("workspace-all", c.Navigation.SelectedNode.Name);
                    Assert.IsTrue(WorkspaceButton(form, "ObjectsDetailsAction").Checked); Assert.IsTrue(WorkspaceButton(form, "ObjectsMetadataAction").Checked);
                    Assert.IsTrue(WorkspaceButton(form, "SharedAdvancedDiagnosticsAction").Checked); Assert.AreSame(owner, form.CompletedComparison); Assert.IsNull(c.SelectedComparison);
                    CollectionAssert.AreEqual(c.Projection.Rows.Where(r => r.SourceOccurrences.Count > 0).Select(r => r.Original.OriginalRowIndex).ToArray(), c.VisibleRows.ToArray());
                }
            });
        }
        [TestMethod, TestCategory("D4T8R2E")]
        public void R2ETechnicalResetRestoresAllRowsAndReportsViewChange()
        {
            Sta(() =>
            {
                using (var form = R2EOpen())
                {
                    Open(form); View(form).SelectedIndex = 2; R2ECombo(form, "MembershipPresenceFilter").SelectedIndex = 2;
                    WorkspaceButton(form, "MembershipResetFiltersAction").PerformClick(); Application.DoEvents();
                    Assert.AreEqual(0, View(form).SelectedIndex); Assert.AreEqual(0, R2ECombo(form, "MembershipPresenceFilter").SelectedIndex);
                    Assert.AreEqual(form.CompletedComparison.Presentation.Rows.Count, Grid(form).Rows.Count); Assert.IsNull(Grid(form).CurrentCell);
                    StringAssert.Contains(R2EItems(form).OfType<ToolStripLabel>().Single(x => x.Name == "MembershipViewStatus").Text, "Filters reset");
                    View(form).SelectedIndex = 3; Open(form); WorkspaceButton(form, "MembershipResetFiltersAction").PerformClick(); Open(form);
                    Assert.AreEqual("all", Control(form).Navigation.SelectedNode.Name); Assert.AreEqual(form.CompletedComparison.Presentation.Rows.Count, Control(form).VisibleRows.Count);
                }
            });
        }
        [TestMethod, TestCategory("D4T8R2E")]
        public void R2EMembershipMasksKeepEveryIndeterminateSubtype()
        {
            Sta(() =>
            {
                var fixture = ExplorerFixture.Saa();
                using (var form = R2EOpen(fixture))
                {
                    var rows = fixture.Presentation.Rows.ToArray(); var csv = ExplorerFixture.Csv(fixture); Open(form);
                    for (int filter = 0; filter < 5; filter++)
                    {
                        R2ECombo(form, "MembershipPresenceFilter").SelectedIndex = filter; Open(form);
                        var presence = filter == 1 ? MembershipPresence.PresentInBoth : filter == 2 ? MembershipPresence.OnlyInSource : filter == 3 ? MembershipPresence.OnlyInTarget : MembershipPresence.Indeterminate;
                        CollectionAssert.AreEqual(rows.Select((r, i) => new { r, i }).Where(x => filter == 0 || x.r.Presence == presence).Select(x => x.i).ToArray(), form.FilteredRows().Select(r => r.OriginalRowIndex).ToArray());
                    }
                    Assert.AreEqual(33, form.FilteredRows().Count); CollectionAssert.AreEqual(rows, fixture.Presentation.Rows.ToArray()); Assert.AreEqual(csv, ExplorerFixture.Csv(fixture));
                }
                using (var form = R2EOpen(DefinitionFixture(true)))
                {
                    R2ECombo(form, "ObjectsMembershipPresenceFilter").SelectedIndex = 1; WorkspaceButton(form, "ObjectsDifferencesFilter").PerformClick(); OpenPrimary(form);
                    Assert.AreEqual(1, Control(form).VisibleRows.Count); Assert.AreEqual(MembershipPresence.PresentInBoth, Control(form).Projection.Rows[0].Original.AuthoritativeRow.Presence);
                    Assert.AreEqual("Different", Control(form).Projection.Rows[0].Original.AuthoritativeRow.DefinitionStatus);
                }
            });
        }
        [TestMethod, TestCategory("D4T8R2E")]
        public void R2EKeyboardSearchFocusAndClearUseVisibleWorkspace()
        {
            Sta(() =>
            {
                using (var form = R2EOpen())
                {
                    Control(form).Grid.Focus(); Assert.IsTrue(R2ECommand(form, Keys.Control | Keys.F)); Assert.IsTrue(R2ESearch(form).TextBox.Focused);
                    R2ESearch(form).Text = "demo"; WorkspaceButton(form, "ObjectsClearSearchAction").PerformClick();
                    Assert.AreEqual("", R2ESearch(form).Text); Assert.IsTrue(R2ESearch(form).TextBox.Focused);
                    Open(form); var prior = View(form).SelectedIndex; R2ECommand(form, Keys.Control | Keys.F);
                    Assert.IsTrue(form.IsTechnicalWorkspace); Assert.AreEqual(prior, View(form).SelectedIndex); Assert.IsFalse(R2ESearch(form).TextBox.Focused);
                    var workspace = form.IsTechnicalWorkspace; R2ECommand(form, Keys.Control | Keys.Shift | Keys.K); Assert.AreEqual(workspace, form.IsTechnicalWorkspace);
                }
            });
        }
        #endregion
        #region R2E Explorer additions B
        [TestMethod, TestCategory("D4T8R2E")]
        public void R2EGridEnterOpensDetailsWithoutTechnicalDisclosure()
        {
            Sta(() =>
            {
                using (var form = R2EOpen())
                {
                    var c = Control(form); int calls = 0; R2EObserveFormatter(c, () => calls++); var row = c.SelectedComparison; var side = c.SelectedSide;
                    Assert.IsFalse(WorkspaceButton(form, "ObjectsDetailsAction").Checked); c.Grid.Focus(); Assert.IsTrue(c.Grid.Focused);
                    Assert.IsTrue(R2ECommand(form, Keys.Enter)); Application.DoEvents(); Assert.IsTrue(WorkspaceButton(form, "ObjectsDetailsAction").Checked);
                    Assert.AreSame(row, c.SelectedComparison); Assert.AreEqual(side, c.SelectedSide); Assert.AreEqual(0, calls);
                    Assert.IsFalse(((CheckBox)Controls(c).Single(x => x.Name == "ExplorerSupportEvidenceToggle")).Checked);
                    Assert.IsFalse(WorkspaceButton(form, "SharedAdvancedDiagnosticsAction").Checked); Assert.IsFalse(form.IsTechnicalWorkspace);
                    c.ApplyFilter(new int[0], false); R2ECommand(form, Keys.Enter); Assert.IsNull(c.SelectedComparison); Assert.AreEqual(0, calls);
                }
            });
        }
        [TestMethod, TestCategory("D4T8R2E")]
        public void R2EOutOfScopeSelectionRestoresExactOriginalRowAndSide()
        {
            Sta(() =>
            {
                using (var form = R2EOpen())
                {
                    Scope(form).SelectedIndex = 1; OpenPrimary(form); var c = Control(form);
                    int index = c.Projection.Rows.Single(r => r.Original.AuthoritativeRow.Presence == MembershipPresence.OnlyInSource).Original.OriginalRowIndex;
                    using (var transitionTrace = new MembershipUiResponsivenessTests.R2ETransitionTrace(form, c, c.Projection.Rows[index].Original, "OutOfScopeOriginalToFlat"))
                    {
                    transitionTrace.Snapshot("checkpoint.before.source.occurrence");
                    c.SelectOccurrence(c.Projection.Rows[index].SourceOccurrences.Single()); var row = c.SelectedComparison;
                    Assert.IsNotNull(row, "Source occurrence must establish the exact comparison row before handoff.");
                    Assert.AreEqual(index, row.OriginalRowIndex); Assert.AreSame(c.Projection.Rows[index].Original.AuthoritativeRow, row.AuthoritativeRow);
                    Assert.AreEqual(DisplayEnvironmentSide.Source, c.SelectedSide);
                    var occurrence = c.SelectedOccurrenceIndex; Assert.AreEqual((int?)c.Projection.Rows[index].SourceOccurrences.Single(), occurrence);
                    var contextField = typeof(MembershipResultsForm).GetField("retainContextWithoutRow", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic);
                    Assert.IsNotNull(contextField);
                    transitionTrace.Snapshot("checkpoint.before.target.scope");
                    Scope(form).SelectedIndex = 0; OpenPrimary(form);
                    transitionTrace.Snapshot("checkpoint.after.target.scope");
                    Assert.IsNotNull(c.SelectedComparison, "An out-of-scope selection must remain actionable.");
                    Assert.AreEqual(occurrence, c.SelectedOccurrenceIndex); Assert.IsTrue(c.Visible); Assert.IsFalse(form.IsTechnicalWorkspace);
                    Assert.IsTrue(WorkspaceButton(form, "ObjectsOriginalRowAction").Enabled);
                    Assert.AreSame(row, c.SelectedComparison); Assert.AreEqual(DisplayEnvironmentSide.Source, c.SelectedSide); Assert.IsFalse(c.VisibleRows.Contains(index));
                    StringAssert.Contains(Controls(c).Single(x => x.Name == "ExplorerAccounting").Text, "outside");
                    transitionTrace.Snapshot("checkpoint.before.original.click");
                    WorkspaceButton(form, "ObjectsOriginalRowAction").PerformClick(); transitionTrace.Snapshot("checkpoint.after.original.click"); Assert.IsTrue(form.IsTechnicalWorkspace); Assert.AreEqual(0, View(form).SelectedIndex);
                    Assert.IsTrue(Grid(form).Visible); transitionTrace.Snapshot("checkpoint.before.flat.current-row.assertion"); Assert.IsNotNull(Grid(form).CurrentRow, "Original Row must establish the exact Flat current row.");
                    Assert.IsNotNull(Grid(form).CurrentRow.DataBoundItem, "Original Row must retain a bound comparison row.");
                    Assert.IsInstanceOfType(Grid(form).CurrentRow.DataBoundItem, typeof(MembershipGridRow));
                    Assert.AreEqual(index, ((MembershipGridRow)Grid(form).CurrentRow.DataBoundItem).OriginalRowIndex);
                    Assert.IsFalse((bool)contextField.GetValue(form), "Exact Flat row intent must replace row-less context.");
                    Assert.AreEqual(DisplayEnvironmentSide.Source, (DisplayEnvironmentSide?)typeof(MembershipResultsForm).GetField("retainedExplorerSide", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic).GetValue(form));
                    Assert.AreSame(row.AuthoritativeRow, ((MembershipGridRow)Grid(form).CurrentRow.DataBoundItem).AuthoritativeRow);
                    OpenPrimary(form); Assert.IsNotNull(c.SelectedComparison, "Return must preserve the exact out-of-scope row.");
                    Assert.AreSame(row, c.SelectedComparison); Assert.AreEqual(DisplayEnvironmentSide.Source, c.SelectedSide); Assert.AreEqual(occurrence, c.SelectedOccurrenceIndex);
                    Assert.IsFalse((bool)contextField.GetValue(form));
                    Assert.AreSame(row.AuthoritativeRow, c.SelectedComparison.AuthoritativeRow); Assert.IsFalse(c.VisibleRows.Contains(index));
                    }
                }
            });
        }
        [TestMethod, TestCategory("D4T8R2E")]
        public void R2ECategoryPageAndEmptySelectionNeverSelectArbitraryRow()
        {
            Sta(() =>
            {
                var many = ExplorerFixture.Make(Enumerable.Range(0, 205).Select(i => ExplorerFixture.Identity(61, "resource" + i)));
                using (var host = R2DHost(many, r => "exact " + r.OriginalRowIndex, out var c))
                {
                    var category = c.Navigation.Nodes.Cast<TreeNode>().Single(n => n.Name == "workspace-Web resources"); c.Navigation.SelectedNode = category;
                    Assert.IsNull(c.SelectedComparison); category.Expand(); var page = category.Nodes[0]; c.Navigation.SelectedNode = page;
                    Assert.IsNull(c.SelectedComparison); Assert.AreEqual(page.Text, c.EvidenceText);
                    c.ApplyFilter(new int[0], false); Assert.IsNull(c.SelectedComparison); Assert.AreEqual(0, c.VisibleRows.Count);
                    Assert.AreEqual("", c.EvidenceText);
                }
                var repeated = ExplorerFixture.Identity(300, null, status: IdentityResolutionStatus.Unsupported);
                using (var form = new MembershipResultsForm(ExplorerFixture.Make(new[] { repeated, repeated }).Completed))
                {
                    form.Show(); Open(form); var c = Control(form); c.SelectOccurrence(0); Assert.AreEqual(2, c.VisibleRows.Count); Assert.IsNull(c.SelectedComparison);
                    StringAssert.Contains(c.EvidenceText, "no arbitrary first mapping");
                    c.SelectComparison(1, DisplayEnvironmentSide.Source); Assert.AreEqual(1, c.SelectedComparison.OriginalRowIndex);
                }
            });
        }
        [TestMethod, TestCategory("D4T8R2E")]
        public void R2EHiddenNavigationClearsEvidenceWithoutFormatting()
        {
            Sta(() =>
            {
                var f = ExplorerFixture.Make(new[] { ExplorerFixture.Identity(61, "one"), ExplorerFixture.Identity(61, "two") }); int calls = 0;
                using (var host = R2DHost(f, r => { calls++; return "EXACT-ROW-" + r.OriginalRowIndex; }, out var c))
                {
                    Assert.AreEqual(0, calls); c.Visible = false; c.SelectComparison(1, DisplayEnvironmentSide.Source); c.ShowDetails(true);
                    Assert.AreEqual(0, calls); Assert.AreEqual("", R2DField<string>(c, "supportEvidence"));
                    c.Visible = true; Assert.AreEqual(0, calls); var support = (CheckBox)Controls(c).Single(x => x.Name == "ExplorerSupportEvidenceToggle"); support.Checked = true;
                    StringAssert.Contains(c.EvidenceText, "EXACT-ROW-1"); Assert.AreEqual(1, calls); Assert.AreEqual(c.EvidenceText, c.SimpleDetailsText); Assert.AreEqual(1, calls);
                    c.Visible = false; c.ApplyFilter(new[] { 0 }, false); c.SelectComparison(0, DisplayEnvironmentSide.Source); Assert.AreEqual(1, calls);
                    Assert.AreEqual("", R2DField<string>(c, "supportEvidence")); c.Visible = true;
                    StringAssert.Contains(c.SimpleDetailsText, "EXACT-ROW-0"); Assert.IsFalse(c.SimpleDetailsText.Contains("EXACT-ROW-1")); Assert.AreEqual(2, calls);
                }
            });
        }
        [TestMethod, TestCategory("D4T8R2E")]
        public void R2EOriginalRowRoundTripPreservesFiltersAndOccurrence()
        {
            Sta(() =>
            {
                using (var form = R2EOpen())
                {
                    Scope(form).SelectedIndex = 1; OpenPrimary(form); R2ECombo(form, "ObjectsMembershipPresenceFilter").SelectedIndex = 2; R2ESearch(form).Text = "source";
                    var c = Control(form); int index = c.VisibleRows.Single(); int occurrence = c.Projection.Rows[index].SourceOccurrences.Single(); c.SelectOccurrence(occurrence);
                    WorkspaceButton(form, "ObjectsDetailsAction").Checked = true; var view = (ComboBox)Controls(c).Single(x => x.Name == "ExplorerDetailsView"); view.SelectedIndex = 1;
                    var support = (CheckBox)Controls(c).Single(x => x.Name == "ExplorerSupportEvidenceToggle"); support.Checked = true; string evidence = c.EvidenceText; support.Checked = false;
                    var original = c.SelectedComparison.AuthoritativeRow; WorkspaceButton(form, "ObjectsOriginalRowAction").PerformClick();
                    Assert.AreSame(original, ((MembershipGridRow)Grid(form).CurrentRow.DataBoundItem).AuthoritativeRow); OpenPrimary(form);
                    Assert.AreEqual(1, Scope(form).SelectedIndex); Assert.AreEqual(2, R2ECombo(form, "ObjectsMembershipPresenceFilter").SelectedIndex); Assert.AreEqual("source", R2ESearch(form).Text);
                    Assert.AreSame(original, c.SelectedComparison.AuthoritativeRow); Assert.AreEqual(occurrence, c.SelectedOccurrenceIndex); Assert.AreEqual(DisplayEnvironmentSide.Source, c.SelectedSide);
                    Assert.AreEqual(1, view.SelectedIndex); Assert.IsFalse(support.Checked); Assert.AreEqual(evidence, c.EvidenceText);
                }
            });
        }
        [TestMethod, TestCategory("D4T8R2E")]
        public void R2EExplorerRestoresFocusAndScrollWithoutStealingToolbarFocus()
        {
            Sta(() =>
            {
                var ids = Enumerable.Range(0, 45).Select(i => ExplorerFixture.Identity(61, "resource" + i)).ToArray();
                using (var form = R2EOpen(SyntheticNames(ExplorerFixture.Make(ids, ids))))
                {
                    var c = Control(form); c.Grid.Focus(); c.Grid.FirstDisplayedScrollingRowIndex = 20; c.Grid.HorizontalScrollingOffset = 80;
                    int top = c.VisibleRows[c.Grid.FirstDisplayedScrollingRowIndex]; int offset = c.Grid.HorizontalScrollingOffset;
                    Open(form); OpenPrimary(form);
                    Assert.IsTrue(c.Grid.Focused); Assert.AreEqual(top, c.VisibleRows[c.Grid.FirstDisplayedScrollingRowIndex]); Assert.AreEqual(offset, c.Grid.HorizontalScrollingOffset);
                    R2ESearch(form).Focus(); Assert.IsTrue(R2ESearch(form).TextBox.Focused); R2ESearch(form).Text = "demo"; Application.DoEvents();
                    Assert.IsTrue(R2ESearch(form).TextBox.Focused); WorkspaceButton(form, "ObjectsClearSearchAction").PerformClick(); Assert.IsTrue(R2ESearch(form).TextBox.Focused);
                }
            });
        }
        [TestMethod, TestCategory("D4T8R2E")]
        public void R2ERapidFilterResetAndWorkspaceChangesDiscardStaleBindings()
        {
            Sta(() =>
            {
                using (var started = new ManualResetEventSlim()) using (var release = new ManualResetEventSlim())
                using (var form = new MembershipResultsForm(R2ETiny().Completed))
                {
                    int builds = 0; var project = form.ProjectExplorer;
                    form.ProjectExplorer = (h, token) => { Interlocked.Increment(ref builds); started.Set(); release.Wait(token); return project(h, token); };
                    try
                    {
                        form.Show(); Pump(() => started.IsSet); Scope(form).SelectedIndex = 2; R2ESearch(form).Text = "alpha";
                        WorkspaceButton(form, "MembershipResetFiltersAction").PerformClick(); View(form).SelectedIndex = 0; release.Set();
                        Pump(() => Grid(form).Visible); Application.DoEvents(); Assert.AreEqual(0, View(form).SelectedIndex); Assert.IsTrue(form.IsTechnicalWorkspace);
                        Assert.IsFalse(ControlOrNull(form)?.Visible == true); OpenPrimary(form); Assert.AreEqual(1, builds); Assert.AreEqual("", R2ESearch(form).Text);
                        Assert.AreEqual(3, form.ExplorerProjection.Rows.Count); Assert.AreEqual(4, form.ExplorerProjection.Occurrences.Count);
                    }
                    finally { release.Set(); }
                }
            });
        }
        private static ExplorerFixture R2ETiny() => SyntheticNames(ExplorerFixture.Make(
            new[] { ExplorerFixture.Identity(61, "alpha"), ExplorerFixture.Identity(61, "sourceOnly") },
            new[] { ExplorerFixture.Identity(61, "alpha"), ExplorerFixture.Identity(61, "targetOnly") }));
        private static MembershipResultsForm R2EOpen(ExplorerFixture fixture = null)
        {
            var form = new MembershipResultsForm((fixture ?? R2ETiny()).Completed) { StartPosition = FormStartPosition.Manual, Location = new Point(-4000, -4000), ShowInTaskbar = false };
            form.Show(); OpenPrimary(form); Scope(form).SelectedIndex = 2; OpenPrimary(form); Application.DoEvents(); var c = Control(form);
            int calls = 0; R2EObserveFormatter(c, () => calls++);
            if (c.VisibleRows.Count > 0)
            {
                int index = c.VisibleRows[0]; c.SelectComparison(index, DisplayEnvironmentSide.Target); Application.DoEvents();
                Assert.AreSame(c.Projection.Rows[index].Original, c.SelectedComparison); Assert.AreEqual(DisplayEnvironmentSide.Target, c.SelectedSide);
                Assert.IsNull(c.SelectedOccurrenceIndex);
            }
            Assert.IsTrue(((SplitContainer)Controls(c).Single(x => x.Name == "ExplorerEvidenceSplit")).Panel2Collapsed);
            Assert.AreEqual(0, calls, "Post-show setup must not disclose technical evidence."); return form;
        }
        private static void R2EObserveFormatter(SolutionObjectsExplorerControl control, Action call)
        {
            var field = typeof(SolutionObjectsExplorerControl).GetField("describe", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);
            var original = (Func<MembershipGridRow, string>)field.GetValue(control);
            field.SetValue(control, (Func<MembershipGridRow, string>)(row => { call(); return original(row); }));
        }
        private static System.Collections.Generic.IEnumerable<ToolStripItem> R2EItems(MembershipResultsForm form) => Controls(form).OfType<ToolStrip>().SelectMany(t => t.Items.Cast<ToolStripItem>());
        private static ToolStripComboBox R2ECombo(MembershipResultsForm form, string name) => R2EItems(form).OfType<ToolStripComboBox>().Single(x => x.Name == name);
        private static ToolStripTextBox R2ESearch(MembershipResultsForm form) => R2EItems(form).OfType<ToolStripTextBox>().Single(x => x.Name == "ObjectsSearchFilter");
        private static bool R2ECommand(MembershipResultsForm form, Keys keys)
        {
            var args = new object[] { Message.Create(form.Handle, 0x100, IntPtr.Zero, IntPtr.Zero), keys };
            return (bool)typeof(MembershipResultsForm).GetMethod("ProcessCmdKey", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance).Invoke(form, args);
        }
        #endregion
    }
}
