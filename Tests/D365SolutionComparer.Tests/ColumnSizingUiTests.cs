using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Runtime.ExceptionServices;
using System.Threading;
using System.Windows.Forms;
using D365SolutionComparer.Models;
using D365SolutionComparer.Models.Membership;
using D365SolutionComparer.Models.Presentation;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace D365SolutionComparer.Tests
{
    [TestClass]
    public class ColumnSizingUiTests
    {
        [TestMethod]
        public void DisplayNameBoundarySupportsIndependentManualWidthChanges()
        {
            WithExplorer((host, panel) =>
            {
                var grid = panel.Explorer.Grid; var before = Widths(grid); var column = grid.Columns["DisplayName"];
                Assert.IsTrue(grid.AllowUserToResizeColumns);
                Assert.AreEqual(DataGridViewAutoSizeColumnsMode.None, grid.AutoSizeColumnsMode);
                Assert.AreEqual(DataGridViewAutoSizeColumnMode.None, column.AutoSizeMode);
                Assert.AreEqual(DataGridViewTriState.True, column.Resizable);
                var boundary = grid.GetColumnDisplayRectangle(column.Index, true);
                Assert.AreEqual(DataGridViewHitTestType.ColumnHeader, grid.HitTest(boundary.Right - 2, grid.ColumnHeadersHeight / 2).Type);
                int events = 0; grid.ColumnWidthChanged += (s, e) => { if (e.Column == column) events++; };
                column.Width += 77;
                Assert.AreEqual(before["DisplayName"] + 77, column.Width); Assert.AreEqual(1, events);
                foreach (var pair in before.Where(p => p.Key != "DisplayName")) Assert.AreEqual(pair.Value, grid.Columns[pair.Key].Width, pair.Key);
            });
        }

        [TestMethod]
        public void EveryPrimaryColumnResizesIndependentlyAndKeepsItsOrder()
        {
            WithExplorer((host, panel) =>
            {
                var grid = panel.Explorer.Grid;
                foreach (DataGridViewColumn column in grid.Columns.Cast<DataGridViewColumn>().Where(c => c.Visible))
                {
                    var before = Widths(grid); int order = column.DisplayIndex;
                    Assert.AreEqual(DataGridViewAutoSizeColumnMode.None, column.AutoSizeMode);
                    Assert.AreEqual(DataGridViewTriState.True, column.Resizable); column.Width += 33;
                    Assert.AreEqual(before[column.Name] + 33, column.Width); Assert.AreEqual(order, column.DisplayIndex);
                    foreach (var pair in before.Where(p => p.Key != column.Name)) Assert.AreEqual(pair.Value, grid.Columns[pair.Key].Width, pair.Key);
                }
            });
        }

        [TestMethod]
        public void InitialWidthsHaveUsefulMinimumsAndClampNarrowRequests()
        {
            WithExplorer((host, panel) =>
            {
                foreach (DataGridViewColumn column in panel.Explorer.Grid.Columns.Cast<DataGridViewColumn>().Where(c => c.Visible))
                {
                    Assert.IsTrue(column.Width > column.MinimumWidth, column.Name);
                    Assert.IsTrue(column.MinimumWidth >= 65, column.Name);
                    column.Width = column.MinimumWidth - 1;
                    Assert.AreEqual(column.MinimumWidth, column.Width, column.Name);
                }
            });
        }

        [TestMethod]
        public void WideColumnsKeepHorizontalScrollingWithoutChangingSelection()
        {
            WithExplorer((host, panel) =>
            {
                Item<ToolStripComboBox>(panel, "SolutionExplorerScope").SelectedIndex = 1;
                var c = panel.Explorer; c.SelectComparison(0, DisplayEnvironmentSide.Source); var selected = c.SelectedComparison;
                Assert.IsNotNull(selected); c.Grid.Columns["DisplayName"].Width = 700;
                Assert.AreEqual(ScrollBars.Both, c.Grid.ScrollBars); Assert.IsTrue(Widths(c.Grid).Values.Sum() > c.Grid.ClientSize.Width);
                c.Grid.HorizontalScrollingOffset = 120; Assert.IsTrue(c.Grid.HorizontalScrollingOffset > 0);
                Assert.AreSame(selected, c.SelectedComparison); Assert.AreEqual(DisplayEnvironmentSide.Source, c.SelectedSide);
            });
        }

        [TestMethod]
        public void ScopeSearchStatusDifferencesAndResetKeepAdjustedWidths()
        {
            WithExplorer((host, panel) =>
            {
                Adjust(panel.Explorer.Grid); var expected = Widths(panel.Explorer.Grid);
                var scope = Item<ToolStripComboBox>(panel, "SolutionExplorerScope");
                foreach (int index in new[] { 1, 2, 0 }) { scope.SelectedIndex = index; AssertWidths(expected, panel.Explorer.Grid); }
                Item<ToolStripTextBox>(panel, "SolutionExplorerSearch").Text = "alpha"; AssertWidths(expected, panel.Explorer.Grid);
                Item<ToolStripComboBox>(panel, "SolutionExplorerStatus").SelectedIndex = 1; AssertWidths(expected, panel.Explorer.Grid);
                Item<ToolStripButton>(panel, "SolutionExplorerDifferences").Checked = true; AssertWidths(expected, panel.Explorer.Grid);
                Item<ToolStripButton>(panel, "SolutionExplorerReset").PerformClick(); AssertWidths(expected, panel.Explorer.Grid);
                Assert.AreEqual(string.Empty, Item<ToolStripTextBox>(panel, "SolutionExplorerSearch").Text);
                Assert.AreEqual(0, Item<ToolStripComboBox>(panel, "SolutionExplorerStatus").SelectedIndex);
            });
        }

        [TestMethod]
        public void CategoryNavigationAndGridRefreshKeepAdjustedWidths()
        {
            WithExplorer((host, panel) =>
            {
                var c = panel.Explorer; Adjust(c.Grid); var expected = Widths(c.Grid);
                Assert.IsTrue(c.Navigation.Nodes.Count > 1);
                foreach (TreeNode category in c.Navigation.Nodes)
                { c.Navigation.SelectedNode = category; AssertWidths(expected, c.Grid); }
                c.Grid.Refresh(); panel.Bind(panel.Projection); AssertWidths(expected, c.Grid);
            });
        }

        [TestMethod]
        public void WindowResizeKeepsWidthsAndAuthoritativeSelection()
        {
            WithExplorer((host, panel) =>
            {
                Item<ToolStripComboBox>(panel, "SolutionExplorerScope").SelectedIndex = 2;
                var c = panel.Explorer; c.SelectComparison(0, DisplayEnvironmentSide.Target); var selected = c.SelectedComparison;
                Assert.IsNotNull(selected); Adjust(c.Grid); var widths = Widths(c.Grid); int? occurrence = c.SelectedOccurrenceIndex;
                host.ClientSize = new Size(1250, 720); Application.DoEvents(); AssertWidths(widths, c.Grid);
                host.ClientSize = new Size(960, 600); Application.DoEvents(); AssertWidths(widths, c.Grid);
                Assert.AreSame(selected, c.SelectedComparison); Assert.AreSame(panel.Projection.Rows[0].Original, selected);
                Assert.AreEqual(DisplayEnvironmentSide.Target, c.SelectedSide); Assert.AreEqual(occurrence, c.SelectedOccurrenceIndex);
            });
        }

        [TestMethod]
        public void ColumnChangesPreserveSelectionSideAndScrollContext()
        {
            WithExplorer((host, panel) =>
            {
                Item<ToolStripComboBox>(panel, "SolutionExplorerScope").SelectedIndex = 1;
                var c = panel.Explorer; c.SelectComparison(0, DisplayEnvironmentSide.Source); var before = c.CaptureNavigation();
                Assert.IsNotNull(c.SelectedComparison); var selected = c.SelectedComparison; Adjust(c.Grid);
                var after = c.CaptureNavigation(); Assert.AreSame(selected, c.SelectedComparison);
                Assert.AreSame(before.Owner, after.Owner); Assert.AreEqual(before.Row, after.Row);
                Assert.AreEqual(before.Side, after.Side); Assert.AreEqual(before.Occurrence, after.Occurrence);
                Assert.AreEqual(before.TopRow, after.TopRow); Assert.AreEqual(before.HorizontalOffset, after.HorizontalOffset);
            });
        }

        [TestMethod]
        public void MainComparisonRebindingPreservesWidthsIncludingEmptyFilters()
        {
            Sta(() =>
            {
                using (var host = new Form { ClientSize = new Size(1100, 650) })
                using (var control = new SolutionComparerControl())
                {
                    Field(control, "suppressSettingsSave", true);
                    host.Controls.Add(control); host.Show(); Application.DoEvents();
                    var grid = (DataGridView)Field(control, "dgvResults");
                    var rows = new List<SolutionCompareResult> { new SolutionCompareResult { UniqueName = "alpha", Status = "Match" },
                        new SolutionCompareResult { UniqueName = "beta", Status = "Publisher Mismatch" } };
                    Field(control, "comparisonResults", rows);
                    var all = (ToolStripMenuItem)Field(control, "miAll"); all.Checked = true;
                    Invoke(control, "BindFilteredResults"); Assert.AreEqual(2, grid.Rows.Count);
                    grid.Columns["SourceDisplayName"].Width += 77; grid.Columns["TargetDisplayName"].Width += 31;
                    var expected = Widths(grid); var order = grid.Columns.Cast<DataGridViewColumn>().ToDictionary(c => c.Name, c => c.DisplayIndex);
                    ((CheckBox)Field(control, "chkChangedOnly")).Checked = true; Assert.AreEqual(1, grid.Rows.Count); AssertWidths(expected, grid);
                    Field(control, "comparisonResults", new List<SolutionCompareResult>()); Invoke(control, "BindFilteredResults"); Assert.AreEqual(0, grid.Rows.Count); AssertWidths(expected, grid);
                    Field(control, "comparisonResults", rows); ((CheckBox)Field(control, "chkChangedOnly")).Checked = false;
                    Assert.AreEqual(2, grid.Rows.Count); AssertWidths(expected, grid);
                    foreach (var pair in order) Assert.AreEqual(pair.Value, grid.Columns[pair.Key].DisplayIndex);
                }
            });
        }

        private static void WithExplorer(Action<Form, SimplifiedSolutionExplorerPanel> check) => Sta(() =>
        {
            var f = ExplorerFixture.Make(new[] { ExplorerFixture.Identity(1, "alpha"), ExplorerFixture.Identity(61, "beta") },
                new[] { ExplorerFixture.Identity(1, "alpha"), ExplorerFixture.Identity(61, "beta") });
            f.Completed = new CompletedMembershipComparison(f.Presentation, f.Presentation.Rows.Select((row, index) =>
                new ComponentDisplayEvidence(index, Named(row.SourceIdentity, DisplayEnvironmentSide.Source),
                    Named(row.TargetIdentity, DisplayEnvironmentSide.Target))), DisplayProjectionOutcome.Completed,
                DisplayCacheAssociation.Verified, DisplayCacheAssociation.Verified);
            using (var host = new Form { ClientSize = new Size(1100, 650) })
            using (var panel = new SimplifiedSolutionExplorerPanel(f.Presentation, row => { Assert.Fail("No technical formatter expected"); return ""; }))
            {
                panel.UseObjectNavigation(); host.Controls.Add(panel); host.Show(); panel.Bind(ExplorerFixture.Project(f)); Application.DoEvents();
                Assert.AreEqual(2, panel.Explorer.VisibleObjects.Count, "Named synthetic fixture must expose both objects");
                check(host, panel);
            }
        });
        private static ComponentSideDisplayEvidence Named(ComponentIdentity identity, DisplayEnvironmentSide side) =>
            ExplorerFixture.Display(identity, side, friendly: new ComponentDisplayValue(DisplayAvailability.Available,
                identity.ComparisonKey, "Synthetic retained friendly name"), name: new ComponentDisplayValue(DisplayAvailability.Available,
                identity.ComparisonKey, "Synthetic retained logical name"));
        private static Dictionary<string, int> Widths(DataGridView grid) => grid.Columns.Cast<DataGridViewColumn>().Where(c => c.Visible).ToDictionary(c => c.Name, c => c.Width);
        private static void Adjust(DataGridView grid) { foreach (DataGridViewColumn c in grid.Columns.Cast<DataGridViewColumn>().Where(c => c.Visible)) c.Width += 23; }
        private static void AssertWidths(Dictionary<string, int> expected, DataGridView grid)
        { CollectionAssert.AreEquivalent(expected.Keys.ToArray(), Widths(grid).Keys.ToArray()); foreach (var pair in expected) Assert.AreEqual(pair.Value, grid.Columns[pair.Key].Width, pair.Key); }
        private static IEnumerable<Control> Descendants(Control root) => root.Controls.Cast<Control>().SelectMany(c => new[] { c }.Concat(Descendants(c)));
        private static T Item<T>(Control root, string name) where T : ToolStripItem => Descendants(root).OfType<ToolStrip>().SelectMany(t => t.Items.OfType<T>()).Single(i => i.Name == name);
        private static object Field(object target, string name) => target.GetType().GetField(name, BindingFlags.NonPublic | BindingFlags.Instance).GetValue(target);
        private static void Field(object target, string name, object value) => target.GetType().GetField(name, BindingFlags.NonPublic | BindingFlags.Instance).SetValue(target, value);
        private static void Invoke(object target, string name) => target.GetType().GetMethod(name, BindingFlags.NonPublic | BindingFlags.Instance).Invoke(target, null);
        private static void Sta(Action action)
        {
            Exception error = null; var thread = new Thread(() => { try { action(); } catch (Exception e) { error = e; } });
            thread.SetApartmentState(ApartmentState.STA); thread.IsBackground = true; thread.Start(); Assert.IsTrue(thread.Join(60000));
            if (error != null) ExceptionDispatchInfo.Capture(error).Throw();
        }
    }
}
