using System;
using System.Drawing;
using System.Linq;
using System.Runtime.ExceptionServices;
using System.Threading;
using System.Windows.Forms;
using D365SolutionComparer.Models.Membership;
using D365SolutionComparer.Services.Membership;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using static D365SolutionComparer.Tests.MembershipTestData;

namespace D365SolutionComparer.Tests
{
    [TestClass]
    public class MembershipUiResponsivenessTests
    {
        [DataTestMethod, TestCategory("Phase2G15Coverage")]
        [DataRow(65)]
        [DataRow(899)]
        public void MembershipGridAccessibilityPreservesRowsWithoutRecursiveOffscreenBounds(int count)
        {
            OnStaThread(() =>
            {
                var snapshot = Snapshot(Enumerable.Range(0, count)
                    .Select(i => Identity("resource-" + i.ToString("D4"), kind: ComponentSemanticKinds.WebResource))
                    .ToArray());
                var presentation = Present(snapshot, snapshot);
                using (var form = new MembershipResultsForm(presentation))
                {
                    ShowLocally(form);
                    var grid = form.Controls.OfType<DataGridView>().Single();
                    Assert.AreEqual(count, grid.Rows.Count);
                    Assert.AreEqual(count, presentation.Summary.PresentInBoth);
                    var visible = grid.Rows[grid.FirstDisplayedScrollingRowIndex];
                    Assert.AreEqual(grid.RectangleToScreen(grid.GetRowDisplayRectangle(visible.Index, true)),
                        visible.AccessibilityObject.Bounds);

                    // Accessibility clients inspect off-screen rows too. Bounds must not recursively
                    // walk all preceding rows, while row/cell names, values and navigation remain intact.
                    foreach (DataGridViewRow row in grid.Rows)
                    {
                        var accessible = row.AccessibilityObject;
                        if (!row.Displayed) Assert.AreEqual(Rectangle.Empty, accessible.Bounds);
                        Assert.AreEqual(AccessibleRole.Row, accessible.Role);
                        Assert.AreEqual(grid.Columns.Count, accessible.GetChildCount());
                        Assert.IsFalse(string.IsNullOrWhiteSpace(accessible.Value));
                    }
                    Assert.AreEqual("resource-" + (count - 1).ToString("D4"),
                        grid.Rows[count - 1].Cells["PortableKey"].Value);
                    Assert.AreEqual(count, presentation.Rows.Count);
                    Assert.AreEqual(1, presentation.Source.Diagnostics.RequestCount);
                    Assert.AreEqual(1, presentation.Target.Diagnostics.RequestCount);
                }
            });
        }

        [TestMethod, TestCategory("Phase2G15Coverage")]
        public void ScrollingRestoresVisibleAccessibleBoundsAndPreservesSelection()
        {
            OnStaThread(() =>
            {
                var snapshot = Snapshot(Enumerable.Range(0, 899)
                    .Select(i => Identity("resource-" + i.ToString("D4"), kind: ComponentSemanticKinds.WebResource))
                    .ToArray());
                using (var form = new MembershipResultsForm(Present(snapshot, snapshot)))
                {
                    ShowLocally(form);
                    var grid = form.Controls.OfType<DataGridView>().Single();
                    grid.CurrentCell = grid.Rows[898].Cells["PortableKey"];
                    grid.FirstDisplayedScrollingRowIndex = 898;
                    Application.DoEvents();
                    var last = grid.Rows[898];
                    Assert.IsTrue(last.Displayed);
                    Assert.AreEqual(grid.RectangleToScreen(grid.GetRowDisplayRectangle(last.Index, true)),
                        last.AccessibilityObject.Bounds);
                    Assert.AreEqual(Rectangle.Empty, grid.Rows[0].AccessibilityObject.Bounds);
                    Assert.AreEqual("resource-0898", last.Cells["PortableKey"].Value);
                    Assert.IsTrue(last.Selected);
                }
            });
        }

        [TestMethod, TestCategory("Phase2G15Coverage")]
        public void EduSizedCoverageDetailsPreservesCompleteEvidenceWithoutInvokingCollectors()
        {
            OnStaThread(() =>
            {
                Func<int, MembershipSnapshot> snapshot = count => Snapshot(Enumerable.Range(0, count)
                    .Select(i => new ComponentIdentity(new SolutionComponentRecord(Guid.NewGuid(), 59, Guid.NewGuid()),
                        IdentityResolutionStatus.Unsupported, null, "Stable unsupported diagnostic",
                        "unsupported:componenttype:59",
                        diagnosticEvidence: new[] { "evidence-" + i + ":" + new string('x', 5000) }))
                    .ToArray());
                var source = snapshot(792);
                var target = snapshot(720);
                var presentation = Present(source, target);
                var builder = new MembershipCoverageDiagnosticsBuilder();
                var sourceCoverage = builder.Build(source);
                var targetCoverage = builder.Build(target);
                var sourceService = new FakeOrganizationService();
                var targetService = new FakeOrganizationService();
                int actions = 0;
                Action capture = () => actions++;
                using (var form = new MembershipCoverageDetailsForm("Source", sourceCoverage,
                    "Target", targetCoverage, presentation, sourceService: sourceService,
                    targetService: targetService, captureAppSettingEvidence: capture
#if DEBUG
                    , captureType92Evidence: capture, captureType59Evidence: capture, discoverType59Evidence: capture,
                    captureCloudFlowEvidence: capture, captureSavedQueryEvidence: capture
#endif
                    ))
                {
                    ShowLocally(form);
                    var tabs = form.Controls.OfType<TabControl>().Single();
                    for (int side = 0; side < 2; side++)
                    {
                        tabs.SelectedIndex = side;
                        Application.DoEvents();
                        var evidence = tabs.TabPages[side].Controls.OfType<RichTextBox>().Single();
                        int count = side == 0 ? 792 : 720;
                        StringAssert.Contains(evidence.Text, "evidence-0:" + new string('x', 5000));
                        StringAssert.Contains(evidence.Text, "evidence-" + (count - 1) + ":" + new string('x', 5000));
                        Assert.AreEqual(count, evidence.Text.Split(new[] { "      evidence-" },
                            StringSplitOptions.None).Length - 1);
                        Assert.IsTrue(evidence.ReadOnly);
                        Assert.IsFalse(evidence.WordWrap);
                    }
                    Assert.AreEqual(0, actions);
                    Assert.AreEqual(0, sourceService.Calls + sourceService.ExecuteCalls);
                    Assert.AreEqual(0, targetService.Calls + targetService.ExecuteCalls);
                    Assert.AreEqual(1512, presentation.Summary.Unsupported);
                    Assert.AreEqual(792, sourceCoverage.TotalCandidates);
                    Assert.AreEqual(720, targetCoverage.TotalCandidates);
                    Assert.AreEqual(1, sourceCoverage.SemanticKinds.Single(k => k.TotalCandidates == 792)
                        .DiagnosticGroups.Count);
                }
            });
        }

        private static MembershipComparisonPresentation Present(MembershipSnapshot source, MembershipSnapshot target) =>
            new MembershipResultPresenter().Create(
                MembershipEnvironmentResult.FromSnapshot("Source", source, 1, TimeSpan.Zero),
                MembershipEnvironmentResult.FromSnapshot("Target", target, 1, TimeSpan.Zero));

        private static void ShowLocally(Form form)
        {
            form.StartPosition = FormStartPosition.Manual;
            form.Location = new Point(-4000, -4000);
            form.Show();
            Application.DoEvents();
        }

        private static void OnStaThread(Action test)
        {
            Exception error = null;
            var thread = new Thread(() =>
            {
                try { test(); }
                catch (Exception ex) { error = ex; }
            }) { IsBackground = true };
            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();
            Assert.IsTrue(thread.Join(TimeSpan.FromSeconds(30)), "Membership UI initialization did not complete.");
            if (error != null) ExceptionDispatchInfo.Capture(error).Throw();
        }
    }
}
