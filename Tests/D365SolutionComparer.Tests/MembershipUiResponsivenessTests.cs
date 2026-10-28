using System;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Runtime.ExceptionServices;
using System.Threading;
using System.Windows.Forms;
using D365SolutionComparer.Models.Membership;
using D365SolutionComparer.Models.Presentation;
using D365SolutionComparer.Services.Membership;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using McTools.Xrm.Connection;
using Microsoft.Xrm.Tooling.Connector;
using Entity = Microsoft.Xrm.Sdk.Entity;
using EntityCollection = Microsoft.Xrm.Sdk.EntityCollection;
using IOrganizationService = Microsoft.Xrm.Sdk.IOrganizationService;
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
                    // This test exercises the existing Flat grid, independently of the initial Explorer.
                    form.Controls.OfType<ToolStrip>().Single(t => t.Name == "MembershipViewBar")
                        .Items.OfType<ToolStripComboBox>().Single(c => c.Name == "MembershipViewSelector").SelectedIndex = 0;
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
                        Assert.AreEqual(grid.Columns.Cast<DataGridViewColumn>().Count(c => c.Visible), accessible.GetChildCount());
                        Assert.IsFalse(string.IsNullOrWhiteSpace(accessible.Value));
                    }
                    Assert.AreEqual("resource-" + (count - 1).ToString("D4"),
                        ((MembershipGridRow)grid.Rows[count - 1].DataBoundItem).PortableKey);
                    StringAssert.Contains(grid.Rows[count - 1].AccessibilityObject.Description, "Membership status: Present in Both");
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
                    // This test exercises the existing Flat grid, independently of the initial Explorer.
                    form.Controls.OfType<ToolStrip>().Single(t => t.Name == "MembershipViewBar")
                        .Items.OfType<ToolStripComboBox>().Single(c => c.Name == "MembershipViewSelector").SelectedIndex = 0;
                    ShowLocally(form);
                    var grid = form.Controls.OfType<DataGridView>().Single();
                    grid.CurrentCell = grid.Rows[898].Cells["FriendlyNameText"];
                    grid.FirstDisplayedScrollingRowIndex = 898;
                    Application.DoEvents();
                    var last = grid.Rows[898];
                    Assert.IsTrue(last.Displayed);
                    Assert.AreEqual(grid.RectangleToScreen(grid.GetRowDisplayRectangle(last.Index, true)),
                        last.AccessibilityObject.Bounds);
                    Assert.AreEqual(Rectangle.Empty, grid.Rows[0].AccessibilityObject.Bounds);
                    Assert.AreEqual("resource-0898", ((MembershipGridRow)last.DataBoundItem).PortableKey);
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

        [DataTestMethod, TestCategory("MembershipSummaryAccounting"), TestCategory("Phase2GPresentation")]
        [DataRow(900)]
        [DataRow(1400)]
        [DataRow(1920)]
        public void CompleteSummaryFitsExistingHeaderAndPreservesGridWithoutRequests(int width)
        {
            OnStaThread(() =>
            {
                var matchOnly = new ComponentIdentity(new SolutionComponentRecord(Guid.NewGuid(), 60, Guid.NewGuid()),
                    IdentityResolutionStatus.Resolved, "synthetic-form", componentTypeKey: ComponentSemanticKinds.SystemForm,
                    inventoryAbsencePolicy: InventoryAbsencePolicy.MatchOnly);
                var presentation = Present(Snapshot(matchOnly, Identity(null, 300, IdentityResolutionStatus.Unsupported)), Snapshot());
                var sourceService = new FakeOrganizationService();
                var targetService = new FakeOrganizationService();
                using (var form = new MembershipResultsForm(presentation, sourceService: sourceService, targetService: targetService))
                {
                    form.Size = new Size(width, 760);
                    // This test exercises the existing Flat grid, independently of the initial Explorer.
                    form.Controls.OfType<ToolStrip>().Single(t => t.Name == "MembershipViewBar")
                        .Items.OfType<ToolStripComboBox>().Single(c => c.Name == "MembershipViewSelector").SelectedIndex = 0;
                    ShowLocally(form);
                    var header = form.Controls.OfType<Panel>().Single();
                    var summary = header.Controls.OfType<Label>().Single(label => label.Name == "MembershipSummary");
                    Assert.AreEqual(174, header.Height);
                    StringAssert.Contains(summary.Text, "Total Comparison Rows: 2");
                    StringAssert.Contains(summary.Text, "Resolved / Indeterminate: 1");
                    StringAssert.Contains(summary.Text, "Unsupported: 1");
                    Assert.IsFalse(summary.Text.Contains("Other / Unavailable"));
                    var measured = TextRenderer.MeasureText(summary.Text, summary.Font,
                        new Size(summary.ClientSize.Width, int.MaxValue), TextFormatFlags.WordBreak | TextFormatFlags.TextBoxControl);
                    Assert.IsTrue(measured.Height <= summary.Height, "Summary text must fit at the existing minimum width.");
                    var diagnostics = header.Controls.OfType<Label>().Single(label => label.Dock == DockStyle.Fill);
                    StringAssert.Contains(diagnostics.Text, "Raw Source Components=2");
                    StringAssert.Contains(diagnostics.Text, "Raw Target Components=0");
                    var grid = form.Controls.OfType<DataGridView>().Single();
                    Assert.AreEqual(9, grid.Columns.Count);
                    Assert.AreEqual(presentation.Rows.Count, grid.Rows.Count);
                    for (int i = 0; i < grid.Rows.Count; i++)
                    {
                        var adapter = (MembershipGridRow)grid.Rows[i].DataBoundItem;
                        Assert.AreSame(presentation.Rows[i], adapter.AuthoritativeRow);
                        Assert.AreEqual(i, adapter.OriginalRowIndex);
                    }
                    Assert.AreEqual(DataGridViewSelectionMode.FullRowSelect, grid.SelectionMode);
                    Assert.IsTrue(grid.ReadOnly);
                    Assert.AreEqual(4, grid.Columns.Cast<DataGridViewColumn>().Count(c => c.Visible));
                    Assert.IsTrue(grid.Columns.Cast<DataGridViewColumn>().Take(1).All(c => c.AutoSizeMode == DataGridViewAutoSizeColumnMode.Fill));
                    Assert.IsTrue(grid.Columns.Cast<DataGridViewColumn>().Skip(1).All(c => c.AutoSizeMode == DataGridViewAutoSizeColumnMode.None));
                    int defaultWidth = grid.Columns.Cast<DataGridViewColumn>().Where(c => c.Visible).Sum(c => c.Width);
                    Assert.IsTrue(defaultWidth <= grid.ClientSize.Width, "Four default columns fit without horizontal scrolling.");
                    header.Controls.OfType<Panel>().Single().Controls.OfType<CheckBox>().Single().Checked = true;
                    Application.DoEvents();
                    Assert.IsTrue(grid.Columns.Cast<DataGridViewColumn>().All(c => c.Visible));
                    int totalWidth = grid.Columns.Cast<DataGridViewColumn>().Sum(c => c.Width);
                    if (width < 1920) Assert.IsTrue(totalWidth > grid.ClientSize.Width, "Narrow windows must scroll rather than shrink optional columns.");
                    else Assert.IsTrue(totalWidth <= grid.ClientSize.Width, "Optional columns should fit the wide test window.");
                    grid.CurrentCell = grid.Rows[0].Cells["ObjectStatusText"]; Application.DoEvents();
                    Assert.IsTrue(grid.HorizontalScrollingOffset > 0 || width >= 1920);
                    Assert.AreSame(presentation.Rows[0], ((MembershipGridRow)grid.CurrentRow.DataBoundItem).AuthoritativeRow);
                    StringAssert.Contains(grid.Columns["ObjectStatusText"].ToolTipText, "object operational status");
                    grid.CurrentCell = grid.Rows[0].Cells["FriendlyNameText"]; Application.DoEvents();
                    Assert.AreEqual(0, grid.HorizontalScrollingOffset);
                    var legend = header.Controls.OfType<Panel>().Single().Controls.OfType<Label>()
                        .Single(label => label.Name == "MembershipLegend");
                    Assert.IsTrue(TextRenderer.MeasureText(legend.Text, legend.Font).Width <= legend.ClientSize.Width,
                        "Legend must fit without obscuring the evidence and coverage actions.");
                    Assert.AreEqual(0, sourceService.Calls + sourceService.ExecuteCalls);
                    Assert.AreEqual(0, targetService.Calls + targetService.ExecuteCalls);
                }
            });
        }

        [DataTestMethod, TestCategory("D4T8R2B"), DataRow(900), DataRow(1400), DataRow(1920)]
        public void R2BPrimaryContextAndWrappingToolbarsFitWithoutHidingTechnicalAccess(int width)
        {
            OnStaThread(() =>
            {
                var fixture = ExplorerFixture.Saa(); var spy = new FakeOrganizationService(); int captures = 0;
                using (var form = new MembershipResultsForm(fixture.Completed, captureAppSettingEvidence: () => captures++, sourceService: spy, targetService: spy))
                {
                    form.Size = new Size(width, 760); ShowLocally(form); form.PerformLayout(); Application.DoEvents();
                    var controls = FoundationControls(form).ToArray();
                    var header = controls.Single(c => c.Name == "ObjectsSnapshotHeader");
                    foreach (var name in new[] { "ObjectsSolutionTitle", "ObjectsSourceState", "ObjectsTargetState", "ObjectsSnapshotContext", "ObjectsGlobalSummary" })
                    {
                        var label = (Label)controls.Single(c => c.Name == name);
                        Assert.IsTrue(label.Visible); Assert.IsFalse(label.AutoEllipsis);
                        Assert.IsTrue(header.ClientRectangle.Contains(header.RectangleToClient(label.RectangleToScreen(label.ClientRectangle))), name);
                        var measured = TextRenderer.MeasureText(label.Text, label.Font, new Size(label.Width, int.MaxValue), TextFormatFlags.WordBreak | TextFormatFlags.TextBoxControl);
                        Assert.IsTrue(measured.Height <= label.Height, name + " must fit without clipping");
                    }
                    StringAssert.Contains(controls.Single(c => c.Name == "ObjectsSolutionTitle").Text, fixture.Presentation.SolutionUniqueName);
                    StringAssert.Contains(controls.Single(c => c.Name == "ObjectsSourceState").Text, fixture.Presentation.Source.Diagnostics.EnvironmentName);
                    StringAssert.Contains(controls.Single(c => c.Name == "ObjectsTargetState").Text, fixture.Presentation.Target.Diagnostics.EnvironmentName);
                    Assert.AreEqual("Snapshot state: Source Complete | Target Complete", controls.Single(c => c.Name == "ObjectsSnapshotContext").Text);
                    var summary = controls.Single(c => c.Name == "ObjectsGlobalSummary"); StringAssert.Contains(summary.Text, "Needs Review: 33");
                    Assert.IsFalse(summary.Text.Contains("requests=")); Assert.IsFalse(summary.Text.Contains("elapsed="));
                    Assert.IsFalse(controls.Single(c => c.Name == "CompleteMembershipDiagnostics").Visible);
                    foreach (var bar in controls.OfType<ToolStrip>().Where(t => t.Name == "ObjectsWorkspaceBar" || t.Name == "MembershipWorkspaceBar"))
                    {
                        Assert.IsTrue(bar.Visible); Assert.AreEqual(ToolStripLayoutStyle.Flow, bar.LayoutStyle);
                        Assert.IsTrue(((FlowLayoutSettings)bar.LayoutSettings).WrapContents);
                        var items = bar.Items.Cast<ToolStripItem>().Where(i => i.Available).ToArray();
                        foreach (var item in items) { Assert.IsTrue(item.Visible); Assert.IsTrue(bar.ClientRectangle.Contains(item.Bounds), item.Name); }
                        for (int i = 0; i < items.Length; i++) for (int j = i + 1; j < items.Length; j++)
                            Assert.IsFalse(items[i].Bounds.IntersectsWith(items[j].Bounds), "Toolbar items must not overlap.");
                    }
                    Assert.IsTrue(controls.Single(c => c.Name == "ObjectsPartialNotice").Visible);
                    Assert.AreEqual(0, captures + spy.Calls + spy.ExecuteCalls + spy.WriteCalls);
                }
            });
        }

        [TestMethod, TestCategory("D4T8R2B")]
        public void R2BPrimaryHeaderDoesNotInventAnUnavailableSnapshotCompletion()
        {
            OnStaThread(() =>
            {
                var source = MembershipEnvironmentResult.Unavailable("Offline source", "SAA", 0, TimeSpan.Zero, "Not read");
                var target = MembershipEnvironmentResult.Unavailable("Offline target", "SAA", 0, TimeSpan.Zero, "Not read");
                var presentation = new MembershipResultPresenter().Create(source, target);
                using (var form = new MembershipResultsForm(presentation))
                {
                    ShowLocally(form);
                    var state = FoundationControls(form).Single(c => c.Name == "ObjectsSnapshotContext");
                    Assert.AreEqual("Snapshot state: Source Unavailable | Target Unavailable", state.Text);
                    Assert.IsFalse(state.Text.Contains("Complete"));
                    StringAssert.Contains(FoundationControls(form).Single(c => c.Name == "ObjectsGlobalSummary").Text, "Compared components: 0");
                }
            });
        }

        private static System.Collections.Generic.IEnumerable<Control> FoundationControls(Control root)
        {
            foreach (Control child in root.Controls)
            {
                yield return child;
                foreach (var nested in FoundationControls(child)) yield return nested;
            }
        }

        [TestMethod, TestCategory("ConnectionRoleLabels")]
        public void ConnectionRoleLabels_TargetBeforeSourceLoadDoesNotOverwriteSource()
        {
            WithConnectionControl((control, source, target) =>
            {
                var sourceDetail = Connection("Source", 1);
                var targetDetail = Connection("Target", 2);
                control.UpdateConnection(source, sourceDetail, string.Empty, null);
                AssertConnectionLabels(control, "Source: Source", "Target: Not connected");
                Assert.IsFalse(ConnectionField<bool>(control, "sourceLoaded"));

                control.UpdateConnection(target, targetDetail, "AdditionalOrganization", null);

                AssertConnectionLabels(control, "Source: Source", "Target: Target");
                Assert.IsFalse(ConnectionField<bool>(control, "sourceLoaded"));
                Assert.AreSame(source, control.Service);
                Assert.AreSame(sourceDetail, control.ConnectionDetail);
                Assert.AreSame(targetDetail, ConnectionField<ConnectionDetail>(control, "targetConnectionDetail"));
                Assert.AreSame(targetDetail, control.AdditionalConnectionDetails.Single());
                Assert.AreEqual(1, targetDetail.ClientRequests);
                Assert.IsNull(ConnectionField<IOrganizationService>(control, "targetService"),
                    "The offline connection returns no client; the callback must not substitute Source.");
                AssertConnectionCalls(source, target, 0, 0);
            });
        }

        [TestMethod, TestCategory("ConnectionRoleLabels")]
        public void ConnectionRoleLabels_SourceFirstInventoryLoadsAndReloadsKeepRoles()
        {
            WithConnectionControl((control, source, target) =>
            {
                control.UpdateConnection(source, Connection("Source", 1), string.Empty, null);
                LoadConnectionInventory(control, "Source");
                AssertConnectionLabels(control, "Source: Source (1 solutions)", "Target: Not connected");
                ConnectOfflineTarget(control, target, Connection("Target", 2));
                AssertConnectionLabels(control, "Source: Source (1 solutions)", "Target: Target");
                LoadConnectionInventory(control, "Target");
                LoadConnectionInventory(control, "Source");
                LoadConnectionInventory(control, "Target");
                AssertConnectionLabels(control, "Source: Source (1 solutions)", "Target: Target (1 solutions)");
                Assert.AreSame(source, control.Service);
                Assert.AreSame(source, ConnectionField<IOrganizationService>(control, "sourceMembershipService"));
                Assert.AreSame(target, ConnectionField<IOrganizationService>(control, "targetService"));
                Assert.AreSame(target, ConnectionField<IOrganizationService>(control, "targetMembershipService"));
                AssertConnectionCalls(source, target, 2, 2);
            });
        }

        [TestMethod, TestCategory("ConnectionRoleLabels")]
        public void ConnectionRoleLabels_TargetCompletesBeforeEitherInventoryLoads()
        {
            WithConnectionControl((control, source, target) =>
            {
                control.UpdateConnection(source, Connection("Source", 1), string.Empty, null);
                ConnectOfflineTarget(control, target, Connection("Target", 2));
                AssertConnectionLabels(control, "Source: Source", "Target: Target");
                Assert.IsFalse(ConnectionField<bool>(control, "sourceLoaded"));
                Assert.IsFalse(ConnectionField<bool>(control, "targetLoaded"));
                LoadConnectionInventory(control, "Target");
                AssertConnectionLabels(control, "Source: Source", "Target: Target (1 solutions)");
                LoadConnectionInventory(control, "Source");
                AssertConnectionLabels(control, "Source: Source (1 solutions)", "Target: Target (1 solutions)");
                AssertConnectionCalls(source, target, 1, 1);
            });
        }

        [TestMethod, TestCategory("ConnectionRoleLabels")]
        public void ConnectionRoleLabels_RepeatedCallbacksPreserveBothServicesAndNames()
        {
            WithConnectionControl((control, source, target) =>
            {
                var sourceDetail = Connection("Source", 1);
                var targetDetail = Connection("Target", 2);
                control.UpdateConnection(source, sourceDetail, string.Empty, null);
                ConnectOfflineTarget(control, target, targetDetail);
                for (int i = 0; i < 3; i++)
                {
                    control.UpdateConnection(target, targetDetail, "AdditionalOrganization", null);
                    control.UpdateConnection(source, sourceDetail, string.Empty, null);
                    AssertConnectionLabels(control, "Source: Source", "Target: Target");
                    Assert.AreSame(source, control.Service);
                    Assert.AreSame(sourceDetail, control.ConnectionDetail);
                    Assert.AreSame(target, ConnectionField<IOrganizationService>(control, "targetService"));
                }
                Assert.AreEqual(1, targetDetail.ClientRequests,
                    "The host already owns the additional connection; repeated callbacks must not reconnect it.");
                Assert.AreEqual(1, control.AdditionalConnectionDetails.Count);
                AssertConnectionCalls(source, target, 0, 0);
            });
        }

        [TestMethod, TestCategory("ConnectionRoleLabels")]
        public void ConnectionRoleLabels_PrimaryNotificationAfterLoadRefreshesOnlySourceName()
        {
            WithConnectionControl((control, source, target) =>
            {
                var sourceDetail = Connection("Source", 1);
                control.UpdateConnection(source, sourceDetail, string.Empty, null);
                ConnectOfflineTarget(control, target, Connection("Target", 2));
                LoadConnectionInventory(control, "Source");
                sourceDetail.ConnectionName = "Renamed source";
                control.UpdateConnection(source, sourceDetail, string.Empty, null);
                AssertConnectionLabels(control, "Source: Renamed source (1 solutions)", "Target: Target");
                Assert.AreSame(source, control.Service);
                Assert.AreSame(target, ConnectionField<IOrganizationService>(control, "targetService"));
                Assert.IsTrue(ConnectionField<bool>(control, "sourceLoaded"));
                AssertConnectionCalls(source, target, 1, 0);
            });
        }

        [DataTestMethod, TestCategory("ConnectionRoleLabels")]
        [DataRow(null, "Target", "Current XrmToolBox connection", "Target")]
        [DataRow("Source", null, "Source", "Target environment")]
        [DataRow(" ", "\t", "Current XrmToolBox connection", "Target environment")]
        public void ConnectionRoleLabels_UnknownNamesNeverBorrowOppositeRole(string sourceName,
            string targetName, string expectedSource, string expectedTarget)
        {
            WithConnectionControl((control, source, target) =>
            {
                control.UpdateConnection(source, Connection(sourceName, 1), string.Empty, null);
                ConnectOfflineTarget(control, target, Connection(targetName, 2));
                AssertConnectionLabels(control, "Source: " + expectedSource, "Target: " + expectedTarget);
                AssertConnectionCalls(source, target, 0, 0);
            });
        }

        [TestMethod, TestCategory("ConnectionRoleLabels")]
        public void ConnectionRoleLabels_MissingPrimaryDetailRemainsUnknown()
        {
            WithConnectionControl((control, source, target) =>
            {
                control.UpdateConnection(source, null, string.Empty, null);
                ConnectOfflineTarget(control, target, Connection("Target", 2));
                AssertConnectionLabels(control, "Source: Current XrmToolBox connection", "Target: Target");
                Assert.IsNull(control.ConnectionDetail);
                Assert.AreSame(source, control.Service);
                AssertConnectionCalls(source, target, 0, 0);
            });
        }

        [TestMethod, TestCategory("ConnectionRoleLabels")]
        public void ConnectionRoleLabels_DuplicateDisplayNamesDoNotEstablishOwnership()
        {
            WithConnectionControl((control, source, target) =>
            {
                var sourceDetail = Connection("Shared label", 1);
                var targetDetail = Connection("Shared label", 2);
                control.UpdateConnection(source, sourceDetail, string.Empty, null);
                ConnectOfflineTarget(control, target, targetDetail);
                control.UpdateConnection(target, targetDetail, "AdditionalOrganization", null);
                AssertConnectionLabels(control, "Source: Shared label", "Target: Shared label");
                Assert.AreSame(source, control.Service);
                Assert.AreSame(sourceDetail, control.ConnectionDetail);
                Assert.AreSame(targetDetail, control.AdditionalConnectionDetails.Single());
                Assert.AreSame(target, ConnectionField<IOrganizationService>(control, "targetService"));
                AssertConnectionCalls(source, target, 0, 0);
            });
        }

        // Exercise the real MultipleConnectionsPluginControlBase callback and collection notification.
        // Only authenticated client creation is replaced; no saved connection or network is used.
        private sealed class OfflineConnectionDetail : ConnectionDetail
        {
            public int ClientRequests { get; private set; }
            public override CrmServiceClient GetCrmServiceClient(bool forceNewService = false)
            {
                ClientRequests++;
                return null;
            }
        }

        private static OfflineConnectionDetail Connection(string name, int id) => new OfflineConnectionDetail
        {
            ConnectionName = name,
            ConnectionId = new Guid(id, 0, 0, new byte[8])
        };

        private static void WithConnectionControl(Action<SolutionComparerControl, FakeOrganizationService,
            FakeOrganizationService> test) => OnStaThread(() =>
        {
            Func<FakeOrganizationService> service = () => new FakeOrganizationService
            {
                RetrievePage = query =>
                {
                    Assert.AreEqual("solution", query.EntityName);
                    var solution = new Entity("solution", new Guid(3, 0, 0, new byte[8]));
                    solution["uniquename"] = "synthetic_solution";
                    solution["friendlyname"] = "Synthetic solution";
                    solution["version"] = "1.0.0.0";
                    return new EntityCollection(new[] { solution });
                }
            };
            using (var host = new Form())
            using (var control = new SolutionComparerControl())
            {
                host.Controls.Add(control);
                ShowLocally(host);
                test(control, service(), service());
                host.Close();
            }
        });

        private static T ConnectionField<T>(SolutionComparerControl control, string name) =>
            (T)typeof(SolutionComparerControl).GetField(name, BindingFlags.Instance | BindingFlags.NonPublic).GetValue(control);

        private static void ConnectOfflineTarget(SolutionComparerControl control, FakeOrganizationService target,
            OfflineConnectionDetail detail)
        {
            control.UpdateConnection(target, detail, "AdditionalOrganization", null);
            Assert.AreEqual(1, detail.ClientRequests);
            // The real Target callback ran with a deliberately null offline client. Supply the existing
            // private service seam only for local solution-read tests, without changing production routing.
            typeof(SolutionComparerControl).GetField("targetService", BindingFlags.Instance | BindingFlags.NonPublic)
                .SetValue(control, target);
            Application.DoEvents();
        }

        private static void LoadConnectionInventory(SolutionComparerControl control, string side) =>
            typeof(SolutionComparerControl).GetMethod("BtnLoad" + side + "_Click",
                BindingFlags.Instance | BindingFlags.NonPublic).Invoke(control, new object[] { null, EventArgs.Empty });

        private static void AssertConnectionLabels(SolutionComparerControl control, string source, string target)
        {
            Assert.AreEqual(source, ConnectionField<Label>(control, "lblSourceEnv").Text);
            Assert.AreEqual(target, ConnectionField<Label>(control, "lblTargetEnv").Text);
        }

        private static void AssertConnectionCalls(FakeOrganizationService source, FakeOrganizationService target,
            int sourceReads, int targetReads)
        {
            Assert.AreEqual(sourceReads, source.Calls);
            Assert.AreEqual(targetReads, target.Calls);
            Assert.AreEqual(0, source.ExecuteCalls + target.ExecuteCalls);
            Assert.AreEqual(0, source.WriteCalls + target.WriteCalls);
        }

        [TestMethod, TestCategory("D4T8R2D")]
        public void R2DAdvancedStartsCollapsedAndRetainsCompleteDiagnostics()
        {
            OnStaThread(() =>
            {
                var f = ExplorerFixture.Make(new[] { ExplorerFixture.Identity(61, "resource") }, new[] { ExplorerFixture.Identity(61, "resource") });
                using (var form = new MembershipResultsForm(f.Completed))
                {
                    ShowLocally(form); var text = (TextBox)FoundationControls(form).Single(c => c.Name == "CompleteMembershipDiagnostics");
                    var shared = R2DAction(form, "SharedAdvancedDiagnosticsAction"); var technical = R2DAction(form, "AdvancedDiagnosticsAction");
                    Assert.IsFalse(shared.Checked); Assert.IsFalse(technical.Checked); Assert.IsFalse(text.Visible); Assert.AreEqual("", text.Text);
                    shared.PerformClick(); Assert.IsTrue(shared.Checked); Assert.IsTrue(technical.Checked); Assert.IsTrue(text.Visible);
                    string expected = MembershipResultsForm.FormatSummary(f.Presentation.Summary) + Environment.NewLine +
                        "Source (Source fixture) diagnostics: requests=27, elapsed=0.00s, Raw Source Components=1, resolved=1, unsupported=0, unresolved=0, ambiguous=0" + Environment.NewLine +
                        "Target (Target fixture) diagnostics: requests=27, elapsed=0.00s, Raw Target Components=1, resolved=1, unsupported=0, unresolved=0, ambiguous=0" + Environment.NewLine +
                        "Source snapshot timestamp: " + f.Presentation.Source.Snapshot.CapturedAt.ToString("o", System.Globalization.CultureInfo.InvariantCulture) + Environment.NewLine +
                        "Target snapshot timestamp: " + f.Presentation.Target.Snapshot.CapturedAt.ToString("o", System.Globalization.CultureInfo.InvariantCulture);
                    Assert.AreEqual(expected, text.Text); Assert.IsTrue(text.ReadOnly); Assert.AreEqual(ScrollBars.Both, text.ScrollBars);
                    R2DAction(form, "TechnicalMembershipAction").PerformClick(); Application.DoEvents(); Assert.IsTrue(text.Visible); Assert.AreEqual(expected, text.Text);
                    technical.PerformClick(); Assert.IsFalse(text.Visible); Assert.IsFalse(shared.Checked);
                    R2DAction(form, "ObjectsBackAction").PerformClick(); Application.DoEvents(); Assert.IsFalse(text.Visible);
                    shared.PerformClick(); Assert.AreEqual(expected, text.Text);
                    var legacy = form.Controls.OfType<Panel>().Single().Controls.OfType<Label>().Single(x => x.Dock == DockStyle.Fill);
                    Assert.IsFalse(legacy.Visible); StringAssert.Contains(legacy.Text, "Raw Source Components=1");
                }
                var source = MembershipEnvironmentResult.Unavailable("Offline source", "SAA", 0, TimeSpan.Zero, "Not read");
                var target = MembershipEnvironmentResult.Unavailable("Offline target", "SAA", 0, TimeSpan.Zero, "Not read");
                using (var form = new MembershipResultsForm(new MembershipResultPresenter().Create(source, target)))
                {
                    ShowLocally(form); R2DAction(form, "SharedAdvancedDiagnosticsAction").PerformClick();
                    string text = FoundationControls(form).Single(c => c.Name == "CompleteMembershipDiagnostics").Text;
                    StringAssert.Contains(text, "Source snapshot timestamp: [Unavailable — no retained snapshot]");
                    StringAssert.Contains(text, "Target snapshot timestamp: [Unavailable — no retained snapshot]");
                }
            });
        }
        [TestMethod, TestCategory("D4T8R2D")]
        public void R2DAdvancedCoverageRoutesRetainOwnersAndDoNotCapture()
        {
            OnStaThread(() =>
            {
                var f = ExplorerFixture.Make(new[] { ExplorerFixture.Identity(61, "resource") }, new[] { ExplorerFixture.Identity(61, "resource") });
                var spy = new FakeOrganizationService(); int captures = 0; var csv = ExplorerFixture.Csv(f);
                using (var form = new MembershipResultsForm(f.Completed, sourceService: spy, targetService: spy, captureAppSettingEvidence: () => captures++))
                {
#if DEBUG
                    form.CaptureType92Evidence = form.CaptureType59Evidence = form.DiscoverType59Evidence = form.CaptureCloudFlowEvidence = form.CaptureSavedQueryEvidence =
                        form.CaptureUnsupportedCoverageInventory = form.CaptureType36Evidence = form.CaptureType31Evidence = form.CaptureType10072Evidence = form.CaptureType300Evidence =
                        form.CaptureType74Evidence = form.CaptureType511Evidence = form.CaptureType10276Evidence = form.CaptureAppActionEvidence = () => captures++;
#endif
                    ShowLocally(form); R2DAction(form, "SharedAdvancedDiagnosticsAction").PerformClick();
                    foreach (bool technical in new[] { false, true })
                    {
                        if (technical) { R2DAction(form, "TechnicalMembershipAction").PerformClick(); Application.DoEvents(); }
                        MembershipComparisonPresentation observedOwner = null;
                        using (var timer = new System.Windows.Forms.Timer { Interval = 10 })
                        {
                            timer.Tick += (sender, args) =>
                            {
                                var coverage = Application.OpenForms.OfType<MembershipCoverageDetailsForm>().FirstOrDefault();
                                if (coverage == null) return;
                                observedOwner = (MembershipComparisonPresentation)typeof(MembershipCoverageDetailsForm).GetField("presentation", BindingFlags.Instance | BindingFlags.NonPublic).GetValue(coverage);
                                timer.Stop(); coverage.Close();
                            };
                            timer.Start(); R2DAction(form, "SharedCoverageDetailsAction").PerformClick();
                        }
                        Assert.AreSame(f.Presentation, observedOwner); Assert.AreSame(f.Completed, form.CompletedComparison);
                        Assert.AreEqual(0, captures + spy.Calls + spy.ExecuteCalls + spy.WriteCalls);
                    }
                    Assert.AreEqual(csv, ExplorerFixture.Csv(f));
                }
            });
        }
        [TestMethod, TestCategory("D4T8R2D")]
        public void R2DAdvancedTogglePreservesHeaderAndDetailsViewport()
        {
            OnStaThread(() =>
            {
                foreach (var dimensions in new[] { new[] { 900, 100 }, new[] { 900, 200 }, new[] { 1600, 100 } })
                {
                    var f = ExplorerFixture.Make(new[] { ExplorerFixture.Identity(61, "resource") }, new[] { ExplorerFixture.Identity(61, "resource") });
                    var spy = new FakeOrganizationService(); int captures = 0;
                    using (var form = new MembershipResultsForm(f.Completed, sourceService: spy, targetService: spy, captureAppSettingEvidence: () => captures++))
                    {
                        form.Size = new Size(dimensions[0], dimensions[0] == 1600 ? 900 : 520);
                        form.Font = new Font("Segoe UI", 9F * dimensions[1] / 100F); ShowLocally(form);
                        var watch = System.Diagnostics.Stopwatch.StartNew();
                        while (form.ExplorerProjection == null && watch.Elapsed.TotalSeconds < 20) { Application.DoEvents(); Thread.Sleep(1); }
                        Assert.IsNotNull(form.ExplorerProjection);
                        var c = form.Controls.OfType<SolutionObjectsExplorerControl>().Single(); c.SelectComparison(0, DisplayEnvironmentSide.Target); c.ShowDetails(true);
                        var owner = c.SelectedComparison.AuthoritativeRow; string overview = c.SimpleDetailsText;
                        var bar = form.Controls.OfType<ToolStrip>().Single(t => t.Name == "MembershipWorkspaceBar");
                        foreach (var item in bar.Items.Cast<ToolStripItem>().Where(x => x.Available)) Assert.IsTrue(bar.ClientRectangle.Contains(item.Bounds), item.Name);
                        var action = R2DAction(form, "SharedAdvancedDiagnosticsAction"); action.PerformClick(); Application.DoEvents();
                        var text = (TextBox)FoundationControls(form).Single(x => x.Name == "CompleteMembershipDiagnostics");
                        Assert.IsTrue(text.Visible); Assert.IsTrue(text.ReadOnly); Assert.AreEqual(ScrollBars.Both, text.ScrollBars);
                        action.PerformClick(); Application.DoEvents(); Assert.IsFalse(text.Visible);
                        Assert.AreEqual(overview, c.SimpleDetailsText); Assert.AreSame(owner, c.SelectedComparison.AuthoritativeRow);
                        var options = FoundationControls(c).Single(x => x.Name == "ExplorerDetailsOptions");
                        Assert.AreEqual(options, FoundationControls(c).Single(x => x.Name == "ExplorerSupportEvidenceToggle").Parent);
                        Assert.AreEqual(options, FoundationControls(c).Single(x => x.Name == "ExplorerDetailsScrollCue").Parent);
                        Assert.IsFalse(((SplitContainer)FoundationControls(c).Single(x => x.Name == "ExplorerEvidenceSplit")).Panel2Collapsed);
                        Assert.AreEqual(0, captures + spy.Calls + spy.ExecuteCalls + spy.WriteCalls);
                    }
                }
            });
        }
        private static ToolStripButton R2DAction(MembershipResultsForm form, string name) => form.Controls.OfType<ToolStrip>()
            .SelectMany(t => t.Items.OfType<ToolStripButton>()).Single(x => x.Name == name);

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
        #region R2E usability additions
        [TestMethod, TestCategory("D4T8R2E")]
        public void R2ENavigationSequencePreservesRowsRawMappingsAndExportInputs()
        {
            OnStaThread(() =>
            {
                var fixture = R2EUsabilityFixture(); var rows = fixture.Presentation.Rows.ToArray(); var projection = ExplorerFixture.Project(fixture);
                var raw = projection.Occurrences.Select(o => o.Original).ToArray(); var csv = ExplorerFixture.Csv(fixture);
                var files = Enumerable.Range(0, 6).Select(_ => System.IO.Path.GetTempFileName()).ToArray();
                try
                {
                    using (var host = new Form()) using (var solutions = new SolutionComparerControl())
                    using (var form = new MembershipResultsForm(fixture.Completed))
                    {
                        host.Controls.Add(solutions); ShowLocally(host);
                        var grid = (DataGridView)typeof(SolutionComparerControl).GetField("dgvResults", BindingFlags.Instance | BindingFlags.NonPublic).GetValue(solutions);
                        grid.DataSource = new System.Collections.Generic.List<Models.SolutionCompareResult> { new Models.SolutionCompareResult
                        {
                            UniqueName = "synthetic_solution", SourceDisplayName = "Source solution", TargetDisplayName = "Target solution", SourceVersion = "1.0", TargetVersion = "2.0",
                            SourcePublisher = "Source publisher", TargetPublisher = "Target publisher", SourcePackageType = "Unmanaged", TargetPackageType = "Managed",
                            PackageTypeStatus = "Managed/Unmanaged Mismatch", Status = "Multiple Differences"
                        } };
                        string[] exports = { "ExportVisibleRowsToCsv", "ExportVisibleRowsToSpreadsheetMl", "ExportVisibleRowsToXlsxOpenXml" };
                        for (int i = 0; i < exports.Length; i++) typeof(SolutionComparerControl).GetMethod(exports[i], BindingFlags.Instance | BindingFlags.NonPublic).Invoke(solutions, new object[] { files[i] });
                        R2EPrepareForm(form); R2ENavigate(form);
                        for (int i = 0; i < exports.Length; i++) typeof(SolutionComparerControl).GetMethod(exports[i], BindingFlags.Instance | BindingFlags.NonPublic).Invoke(solutions, new object[] { files[i + 3] });
                        for (int i = 0; i < 2; i++) Assert.AreEqual(R2ENormalizeTimestamp(System.IO.File.ReadAllText(files[i])), R2ENormalizeTimestamp(System.IO.File.ReadAllText(files[i + 3])));
                        string before = R2EXlsxSemantics(files[2]); string after = R2EXlsxSemantics(files[5]); Assert.AreEqual(before, after);
                        StringAssert.Contains(after, "synthetic_solution"); StringAssert.Contains(after, "Source publisher"); StringAssert.Contains(after, "Target publisher");
                        Assert.AreEqual(csv, ExplorerFixture.Csv(fixture)); CollectionAssert.AreEqual(rows, fixture.Presentation.Rows.ToArray());
                        Assert.AreEqual(projection.Rows.Count, form.ExplorerProjection.Rows.Count); Assert.AreEqual(raw.Length, form.ExplorerProjection.Occurrences.Count);
                        for (int i = 0; i < raw.Length; i++)
                        {
                            var current = form.ExplorerProjection.Occurrences[i].Original;
                            Assert.AreSame(raw[i].Record, current.Record); Assert.AreSame(raw[i].Snapshot, current.Snapshot); Assert.AreEqual(raw[i].Side, current.Side);
                            CollectionAssert.AreEqual(raw[i].ComparisonRowIndices.ToArray(), current.ComparisonRowIndices.ToArray());
                        }
                        for (int i = 0; i < rows.Length; i++) Assert.AreSame(rows[i], form.ExplorerProjection.Rows[i].Original.AuthoritativeRow);
                    }
                }
                finally { foreach (var file in files) System.IO.File.Delete(file); }
            });
        }
        [TestMethod, TestCategory("D4T8R2E")]
        public void R2ENavigationSequenceKeepsAllServicesAndCaptureDelegatesIdle()
        {
            OnStaThread(() =>
            {
                var fixture = R2EUsabilityFixture(); var source = new FakeOrganizationService(); var target = new FakeOrganizationService(); int captures = 0;
                using (var form = new MembershipResultsForm(fixture.Completed, sourceService: source, targetService: target, captureAppSettingEvidence: () => captures++))
                {
#if DEBUG
                    form.CaptureType92Evidence = form.CaptureType59Evidence = form.DiscoverType59Evidence = form.CaptureCloudFlowEvidence = form.CaptureSavedQueryEvidence =
                        form.CaptureUnsupportedCoverageInventory = form.CaptureType36Evidence = form.CaptureType31Evidence = form.CaptureType10072Evidence = form.CaptureType300Evidence =
                        form.CaptureType74Evidence = form.CaptureType511Evidence = form.CaptureType10276Evidence = form.CaptureAppActionEvidence = () => captures++;
#endif
                    R2EPrepareForm(form); R2ENavigate(form);
                    R2EAction(form, "SharedAdvancedDiagnosticsAction").PerformClick(); R2EAction(form, "SharedAdvancedDiagnosticsAction").PerformClick();
                    form.BeginInvoke(new Action(() =>
                    {
                        var coverage = Application.OpenForms.Cast<Form>().Single(x => x is MembershipCoverageDetailsForm);
                        Assert.IsTrue(coverage.Controls.Count > 0); coverage.Close();
                    }));
                    R2EAction(form, "SharedCoverageDetailsAction").PerformClick();
                    Assert.AreEqual(0, captures + source.Calls + source.ExecuteCalls + source.WriteCalls + target.Calls + target.ExecuteCalls + target.WriteCalls);
                }
            });
        }
        [DataTestMethod, TestCategory("D4T8R2E")]
        [DataRow(900, 100), DataRow(900, 125), DataRow(900, 150), DataRow(900, 200)]
        [DataRow(1100, 100), DataRow(1100, 125), DataRow(1100, 150), DataRow(1100, 200)]
        [DataRow(1400, 100), DataRow(1400, 125), DataRow(1400, 150), DataRow(1400, 200)]
        [DataRow(1600, 100), DataRow(1600, 125), DataRow(1600, 150), DataRow(1600, 200)]
        public void R2ENavigationControlsRemainReachableAtFontScale(int width, int percent)
        {
            OnStaThread(() =>
            {
                using (var form = new MembershipResultsForm(R2EUsabilityFixture().Completed))
                {
                    form.Size = new Size(width, 520); form.Font = new Font("Segoe UI", 9F * percent / 100F); R2EPrepareForm(form);
                    Assert.IsNotNull(form.Controls.OfType<SolutionObjectsExplorerControl>().Single().SelectedComparison, "Exact post-show Objects selection must exist before handoff.");
                    var c = form.Controls.OfType<SolutionObjectsExplorerControl>().Single(); var owner = form.CompletedComparison; var selected = c.SelectedComparison.AuthoritativeRow;
                    var selectedRow = c.SelectedComparison; var selectedSide = c.SelectedSide; var selectedOccurrence = c.SelectedOccurrenceIndex;
                    using (var transitionTrace = new R2ETransitionTrace(form, c, selectedRow, "ObjectsToFlat:" + width + ":" + percent))
                    {
                    Assert.AreSame(c.Projection.Rows[selectedRow.OriginalRowIndex].Original, selectedRow);
                    Assert.AreEqual(DisplayEnvironmentSide.Target, selectedSide); Assert.IsNull(selectedOccurrence);
                    var contextField = typeof(MembershipResultsForm).GetField("retainContextWithoutRow", BindingFlags.Instance | BindingFlags.NonPublic);
                    Assert.IsNotNull(contextField);
                    foreach (bool technical in new[] { false, true })
                    {
                        if (technical)
                        {
                            Assert.IsNotNull(c.SelectedComparison, "Objects selection must exist before workspace handoff.");
                            Assert.AreSame(selectedRow, c.SelectedComparison); Assert.AreSame(selected, c.SelectedComparison.AuthoritativeRow);
                            Assert.AreEqual(selectedSide, c.SelectedSide); Assert.AreEqual(selectedOccurrence, c.SelectedOccurrenceIndex);
                        }
                        if (technical) transitionTrace.Snapshot("checkpoint.before.technical.click");
                        if (technical) { R2EAction(form, "TechnicalMembershipAction").PerformClick(); transitionTrace.Snapshot("checkpoint.after.technical.click"); R2EWait(() => form.IsTechnicalWorkspace); }
                        if (technical)
                        {
                            transitionTrace.Snapshot("checkpoint.after.technical.wait");
                            Assert.AreEqual(0, R2EUiItems(form).OfType<ToolStripComboBox>().Single(item => item.Name == "MembershipViewSelector").SelectedIndex);
                            var flat = form.Controls.OfType<DataGridView>().Single(); Assert.IsTrue(flat.Visible);
                            Assert.IsNotNull(flat.CurrentRow, "Objects to Technical must retain the exact Flat current row.");
                            Assert.IsNotNull(flat.CurrentRow.DataBoundItem, "Flat arrival must retain a bound comparison row.");
                            Assert.IsInstanceOfType(flat.CurrentRow.DataBoundItem, typeof(MembershipGridRow));
                            Assert.AreEqual(selectedRow.OriginalRowIndex, ((MembershipGridRow)flat.CurrentRow.DataBoundItem).OriginalRowIndex);
                            Assert.AreSame(selected, ((MembershipGridRow)flat.CurrentRow.DataBoundItem).AuthoritativeRow);
                            transitionTrace.Snapshot("checkpoint.before.context.assertion");
                            Assert.IsFalse((bool)contextField.GetValue(form), "A retained row must not carry row-less context intent.");
                            Assert.AreEqual(selectedSide, (DisplayEnvironmentSide?)typeof(MembershipResultsForm).GetField("retainedExplorerSide", BindingFlags.Instance | BindingFlags.NonPublic).GetValue(form));
                            Assert.AreEqual(selectedOccurrence, (int?)typeof(MembershipResultsForm).GetField("retainedExplorerOccurrence", BindingFlags.Instance | BindingFlags.NonPublic).GetValue(form));
                        }
                        form.PerformLayout(); Application.DoEvents();
                        foreach (var bar in form.Controls.OfType<ToolStrip>().Where(t => t.Visible))
                            foreach (ToolStripItem item in bar.Items)
                                if (item.Available && !string.IsNullOrEmpty(item.Name))
                                { Assert.IsTrue(item.Visible, item.Name); Assert.IsTrue(bar.ClientRectangle.Contains(item.Bounds), item.Name + " must fit its native wrapped toolbar."); }
                        Assert.IsTrue(R2EAction(form, "MembershipResetFiltersAction").Visible); Assert.IsFalse(string.IsNullOrWhiteSpace(R2EAction(form, "MembershipResetFiltersAction").AccessibleName));
                        Assert.IsTrue(c.Grid.VirtualMode); Assert.AreEqual(DataGridViewAutoSizeRowsMode.None, c.Grid.AutoSizeRowsMode); Assert.AreSame(owner, form.CompletedComparison);
                        Assert.IsTrue(form.ClientSize.Width > 0 && form.ClientSize.Height > 0);
                    }
                    R2EAction(form, "ObjectsBackAction").PerformClick(); R2EWait(() => !form.IsTechnicalWorkspace && c.Visible);
                    Assert.IsNotNull(c.SelectedComparison, "Objects return must preserve the captured row rather than publish a focus-generated clear.");
                    Assert.AreSame(selectedRow, c.SelectedComparison); Assert.AreEqual(selectedSide, c.SelectedSide); Assert.AreEqual(selectedOccurrence, c.SelectedOccurrenceIndex);
                    Assert.IsFalse((bool)contextField.GetValue(form));
                    Assert.AreSame(selected, c.SelectedComparison.AuthoritativeRow);
                    var options = FoundationControls(c).Single(x => x.Name == "ExplorerDetailsOptions");
                    Assert.AreEqual(options, FoundationControls(c).Single(x => x.Name == "ExplorerDetailsScrollCue").Parent);
                    Assert.AreEqual(options, FoundationControls(c).Single(x => x.Name == "ExplorerSupportEvidenceToggle").Parent);
                    Assert.IsTrue(((LinkLabel)FoundationControls(c).Single(x => x.Name == "ExplorerExpandDetails")).TabStop);
                    Assert.IsTrue(R2EUiItems(form).OfType<ToolStripComboBox>().Single(x => x.Name == "ObjectsMembershipPresenceFilter").Control.TabStop);
                    System.Diagnostics.Debug.WriteLine("R2E synthetic font scale " + percent + "%; width=" + width + "; actual DeviceDpi=" + form.DeviceDpi + "; not OS DPI acceptance");
                    }
                }
            });
        }
        internal sealed class R2ETransitionTrace : IDisposable
        {
            private readonly MembershipResultsForm form;
            private readonly SolutionObjectsExplorerControl explorer;
            private readonly MembershipGridRow expected;
            private readonly string caseLabel;
            private readonly DataGridView flat;
            private readonly ToolStripComboBox view;
            private readonly Action originalIntent, intentWrapper;
            private readonly Action<MembershipGridRow> originalSelection, selectionWrapper;
            private readonly EventHandler explorerVisibility, flatVisibility, flatCurrent, flatSelection, viewSelection;
            private readonly TreeViewCancelEventHandler treeBefore;
            private readonly TreeViewEventHandler treeAfter;
            private readonly DataGridViewBindingCompleteEventHandler flatBinding;
            private int sequence;
            internal R2ETransitionTrace(MembershipResultsForm form, SolutionObjectsExplorerControl explorer, MembershipGridRow expected, string caseLabel)
            {
                this.form = form; this.explorer = explorer; this.expected = expected; this.caseLabel = caseLabel;
                flat = form.Controls.OfType<DataGridView>().Single();
                view = form.Controls.OfType<ToolStrip>().SelectMany(bar => bar.Items.OfType<ToolStripComboBox>()).Single(item => item.Name == "MembershipViewSelector");
                originalIntent = explorer.NavigationIntentChanged; originalSelection = explorer.SelectedComparisonChanged;
                intentWrapper = () => { Snapshot("intent.enter"); try { originalIntent?.Invoke(); } finally { Snapshot("intent.exit"); } };
                selectionWrapper = row => { Snapshot("selection.enter argument=" + Row(row)); try { originalSelection?.Invoke(row); } finally { Snapshot("selection.exit argument=" + Row(row)); } };
                explorerVisibility = (sender, args) => Snapshot("explorer.visible");
                flatVisibility = (sender, args) => Snapshot("flat.visible");
                flatCurrent = (sender, args) => Snapshot("flat.current-cell");
                flatSelection = (sender, args) => Snapshot("flat.selection");
                viewSelection = (sender, args) => Snapshot("view.selected-index");
                treeBefore = (sender, args) => Snapshot("tree.before-select action=" + args.Action + " proposed=" + Path(args.Node));
                treeAfter = (sender, args) => Snapshot("tree.after-select action=" + args.Action);
                flatBinding = (sender, args) => Snapshot("flat.binding-complete");
                explorer.NavigationIntentChanged = intentWrapper; explorer.SelectedComparisonChanged = selectionWrapper;
                explorer.VisibleChanged += explorerVisibility; explorer.Navigation.BeforeSelect += treeBefore; explorer.Navigation.AfterSelect += treeAfter;
                flat.VisibleChanged += flatVisibility; flat.CurrentCellChanged += flatCurrent; flat.SelectionChanged += flatSelection;
                flat.DataBindingComplete += flatBinding; view.SelectedIndexChanged += viewSelection;
                Snapshot("trace.attached");
            }
            private static object Read(object value, string name) => value?.GetType().GetField(name, BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic)?.GetValue(value);
            private static string Ref(object value) => value == null ? "null" : System.Runtime.CompilerServices.RuntimeHelpers.GetHashCode(value).ToString(System.Globalization.CultureInfo.InvariantCulture);
            private string Row(MembershipGridRow row) => row == null ? "null" : "index=" + row.OriginalRowIndex + ";wrapper=" + Ref(row) + ";authoritative=" + Ref(row.AuthoritativeRow) + ";expectedWrapper=" + ReferenceEquals(row, expected) + ";expectedAuthoritative=" + ReferenceEquals(row.AuthoritativeRow, expected?.AuthoritativeRow);
            private static string Path(TreeNode node)
            {
                var parts = new System.Collections.Generic.List<string>();
                while (node != null) { parts.Insert(0, node.Name); node = node.Parent; }
                return string.Join("/", parts);
            }
            private static string Quote(string value) => "\"" + string.Concat((value ?? "null").Select(character => char.IsControl(character) ? "\\u" + ((int)character).ToString("x4", System.Globalization.CultureInfo.InvariantCulture) : character == '\\' ? "\\\\" : character == '"' ? "\\\"" : character.ToString())) + "\"";
            internal void Snapshot(string checkpoint)
            {
                var state = new System.Collections.Generic.SortedDictionary<string, string>();
                Action<string, object> put = (name, value) => state.Add(name, value == null ? "null" : Convert.ToString(value, System.Globalization.CultureInfo.InvariantCulture));
                put("case", caseLabel); put("sequence", ++sequence); put("checkpoint", checkpoint);
                put("workspace.technical", form.IsTechnicalWorkspace); put("form.visible", form.Visible); put("form.disposed", form.IsDisposed);
                put("explorer.visible", explorer.Visible); put("explorer.primary", explorer.IsPrimaryWorkspace); put("explorer.scope", explorer.WorkspaceScope);
                put("explorer.row", Row(explorer.SelectedComparison)); put("explorer.side", explorer.SelectedSide); put("explorer.occurrence", explorer.SelectedOccurrenceIndex);
                put("projection.ref", Ref(explorer.Projection)); put("completed.ref", Ref(form.CompletedComparison));
                var hierarchy = Read(form, "hierarchyControl") as Control; put("hierarchy.visible", hierarchy?.Visible);
                foreach (var name in new[] { "retainedSelectionRow", "retainedExplorerSide", "retainedExplorerOccurrence", "retainContextWithoutRow", "selectionIntentRevision", "hierarchyBookmarkRevision", "viewRequestVersion", "restoringNavigation", "navigationTransaction", "switchingWorkspace", "restoreFocusPending", "focusWorkspace" }) put("form." + name, Read(form, name));
                int? retained = (int?)Read(form, "retainedSelectionRow");
                var rows = (System.Collections.IEnumerable)Read(form, "gridRows"); put("form.retainedRowIdentity", Row(retained.HasValue ? rows.Cast<MembershipGridRow>().FirstOrDefault(row => row.OriginalRowIndex == retained.Value) : null));
                foreach (var name in new[] { "restoringNavigation", "preparing", "navigationContextChanged", "technicalContextRevision", "lastNavigationFocus" }) put("explorer." + name, Read(explorer, name));
                var node = explorer.Navigation.SelectedNode; var tag = node?.Tag;
                put("tree.path", Path(node)); put("tree.tagType", tag?.GetType().Name); put("tree.tagRef", Ref(tag));
                foreach (var name in new[] { "Side", "Occurrence", "Description", "Technical", "Start", "Count" }) put("tree.tag." + name, Read(tag, name));
                var tagRows = Read(tag, "Rows") as System.Collections.IEnumerable; put("tree.tag.Rows", tagRows == null ? null : string.Join(",", tagRows.Cast<object>()));
                put("view.index", view.SelectedIndex); put("flat.visible", flat.Visible); put("flat.currentCell", flat.CurrentCell == null ? "null" : flat.CurrentCell.RowIndex + ":" + flat.CurrentCell.ColumnIndex);
                put("flat.currentRowIndex", flat.CurrentRow?.Index); put("flat.bindingType", flat.CurrentRow?.DataBoundItem?.GetType().Name); put("flat.row", Row(flat.CurrentRow?.DataBoundItem as MembershipGridRow));
                var bookmarks = (System.Collections.IDictionary)Read(form, "explorerBookmarks"); var saved = new System.Collections.Generic.List<string>();
                foreach (System.Collections.DictionaryEntry entry in bookmarks)
                {
                    var bookmark = entry.Value;
                    saved.Add(entry.Key + ";owner=" + Ref(Read(bookmark, "Owner")) + ";row=" + Read(bookmark, "Row") + ";side=" + Read(bookmark, "Side") + ";occurrence=" + Read(bookmark, "Occurrence") + ";path=" + Read(bookmark, "SelectedPath") + ";focus=" + Read(bookmark, "Focus"));
                }
                put("form.bookmarks", string.Join(" | ", saved.OrderBy(item => item, StringComparer.Ordinal)));
                put("stack", new System.Diagnostics.StackTrace(1, false).ToString());
                Console.WriteLine("R2E_TRACE " + "{" + string.Join(",", state.Select(entry => Quote(entry.Key) + ":" + Quote(entry.Value))) + "}");
            }
            public void Dispose()
            {
                Snapshot("trace.detaching");
                if (ReferenceEquals(explorer.NavigationIntentChanged, intentWrapper)) explorer.NavigationIntentChanged = originalIntent;
                if (ReferenceEquals(explorer.SelectedComparisonChanged, selectionWrapper)) explorer.SelectedComparisonChanged = originalSelection;
                explorer.VisibleChanged -= explorerVisibility; explorer.Navigation.BeforeSelect -= treeBefore; explorer.Navigation.AfterSelect -= treeAfter;
                flat.VisibleChanged -= flatVisibility; flat.CurrentCellChanged -= flatCurrent; flat.SelectionChanged -= flatSelection;
                flat.DataBindingComplete -= flatBinding; view.SelectedIndexChanged -= viewSelection;
            }
        }
        private static ExplorerFixture R2EUsabilityFixture()
        {
            var ids = Enumerable.Range(0, 6).Select(i => ExplorerFixture.Identity(61, "resource" + i)).ToArray(); var f = ExplorerFixture.Make(ids, ids);
            var p = new MembershipResultPresenter().Create(
                MembershipEnvironmentResult.FromSnapshot("Source environment with a retained long display name 1234567890", f.Presentation.Source.Snapshot, 27, TimeSpan.Zero),
                MembershipEnvironmentResult.FromSnapshot("Target environment with a retained long display name 1234567890", f.Presentation.Target.Snapshot, 27, TimeSpan.Zero));
            f.Presentation = p; f.Completed = new CompletedMembershipComparison(p, p.Rows.Select((row, index) => new ComponentDisplayEvidence(index,
                ExplorerFixture.Display(row.SourceIdentity, DisplayEnvironmentSide.Source, friendly: new ComponentDisplayValue(DisplayAvailability.Available, "Source component with a retained long display name 1234567890", "Synthetic")),
                ExplorerFixture.Display(row.TargetIdentity, DisplayEnvironmentSide.Target, friendly: new ComponentDisplayValue(DisplayAvailability.Available, "Target component with a retained long display name 1234567890", "Synthetic")))),
                DisplayProjectionOutcome.Completed, DisplayCacheAssociation.Verified, DisplayCacheAssociation.Verified); return f;
        }
        private static void R2EPrepareForm(MembershipResultsForm form)
        {
            ShowLocally(form); R2EWait(() => form.Controls.OfType<SolutionObjectsExplorerControl>().Any(control => control.Visible));
            var c = form.Controls.OfType<SolutionObjectsExplorerControl>().Single(); int calls = 0;
            var field = typeof(SolutionObjectsExplorerControl).GetField("describe", BindingFlags.Instance | BindingFlags.NonPublic);
            var formatter = (Func<MembershipGridRow, string>)field.GetValue(c); field.SetValue(c, (Func<MembershipGridRow, string>)(row => { calls++; return formatter(row); }));
            c.SelectComparison(0, DisplayEnvironmentSide.Target); Application.DoEvents(); Assert.AreEqual(0, c.SelectedComparison.OriginalRowIndex);
            Assert.AreEqual(DisplayEnvironmentSide.Target, c.SelectedSide); Assert.IsNull(c.SelectedOccurrenceIndex); Assert.AreEqual(0, calls);
        }
        private static void R2ENavigate(MembershipResultsForm form)
        {
            var c = form.Controls.OfType<SolutionObjectsExplorerControl>().Single();
            var scope = R2EUiItems(form).OfType<ToolStripComboBox>().Single(x => x.Name == "ObjectsScopeSelector");
            var search = R2EUiItems(form).OfType<ToolStripTextBox>().Single(x => x.Name == "ObjectsSearchFilter");
            var mask = R2EUiItems(form).OfType<ToolStripComboBox>().Single(x => x.Name == "ObjectsMembershipPresenceFilter");
            for (int i = 0; i < 3; i++) { scope.SelectedIndex = i; R2EWait(() => c.WorkspaceScope == (ObjectsWorkspaceScope)i && c.Visible); }
            search.Text = "no retained match"; Assert.AreEqual(0, c.VisibleRows.Count); R2EAction(form, "ObjectsClearSearchAction").PerformClick();
            mask.SelectedIndex = 1; R2EAction(form, "MembershipResetFiltersAction").PerformClick(); R2EWait(() => c.Visible);
            R2EAction(form, "TechnicalMembershipAction").PerformClick(); var view = R2EUiItems(form).OfType<ToolStripComboBox>().Single(x => x.Name == "MembershipViewSelector");
            for (int i = 0; i < 4; i++)
            {
                view.SelectedIndex = i;
                R2EWait(() => i == 1 ? form.Controls.OfType<MembershipHierarchyControl>().Any(h => h.Visible) : i == 3 ? c.Visible : form.Controls.OfType<DataGridView>().Any(g => g.Visible));
            }
            R2EAction(form, "ObjectsBackAction").PerformClick(); R2EWait(() => !form.IsTechnicalWorkspace && c.Visible);
            c.SelectComparison(0, DisplayEnvironmentSide.Target); R2EAction(form, "ObjectsDetailsAction").Checked = true;
            var support = (CheckBox)FoundationControls(c).Single(x => x.Name == "ExplorerSupportEvidenceToggle"); support.Checked = true; var evidence = c.EvidenceText;
            support.Checked = false; Assert.IsFalse(string.IsNullOrEmpty(evidence));
            R2EAction(form, "ObjectsOriginalRowAction").PerformClick(); R2EAction(form, "ObjectsBackAction").PerformClick(); R2EWait(() => !form.IsTechnicalWorkspace && c.Visible);
        }
        private static System.Collections.Generic.IEnumerable<ToolStripItem> R2EUiItems(Form form) => FoundationControls(form).OfType<ToolStrip>().SelectMany(t => t.Items.Cast<ToolStripItem>());
        private static ToolStripButton R2EAction(Form form, string name) => R2EUiItems(form).OfType<ToolStripButton>().Single(x => x.Name == name);
        private static void R2EWait(Func<bool> ready)
        {
            var watch = System.Diagnostics.Stopwatch.StartNew(); while (!ready() && watch.Elapsed.TotalSeconds < 20) { Application.DoEvents(); Thread.Sleep(1); }
            Application.DoEvents(); Assert.IsTrue(ready(), "R2E local binding did not finish within the existing UI budget.");
        }
        private static string R2ENormalizeTimestamp(string text) => System.Text.RegularExpressions.Regex.Replace(text, @"\d{4}-\d{2}-\d{2} \d{2}:\d{2}", "export timestamp");
        private static object R2EProperty(object value, string name) => value?.GetType().GetProperty(name).GetValue(value);
        private static string R2EXlsxSemantics(string path)
        {
            var type = Assembly.Load("DocumentFormat.OpenXml").GetType("DocumentFormat.OpenXml.Packaging.SpreadsheetDocument", true);
            using (var document = (IDisposable)type.GetMethod("Open", new[] { typeof(string), typeof(bool) }).Invoke(null, new object[] { path, false }))
            {
                var part = R2EProperty(document, "WorkbookPart"); var result = new System.Collections.Generic.List<string>();
                var workbook = new System.Xml.XmlDocument(); workbook.LoadXml((string)R2EProperty(R2EProperty(part, "Workbook"), "OuterXml"));
                foreach (System.Xml.XmlNode sheet in workbook.GetElementsByTagName("sheet", "http://schemas.openxmlformats.org/spreadsheetml/2006/main"))
                    result.Add("Sheet:" + sheet.Attributes["name"].Value + ":" + sheet.Attributes["sheetId"].Value);
                var worksheets = ((System.Collections.IEnumerable)R2EProperty(part, "WorksheetParts")).Cast<object>().OrderBy(w => Convert.ToString(R2EProperty(w, "Uri"))).ToArray();
                Assert.AreEqual(1, worksheets.Length);
                foreach (var worksheet in worksheets) result.Add((string)R2EProperty(R2EProperty(worksheet, "Worksheet"), "OuterXml"));
                result.Add((string)R2EProperty(R2EProperty(R2EProperty(part, "WorkbookStylesPart"), "Stylesheet"), "OuterXml"));
                var strings = R2EProperty(part, "SharedStringTablePart"); if (strings != null) result.Add((string)R2EProperty(R2EProperty(strings, "SharedStringTable"), "OuterXml"));
                return R2ENormalizeTimestamp(string.Join(Environment.NewLine, result));
            }
        }
        #endregion
    }
}
