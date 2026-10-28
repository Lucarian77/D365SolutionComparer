using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.ExceptionServices;
using System.Threading;
using System.Windows.Forms;
using D365SolutionComparer.Models.ComponentDetails;
using D365SolutionComparer.Models.Identity;
using D365SolutionComparer.Models.Membership;
using D365SolutionComparer.Models.Presentation;
using D365SolutionComparer.Services.Membership;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace D365SolutionComparer.Tests
{
    [TestClass, TestCategory("H2AHierarchyUi")]
    public class MembershipHierarchyUiTests
    {
        public TestContext TestContext { get; set; }

        [TestMethod]
        public void ExplicitFlatRetainsOriginalGridAndProjectionIsOncePerCompletedComparison()
        {
            Sta(() =>
            {
                var f = Fixture(); var service = new FakeOrganizationService(); int projections = 0;
                using (var form = new MembershipResultsForm(f.Completed, sourceService: service, targetService: service))
                {
                    ViewSelector(form).SelectedIndex = 0; form.Show(); Application.DoEvents(); var grid = Grid(form); var data = grid.DataSource;
                    Assert.AreEqual("Flat View", ViewSelector(form).SelectedItem); Assert.IsTrue(grid.Visible);
                    Assert.IsNull(form.HierarchyProjection);
                    form.ProjectHierarchy = token => { Interlocked.Increment(ref projections); return ComponentHierarchyProjector.Project(f.Completed, cancellationToken: token); };
                    SwitchHierarchy(form); var projection = form.HierarchyProjection;
                    ViewSelector(form).SelectedIndex = 0; Assert.IsTrue(grid.Visible); Assert.AreSame(data, grid.DataSource);
                    SwitchHierarchy(form); Assert.AreSame(projection, form.HierarchyProjection); Assert.AreEqual(1, projections);
                    CollectionAssert.AreEqual(new[] { "Component", "Type", "Membership Status", "Definition Status" },
                        grid.Columns.Cast<DataGridViewColumn>().Where(c => c.Visible).Select(c => c.HeaderText).ToArray());
                    Assert.AreEqual(0, service.Calls + service.ExecuteCalls + service.WriteCalls);
                    Assert.AreEqual(27, f.Presentation.Source.Diagnostics.RequestCount);
                }
            });
        }

        [DataTestMethod]
        [DataRow(2), DataRow(10), DataRow(26), DataRow(59), DataRow(60)]
        public void RootLinkedSubcomponentsAreLabelledInclusionNotOwnership(int type)
        {
            var parent = Identity(1, 1, "account"); var child = Identity(2, type, "child", root: parent.Record.SolutionComponentId);
            var f = Make(new[] { parent, child }); var model = Model(f);
            var parentNode = Walk(model.Root).Single(n => n.Occurrence?.Record.SolutionComponentId == parent.Record.SolutionComponentId);
            var group = parentNode.Children.Single(n => n.Text.StartsWith("Inclusion-root links", StringComparison.Ordinal));
            Assert.IsTrue(group.Text.Contains("not structural ownership")); Assert.AreSame(child.Record, group.Children.Single().Occurrence.Record);
            Assert.IsTrue(Walk(model.Root).Any(n => n.Text == "Parent Not Established"));
            Assert.AreEqual(HierarchyPlacement.ParentNotEstablished, parentNode.Occurrence.Placement);
        }

        [TestMethod]
        public void VerifiedColumnMetadataUsesSeparateStructuralBranchWithoutKeyParsing()
        {
            var table = Identity(1, 1, "account"); var column = Identity(2, 2, "unparseable-column-key");
            var f = Make(new[] { table, column }, parentScope: "account"); var model = Model(f);
            var group = Walk(model.Root).Single(n => n.Text == "Verified structural parent links");
            Assert.AreSame(column.Record, group.Children.Single().Occurrence.Record);
            Assert.AreEqual(HierarchyPlacement.VerifiedStructuralParent, group.Children.Single().Occurrence.Placement);
        }

        [DataTestMethod]
        [DataRow("missing"), DataRow("self"), DataRow("cycle"), DataRow("duplicate")]
        public void InvalidRootEvidenceRemainsUnknownAndNeverGuessesParent(string mode)
        {
            var a = Identity(1, 1, "account", root: mode == "self" ? Id(1) : mode == "missing" ? Id(999) : Id(2));
            var b = Identity(2, 1, "contact", root: mode == "cycle" ? Id(1) : (Guid?)null);
            var list = new List<ComponentIdentity> { a, b }; if (mode == "duplicate") list.Add(Identity(2, 1, "duplicate"));
            var model = Model(Make(list)); var node = Walk(model.Root).Single(n => ReferenceEquals(n.Occurrence?.Identity, a));
            Assert.AreEqual(HierarchyPlacement.ParentNotEstablished, node.Occurrence.Placement);
            StringAssert.Contains(MembershipHierarchyView.Describe(node), "not top-level proof");
            Assert.IsFalse(node.Children.Any(n => n.Text.StartsWith("Inclusion-root links", StringComparison.Ordinal)));
        }

        [TestMethod]
        public void SourceTargetLocalParentsAndDifferentNamesRemainSeparate()
        {
            var sourceParent = Identity(1, 1, "account"); var targetParent = Identity(11, 1, "contact");
            var sourceChild = Identity(2, 2, "column", root: Id(1)); var targetChild = Identity(12, 2, "column", root: Id(11));
            var f = Make(new[] { sourceParent, sourceChild }, new[] { targetParent, targetChild }); var model = Model(f);
            var occurrences = Walk(model.Root).Where(n => n.Occurrence != null).ToList();
            Assert.AreEqual(4, occurrences.Count); Assert.AreEqual(2, occurrences.Count(n => n.Occurrence.Side == DisplayEnvironmentSide.Source));
            var a = occurrences.Single(n => ReferenceEquals(n.Occurrence.Identity, sourceChild));
            var b = occurrences.Single(n => ReferenceEquals(n.Occurrence.Identity, targetChild));
            Assert.AreNotEqual(a.Occurrence.Associations.Single().ReferencedLocalId, b.Occurrence.Associations.Single().ReferencedLocalId);
            Assert.AreSame(a.Row.AuthoritativeRow, b.Row.AuthoritativeRow); Assert.AreNotEqual(a.Text, b.Text);
        }

        [DataTestMethod]
        [DataRow(IdentityResolutionStatus.Unsupported), DataRow(IdentityResolutionStatus.Unresolved), DataRow(IdentityResolutionStatus.Ambiguous)]
        public void VerifiedParentCannotUpgradeChildClassification(IdentityResolutionStatus status)
        {
            var parent = Identity(1, 1, "account"); var child = Identity(2, 300, null, root: Id(1), status: status);
            var f = Make(new[] { parent, child }); var model = Model(f);
            var node = Walk(model.Root).Single(n => n.Occurrence?.Identity == child);
            Assert.AreEqual(status, node.Occurrence.Identity.Status);
            Assert.AreEqual(HierarchyPlacement.VerifiedInclusionRoot, node.Occurrence.Placement);
            Assert.AreSame(f.Presentation.Rows.Single(r => r.SourceIdentity?.Record == child.Record), node.Row.AuthoritativeRow);
            StringAssert.Contains(node.Text, "Indeterminate");
        }

        [TestMethod]
        public void SourceOnlyFilteringKeepsAncestorsAsContextAndCountsOnlyAuthoritativeRows()
        {
            var f = Fixture(); var p = ComponentHierarchyProjector.Project(f.Completed);
            var rows = MembershipGridRow.CreateRows(f.Presentation, f.Completed);
            var selected = rows.Where(r => r.AuthoritativeRow.Presence == MembershipPresence.OnlyInSource).Select(r => r.OriginalRowIndex).ToList();
            var model = MembershipHierarchyView.Create(p, rows, selected, false);
            Assert.AreEqual(selected.Count, model.ComparisonRowCount);
            var parent = Walk(model.Root).Single(n => n.Occurrence?.Record.ComponentType == 1);
            Assert.IsTrue(parent.NavigationContext); Assert.IsNull(parent.Row);
            Assert.IsTrue(Walk(model.Root).Any(n => n.Occurrence?.Record.ComponentType == 2 && !n.NavigationContext));
            Assert.IsFalse(Walk(model.Root).Any(n => n.Occurrence?.Record.ComponentType == 300));
            Assert.AreEqual(840, Baseline().Summary.TotalComparisonRows);
        }

        [TestMethod]
        public void FilterPersistsAcrossViewSwitchesAndSharedComparisonCountIsNotOccurrenceCount()
        {
            Sta(() =>
            {
                var f = Fixture(); using (var form = new MembershipResultsForm(f.Completed))
                {
                    form.Show(); Filter(form).SelectedIndex = 2; int count = Grid(form).Rows.Count;
                    SwitchHierarchy(form); var model = CurrentViewModel(form); Assert.AreEqual(count, model.ComparisonRowCount);
                    ViewSelector(form).SelectedIndex = 0; Assert.AreEqual(2, Filter(form).SelectedIndex); Assert.AreEqual(count, Grid(form).Rows.Count);
                    SwitchHierarchy(form); Assert.AreEqual(2, Filter(form).SelectedIndex);
                    Assert.AreSame(f.Presentation, form.HierarchyProjection.Presentation);
                }
            });
        }

        [DataTestMethod]
        [DataRow(60), DataRow(91), DataRow(92)]
        public void CollapsedComparisonRowsStillHaveEveryRepeatedOccurrence(int type)
        {
            var one = Identity(1, type, "key", objectId: Id(500)); var two = Identity(2, type, "key", objectId: Id(500));
            var f = Make(new[] { one, two }); var model = Model(f);
            Assert.AreEqual(2, Walk(model.Root).Count(n => n.Occurrence != null)); Assert.AreEqual(1, model.ComparisonRowCount);
            var links = Walk(model.Root).Where(n => n.Occurrence != null).Select(n => n.Row.AuthoritativeRow).ToList();
            Assert.AreSame(links[0], links[1]);
        }

        [TestMethod]
        public void MultipleRowsAndUnmappedOccurrencesAreAccessibleWithoutInventedClassifications()
        {
            var identity = Identity(1, 61, "web"); var unmapped = Identity(2, 300, null, status: IdentityResolutionStatus.Unsupported);
            var original = Make(new[] { identity });
            var rows = new[] { original.Presentation.Rows[0], original.Presentation.Rows[0] };
            var p = new MembershipComparisonPresentation("sample", Env(new[] { identity, unmapped }, false), Env(new ComponentIdentity[0], true), rows);
            var f = Enrich(p); var model = Model(f); var node = Walk(model.Root).Single(n => n.Occurrence?.Identity == identity);
            Assert.IsNull(node.Row); Assert.AreEqual(2, node.Children.Count); Assert.AreSame(f.Presentation.Rows[0], node.Children[0].Row.AuthoritativeRow);
            var unknown = Walk(model.Root).Single(n => n.Occurrence?.Identity == unmapped);
            StringAssert.Contains(unknown.Text, "Comparison mapping unavailable"); Assert.IsNull(unknown.Row);
            Assert.AreEqual(2, model.ComparisonRowCount);
        }

        [TestMethod]
        public void UnmappedComparisonRowHasDedicatedNavigationEntry()
        {
            var f = Fixture(); var p = new MembershipComparisonPresentation("sample", Env(new ComponentIdentity[0], false),
                Env(new ComponentIdentity[0], true), f.Presentation.Rows);
            var model = Model(Enrich(p)); Assert.AreEqual(p.Rows.Count, model.ComparisonRowCount);
            Assert.AreEqual(0, model.VisibleOccurrenceCount);
            Assert.IsTrue(Walk(model.Root).Any(n => n.Text == "Comparison rows without retained raw mapping"));
            Assert.AreEqual(p.Rows.Count, Walk(model.Root).Count(n => n.Row != null));
        }

        [TestMethod]
        public void AppSettingHandoffIsDeferredAndNormalUnknownPlacementIsHonest()
        {
            var app = Identity(1, 80, "app"); var setting = Identity(2, 10075, "composite");
            var model = Model(Make(new[] { app, setting })); var node = Walk(model.Root).Single(n => n.Occurrence?.Identity == setting);
            Assert.AreEqual(HierarchyPlacement.ParentNotEstablished, node.Occurrence.Placement);
            Assert.IsFalse(Walk(model.Root).Any(n => n.Text == "Verified structural parent links"));
        }

        [TestMethod]
        public void ConflictingVerifiedParentsAreRetainedWithoutArbitraryAttachment()
        {
            var a = Identity(1, 1, "account"); var b = Identity(2, 1, "contact"); var c = Identity(3, 2, "column", root: Id(1));
            var f = Make(new[] { a, b, c }, parentScope: "contact"); var model = Model(f);
            Assert.IsTrue(Walk(model.Root).Any(n => n.Text == "Multiple verified associations — no single display parent"));
            var node = Walk(model.Root).Single(n => n.Occurrence?.Identity == c);
            Assert.IsTrue(node.Occurrence.HasDifferentParentAssociations);
            StringAssert.Contains(MembershipHierarchyView.Describe(node), "no single parent selected");
        }

        [TestMethod]
        public void VerifiedNamesUseRetainedEvidenceOtherwiseHonestTechnicalFallback()
        {
            var f = Fixture(); var model = Model(f);
            StringAssert.StartsWith(Walk(model.Root).Single(n => n.Occurrence?.Record.ComponentType == 1 && n.Occurrence.Side == DisplayEnvironmentSide.Source).Text, "Account");
            var missing = Model(Enrich(f.Presentation, names: false));
            Assert.IsTrue(Walk(missing.Root).Where(n => n.Occurrence != null).All(n => n.Text.Contains("Name unavailable")));
            Assert.IsFalse(Walk(missing.Root).Where(n => n.Occurrence != null).Any(n => n.Text.Contains("opaque")));
        }

        [TestMethod]
        public void HierarchyDetailsReuseOriginalDisplayAndDefinitionEvidence()
        {
            Sta(() =>
            {
                var f = Fixture(); using (var form = new MembershipResultsForm(f.Completed))
                using (var timer = new System.Windows.Forms.Timer { Interval = 20 })
                {
                    form.Show(); SwitchHierarchy(form); var control = Control(form); ExpandAllBounded(control.Tree, 100);
                    var node = TreeNodes(control.Tree).First(n => (n.Tag as MembershipHierarchyNode)?.Occurrence?.Record.ComponentType == 2);
                    control.Tree.SelectedNode = node; Application.DoEvents();
                    var selected = control.SelectedComparison; Assert.IsNotNull(selected);
                    var text = control.Controls.OfType<SplitContainer>().Single().Panel2.Controls.OfType<TextBox>().Single().Text;
                    StringAssert.Contains(text, "SolutionComponentId (local audit)"); StringAssert.Contains(text, "Membership status: " + selected.MembershipStatus);
                    StringAssert.Contains(text, "not schema ownership"); StringAssert.Contains(text, "Original comparison row index: " + selected.OriginalRowIndex);
                    bool opened = false;
                    timer.Tick += (s, e) =>
                    {
                        var dialog = Application.OpenForms.Cast<Form>().OfType<ComponentDefinitionDetailsForm>().FirstOrDefault();
                        if (dialog == null) return; opened = true; timer.Stop(); dialog.Close();
                    };
                    timer.Start(); var link = control.Controls.OfType<SplitContainer>().Single().Panel2.Controls.OfType<LinkLabel>().Single();
                    typeof(LinkLabel).GetMethod("OnLinkClicked", BindingFlags.Instance | BindingFlags.NonPublic)
                        .Invoke(link, new object[] { new LinkLabelLinkClickedEventArgs(link.Links[0]) });
                    Assert.IsTrue(opened); Assert.AreSame(f.Presentation.Rows[selected.OriginalRowIndex].DefinitionDetail, selected.AuthoritativeRow.DefinitionDetail);
                }
            });
        }

        [DataTestMethod]
        [DataRow("exception"), DataRow("cancelled"), DataRow("failed")]
        public void HierarchyFailureReturnsToFlatKeepsFilterAndSanitizesDiagnostics(string mode)
        {
            Sta(() =>
            {
                var f = Fixture(); using (var form = new MembershipResultsForm(f.Completed))
                {
                    ViewSelector(form).SelectedIndex = 0; form.Show(); Filter(form).SelectedIndex = 2;
                    form.ProjectHierarchy = token => mode == "exception" ? throw new InvalidOperationException("secret URL payload") :
                        ComponentHierarchyProjector.ProjectCore(f.Presentation, f.Completed, null, null,
                            mode == "cancelled" ? new CancellationToken(true) : token, _ => { if (mode == "failed") throw new InvalidOperationException("secret"); });
                    ViewSelector(form).SelectedIndex = 1;
                    Pump(() => form.HierarchyDiagnostics.Count > 0);
                    Assert.AreEqual(0, ViewSelector(form).SelectedIndex); Assert.IsTrue(Grid(form).Visible); Assert.AreEqual(2, Filter(form).SelectedIndex);
                    Assert.IsFalse(string.Join("", form.HierarchyDiagnostics).Contains("secret"));
                    Assert.AreEqual(f.Presentation.Rows.Count(r => r.Presence == MembershipPresence.OnlyInSource), Grid(form).Rows.Count);
                }
            });
        }

        [TestMethod]
        public void RenderingFailureAlsoReturnsToFlatWithoutPartialTreeOrRetrieval()
        {
            Sta(() =>
            {
                var f = Fixture(); var spy = new FakeOrganizationService(); using (var form = new MembershipResultsForm(f.Completed, sourceService: spy, targetService: spy))
                {
                    form.Show(); SwitchHierarchy(form);
                    Control(form).RenderingFailed(new InvalidOperationException("secret"));
                    Assert.AreEqual(0, ViewSelector(form).SelectedIndex); Assert.IsTrue(Grid(form).Visible);
                    Assert.IsFalse(string.Join("", form.HierarchyDiagnostics).Contains("secret"));
                    Assert.AreEqual(0, spy.Calls + spy.ExecuteCalls + spy.WriteCalls);
                }
            });
        }

        [TestMethod]
        public void ClosingDuringProjectionCancelsLocalWorkWithoutTouchingResults()
        {
            Sta(() =>
            {
                var f = Fixture(); using (var started = new ManualResetEventSlim()) using (var exited = new ManualResetEventSlim())
                {
                    var form = new MembershipResultsForm(f.Completed); ViewSelector(form).SelectedIndex = 0;
                    form.ProjectHierarchy = token => { started.Set(); token.WaitHandle.WaitOne(); exited.Set(); token.ThrowIfCancellationRequested(); return null; };
                    bool shown = false; form.Shown += (sender, args) => shown = true; form.Show(); Pump(() => shown);
                    ViewSelector(form).SelectedIndex = 1; Assert.IsTrue(started.Wait(5000)); form.Close(); form.Dispose();
                    Assert.IsTrue(exited.Wait(5000)); Application.DoEvents(); Assert.AreEqual(3, f.Presentation.Rows.Count);
                }
            });
        }

        [TestMethod]
        public void RapidViewSwitchDiscardsStaleRenderAndReusesCompletedProjection()
        {
            Sta(() =>
            {
                var f = Fixture(); using (var gate = new ManualResetEventSlim()) using (var started = new ManualResetEventSlim())
                using (var form = new MembershipResultsForm(f.Completed))
                {
                    ViewSelector(form).SelectedIndex = 0;
                    form.ProjectHierarchy = token => { started.Set(); gate.Wait(token); return ComponentHierarchyProjector.Project(f.Completed, cancellationToken: token); };
                    bool shown = false; form.Shown += (sender, args) => shown = true; form.Show(); Pump(() => shown);
                    ViewSelector(form).SelectedIndex = 1; Assert.IsTrue(started.Wait(5000)); ViewSelector(form).SelectedIndex = 0; gate.Set();
                    var watch = Stopwatch.StartNew(); while (watch.ElapsedMilliseconds < 100) { Application.DoEvents(); Thread.Sleep(1); }
                    Assert.IsTrue(Grid(form).Visible); Assert.IsFalse(form.Controls.OfType<MembershipHierarchyControl>().Any(c => c.Visible));
                    SwitchHierarchy(form); Assert.IsTrue(Control(form).Visible);
                }
            });
        }

        [TestMethod]
        public void Current840BaselineAndCsvAreUntouchedByViewFiltering()
        {
            var f = Enrich(Baseline()); var p = f.Presentation; var csv = new MembershipCoverageCsvExporter().CreateCsv(p);
            var projection = ComponentHierarchyProjector.Project(f.Completed); var rows = MembershipGridRow.CreateRows(p, f.Completed);
            foreach (var presence in new[] { MembershipPresence.PresentInBoth, MembershipPresence.OnlyInSource, MembershipPresence.OnlyInTarget, MembershipPresence.Indeterminate })
            {
                var selected = rows.Where(r => r.AuthoritativeRow.Presence == presence).Select(r => r.OriginalRowIndex).ToList();
                Assert.AreEqual(selected.Count, MembershipHierarchyView.Create(projection, rows, selected, false).ComparisonRowCount);
            }
            var s = p.Summary; Assert.AreEqual(840, s.TotalComparisonRows); Assert.AreEqual(672, s.PresentInBoth);
            Assert.AreEqual(103, s.SourceOnly); Assert.AreEqual(32, s.TargetOnly); Assert.AreEqual(26, s.IndeterminateUnsupported);
            Assert.AreEqual(4, s.IndeterminateAmbiguous); Assert.AreEqual(3, s.IndeterminateResolved);
            Assert.AreEqual(csv, new MembershipCoverageCsvExporter().CreateCsv(p));
        }

        [TestMethod]
        public void BindingIsExplicitlyMarshalledEvenWithoutAmbientUiSynchronizationContext()
        {
            Sta(() =>
            {
                bool previous = System.Windows.Forms.Control.CheckForIllegalCrossThreadCalls;
                System.Windows.Forms.Control.CheckForIllegalCrossThreadCalls = true;
                try
                {
                    using (var form = new MembershipResultsForm(Fixture().Completed))
                    {
                        form.Show(); SynchronizationContext.SetSynchronizationContext(null);
                        SwitchHierarchy(form); ViewSelector(form).SelectedIndex = 0; SwitchHierarchy(form);
                        Assert.AreEqual(0, form.HierarchyDiagnostics.Count); Assert.IsFalse(Control(form).InvokeRequired);
                    }
                }
                finally { System.Windows.Forms.Control.CheckForIllegalCrossThreadCalls = previous; }
            });
        }

        [TestMethod]
        public void ViewSwitchPreservesHierarchyExpansionAndSelectedOccurrence()
        {
            Sta(() =>
            {
                using (var form = new MembershipResultsForm(Fixture().Completed))
                {
                    form.Show(); SwitchHierarchy(form); ExpandAllBounded(Control(form).Tree, 100);
                    var node = TreeNodes(Control(form).Tree).First(n => (n.Tag as MembershipHierarchyNode)?.Occurrence?.Record.ComponentType == 2);
                    Control(form).Tree.SelectedNode = node; var selected = Control(form).SelectedComparison;
                    ViewSelector(form).SelectedIndex = 0; SwitchHierarchy(form);
                    Assert.AreSame(node, Control(form).Tree.SelectedNode); Assert.AreSame(selected, Control(form).SelectedComparison);
                }
            });
        }

        [TestMethod]
        public void ActualSelectionRenderingFaultClearsPartialTreeAndPreservesFlatResults()
        {
            Sta(() =>
            {
                var f = Fixture(); using (var form = new MembershipResultsForm(f.Completed))
                {
                    form.Show(); SwitchHierarchy(form); var control = Control(form); ExpandAllBounded(control.Tree, 100);
                    typeof(MembershipHierarchyControl).GetField("describeRow", BindingFlags.Instance | BindingFlags.NonPublic)
                        .SetValue(control, new Func<MembershipGridRow, string>(_ => throw new InvalidOperationException("sensitive payload")));
                    control.Tree.SelectedNode = TreeNodes(control.Tree).First(n => (n.Tag as MembershipHierarchyNode)?.Row != null);
                    Assert.AreEqual(0, ViewSelector(form).SelectedIndex); Assert.IsTrue(Grid(form).Visible);
                    Assert.AreEqual(0, control.Tree.Nodes.Count); Assert.AreEqual(f.Presentation.Rows.Count, Grid(form).Rows.Count);
                    Assert.IsFalse(string.Join("", form.HierarchyDiagnostics).Contains("sensitive"));
                }
            });
        }

        [TestMethod]
        public void EmptyAndUnavailableInventoriesRemainHonest()
        {
            var empty = Model(Make(new ComponentIdentity[0])); Assert.AreEqual(0, empty.ComparisonRowCount); Assert.AreEqual(0, empty.VisibleOccurrenceCount);
            var p = new MembershipComparisonPresentation("sample", MembershipEnvironmentResult.Unavailable("Source", "sample", 27, TimeSpan.Zero, "Unavailable"),
                Env(new ComponentIdentity[0], true), new MembershipResultRow[0]);
            var unavailable = Model(Enrich(p));
            Assert.IsTrue(Walk(unavailable.Root).Any(n => n.Text.Contains("not absence evidence")));
        }

        [TestMethod]
        public void ScopedAppSettingAssociationsAlreadyPresentInH1AreShownWithoutH2BHandoff()
        {
            var parent = Identity(1, 80, "app"); var child = Identity(2, 10075, "setting"); var f = Make(new[] { parent, child });
            var evidence = new ComponentHierarchyParentEvidence(f.Presentation.Source.Snapshot, DisplayEnvironmentSide.Source,
                new[] { new HierarchyStructuralObservation(child.Record.SolutionComponentId, child.Record.ObjectId, parent.Record.ObjectId,
                    parent.SemanticKind, parent.ComparisonKey, "Synthetic retained and verified parent binding", true) });
            var projection = ComponentHierarchyProjector.Project(f.Completed, sourceParents: evidence);
            var rows = MembershipGridRow.CreateRows(f.Presentation, f.Completed);
            var model = MembershipHierarchyView.Create(projection, rows, Enumerable.Range(0, rows.Count), true);
            var group = Walk(model.Root).Single(n => n.Text == "Verified structural parent links");
            Assert.AreSame(child, group.Children.Single().Occurrence.Identity); Assert.IsTrue(Walk(model.Root).Any(n => n.Text.StartsWith("Model-driven Apps", StringComparison.Ordinal)));
        }

        [TestMethod]
        public void InternalNameFallbackIsVerifiedAndDoesNotParsePortableKeys()
        {
            var f = Fixture();
            var evidence = f.Completed.DisplayEvidence.Select(e => new ComponentDisplayEvidence(e.RowIndex,
                InternalSide(e.SourceDisplayEvidence), InternalSide(e.TargetDisplayEvidence)));
            var completed = new CompletedMembershipComparison(f.Presentation, evidence, DisplayProjectionOutcome.Completed,
                DisplayCacheAssociation.Verified, DisplayCacheAssociation.Verified);
            var rows = MembershipGridRow.CreateRows(f.Presentation, completed);
            var model = MembershipHierarchyView.Create(ComponentHierarchyProjector.Project(completed), rows, Enumerable.Range(0, rows.Count), true);
            Assert.IsTrue(Walk(model.Root).Where(n => n.Occurrence != null).All(n => n.Text.StartsWith("verified_logical_name", StringComparison.Ordinal)));
        }
        private static ComponentSideDisplayEvidence InternalSide(ComponentSideDisplayEvidence s) =>
            new ComponentSideDisplayEvidence(s.EnvironmentSide, s.SolutionComponentId, s.LocalObjectId, s.RawComponentType,
                s.ComponentCategory, new ComponentDisplayValue(DisplayAvailability.NotRetrieved),
                new ComponentDisplayValue(DisplayAvailability.Available, "verified_logical_name", "Retained internal name evidence"),
                s.BackingCorrelationStatus, s.ParentDisplayContext);

        [TestMethod]
        public void ConflictingRetainedNamesDoNotChooseOneArbitrarily()
        {
            var identity = Identity(1, 61, "opaque"); var original = Make(new[] { identity });
            var p = new MembershipComparisonPresentation("sample", original.Presentation.Source, original.Presentation.Target,
                new[] { original.Presentation.Rows[0], original.Presentation.Rows[0] });
            var f = Enrich(p); var values = f.Completed.DisplayEvidence.Select((e, i) => new ComponentDisplayEvidence(i,
                new ComponentSideDisplayEvidence(DisplayEnvironmentSide.Source, identity.Record.SolutionComponentId, identity.Record.ObjectId, 61,
                    e.SourceDisplayEvidence.ComponentCategory, new ComponentDisplayValue(DisplayAvailability.Available, "Name " + i, "Retained metadata"),
                    e.SourceDisplayEvidence.InternalName, e.SourceDisplayEvidence.BackingCorrelationStatus, e.SourceDisplayEvidence.ParentDisplayContext), e.TargetDisplayEvidence));
            var completed = new CompletedMembershipComparison(f.Presentation, values, DisplayProjectionOutcome.Completed, DisplayCacheAssociation.Verified, DisplayCacheAssociation.Verified);
            var rows = MembershipGridRow.CreateRows(f.Presentation, completed);
            var model = MembershipHierarchyView.Create(ComponentHierarchyProjector.Project(completed), rows, Enumerable.Range(0, rows.Count), true);
            StringAssert.Contains(Walk(model.Root).Single(n => n.Occurrence != null).Text, "Conflicting retained names");
        }

        [TestMethod]
        public void CancelledViewAdapterNeverReturnsPartialNavigation()
        {
            var f = Fixture(); var projection = ComponentHierarchyProjector.Project(f.Completed); var rows = MembershipGridRow.CreateRows(f.Presentation, f.Completed);
            Assert.ThrowsException<OperationCanceledException>(() => MembershipHierarchyView.Create(projection, rows,
                Enumerable.Range(0, rows.Count), true, new CancellationToken(true)));
        }

        [TestMethod]
        public void NativeHierarchyCanResizeAtMinimumAndLargerFormSizes()
        {
            Sta(() =>
            {
                using (var form = new MembershipResultsForm(Fixture().Completed))
                {
                    form.Show(); SwitchHierarchy(form); form.Size = form.MinimumSize; Application.DoEvents();
                    var header = form.Controls.OfType<Panel>().Single();
                    Assert.AreEqual(174, header.Height);
                    Assert.IsTrue(Control(form).Top >= Bar(form).Bottom); Assert.IsTrue(Bar(form).Top >= header.Bottom);
                    var split = Control(form).Controls.OfType<SplitContainer>().Single(); Assert.IsTrue(split.Panel1.Width >= 200); Assert.IsTrue(split.Panel2.Width >= 200);
                    form.Size = new Size(1600, 900); Application.DoEvents(); Assert.IsTrue(split.Height > 500); Assert.IsTrue(Control(form).Visible);
                }
            });
        }

        [TestMethod]
        public void DeepTenThousandOccurrenceHierarchyUsesLazyNativeExpansion()
        {
            Sta(() =>
            {
                var f = Make(Enumerable.Range(1, 10000).Select(i => Identity(i, 2, "column" + i, root: i == 1 ? (Guid?)null : Id(i - 1))));
                using (var form = new MembershipResultsForm(f.Completed))
                {
                    form.Show(); var watch = Stopwatch.StartNew(); SwitchHierarchy(form); ExpandAllBounded(Control(form).Tree, 20); watch.Stop();
                    Assert.AreEqual(10000, form.HierarchyProjection.Occurrences.Count); Assert.IsTrue(TreeNodes(Control(form).Tree).Count() < 100);
                    TestContext?.WriteLine("H2A actual WinForms synthetic deep_records=10000; show_and_expand_ms={0:F3}; native_nodes={1}",
                        watch.Elapsed.TotalMilliseconds, TreeNodes(Control(form).Tree).Count());
                }
            });
        }

        [DataTestMethod]
        [DataRow(840), DataRow(10000)]
        public void ActualNativeControlsRenderLazilyAndExpandInBoundedPages(int count)
        {
            Sta(() =>
            {
                using (var warm = new MembershipResultsForm(Fixture().Completed)) { warm.Show(); SwitchHierarchy(warm); }
                var f = Make(Enumerable.Range(1, count).Select(i => Identity(i, 61, "web" + i)).ToList());
                using (var form = new MembershipResultsForm(f.Completed))
                {
                    form.Show(); Application.DoEvents(); var watch = Stopwatch.StartNew(); SwitchHierarchy(form); watch.Stop();
                    var tree = Control(form).Tree; int initial = TreeNodes(tree).Count(); Assert.IsTrue(initial < 30, "Initial native node population must be bounded.");
                    var expand = Stopwatch.StartNew(); ExpandAllBounded(tree, 5); expand.Stop();
                    Assert.IsTrue(TreeNodes(tree).Count() < 600, "Only the first page should materialize.");
                    var model = CurrentViewModel(form); Assert.AreEqual(count, model.VisibleOccurrenceCount);
                    TestContext?.WriteLine("H2A actual WinForms synthetic records={0}; show_local_hierarchy_ms={1:F3}; initial_native_nodes={2}; expand_first_page_ms={3:F3}; rendered_nodes={4}; requests=0; writes=0",
                        count, watch.Elapsed.TotalMilliseconds, initial, expand.Elapsed.TotalMilliseconds, TreeNodes(tree).Count());
                    Assert.IsTrue(watch.Elapsed.TotalSeconds < 15); Assert.IsTrue(expand.Elapsed.TotalSeconds < 10);
                }
            });
        }

        [TestMethod]
        public void SyntheticExecutableUiScreenshotsCanBeCapturedWithoutLiveConnections()
        {
            var directory = Environment.GetEnvironmentVariable("H2A_EVIDENCE_DIRECTORY");
            if (string.IsNullOrWhiteSpace(directory)) return;
            Sta(() =>
            {
                Directory.CreateDirectory(directory); using (var form = new MembershipResultsForm(Fixture().Completed))
                {
                    form.Text += " — Synthetic H2A validation fixture"; ViewSelector(form).SelectedIndex = 0; form.Show(); Application.DoEvents(); Save(form, directory, "01-flat.png");
                    SwitchHierarchy(form); Save(form, directory, "02-hierarchy.png"); ExpandAllBounded(Control(form).Tree, 100);
                    Save(form, directory, "03-expanded.png");
                    Control(form).Tree.SelectedNode = TreeNodes(Control(form).Tree).First(n => (n.Tag as MembershipHierarchyNode)?.Occurrence?.Record.ComponentType == 300);
                    Save(form, directory, "04-parent-not-established.png");
                    Filter(form).SelectedIndex = 2; Pump(() => Status(form).Text.Contains("occurrences")); ExpandAllBounded(Control(form).Tree, 100);
                    Save(form, directory, "05-status-filter.png");
                    Control(form).Tree.SelectedNode = TreeNodes(Control(form).Tree).First(n => (n.Tag as MembershipHierarchyNode)?.Occurrence?.Record.ComponentType == 2);
                    Save(form, directory, "06-selected-details.png");
                }
            });
        }

        private static void Save(Form form, string directory, string name)
        { Application.DoEvents(); using (var image = new Bitmap(form.Width, form.Height)) { form.DrawToBitmap(image, new Rectangle(Point.Empty, form.Size)); image.Save(Path.Combine(directory, name), ImageFormat.Png); } }
        private static MembershipHierarchyView CurrentViewModel(MembershipResultsForm form) => MembershipHierarchyView.Create(form.HierarchyProjection,
            MembershipGridRow.CreateRows(form.CompletedComparison.Presentation, form.CompletedComparison), form.FilteredRows().Select(r => r.OriginalRowIndex), Filter(form).SelectedIndex == 0);
        private static IEnumerable<MembershipHierarchyNode> Walk(MembershipHierarchyNode root)
        { var stack = new Stack<MembershipHierarchyNode>(); stack.Push(root); while (stack.Count > 0) { var n = stack.Pop(); yield return n; foreach (var child in n.Children.Reverse()) stack.Push(child); } }
        private static IEnumerable<TreeNode> TreeNodes(TreeView tree)
        { var stack = new Stack<TreeNode>(tree.Nodes.Cast<TreeNode>().Reverse()); while (stack.Count > 0) { var n = stack.Pop(); yield return n; foreach (TreeNode child in n.Nodes.Cast<TreeNode>().Reverse()) stack.Push(child); } }
        private static void ExpandAllBounded(TreeView tree, int limit)
        { int count = 0; foreach (var n in TreeNodes(tree)) { if (++count > limit) break; n.Expand(); } Application.DoEvents(); }
        private static void SwitchHierarchy(MembershipResultsForm form)
        { ViewSelector(form).SelectedIndex = 1; Pump(() => form.Controls.OfType<MembershipHierarchyControl>().Any(c => c.Visible) || form.HierarchyDiagnostics.Count > 0); Assert.IsTrue(Control(form).Visible, string.Join(";", form.HierarchyDiagnostics)); }
        private static void Pump(Func<bool> ready)
        { var watch = Stopwatch.StartNew(); while (!ready() && watch.Elapsed.TotalSeconds < 20) { Application.DoEvents(); Thread.Sleep(1); } Application.DoEvents(); Assert.IsTrue(ready(), "Hierarchy UI did not complete."); }
        private static ToolStrip Bar(MembershipResultsForm f) => f.Controls.OfType<ToolStrip>().Single(bar => bar.Name == "MembershipViewBar");
        private static ToolStripComboBox ViewSelector(MembershipResultsForm f) => Bar(f).Items.OfType<ToolStripComboBox>().Single(i => i.Name == "MembershipViewSelector");
        private static ToolStripComboBox Filter(MembershipResultsForm f) => Bar(f).Items.OfType<ToolStripComboBox>().Single(i => i.Name == "MembershipPresenceFilter");
        private static ToolStripLabel Status(MembershipResultsForm f) => Bar(f).Items.OfType<ToolStripLabel>().Single(i => i.Name == "MembershipViewStatus");
        private static MembershipHierarchyControl Control(MembershipResultsForm f) => f.Controls.OfType<MembershipHierarchyControl>().Single();
        private static DataGridView Grid(MembershipResultsForm f) => f.Controls.OfType<DataGridView>().Single();
        private static void Sta(Action action)
        {
            Exception error = null; var thread = new Thread(() => { try { action(); } catch (Exception ex) { error = ex; } }) { IsBackground = true };
            thread.SetApartmentState(ApartmentState.STA); thread.Start(); Assert.IsTrue(thread.Join(TimeSpan.FromSeconds(120)), "H2A UI timeout.");
            if (error != null) ExceptionDispatchInfo.Capture(error).Throw();
        }
        private sealed class Data { internal MembershipComparisonPresentation Presentation; internal CompletedMembershipComparison Completed; }
        private static Guid Id(int n) => new Guid(n, 0, 0, new byte[8]);
        private static ComponentIdentity Identity(int id, int type, string key, Guid? root = null,
            IdentityResolutionStatus status = IdentityResolutionStatus.Resolved, Guid? objectId = null) => new ComponentIdentity(
                new SolutionComponentRecord(Id(id), type, objectId ?? Id(id + 50000), 1, root, true), status,
                status == IdentityResolutionStatus.Resolved ? key : null, "Original synthetic diagnostic", componentTypeKey: type == 1 ? ComponentSemanticKinds.Table :
                type == 2 ? ComponentSemanticKinds.Column : type == 80 ? ComponentSemanticKinds.AppModule : type == 10075 ? ComponentSemanticKinds.AppSetting :
                type == 60 ? ComponentSemanticKinds.SystemForm : type == 91 ? ComponentSemanticKinds.PluginAssembly : type == 92 ? ComponentSemanticKinds.SdkMessageProcessingStep : "componenttype:" + type,
                diagnosticEvidence: new[] { "Original diagnostic evidence" });
        private static MembershipEnvironmentResult Env(IEnumerable<ComponentIdentity> items, bool target) => MembershipEnvironmentResult.FromSnapshot(target ? "Target" : "Source",
            MembershipSnapshot.Complete(new SolutionIdentity(new EnvironmentIdentity(Id(target ? 90002 : 90001), target ? "Target" : "Source"), Id(target ? 80002 : 80001), "sample"),
                items, DateTimeOffset.UtcNow), 27, TimeSpan.Zero);
        private static Data Make(IEnumerable<ComponentIdentity> source, IEnumerable<ComponentIdentity> target = null, string parentScope = null) =>
            Enrich(new MembershipResultPresenter().Create(Env(source, false), Env(target ?? new ComponentIdentity[0], true)), parentScope);
        private static Data Fixture() => Make(new[] { Identity(1, 1, "account"), Identity(2, 2, "account.name", root: Id(1)),
            Identity(3, 300, null, status: IdentityResolutionStatus.Unsupported) }, new[] { Identity(11, 1, "account") });
        private static Data Enrich(MembershipComparisonPresentation original, string parentScope = null, bool names = true)
        {
            var p = new MembershipComparisonPresentation("sample", original.Source, original.Target, original.Rows.Select(r => r.WithDefinition("Not Compared", "",
                new ComponentDefinitionDetailPresentation(r.ComponentKind, r.PortableKey, r.MembershipStatus, "Not Compared",
                    new ComponentPropertyPresentation[0], "Original definition diagnostic", new[] { "Source definition audit" }, new[] { "Target definition audit" }))));
            Func<DisplayEnvironmentSide, ComponentIdentity, ComponentSideDisplayEvidence> side = (role, i) =>
            {
                string label = i?.Record.ComponentType == 1 ? "Account" : i?.Record.ComponentType == 2 ? "Account Name" : i?.Record.ComponentType == 300 ? "Canvas App evidence" : "Component";
                var value = names && i != null ? new ComponentDisplayValue(DisplayAvailability.Available, label, "Retained synthetic verified metadata") : new ComponentDisplayValue(DisplayAvailability.NotRetrieved);
                return new ComponentSideDisplayEvidence(role, i?.Record.SolutionComponentId, i?.Record.ObjectId, i?.Record.ComponentType,
                    value, value, new ComponentDisplayValue(DisplayAvailability.NotRetrieved), DisplayBackingCorrelationStatus.VerifiedMetadata,
                    parentScope == null ? new ComponentDisplayValue(DisplayAvailability.NotRetrieved) :
                        new ComponentDisplayValue(DisplayAvailability.Available, parentScope, "AttributeMetadata.EntityLogicalName"));
            };
            return new Data { Presentation = p, Completed = new CompletedMembershipComparison(p, p.Rows.Select((r, index) => new ComponentDisplayEvidence(index,
                side(DisplayEnvironmentSide.Source, r.SourceIdentity), side(DisplayEnvironmentSide.Target, r.TargetIdentity))), DisplayProjectionOutcome.Completed,
                DisplayCacheAssociation.Verified, DisplayCacheAssociation.Verified) };
        }
        private static MembershipHierarchyView Model(Data f) => MembershipHierarchyView.Create(ComponentHierarchyProjector.Project(f.Completed),
            MembershipGridRow.CreateRows(f.Presentation, f.Completed), Enumerable.Range(0, f.Presentation.Rows.Count), true);
        private static MembershipComparisonPresentation Baseline()
        {
            var source = new List<ComponentIdentity>(); var target = new List<ComponentIdentity>(); int id = 1;
            for (int i = 0; i < 672; i++) { source.Add(Identity(id++, 61, "shared" + i)); target.Add(Identity(id++, 61, "shared" + i)); }
            for (int i = 0; i < 103; i++) source.Add(Identity(id++, 61, "source" + i));
            for (int i = 0; i < 32; i++) target.Add(Identity(id++, 61, "target" + i));
            for (int i = 0; i < 13; i++) { source.Add(Identity(id++, 300, null, status: IdentityResolutionStatus.Unsupported)); target.Add(Identity(id++, 300, null, status: IdentityResolutionStatus.Unsupported)); }
            for (int i = 0; i < 2; i++) { source.Add(Identity(id++, 92, null, status: IdentityResolutionStatus.Ambiguous)); target.Add(Identity(id++, 92, null, status: IdentityResolutionStatus.Ambiguous)); }
            Func<ComponentIdentity> indeterminate = () => new ComponentIdentity(new SolutionComponentRecord(Id(id++), 60, Id(id + 50000)),
                IdentityResolutionStatus.Resolved, "match-only" + id, componentTypeKey: ComponentSemanticKinds.SystemForm, inventoryAbsencePolicy: InventoryAbsencePolicy.MatchOnly);
            source.Add(indeterminate()); target.Add(indeterminate()); target.Add(indeterminate());
            return new MembershipResultPresenter().Create(Env(source, false), Env(target, true));
        }
        #region R2E hierarchy additions
        [TestMethod, TestCategory("D4T8R2E")]
        public void R2EHierarchyKeyboardRoutesRetainNativeExpansionAndExactRow()
        {
            Sta(() =>
            {
                using (var form = new MembershipResultsForm(Fixture().Completed))
                {
                    form.Show(); SwitchHierarchy(form); var c = Control(form); var tree = c.Tree; var source = tree.Nodes[0].Nodes[0];
                    tree.SelectedNode = source; source.Collapse(); tree.Focus(); Assert.IsTrue(tree.Focused);
                    R2EHierarchyKey(tree, Keys.Right); Application.DoEvents(); Assert.IsTrue(source.IsExpanded);
                    typeof(System.Windows.Forms.Control).GetMethod("OnKeyDown", BindingFlags.NonPublic | BindingFlags.Instance).Invoke(tree, new object[] { new KeyEventArgs(Keys.Enter) });
                    Assert.IsFalse(source.IsExpanded); Assert.IsNull(c.SelectedComparison);
                    c.SelectComparison(0, DisplayEnvironmentSide.Source); Application.DoEvents(); Assert.IsNotNull(c.SelectedComparison);
                    var row = c.SelectedComparison.AuthoritativeRow; var side = c.SelectedSide;
                    form.SelectNextControl(tree, true, true, true, true); Application.DoEvents();
                    Assert.AreSame(row, c.SelectedComparison.AuthoritativeRow); Assert.AreEqual(side, c.SelectedSide); Assert.IsFalse(tree.Focused);
                    Assert.AreEqual(3, form.CompletedComparison.Presentation.Rows.Count);
                }
            });
        }
        [TestMethod, TestCategory("D4T8R2E")]
        public void R2EHierarchyRoundTripRestoresExactOccurrenceAndExpansion()
        {
            Sta(() =>
            {
                var one = Identity(1, 60, "repeat", objectId: Id(500)); var two = Identity(2, 60, "repeat", objectId: Id(500));
                var f = Make(new[] { one, two });
                using (var form = new MembershipResultsForm(f.Completed))
                {
                    form.Show(); SwitchHierarchy(form); var c = Control(form); ExpandAllBounded(c.Tree, 100);
                    var occurrences = TreeNodes(c.Tree).Where(n => (n.Tag as MembershipHierarchyNode)?.Occurrence != null).ToArray(); Assert.AreEqual(2, occurrences.Length);
                    c.Tree.SelectedNode = occurrences[1]; Application.DoEvents(); var exact = c.SelectedOccurrenceIndex; var row = c.SelectedComparison.AuthoritativeRow;
                    Assert.AreEqual(((MembershipHierarchyNode)occurrences[1].Tag).Occurrence.OccurrenceIndex, exact); Assert.AreEqual(DisplayEnvironmentSide.Source, c.SelectedSide);
                    occurrences[1].Expand(); ViewSelector(form).SelectedIndex = 0; SwitchHierarchy(form);
                    Assert.AreSame(occurrences[1], c.Tree.SelectedNode); Assert.AreEqual(exact, c.SelectedOccurrenceIndex); Assert.AreSame(row, c.SelectedComparison.AuthoritativeRow);
                    Assert.IsTrue(c.Tree.SelectedNode.IsExpanded); Assert.AreEqual(1, form.HierarchyProjection.ComparisonRows.Count); Assert.AreEqual(2, form.HierarchyProjection.Occurrences.Count);
                }
                var pageSource = Enumerable.Range(0, 205).Select(i => Identity(1000 + i, 61, "resource" + i)).Concat(new[] { Identity(9000, 61, "only-source") }).ToArray();
                var pageTarget = Enumerable.Range(0, 205).Select(i => Identity(3000 + i, 61, "resource" + i)).ToArray();
                using (var form = new MembershipResultsForm(Make(pageSource, pageTarget).Completed))
                {
                    form.Show(); SwitchHierarchy(form); var c = Control(form); var source = c.Tree.Nodes[0].Nodes[0]; source.Expand();
                    var group = source.Nodes.Cast<TreeNode>().Single(n => (n.Tag as MembershipHierarchyNode)?.Children.Count > 200); group.Expand();
                    var page = group.Nodes.Cast<TreeNode>().Single(n => n.Tag != null && n.Tag.GetType().Name == "Page" && R2EHierarchyPageStart(n) == 200);
                    c.Tree.SelectedNode = page; Assert.IsNull(c.SelectedComparison); Assert.AreEqual(200, R2EHierarchyPageStart(c.Tree.SelectedNode));
                    Filter(form).SelectedIndex = 1; SwitchHierarchy(form);
                    Assert.IsNull(c.SelectedComparison); Assert.AreEqual("Page", c.Tree.SelectedNode.Tag.GetType().Name); Assert.AreEqual(200, R2EHierarchyPageStart(c.Tree.SelectedNode));
                    var filteredPage = c.Tree.SelectedNode; ViewSelector(form).SelectedIndex = 0; Assert.IsNull(Grid(form).CurrentCell); SwitchHierarchy(form);
                    Assert.AreSame(filteredPage, c.Tree.SelectedNode); Assert.IsNull(c.SelectedComparison);
                    Filter(form).SelectedIndex = 0; SwitchHierarchy(form); Assert.AreEqual(200, R2EHierarchyPageStart(c.Tree.SelectedNode)); Assert.IsNull(c.SelectedComparison);
                    Assert.AreEqual(206, form.FilteredRows().Count); Assert.AreEqual(411, form.HierarchyProjection.Occurrences.Count);
                }
            });
        }
        // R2E correction helpers stay within the existing R2E additions.
        private static int R2EHierarchyPageStart(TreeNode node) => (int)node.Tag.GetType().GetField("Start", BindingFlags.Instance | BindingFlags.NonPublic).GetValue(node.Tag);
        [TestMethod, TestCategory("D4T8R2E")]
        public void R2EHierarchyFilterRestoreKeepsContextAncestorsAndNoFalseSelection()
        {
            Sta(() =>
            {
                var f = Fixture(); using (var form = new MembershipResultsForm(f.Completed))
                {
                    form.Show(); SwitchHierarchy(form); var c = Control(form); ExpandAllBounded(c.Tree, 100);
                    var column = TreeNodes(c.Tree).Single(n => (n.Tag as MembershipHierarchyNode)?.Occurrence?.Record.ComponentType == 2);
                    c.Tree.SelectedNode = column; var row = c.SelectedComparison.AuthoritativeRow; var raw = c.SelectedOccurrenceIndex;
                    Filter(form).SelectedIndex = 2; SwitchHierarchy(form); Assert.AreSame(row, c.SelectedComparison.AuthoritativeRow); Assert.AreEqual(raw, c.SelectedOccurrenceIndex);
                    ExpandAllBounded(c.Tree, 100); var context = TreeNodes(c.Tree).Single(n => (n.Tag as MembershipHierarchyNode)?.Occurrence?.Record.ComponentType == 1);
                    Assert.IsTrue(((MembershipHierarchyNode)context.Tag).NavigationContext); c.Tree.SelectedNode = context; Assert.IsNull(c.SelectedComparison);
                    var reset = form.Controls.OfType<ToolStrip>().SelectMany(t => t.Items.OfType<ToolStripButton>()).Single(x => x.Name == "MembershipResetFiltersAction");
                    reset.PerformClick(); SwitchHierarchy(form); Assert.AreEqual(0, Filter(form).SelectedIndex); Assert.IsNull(c.SelectedComparison); Assert.AreSame(c.Tree.Nodes[0], c.Tree.SelectedNode);
                    Assert.AreEqual(f.Presentation.Rows.Count, form.FilteredRows().Count);
                    Assert.IsTrue(form.HierarchyProjection.Occurrences.Any(o => o.Placement == HierarchyPlacement.VerifiedInclusionRoot));
                    var sourceCategory = c.Tree.Nodes[0].Nodes[0]; c.Tree.SelectedNode = sourceCategory; Assert.IsNull(c.SelectedComparison);
                    Filter(form).SelectedIndex = 1; SwitchHierarchy(form);
                    Assert.IsNull(c.SelectedComparison); Assert.AreSame(c.Tree.Nodes[0].Nodes[0], c.Tree.SelectedNode);
                    ViewSelector(form).SelectedIndex = 0; Assert.IsNull(Grid(form).CurrentCell); SwitchHierarchy(form);
                    Assert.IsNull(c.SelectedComparison); Assert.AreSame(c.Tree.Nodes[0].Nodes[0], c.Tree.SelectedNode);
                    c.ClearNavigationSelection(); Assert.IsNull(c.Tree.SelectedNode); Assert.IsNull(c.SelectedComparison);
                    Filter(form).SelectedIndex = 0; SwitchHierarchy(form); Assert.IsNull(c.Tree.SelectedNode); Assert.IsNull(c.SelectedComparison);
                    ViewSelector(form).SelectedIndex = 0; Assert.IsNull(Grid(form).CurrentCell); SwitchHierarchy(form);
                    Assert.IsNull(c.Tree.SelectedNode); Assert.IsNull(c.SelectedComparison);
                    ViewSelector(form).SelectedIndex = 0; var flat = Grid(form); var current = (MembershipGridRow)flat.Rows[0].DataBoundItem;
                    flat.CurrentCell = flat.Rows[0].Cells[0];
                    typeof(DataGridView).GetMethod("OnCellClick", BindingFlags.Instance | BindingFlags.NonPublic).Invoke(flat, new object[] { new DataGridViewCellEventArgs(0, 0) });
                    SwitchHierarchy(form); Assert.IsNotNull(c.SelectedComparison); Assert.AreSame(current.AuthoritativeRow, c.SelectedComparison.AuthoritativeRow);
                }
            });
        }
        [TestMethod, TestCategory("D4T8R2E")]
        public void R2EHierarchyFocusScrollAndCloseKeepOwnerGuards()
        {
            Sta(() =>
            {
                int calls = 0; string ownerText = "OWNER-ONE"; Form host = null; bool closeDuringFormat = false; var published = new List<string>();
                var first = Fixture(); var second = Make(new[] { Identity(91, 61, "replacement") }); var ownerOne = new object(); var ownerTwo = new object();
                using (host = new Form { Size = new Size(1100, 700), StartPosition = FormStartPosition.Manual, Location = new Point(-4000, -4000), ShowInTaskbar = false })
                using (var c = new MembershipHierarchyControl(row =>
                {
                    calls++;
                    if (ownerText == "OWNER-TWO") Assert.AreSame(second.Presentation.Rows[0], row.AuthoritativeRow);
                    if (closeDuringFormat) host.Close();
                    return ownerText + "|ROW=" + row.OriginalRowIndex;
                }, _ => { }))
                {
                    c.Visible = false; var input = new TextBox { Dock = DockStyle.Top }; host.Controls.Add(c); host.Controls.Add(input);
                    var evidence = c.Controls.OfType<SplitContainer>().Single().Panel2.Controls.OfType<TextBox>().Single();
                    evidence.TextChanged += (sender, args) => published.Add(evidence.Text);
                    c.Bind(Model(first), ownerOne); host.Show(); Application.DoEvents(); c.SelectComparison(0, DisplayEnvironmentSide.Source); Application.DoEvents();
                    Assert.IsNotNull(c.SelectedComparison); Assert.AreSame(first.Presentation.Rows[0], c.SelectedComparison.AuthoritativeRow);
                    Assert.AreEqual(DisplayEnvironmentSide.Source, c.SelectedSide); Assert.AreEqual(0, calls, "Post-show hidden selection must not format full evidence.");
                    var bookmark = c.CaptureNavigation(); input.Focus(); Assert.IsTrue(input.Focused);
                    c.RestoreNavigation(bookmark, 0, DisplayEnvironmentSide.Source); Assert.AreEqual(0, calls); Assert.IsTrue(input.Focused); Assert.AreEqual("", evidence.Text);
                    c.Visible = true; Application.DoEvents(); Assert.AreEqual(1, calls); StringAssert.Contains(evidence.Text, "OWNER-ONE|ROW=0"); Assert.IsTrue(input.Focused);
                    c.Visible = false; c.RestoreNavigation(bookmark, 0, DisplayEnvironmentSide.Source); Assert.AreEqual(1, calls); Assert.AreEqual("", evidence.Text); Assert.IsTrue(input.Focused);
                    ownerText = "OWNER-TWO"; c.Bind(Model(second), ownerTwo); c.RestoreNavigation(bookmark, 0, DisplayEnvironmentSide.Source);
                    Assert.IsNull(c.SelectedComparison); Assert.AreEqual(1, calls); Assert.IsFalse(evidence.Text.Contains("OWNER-ONE"));
                    c.SelectComparison(0, DisplayEnvironmentSide.Source); Assert.AreSame(second.Presentation.Rows[0], c.SelectedComparison.AuthoritativeRow); Assert.AreEqual(1, calls);
                    c.Visible = true; Application.DoEvents(); Assert.AreEqual(2, calls); StringAssert.Contains(evidence.Text, "OWNER-TWO|ROW=0"); Assert.IsFalse(evidence.Text.Contains("OWNER-ONE")); Assert.IsTrue(input.Focused);
                    c.Visible = false; c.RestoreNavigation(c.CaptureNavigation(), 0, DisplayEnvironmentSide.Source); Assert.AreEqual(2, calls);
                    ownerText = "CLOSED-PUBLICATION"; closeDuringFormat = true; c.Visible = true; Application.DoEvents();
                    Assert.IsTrue(host.IsDisposed); Assert.AreEqual(3, calls); Assert.IsFalse(published.Any(text => text.Contains("CLOSED-PUBLICATION")), "A closed control cannot publish a retained formatter result.");
                }
                var newer = Make(new[] { Identity(101, 61, "one"), Identity(102, 61, "two") }, new[] { Identity(201, 61, "one"), Identity(202, 61, "two") });
                int newerCalls = 0;
                using (var nextHost = new Form { Size = new Size(1100, 700), StartPosition = FormStartPosition.Manual, Location = new Point(-4000, -4000), ShowInTaskbar = false })
                using (var c = new MembershipHierarchyControl(row => { newerCalls++; return "CURRENT-OWNER|ROW=" + row.OriginalRowIndex; }, _ => { }))
                {
                    c.Visible = false; nextHost.Controls.Add(c); var input = new TextBox { Dock = DockStyle.Top }; nextHost.Controls.Add(input);
                    c.Bind(Model(newer), newer.Presentation); nextHost.Show(); Application.DoEvents(); c.SelectComparison(0, DisplayEnvironmentSide.Source);
                    Assert.AreSame(newer.Presentation.Rows[0], c.SelectedComparison.AuthoritativeRow); Assert.AreEqual(DisplayEnvironmentSide.Source, c.SelectedSide); Assert.AreEqual(0, newerCalls);
                    var sourceBookmark = c.CaptureNavigation(); c.SelectComparison(0, DisplayEnvironmentSide.Target);
                    c.RestoreNavigation(sourceBookmark, 0, DisplayEnvironmentSide.Target);
                    Assert.AreEqual(DisplayEnvironmentSide.Target, c.SelectedSide); Assert.AreSame(newer.Presentation.Rows[0], c.SelectedComparison.AuthoritativeRow); Assert.AreEqual(0, newerCalls);
                    c.SelectComparison(1, DisplayEnvironmentSide.Target); c.RestoreNavigation(sourceBookmark, 1, DisplayEnvironmentSide.Target);
                    Assert.AreEqual(DisplayEnvironmentSide.Target, c.SelectedSide); Assert.AreSame(newer.Presentation.Rows[1], c.SelectedComparison.AuthoritativeRow); Assert.AreEqual(0, newerCalls);
                    input.Focus(); c.Visible = true; Application.DoEvents(); Assert.IsTrue(input.Focused); Assert.AreEqual(1, newerCalls);
                    var text = c.Controls.OfType<SplitContainer>().Single().Panel2.Controls.OfType<TextBox>().Single(); StringAssert.Contains(text.Text, "CURRENT-OWNER|ROW=1"); Assert.IsFalse(text.Text.Contains("CURRENT-OWNER|ROW=0"));
                    c.Visible = false; var rowBookmark = c.CaptureNavigation(); c.Tree.SelectedNode = c.Tree.Nodes[0].Nodes[0];
                    c.RestoreNavigation(rowBookmark, null, null); Assert.IsNull(c.SelectedComparison); Assert.AreSame(c.Tree.Nodes[0].Nodes[0], c.Tree.SelectedNode); Assert.AreEqual(1, newerCalls);
                    c.ClearNavigationSelection(); c.RestoreNavigation(rowBookmark, null, null); Assert.IsNull(c.SelectedComparison); Assert.IsNull(c.Tree.SelectedNode); Assert.AreEqual(1, newerCalls);
                }
            });
        }
        [System.Runtime.InteropServices.DllImport("user32.dll", EntryPoint = "SendMessageW")]
        private static extern IntPtr R2EHierarchySendKey(IntPtr window, int message, IntPtr key, IntPtr detail);
        private static void R2EHierarchyKey(TreeView tree, Keys key)
        {
            R2EHierarchySendKey(tree.Handle, 0x100, (IntPtr)(int)key, IntPtr.Zero);
            R2EHierarchySendKey(tree.Handle, 0x101, (IntPtr)(int)key, IntPtr.Zero);
        }
        #endregion
    }
}
