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
    [TestClass, TestCategory("H2CPresentation")]
    public class MembershipHierarchyPresentationTests
    {
        public TestContext TestContext { get; set; }

        [DataTestMethod]
        [DataRow("OnlineMode"), DataRow("ChannelAvailability"), DataRow("EnableRoleBasedOfflineProfileAccess")]
        [DataRow("msdyn_enablertecopilot"), DataRow("msdyn_rtecopilotdefaultonruncompleted"), DataRow("namespace:setting")]
        public void VerifiedAppSettingNamesAreDisplayFormattingOnlyAndKeepFullIdentity(string setting)
        {
            var source = Identity(10075, Key("app", setting)); var target = Identity(10075, source.ComparisonKey);
            var f = Make(new[] { source }, new[] { target }); var csv = Csv(f); var rows = Rows(f);
            Assert.AreEqual(setting, rows.Single().ComponentText);
            Assert.AreEqual(source.ComparisonKey, rows.Single().PortableKey);
            Assert.AreEqual(DisplayAvailability.NotRetrieved, rows.Single().SourceDisplayEvidence.FriendlyName.Availability);
            var model = Model(f); var sourceNode = Walk(model.Root).Single(n => n.Occurrence?.Identity == source);
            StringAssert.StartsWith(sourceNode.Text, setting); StringAssert.Contains(MembershipHierarchyView.Describe(sourceNode), source.ComparisonKey);
            StringAssert.Contains(MembershipHierarchyView.Describe(sourceNode), "Setting definition identity (from verified membership key): " + setting);
            Assert.AreEqual(HierarchyPlacement.ParentNotEstablished, sourceNode.Occurrence.Placement);
            Assert.IsFalse(Walk(model.Root).Any(n => n.Text.StartsWith("AppSettings ("))); // Key qualifier cannot create parent association.
            Assert.AreEqual(csv, Csv(f)); Assert.AreSame(f.Presentation.Rows.Single(), rows.Single().AuthoritativeRow);
        }

        [DataTestMethod]
        [DataRow("opaque"), DataRow("appsetting:v1:3:app:8:broken"), DataRow("appsetting:v1:03:app:4:name")]
        [DataRow("appsetting:v1:3:app:4:name:extra"), DataRow("appsetting:v1:999999999999999:app:4:name")]
        [DataRow("appsetting:v1:3:app:1: "), DataRow("appsetting:v2:3:app:4:name")]
        public void AppSettingNameUnavailableRetainsVerifiedCompleteKey(string key)
        {
            var f = Make(new[] { Identity(10075, key) });
            StringAssert.Contains(Rows(f).Single().ComponentText, key);
            StringAssert.StartsWith(Walk(Model(f).Root).Single(n => n.Occurrence != null).Text, key);
        }

        [DataTestMethod]
        [DataRow(IdentityResolutionStatus.Unresolved), DataRow(IdentityResolutionStatus.Unsupported), DataRow(IdentityResolutionStatus.Ambiguous)]
        public void CandidateDiagnosticOrLocalGuidCannotManufactureAppSettingName(IdentityResolutionStatus status)
        {
            var item = Identity(10075, null, status: status); var f = Make(new[] { item });
            var node = Walk(Model(f).Root).Single(n => n.Occurrence != null);
            StringAssert.Contains(node.Text, "Name unavailable"); Assert.AreEqual(status, node.Occurrence.Identity.Status);
            Assert.IsFalse(node.Text.Contains("OnlineMode")); Assert.IsNull(item.ComparisonKey);
        }

        [TestMethod]
        public void AppSettingParentQualifierIsNotConfusedWithVerifiedParentOwnership()
        {
            var f = AppSettings(); var model = Model(f); var rows = Rows(f); var csv = Csv(f);
            foreach (var side in new[] { DisplayEnvironmentSide.Source, DisplayEnvironmentSide.Target })
            {
                var app = Walk(model.Root).Single(n => n.Occurrence?.Record.ComponentType == 80 && n.Occurrence.Side == side);
                var group = app.Children.Single(n => n.Text == "AppSettings (5)");
                Assert.AreEqual(5, group.Children.Count); StringAssert.Contains(MembershipHierarchyView.Describe(group), "raw occurrences");
                foreach (var child in group.Children)
                {
                    Assert.AreEqual(side, child.Occurrence.Side); Assert.AreEqual(HierarchyPlacement.VerifiedStructuralParent, child.Occurrence.Placement);
                    var text = MembershipHierarchyView.Describe(child);
                    StringAssert.Contains(text, "Verified parent identity: app"); StringAssert.Contains(text, "Verified parent component: Field Service Mobile");
                    StringAssert.Contains(text, "StructuralParent: Verified"); StringAssert.Contains(text, "Bound normal AppSetting evidence");
                    Assert.AreSame(f.Presentation.Rows[child.Row.OriginalRowIndex], child.Row.AuthoritativeRow);
                }
            }
            Assert.AreEqual(12, model.VisibleOccurrenceCount); Assert.AreEqual(6, model.ComparisonRowCount); Assert.AreEqual(csv, Csv(f));
            Assert.IsTrue(rows.All(r => r.MembershipStatus == "Present in Both"));
        }

        [DataTestMethod]
        [DataRow("account.syn_regionlocation"), DataRow("appointment.regardingobjectid"), DataRow("contact.regardingobjectid")]
        public void FlatColumnsUseCompleteVerifiedQualifiedIdentity(string key)
        {
            var identity = Identity(2, key); var f = Make(new[] { identity }, new[] { Identity(2, key) }, names: "Short display name");
            Assert.AreEqual(key, Rows(f).Single().ComponentText); Assert.AreEqual(key, Rows(f).Single().PortableKey);
            var node = Walk(Model(f).Root).First(n => n.Occurrence != null); StringAssert.StartsWith(node.Text, key);
        }

        [DataTestMethod]
        [DataRow(true), DataRow(false)]
        public void OneSidedColumnsKeepExplicitSideAndOriginalClassification(bool source)
        {
            var item = Identity(2, "appointment.regardingobjectid"); var f = Make(source ? new[] { item } : new ComponentIdentity[0], source ? null : new[] { item });
            var row = Rows(f).Single(); StringAssert.Contains(row.ComponentText, (source ? "Source: " : "Target: ") + item.ComparisonKey);
            Assert.AreEqual(source ? "Source Only" : "Target Only", row.MembershipStatus);
            Assert.AreEqual(source ? MembershipPresence.OnlyInSource : MembershipPresence.OnlyInTarget, row.AuthoritativeRow.Presence);
            Assert.AreSame(item.Record, Walk(Model(f).Root).Single(n => n.Occurrence != null).Occurrence.Record);
        }

        [TestMethod]
        public void SameColumnNameUnderDifferentTablesRemainsDistinct()
        {
            var f = Make(new[] { Identity(2, "account.regardingobjectid"), Identity(2, "appointment.regardingobjectid") }, names: "regardingobjectid");
            CollectionAssert.AreEquivalent(new[] { "Source: account.regardingobjectid", "Source: appointment.regardingobjectid" }, Rows(f).Select(r => r.ComponentText.Split('\n')[0].TrimEnd('\r')).ToArray());
            Assert.AreEqual(2, f.Presentation.Summary.SourceOnly);
        }

        [DataTestMethod]
        [DataRow("unqualified"), DataRow("account."), DataRow(".column"), DataRow("account.column.extra")]
        [DataRow("account.display label"), DataRow("column:v1:account:name")]
        public void UnverifiableQualificationKeepsSafeExistingNameAndTechnicalEvidence(string key)
        {
            var f = Make(new[] { Identity(2, key) }, names: "Verified retained display");
            StringAssert.Contains(Rows(f).Single().ComponentText, "Verified retained display"); Assert.IsNull(MembershipGridRow.QualifiedColumnIdentity(f.Presentation.Rows[0].SourceIdentity));
            StringAssert.Contains(MembershipHierarchyView.Describe(Walk(Model(f).Root).Single(n => n.Occurrence != null)), key);
        }

        [DataTestMethod]
        [DataRow(true), DataRow(false)]
        public void InclusionAndStructuralGroupsHaveSimpleLabelsButSeparateEvidence(bool inclusion)
        {
            var table = Identity(1, "account"); var child = Identity(2, "account.name", root: inclusion ? table.Record.SolutionComponentId : (Guid?)null);
            var f = Make(new[] { table, child }, tableScope: inclusion ? null : "account"); var model = Model(f);
            var parent = Walk(model.Root).Single(n => n.Occurrence?.Identity == table); var group = parent.Children.Single(n => n.Text == "Columns (1)");
            var node = group.Children.Single(); StringAssert.StartsWith(node.Text, "name");
            var detail = MembershipHierarchyView.Describe(node);
            StringAssert.Contains(detail, inclusion ? "InclusionRoot: Verified" : "StructuralParent: Verified");
            StringAssert.Contains(MembershipHierarchyView.Describe(group), "not structural ownership");
            Assert.AreEqual(inclusion ? HierarchyPlacement.VerifiedInclusionRoot : HierarchyPlacement.VerifiedStructuralParent, node.Occurrence.Placement);
            Assert.AreSame(child.Record, node.Occurrence.Record);
        }

        [TestMethod]
        public void DifferentRootParentsDoNotLoseSourceTargetIdentityOrRowLocator()
        {
            var sourceParent = Identity(1, "account"); var targetParent = Identity(1, "contact");
            var sourceChild = Identity(2, "account.name", root: sourceParent.Record.SolutionComponentId);
            var targetChild = Identity(2, "account.name", root: targetParent.Record.SolutionComponentId);
            var f = Make(new[] { sourceParent, sourceChild }, new[] { targetParent, targetChild }); var model = Model(f);
            var a = Walk(model.Root).Single(n => n.Occurrence?.Identity == sourceChild); var b = Walk(model.Root).Single(n => n.Occurrence?.Identity == targetChild);
            StringAssert.StartsWith(a.Text, "name"); StringAssert.StartsWith(b.Text, "account.name");
            StringAssert.Contains(MembershipHierarchyView.Describe(a), "Verified parent identity: account");
            StringAssert.Contains(MembershipHierarchyView.Describe(b), "Verified parent identity: contact");
            Assert.AreSame(a.Row.AuthoritativeRow, b.Row.AuthoritativeRow); Assert.AreNotEqual(a.Occurrence.Side, b.Occurrence.Side);
        }

        [TestMethod]
        public void NavigationCategoriesAreNotTopLevelPlacementProof()
        {
            var f = Make(new[] { Identity(1, "account"), Identity(80, "app"), Identity(380, "env"), Identity(9, "choice"), Identity(61, "script"), Identity(300, null, status: IdentityResolutionStatus.Unsupported) });
            var model = Model(f); var source = model.Root.Children.First();
            CollectionAssert.AreEquivalent(new[] { "Tables (1)", "Model-driven Apps (1)", "Environment Variables (1)", "Choices (1)", "Web Resources (1)", "Other Components (1)" }, source.Children.Select(n => n.Text).ToArray());
            foreach (var group in source.Children) StringAssert.Contains(MembershipHierarchyView.Describe(group), "not verified top-level");
            foreach (var occurrence in Walk(source).Where(n => n.Occurrence != null))
            { Assert.AreEqual(HierarchyPlacement.ParentNotEstablished, occurrence.Occurrence.Placement); StringAssert.Contains(MembershipHierarchyView.Describe(occurrence), "Hierarchy placement: ParentNotEstablished"); }
        }

        [DataTestMethod]
        [DataRow(IdentityResolutionStatus.Unsupported), DataRow(IdentityResolutionStatus.Unresolved), DataRow(IdentityResolutionStatus.Ambiguous)]
        public void UncertainComponentsRemainAccessibleAndUnclassifiedByNavigation(IdentityResolutionStatus state)
        {
            var item = Identity(300, null, status: state); var f = Make(new[] { item }); var model = Model(f); var node = Walk(model.Root).Single(n => n.Occurrence != null);
            Assert.AreEqual(1, model.VisibleOccurrenceCount); Assert.AreEqual(1, model.ComparisonRowCount); Assert.AreSame(item, node.Occurrence.Identity);
            StringAssert.Contains(node.Text, "Indeterminate"); Assert.AreEqual(MembershipPresence.Indeterminate, node.Row.AuthoritativeRow.Presence);
        }

        [TestMethod]
        public void RawAndComparisonAccountingRemainIndependentAndLossless()
        {
            var source = new List<ComponentIdentity>(); var target = new List<ComponentIdentity>();
            for (int i = 0; i < 672; i++) { source.Add(Identity(61, "both" + i)); target.Add(Identity(61, "both" + i)); }
            for (int i = 0; i < 103; i++) source.Add(Identity(61, "source" + i)); for (int i = 0; i < 32; i++) target.Add(Identity(61, "target" + i));
            for (int i = 0; i < 13; i++) { source.Add(Identity(300, null, status: IdentityResolutionStatus.Unsupported)); target.Add(Identity(300, null, status: IdentityResolutionStatus.Unsupported)); }
            for (int i = 0; i < 3; i++) source.Add(Identity(92, null, status: IdentityResolutionStatus.Ambiguous)); target.Add(Identity(92, null, status: IdentityResolutionStatus.Ambiguous));
            source.Add(Identity(60, "source-match-only", matchOnly: true)); target.Add(Identity(60, "target-match-only-1", matchOnly: true)); target.Add(Identity(60, "target-match-only-2", matchOnly: true));
            var f = Make(source, target); var csv = Csv(f); var model = Model(f); var summary = f.Presentation.Summary;
            Assert.AreEqual(792, f.Presentation.Source.Snapshot.Components.Count); Assert.AreEqual(720, f.Presentation.Target.Snapshot.Components.Count);
            Assert.AreEqual(1512, model.VisibleOccurrenceCount); Assert.AreEqual(840, model.ComparisonRowCount);
            Assert.AreEqual(672, summary.PresentInBoth); Assert.AreEqual(103, summary.SourceOnly); Assert.AreEqual(32, summary.TargetOnly);
            Assert.AreEqual(26, summary.IndeterminateUnsupported); Assert.AreEqual(4, summary.Ambiguous); Assert.AreEqual(3, summary.IndeterminateResolved);
            Assert.AreEqual(1512, Walk(model.Root).Count(n => n.Occurrence != null)); Assert.AreEqual(840, Walk(model.Root).Where(n => n.Row != null).Select(n => n.Row.OriginalRowIndex).Distinct().Count());
            Assert.AreEqual(csv, Csv(f));
        }

        [TestMethod]
        public void RepeatedRawReferencesRemainSeparateEvenWhenComparisonCollapsesThem()
        {
            var item = Identity(60, "form"); var f = Make(new[] { item, item }, new[] { Identity(60, "form") }); var model = Model(f);
            Assert.AreEqual(3, model.VisibleOccurrenceCount); Assert.AreEqual(1, model.ComparisonRowCount);
            Assert.AreEqual(3, Walk(model.Root).Count(n => n.Occurrence != null)); Assert.AreEqual(1, f.Presentation.Summary.PresentInBoth);
        }

        [TestMethod]
        public void NativeViewsSelectionMetadataAndOriginalDetailsRemainAvailableWithZeroRequests()
        {
            Sta(() =>
            {
                var f = AppSettings(); var spy = new FakeOrganizationService(); var csv = Csv(f); int calls = 0;
                using (var form = new MembershipResultsForm(f.Completed, sourceService: spy, targetService: spy))
                {
                    Selector(form).SelectedIndex = 0; form.Show(); Application.DoEvents(); Assert.AreEqual("Flat View", Selector(form).SelectedItem);
                    CollectionAssert.AreEqual(new[] { "Flat View", "Hierarchy View", "Differences Only", "Solution Objects Explorer" }, Selector(form).Items.Cast<string>().ToArray());
                    var originalProjection = form.ProjectHierarchy; form.ProjectHierarchy = token => { calls++; return originalProjection(token); };
                    Switch(form); var control = form.Controls.OfType<MembershipHierarchyControl>().Single();
                    Assert.IsTrue(WalkNative(control.Tree).Any(n => n.Text == "Source: SYN-THET-IAB"));
                    control.SelectComparison(2, DisplayEnvironmentSide.Target); var selected = control.SelectedComparison;
                    Assert.IsNotNull(selected); Assert.AreEqual(DisplayEnvironmentSide.Target, control.SelectedSide);
                    var selectedNode = control.Tree.SelectedNode; selectedNode.Collapse(); selectedNode.Expand(); Assert.AreSame(selectedNode, control.Tree.SelectedNode);
                    StringAssert.Contains(Controls(control).OfType<TextBox>().Single().Text, "Original comparison row index: " + selected.OriginalRowIndex);
                    Selector(form).SelectedIndex = 0; Assert.AreEqual(selected.OriginalRowIndex, ((MembershipGridRow)Grid(form).CurrentRow.DataBoundItem).OriginalRowIndex);
                    Switch(form); Assert.AreSame(selectedNode, control.Tree.SelectedNode); Assert.AreSame(selected, control.SelectedComparison);
                    var metadata = Controls(form).OfType<CheckBox>().Single(c => c.Name == "OptionalMetadataToggle"); metadata.Checked = true;
                    Assert.AreEqual(9, Grid(form).Columns.Cast<DataGridViewColumn>().Count(c => c.Visible)); metadata.Checked = false;
                    Selector(form).SelectedIndex = 2; Assert.AreEqual(0, Grid(form).Rows.Count); Selector(form).SelectedIndex = 0; Assert.AreEqual(6, Grid(form).Rows.Count);
                    Switch(form); Assert.AreEqual(1, calls); Assert.AreEqual(0, spy.Calls + spy.ExecuteCalls + spy.WriteCalls); Assert.AreEqual(csv, Csv(f));
                }
            });
        }

        [TestMethod]
        public void DifferencesOnlyIncludesConfirmedDifferencesAndIndeterminateObservations()
        {
            Sta(() =>
            {
                var f = Make(new[] { Identity(61, "both"), Identity(61, "source"), Identity(300, null, status: IdentityResolutionStatus.Unsupported), Identity(92, null, status: IdentityResolutionStatus.Ambiguous) }, new[] { Identity(61, "both"), Identity(61, "target") });
                var rows = f.Presentation.Rows.Select(r => r.Presence == MembershipPresence.PresentInBoth ? r.WithDefinition("Different", "synthetic", r.DefinitionDetail) : r).ToList();
                f = Enrich(new MembershipComparisonPresentation("SAA", f.Presentation.Source, f.Presentation.Target, rows));
                using (var form = new MembershipResultsForm(f.Completed))
                {
                    form.Show(); Selector(form).SelectedIndex = 2; Assert.AreEqual(rows.Count, Grid(form).Rows.Count);
                    Assert.IsTrue(form.FilteredRows().Any(r => r.AuthoritativeRow.Presence == MembershipPresence.Indeterminate));
                    Assert.IsTrue(form.FilteredRows().Any(r => r.DefinitionStatus == "Different")); Assert.AreSame(rows[0], f.Presentation.Rows[0]);
                }
            });
        }

        [DataTestMethod]
        [DataRow(840), DataRow(10000)]
        public void NativeHierarchyOpenAndExpansionRemainBoundedWithoutNetwork(int count)
        {
            Sta(() =>
            {
                var f = Make(Enumerable.Range(0, count).Select(i => Identity(61, "script" + i))); var spy = new FakeOrganizationService();
                using (var form = new MembershipResultsForm(f.Completed, sourceService: spy, targetService: spy))
                {
                    form.Show(); var watch = Stopwatch.StartNew(); Switch(form); watch.Stop(); var control = form.Controls.OfType<MembershipHierarchyControl>().Single();
                    int initial = WalkNative(control.Tree).Count(); Assert.IsTrue(initial < 30);
                    var expansion = Stopwatch.StartNew();
                    // Expand through newly materialized family levels and the first bounded page;
                    // a single snapshot walk would only expand ancestors, not actual component entries.
                    for (int pass = 0; pass < 5; pass++)
                        foreach (var node in WalkNative(control.Tree).Where(n =>
                            n.Tag is MembershipHierarchyNode model && model.Occurrence == null && model.Row == null && model.Children.Count > 0 ||
                            n.Text.StartsWith("Items 1–", StringComparison.Ordinal)).ToArray()) node.Expand();
                    expansion.Stop();
                    Assert.AreEqual(200, WalkNative(control.Tree).Count(n => (n.Tag as MembershipHierarchyNode)?.Occurrence != null));
                    Assert.IsTrue(WalkNative(control.Tree).Count() < 600); Assert.AreEqual(count, form.HierarchyProjection.Occurrences.Count);
                    Assert.IsTrue(watch.Elapsed.TotalSeconds < 15); Assert.IsTrue(expansion.Elapsed.TotalSeconds < 10);
                    Assert.AreEqual(0, spy.Calls + spy.ExecuteCalls + spy.WriteCalls);
                    TestContext.WriteLine("H2C native synthetic count={0}; hierarchy_open_ms={1:F3}; expansion_ms={2:F3}; initial_nodes={3}; expanded_nodes={4}; requests=0; writes=0", count, watch.Elapsed.TotalMilliseconds, expansion.Elapsed.TotalMilliseconds, initial, WalkNative(control.Tree).Count());
                }
            });
        }

        [TestMethod]
        public void SyntheticScreenshotsCanBeCapturedWithoutLiveConnections()
        {
            var path = System.Environment.GetEnvironmentVariable("H2C_EVIDENCE_DIRECTORY"); if (string.IsNullOrWhiteSpace(path)) return;
            Sta(() =>
            {
                Directory.CreateDirectory(path); var f = AppSettings();
                using (var form = new MembershipResultsForm(f.Completed))
                {
                    form.Text += " — Synthetic H2C presentation fixture"; Selector(form).SelectedIndex = 0; form.Show(); Save(form, path, "01-flat.png"); Switch(form);
                    var control = form.Controls.OfType<MembershipHierarchyControl>().Single(); control.SelectComparison(2, DisplayEnvironmentSide.Source); Save(form, path, "02-appsettings.png");
                    control.SelectComparison(2, DisplayEnvironmentSide.Target); Save(form, path, "03-target-appsetting.png");
                }
            });
        }

        private sealed class Data { internal MembershipComparisonPresentation Presentation; internal CompletedMembershipComparison Completed; }
        private static string Key(string parent, string name) => AppSettingResolutionOperation.PortableKey(parent, name);
        private static ComponentIdentity Identity(int type, string key, Guid? root = null, IdentityResolutionStatus status = IdentityResolutionStatus.Resolved, bool matchOnly = false) =>
            new ComponentIdentity(new SolutionComponentRecord(Guid.NewGuid(), type, Guid.NewGuid(), rootSolutionComponentId: root), status,
                status == IdentityResolutionStatus.Resolved ? key : null, "Original synthetic diagnostic; no name inferred",
                componentTypeKey: type == 10075 ? ComponentSemanticKinds.AppSetting : type == 9 ? ComponentSemanticKinds.GlobalChoice : null,
                registeredDefinition: type == 10075 ? new SolutionComponentDefinitionIdentity(10075, "AppSetting", "appsetting") : null,
                inventoryAbsencePolicy: matchOnly ? InventoryAbsencePolicy.MatchOnly : InventoryAbsencePolicy.CompleteInventory);
        private static MembershipEnvironmentResult Environment(IEnumerable<ComponentIdentity> items, bool target) => MembershipEnvironmentResult.FromSnapshot(target ? "SYN-THET-IAC" : "SYN-THET-IAB",
            MembershipSnapshot.Complete(new SolutionIdentity(new EnvironmentIdentity(Guid.NewGuid(), target ? "SYN-THET-IAC" : "SYN-THET-IAB"), Guid.NewGuid(), "SAA"), items, new DateTimeOffset(2026, 10, 9, 0, 0, 0, TimeSpan.Zero)), 27, TimeSpan.Zero);
        private static Data Make(IEnumerable<ComponentIdentity> source, IEnumerable<ComponentIdentity> target = null, string names = null, string tableScope = null) =>
            Enrich(new MembershipResultPresenter().Create(Environment(source, false), Environment(target ?? new ComponentIdentity[0], true)), names, tableScope);
        private static Data Enrich(MembershipComparisonPresentation p, string names = null, string tableScope = null)
        {
            Func<ComponentIdentity, DisplayEnvironmentSide, ComponentSideDisplayEvidence> side = (i, role) =>
            {
                var value = names == null || i == null ? new ComponentDisplayValue(DisplayAvailability.NotRetrieved) : new ComponentDisplayValue(DisplayAvailability.Available, names, "Synthetic verified display metadata");
                return new ComponentSideDisplayEvidence(role, i?.Record.SolutionComponentId, i?.Record.ObjectId, i?.Record.ComponentType,
                    new ComponentDisplayValue(DisplayAvailability.NotRetrieved), value, new ComponentDisplayValue(DisplayAvailability.NotRetrieved), DisplayBackingCorrelationStatus.VerifiedMetadata,
                    tableScope == null ? new ComponentDisplayValue(DisplayAvailability.NotRetrieved) : new ComponentDisplayValue(DisplayAvailability.Available, tableScope, "AttributeMetadata.EntityLogicalName"));
            };
            return new Data { Presentation = p, Completed = new CompletedMembershipComparison(p, p.Rows.Select((r, index) => new ComponentDisplayEvidence(index,
                side(r.SourceIdentity, DisplayEnvironmentSide.Source), side(r.TargetIdentity, DisplayEnvironmentSide.Target))), DisplayProjectionOutcome.Completed,
                DisplayCacheAssociation.Verified, DisplayCacheAssociation.Verified) };
        }
        private static Data AppSettings()
        {
            var names = new[] { "OnlineMode", "ChannelAvailability", "EnableRoleBasedOfflineProfileAccess", "msdyn_enablertecopilot", "msdyn_rtecopilotdefaultonruncompleted" };
            var source = new List<ComponentIdentity> { Identity(80, "app") }; var target = new List<ComponentIdentity> { Identity(80, "app") };
            foreach (var name in names) { source.Add(Identity(10075, Key("app", name))); target.Add(Identity(10075, Key("app", name))); }
            var f = Make(source, target);
            var display = f.Completed.DisplayEvidence.Select(e => new ComponentDisplayEvidence(e.RowIndex, AppDisplay(e.SourceDisplayEvidence), AppDisplay(e.TargetDisplayEvidence))).ToList();
            f.Completed = new CompletedMembershipComparison(new CompletedMembershipComparison(f.Presentation, display, DisplayProjectionOutcome.Completed), Parents(f.Presentation.Source.Snapshot, DisplayEnvironmentSide.Source), Parents(f.Presentation.Target.Snapshot, DisplayEnvironmentSide.Target));
            return f;
        }
        private static ComponentSideDisplayEvidence AppDisplay(ComponentSideDisplayEvidence side) => side.RawComponentType != 80 ? side :
            new ComponentSideDisplayEvidence(side.EnvironmentSide, side.SolutionComponentId, side.LocalObjectId, side.RawComponentType, side.ComponentCategory,
                new ComponentDisplayValue(DisplayAvailability.Available, "Field Service Mobile", "Synthetic verified display metadata"), side.InternalName, side.BackingCorrelationStatus, side.ParentDisplayContext);
        private static ComponentHierarchyParentEvidence Parents(MembershipSnapshot snapshot, DisplayEnvironmentSide side)
        {
            var invocation = new object(); var read = new object(); var emitter = new AppSettingParentEvidenceEmission(snapshot, side, invocation, read);
            var parent = snapshot.Components[0].Record.ObjectId.Value;
            emitter.Observe(new AppSettingReadObservation(AppSettingReadKind.Parent, parent, NormalReadDisposition.Unique, 1, parent, "app"));
            foreach (var identity in snapshot.Components.Skip(1))
            {
                var definition = Guid.NewGuid(); string qualifier, name; Assert.IsTrue(MembershipGridRow.AppSettingIdentityNames(identity, out qualifier, out name));
                emitter.Observe(new AppSettingReadObservation(AppSettingReadKind.Child, identity.Record.ObjectId.Value, NormalReadDisposition.Unique, 1, identity.Record.ObjectId,
                    parentId: parent, parentTableVerified: true, definitionId: definition, definitionTableVerified: true));
                emitter.Observe(new AppSettingReadObservation(AppSettingReadKind.Definition, definition, NormalReadDisposition.Unique, 1, definition, name));
            }
            return ComponentHierarchyProjector.RetainAppSettingParentsFromNormalRead(emitter.SealAgainstFinalSnapshot(snapshot, CancellationToken.None), snapshot, side, invocation, read);
        }
        private static IReadOnlyList<MembershipGridRow> Rows(Data f) => MembershipGridRow.CreateRows(f.Presentation, f.Completed, refinedPresentation: true);
        private static MembershipHierarchyView Model(Data f) => MembershipHierarchyView.CreateAdministratorView(ComponentHierarchyProjector.Project(f.Completed, f.Completed.SourceHierarchyParents, f.Completed.TargetHierarchyParents), Rows(f), Enumerable.Range(0, f.Presentation.Rows.Count), true);
        private static string Csv(Data f) => new MembershipCoverageCsvExporter().CreateCsv(f.Presentation);
        private static IEnumerable<MembershipHierarchyNode> Walk(MembershipHierarchyNode root)
        { var stack = new Stack<MembershipHierarchyNode>(); stack.Push(root); while (stack.Count > 0) { var n = stack.Pop(); yield return n; foreach (var c in n.Children.Reverse()) stack.Push(c); } }
        private static IEnumerable<TreeNode> WalkNative(TreeView tree)
        { var stack = new Stack<TreeNode>(tree.Nodes.Cast<TreeNode>().Reverse()); while (stack.Count > 0) { var n = stack.Pop(); yield return n; foreach (var c in n.Nodes.Cast<TreeNode>().Reverse()) stack.Push(c); } }
        private static IEnumerable<Control> Controls(Control root)
        { foreach (Control child in root.Controls) { yield return child; foreach (var descendant in Controls(child)) yield return descendant; } }
        private static ToolStripComboBox Selector(MembershipResultsForm form) => Controls(form).OfType<ToolStrip>().SelectMany(s => s.Items.OfType<ToolStripComboBox>()).Single(c => c.Name == "MembershipViewSelector");
        private static DataGridView Grid(MembershipResultsForm form) => Controls(form).OfType<DataGridView>().Single();
        private static void Switch(MembershipResultsForm form)
        { Selector(form).SelectedIndex = 1; var timeout = Stopwatch.StartNew(); while (form.HierarchyProjection == null && timeout.Elapsed.TotalSeconds < 15) { Application.DoEvents(); Thread.Sleep(1); } Assert.IsNotNull(form.HierarchyProjection, string.Join(";", form.HierarchyDiagnostics)); }
        private static void Save(Form form, string path, string name)
        { Application.DoEvents(); using (var image = new Bitmap(form.Width, form.Height)) { form.DrawToBitmap(image, new Rectangle(Point.Empty, form.Size)); image.Save(Path.Combine(path, name), ImageFormat.Png); } }
        private static void Sta(Action action)
        { Exception error = null; var thread = new Thread(() => { try { action(); } catch (Exception ex) { error = ex; } }) { IsBackground = true }; thread.SetApartmentState(ApartmentState.STA); thread.Start(); Assert.IsTrue(thread.Join(60000)); if (error != null) ExceptionDispatchInfo.Capture(error).Throw(); }
    }
}
