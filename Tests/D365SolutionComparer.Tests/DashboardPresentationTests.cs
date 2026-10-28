using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Runtime.ExceptionServices;
using System.Threading;
using System.Windows.Forms;
using D365SolutionComparer.Models.Membership;
using D365SolutionComparer.Models.Presentation;
using D365SolutionComparer.Services.Membership;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Microsoft.Xrm.Sdk;

namespace D365SolutionComparer.Tests
{
    [TestClass]
    public class DashboardPresentationTests
    {
        [DataTestMethod]
        [DataRow("equal"), DataRow("different"), DataRow("source-unavailable"), DataRow("target-unavailable")]
        [DataRow("both-unavailable"), DataRow("null"), DataRow("empty")]
        public void DashboardNamesUseIndependentVerifiedLocalNameFields(string scenario)
        {
            var f = Fixture();
            var source = Cache(f.Presentation.Source.Snapshot, out var s);
            var target = Cache(f.Presentation.Target.Snapshot, out var t);
            var left = Row(f.Presentation.Source.Snapshot, "Offline dashboard alpha", 0);
            var right = Row(f.Presentation.Target.Snapshot, scenario == "different" ? "Offline dashboard beta" : "Offline dashboard alpha", 0);
            if (scenario == "source-unavailable" || scenario == "both-unavailable") left.Attributes.Remove("name");
            if (scenario == "target-unavailable" || scenario == "both-unavailable") right.Attributes.Remove("name");
            if (scenario == "null") { left["name"] = null; right["name"] = null; }
            if (scenario == "empty") { left["name"] = ""; right["name"] = ""; }
            source.MetadataCache.StoreEntityRow(left, left.Id); target.MetadataCache.StoreEntityRow(right, right.Id);
            var projection = Read(f, source, target, s, t); var item = new SolutionObjectNavigation(projection).Items(ObjectsWorkspaceScope.Compare).Single();
            string expected = scenario == "equal" ? "Offline dashboard alpha" : scenario == "different" ? "Source: Offline dashboard alpha \u00b7 Target: Offline dashboard beta" :
                scenario == "source-unavailable" ? "Source: [Not available] \u00b7 Target: Offline dashboard alpha" :
                scenario == "target-unavailable" ? "Source: Offline dashboard alpha \u00b7 Target: [Not available]" : "Source: [Not available] \u00b7 Target: [Not available]";
            Assert.AreEqual(expected, item.Value(true));
            Assert.AreEqual("retained_technical_unique", projection.Rows[0].SourceDisplay.InternalName.Value);
            Assert.AreEqual("retained_technical_unique", projection.Rows[0].TargetDisplay.InternalName.Value);
            Assert.AreEqual("systemform.uniquename", projection.Rows[0].SourceDisplay.InternalName.DisplayEvidenceSource);
            Assert.AreEqual("Dashboards", item.Category); Assert.AreEqual(1, item.Rows.Count);
        }

        [DataTestMethod, DataRow(true), DataRow(false)]
        public void SingleSideDashboardNameNeverClaimsTheOtherSide(bool sourceOnly)
        {
            var f = Fixture(sourceOnly ? "source" : "target");
            var context = Cache(sourceOnly ? f.Presentation.Source.Snapshot : f.Presentation.Target.Snapshot, out var spy);
            var row = Row(sourceOnly ? f.Presentation.Source.Snapshot : f.Presentation.Target.Snapshot, "Offline one-side dashboard", 0);
            context.MetadataCache.StoreEntityRow(row, row.Id);
            var projection = Read(f, sourceOnly ? context : null, sourceOnly ? null : context, sourceOnly ? spy : null, sourceOnly ? null : spy);
            var item = new SolutionObjectNavigation(projection).Items(sourceOnly ? ObjectsWorkspaceScope.Source : ObjectsWorkspaceScope.Target).Single();
            Assert.AreEqual(sourceOnly ? "Source: Offline one-side dashboard \u00b7 Target: [Not available]" : "Source: [Not available] \u00b7 Target: Offline one-side dashboard", item.Value(true));
            Assert.AreEqual(1, item.Locals.Count); Assert.AreEqual(sourceOnly ? DisplayEnvironmentSide.Source : DisplayEnvironmentSide.Target, item.Locals[0].Side);
        }

        [DataTestMethod, DataRow(0, "Dashboard"), DataRow(10, "Dashboard"), DataRow(13, "Dashboard"), DataRow(103, "Power BI Embedded")]
        public void VerifiedCachedSubtypeControlsThePresentationLabel(int type, string expected)
        {
            var f = Fixture("source"); var context = Cache(f.Presentation.Source.Snapshot, out var spy);
            // The title deliberately contains the other subtype's wording; it is not a discriminator.
            var row = Row(f.Presentation.Source.Snapshot, type == 103 ? "Ordinary synthetic report" : "Power BI Embedded synthetic title", type);
            context.MetadataCache.StoreEntityRow(row, row.Id); var projection = Read(f, context, null, spy, null);
            var item = new SolutionObjectNavigation(projection).Items(ObjectsWorkspaceScope.Source).Single();
            Assert.AreEqual(expected, item.Type); Assert.AreEqual("Dashboards", item.Category);
            Assert.AreEqual(type, projection.Rows[0].SourceDisplay.MakerObject.DashboardFormType);
            Assert.AreEqual(row["name"], item.Locals[0].LogicalName);
        }

        [DataTestMethod, DataRow("missing"), DataRow("unknown"), DataRow("malformed")]
        public void UnverifiedFormSubtypeDoesNotInventDashboardClassification(string scenario)
        {
            var f = Fixture("source"); var context = Cache(f.Presentation.Source.Snapshot, out var spy);
            var row = Row(f.Presentation.Source.Snapshot, "Power BI Embedded is only a title", 0);
            if (scenario == "missing") row.Attributes.Remove("type");
            else row["type"] = scenario == "unknown" ? (object)new OptionSetValue(999) : "103";
            context.MetadataCache.StoreEntityRow(row, row.Id); var projection = Read(f, context, null, spy, null);
            Assert.IsNull(projection.Rows[0].SourceDisplay.MakerObject);
            Assert.AreEqual(0, new SolutionObjectNavigation(projection).Items(ObjectsWorkspaceScope.Source).Count);
        }

        [TestMethod]
        public void DetachedDashboardWithoutRetainedSubtypeKeepsGeneralLabel()
        {
            var identity = Identity();
            var maker = new MakerObjectDisplayEvidence(identity, MakerObjectKind.Dashboard, identity.Record.ObjectId.Value,
                new ComponentDisplayValue(DisplayAvailability.Available, "Offline dashboard", "systemform.name"),
                new ComponentDisplayValue(DisplayAvailability.NotRetrieved), DisplayBackingCorrelationStatus.Unique);
            Assert.IsTrue(maker.IsFor(identity)); Assert.IsNull(maker.DashboardFormType); Assert.AreEqual("Dashboard", maker.Type);
        }

        [DataTestMethod, DataRow("Resolved"), DataRow("Unsupported"), DataRow("Ambiguous")]
        public void NameProjectionPreservesIdentityReferencesStatusesAndRawMappings(string resolution)
        {
            var status = (IdentityResolutionStatus)Enum.Parse(typeof(IdentityResolutionStatus), resolution);
            var f = Fixture(status: status); var source = Cache(f.Presentation.Source.Snapshot, out var s); var target = Cache(f.Presentation.Target.Snapshot, out var t);
            var left = Row(f.Presentation.Source.Snapshot, "Same display value is not identity", 103);
            var right = Row(f.Presentation.Target.Snapshot, "Same display value is not identity", 103);
            source.MetadataCache.StoreEntityRow(left, left.Id); target.MetadataCache.StoreEntityRow(right, right.Id);
            var originals = f.Presentation.Rows.ToArray(); var projection = Read(f, source, target, s, t);
            Assert.AreEqual(status == IdentityResolutionStatus.Resolved ? 1 : 2, new SolutionObjectNavigation(projection).Items(ObjectsWorkspaceScope.Compare).Count);
            for (int i = 0; i < originals.Length; i++)
            {
                Assert.AreSame(originals[i], projection.Rows[i].Original.AuthoritativeRow);
                Assert.AreSame(originals[i].SourceIdentity, projection.Rows[i].Original.AuthoritativeRow.SourceIdentity);
                Assert.AreSame(originals[i].TargetIdentity, projection.Rows[i].Original.AuthoritativeRow.TargetIdentity);
                Assert.AreEqual(status, (originals[i].SourceIdentity ?? originals[i].TargetIdentity).Status);
            }
            Assert.IsTrue(projection.Rows.All(r => r.SourceOccurrences.Concat(r.TargetOccurrences).All(index =>
                projection.Occurrences[index].Original.ComparisonRowIndices.Contains(r.OriginalRowIndex))));
        }

        [TestMethod]
        public void DifferingVerifiedSubtypesKeepTheirSourceTargetAttribution()
        {
            var f = Fixture(); var source = Cache(f.Presentation.Source.Snapshot, out var s); var target = Cache(f.Presentation.Target.Snapshot, out var t);
            var left = Row(f.Presentation.Source.Snapshot, "Offline alpha", 0); var right = Row(f.Presentation.Target.Snapshot, "Offline alpha", 103);
            source.MetadataCache.StoreEntityRow(left, left.Id); target.MetadataCache.StoreEntityRow(right, right.Id);
            var projection = Read(f, source, target, s, t); var item = new SolutionObjectNavigation(projection).Items(ObjectsWorkspaceScope.Compare).Single();
            Assert.AreEqual("Source: Dashboard \u00b7 Target: Power BI Embedded", item.Type);
            Assert.AreEqual("Offline alpha", item.Value(true));
        }

        [TestMethod]
        public void SwappedOwnerCacheCannotSupplyOppositeSideDashboardName()
        {
            var f = Fixture(); var source = Cache(f.Presentation.Source.Snapshot, out var s); var target = Cache(f.Presentation.Target.Snapshot, out var t);
            var left = Row(f.Presentation.Source.Snapshot, "Verified Source only", 103); source.MetadataCache.StoreEntityRow(left, left.Id);
            new MakerObjectDisplayReader().Read(source, f.Presentation.Source.Snapshot, CancellationToken.None);
            new MakerObjectDisplayReader().Read(target, f.Presentation.Target.Snapshot, CancellationToken.None);
            var good = ComponentDisplayEvidenceProjector.Project(f.Presentation, source.MetadataCache, source.MetadataCache);
            Assert.IsNotNull(good[0].SourceDisplayEvidence.MakerObject); Assert.IsNull(good[0].TargetDisplayEvidence.MakerObject);
            Assert.AreEqual(DisplayAvailability.NotRetrieved, good[0].TargetDisplayEvidence.FriendlyName.Availability);
            var copy = MembershipSnapshot.Complete(f.Presentation.Source.Snapshot.Solution, f.Presentation.Source.Snapshot.Components, f.Presentation.Source.Snapshot.CapturedAt);
            MakerObjectDisplayEvidence unused; Assert.IsFalse(source.MetadataCache.TryGetMakerObjectDisplay(copy, copy.Components[0], out unused));
            Assert.AreEqual(1, s.ExecuteCalls); Assert.AreEqual(1, t.ExecuteCalls); Assert.AreEqual(0, s.Calls + t.Calls);
        }

        [DataTestMethod, DataRow(0), DataRow(1), DataRow(2)]
        public void SelectedDashboardDetailsAndGridUseTheSameReadableName(int scope)
        {
            Sta(() =>
            {
                var f = Fixture(); var source = Cache(f.Presentation.Source.Snapshot, out var s); var target = Cache(f.Presentation.Target.Snapshot, out var t);
                var left = Row(f.Presentation.Source.Snapshot, "Offline dashboard UI", 103); var right = Row(f.Presentation.Target.Snapshot, "Offline dashboard UI", 103);
                left.Attributes.Remove("uniquename"); right.Attributes.Remove("uniquename");
                source.MetadataCache.StoreEntityRow(left, left.Id); target.MetadataCache.StoreEntityRow(right, right.Id);
                var projection = Read(f, source, target, s, t);
                using (var host = new Form { ClientSize = new Size(1100, 650) })
                using (var panel = new SimplifiedSolutionExplorerPanel(f.Presentation, row => { Assert.Fail("No full technical disclosure expected"); return ""; }))
                {
                    panel.UseObjectNavigation(); host.Controls.Add(panel); host.Show(); panel.Bind(projection);
                    Item<ToolStripComboBox>(panel, "SolutionExplorerScope").SelectedIndex = scope;
                    panel.Explorer.SelectComparison(0); Assert.IsNotNull(panel.Explorer.SelectedComparison);
                    Assert.AreSame(f.Presentation.Rows[0], panel.Explorer.SelectedComparison.AuthoritativeRow);
                    Assert.AreEqual("Offline dashboard UI", panel.Explorer.Grid.Rows[0].Cells["Name"].Value);
                    Assert.AreEqual("Power BI Embedded", panel.Explorer.Grid.Rows[0].Cells["Type"].Value);
                    var details = Descendants(panel).OfType<TextBox>().Single(c => c.Name == "SolutionExplorerSelectedIdentity");
                    StringAssert.Contains(details.Text, "Name: Offline dashboard UI"); StringAssert.Contains(details.Text, "Component Type: Power BI Embedded");
                    Assert.IsFalse(details.Text.Contains("Logical / Schema Name:"));
                }
            });
        }

        private static ComponentIdentity Identity(IdentityResolutionStatus status = IdentityResolutionStatus.Resolved) =>
            new ComponentIdentity(new SolutionComponentRecord(Guid.NewGuid(), 60, Guid.NewGuid()), status,
                status == IdentityResolutionStatus.Resolved ? "offline-established-systemform-contract" : null,
                semanticKind: ComponentSemanticKinds.SystemForm);
        private static ExplorerFixture Fixture(string side = "both", IdentityResolutionStatus status = IdentityResolutionStatus.Resolved) =>
            ExplorerFixture.Make(side == "target" ? new ComponentIdentity[0] : new[] { Identity(status) },
                side == "source" ? new ComponentIdentity[0] : new[] { Identity(status) });
        private static DataverseReadContext Cache(MembershipSnapshot owner, out FakeOrganizationService service)
        {
            service = new FakeOrganizationService { ExecuteRequest = request =>
            {
                Assert.IsInstanceOfType(request, typeof(WhoAmIRequest)); var response = new WhoAmIResponse();
                response.Results["OrganizationId"] = owner.Environment.OrganizationId; return response;
            }};
            var context = new DataverseReadContext(service, owner.Environment, CancellationToken.None);
            context.MetadataCache.BeginNormalDisplayNames(owner);
            context.MetadataCache.SealNormalDisplayNames(owner);
            return context;
        }
        private static Entity Row(MembershipSnapshot owner, string name, int type)
        {
            Guid id = owner.Components[0].Record.ObjectId.Value;
            return new Entity("systemform", id) { ["formid"] = id, ["name"] = name, ["uniquename"] = "retained_technical_unique", ["type"] = new OptionSetValue(type) };
        }
        private static SolutionObjectsExplorerProjection Read(ExplorerFixture f, DataverseReadContext source, DataverseReadContext target,
            FakeOrganizationService left, FakeOrganizationService right)
        {
            var original = f.Presentation.Rows.ToArray(); string csv = ExplorerFixture.Csv(f);
            if (source != null) new MakerObjectDisplayReader().Read(source, f.Presentation.Source.Snapshot, CancellationToken.None);
            if (target != null) new MakerObjectDisplayReader().Read(target, f.Presentation.Target.Snapshot, CancellationToken.None);
            f.Completed = new CompletedMembershipComparison(f.Presentation, ComponentDisplayEvidenceProjector.Project(f.Presentation,
                source?.MetadataCache, target?.MetadataCache), DisplayProjectionOutcome.Completed);
            var projection = ExplorerFixture.Project(f); Assert.AreEqual(original.Length, projection.Rows.Count);
            Assert.AreEqual(csv, ExplorerFixture.Csv(f));
            for (int i = 0; i < original.Length; i++) Assert.AreSame(original[i], projection.Rows[i].Original.AuthoritativeRow);
            foreach (var spy in new[] { left, right }.Where(s => s != null)) { Assert.AreEqual(1, spy.ExecuteCalls); Assert.AreEqual(0, spy.Calls + spy.WriteCalls); }
            return projection;
        }
        private static IEnumerable<Control> Descendants(Control root) => root.Controls.Cast<Control>().SelectMany(c => new[] { c }.Concat(Descendants(c)));
        private static T Item<T>(Control root, string name) where T : ToolStripItem => Descendants(root).OfType<ToolStrip>().SelectMany(t => t.Items.OfType<T>()).Single(i => i.Name == name);
        private static void Sta(Action action)
        {
            Exception error = null; var thread = new Thread(() => { try { action(); } catch (Exception e) { error = e; } });
            thread.SetApartmentState(ApartmentState.STA); thread.IsBackground = true; thread.Start(); Assert.IsTrue(thread.Join(60000));
            if (error != null) ExceptionDispatchInfo.Capture(error).Throw();
        }
    }
}
