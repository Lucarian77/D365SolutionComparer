using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using D365SolutionComparer.Models.ComponentDetails;
using D365SolutionComparer.Models.Membership;
using D365SolutionComparer.Models.Presentation;
using D365SolutionComparer.Services.Membership;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Microsoft.Xrm.Sdk;

namespace D365SolutionComparer.Tests
{
    [TestClass]
    public class FinalMakerPresentationTests
    {
        [DataTestMethod, DataRow(0, "Process (Workflow)"), DataRow(1, "Process (Dialog)"), DataRow(3, "Process (Action)"), DataRow(4, "Process (Business Process Flow)")]
        public void ProcessNameAndSubtypeComeFromVerifiedWorkflowFields(int category, string expected)
        {
            var f = ExplorerFixture.Make(new[] { Identity(29, "process") }); var owner = f.Presentation.Source.Snapshot;
            var spy = new FakeOrganizationService { ExecuteRequest = request => { Assert.IsInstanceOfType(request, typeof(WhoAmIRequest)); var response = new WhoAmIResponse(); response.Results["OrganizationId"] = owner.Environment.OrganizationId; return response; } };
            var context = new DataverseReadContext(spy, owner.Environment, CancellationToken.None); var cache = context.MetadataCache;
            cache.BeginNormalDisplayNames(owner); cache.SealNormalDisplayNames(owner); var identity = owner.Components[0]; Guid id = identity.Record.ObjectId.Value;
            cache.WorkflowDefinitions[id] = new Entity("workflow", id) { ["workflowid"] = id, ["name"] = "Synthetic process title", ["uniquename"] = "technical_process_identifier", ["type"] = new OptionSetValue(1), ["category"] = new OptionSetValue(category) };
            new MakerObjectDisplayReader().Read(context, owner, CancellationToken.None);
            var p = Project(f, cache); var item = new SolutionObjectNavigation(p).Items(ObjectsWorkspaceScope.Source).Single();
            StringAssert.Contains(item.Value(true), "Synthetic process title"); Assert.AreEqual(expected, item.Type);
            Assert.AreEqual("technical_process_identifier", p.Rows[0].SourceDisplay.InternalName.Value);
            Assert.AreSame(identity, p.Rows[0].Original.AuthoritativeRow.SourceIdentity);
            Assert.AreEqual(1, spy.ExecuteCalls); Assert.AreEqual(0, spy.Calls + spy.WriteCalls);
        }

        [DataTestMethod, DataRow(-1), DataRow(999)]
        public void MissingOrUnknownProcessSubtypeIsNotInferredFromTitle(int category)
        {
            var identity = Identity(29, "process"); var maker = new MakerObjectDisplayEvidence(identity, MakerObjectKind.Process, identity.Record.ObjectId.Value,
                Available("Business Process Flow is only a title", "workflow.name"), Available("technical", "workflow.uniquename"), DisplayBackingCorrelationStatus.Unique,
                processCategory: category < 0 ? (int?)null : category);
            Assert.AreEqual("Process", maker.Type); Assert.IsTrue(maker.IsFor(identity));
        }

        [DataTestMethod, DataRow("equal"), DataRow("different"), DataRow("source-only"), DataRow("target-only"), DataRow("missing")]
        public void EmailTemplateNameUsesItsRetainedTitleWithoutPairing(string mode)
        {
            var left = mode == "target-only" ? new ComponentIdentity[0] : new[] { Identity(36) };
            var right = mode == "source-only" ? new ComponentIdentity[0] : new[] { Identity(36) };
            var f = ExplorerFixture.Make(left, right); var p = Project(f, Cache(f.Presentation.Source.Snapshot, mode == "missing" ? null : "Synthetic email"),
                Cache(f.Presentation.Target.Snapshot, mode == "missing" ? null : mode == "different" ? "Other synthetic email" : "Synthetic email"));
            if (mode == "missing") { Assert.IsNull(p.Rows[0].PrimaryName(DisplayEnvironmentSide.Source, false)); return; }
            var items = new SolutionObjectNavigation(p).Items(ObjectsWorkspaceScope.Compare); Assert.AreEqual(left.Length + right.Length, items.Count);
            Assert.IsTrue(items.All(i => i.Status(p) == "Needs Review"));
            foreach (var item in items) Assert.IsTrue(item.Value(true).Contains("Synthetic email") || item.Value(true).Contains("Other synthetic email"));
            Assert.IsTrue(items.All(i => i.Locals.Count == 1));
        }

        [TestMethod]
        public void DistinctEmailRecordsWithIdenticalTitlesRemainDistinct()
        {
            var f = ExplorerFixture.Make(new[] { Identity(36), Identity(36) }, new[] { Identity(36), Identity(36) });
            var p = Project(f, Cache(f.Presentation.Source.Snapshot, "Same title"), Cache(f.Presentation.Target.Snapshot, "Same title"));
            var items = new SolutionObjectNavigation(p).Items(ObjectsWorkspaceScope.Compare);
            Assert.AreEqual(4, items.Count); Assert.AreEqual(4, items.Select(i => i.Key).Distinct().Count());
            Assert.IsTrue(items.All(i => i.Status(p) == "Needs Review"));
        }

        [TestMethod]
        public void MaskingDisplayUsesCanonicalMaskObservationInsteadOfForeignWrapperName()
        {
            var f = ExplorerFixture.Make(new[] { Mask() }); var identity = f.Presentation.Rows[0].SourceIdentity;
            var maker = new MakerObjectDisplayEvidence(identity, MakerObjectKind.SecuredMaskingRule, identity.Record.ObjectId.Value,
                Available("Synthetic mask name", "syntheticmask.caption"), new ComponentDisplayValue(DisplayAvailability.NotRetrieved), DisplayBackingCorrelationStatus.Unique);
            var side = new ComponentSideDisplayEvidence(DisplayEnvironmentSide.Source, identity.Record.SolutionComponentId, identity.Record.ObjectId, 74,
                new ComponentDisplayValue(DisplayAvailability.NotRetrieved), Available("Foreign team title", "NormalMembershipDiagnostic.teamtemplate.teamtemplatename"),
                new ComponentDisplayValue(DisplayAvailability.NotRetrieved), DisplayBackingCorrelationStatus.Unique, new ComponentDisplayValue(DisplayAvailability.NotRetrieved), makerObject: maker);
            f.Completed = new CompletedMembershipComparison(f.Presentation, new[] { new ComponentDisplayEvidence(0, side, ExplorerFixture.Display(null, DisplayEnvironmentSide.Target)) }, DisplayProjectionOutcome.Completed);
            var item = new SolutionObjectNavigation(ExplorerFixture.Project(f)).Items(ObjectsWorkspaceScope.Source).Single();
            Assert.AreEqual("Synthetic mask name", item.Value(false)); Assert.IsFalse(item.Value(false).Contains("Foreign team"));
            StringAssert.Contains(item.Value(true), "Synthetic mask name");
        }

        [TestMethod]
        public void EqualLocalIdsAcrossTeamAndMaskCategoriesCannotContaminateNames()
        {
            Guid id = Guid.NewGuid(); var team = new ComponentIdentity(new SolutionComponentRecord(Guid.NewGuid(), 511, id), IdentityResolutionStatus.Unsupported, semanticKind: ComponentSemanticKinds.TeamTemplate);
            var mask = Mask(id); var f = ExplorerFixture.Make(new[] { team, mask }); var p = Project(f, Cache(f.Presentation.Source.Snapshot, "Synthetic retained"));
            var items = new SolutionObjectNavigation(p).Items(ObjectsWorkspaceScope.Source);
            Assert.AreEqual(2, items.Count); Assert.AreEqual(2, items.Select(i => i.Category).Distinct().Count());
            Assert.AreEqual("Synthetic mask", items.Single(i => i.Category == "Secured Masking Rule").Value(false));
            Assert.AreEqual("Synthetic team", items.Single(i => i.Category == "Team template").Value(false));
        }

        [DataTestMethod, DataRow(29), DataRow(36), DataRow(74)]
        public void OppositeSideCacheNeverEstablishesAName(int type)
        {
            var f = ExplorerFixture.Make(new[] { type == 74 ? Mask() : Identity(type) }, new[] { type == 74 ? Mask() : Identity(type) });
            var cache = Cache(f.Presentation.Source.Snapshot, "Source only metadata"); var p = Project(f, cache, cache);
            Assert.IsNull(p.Rows.Single(r => r.TargetDisplay != null).PrimaryName(DisplayEnvironmentSide.Target, false));
        }

        [DataTestMethod, DataRow(2), DataRow(5)]
        public void PresentationDoesNotExpandTheApprovedProcessCategoryWhitelist(int category)
        {
            var f = ExplorerFixture.Make(new[] { Identity(29) }); var owner = f.Presentation.Source.Snapshot;
            var spy = new FakeOrganizationService { ExecuteRequest = request => { var response = new WhoAmIResponse(); response.Results["OrganizationId"] = owner.Environment.OrganizationId; return response; } };
            var context = new DataverseReadContext(spy, owner.Environment, CancellationToken.None); context.MetadataCache.BeginNormalDisplayNames(owner); context.MetadataCache.SealNormalDisplayNames(owner);
            Guid id = owner.Components[0].Record.ObjectId.Value; context.MetadataCache.WorkflowDefinitions[id] = new Entity("workflow", id) { ["workflowid"] = id, ["name"] = "Synthetic process", ["type"] = new OptionSetValue(1), ["category"] = new OptionSetValue(category) };
            new MakerObjectDisplayReader().Read(context, owner, CancellationToken.None); var p = Project(f, context.MetadataCache);
            Assert.AreEqual(0, new SolutionObjectNavigation(p).Items(ObjectsWorkspaceScope.Source).Count); Assert.AreEqual(1, spy.ExecuteCalls); Assert.AreEqual(0, spy.Calls);
        }

        [TestMethod]
        public void SyntheticOwnerCountShapeAndScopesRemain159WithFiveEightFour()
        {
            var left = new List<ComponentIdentity>(); var right = new List<ComponentIdentity>();
            for (int i = 0; i < 142; i++) { left.Add(Identity(61, "web" + i)); right.Add(Identity(61, "web" + i)); }
            for (int i = 0; i < 2; i++) { left.Add(Identity(29, "shared" + i)); right.Add(Identity(29, "shared" + i)); }
            left.Add(Identity(29, "source-a")); left.Add(Identity(29, "source-b")); right.Add(Identity(29, "target-only"));
            for (int i = 0; i < 4; i++) { left.Add(Identity(36)); right.Add(Identity(36)); }
            for (int i = 0; i < 2; i++) { left.Add(Identity(31)); right.Add(Identity(31)); }
            var f = ExplorerFixture.Make(left, right); string csv = ExplorerFixture.Csv(f); var originals = f.Presentation.Rows.ToArray();
            var p = Project(f, Cache(f.Presentation.Source.Snapshot, "Synthetic names"), Cache(f.Presentation.Target.Snapshot, "Synthetic names"));
            var model = new SolutionObjectNavigation(p); var all = model.Items(ObjectsWorkspaceScope.Compare);
            Assert.AreEqual(159, all.Count); Assert.AreEqual(5, all.Count(i => i.Category == "Processes"));
            Assert.AreEqual(8, all.Count(i => i.Category == "Email templates")); Assert.AreEqual(4, all.Count(i => i.Category == "Reports"));
            Assert.AreEqual(152, model.Items(ObjectsWorkspaceScope.Source).Count); Assert.AreEqual(151, model.Items(ObjectsWorkspaceScope.Target).Count);
            Assert.IsTrue(model.Items(ObjectsWorkspaceScope.Source).All(i => i.Has(DisplayEnvironmentSide.Source)));
            Assert.IsFalse(model.Items(ObjectsWorkspaceScope.Source).Any(i => i.Locals.All(l => l.Side == DisplayEnvironmentSide.Target)));
            Assert.AreEqual(csv, ExplorerFixture.Csv(f)); for (int i = 0; i < originals.Length; i++) Assert.AreSame(originals[i], p.Rows[i].Original.AuthoritativeRow);
        }

        [TestMethod]
        public void ReportNameAndPropertyEvidenceContractsRemainIntact()
        {
            var f = ExplorerFixture.Make(new[] { Identity(31) }); var p = Project(f, Cache(f.Presentation.Source.Snapshot, "Synthetic report"));
            Assert.AreEqual("Synthetic report", p.Rows[0].PrimaryName(DisplayEnvironmentSide.Source, false));
            var different = new ComponentPropertyPresentation("defaultvalue", "synthetic-left", "synthetic-right", true, "Different");
            var match = new ComponentPropertyPresentation("type", "same", "same", false, "Match");
            Assert.AreSame(different, SolutionObjectsExplorerRow.OrderProperties(new[] { match, different })[0]);
            Assert.AreEqual("Unavailable", SolutionObjectsExplorerRow.PropertyComparison(new ComponentPropertyPresentation("field", null, null, false, "Match")));
        }

        private static ComponentIdentity Identity(int type, string key = null) => new ComponentIdentity(new SolutionComponentRecord(Guid.NewGuid(), type, Guid.NewGuid()),
            key == null ? IdentityResolutionStatus.Unsupported : IdentityResolutionStatus.Resolved, key,
            semanticKind: type == 29 ? ComponentSemanticKinds.Process : null);
        private static ComponentIdentity Mask(Guid? id = null) => new ComponentIdentity(new SolutionComponentRecord(Guid.NewGuid(), 74, id ?? Guid.NewGuid()),
            IdentityResolutionStatus.Unsupported, registeredDefinition: new SolutionComponentDefinitionIdentity(74, "MaskingRule", "syntheticmask"));
        private static ComponentDisplayValue Available(string value, string provenance) => new ComponentDisplayValue(DisplayAvailability.Available, value, provenance);
        private static DataverseComponentMetadataCache Cache(MembershipSnapshot owner, string title)
        {
            var cache = new DataverseComponentMetadataCache(); cache.BeginNormalDisplayNames(owner);
            foreach (int type in new[] { 31, 36, 511 })
            {
                var records = owner.Components.Where(i => i.Record.ComponentType == type).ToArray(); if (records.Length == 0) continue;
                string table = type == 31 ? "report" : type == 36 ? "template" : "teamtemplate", primary = type == 31 ? "reportid" : type == 36 ? "templateid" : "teamtemplateid", field = type == 36 ? "title" : type == 511 ? "teamtemplatename" : "name";
                var ids = records.Select(i => i.Record.ObjectId.Value).Distinct().ToList(); var rows = ids.ToDictionary(id => id, id => new List<Entity>());
                foreach (Guid id in ids) { var row = new Entity(table, id) { [primary] = id }; if (title != null) row[field] = type == 511 ? "Synthetic team" : title; rows[id].Add(row); }
                cache.ObserveNormalDisplayNames(type, new DiagnosticQueryResult(ids, rows, new Dictionary<Guid, string>(), ids.ToDictionary(id => id, id => new List<Entity>()), ids.Count, true));
            }
            cache.SealNormalDisplayNames(owner); var makers = new Dictionary<SolutionComponentRecord, MakerObjectDisplayEvidence>();
            foreach (var i in owner.Components)
            {
                Guid id = i.Record.ObjectId.Value;
                if (i.Record.ComponentType == 29) makers[i.Record] = new MakerObjectDisplayEvidence(i, MakerObjectKind.Process, id,
                    Available(title, "workflow.name"), Available("technical_process", "workflow.uniquename"), DisplayBackingCorrelationStatus.Unique, processCategory: 4);
                if (i.Record.ComponentType == 74) makers[i.Record] = new MakerObjectDisplayEvidence(i, MakerObjectKind.SecuredMaskingRule, id,
                    Available("Synthetic mask", "syntheticmask.caption"), new ComponentDisplayValue(DisplayAvailability.NotRetrieved), DisplayBackingCorrelationStatus.Unique);
                if (i.Record.ComponentType == 61) cache.StoreEntityRow(new Entity("webresource", id) { ["webresourceid"] = id, ["name"] = i.ComparisonKey, ["displayname"] = i.ComparisonKey }, id);
            }
            cache.SealMakerObjectDisplay(owner, makers, new string[0]); return cache;
        }
        private static SolutionObjectsExplorerProjection Project(ExplorerFixture f, DataverseComponentMetadataCache source, DataverseComponentMetadataCache target = null)
        {
            f.Completed = new CompletedMembershipComparison(f.Presentation, ComponentDisplayEvidenceProjector.Project(f.Presentation, source, target), DisplayProjectionOutcome.Completed); return ExplorerFixture.Project(f);
        }
    }
}
