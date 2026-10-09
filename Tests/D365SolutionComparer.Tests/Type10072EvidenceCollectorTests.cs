using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.ServiceModel;
using System.Threading;
using System.Windows.Forms;
using D365SolutionComparer.Models.Membership;
using D365SolutionComparer.Models.ComponentDetails;
using D365SolutionComparer.Services.Membership;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Query;
using static D365SolutionComparer.Tests.MembershipTestData;

namespace D365SolutionComparer.Tests
{
    [TestClass, TestCategory("Phase2GType10072Evidence")]
    public class Type10072EvidenceCollectorTests
    {
        [TestMethod]
        public void NormalApplicationCompositionDoesNotRetrieveAppElementEvidenceOrAddWhoAmI()
        {
            var solution = Solution(); var raw = ComponentRow(solution, 10072);
            var service = Service(solution, query =>
            {
                if (query.EntityName == "solution") return Rows(SolutionRow(solution));
                if (query.EntityName == "solutioncomponent") return Rows(raw);
                Assert.AreEqual("solutioncomponentdefinition", query.EntityName);
                if (query.Criteria.Conditions.Single().AttributeName == "primaryentityname") return Rows();
                return Rows(new Entity("solutioncomponentdefinition", Guid.NewGuid()) { ["objecttypecode"] = 10072, ["name"] = "AppElement", ["primaryentityname"] = "appelement" });
            });
            var result = SolutionComparerControl.CreateMembershipComparisonOperation().ReadAndResolve(service, solution, CancellationToken.None);
            Assert.AreEqual(IdentityResolutionStatus.Unsupported, result.Membership.Components.Single().Status);
            Assert.IsNull(result.Membership.Components.Single().ComparisonKey);
            Assert.AreEqual(4, service.Calls); Assert.AreEqual(1, service.ExecuteCalls); Assert.AreEqual(0, service.WriteCalls);
        }
        [TestMethod]
        public void NormalMembershipRemainsIndeterminateWithoutAnAppElementDefinitionContract()
        {
            var source = Snapshot(new ComponentIdentity(new SolutionComponentRecord(Guid.NewGuid(), 10072, Guid.NewGuid()),
                IdentityResolutionStatus.Unsupported, semanticKind: "registered:solutioncomponentdefinition:appelement"));
            var target = Snapshot(new ComponentIdentity(new SolutionComponentRecord(Guid.NewGuid(), 10072, Guid.NewGuid()),
                IdentityResolutionStatus.Unsupported, semanticKind: "registered:solutioncomponentdefinition:appelement"));
            var results = new SolutionMembershipComparer().Compare(source, target);
            Assert.IsTrue(results.All(r => r.Presence == MembershipPresence.Indeterminate));
            Assert.IsNull(ComponentDefinitionContractCatalog.For("registered:solutioncomponentdefinition:appelement"));
#if !DEBUG
            var assembly = typeof(SolutionComparerControl).Assembly;
            Assert.IsNull(assembly.GetType("D365SolutionComparer.Services.Membership.Type10072EvidenceCollector"));
            Assert.IsNull(assembly.GetType("D365SolutionComparer.AppElementEvidenceResultsForm"));
            Assert.IsNull(assembly.GetType("D365SolutionComparer.Type10072EvidenceResultsForm"));
            Assert.IsNull(typeof(MembershipResultsForm).GetProperty("CaptureType10072Evidence", BindingFlags.Instance | BindingFlags.NonPublic));
            Assert.IsNull(typeof(SolutionComparerControl).GetMethod("CaptureType10072Evidence", BindingFlags.Instance | BindingFlags.NonPublic));
            Assert.IsFalse(typeof(MembershipCoverageDetailsForm).GetConstructors().SelectMany(c => c.GetParameters()).Any(p => p.Name == "captureType10072Evidence"));
#endif
        }
#if DEBUG
        private const string Secret = "PRIVATE-APPELEMENT-CONFIGURATION-NOT-FOR-EXPORT";
        [TestMethod]
        public void DebugCaptureButtonIsExplicitAndDoesNotRunDuringCoverageFormInitialization()
        {
            Exception error = null;
            var thread = new Thread(() =>
            {
                try
                {
                    var pair = new Pair(); pair.Source.Add(); pair.Target.Add(); var source = pair.Source.Snapshot(); var target = pair.Target.Snapshot();
                    var presentation = new MembershipResultPresenter().Create(MembershipEnvironmentResult.FromSnapshot("Source", source, 4, TimeSpan.Zero),
                        MembershipEnvironmentResult.FromSnapshot("Target", target, 4, TimeSpan.Zero)); int calls = 0;
                    using (var form = new MembershipCoverageDetailsForm("Source", new MembershipCoverageDiagnosticsBuilder().Build(source), "Target",
                        new MembershipCoverageDiagnosticsBuilder().Build(target), presentation, captureType10072Evidence: () => calls++))
                    {
                        form.StartPosition = FormStartPosition.Manual; form.Location = new System.Drawing.Point(-4000, -4000); form.Show(); Application.DoEvents();
                        var button = form.Controls.OfType<FlowLayoutPanel>().SelectMany(p => p.Controls.OfType<Button>()).Single(b => b.Text == "Capture Type 10072 AppElement Evidence...");
                        Assert.IsTrue(button.Enabled); Assert.AreEqual(0, calls); button.PerformClick(); Assert.AreEqual(1, calls);
                    }
                    Assert.AreEqual(0, pair.Source.Service.Calls + pair.Target.Service.Calls);
                    using (var form = new Type10072EvidenceResultsForm("redacted")) Assert.IsTrue(form.Controls.OfType<RichTextBox>().Single().ReadOnly);
                }
                catch (Exception ex) { error = ex; }
            }) { IsBackground = true };
            thread.SetApartmentState(ApartmentState.STA); thread.Start(); Assert.IsTrue(thread.Join(TimeSpan.FromSeconds(30)));
            if (error != null) System.Runtime.ExceptionServices.ExceptionDispatchInfo.Capture(error).Throw();
        }
        [TestMethod]
        public void DistinctParentIdsWithSamePortableNameAreAmbiguousEvenWhenReusedFromSnapshot()
        {
            var pair = new Pair(); pair.Source.Add(); var other = pair.Source.Add(); var parent = Guid.NewGuid();
            other["appmoduleid"] = new EntityReference("appmodule", parent); pair.Source.Context.Add(Resolved(80, parent, pair.Source.ParentKey));
            var report = pair.Capture(); Assert.IsTrue(report.Source.Rows.Values.All(r => r.ParentStatus == "Ambiguous" && r.CandidateA == null));
        }
        [DataTestMethod, DataRow(80), DataRow(61)]
        public void SnapshotCanonicalCollisionBlocksParentOrReferenceEvenWhenOnlyOneIsReferenced(int type)
        {
            var pair = new Pair(); pair.Source.Add(); pair.Source.Context.Add(Resolved(type, Guid.NewGuid(), type == 80 ? pair.Source.ParentKey : pair.Source.ReferenceKey));
            var report = pair.Capture(); Assert.IsNull(report.Source.Rows.Values.Single().CandidateA);
            Assert.IsFalse(report.Pairs.Any(p => p.Outcome == "SemanticPair"));
        }
        [TestMethod]
        public void UntypedObjectIdDoesNotAssumeComponentTypeSemanticsWithoutRelationshipMetadata()
        {
            var pair = new Pair(); var row = pair.Source.Add();
            pair.Source.Attributes.Add(Attribute("objectid", new UniqueIdentifierAttributeMetadata())); pair.Source.Attributes.Add(Attribute("componenttype", new PicklistAttributeMetadata()));
            row["objectid"] = ((EntityReference)row["componentid"]).Id; row["componenttype"] = new OptionSetValue(61); row.Attributes.Remove("componentid");
            var report = pair.Capture(); Assert.IsNull(report.Source.Rows.Values.Single().CandidateA); StringAssert.Contains(report.Build(), "no referenced table/identity relationship assumed");
        }
        [TestMethod]
        public void ValidatedGuidReferenceReusesOnlyExactCompletedComponentIdentity()
        {
            var pair = new Pair(); var row = pair.Source.Add(); var id = ((EntityReference)row["componentid"]).Id;
            pair.Source.Attributes.RemoveAll(a => a.LogicalName == "componentid"); pair.Source.Attributes.Add(Attribute("componentid", new UniqueIdentifierAttributeMetadata()));
            row["componentid"] = id; pair.Source.Relationships = new[] { new OneToManyRelationshipMetadata { SchemaName = "appelement_reference",
                ReferencingEntity = "appelement", ReferencingAttribute = "componentid", ReferencedEntity = "webresource", ReferencedAttribute = "webresourceid" } };
            var report = pair.Capture(); Assert.IsNotNull(report.Source.Rows.Values.Single().CandidateA); Assert.AreEqual(1, pair.Source.Service.Calls);
        }
        [TestMethod]
        public void AuditOwnerLookupDoesNotReplaceOrObscureTheVerifiedComponentReference()
        {
            var pair = new Pair(); var row = pair.Source.Add(); pair.Source.Attributes.Add(Attribute("createdby", new LookupAttributeMetadata { Targets = new[] { "systemuser" } }));
            row["createdby"] = new EntityReference("systemuser", Guid.NewGuid()); var report = pair.Capture(); Assert.IsNotNull(report.Source.Rows.Values.Single().CandidateA);
        }
        [TestMethod]
        public void LifecycleHashesAndInstallationIdsAreComparedOnlyAfterSemanticPairing()
        {
            var pair = new Pair(); pair.Source.Add(); var target = pair.Target.Add(); target["configjson"] = "different private config";
            var report = pair.Capture(); Assert.IsTrue(report.Pairs.Single().Categories.Contains("DifferentUniqueId"));
            Assert.IsTrue(report.Pairs.Single().Categories.Contains("DifferentDefinition")); Assert.AreEqual("SemanticPair", report.Pairs.Single().Outcome);
            Assert.IsFalse(report.Build().Contains("different private config"));
        }
        [TestMethod]
        public void ZeroMembersStopBeforeMetadataAndBackingReads()
        {
            var pair = new Pair(); var report = pair.Capture();
            Assert.AreEqual(0, pair.Source.Service.Calls + pair.Target.Service.Calls + pair.Source.Service.ExecuteCalls + pair.Target.Service.ExecuteCalls);
            StringAssert.Contains(report.Build(), "raw=0"); Assert.AreEqual(0, report.Pairs.Count);
        }
        [TestMethod]
        public void MetadataValidatedDirectCorrelationAndDifferentIdsProduceUniqueSemanticEvidenceOnly()
        {
            var pair = new Pair(); var left = pair.Source.Add(); var right = pair.Target.Add();
            pair.Source.Rows[0]["ismanaged"] = false; pair.Target.Rows[0]["ismanaged"] = true;
            Assert.AreNotEqual(left.Id, right.Id);
            var before = new SolutionMembershipComparer().Compare(pair.Source.Snapshot(), pair.Target.Snapshot());
            var report = pair.Capture(); var entry = report.Pairs.Single();
            Assert.AreEqual("SemanticPair", entry.Outcome); Assert.AreEqual(left.Id, entry.Source.PrimaryId); Assert.AreEqual(right.Id, entry.Target.PrimaryId);
            Assert.AreEqual(entry.Source.CandidateA, entry.Target.CandidateA); Assert.AreEqual(entry.Source.CandidateB, entry.Target.CandidateB);
            Assert.AreEqual("Unique", entry.Source.ParentStatus); Assert.AreEqual("Unique", entry.Source.ReferenceStatus);
            StringAssert.Contains(report.Build(), "DifferentPrimaryId"); StringAssert.Contains(report.Build(), "UnmanagedToManaged");
            StringAssert.Contains(report.Build(), "ObjectIdEqualsPrimaryId=True");
            Assert.AreEqual(1, pair.Source.Service.Calls); Assert.AreEqual(2, pair.Source.Service.ExecuteCalls);
            Assert.IsFalse(report.Build().Contains(Secret)); StringAssert.Contains(report.Build(), "sha256=");
            var after = new SolutionMembershipComparer().Compare(pair.Source.Snapshot(), pair.Target.Snapshot());
            CollectionAssert.AreEqual(before.Select(r => r.Presence).ToArray(), after.Select(r => r.Presence).ToArray());
            Assert.IsTrue(report.Source.Raw.All(r => r.Status == IdentityResolutionStatus.Unsupported && r.ComparisonKey == null));
            Assert.AreEqual(0, pair.Source.Service.WriteCalls + pair.Target.Service.WriteCalls);
        }
        [TestMethod]
        public void DifferentSemanticCandidatesWithDifferentPrimaryIdsNeverPairByCandidateBOrContentHash()
        {
            var pair = new Pair(); pair.Source.Add(); pair.Target.ReferenceKey = "publisher/changed.js"; pair.Target.Add();
            var report = pair.Capture(); Assert.AreEqual(2, report.Pairs.Count);
            Assert.IsTrue(report.Pairs.All(p => p.Outcome == "OneSidedEvidence"));
            Assert.AreEqual(report.Source.Rows.Values.Single().CandidateB, report.Target.Rows.Values.Single().CandidateB);
            Assert.AreEqual(report.Source.Rows.Values.Single().Content["configjson"].Sha256, report.Target.Rows.Values.Single().Content["configjson"].Sha256);
        }
        [TestMethod]
        public void CandidateACollisionCannotBeRepairedByDifferentBNamesOrGuidOverlap()
        {
            var pair = new Pair(); var first = pair.Source.Add(unique: "first"); var second = pair.Source.Add(unique: "second");
            var secondReference = ((EntityReference)second["componentid"]).Id;
            second["componentid"] = first["componentid"]; pair.Source.Context.RemoveAll(c => c.Record.ObjectId == secondReference);
            pair.Target.Add(id: first.Id, unique: "first");
            var report = pair.Capture(); Assert.IsTrue(report.Source.Rows.Values.All(r => r.DuplicateA));
            Assert.IsTrue(report.Pairs.All(p => p.Outcome == "Ambiguous"));
            Assert.IsFalse(report.Source.Rows.Values.Any(r => r.DuplicateB)); StringAssert.Contains(report.Build(), "Candidate B cannot repair it");
        }
        [TestMethod]
        public void CandidateBCollisionsAreReportedWithoutRepairingOrInvalidatingUniqueA()
        {
            var pair = new Pair(); pair.Source.Add(); pair.Source.ReferenceKey = "publisher/another.js"; pair.Source.Add();
            pair.Target.Add(); pair.Target.ReferenceKey = "publisher/another.js"; pair.Target.Add();
            var report = pair.Capture(); Assert.IsTrue(report.Source.Rows.Values.All(r => r.DuplicateB));
            Assert.IsTrue(report.Pairs.All(p => p.Outcome == "SemanticPair")); Assert.IsFalse(report.Source.Rows.Values.Any(r => r.DuplicateA));
        }
        [DataTestMethod, DataRow("parent"), DataRow("reference"), DataRow("type"), DataRow("unknownReference"), DataRow("uncertainReference")]
        public void MissingContextOrUncertainReferenceLeavesAIncomplete(string missing)
        {
            var pair = new Pair(); var row = pair.Source.Add();
            if (missing == "parent") row.Attributes.Remove("appmoduleid");
            if (missing == "reference") row.Attributes.Remove("componentid");
            if (missing == "type") row.Attributes.Remove("elementtype");
            if (missing == "unknownReference") row["componentid"] = new EntityReference("unknown", Guid.NewGuid());
            if (missing == "uncertainReference") pair.Source.Context[1] = new ComponentIdentity(pair.Source.Context[1].Record, IdentityResolutionStatus.Ambiguous);
            var report = pair.Capture(); Assert.IsNull(report.Source.Rows.Values.Single().CandidateA);
            Assert.AreEqual("Incomplete", report.Pairs.Single().Outcome);
        }
        [DataTestMethod, DataRow("parents"), DataRow("references"), DataRow("parentIdentity")]
        public void MultipleOrAmbiguousContextNeverCreatesA(string kind)
        {
            var pair = new Pair(); var row = pair.Source.Add();
            if (kind == "parents")
            {
                pair.Source.Attributes.Add(Attribute("parentappmoduleid", new LookupAttributeMetadata { Targets = new[] { "appmodule" } }));
                row["parentappmoduleid"] = new EntityReference("appmodule", Guid.NewGuid());
            }
            if (kind == "references")
            {
                pair.Source.Attributes.Add(Attribute("referencedcomponentid", new LookupAttributeMetadata { Targets = new[] { "webresource" } }));
                var id = Guid.NewGuid(); row["referencedcomponentid"] = new EntityReference("webresource", id);
                pair.Source.Context.Add(Resolved(61, id, "different.js"));
            }
            if (kind == "parentIdentity") pair.Source.Context[0] = new ComponentIdentity(pair.Source.Context[0].Record, IdentityResolutionStatus.Ambiguous);
            var report = pair.Capture(); Assert.IsNull(report.Source.Rows.Values.Single().CandidateA);
            Assert.IsFalse(report.Pairs.Any(p => p.Outcome == "SemanticPair"));
        }
        [TestMethod]
        public void ParentAppModuleRetrievalIsMetadataValidatedScopedDeduplicatedAndNeverAnEnvironmentScan()
        {
            var pair = new Pair(); pair.Source.Add(); pair.Source.ReferenceKey = "second.js"; pair.Source.Add();
            pair.Source.Context.RemoveAll(c => c.Record.ComponentType == 80);
            var report = pair.Capture(); Assert.IsTrue(report.Source.Rows.Values.All(r => r.ParentStatus == "Unique"));
            Assert.AreEqual(2, pair.Source.Service.Calls); Assert.AreEqual(3, pair.Source.Service.ExecuteCalls);
            Assert.AreEqual(1, pair.Source.Queries.Single(q => q.EntityName == "appmodule").Criteria.Conditions.Single().Values.Count);
            StringAssert.Contains(report.Build(), "appmodule.uniquename");
        }
        [DataTestMethod, DataRow("blank"), DataRow("missing"), DataRow("duplicate"), DataRow("unreadable")]
        public void ParentBackingUncertaintyBlocksCandidates(string kind)
        {
            var pair = new Pair(); pair.Source.Add(); pair.Source.Context.RemoveAll(c => c.Record.ComponentType == 80);
            if (kind == "blank") pair.Source.Parent["uniquename"] = " ";
            if (kind == "missing") pair.Source.ParentRows.Clear();
            if (kind == "duplicate") pair.Source.ParentRows.Add(pair.Source.Parent);
            if (kind == "unreadable") Set(pair.Source.ParentAttributes.Single(a => a.LogicalName == "uniquename"), "IsValidForRead", false);
            var report = pair.Capture(); Assert.IsNull(report.Source.Rows.Values.Single().CandidateA); Assert.IsNull(report.Source.Rows.Values.Single().CandidateB);
        }
        [TestMethod]
        public void MetadataGuidParentRelationshipRequiresExactPrimaryKeyAndCannotGuessUniqueId()
        {
            var pair = new Pair(); var row = pair.Source.Add(); pair.Source.Context.RemoveAll(c => c.Record.ComponentType == 80);
            pair.Source.Attributes.RemoveAll(a => a.LogicalName == "appmoduleid"); pair.Source.Attributes.Add(Attribute("appmoduleid", new UniqueIdentifierAttributeMetadata()));
            row["appmoduleid"] = pair.Source.Parent.Id;
            pair.Source.Relationships = new[] { new OneToManyRelationshipMetadata { SchemaName = "appelement_parent", ReferencingEntity = "appelement", ReferencingAttribute = "appmoduleid", ReferencedEntity = "appmodule", ReferencedAttribute = "appmoduleid" } };
            var report = pair.Capture(); Assert.AreEqual("Unique", report.Source.Rows.Values.Single().ParentStatus);
            pair.Source.Relationships[0].ReferencedAttribute = "appmoduleidunique";
            report = pair.Capture(); Assert.IsNull(report.Source.Rows.Values.Single().CandidateA);
        }
        [TestMethod]
        public void UnavailableFieldsAreNotQueriedAndDisplayNameAloneCannotProduceCandidates()
        {
            var pair = new Pair(); var row = pair.Source.Add();
            foreach (var field in new[] { "elementtype", "uniquename", "componentid" }) Set(pair.Source.Attributes.Single(a => a.LogicalName == field), "IsValidForRead", false);
            var report = pair.Capture(); Assert.IsNull(report.Source.Rows.Values.Single().CandidateA); Assert.IsNull(report.Source.Rows.Values.Single().CandidateB);
            Assert.IsFalse(pair.Source.Queries.Single().ColumnSet.Columns.Contains("componentid"));
            StringAssert.Contains(report.Build(), "UnavailableOrNotQueried");
        }
        [DataTestMethod, DataRow(200, 1), DataRow(201, 2)]
        public void DistinctIdBatchingNeverExceeds200AndDoesNotQueryPerElement(int count, int calls)
        {
            var pair = new Pair(); for (int i = 0; i < count; i++) pair.Source.Add();
            var report = pair.Capture(); Assert.AreEqual(calls, pair.Source.Service.Calls);
            Assert.AreEqual(count, report.Source.Rows.Count); Assert.AreEqual(2, pair.Source.Service.ExecuteCalls);
            Assert.IsTrue(pair.Source.Queries.All(q => q.Criteria.Conditions.Single().Values.Count <= 200));
        }
        [TestMethod]
        public void TerminalPagingCookieIsCompleteWithoutAnotherRead()
        {
            var pair = new Pair(); var row = pair.Source.Add(); pair.Source.Page = q => new EntityCollection(new List<Entity> { row }) { MoreRecords = false, PagingCookie = "private-cookie" };
            var report = pair.Capture(); Assert.AreEqual("Unique", report.Source.Rows.Values.Single().Status); Assert.AreEqual(1, pair.Source.Service.Calls);
            StringAssert.Contains(report.Build(), "PagingCookieSupplied=True"); Assert.IsFalse(report.Build().Contains("private-cookie"));
        }
        [DataTestMethod, DataRow(true), DataRow(false)]
        public void GenuinePagingFinalizesOnlyAfterTerminalPageAndDeduplicatesIdenticalCrossPageRows(bool cookie)
        {
            var pair = new Pair(); var one = pair.Source.Add(); var two = pair.Source.Add();
            pair.Source.Page = q => q.PageInfo.PageNumber == 1 ? new EntityCollection(new List<Entity> { one }) { MoreRecords = true, PagingCookie = cookie ? "page1" : null }
                : new EntityCollection(new List<Entity> { one, two }) { MoreRecords = false, PagingCookie = cookie ? "terminal" : null };
            var report = pair.Capture(); Assert.IsTrue(report.Source.Rows.Values.All(r => r.Status == "Unique")); Assert.AreEqual(2, pair.Source.Service.Calls);
            StringAssert.Contains(report.Build(), "pageCount=2");
        }
        [DataTestMethod, DataRow("missing"), DataRow("duplicate"), DataRow("conflict"), DataRow("foreign"), DataRow("null"), DataRow("stalled"), DataRow("fault")]
        public void BackingFailuresCannotProduceSemanticCandidates(string kind)
        {
            var pair = new Pair(); var row = pair.Source.Add();
            if (kind == "missing") pair.Source.Rows.Clear();
            if (kind == "duplicate") pair.Source.Rows.Add(row);
            if (kind == "conflict") row["appelementid"] = Guid.NewGuid();
            if (kind == "foreign") pair.Source.Page = q => Rows(new Entity("appelement", Guid.NewGuid()) { ["appelementid"] = Guid.NewGuid() });
            if (kind == "null") pair.Source.Page = q => null;
            if (kind == "stalled") pair.Source.Page = q => new EntityCollection(new List<Entity> { row }) { MoreRecords = true };
            if (kind == "fault") pair.Source.Page = q => throw new InvalidOperationException(Secret);
            var report = pair.Capture(); Assert.IsNull(report.Source.Rows.Values.Single().CandidateA);
            Assert.AreNotEqual("Unique", report.Source.Rows.Values.Single().Status); Assert.IsFalse(report.Build().Contains(Secret));
        }
        [TestMethod]
        public void ConflictingCrossPagePayloadIsDuplicateNotCollapsed()
        {
            var pair = new Pair(); var row = pair.Source.Add(); var other = pair.Source.Add();
            var conflict = new Entity("appelement", row.Id); foreach (var a in row.Attributes) conflict[a.Key] = a.Value; conflict["name"] = "changed";
            pair.Source.Page = q => q.PageInfo.PageNumber == 1 ? new EntityCollection(new List<Entity> { row }) { MoreRecords = true }
                : new EntityCollection(new List<Entity> { conflict, other });
            var report = pair.Capture(); Assert.AreEqual("Duplicate", report.Source.Rows[row.Id].Status); Assert.IsNull(report.Source.Rows[row.Id].CandidateA);
        }
        [TestMethod]
        public void RepeatedRawMembershipDoesNotBecomeDuplicateBackingOrCandidate()
        {
            var pair = new Pair(); var row = pair.Source.Add(); pair.Source.Reference(row.Id); pair.Target.Add();
            var report = pair.Capture(); Assert.AreEqual(2, report.Source.Raw.Count); Assert.AreEqual(1, report.Source.Rows.Count);
            Assert.AreEqual("SemanticPair", report.Pairs.Single().Outcome); Assert.IsFalse(report.Source.Rows[row.Id].DuplicateA);
            StringAssert.Contains(report.Build(), "RepeatedMembershipOnly");
        }
        [DataTestMethod, DataRow("metadata"), DataRow("backing"), DataRow("parentMetadata"), DataRow("parentBacking"), DataRow("before")]
        public void CancellationAlwaysPropagatesWithoutCompletingEvidence(string stage)
        {
            var pair = new Pair(); pair.Source.Add(); using (var cancel = new CancellationTokenSource())
            {
                if (stage.StartsWith("parent", StringComparison.Ordinal)) pair.Source.Context.RemoveAll(c => c.Record.ComponentType == 80);
                var execute = pair.Source.Service.ExecuteRequest; var read = pair.Source.Service.RetrievePage;
                pair.Source.Service.ExecuteRequest = req => { var value = execute(req); if ((stage == "metadata" && ((RetrieveEntityRequest)req).LogicalName == "appelement") || (stage == "parentMetadata" && ((RetrieveEntityRequest)req).LogicalName == "appmodule")) cancel.Cancel(); return value; };
                pair.Source.Service.RetrievePage = q => { var value = read(q); if ((stage == "backing" && q.EntityName == "appelement") || (stage == "parentBacking" && q.EntityName == "appmodule")) cancel.Cancel(); return value; };
                if (stage == "before") cancel.Cancel();
                Assert.ThrowsException<OperationCanceledException>(() => pair.Capture(cancel.Token));
            }
        }
        [DataTestMethod, DataRow("fault"), DataRow("incomplete"), DataRow("duplicateSchema"), DataRow("parentFault")]
        public void MetadataFaultOrIncompleteSchemaDoesNotGuessColumns(string failure)
        {
            var pair = new Pair(); pair.Source.Add(); var normal = pair.Source.Service.ExecuteRequest;
            if (failure == "parentFault") pair.Source.Context.RemoveAll(c => c.Record.ComponentType == 80);
            pair.Source.Service.ExecuteRequest = req =>
            {
                if (failure == "fault" || failure == "parentFault" && ((RetrieveEntityRequest)req).LogicalName == "appmodule") throw new InvalidOperationException(Secret);
                if (failure == "incomplete") return new RetrieveEntityResponse();
                return normal(req);
            };
            if (failure == "duplicateSchema") pair.Source.Attributes.Add(pair.Source.Attributes[0]);
            var report = pair.Capture(); Assert.IsNull(report.Source.Rows.Values.Single().CandidateA); Assert.IsFalse(report.Build().Contains(Secret));
            if (failure != "parentFault") Assert.AreEqual(0, pair.Source.Service.Calls);
        }
        [TestMethod]
        public void CaseWhitespaceAndLocalInstallationIdsDoNotChangeSemanticCandidates()
        {
            var pair = new Pair(); pair.Source.Add(); pair.Target.ParentKey = "  PUBLISHER_APP  "; pair.Target.ReferenceKey = "PUBLISHER/VIEW.JS"; pair.Target.Add(unique: "  ELEMENT_ONE  ");
            var report = pair.Capture(); Assert.AreEqual("SemanticPair", report.Pairs.Single().Outcome);
            Assert.IsTrue(StringComparer.OrdinalIgnoreCase.Equals(report.Source.Rows.Values.Single().CandidateB, report.Target.Rows.Values.Single().CandidateB));
        }
        [TestMethod]
        public void ReportIncludesAllRequiredSectionsAndNoBinaryOrRawConfigurationIsRequested()
        {
            var pair = new Pair(); pair.Source.Add(); pair.Target.Add(); var report = pair.Capture(); var text = report.Build();
            foreach (var section in new[] { "RAW TYPE 10072 MEMBERSHIP", "BACKING APPELEMENT CORRELATION", "READABLE / UNAVAILABLE SCHEMA", "PARENT / REFERENCE RELATIONSHIP DISCOVERY", "CANDIDATE IDENTITY ANALYSIS", "DUPLICATE / REPEATED ANALYSIS", "SOURCE / TARGET FIELD COMPARISON", "LIFECYCLE CORRELATION MATRIX", "DIFFERING PRIMARY-ID SEMANTIC PAIRS", "PORTABILITY ASSESSMENT", "EXACT REQUEST LEDGER" }) StringAssert.Contains(text, section);
            Assert.IsFalse(text.Contains(Secret)); Assert.IsFalse(pair.Source.Queries.Single().ColumnSet.Columns.Contains("binarycontent"));
            Assert.IsFalse(pair.Source.Queries.Single().ColumnSet.Columns.Contains("attachment")); Assert.AreEqual(text, report.Build());
            StringAssert.Contains(text, "AdditionalWhoAmI=0"); StringAssert.Contains(text, "NormalMembershipEvidenceRequests=0");
        }
        [TestMethod]
        public void FullQuerySuccessKeepsSingleBatchedReadAndDoesNotRetry()
        {
            var pair = new Pair(); pair.Source.Add(); var report = pair.Capture();
            Assert.AreEqual(1, pair.Source.Service.Calls); Assert.AreEqual(2, pair.Source.Service.ExecuteCalls);
            Assert.AreEqual("Unique", report.Source.Rows.Values.Single().Status);
            Assert.IsFalse(report.Build().Contains("Minimal AppElement retry"));
        }
        [TestMethod]
        public void ReadableLookupShadowsAreExcludedWhileLookupsAndHashOnlyPublishConfigurationRemain()
        {
            var pair = new Pair(); pair.Source.Add(); AddRetrySchema(pair.Source); AddLookupShadows(pair.Source);
            var report = pair.Capture(); var query = pair.Source.Queries.Single(); var row = report.Source.Rows.Values.Single();
            foreach (var field in ShadowFields.Concat(new[] { "diagnosticlookupname", "diagnosticlookupyominame" }))
            {
                Assert.IsFalse(query.ColumnSet.Columns.Contains(field));
                StringAssert.Contains(report.Build(), "appelement." + field + "; type=String; readable=True; capture=ExcludedLookupShadow");
            }
            foreach (var lookup in new[] { "canvasappid", "createdby", "createdonbehalfby", "modifiedby", "modifiedonbehalfby", "organizationid", "parentappmoduleid", "diagnosticlookup" })
                Assert.IsTrue(query.ColumnSet.Columns.Contains(lookup));
            Assert.IsTrue(query.ColumnSet.Columns.Contains("name")); Assert.IsTrue(query.ColumnSet.Columns.Contains("uniquename"));
            Assert.IsTrue(query.ColumnSet.Columns.Contains("publishconfiguration")); Assert.IsTrue(row.Content["publishconfiguration"].Known);
            Assert.IsFalse(report.Build().Contains(Secret)); Assert.IsFalse(report.Build().Contains("Minimal AppElement retry"));
            Assert.AreEqual("Unique", row.Status); Assert.IsNotNull(row.CandidateA);
        }
        [DataTestMethod, DataRow("uniquename", false), DataRow("publishconfiguration", true)]
        public void UnexpectedFaultStillInvokesBoundedFallbackWithLookupShadowsExcluded(string badColumn, bool criticalComplete)
        {
            var pair = new Pair(); pair.Source.Add(); AddRetrySchema(pair.Source); AddLookupShadows(pair.Source); FailColumns(pair.Source, new[] { badColumn });
            var report = pair.Capture(); var row = report.Source.Rows.Values.Single();
            Assert.AreEqual("Unique", row.Status); Assert.AreEqual(criticalComplete, row.CriticalComplete);
            Assert.AreEqual(criticalComplete, row.CandidateA != null);
            StringAssert.Contains(report.Build(), "Minimal AppElement retry"); StringAssert.Contains(report.Build(), "Attribute faulted: " + badColumn);
            Assert.IsTrue(pair.Source.Service.Calls <= Type10072EvidenceCollector.MaxIsolationGroups + 3);
            Assert.IsTrue(pair.Source.Queries.All(q => !ShadowFields.Any(q.ColumnSet.Columns.Contains) &&
                !q.ColumnSet.Columns.Contains("diagnosticlookupname") && !q.ColumnSet.Columns.Contains("diagnosticlookupyominame")));
            Assert.AreEqual(0, pair.Source.Service.WriteCalls);
        }
        [TestMethod]
        public void FullFaultThenMinimalSuccessSeparatesOptionalReadsAndRestoresEvidence()
        {
            var pair = new Pair(); pair.Source.Add(); AddRetrySchema(pair.Source);
            var normal = pair.Source.Service.RetrievePage; int count = 0;
            pair.Source.Service.RetrievePage = q => { var rows = normal(q); if (q.EntityName == "appelement" && ++count == 1) throw SdkFault(); return rows; };
            var report = pair.Capture(); var row = report.Source.Rows.Values.Single();
            Assert.AreEqual("Unique", row.Status); Assert.IsTrue(row.CriticalComplete); Assert.IsNotNull(row.CandidateA);
            Assert.AreEqual(3, pair.Source.Service.Calls); Assert.AreEqual(2, pair.Source.Service.ExecuteCalls);
            CollectionAssert.AreEquivalent(new[] { "appelementid", "parentappmoduleid", "objectid", "objectidtype", "name", "uniquename", "componentidunique", "componentstate", "ismanaged", "canvasappid" },
                pair.Source.Queries[1].ColumnSet.Columns.ToArray());
            Assert.IsFalse(pair.Source.Queries[1].ColumnSet.Columns.Contains("publishconfiguration"));
            Assert.IsTrue(pair.Source.Queries[2].ColumnSet.Columns.Contains("publishconfiguration"));
        }
        [DataTestMethod, DataRow(1), DataRow(2)]
        public void OptionalAttributeFaultsAreIsolatedWithoutLosingPrimaryOrIdentityEvidence(int badCount)
        {
            var pair = new Pair(); pair.Source.Add(); AddRetrySchema(pair.Source);
            var bad = badCount == 1 ? new[] { "publishconfiguration" } : new[] { "publishconfiguration", "displayname" };
            FailColumns(pair.Source, bad); var report = pair.Capture(); var row = report.Source.Rows.Values.Single();
            Assert.AreEqual("Unique", row.Status); Assert.IsTrue(row.CriticalComplete); Assert.IsNotNull(row.CandidateA);
            Assert.AreEqual("Unavailable", row.Content["publishconfiguration"].Presence); Assert.IsFalse(row.Content["publishconfiguration"].Known);
            foreach (var field in bad) StringAssert.Contains(report.Build(), "Attribute faulted: " + field);
            Assert.IsTrue(row.Content["configjson"].Known); Assert.IsTrue(pair.Source.Service.Calls <= 67);
            Assert.AreEqual(2, pair.Source.Service.ExecuteCalls); Assert.AreEqual(0, pair.Source.Service.WriteCalls);
        }
        [TestMethod]
        public void IdentityCriticalFaultRetainsPrimaryAndOtherFieldsButBlocksCandidates()
        {
            var pair = new Pair(); var backing = pair.Source.Add(); AddRetrySchema(pair.Source); FailColumns(pair.Source, new[] { "uniquename" });
            var report = pair.Capture(); var row = report.Source.Rows.Values.Single();
            Assert.AreEqual("Unique", row.Status); Assert.AreEqual(backing.Id, row.PrimaryId); Assert.IsFalse(row.CriticalComplete);
            Assert.IsNull(row.CandidateA); Assert.IsNull(row.CandidateB); Assert.IsNull(row.Get("uniquename")); Assert.IsNotNull(row.Get("name"));
            StringAssert.Contains(report.Build(), "IdentityCriticalUnavailable");
            Assert.IsFalse(pair.Source.Queries.Skip(2).Any(q => q.ColumnSet.Columns.Contains("publishconfiguration")));
        }
        [TestMethod]
        public void MinimalGroupFaultCanRecoverWhenPrimaryAndIsolatedGroupsSucceed()
        {
            var pair = new Pair(); pair.Source.Add(); AddRetrySchema(pair.Source);
            var normal = pair.Source.Service.RetrievePage; int count = 0;
            pair.Source.Service.RetrievePage = q => { var rows = normal(q); if (q.EntityName == "appelement" && ++count <= 2) throw SdkFault(); return rows; };
            var report = pair.Capture(); var row = report.Source.Rows.Values.Single();
            CollectionAssert.AreEqual(new[] { "appelementid" }, pair.Source.Queries[2].ColumnSet.Columns.ToArray());
            Assert.AreEqual("Unique", row.Status); Assert.IsTrue(row.CriticalComplete); Assert.IsNotNull(row.CandidateA);
            Assert.AreEqual(5, pair.Source.Service.Calls);
        }
        [TestMethod]
        public void PrimaryOnlyFaultStopsIsolationAndPreventsParentReads()
        {
            var pair = new Pair(); pair.Source.Add(); pair.Source.Context.RemoveAll(c => c.Record.ComponentType == 80); AddRetrySchema(pair.Source);
            FailColumns(pair.Source, new[] { "appelementid" }); var report = pair.Capture();
            Assert.AreEqual("Faulted", report.Source.Rows.Values.Single().Status); Assert.IsNull(report.Source.Rows.Values.Single().CandidateA);
            Assert.AreEqual(3, pair.Source.Service.Calls); Assert.AreEqual(1, pair.Source.Service.ExecuteCalls);
            Assert.IsTrue(pair.Source.Queries.All(q => q.EntityName == "appelement"));
        }
        [TestMethod]
        public void FaultReportIncludesOnlySafeSdkCodeAndClassNotServiceMessageOrTrace()
        {
            var pair = new Pair(); pair.Source.Add(); AddRetrySchema(pair.Source); FailColumns(pair.Source, new[] { "publishconfiguration" });
            var text = pair.Capture().Build(); StringAssert.Contains(text, "FaultType=OrganizationServiceFault"); StringAssert.Contains(text, "SDKErrorCode=0x80040216");
            Assert.IsFalse(text.Contains(Secret)); Assert.IsFalse(text.Contains("https://")); Assert.IsFalse(text.Contains("Bearer")); Assert.IsFalse(text.Contains("STACK-TRACE"));
        }
        [TestMethod]
        public void EveryRetryKeepsTheSelectedIdsAndMaximum200BatchBoundary()
        {
            var pair = new Pair(); for (int i = 0; i < 201; i++) pair.Source.Add(); AddRetrySchema(pair.Source);
            FailColumns(pair.Source, new[] { "publishconfiguration" }); var report = pair.Capture();
            Assert.IsTrue(report.Source.Rows.Values.All(r => r.Status == "Unique"));
            var selected = pair.Source.Rows.Select(r => r.Id).ToHashSet();
            Assert.IsTrue(pair.Source.Queries.All(q => q.Criteria.Conditions.Count == 1 && q.Criteria.Conditions[0].Operator == ConditionOperator.In &&
                q.Criteria.Conditions[0].Values.Count <= 200 && q.Criteria.Conditions[0].Values.Cast<Guid>().All(selected.Contains)));
            Assert.AreEqual(2, pair.Source.Service.ExecuteCalls);
        }
        [TestMethod]
        public void RetryPagingHonorsTerminalCookieAndCrossPageDeduplication()
        {
            var pair = new Pair(); var one = pair.Source.Add(); var two = pair.Source.Add(); AddRetrySchema(pair.Source);
            var normal = pair.Source.Service.RetrievePage;
            pair.Source.Service.RetrievePage = q =>
            {
                var rows = normal(q); if (q.ColumnSet.Columns.Contains("publishconfiguration")) throw SdkFault();
                if (q.ColumnSet.Columns.Contains("name") && !q.ColumnSet.Columns.Contains("configjson"))
                {
                    var filtered = rows.Entities.Where(r => q.PageInfo.PageNumber != 1 || r.Id == one.Id).ToList();
                    return new EntityCollection(filtered) { MoreRecords = q.PageInfo.PageNumber == 1, PagingCookie = q.PageInfo.PageNumber == 1 ? "retry-page1" : "terminal-private-cookie" };
                }
                return rows;
            };
            var report = pair.Capture(); Assert.IsTrue(report.Source.Rows.Values.All(r => r.Status == "Unique"));
            Assert.IsTrue(pair.Source.Queries.Any(q => q.PageInfo.PageNumber == 2)); StringAssert.Contains(report.Build(), "PagingCookieSupplied=True");
            Assert.IsFalse(report.Build().Contains("terminal-private-cookie"));
        }
        [TestMethod]
        public void CancellationDuringIsolationPropagatesWithoutParentOrFurtherReads()
        {
            var pair = new Pair(); pair.Source.Add(); AddRetrySchema(pair.Source); FailColumns(pair.Source, new[] { "publishconfiguration" });
            using (var cancel = new CancellationTokenSource())
            {
                var normal = pair.Source.Service.RetrievePage;
                pair.Source.Service.RetrievePage = q => { if (q.ColumnSet.Columns.Count == 2 && q.ColumnSet.Columns.Contains("publishconfiguration")) cancel.Cancel(); return normal(q); };
                Assert.ThrowsException<OperationCanceledException>(() => pair.Capture(cancel.Token));
                Assert.AreEqual(1, pair.Source.Service.ExecuteCalls); Assert.AreEqual(0, pair.Source.Service.WriteCalls);
            }
        }
        [TestMethod]
        public void IsolationBudgetExcludesUnprovenOptionalFieldsAndNeverTreatsThemAsBlankHashes()
        {
            var pair = new Pair(); pair.Source.Add(); AddRetrySchema(pair.Source);
            var bad = Enumerable.Range(0, 80).Select(i => "audit" + i.ToString("D2")).ToArray();
            foreach (var field in bad) pair.Source.Attributes.Add(Attribute(field, new StringAttributeMetadata()));
            FailColumns(pair.Source, bad); var report = pair.Capture(); var row = report.Source.Rows.Values.Single();
            Assert.AreEqual("Unique", row.Status); Assert.IsTrue(pair.Source.Service.Calls <= Type10072EvidenceCollector.MaxIsolationGroups + 2);
            StringAssert.Contains(report.Build(), "after isolation limit=64"); Assert.IsTrue(bad.All(f => !row.Content[f].Known));
        }
        [TestMethod]
        public void OptionalFaultNeverProducesFalseSameDefinitionWhenBothSidesLackThatHash()
        {
            var pair = new Pair(); pair.Source.Add(); pair.Target.Add();
            foreach (var side in new[] { pair.Source, pair.Target }) { AddRetrySchema(side); FailColumns(side, new[] { "publishconfiguration" }); }
            var report = pair.Capture(); Assert.AreEqual("SemanticPair", report.Pairs.Single().Outcome);
            Assert.IsFalse(report.Pairs.Single().Categories.Contains("SameDefinition"));
        }
        private static FaultException<OrganizationServiceFault> SdkFault() => new FaultException<OrganizationServiceFault>(
            new OrganizationServiceFault { ErrorCode = unchecked((int)0x80040216), Message = Secret + " https://private.service Bearer TOKEN", TraceText = "STACK-TRACE " + Secret },
            new FaultReason(Secret));
        [TestMethod]
        public void DependencyOnlyCanvasAppsCorrelateExplicitIdsAndCompleteDiagnosticAppElementPair()
        {
            var pair = new Pair(); var left = pair.Source.AddCanvas(); var right = pair.Target.AddCanvas();
            pair.Source.AddCanvasElement(left); pair.Target.AddCanvasElement(right);
            var report = pair.Capture(); var matched = report.Pairs.Single();
            Assert.AreNotEqual(left.Id, right.Id); Assert.AreEqual("SemanticPair", matched.Outcome);
            Assert.AreEqual("VerifiedDependencyOnlyCanvasAppIdentity", matched.Source.CanvasDependencyState);
            Assert.AreEqual("VerifiedDependencyOnlyCanvasAppIdentity", matched.Target.CanvasDependencyState);
            Assert.IsNotNull(matched.Source.CandidateA); Assert.AreNotEqual(matched.Source.PrimaryId, matched.Target.PrimaryId);
            Assert.AreEqual(0, report.Source.CanvasDependencies.Raw.Count); Assert.AreEqual(left.Id, report.Source.CanvasDependencies.Rows.Single().Value.PrimaryId);
            var canvasQueries = pair.Source.Queries.Where(q => q.EntityName == "canvasapp").ToArray(); Assert.AreEqual(2, canvasQueries.Length);
            Assert.IsTrue(canvasQueries.All(q => q.Criteria.Conditions.Single().Values.Cast<Guid>().SequenceEqual(new[] { left.Id })));
            Assert.IsTrue(pair.Source.Raw.Concat(pair.Target.Raw).All(c => c.Status == IdentityResolutionStatus.Unsupported && c.ComparisonKey == null));
        }
        [DataTestMethod, DataRow(true), DataRow(false)]
        public void VerifiedCanvasReferenceRequiresIndependentExplicitElementType(bool hasType)
        {
            var pair = new Pair();
            foreach (var side in new[] { pair.Source, pair.Target })
            {
                var row = side.AddCanvasElement(side.AddCanvas()); row.Attributes.Remove("elementtype");
                side.Attributes.Add(Attribute("objectidtype", new PicklistAttributeMetadata()));
                if (hasType) row["objectidtype"] = new OptionSetValue(300);
            }
            var report = pair.Capture();
            foreach (var row in report.Source.Rows.Values.Concat(report.Target.Rows.Values))
            {
                Assert.AreEqual("VerifiedDependencyOnlyCanvasAppIdentity", row.CanvasDependencyState);
                Assert.IsTrue(row.ParentIdentityComplete); Assert.IsTrue(row.ReferencedComponentIdentityComplete);
                Assert.AreEqual(hasType, row.ElementTypeComplete); Assert.AreEqual(hasType, row.CompleteA);
                Assert.IsNotNull(row.CandidateB); // Descriptive B cannot repair an absent explicit type.
                if (hasType) { Assert.AreEqual("objectidtype", row.ElementTypeField); Assert.AreEqual("None", row.CandidateABlockingReason); }
                else { Assert.IsNull(row.CandidateA); StringAssert.Contains(row.CandidateABlockingReason, "objectidtype are blank or unavailable"); }
            }
            StringAssert.Contains(report.Build(), "ParentIdentityComplete=True");
            StringAssert.Contains(report.Build(), "ReferencedComponentIdentityComplete=True");
            StringAssert.Contains(report.Build(), "ElementTypeComplete=" + hasType);
            StringAssert.Contains(report.Build(), "CompleteA=" + hasType);
            StringAssert.Contains(report.Build(), "BlockingReason=");
        }
        [DataTestMethod, DataRow(false), DataRow(true)]
        public void OrganizationReferenceIsAuditOnlyAndDoesNotBlockVerifiedCanvasCandidate(bool guidRelationship)
        {
            var pair = new Pair();
            foreach (var side in new[] { pair.Source, pair.Target })
            {
                var row = side.AddCanvasElement(side.AddCanvas());
                side.Attributes.Add(Attribute("organizationid", guidRelationship ? (AttributeMetadata)new UniqueIdentifierAttributeMetadata()
                    : new LookupAttributeMetadata { Targets = new[] { "organization" } }));
                row["organizationid"] = guidRelationship ? (object)Guid.NewGuid() : new EntityReference("organization", Guid.NewGuid());
                if (guidRelationship) side.Relationships = new[] { new OneToManyRelationshipMetadata {
                    ReferencingEntity = "appelement", ReferencingAttribute = "organizationid", ReferencedEntity = "organization", ReferencedAttribute = "organizationid" } };
            }
            var report = pair.Capture(); Assert.AreEqual("SemanticPair", report.Pairs.Single().Outcome);
            foreach (var row in report.Source.Rows.Values.Concat(report.Target.Rows.Values))
            {
                Assert.IsNotNull(row.Get("organizationid")); Assert.IsFalse(row.ReferenceUncertain);
                Assert.AreEqual("Unique", row.ReferenceStatus); Assert.IsTrue(row.CompleteA);
                Assert.IsTrue(row.ParentIdentityComplete && row.ElementTypeComplete && row.ReferencedComponentIdentityComplete);
            }
            Assert.IsFalse(report.Build().Contains("Reference organization has no already-verified snapshot identity mapping"));
            Assert.IsTrue(pair.Source.Queries.Concat(pair.Target.Queries).All(q => q.EntityName != "organization"));
        }
        [TestMethod]
        public void VerifiedCanvasDependencyDoesNotRepairIncompleteParentIdentity()
        {
            var pair = new Pair(); var left = pair.Source.AddCanvasElement(pair.Source.AddCanvas()); pair.Target.AddCanvasElement(pair.Target.AddCanvas());
            left.Attributes.Remove("appmoduleid"); var row = pair.Capture().Source.Rows.Values.Single();
            Assert.AreEqual("VerifiedDependencyOnlyCanvasAppIdentity", row.CanvasDependencyState);
            Assert.IsFalse(row.ParentIdentityComplete); Assert.IsTrue(row.ElementTypeComplete && row.ReferencedComponentIdentityComplete);
            Assert.IsFalse(row.CompleteA); StringAssert.Contains(row.CandidateABlockingReason, "Parent AppModule identity incomplete");
        }
        [TestMethod]
        public void CompletedMemberEvidenceAndDependencyOnlyEvidenceCompleteBothDistinctPairsWithoutMemberRequery()
        {
            var pair = TwoCanvasPairs(); var cached = pair.CaptureMembers(); int before = pair.Source.Queries.Count;
            var report = pair.Capture(cached: cached); Assert.AreEqual(2, report.Pairs.Count(p => p.Outcome == "SemanticPair"));
            Assert.AreEqual(1, report.Source.Rows.Values.Count(r => r.CanvasDependencyState == "VerifiedType300SemanticIdentity"));
            Assert.AreEqual(1, report.Source.Rows.Values.Count(r => r.CanvasDependencyState == "VerifiedDependencyOnlyCanvasAppIdentity"));
            Assert.IsTrue(report.Source.Rows.Values.All(r => r.CandidateA != null && !r.DuplicateA));
            Assert.IsTrue(report.Source.Rows.Values.Concat(report.Target.Rows.Values).All(r => r.ParentIdentityComplete &&
                r.ElementTypeComplete && r.ReferencedComponentIdentityComplete && r.CompleteA && r.CandidateABlockingReason == "None"));
            var memberId = pair.Source.Context.Single(c => c.Record.ComponentType == 300).Record.ObjectId.Value;
            Assert.IsTrue(pair.Source.Queries.Skip(before).Where(q => q.EntityName == "canvasapp").All(q => !q.Criteria.Conditions.Single().Values.Contains(memberId)));
            Assert.AreEqual(2, report.Source.Requests.Count(r => r.StartsWith("RetrieveMultiple canvasapp;")));
            Assert.IsFalse(report.Source.Requests.Any(r => r.StartsWith("Execute RetrieveEntity(canvasapp,")));
            StringAssert.Contains(report.Build(), "UniqueDifferingPrimaryIdPairs=2"); StringAssert.Contains(report.Build(), "CanvasCandidateACollisionGroups=0");
            Assert.IsFalse(cached.Source.Rows.Values.Any(r => r.DuplicateA)); Assert.IsFalse(cached.Target.Rows.Values.Any(r => r.DuplicateA));
        }
        [TestMethod]
        public void CompletedCanvasCandidatePIsReusedReadOnlyAndAppElementEvidenceCannotRepairIt()
        {
            var pair = TwoCanvasPairs(); var cached = pair.CaptureMembers();
            var member = cached.Source.Rows.Values.Single(); Assert.IsTrue(member.CompleteP);
            member.RuntimeColumns.Remove(Type300EvidenceCollector.ProposedIdentifierField);
            Type300EvidenceCollector.EvaluateCandidateP(member, cached.Source.Snapshot);
            Assert.IsFalse(member.CompleteP); Assert.IsNotNull(member.CandidateA);
            int before = pair.Source.Queries.Count;
            var report = pair.Capture(cached: cached);
            Assert.AreSame(member, report.Source.CanvasReferenceRows[member.ObjectId]); Assert.IsFalse(member.CompleteP);
            Assert.AreEqual(2, report.Pairs.Count(p => p.Outcome == "SemanticPair")); // Existing A evidence is independent from P.
            Assert.IsTrue(pair.Source.Queries.Skip(before).Where(q => q.EntityName == "canvasapp")
                .All(q => !q.Criteria.Conditions.Single().Values.Contains(member.ObjectId)));
            Assert.AreEqual(2, report.Source.Requests.Count(r => r.StartsWith("RetrieveMultiple canvasapp;")));
            Assert.IsTrue(report.Source.CanvasReferenceRows.Values.Where(r => r.ObjectId != member.ObjectId).All(r => r.CompleteP));
        }
        [TestMethod]
        public void DifferentDependencyCandidateCannotPairThroughIdenticalCandidateBOrHash()
        {
            var pair = new Pair(); var left = pair.Source.AddCanvas(name: "left"); var right = pair.Target.AddCanvas(name: "right");
            pair.Source.AddCanvasElement(left); pair.Target.AddCanvasElement(right);
            var report = pair.Capture(); Assert.IsTrue(report.Source.Rows.Values.Concat(report.Target.Rows.Values).All(r => r.CandidateA == null && r.CanvasDependencyState == "Incomplete"));
            Assert.AreEqual(report.Source.CanvasDependencies.Rows.Single().Value.CandidateB, report.Target.CanvasDependencies.Rows.Single().Value.CandidateB);
        }
        [TestMethod]
        public void MissingDependencyCanvasRowDoesNotBecomeProductionMissingOrOneSidedIdentity()
        {
            var pair = new Pair(); var left = pair.Source.AddCanvas(); var right = pair.Target.AddCanvas();
            pair.Source.AddCanvasElement(left); pair.Target.AddCanvasElement(right); pair.Target.CanvasRows.Clear();
            var report = pair.Capture(); Assert.AreEqual("Missing", report.Target.Rows.Values.Single().CanvasDependencyState);
            Assert.IsNull(report.Source.Rows.Values.Single().CandidateA); Assert.IsNull(report.Target.Rows.Values.Single().CandidateA);
            Assert.IsTrue(report.Pairs.All(p => p.Outcome == "Incomplete"));
        }
        [DataTestMethod, DataRow(false), DataRow(true)]
        public void DuplicateOrConflictingDependencyBackingRowsStayAmbiguous(bool conflict)
        {
            var pair = new Pair(); var left = pair.Source.AddCanvas(); var right = pair.Target.AddCanvas();
            pair.Source.AddCanvasElement(left); pair.Target.AddCanvasElement(right);
            pair.Source.CanvasPage = q => { var row = left; var copy = new Entity("canvasapp", row.Id);
                foreach (var field in q.ColumnSet.Columns) if (row.Contains(field)) copy[field] = row[field];
                var extra = new Entity("canvasapp", row.Id); foreach (var field in copy.Attributes) extra[field.Key] = field.Value;
                if (conflict) extra["name"] = "conflict"; return Rows(copy, extra); };
            var report = pair.Capture(); Assert.AreEqual("Ambiguous", report.Source.Rows.Values.Single().CanvasDependencyState);
            Assert.IsNull(report.Source.Rows.Values.Single().CandidateA); Assert.IsNull(report.Target.Rows.Values.Single().CandidateA);
        }
        [TestMethod]
        public void MemberDependencyCandidateCollisionBlocksBothAndCandidateBNeverRepairsIt()
        {
            var pair = TwoCanvasPairs(); foreach (var row in pair.Source.CanvasRows.Concat(pair.Target.CanvasRows)) row["name"] = "collision";
            foreach (var side in new[] { pair.Source, pair.Target }) side.CanvasRows[1]["displayname"] = "Different descriptive B";
            var cached = pair.CaptureMembers(); var report = pair.Capture(cached: cached);
            Assert.IsTrue(report.Source.Rows.Values.Concat(report.Target.Rows.Values).All(r => r.CanvasDependencyState == "Ambiguous" && r.CandidateA == null));
            StringAssert.Contains(report.Build(), "CanvasCandidateACollisionGroups=2");
            Assert.IsTrue(cached.Source.Rows.Values.All(r => !r.DuplicateA));
        }
        [DataTestMethod, DataRow("configuration", true), DataRow("name", false)]
        public void RuntimeDependencyColumnFaultUsesExistingIsolationAndPreservesCriticalSafeguards(string field, bool complete)
        {
            var pair = new Pair(); var left = pair.Source.AddCanvas(); var right = pair.Target.AddCanvas();
            pair.Source.AddCanvasElement(left); pair.Target.AddCanvasElement(right);
            var normal = pair.Source.Service.RetrievePage;
            pair.Source.Service.RetrievePage = q => { var rows = normal(q); if (q.EntityName == "canvasapp" && q.ColumnSet.Columns.Contains(field)) throw SdkFault(); return rows; };
            var report = pair.Capture(); Assert.AreEqual(complete, report.Source.Rows.Values.Single().CandidateA != null);
            Assert.AreEqual(complete, report.Target.Rows.Values.Single().CandidateA != null);
            StringAssert.Contains(report.Build(), "Attribute faulted: " + field); Assert.IsFalse(report.Build().Contains(Secret));
        }
        [TestMethod]
        public void DependencyPrimaryFaultIsReportedFaultedWithoutFallbackIdentity()
        {
            var pair = new Pair(); var left = pair.Source.AddCanvas(); var right = pair.Target.AddCanvas();
            pair.Source.AddCanvasElement(left); pair.Target.AddCanvasElement(right);
            var normal = pair.Source.Service.RetrievePage; pair.Source.Service.RetrievePage = q => { var rows = normal(q); if (q.EntityName == "canvasapp") throw SdkFault(); return rows; };
            var report = pair.Capture(); Assert.AreEqual("Faulted", report.Source.Rows.Values.Single().CanvasDependencyState); Assert.IsNull(report.Source.Rows.Values.Single().CandidateA);
        }
        [DataTestMethod, DataRow(""), DataRow("9c363a41-5f5b-4245-9b66-c16a5d837348")]
        public void BlankOrGuidOnlyDependencyInternalIdentifierStaysIncomplete(string name)
        {
            var pair = new Pair(); pair.Source.AddCanvasElement(pair.Source.AddCanvas(name: name)); pair.Target.AddCanvasElement(pair.Target.AddCanvas(name: name));
            Assert.IsTrue(pair.Capture().Source.Rows.Values.All(r => r.CandidateA == null && r.CanvasDependencyState == "Incomplete"));
        }
        [TestMethod]
        public void CanvasVerificationCannotBypassUncertainOtherComponentReference()
        {
            var pair = new Pair(); var left = pair.Source.AddCanvasElement(pair.Source.AddCanvas()); pair.Target.AddCanvasElement(pair.Target.AddCanvas());
            left["componentid"] = new EntityReference("webresource", Guid.NewGuid()); var report = pair.Capture();
            Assert.AreEqual("VerifiedDependencyOnlyCanvasAppIdentity", report.Source.Rows.Values.Single().CanvasDependencyState);
            Assert.IsNull(report.Source.Rows.Values.Single().CandidateA); Assert.IsTrue(report.Source.Rows.Values.Single().ReferenceUncertain);
        }
        [TestMethod]
        public void MemberCacheMustBelongToExactSnapshotsAndCannotBeGuessedOrRequeried()
        {
            var pair = TwoCanvasPairs(); var cached = pair.CaptureMembers();
            var report = new Type10072EvidenceCollector().Capture(pair.Source.Service, pair.Source.Snapshot(), "1", pair.Target.Service, pair.Target.Snapshot(), "2", CancellationToken.None, completedType300Evidence: cached);
            Assert.AreEqual(1, report.Source.Rows.Values.Count(r => r.CanvasDependencyState == "Incomplete"));
            Assert.AreEqual(1, report.Source.Rows.Values.Count(r => r.CanvasDependencyState == "VerifiedDependencyOnlyCanvasAppIdentity"));
        }
        [TestMethod]
        public void CancellationDuringDependencyRetrievalDoesNotContinueOrChangeMembership()
        {
            var pair = new Pair(); pair.Source.AddCanvasElement(pair.Source.AddCanvas()); pair.Target.AddCanvasElement(pair.Target.AddCanvas());
            using (var cancellation = new CancellationTokenSource())
            {
                var normal = pair.Source.Service.RetrievePage; pair.Source.Service.RetrievePage = q => { if (q.EntityName == "canvasapp") cancellation.Cancel(); return normal(q); };
                Assert.ThrowsException<OperationCanceledException>(() => pair.Capture(cancellation.Token));
                Assert.IsFalse(pair.Target.Queries.Any(q => q.EntityName == "canvasapp"));
                Assert.IsTrue(pair.Source.Raw.All(c => c.Status == IdentityResolutionStatus.Unsupported && c.ComparisonKey == null));
                Assert.AreEqual(0, pair.Source.Service.WriteCalls + pair.Target.Service.WriteCalls);
            }
        }
        [TestMethod]
        public void ExplicitPicklistTypeValueIsVerifiedOnlyThroughSelectedAttributeOptionMetadata()
        {
            var pair = new Pair(); var backing = pair.Source.Add(); TypeOptions(pair.Source, "elementtype", 7);
            var report = pair.Capture(); var row = report.Source.Rows[backing.Id];
            Assert.IsTrue(row.IndependentTypeComplete); Assert.AreEqual("VerifiedExplicitOptionDiscriminatorEvidence", row.IndependentTypeStatus);
            StringAssert.Contains(report.Build(), "NumericValue=7"); StringAssert.Contains(report.Build(), "OptionSetName=appelement_kind_test");
            StringAssert.Contains(report.Build(), "LabelsAuditOnly=[LCID=1033:Fixture type label]");
            Assert.AreEqual("7", row.ElementType); Assert.IsFalse(row.CandidateA.Contains("Fixture type label"));
            Assert.IsTrue(row.TypeAnalysis.Any(a => a.Contains("ActualValue=7") && a.Contains("VerifiedOptionMetadataValue")));
        }

        [DataTestMethod, DataRow("blank"), DataRow("unavailable"), DataRow("unknownOption"), DataRow("ambiguousOption")]
        public void BlankUnavailableUnknownOrAmbiguousTypeCannotBeIndependentlyEstablished(string state)
        {
            var pair = new Pair(); var backing = pair.Source.Add(); TypeOptions(pair.Source, "elementtype", 7);
            if (state == "blank") backing.Attributes.Remove("elementtype");
            if (state == "unavailable") Set(pair.Source.Attributes.Single(a => a.LogicalName == "elementtype"), "IsValidForRead", false);
            if (state == "unknownOption") backing["elementtype"] = new OptionSetValue(9);
            if (state == "ambiguousOption") ((PicklistAttributeMetadata)pair.Source.Attributes.Single(a => a.LogicalName == "elementtype")).OptionSet.Options.Add(new OptionMetadata(new Microsoft.Xrm.Sdk.Label("Other label", 1033), 7));
            var report = pair.Capture(); var row = report.Source.Rows[backing.Id];
            Assert.IsFalse(row.IndependentTypeComplete); StringAssert.Contains(report.Build(), "No independent element-type source found");
            if (state == "blank" || state == "unavailable") { Assert.IsFalse(row.ElementTypeComplete); Assert.IsNull(row.CandidateA); }
            if (state == "unavailable") Assert.IsFalse(pair.Source.Queries.Any(q => q.ColumnSet.Columns.Contains("elementtype")));
        }

        [TestMethod]
        public void ConflictingExplicitFieldsRequireAnIndependentMappingAndCannotBeCollapsedByLabels()
        {
            var pair = new Pair(); var backing = pair.Source.Add(); TypeOptions(pair.Source, "elementtype", 7);
            pair.Source.Attributes.Add(Attribute("objectidtype", new PicklistAttributeMetadata())); TypeOptions(pair.Source, "objectidtype", 8);
            backing["objectidtype"] = new OptionSetValue(8);
            var report = pair.Capture(); var row = report.Source.Rows[backing.Id];
            Assert.IsFalse(row.IndependentTypeComplete); StringAssert.Contains(row.IndependentTypeStatus, "AmbiguousMultipleDiscriminators");
            Assert.AreEqual("7", row.ElementType); // Existing A precedence/construction is unchanged; this investigation does not repair or rewrite it.
            Assert.IsTrue(row.TypeAnalysis.Count(a => a.Contains("VerifiedOptionMetadataValue")) == 2);
        }

        [DataTestMethod, DataRow("canvasapp", true), DataRow("99999", false), DataRow("unavailable_entity", false)]
        public void EntityNameTargetIsIndependentlyVerifiedWithoutInferringAppElementKind(string value, bool verified)
        {
            var pair = new Pair(); var backing = pair.Source.Add(); backing.Attributes.Remove("elementtype");
            pair.Source.Attributes.Add(Attribute("objectidtype", new EntityNameAttributeMetadata())); backing["objectidtype"] = value;
            var normal = pair.Source.Service.ExecuteRequest;
            pair.Source.Service.ExecuteRequest = req => {
                if (((RetrieveEntityRequest)req).LogicalName == "unavailable_entity") throw new InvalidOperationException("private server data");
                return normal(req);
            };
            var report = pair.Capture(); var row = report.Source.Rows[backing.Id];
            Assert.IsFalse(row.IndependentTypeComplete);
            Assert.AreEqual(verified, row.TypeAnalysis.Any(a => a.Contains("VerifiedEntityNameTarget")));
            StringAssert.Contains(report.Build(), "No independent element-type source found");
            Assert.IsFalse(report.Build().Contains("private server data"));
            Assert.IsFalse(pair.Source.Queries.Any(q => q.EntityName == "canvasapp"));
        }

        [TestMethod]
        public void NewPlausibleFieldIsObservedWithoutChangingExistingCandidateConstruction()
        {
            var pair = new Pair(); var backing = pair.Source.Add(); backing.Attributes.Remove("elementtype");
            pair.Source.Attributes.Add(Attribute("appelementtype", new PicklistAttributeMetadata())); TypeOptions(pair.Source, "appelementtype", 3);
            backing["appelementtype"] = new OptionSetValue(3);
            var report = pair.Capture(); var row = report.Source.Rows[backing.Id];
            Assert.IsTrue(row.IndependentTypeComplete); Assert.IsTrue(row.ParentIdentityComplete && row.ReferencedComponentIdentityComplete);
            Assert.IsFalse(row.ElementTypeComplete); Assert.IsNull(row.CandidateA); Assert.IsNotNull(row.CandidateB);
            StringAssert.Contains(report.Build(), "CandidateAConstructionUnchanged=True");
        }

        [TestMethod]
        public void VerifiedTypeCompletesOnlyTypeEvidenceAndCannotSupplyAnUnresolvedParentOrReference()
        {
            var pair = new Pair(); var backing = pair.Source.Add(); TypeOptions(pair.Source, "elementtype", 7);
            pair.Source.Context.Clear(); pair.Source.ParentRows.Clear();
            var report = pair.Capture(); var row = report.Source.Rows[backing.Id];
            Assert.IsTrue(row.IndependentTypeComplete && row.ElementTypeComplete);
            Assert.IsFalse(row.ParentIdentityComplete || row.ReferencedComponentIdentityComplete || row.CompleteA);
        }

        [TestMethod]
        public void CanvasRelationshipAndCachedDependencyCannotManufactureMissingElementType()
        {
            var pair = TwoCanvasPairs(); var cached = pair.CaptureMembers();
            foreach (var side in new[] { pair.Source, pair.Target }) foreach (var row in side.Rows) row.Attributes.Remove("elementtype");
            var report = pair.Capture(cached: cached);
            Assert.IsTrue(report.Source.Rows.Values.Concat(report.Target.Rows.Values).All(r => r.ParentIdentityComplete && r.ReferencedComponentIdentityComplete &&
                !r.ElementTypeComplete && !r.IndependentTypeComplete && !r.CompleteA));
            Assert.AreEqual(2, report.Source.Requests.Count(r => r.StartsWith("RetrieveMultiple canvasapp;"))); // existing critical + optional dependency batches; completed member reused
            Assert.IsTrue(report.Source.Rows.Values.All(r => r.CandidateB != null));
            StringAssert.Contains(report.Build(), "Canvas App/parent/GUID/B/hash evidence cannot supply type");
        }

        [TestMethod]
        public void TypeInvestigationLeavesCandidateKeysAndProductionMembershipUnchanged()
        {
            var pair = new Pair(); var backing = pair.Source.Add(); var before = pair.Capture();
            TypeOptions(pair.Source, "elementtype", 7); var after = pair.Capture();
            Assert.AreEqual(before.Source.Rows[backing.Id].CandidateA, after.Source.Rows[backing.Id].CandidateA);
            Assert.AreEqual(before.Source.Rows[backing.Id].CandidateB, after.Source.Rows[backing.Id].CandidateB);
            Assert.IsTrue(pair.Source.Raw.All(r => r.Status == IdentityResolutionStatus.Unsupported && r.ComparisonKey == null));
            Assert.IsTrue(new SolutionMembershipComparer().Compare(pair.Source.Snapshot(), pair.Target.Snapshot()).Where(r => r.Source?.Record.ComponentType == 10072 || r.Target?.Record.ComponentType == 10072).All(r => r.Presence == MembershipPresence.Indeterminate));
            Assert.AreEqual(0, pair.Source.Service.WriteCalls + pair.Target.Service.WriteCalls);
        }

        [TestMethod]
        public void RegistrationReviewIsMetadataFirstRawTypeScopedAndNeverSuppliesChildType()
        {
            var pair = new Pair(); var backing = pair.Source.Add(); backing.Attributes.Remove("elementtype");
            var normal = pair.Source.Service.ExecuteRequest;
            pair.Source.Service.ExecuteRequest = req => {
                if (((RetrieveEntityRequest)req).LogicalName != "solutioncomponentdefinition") return normal(req);
                var metadata = new EntityMetadata { LogicalName = "solutioncomponentdefinition" };
                Set(metadata, "PrimaryIdAttribute", "solutioncomponentdefinitionid");
                Set(metadata, "Attributes", new[] { Attribute("solutioncomponentdefinitionid", new UniqueIdentifierAttributeMetadata()),
                    Attribute("objecttypecode", new IntegerAttributeMetadata()), Attribute("name", new StringAttributeMetadata()),
                    Attribute("primaryentityname", new StringAttributeMetadata()), Attribute("componenttype", new IntegerAttributeMetadata()) });
                var response = new RetrieveEntityResponse(); response.Results["EntityMetadata"] = metadata; return response;
            };
            var registration = new Entity("solutioncomponentdefinition", Guid.NewGuid()); registration["solutioncomponentdefinitionid"] = registration.Id;
            registration["objecttypecode"] = 10072; registration["name"] = "AppElement"; registration["primaryentityname"] = "appelement"; registration["componenttype"] = 300;
            var normalPage = pair.Source.Service.RetrievePage;
            pair.Source.Service.RetrievePage = q => {
                if (q.EntityName != "solutioncomponentdefinition") return normalPage(q);
                pair.Source.Queries.Add(q); Assert.AreEqual(ConditionOperator.Equal, q.Criteria.Conditions.Single().Operator);
                Assert.AreEqual(10072, q.Criteria.Conditions.Single().Values.Single()); Assert.AreEqual(200, q.PageInfo.Count);
                return new EntityCollection(new List<Entity> { registration }) { PagingCookie = "private-cookie", MoreRecords = false };
            };
            var report = pair.Capture(); Assert.IsFalse(report.Source.Rows[backing.Id].IndependentTypeComplete); Assert.IsNull(report.Source.Rows[backing.Id].CandidateA);
            StringAssert.Contains(report.Build(), "distinctRows=1; pageCount=1"); StringAssert.Contains(report.Build(), "Field=componenttype; Value=300");
            Assert.IsFalse(report.Build().Contains("private-cookie"));
            Assert.AreEqual(report.Source.Requests.Count, pair.Source.Service.Calls + pair.Source.Service.ExecuteCalls);
            Assert.AreEqual(0, pair.Target.Service.Calls + pair.Target.Service.ExecuteCalls);
        }

        [TestMethod]
        public void TypeMetadataReadFaultCannotInvalidateExistingBackingParentOrReferenceEvidence()
        {
            var pair = new Pair(); var backing = pair.Source.Add(); var normal = pair.Source.Service.ExecuteRequest;
            pair.Source.Service.ExecuteRequest = req => ((RetrieveEntityRequest)req).LogicalName == "solutioncomponentdefinition" ? throw new InvalidOperationException("private metadata") : normal(req);
            var report = pair.Capture(); var row = report.Source.Rows[backing.Id];
            Assert.AreEqual("Unique", row.Status); Assert.IsTrue(row.ParentIdentityComplete && row.ReferencedComponentIdentityComplete);
            Assert.IsNotNull(row.CandidateA); Assert.IsFalse(report.Build().Contains("private metadata"));
            StringAssert.Contains(report.Build(), "Registered-definition metadata unavailable");
        }

        [TestMethod]
        public void RuntimeFaultingTypeColumnRemainsUnavailableWithoutLosingBackingCorrelation()
        {
            var pair = new Pair(); var backing = pair.Source.Add(); TypeOptions(pair.Source, "elementtype", 7);
            FailColumns(pair.Source, new[] { "elementtype" });
            var report = pair.Capture(); var row = report.Source.Rows[backing.Id];
            Assert.AreEqual("Unique", row.Status); Assert.IsTrue(row.ParentIdentityComplete && row.ReferencedComponentIdentityComplete);
            Assert.IsFalse(row.IndependentTypeComplete || row.ElementTypeComplete || row.CompleteA);
            StringAssert.Contains(report.Build(), "appelement.elementtype; AttributeType=Picklist; MetadataReadable=True; Selected=True");
            Assert.IsTrue(row.TypeAnalysis.Any(a => a.Contains("Field=elementtype; RuntimeReadable=False; ActualValue=Unavailable")));
            Assert.IsTrue(pair.Source.Queries.All(q => q.Criteria.Conditions.Single().Operator == ConditionOperator.In));
        }

        [TestMethod]
        public void EntityNameInvestigationReusesExactSnapshotCanvasMetadataWithoutIdentityRepair()
        {
            var pair = TwoCanvasPairs(); var cached = pair.CaptureMembers();
            foreach (var side in new[] { pair.Source, pair.Target }) {
                side.Attributes.Add(Attribute("objectidtype", new EntityNameAttributeMetadata()));
                foreach (var row in side.Rows) { row.Attributes.Remove("elementtype"); row["objectidtype"] = "canvasapp"; }
            }
            var report = pair.Capture(cached: cached);
            Assert.IsTrue(report.Source.Rows.Values.All(r => !r.IndependentTypeComplete));
            Assert.IsFalse(report.Source.Requests.Any(r => r.StartsWith("Execute RetrieveEntity(canvasapp,")));
            Assert.AreSame(cached.Source.Rows.Values.Single(), report.Source.CanvasReferenceRows[cached.Source.Rows.Keys.Single()]);
            Assert.IsTrue(report.Source.TypeSchema.Any(v => v.Contains("target metadata reused")));
        }

        [TestMethod]
        public void CancellationDuringRegistrationMetadataStopsInvestigationWithoutFurtherReads()
        {
            var pair = new Pair(); pair.Source.Add(); pair.Target.Add(); var normal = pair.Source.Service.ExecuteRequest;
            using (var cancel = new CancellationTokenSource()) {
                pair.Source.Service.ExecuteRequest = req => { if (((RetrieveEntityRequest)req).LogicalName == "solutioncomponentdefinition") cancel.Cancel(); return normal(req); };
                Assert.ThrowsException<OperationCanceledException>(() => pair.Capture(cancel.Token));
                Assert.AreEqual(0, pair.Target.Service.Calls + pair.Target.Service.ExecuteCalls);
                Assert.IsTrue(pair.Source.Raw.All(r => r.Status == IdentityResolutionStatus.Unsupported && r.ComparisonKey == null));
            }
        }

        private static void TypeOptions(Fixture side, string field, int value)
        {
            var attribute = (PicklistAttributeMetadata)side.Attributes.Single(a => a.LogicalName == field);
            attribute.OptionSet = new OptionSetMetadata { Name = "appelement_kind_test", IsGlobal = false };
            attribute.OptionSet.Options.Add(new OptionMetadata(new Microsoft.Xrm.Sdk.Label("Fixture type label", 1033), value));
        }

        private static Pair TwoCanvasPairs()
        {
            var pair = new Pair(); foreach (var side in new[] { pair.Source, pair.Target })
            {
                side.AddCanvasElement(side.AddCanvas(member: true, name: "ava_casemanagementsystemdefaultcommandlibrary_d992b"), "first");
                side.AddCanvasElement(side.AddCanvas(name: "publisher_dependency_only_canvas"), "second");
            }
            return pair;
        }
        private static readonly string[] ShadowFields = { "canvasappidname", "createdbyname", "createdbyyominame", "createdonbehalfbyname", "createdonbehalfbyyominame",
            "modifiedbyname", "modifiedbyyominame", "modifiedonbehalfbyname", "modifiedonbehalfbyyominame", "organizationidname", "parentappmoduleidname" };
        private static void AddLookupShadows(Fixture side)
        {
            foreach (var field in ShadowFields) side.Attributes.Add(Attribute(field, new StringAttributeMetadata()));
            foreach (var lookup in new[] { "createdby", "createdonbehalfby", "modifiedby", "modifiedonbehalfby", "organizationid", "diagnosticlookup" })
                side.Attributes.Add(Attribute(lookup, new LookupAttributeMetadata()));
            foreach (var suffix in new[] { "name", "yominame" })
            {
                var shadow = Attribute("diagnosticlookup" + suffix, new StringAttributeMetadata());
                Set(shadow, "AttributeOf", "diagnosticlookup"); side.Attributes.Add(shadow);
            }
        }
        private static void AddRetrySchema(Fixture side)
        {
            side.Attributes.Add(Attribute("parentappmoduleid", new LookupAttributeMetadata { Targets = new[] { "appmodule" } }));
            side.Attributes.Add(Attribute("objectid", new UniqueIdentifierAttributeMetadata())); side.Attributes.Add(Attribute("objectidtype", new PicklistAttributeMetadata()));
            side.Attributes.Add(Attribute("componentidunique", new UniqueIdentifierAttributeMetadata())); side.Attributes.Add(Attribute("canvasappid", new LookupAttributeMetadata { Targets = new[] { "canvasapp" } }));
            side.Attributes.Add(Attribute("publishconfiguration", new MemoAttributeMetadata()));
            foreach (var row in side.Rows) row["publishconfiguration"] = Secret;
        }
        private static void FailColumns(Fixture side, string[] bad)
        {
            var normal = side.Service.RetrievePage;
            side.Service.RetrievePage = q => { var rows = normal(q); if (q.EntityName == "appelement" && bad.Any(q.ColumnSet.Columns.Contains)) throw SdkFault(); return rows; };
        }
        private static ComponentIdentity Resolved(int type, Guid id, string key) => new ComponentIdentity(new SolutionComponentRecord(Guid.NewGuid(), type, id), IdentityResolutionStatus.Resolved, key);
        private static AttributeMetadata Attribute(string name, AttributeMetadata attribute) { attribute.LogicalName = name; Set(attribute, "IsValidForRead", true); return attribute; }
        private static void Set(object obj, string name, object value) => obj.GetType().GetProperty(name).SetValue(obj, value, null);
        private sealed class Pair
        {
            internal readonly Fixture Source = new Fixture(), Target = new Fixture();
            internal AppElementEvidenceReport Capture(CancellationToken token = default(CancellationToken), CanvasAppEvidenceReport cached = null) =>
                new Type10072EvidenceCollector().Capture(Source.Service, cached?.Source.Snapshot ?? Source.Snapshot(), "1", Target.Service, cached?.Target.Snapshot ?? Target.Snapshot(), "2", token, completedType300Evidence: cached);
            internal CanvasAppEvidenceReport CaptureMembers() => new Type300EvidenceCollector().Capture(Source.Service, Source.Snapshot(), "1", Target.Service, Target.Snapshot(), "2", CancellationToken.None);
        }
        private sealed class Fixture
        {
            internal readonly FakeOrganizationService Service;
            internal readonly List<Entity> Rows = new List<Entity>(), ParentRows = new List<Entity>();
            internal readonly List<Entity> CanvasRows = new List<Entity>();
            internal readonly List<AttributeMetadata> CanvasAttributes = new List<AttributeMetadata>();
            internal Func<QueryExpression, EntityCollection> CanvasPage;
            internal readonly List<ComponentIdentity> Raw = new List<ComponentIdentity>(), Context = new List<ComponentIdentity>();
            internal readonly List<QueryExpression> Queries = new List<QueryExpression>();
            internal readonly List<AttributeMetadata> Attributes = new List<AttributeMetadata>(), ParentAttributes = new List<AttributeMetadata>();
            internal OneToManyRelationshipMetadata[] Relationships = new OneToManyRelationshipMetadata[0];
            internal Func<QueryExpression, EntityCollection> Page;
            internal string ParentKey = "publisher_app", ReferenceKey = "publisher/view.js";
            internal readonly Entity Parent = new Entity("appmodule", Guid.NewGuid());
            private readonly D365SolutionComparer.Models.Identity.SolutionIdentity solution = Solution();
            internal Fixture()
            {
                CanvasAttributes.Add(Attribute("canvasappid", new UniqueIdentifierAttributeMetadata()));
                foreach (var field in new[] { "name", "displayname", "uniquecanvasappid" }) CanvasAttributes.Add(Attribute(field, new StringAttributeMetadata()));
                CanvasAttributes.Add(Attribute("configuration", new MemoAttributeMetadata())); CanvasAttributes.Add(Attribute("appcomponenttype", new PicklistAttributeMetadata()));
                Attributes.Add(Attribute("appelementid", new UniqueIdentifierAttributeMetadata()));
                Attributes.Add(Attribute("appelementidunique", new UniqueIdentifierAttributeMetadata()));
                foreach (var field in new[] { "name", "uniquename", "logicalname", "displayname" }) Attributes.Add(Attribute(field, new StringAttributeMetadata()));
                Attributes.Add(Attribute("elementtype", new PicklistAttributeMetadata())); Attributes.Add(Attribute("ismanaged", new BooleanAttributeMetadata()));
                Attributes.Add(Attribute("componentstate", new PicklistAttributeMetadata()));
                Attributes.Add(Attribute("appmoduleid", new LookupAttributeMetadata { Targets = new[] { "appmodule" } }));
                Attributes.Add(Attribute("componentid", new LookupAttributeMetadata { Targets = new[] { "webresource" } }));
                Attributes.Add(Attribute("configjson", new MemoAttributeMetadata()));
                Attributes.Add(Attribute("binarycontent", new StringAttributeMetadata())); Attributes.Add(Attribute("attachment", new StringAttributeMetadata()));
                ParentAttributes.Add(Attribute("appmoduleid", new UniqueIdentifierAttributeMetadata())); ParentAttributes.Add(Attribute("uniquename", new StringAttributeMetadata()));
                Parent["appmoduleid"] = Parent.Id; Parent["uniquename"] = ParentKey; ParentRows.Add(Parent);
                Service = new FakeOrganizationService
                {
                    ExecuteRequest = request =>
                    {
                        Assert.IsInstanceOfType(request, typeof(RetrieveEntityRequest)); var req = (RetrieveEntityRequest)request;
                        Assert.AreEqual(EntityFilters.Attributes | EntityFilters.Relationships, req.EntityFilters); Assert.IsFalse(req.RetrieveAsIfPublished);
                        var metadata = new EntityMetadata { LogicalName = req.LogicalName }; Set(metadata, "PrimaryIdAttribute", req.LogicalName + "id");
                        Set(metadata, "Attributes", (req.LogicalName == "appelement" ? Attributes : req.LogicalName == "canvasapp" ? CanvasAttributes : ParentAttributes).ToArray());
                        Set(metadata, "ManyToOneRelationships", req.LogicalName == "appelement" ? Relationships : new OneToManyRelationshipMetadata[0]);
                        var response = new RetrieveEntityResponse(); response.Results["EntityMetadata"] = metadata; return response;
                    },
                    RetrievePage = query =>
                    {
                        Queries.Add(query); var condition = query.Criteria.Conditions.Single(); Assert.AreEqual(ConditionOperator.In, condition.Operator);
                        Assert.IsTrue(condition.Values.Count <= 200); var ids = condition.Values.Cast<Guid>().ToArray();
                        if (query.EntityName == "appelement" && Page != null) return Page(query);
                        if (query.EntityName == "canvasapp" && CanvasPage != null) return CanvasPage(query);
                        return new EntityCollection((query.EntityName == "appelement" ? Rows : query.EntityName == "canvasapp" ? CanvasRows : ParentRows).Where(r => ids.Contains(r.Id)).Select(r =>
                        { var copy = new Entity(r.LogicalName, r.Id); foreach (var field in query.ColumnSet.Columns) if (r.Contains(field)) copy[field] = r[field]; return copy; }).ToList());
                    }
                };
            }
            internal Entity Add(Guid? id = null, string unique = "element_one")
            {
                var reference = Guid.NewGuid(); var row = new Entity("appelement", id ?? Guid.NewGuid())
                {
                    ["appelementidunique"] = Guid.NewGuid(), ["name"] = "Display element", ["displayname"] = "Localized display",
                    ["uniquename"] = unique, ["elementtype"] = new OptionSetValue(7), ["ismanaged"] = false, ["componentstate"] = new OptionSetValue(0),
                    ["appmoduleid"] = new EntityReference("appmodule", Parent.Id), ["componentid"] = new EntityReference("webresource", reference),
                    ["configjson"] = Secret
                };
                row["appelementid"] = row.Id; Rows.Add(row); Reference(row.Id);
                if (!Context.Any(c => c.Record.ComponentType == 80)) Context.Add(Resolved(80, Parent.Id, ParentKey.Trim()));
                Context.Add(Resolved(61, reference, ReferenceKey)); return row;
            }
            internal void Reference(Guid id) => Raw.Add(new ComponentIdentity(new SolutionComponentRecord(Guid.NewGuid(), 10072, id), IdentityResolutionStatus.Unsupported,
                registeredDefinition: new SolutionComponentDefinitionIdentity(10072, "AppElement", "appelement")));
            internal Entity AddCanvas(bool member = false, string name = "publisher_dependency_canvas")
            {
                var row = new Entity("canvasapp", Guid.NewGuid()) { ["name"] = name, ["displayname"] = "Descriptive canvas",
                    ["uniquecanvasappid"] = Guid.NewGuid().ToString("D"), ["configuration"] = Secret, ["appcomponenttype"] = new OptionSetValue(1) };
                row["canvasappid"] = row.Id; CanvasRows.Add(row);
                if (member) Context.Add(new ComponentIdentity(new SolutionComponentRecord(Guid.NewGuid(), 300, row.Id), IdentityResolutionStatus.Unsupported));
                return row;
            }
            internal Entity AddCanvasElement(Entity canvas, string unique = "canvas_element")
            {
                if (!Attributes.Any(a => a.LogicalName == "canvasappid")) Attributes.Add(Attribute("canvasappid", new LookupAttributeMetadata { Targets = new[] { "canvasapp" } }));
                var row = Add(unique: unique); row.Attributes.Remove("componentid"); row["canvasappid"] = new EntityReference("canvasapp", canvas.Id); return row;
            }
            internal MembershipSnapshot Snapshot() => MembershipSnapshot.Complete(solution, Raw.Concat(Context), DateTimeOffset.UtcNow);
        }
#endif
    }
}
