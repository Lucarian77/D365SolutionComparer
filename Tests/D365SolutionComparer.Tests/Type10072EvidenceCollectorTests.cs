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
            Assert.AreEqual(1, pair.Source.Service.Calls); Assert.AreEqual(1, pair.Source.Service.ExecuteCalls);
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
            Assert.AreEqual(2, pair.Source.Service.Calls); Assert.AreEqual(2, pair.Source.Service.ExecuteCalls);
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
            Assert.AreEqual(count, report.Source.Rows.Count); Assert.AreEqual(1, pair.Source.Service.ExecuteCalls);
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
            Assert.AreEqual(1, pair.Source.Service.Calls); Assert.AreEqual(1, pair.Source.Service.ExecuteCalls);
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
            Assert.AreEqual(3, pair.Source.Service.Calls); Assert.AreEqual(1, pair.Source.Service.ExecuteCalls);
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
            Assert.AreEqual(1, pair.Source.Service.ExecuteCalls); Assert.AreEqual(0, pair.Source.Service.WriteCalls);
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
            Assert.AreEqual(1, pair.Source.Service.ExecuteCalls);
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
            internal AppElementEvidenceReport Capture(CancellationToken token = default(CancellationToken)) => new Type10072EvidenceCollector().Capture(Source.Service, Source.Snapshot(), "1", Target.Service, Target.Snapshot(), "2", token);
        }
        private sealed class Fixture
        {
            internal readonly FakeOrganizationService Service;
            internal readonly List<Entity> Rows = new List<Entity>(), ParentRows = new List<Entity>();
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
                        Set(metadata, "Attributes", (req.LogicalName == "appelement" ? Attributes : ParentAttributes).ToArray());
                        Set(metadata, "ManyToOneRelationships", req.LogicalName == "appelement" ? Relationships : new OneToManyRelationshipMetadata[0]);
                        var response = new RetrieveEntityResponse(); response.Results["EntityMetadata"] = metadata; return response;
                    },
                    RetrievePage = query =>
                    {
                        Queries.Add(query); var condition = query.Criteria.Conditions.Single(); Assert.AreEqual(ConditionOperator.In, condition.Operator);
                        Assert.IsTrue(condition.Values.Count <= 200); var ids = condition.Values.Cast<Guid>().ToArray();
                        if (query.EntityName == "appelement" && Page != null) return Page(query);
                        return new EntityCollection((query.EntityName == "appelement" ? Rows : ParentRows).Where(r => ids.Contains(r.Id)).Select(r =>
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
            internal MembershipSnapshot Snapshot() => MembershipSnapshot.Complete(solution, Raw.Concat(Context), DateTimeOffset.UtcNow);
        }
#endif
    }
}
