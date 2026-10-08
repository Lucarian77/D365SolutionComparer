using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.ServiceModel;
using System.Threading;
using System.Windows.Forms;
using D365SolutionComparer.Models.Membership;
using D365SolutionComparer.Services.Membership;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Microsoft.Xrm.Sdk;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Metadata.Query;
using Microsoft.Xrm.Sdk.Query;
using static D365SolutionComparer.Tests.MembershipTestData;

namespace D365SolutionComparer.Tests
{
    [TestClass, TestCategory("Phase2GType10276Evidence")]
    public class Type10276EvidenceCollectorTests
    {
        [TestMethod]
        public void NormalType10276RemainsUnsupportedIndeterminateWithoutProductionKeyOrContract()
        {
            var source = Snapshot(Unsupported()); var target = Snapshot(Unsupported());
            Assert.IsTrue(new SolutionMembershipComparer().Compare(source, target).All(r => r.Presence == MembershipPresence.Indeterminate));
            Assert.IsTrue(source.Components.Concat(target.Components).All(c => c.Status == IdentityResolutionStatus.Unsupported && c.ComparisonKey == null));
            Assert.IsNull(D365SolutionComparer.Models.ComponentDetails.ComponentDefinitionContractCatalog.For(source.Components.Single().SemanticKind));
        }
        private static ComponentIdentity Unsupported() => new ComponentIdentity(new SolutionComponentRecord(Guid.NewGuid(), 10276, Guid.NewGuid()),
            IdentityResolutionStatus.Unsupported, registeredDefinition: new SolutionComponentDefinitionIdentity(10276, "AiSkillConfig", "fixture_ai_config"));
        [TestMethod]
        public void PublicResolverHasNoProductionAiSkillConfigKey()
        {
            var solution = Solution(); var record = new SolutionComponentRecord(Guid.NewGuid(), 10276, Guid.NewGuid());
            var service = Service(solution, q => q.EntityName == "solutioncomponentdefinition" &&
                q.Criteria.Conditions.Any(c => c.AttributeName == "objecttypecode" && c.Values.Contains(10276))
                ? Rows(new Entity("solutioncomponentdefinition", Guid.NewGuid()) { ["objecttypecode"] = 10276,
                    ["name"] = "AISkillConfig", ["primaryentityname"] = "aiskillconfig" }) : Rows());
            var identity = new DataverseComponentIdentityResolver().Resolve(service, solution.Environment, record, CancellationToken.None);
            Assert.AreEqual(IdentityResolutionStatus.Unsupported, identity.Status); Assert.IsNull(identity.ComparisonKey);
            Assert.IsNull(D365SolutionComparer.Models.ComponentDetails.ComponentDefinitionContractCatalog.For(identity.SemanticKind));
            Assert.IsTrue(new SolutionMembershipComparer().Compare(Snapshot(identity), Snapshot()).All(r => r.Presence == MembershipPresence.Indeterminate));
            Assert.AreEqual(0, service.WriteCalls);
        }
        [TestMethod]
        public void CollectorFormAndCapturePathAreAbsentFromRelease()
        {
            var assembly = typeof(SolutionComparerControl).Assembly;
#if DEBUG
            Assert.IsNotNull(assembly.GetType("D365SolutionComparer.Services.Membership.Type10276EvidenceCollector"));
            Assert.IsNotNull(typeof(MembershipResultsForm).GetProperty("CaptureType10276Evidence", BindingFlags.Instance | BindingFlags.NonPublic));
#else
            Assert.IsNull(assembly.GetType("D365SolutionComparer.Services.Membership.Type10276EvidenceCollector"));
            Assert.IsNull(assembly.GetType("D365SolutionComparer.Type10276EvidenceResultsForm"));
            Assert.IsNull(typeof(MembershipResultsForm).GetProperty("CaptureType10276Evidence", BindingFlags.Instance | BindingFlags.NonPublic));
            Assert.IsNull(typeof(SolutionComparerControl).GetMethod("CaptureType10276Evidence", BindingFlags.Instance | BindingFlags.NonPublic));
            Assert.IsFalse(typeof(MembershipCoverageDetailsForm).GetConstructors().SelectMany(c => c.GetParameters()).Any(p => p.Name == "captureType10276Evidence"));
#endif
        }
#if DEBUG
        private const string Secret = "PRIVATE-EXPRESSION-CONTENT-AND-SERVICE-DETAILS";
        private const string EntityName = "fixture_ai_config", Primary = "fixture_configid";
        [TestMethod]
        public void CaptureActionIsExplicitAndCoverageInitializationDoesNotReadDataverse()
        {
            Exception error = null;
            var thread = new Thread(() =>
            {
                try
                {
                    var pair = new Pair(); pair.Source.Add(); pair.Target.Add(); var source = pair.Source.Snapshot(); var target = pair.Target.Snapshot();
                    var presentation = new MembershipResultPresenter().Create(MembershipEnvironmentResult.FromSnapshot("Source", source, 0, TimeSpan.Zero),
                        MembershipEnvironmentResult.FromSnapshot("Target", target, 0, TimeSpan.Zero)); int calls = 0;
                    using (var form = new MembershipCoverageDetailsForm("Source", new MembershipCoverageDiagnosticsBuilder().Build(source), "Target",
                        new MembershipCoverageDiagnosticsBuilder().Build(target), presentation, captureType10276Evidence: () => calls++))
                    {
                        form.StartPosition = FormStartPosition.Manual; form.Location = new System.Drawing.Point(-4000, -4000); form.Show(); Application.DoEvents();
                        var button = form.Controls.OfType<FlowLayoutPanel>().SelectMany(p => p.Controls.OfType<Button>()).Single(b => b.Text == "Capture Type 10276 AI Skill Config Evidence...");
                        Assert.IsTrue(button.Enabled); Assert.AreEqual(0, calls); button.PerformClick(); Assert.AreEqual(1, calls);
                    }
                    Assert.AreEqual(0, pair.Source.Service.Calls + pair.Target.Service.Calls + pair.Source.Service.ExecuteCalls + pair.Target.Service.ExecuteCalls);
                    using (var form = new Type10276EvidenceResultsForm("hash-only evidence")) Assert.IsTrue(form.Controls.OfType<RichTextBox>().Single().ReadOnly);
                }
                catch (Exception ex) { error = ex; }
            }) { IsBackground = true };
            thread.SetApartmentState(ApartmentState.STA); thread.Start(); Assert.IsTrue(thread.Join(TimeSpan.FromSeconds(30)));
            if (error != null) System.Runtime.ExceptionServices.ExceptionDispatchInfo.Capture(error).Throw();
        }
        [TestMethod]
        public void ZeroMembersStopBeforeMetadataAndBackingReads()
        {
            var pair = new Pair(); var report = pair.Capture(); Assert.AreEqual(0, report.Pairs.Count);
            Assert.AreEqual(0, pair.Source.Service.Calls + pair.Target.Service.Calls + pair.Source.Service.ExecuteCalls + pair.Target.Service.ExecuteCalls);
        }
        [TestMethod]
        public void ZeroTargetIsDiagnosticOneSidedEvidenceOnly()
        {
            var pair = new Pair(); pair.Source.Add(); var report = pair.Capture(); Assert.AreEqual("OneSidedEvidence", report.Pairs.Single().Outcome);
            Assert.AreEqual(0, pair.Target.Service.Calls + pair.Target.Service.ExecuteCalls);
        }
        [TestMethod]
        public void RegisteredMappingDiscoversActualEntityAndPrimaryWithoutAssumingAiskillconfigTable()
        {
            var pair = new Pair(); var left = pair.Source.Add(unique: " publisher_Skill "); var right = pair.Target.Add(unique: "PUBLISHER_skill"); right["ismanaged"] = true;
            var report = pair.Capture(); var match = report.Pairs.Single(); Assert.AreEqual("SemanticPair", match.Outcome);
            Assert.AreEqual(EntityName, report.Source.EntityName); Assert.AreEqual(Primary, report.Source.PrimaryId);
            Assert.AreEqual(left.Id, match.Source.PrimaryId); Assert.AreNotEqual(left.Id, right.Id);
            Assert.IsTrue(match.Categories.Contains("DifferentPrimaryId") && match.Categories.Contains("UnmanagedToManaged"));
            Assert.AreEqual(1, pair.Source.Service.ExecuteCalls); Assert.AreEqual(2, pair.Source.Service.Calls);
            CollectionAssert.AreEquivalent(new[] { Primary, "uniquename", "entitylogicalname" }, pair.Source.Queries.First().ColumnSet.Columns.ToArray());
        }
        [DataTestMethod, DataRow(false), DataRow(true)]
        public void MissingOrConflictingRegisteredMappingNeverGuessesBackingTable(bool conflict)
        {
            var pair = new Pair(); var row = pair.Source.Add(); pair.Source.Raw.Clear();
            pair.Source.Raw.Add(new ComponentIdentity(new SolutionComponentRecord(Guid.NewGuid(), 10276, row.Id), IdentityResolutionStatus.Unsupported));
            if (conflict) pair.Source.Reference(Guid.NewGuid());
            var report = pair.Capture(); Assert.IsNull(report.Source.EntityName); Assert.AreEqual("Incomplete", report.Source.Rows[row.Id].Status);
            Assert.AreEqual(conflict ? 0 : 1, pair.Source.Service.ExecuteCalls + pair.Source.Service.Calls);
            Assert.IsTrue(pair.Source.Queries.All(q => q.EntityName == "solutioncomponentdefinition"));
        }
        [TestMethod]
        public void VerifiedTableScopeUsesSnapshotPortableKeyAcrossDifferentParentGuids()
        {
            var pair = new Pair(); foreach (var side in new[] { pair.Source, pair.Target }) side.Scope(side.Add());
            var match = pair.Capture().Pairs.Single(); Assert.AreEqual("SemanticPair", match.Outcome);
            Assert.AreEqual("VerifiedSnapshotScope", match.Source.ParentStatus); Assert.IsTrue(match.Source.ParentComplete);
            Assert.IsTrue(pair.Source.Queries.Concat(pair.Target.Queries).All(q => q.EntityName == EntityName));
        }
        [DataTestMethod, DataRow(false), DataRow(true)]
        public void MissingOrAmbiguousParentCannotBeRepairedByCandidateB(bool ambiguous)
        {
            var pair = new Pair(); var row = pair.Source.Add(); var parent = pair.Source.Scope(row);
            pair.Target.Scope(pair.Target.Add());
            if (ambiguous) pair.Source.Context.Add(new ComponentIdentity(new SolutionComponentRecord(Guid.NewGuid(), 1, Guid.NewGuid()), IdentityResolutionStatus.Resolved, "ava_case"));
            else pair.Source.Context.Clear();
            var report = pair.Capture(); var found = report.Source.Rows[row.Id]; Assert.IsNull(found.CandidateA);
            Assert.IsNotNull(found.CandidateB);
            Assert.IsFalse(found.ParentComplete); Assert.AreEqual(ambiguous ? "Ambiguous" : "Incomplete", found.ParentStatus);
            Assert.IsFalse(report.Pairs.Any(p => p.Outcome == "SemanticPair"));
        }
        [TestMethod]
        public void BlankSuccessfullyReadParentIsNotPresentAndDoesNotBlock()
        {
            var pair = new Pair(); var row = pair.Source.Add(); pair.Source.Scope(row); row.Attributes.Remove("parenttableid");
            var found = pair.Capture().Source.Rows[row.Id]; Assert.IsNotNull(found.CandidateA); Assert.IsTrue(found.ParentComplete);
            Assert.AreEqual("NotPresent", found.References["parenttableid"].Status);
        }
        [TestMethod]
        public void CandidateACollisionCannotBeRepairedByDistinctDisplayNamesHashesOrLocalIds()
        {
            var pair = new Pair(); pair.Source.Add(); var extra = pair.Source.Add(unique: "PUBLISHER_skill"); extra["name"] = "Other display"; extra["expression"] = "Other expression"; pair.Target.Add();
            var report = pair.Capture(); Assert.IsTrue(report.Pairs.All(p => p.Outcome == "Ambiguous"));
            Assert.IsTrue(report.Source.Rows.Values.All(r => r.DuplicateA)); Assert.AreEqual(2, report.Source.Rows.Values.Select(r => r.CandidateB).Distinct().Count());
        }
        [TestMethod]
        public void DisplayNameOnlyOrBlankStrongestIdentifierCannotCreateA()
        {
            var pair = new Pair(); var row = pair.Source.Add(unique: " "); pair.Source.Attributes.Add(Attribute("schemaname", new StringAttributeMetadata())); row["schemaname"] = "weaker_schema";
            var found = pair.Capture().Source.Rows[row.Id]; Assert.IsNull(found.CandidateA); Assert.IsNotNull(found.CandidateB);
        }
        [TestMethod]
        public void DifferentAValuesNeverPairThroughEqualHashesDisplayNamesOrGuids()
        {
            var pair = new Pair(); var row = pair.Source.Add(unique: "one"); pair.Target.Add(row.Id, "two"); var report = pair.Capture();
            Assert.IsTrue(report.Pairs.All(p => p.Outcome == "OneSidedEvidence"));
        }
        [TestMethod]
        public void MissingBackingRecordIsNotMembershipAbsence()
        {
            var pair = new Pair(); var row = pair.Source.Add(); pair.Source.Data.Clear(); var report = pair.Capture();
            Assert.AreEqual("Missing", report.Source.Rows[row.Id].Status); Assert.IsTrue(report.Pairs.All(p => p.Outcome == "Incomplete"));
        }
        [DataTestMethod, DataRow(false), DataRow(true)]
        public void DuplicateOrConflictingBackingRowsAreAmbiguous(bool conflict)
        {
            var pair = new Pair(); var row = pair.Source.Add(); var normal = pair.Source.Service.RetrievePage;
            pair.Source.Service.RetrievePage = q => { var result = normal(q); var copy = Clone(result.Entities.Single(), q);
                if (conflict) copy["uniquename"] = "conflict"; result.Entities.Add(copy); return result; };
            var found = pair.Capture().Source.Rows[row.Id]; Assert.AreEqual("Duplicate", found.Status); Assert.IsNull(found.CandidateA);
        }
        [TestMethod]
        public void RepeatedRawMembershipDoesNotCreateDuplicateBackingCandidate()
        {
            var pair = new Pair(); var row = pair.Source.Add(); pair.Source.Reference(row.Id); pair.Target.Add(); var report = pair.Capture();
            Assert.AreEqual(2, report.Source.Raw.Count); Assert.AreEqual(1, report.Source.Rows.Count); Assert.AreEqual("SemanticPair", report.Pairs.Single().Outcome);
        }
        [DataTestMethod, DataRow("expression", true), DataRow("uniquename", false)]
        public void RuntimeReadableColumnFaultUsesBoundedIsolationAndProtectsCriticalEvidence(string field, bool candidateComplete)
        {
            var pair = new Pair(); var row = pair.Source.Add(); Fail(pair.Source, field); var report = pair.Capture(); var found = report.Source.Rows[row.Id];
            Assert.AreEqual("Unique", found.Status); Assert.AreEqual(candidateComplete, found.CandidateA != null);
            StringAssert.Contains(report.Build(), "Attribute faulted: " + field); Assert.IsFalse(report.Build().Contains(Secret));
            Assert.IsTrue(pair.Source.Service.Calls <= Type10276EvidenceCollector.MaxIsolationGroups + 2);
        }
        [TestMethod]
        public void PrimaryQueryFaultIsFaultedAndSafeWithoutWideningScope()
        {
            var pair = new Pair(); var row = pair.Source.Add(); Fail(pair.Source, Primary); var report = pair.Capture();
            Assert.AreEqual("Faulted", report.Source.Rows[row.Id].Status); Assert.IsNull(report.Source.Rows[row.Id].CandidateA);
            StringAssert.Contains(report.Build(), "0x8004023B"); Assert.IsFalse(report.Build().Contains(Secret));
            Assert.IsTrue(pair.Source.Queries.All(q => q.Criteria.Conditions.Single().Values.Cast<Guid>().SequenceEqual(new[] { row.Id })));
        }
        [TestMethod]
        public void UnreadableStrongIdentifierLeavesAIncomplete()
        {
            var pair = new Pair(); var row = pair.Source.Add(); Set(pair.Source.Attributes.Single(a => a.LogicalName == "uniquename"), "IsValidForRead", false);
            Assert.IsNull(pair.Capture().Source.Rows[row.Id].CandidateA);
            Assert.IsTrue(pair.Source.Queries.All(q => !q.ColumnSet.Columns.Contains("uniquename")));
        }
        [TestMethod]
        public void MetadataFaultStopsBeforeBackingAndDoesNotExposeServiceDetails()
        {
            var pair = new Pair(); var row = pair.Source.Add(); pair.Source.Service.ExecuteRequest = r => { throw new InvalidOperationException(Secret); };
            var report = pair.Capture(); Assert.AreEqual("Faulted", report.Source.Rows[row.Id].Status); Assert.AreEqual(0, pair.Source.Service.Calls);
            Assert.IsFalse(report.Build().Contains(Secret));
        }
        [TestMethod]
        public void TerminalCookieDoesNotInvalidateCompletedOnePage()
        {
            var pair = new Pair(); var row = pair.Source.Add(); var normal = pair.Source.Service.RetrievePage;
            pair.Source.Service.RetrievePage = q => { var result = normal(q); result.PagingCookie = Secret; return result; };
            var report = pair.Capture(); Assert.AreEqual("Unique", report.Source.Rows[row.Id].Status); Assert.IsNotNull(report.Source.Rows[row.Id].CandidateA);
            StringAssert.Contains(report.Build(), "PagingCookieSupplied=True"); Assert.IsFalse(report.Build().Contains(Secret));
        }
        [TestMethod]
        public void GenuinePagingDeduplicatesIdenticalCrossPageRowsAndWaitsForTerminal()
        {
            var pair = new Pair(); var first = pair.Source.Add(unique: "one"); var second = pair.Source.Add(unique: "two");
            pair.Source.Service.RetrievePage = q => q.PageInfo.PageNumber == 1
                ? new EntityCollection(new List<Entity> { Clone(first, q) }) { MoreRecords = true, PagingCookie = "page1" }
                : new EntityCollection(new List<Entity> { Clone(first, q), Clone(second, q) }) { PagingCookie = "terminal" };
            var report = pair.Capture(); Assert.IsTrue(report.Source.Rows.Values.All(r => r.Status == "Unique"));
            Assert.AreEqual(4, pair.Source.Service.Calls); StringAssert.Contains(report.Build(), "pageCount=2");
        }
        [TestMethod]
        public void StalledPagingDoesNotFinalizeCandidates()
        {
            var pair = new Pair(); var row = pair.Source.Add(); pair.Source.Service.RetrievePage = q => new EntityCollection(new List<Entity> { Clone(row, q) }) { MoreRecords = true, PagingCookie = "same" };
            var found = pair.Capture().Source.Rows[row.Id]; Assert.AreEqual("Incomplete", found.Status); Assert.IsNull(found.CandidateA); StringAssert.Contains(found.Reason, "Stalled paging");
        }
        [DataTestMethod, DataRow(200, 2), DataRow(201, 4)]
        public void BatchingIsByDistinctSelectedIds(int count, int calls)
        {
            var pair = new Pair(); for (int i = 0; i < count; i++) pair.Source.Add(unique: "publisher_skill_" + i);
            var report = pair.Capture(); Assert.AreEqual(count, report.Source.Rows.Count); Assert.AreEqual(calls, pair.Source.Service.Calls);
            Assert.IsTrue(pair.Source.Queries.All(q => q.Criteria.Conditions.Single().Values.Count <= 200));
        }
        [TestMethod]
        public void BinaryShadowAndLargeRawContentNeverLeakIntoEvidence()
        {
            var pair = new Pair(); var row = pair.Source.Add(); pair.Source.Attributes.Add(Attribute("binarypayload", new StringAttributeMetadata()));
            pair.Source.Attributes.Add(Attribute("organizationid", new LookupAttributeMetadata { Targets = new[] { "organization" } }));
            pair.Source.Attributes.Add(Attribute("organizationidname", new StringAttributeMetadata()));
            row["binarypayload"] = Secret; row["organizationidname"] = Secret; row["organizationid"] = new EntityReference("organization", Guid.NewGuid());
            var report = pair.Capture(); Assert.IsNotNull(report.Source.Rows[row.Id].CandidateA); Assert.IsFalse(report.Build().Contains(Secret));
            Assert.IsTrue(pair.Source.Queries.All(q => !q.ColumnSet.Columns.Contains("binarypayload") && !q.ColumnSet.Columns.Contains("organizationidname")));
            Assert.IsTrue(report.Source.Rows[row.Id].Content["expression"].Known);
        }
        [TestMethod]
        public void CancellationStopsBeforeTargetAndDoesNotChangeProductionEvidence()
        {
            var pair = new Pair(); pair.Source.Add(); pair.Target.Add(); using (var cancellation = new CancellationTokenSource())
            {
                var normal = pair.Source.Service.RetrievePage; pair.Source.Service.RetrievePage = q => { cancellation.Cancel(); return normal(q); };
                Assert.ThrowsException<OperationCanceledException>(() => pair.Capture(cancellation.Token)); Assert.AreEqual(0, pair.Target.Service.Calls + pair.Target.Service.ExecuteCalls);
            }
            Assert.IsTrue(pair.Source.Raw.Concat(pair.Target.Raw).All(c => c.Status == IdentityResolutionStatus.Unsupported && c.ComparisonKey == null));
            Assert.AreEqual(0, pair.Source.Service.WriteCalls + pair.Target.Service.WriteCalls);
        }
        [TestMethod]
        public void ReportHasAllSectionsAndCaptureCannotChangeMembershipResults()
        {
            var pair = new Pair(); pair.Source.Add(); pair.Target.Add(); var before = new SolutionMembershipComparer().Compare(pair.Source.Snapshot(), pair.Target.Snapshot()).Select(r => r.Presence).ToArray();
            var report = pair.Capture(); var text = report.Build(); var after = new SolutionMembershipComparer().Compare(pair.Source.Snapshot(), pair.Target.Snapshot()).Select(r => r.Presence).ToArray();
            CollectionAssert.AreEqual(before, after); Assert.AreEqual(text, report.Build());
            foreach (var heading in new[] { "RAW TYPE 10276 MEMBERSHIP", "REGISTERED-DEFINITION / BACKING-ENTITY DISCOVERY", "PORTABLE PARENT / REFERENCE RESOLUTION", "BACKING-RECORD CORRELATION", "READABLE / UNAVAILABLE SCHEMA", "RUNTIME-READABLE / FAULTED COLUMNS",
                "PARENT / REFERENCE RELATIONSHIP DISCOVERY", "CANDIDATE A / B ANALYSIS", "DUPLICATE / COLLISION ANALYSIS", "SELECTED SOURCE / TARGET RECORD FIELD COMPARISON",
                "ONE-SIDED LIFECYCLE EVIDENCE", "CROSS-ENVIRONMENT PAIRING STATUS", "PORTABILITY ASSESSMENT", "EXACT REQUEST LEDGER" }) StringAssert.Contains(text, heading);
            StringAssert.Contains(text, "TotalReads=3"); StringAssert.Contains(text, "NormalMembershipEvidenceRequests=0");
            StringAssert.Contains(text, "AdditionalWhoAmI=0"); StringAssert.Contains(text, "Writes=0");
        }
        [TestMethod]
        public void UnreadableExposedScopeIsNotQueriedAndCannotBeBypassedByParentLookup()
        {
            var pair = new Pair(); var row = pair.Source.Add(); pair.Source.Scope(row);
            Set(pair.Source.Attributes.Single(a => a.LogicalName == "entitylogicalname"), "IsValidForRead", false);
            var report = pair.Capture(); Assert.AreEqual("Unique", report.Source.Rows[row.Id].Status); Assert.IsNull(report.Source.Rows[row.Id].CandidateA);
            Assert.IsTrue(pair.Source.Queries.All(q => !q.ColumnSet.Columns.Contains("entitylogicalname")));
        }
        [TestMethod]
        public void TextScopeReusesCompleteSnapshotTableIdentityWithoutAdditionalMetadata()
        {
            var pair = new Pair(); pair.Source.Add(); pair.Target.Add(); var report = pair.Capture();
            Assert.AreEqual("SemanticPair", report.Pairs.Single().Outcome);
            Assert.AreEqual("ava_case", report.Pairs.Single().Source.TableKey);
            Assert.AreEqual(1, pair.Source.Service.ExecuteCalls);
            StringAssert.Contains(report.Build(), "Reused verified snapshot Table identity");
        }
        [TestMethod]
        public void SolutionOrganizationOwnerAndInstallationReferencesAreAuditOnly()
        {
            var pair = new Pair(); var row = pair.Source.Add();
            foreach (var field in new[] { "solutionid", "organizationid", "ownerid", "relatedidunique" }) {
                pair.Source.Attributes.Add(Attribute(field, new LookupAttributeMetadata { Targets = new[] { "unverified_local_entity" } }));
                row[field] = new EntityReference("unverified_local_entity", Guid.NewGuid());
            }
            var found = pair.Capture().Source.Rows[row.Id]; Assert.IsNotNull(found.CandidateA); Assert.IsTrue(found.ParentComplete);
            Assert.IsTrue(found.Fields.Keys.Contains("solutionid"));
        }
        [TestMethod]
        public void ScopedRegisteredDefinitionDiscoveryDoesNotRequireAssumedBackingEntity()
        {
            var pair = new Pair(); var row = pair.Source.Add(); pair.Source.Raw.Clear();
            pair.Source.Raw.Add(new ComponentIdentity(new SolutionComponentRecord(Guid.NewGuid(), 10276, row.Id), IdentityResolutionStatus.Unsupported));
            var normal = pair.Source.Service.RetrievePage;
            pair.Source.Service.RetrievePage = q => q.EntityName == "solutioncomponentdefinition" ? new EntityCollection(new List<Entity> {
                new Entity("solutioncomponentdefinition", Guid.NewGuid()) { ["objecttypecode"] = 10276, ["name"] = "AiSkillConfig", ["primaryentityname"] = EntityName }
            }) { PagingCookie = "terminal" } : normal(q);
            var report = pair.Capture(); Assert.AreEqual(EntityName, report.Source.EntityName); Assert.AreEqual("Unique", report.Source.Rows[row.Id].Status);
            Assert.IsNotNull(report.Source.Rows[row.Id].CandidateA); Assert.AreEqual(3, pair.Source.Service.Calls);
            StringAssert.Contains(report.Build(), "Discovered registered definition");
        }
        [TestMethod]
        public void ConflictingRegisteredDefinitionRecordsStopBeforeBackingSchemaOrQuery()
        {
            var pair = new Pair(); var row = pair.Source.Add(); pair.Source.Raw.Clear();
            pair.Source.Raw.Add(new ComponentIdentity(new SolutionComponentRecord(Guid.NewGuid(), 10276, row.Id), IdentityResolutionStatus.Unsupported));
            pair.Source.Service.RetrievePage = q => new EntityCollection(new List<Entity> {
                new Entity("solutioncomponentdefinition", Guid.NewGuid()) { ["objecttypecode"] = 10276 },
                new Entity("solutioncomponentdefinition", Guid.NewGuid()) { ["objecttypecode"] = 10276 }
            });
            var report = pair.Capture(); Assert.IsNull(report.Source.EntityName); Assert.AreEqual(0, pair.Source.Service.ExecuteCalls);
            Assert.AreEqual(1, pair.Source.Service.Calls); Assert.IsNull(report.Source.Rows[row.Id].CandidateA);
        }
        [TestMethod]
        public void TwoSelectedSourceMembersAndZeroTargetNeverInventPairOrAbsence()
        {
            var pair = new Pair(); pair.Source.Add(unique: "skill_one"); pair.Source.Add(unique: "skill_two");
            var before = new SolutionMembershipComparer().Compare(pair.Source.Snapshot(), pair.Target.Snapshot()).Select(r => r.Presence).ToArray();
            var report = pair.Capture(); Assert.AreEqual(2, report.Source.Raw.Count); Assert.AreEqual(2, report.Source.Rows.Count);
            Assert.IsTrue(report.Source.Rows.Values.All(r => r.Status == "Unique" && r.CandidateA != null));
            Assert.IsTrue(report.Pairs.All(p => p.Outcome == "OneSidedEvidence"));
            Assert.AreEqual(0, pair.Target.Service.Calls + pair.Target.Service.ExecuteCalls);
            StringAssert.Contains(report.Build(), "No selected Type 10276 membership references on this side; no backing lookup performed; this does not prove backing-record absence.");
            StringAssert.Contains(report.Build(), "Actual cross-environment semantic pairs=0");
            StringAssert.Contains(report.Build(), "One-sided selected membership evidence is not backing-record absence proof");
            CollectionAssert.AreEqual(before, new SolutionMembershipComparer().Compare(pair.Source.Snapshot(), pair.Target.Snapshot()).Select(r => r.Presence).ToArray());
            Assert.AreEqual(3, report.Source.Requests.Count); Assert.AreEqual(0, report.Target.Requests.Count);
            Assert.AreEqual(0, pair.Source.Service.WriteCalls + pair.Target.Service.WriteCalls);
        }
        [TestMethod]
        public void DisplayPrimaryNameIsNeverAnInternalSemanticIdentifier()
        {
            var pair = new Pair(); var row = pair.Source.Add(); pair.Source.Attributes.RemoveAll(a => a.LogicalName == "uniquename");
            var found = pair.Capture().Source.Rows[row.Id]; Assert.IsNull(found.CandidateA); Assert.IsNotNull(found.CandidateB);
            StringAssert.Contains(found.BlockingReason, "No independently semantic internal");
        }
        [TestMethod]
        public void GuidValuedInternalIdentifierDoesNotEstablishPortability()
        {
            var pair = new Pair(); var row = pair.Source.Add(unique: Guid.NewGuid().ToString("D"));
            var found = pair.Capture().Source.Rows[row.Id]; Assert.IsNull(found.CandidateA);
            StringAssert.Contains(found.BlockingReason, "GUID; portability unproven");
        }
        [TestMethod]
        public void InternalIdentifierWithoutExposedTableStillRemainsEvidenceHypothesis()
        {
            var pair = new Pair(); var row = pair.Source.Add(); pair.Source.Attributes.RemoveAll(a => a.LogicalName == "entitylogicalname");
            var found = pair.Capture().Source.Rows[row.Id]; Assert.IsNotNull(found.CandidateA); Assert.AreEqual("NotExposed", found.TableStatus);
            Assert.IsTrue(found.CandidateA.StartsWith("aiskillconfig-candidate-a:"));
            Assert.AreEqual(IdentityResolutionStatus.Unsupported, pair.Source.Raw.Single().Status); Assert.IsNull(pair.Source.Raw.Single().ComparisonKey);
        }
        [TestMethod]
        public void OptionalIncompleteResponseCannotInvalidateCriticalCorrelation()
        {
            var pair = new Pair(); var row = pair.Source.Add(); var normal = pair.Source.Service.RetrievePage;
            pair.Source.Service.RetrievePage = q => q.ColumnSet.Columns.Contains("expression") ? null : normal(q);
            var found = pair.Capture().Source.Rows[row.Id]; Assert.AreEqual("Unique", found.Status); Assert.IsNotNull(found.CandidateA);
            Assert.AreEqual("Unavailable", found.Content["expression"].Presence);
        }
        [TestMethod]
        public void PromptAndConfigurationBodiesAreHashOnlyEvenWhenShort()
        {
            var pair = new Pair(); var row = pair.Source.Add();
            foreach (var field in new[] { "prompt", "configuration", "jsonconfig", "xmlconfig" }) {
                pair.Source.Attributes.Add(Attribute(field, new StringAttributeMetadata())); row[field] = Secret;
            }
            var report = pair.Capture(); Assert.IsFalse(report.Build().Contains(Secret));
            foreach (var field in new[] { "prompt", "configuration", "jsonconfig", "xmlconfig" }) Assert.IsTrue(report.Source.Rows[row.Id].Content[field].Known);
        }
        [TestMethod]
        public void CancellationDuringBoundedIsolationStopsImmediately()
        {
            var pair = new Pair(); pair.Source.Add(); Fail(pair.Source, "expression"); var normal = pair.Source.Service.RetrievePage;
            using (var cancellation = new CancellationTokenSource()) {
                pair.Source.Service.RetrievePage = q => { if (pair.Source.Service.Calls == 3) cancellation.Cancel(); return normal(q); };
                Assert.ThrowsException<OperationCanceledException>(() => pair.Capture(cancellation.Token));
                Assert.AreEqual(0, pair.Target.Service.ExecuteCalls + pair.Target.Service.Calls);
            }
        }
        [TestMethod]
        public void InvalidExposedScopeTypeCannotBeTreatedAsUnexposed()
        {
            var pair = new Pair(); var row = pair.Source.Add(); pair.Source.Attributes.RemoveAll(a => a.LogicalName == "entitylogicalname");
            pair.Source.Attributes.Add(Attribute("entitylogicalname", new PicklistAttributeMetadata())); row["entitylogicalname"] = new OptionSetValue(12);
            var found = pair.Capture().Source.Rows[row.Id]; Assert.AreEqual("Incomplete", found.TableStatus); Assert.IsNull(found.CandidateA);
        }
        [TestMethod]
        public void RegisteredDiscoveryFaultIsExplicitWithoutBackingOrAbsenceFallback()
        {
            var pair = new Pair(); var row = pair.Source.Add(); pair.Source.Raw.Clear();
            pair.Source.Raw.Add(new ComponentIdentity(new SolutionComponentRecord(Guid.NewGuid(), 10276, row.Id), IdentityResolutionStatus.Unsupported));
            pair.Source.Service.RetrievePage = q => { throw new InvalidOperationException(Secret); };
            var report = pair.Capture(); Assert.AreEqual("Faulted", report.Source.Rows[row.Id].Status); Assert.IsNull(report.Source.EntityName);
            Assert.AreEqual(0, pair.Source.Service.ExecuteCalls); Assert.AreEqual(1, pair.Source.Service.Calls); Assert.IsFalse(report.Build().Contains(Secret));
            Assert.AreEqual("Incomplete", report.Pairs.Single().Outcome);
        }
        [TestMethod]
        public void ConflictingPrimaryKeyDoesNotFinalizeBackingCorrelation()
        {
            var pair = new Pair(); var row = pair.Source.Add(); var normal = pair.Source.Service.RetrievePage;
            pair.Source.Service.RetrievePage = q => { var response = normal(q); response.Entities.Single()[Primary] = Guid.NewGuid(); return response; };
            var found = pair.Capture().Source.Rows[row.Id]; Assert.AreEqual("Incomplete", found.Status); Assert.IsNull(found.CandidateA);
        }
        [DataTestMethod, DataRow("appmodule", 80), DataRow("workflow", 29)]
        public void ParentReferencesReuseVerifiedAppAndProcessSnapshotIdentity(string entity, int type)
        {
            var pair = new Pair();
            foreach (var side in new[] { pair.Source, pair.Target }) {
                var row = side.Add(); var id = Guid.NewGuid();
                side.Attributes.Add(Attribute("parentcomponentid", new LookupAttributeMetadata { Targets = new[] { entity } }));
                row["parentcomponentid"] = new EntityReference(entity, id);
                side.Context.Add(new ComponentIdentity(new SolutionComponentRecord(Guid.NewGuid(), type, id), IdentityResolutionStatus.Resolved, "publisher_parent"));
            }
            var report = pair.Capture(); Assert.AreEqual("SemanticPair", report.Pairs.Single().Outcome);
            Assert.IsTrue(report.Pairs.Single().Source.ParentComplete);
            Assert.IsTrue(pair.Source.Queries.Concat(pair.Target.Queries).All(q => q.EntityName == EntityName));
            StringAssert.Contains(report.Build(), "reused verified snapshot identity");
        }
        [TestMethod]
        public void ReferencedUnsupportedComponentDoesNotUseLocalGuidOrDisplayAsScope()
        {
            var pair = new Pair(); var row = pair.Source.Add();
            pair.Source.Attributes.Add(Attribute("modelid", new LookupAttributeMetadata { Targets = new[] { "unverified_model" } }));
            row["modelid"] = new EntityReference("unverified_model", Guid.NewGuid());
            var found = pair.Capture().Source.Rows[row.Id]; Assert.IsFalse(found.ParentComplete); Assert.IsNull(found.CandidateA);
            Assert.IsNotNull(found.CandidateB); StringAssert.Contains(found.BlockingReason, "Parent/reference identity");
        }
        [TestMethod]
        public void BlankObjectIdsRemainIncompleteInsteadOfZeroMembershipOrAbsence()
        {
            var pair = new Pair(); pair.Source.Raw.Add(new ComponentIdentity(new SolutionComponentRecord(Guid.NewGuid(), 10276, null), IdentityResolutionStatus.Unsupported));
            var report = pair.Capture(); Assert.AreEqual(1, report.Source.Raw.Count); Assert.AreEqual(0, report.Source.Rows.Count);
            Assert.AreEqual("Incomplete", report.Pairs.Single().Outcome); Assert.AreEqual(0, pair.Source.Service.Calls + pair.Source.Service.ExecuteCalls);
            StringAssert.Contains(report.Build(), "Selected membership references have no usable nonblank object IDs");
        }
        [TestMethod]
        public void SemanticUniqueNameIsNotMisreportedAsInstallationUniqueId()
        {
            var pair = new Pair(); pair.Source.Add(); pair.Target.Add(); var report = pair.Capture();
            Assert.AreEqual("SemanticPair", report.Pairs.Single().Outcome);
            Assert.IsFalse(report.Pairs.Single().Categories.Contains("SameUniqueId"));
        }
        [DataTestMethod, DataRow("aimodel"), DataRow("generationapi"), DataRow("sdkmessageid")]
        public void BlankOptionalLookupIsNotPresentAndDoesNotBlockCandidateA(string field)
        {
            var pair = new Pair(); var row = LiveShape(pair.Source); SnapshotColumn(pair.Source, row, "description");
            pair.Source.Attributes.Add(Attribute(field, new LookupAttributeMetadata { Targets = new[] { "unverified_optional_entity" } }));
            var found = pair.Capture().Source.Rows[row.Id]; Assert.AreEqual("NotPresent", found.References[field].Status);
            Assert.IsTrue(found.ParentComplete); Assert.IsNotNull(found.CandidateA); Assert.AreEqual("ava_casecontact", found.TableKey);
        }
        [TestMethod]
        public void LookupShadowsNeverOverrideVerifiedEntityOrBecomeIndependentReferences()
        {
            var pair = new Pair(); var row = LiveShape(pair.Source); SnapshotColumn(pair.Source, row, "casecontactname");
            foreach (var field in new[] { "entityname", "attributename", "aimodelname", "aimodelyominame" })
                pair.Source.Attributes.Add(Attribute(field, new StringAttributeMetadata()));
            pair.Source.Attributes.Add(Attribute("aimodel", new LookupAttributeMetadata { Targets = new[] { "aimodel" } }));
            var report = pair.Capture(); var found = report.Source.Rows[row.Id];
            Assert.AreEqual("Verified", found.TableStatus); Assert.AreEqual("ava_casecontact", found.TableKey); Assert.IsNotNull(found.CandidateA);
            Assert.IsFalse(report.Source.ScopeFields.Contains("entityname"));
            Assert.IsTrue(found.References.Keys.All(k => !k.EndsWith("name", StringComparison.Ordinal)));
            Assert.IsTrue(pair.Source.Queries.All(q => q.ColumnSet.Columns.All(c => !new[] { "entityname", "attributename", "aimodelname", "aimodelyominame" }.Contains(c))));
            StringAssert.Contains(report.Build(), "aimodel: NotPresent");
        }
        [TestMethod]
        public void SnapshotAttributeIdentityCompletesItsPortionWithoutExtraMetadataRead()
        {
            var pair = new Pair(); var row = LiveShape(pair.Source); SnapshotColumn(pair.Source, row, "description");
            var report = pair.Capture(); var found = report.Source.Rows[row.Id];
            Assert.AreEqual("VerifiedSnapshot", found.References["attribute"].Status);
            Assert.AreEqual("ava_casecontact.description", found.References["attribute"].PortableKey);
            Assert.IsTrue(found.ParentComplete); Assert.IsNotNull(found.CandidateA); Assert.AreEqual(0, report.Source.AttributeMetadataSubrequests);
            Assert.AreEqual(1, pair.Source.Service.ExecuteCalls); Assert.AreEqual(0, pair.Target.Service.ExecuteCalls + pair.Target.Service.Calls);
        }
        [TestMethod]
        public void MissingSnapshotAttributeUsesOneSelectedSourceMetadataBatchAndNoTargetReads()
        {
            var pair = new Pair(); var row = LiveShape(pair.Source); MetadataAttributes(pair.Source);
            var report = pair.Capture(); var found = report.Source.Rows[row.Id];
            Assert.AreEqual("VerifiedScopedMetadata", found.References["attribute"].Status);
            Assert.AreEqual("ava_casecontact.column_" + ((EntityReference)row["attribute"]).Id.ToString("N"), found.References["attribute"].PortableKey); Assert.IsNotNull(found.CandidateA);
            Assert.AreEqual(1, report.Source.AttributeMetadataBatches); Assert.AreEqual(1, report.Source.AttributeMetadataSubrequests);
            Assert.AreEqual(4, report.Source.Requests.Count); Assert.AreEqual(0, report.Target.Requests.Count);
            StringAssert.Contains(report.Build(), "Actual cross-environment semantic pairs=0");
            Assert.AreEqual(0, pair.Source.Service.WriteCalls + pair.Target.Service.WriteCalls);
        }
        [TestMethod]
        public void UnresolvedAttributeMetadataRemainsIncompleteWithoutDisplayOrHashRepair()
        {
            var pair = new Pair(); var row = LiveShape(pair.Source); MetadataAttributes(pair.Source, missing: true);
            var report = pair.Capture(); var found = report.Source.Rows[row.Id];
            Assert.IsFalse(found.ParentComplete); Assert.IsNull(found.CandidateA); Assert.IsNotNull(found.CandidateB);
            Assert.AreEqual("Incomplete", found.References["attribute"].Status); Assert.AreEqual("Incomplete", report.Pairs.Single().Outcome);
        }
        [TestMethod]
        public void AmbiguousSnapshotAttributeCannotBeRepairedByMetadataFallback()
        {
            var pair = new Pair(); var row = LiveShape(pair.Source); var id = SnapshotColumn(pair.Source, row, "name");
            pair.Source.Context.Add(new ComponentIdentity(new SolutionComponentRecord(Guid.NewGuid(), 2, id), IdentityResolutionStatus.Resolved, "ava_casecontact.other"));
            var report = pair.Capture(); var found = report.Source.Rows[row.Id];
            Assert.AreEqual("Ambiguous", found.References["attribute"].Status); Assert.IsFalse(found.ParentComplete); Assert.IsNull(found.CandidateA);
            Assert.AreEqual(0, report.Source.AttributeMetadataBatches); Assert.AreEqual("Ambiguous", report.Pairs.Single().Outcome);
        }
        [TestMethod]
        public void DistinctAttributesDistinguishCandidatesWithIdenticalUniqueNames()
        {
            var pair = new Pair(); var first = LiveShape(pair.Source); SnapshotColumn(pair.Source, first, "firstcolumn");
            var second = LiveShape(pair.Source); SnapshotColumn(pair.Source, second, "secondcolumn");
            var report = pair.Capture(); Assert.IsTrue(report.Source.Rows.Values.All(r => r.CandidateA != null && !r.DuplicateA));
            Assert.AreNotEqual(report.Source.Rows[first.Id].CandidateA, report.Source.Rows[second.Id].CandidateA);
            Assert.AreEqual(0, report.Target.Requests.Count);
        }
        [TestMethod]
        public void IncompleteSnapshotColumnCanBeIndependentlyVerifiedByScopedMetadata()
        {
            var pair = new Pair(); var row = LiveShape(pair.Source); var id = ((EntityReference)row["attribute"]).Id;
            pair.Source.Context.Add(new ComponentIdentity(new SolutionComponentRecord(Guid.NewGuid(), 2, id), IdentityResolutionStatus.Unresolved));
            MetadataAttributes(pair.Source); var found = pair.Capture().Source.Rows[row.Id];
            Assert.AreEqual("VerifiedScopedMetadata", found.References["attribute"].Status); Assert.IsNotNull(found.CandidateA);
        }
        [TestMethod]
        public void ForeignTableAttributeMetadataCannotCompleteCandidateA()
        {
            var pair = new Pair(); var row = LiveShape(pair.Source); MetadataAttributes(pair.Source, parent: "another_table");
            var found = pair.Capture().Source.Rows[row.Id]; Assert.IsNull(found.CandidateA); Assert.IsFalse(found.ParentComplete);
        }
        [TestMethod]
        public void SnapshotColumnFromDifferentTableCannotCompleteOrTriggerFallback()
        {
            var pair = new Pair(); var row = LiveShape(pair.Source); var id = ((EntityReference)row["attribute"]).Id;
            pair.Source.Context.Add(new ComponentIdentity(new SolutionComponentRecord(Guid.NewGuid(), 2, id), IdentityResolutionStatus.Resolved, "another_table.name"));
            var report = pair.Capture(); Assert.IsNull(report.Source.Rows[row.Id].CandidateA); Assert.AreEqual(0, report.Source.AttributeMetadataBatches);
        }
        [DataTestMethod, DataRow(false), DataRow(true)]
        public void DuplicateAttributeMetadataResponseOrCanonicalCollisionIsAmbiguous(bool canonicalCollision)
        {
            var pair = new Pair(); var first = LiveShape(pair.Source); if (canonicalCollision) LiveShape(pair.Source);
            MetadataAttributes(pair.Source, duplicate: !canonicalCollision, sameLogicalName: canonicalCollision);
            var report = pair.Capture(); Assert.IsTrue(report.Source.Rows.Values.All(r => r.References["attribute"].Status == "Ambiguous" && r.CandidateA == null));
        }
        [TestMethod]
        public void AttributeMetadataFaultIsSafeAndConservative()
        {
            var pair = new Pair(); var row = LiveShape(pair.Source); MetadataAttributes(pair.Source, fault: true);
            var report = pair.Capture(); Assert.AreEqual("Faulted", report.Source.Rows[row.Id].References["attribute"].Status);
            Assert.IsNull(report.Source.Rows[row.Id].CandidateA); Assert.IsFalse(report.Build().Contains(Secret)); StringAssert.Contains(report.Build(), "0x8004023B");
        }
        [TestMethod]
        public void SharedAttributeReferencesAreDeduplicatedAcrossSelectedRecords()
        {
            var pair = new Pair(); var first = LiveShape(pair.Source); var second = LiveShape(pair.Source);
            second["attribute"] = first["attribute"]; second["uniquename"] = "publisher_second"; MetadataAttributes(pair.Source);
            var report = pair.Capture(); Assert.AreEqual(1, report.Source.AttributeMetadataBatches); Assert.AreEqual(1, report.Source.AttributeMetadataSubrequests);
            Assert.IsTrue(report.Source.Rows.Values.All(r => r.CandidateA != null));
        }
        [TestMethod]
        public void TargetDoesNotPerformScopedAttributeFallbackEvenWithMembers()
        {
            var pair = new Pair(); var row = LiveShape(pair.Target); MetadataAttributes(pair.Target);
            var report = pair.Capture(); Assert.IsNull(report.Target.Rows[row.Id].CandidateA); Assert.AreEqual(0, report.Target.AttributeMetadataBatches);
            Assert.AreEqual(1, pair.Target.Service.ExecuteCalls); Assert.AreEqual(0, pair.Source.Service.ExecuteCalls + pair.Source.Service.Calls);
        }
        [DataTestMethod, DataRow(200, 1), DataRow(201, 2)]
        public void ScopedAttributeMetadataReadsBatchAtTwoHundredIds(int count, int batches)
        {
            var pair = new Pair(); for (int i = 0; i < count; i++) LiveShape(pair.Source)["uniquename"] = "skill_" + i;
            MetadataAttributes(pair.Source); var report = pair.Capture(); Assert.AreEqual(batches, report.Source.AttributeMetadataBatches);
            Assert.AreEqual(count, report.Source.AttributeMetadataSubrequests); Assert.IsTrue(report.Source.Rows.Values.All(r => r.CandidateA != null));
            Assert.AreEqual(0, report.Target.Requests.Count);
        }
        [TestMethod]
        public void CancellationDuringAttributeMetadataStopsWithoutTargetReads()
        {
            var pair = new Pair(); LiveShape(pair.Source); MetadataAttributes(pair.Source); var normal = pair.Source.Service.ExecuteRequest;
            using (var cancellation = new CancellationTokenSource()) {
                pair.Source.Service.ExecuteRequest = r => { if (r is ExecuteMultipleRequest) cancellation.Cancel(); return normal(r); };
                Assert.ThrowsException<OperationCanceledException>(() => pair.Capture(cancellation.Token));
                Assert.AreEqual(0, pair.Target.Service.ExecuteCalls + pair.Target.Service.Calls);
            }
        }
        [TestMethod]
        public void UnreadableOptionalLookupCannotBeAssumedNotPresent()
        {
            var pair = new Pair(); var row = LiveShape(pair.Source);
            var lookup = Attribute("aimodel", new LookupAttributeMetadata { Targets = new[] { "aimodel" } }); Set(lookup, "IsValidForRead", false); pair.Source.Attributes.Add(lookup);
            var found = pair.Capture().Source.Rows[row.Id]; Assert.AreEqual("Incomplete", found.References["aimodel"].Status); Assert.IsNull(found.CandidateA);
        }
        [TestMethod]
        public void ScopedAttributeIdentityConflictingWithSnapshotColumnIsAmbiguous()
        {
            var pair = new Pair(); var row = LiveShape(pair.Source);
            pair.Source.Context.Add(new ComponentIdentity(new SolutionComponentRecord(Guid.NewGuid(), 2, Guid.NewGuid()),
                IdentityResolutionStatus.Resolved, "ava_casecontact.same_column"));
            MetadataAttributes(pair.Source, sameLogicalName: true); var found = pair.Capture().Source.Rows[row.Id];
            Assert.AreEqual("Ambiguous", found.References["attribute"].Status); Assert.IsFalse(found.ParentComplete); Assert.IsNull(found.CandidateA);
        }
        [TestMethod]
        public void BlankOptionalRelationshipsDoNotChangeCandidateIdentity()
        {
            var pair = new Pair(); var row = LiveShape(pair.Source); SnapshotColumn(pair.Source, row, "name");
            var original = pair.Capture().Source.Rows[row.Id].CandidateA;
            pair.Source.Attributes.Add(Attribute("aimodel", new LookupAttributeMetadata { Targets = new[] { "aimodel" } }));
            var corrected = pair.Capture().Source.Rows[row.Id].CandidateA; Assert.IsNotNull(original); Assert.AreEqual(original, corrected);
        }
        private static Entity LiveShape(Fixture side)
        {
            var row = side.Add(); side.Attributes.RemoveAll(a => a.LogicalName == "entitylogicalname"); row.Attributes.Remove("entitylogicalname");
            var table = side.Context.Single(c => c.Record.ComponentType == 1); side.Context.Remove(table);
            side.Context.Add(new ComponentIdentity(table.Record, IdentityResolutionStatus.Resolved, "ava_casecontact"));
            foreach (var field in new[] { "entity", "attribute" }) if (!side.Attributes.Any(a => a.LogicalName == field))
                side.Attributes.Add(Attribute(field, new LookupAttributeMetadata { Targets = new[] { field } }));
            row["entity"] = new EntityReference("entity", table.Record.ObjectId.Value);
            row["attribute"] = new EntityReference("attribute", Guid.NewGuid()); return row;
        }
        private static Guid SnapshotColumn(Fixture side, Entity row, string logicalName)
        {
            var id = ((EntityReference)row["attribute"]).Id;
            side.Context.Add(new ComponentIdentity(new SolutionComponentRecord(Guid.NewGuid(), 2, id), IdentityResolutionStatus.Resolved, "ava_casecontact." + logicalName)); return id;
        }
        private static void MetadataAttributes(Fixture side, bool missing = false, bool duplicate = false, bool sameLogicalName = false,
            bool fault = false, string parent = "ava_casecontact")
        {
            var normal = side.Service.ExecuteRequest;
            side.Service.ExecuteRequest = r => {
                if (!(r is ExecuteMultipleRequest)) return normal(r);
                var request = (ExecuteMultipleRequest)r; Assert.IsTrue(request.Settings.ContinueOnError && request.Settings.ReturnResponses);
                Assert.IsTrue(request.Requests.Count <= 200); var response = new ExecuteMultipleResponse();
                var items = new ExecuteMultipleResponseItemCollection(); response.Results["Responses"] = items;
                for (int i = 0; i < request.Requests.Count; i++) {
                    Assert.IsInstanceOfType(request.Requests[i], typeof(RetrieveAttributeRequest)); var child = (RetrieveAttributeRequest)request.Requests[i];
                    Assert.AreEqual("ava_casecontact", child.EntityLogicalName); Assert.IsFalse(child.RetrieveAsIfPublished);
                    Assert.IsTrue(side.Data.Any(d => d.GetAttributeValue<EntityReference>("attribute")?.Id == child.MetadataId));
                    if (missing) continue;
                    var attribute = new StringAttributeMetadata { MetadataId = child.MetadataId, LogicalName = sameLogicalName ? "same_column" : "column_" + child.MetadataId.ToString("N") };
                    Set(attribute, "EntityLogicalName", parent); var result = new RetrieveAttributeResponse(); result.Results["AttributeMetadata"] = attribute;
                    items.Add(new ExecuteMultipleResponseItem { RequestIndex = i, Response = fault ? null : result,
                        Fault = fault ? new OrganizationServiceFault { ErrorCode = unchecked((int)0x8004023B), Message = Secret, TraceText = Secret } : null });
                    if (duplicate) items.Add(new ExecuteMultipleResponseItem { RequestIndex = i, Response = result });
                }
                return response;
            };
        }
        private static Entity Clone(Entity row, QueryExpression query)
        { var copy = new Entity(row.LogicalName, row.Id); foreach (var field in query.ColumnSet.Columns) if (row.Contains(field)) copy[field] = row[field]; return copy; }
        private static void Fail(Fixture fixture, string field)
        {
            var normal = fixture.Service.RetrievePage; fixture.Service.RetrievePage = q => { if (q.ColumnSet.Columns.Contains(field))
                throw new FaultException<OrganizationServiceFault>(new OrganizationServiceFault { ErrorCode = unchecked((int)0x8004023B), Message = Secret, TraceText = Secret }, new FaultReason(Secret)); return normal(q); };
        }
        private static AttributeMetadata Attribute(string name, AttributeMetadata attribute)
        { attribute.LogicalName = name; Set(attribute, "IsValidForRead", true); return attribute; }
        private static void Set(object obj, string name, object value) => obj.GetType().GetProperty(name).SetValue(obj, value, null);
        private sealed class Pair
        {
            internal readonly Fixture Source = new Fixture(), Target = new Fixture();
            internal AiSkillConfigEvidenceReport Capture(CancellationToken token = default(CancellationToken)) => new Type10276EvidenceCollector().Capture(Source.Service, Source.Snapshot(), "1", Target.Service, Target.Snapshot(), "2", token);
        }
        private sealed class Fixture
        {
            internal readonly List<Entity> Data = new List<Entity>();
            internal readonly List<ComponentIdentity> Raw = new List<ComponentIdentity>(), Context = new List<ComponentIdentity>();
            internal readonly List<AttributeMetadata> Attributes = new List<AttributeMetadata>();
            internal readonly List<QueryExpression> Queries = new List<QueryExpression>();
            internal string BackingEntity = EntityName;
            internal readonly FakeOrganizationService Service;
            internal Fixture()
            {
                Attributes.Add(Attribute(Primary, new UniqueIdentifierAttributeMetadata()));
                foreach (var name in new[] { "name", "uniquename", "installationidunique" }) Attributes.Add(Attribute(name, new StringAttributeMetadata()));
                Attributes.Add(Attribute("expression", new MemoAttributeMetadata())); Attributes.Add(Attribute("ismanaged", new BooleanAttributeMetadata()));
                Attributes.Add(Attribute("ruletype", new PicklistAttributeMetadata()));
                Attributes.Add(Attribute("entitylogicalname", new StringAttributeMetadata()));
                Service = new FakeOrganizationService {
                    ExecuteRequest = r => {
                        if (r is RetrieveMetadataChangesRequest) {
                            var requestScope = (RetrieveMetadataChangesRequest)r;
                            Assert.IsTrue(requestScope.Query.Criteria.Conditions.Count <= 200);
                            var entities = new List<EntityMetadata>();
                            foreach (var condition in requestScope.Query.Criteria.Conditions) {
                                var name = condition.PropertyName == "LogicalName" ? (string)condition.Value : "ava_case";
                                var table = Context.FirstOrDefault(c => c.Record.ComponentType == 1 && c.ComparisonKey == name);
                                var m = new EntityMetadata { LogicalName = name, MetadataId = table?.Record.ObjectId ?? Guid.NewGuid() };
                                Set(m, "ObjectTypeCode", condition.PropertyName == "ObjectTypeCode" ? condition.Value : 1234);
                                entities.Add(m);
                            }
                            var collection = new EntityMetadataCollection(); foreach (var entity in entities) collection.Add(entity);
                            var scopeResponse = new RetrieveMetadataChangesResponse(); scopeResponse.Results["EntityMetadata"] = collection; return scopeResponse;
                        }
                        Assert.IsInstanceOfType(r, typeof(RetrieveEntityRequest)); var request = (RetrieveEntityRequest)r;
                        Assert.AreEqual(BackingEntity, request.LogicalName); Assert.AreEqual(EntityFilters.Entity | EntityFilters.Attributes | EntityFilters.Relationships, request.EntityFilters);
                        Assert.IsFalse(request.RetrieveAsIfPublished); var metadata = new EntityMetadata { LogicalName = request.LogicalName };
                        Set(metadata, "PrimaryIdAttribute", Primary);
                        Set(metadata, "PrimaryNameAttribute", "name");
                        Set(metadata, "Attributes", Attributes.ToArray());
                        var result = new RetrieveEntityResponse(); result.Results["EntityMetadata"] = metadata; return result; },
                    RetrievePage = q => { Queries.Add(q);
                        if (q.EntityName == "solutioncomponentdefinition") {
                            Assert.AreEqual(10276, q.Criteria.Conditions.Single().Values.Single()); return new EntityCollection();
                        }
                        Assert.AreEqual(BackingEntity, q.EntityName); var filter = q.Criteria.Conditions.Single();
                        Assert.AreEqual(Primary, filter.AttributeName); Assert.AreEqual(ConditionOperator.In, filter.Operator); Assert.IsTrue(filter.Values.Count <= 200);
                        return new EntityCollection(Data.Where(row => filter.Values.Contains(row.Id)).Select(row => Clone(row, q)).ToList()); }
                };
            }
            internal Entity Add(Guid? id = null, string unique = "publisher_skill")
            {
                var row = new Entity(EntityName, id ?? Guid.NewGuid()) { ["uniquename"] = unique, ["name"] = "Display skill", ["entitylogicalname"] = "ava_case",
                    ["installationidunique"] = Guid.NewGuid().ToString("D"), ["ismanaged"] = false, ["expression"] = Secret, ["ruletype"] = new OptionSetValue(1) };
                row[Primary] = row.Id; Data.Add(row); Reference(row.Id);
                if (!Context.Any(c => c.Record.ComponentType == 1)) Context.Add(new ComponentIdentity(
                    new SolutionComponentRecord(Guid.NewGuid(), 1, Guid.NewGuid()), IdentityResolutionStatus.Resolved, "ava_case"));
                return row;
            }
            internal void Reference(Guid id) => Raw.Add(new ComponentIdentity(new SolutionComponentRecord(Guid.NewGuid(), 10276, id), IdentityResolutionStatus.Unsupported,
                registeredDefinition: new SolutionComponentDefinitionIdentity(10276, "AiSkillConfig", EntityName)));
            internal Guid Scope(Entity row)
            {
                if (!Attributes.Any(a => a.LogicalName == "parenttableid")) Attributes.Add(Attribute("parenttableid", new LookupAttributeMetadata { Targets = new[] { "entity" } }));
                var parent = Context.Single(c => c.Record.ComponentType == 1).Record.ObjectId.Value;
                row["parenttableid"] = new EntityReference("entity", parent); return parent;
            }
            internal MembershipSnapshot Snapshot() => MembershipSnapshot.Complete(Solution(), Raw.Concat(Context), DateTimeOffset.UtcNow);
        }
#endif
    }
}
