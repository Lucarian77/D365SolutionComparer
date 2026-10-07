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
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Metadata.Query;
using Microsoft.Xrm.Sdk.Query;
using static D365SolutionComparer.Tests.MembershipTestData;

namespace D365SolutionComparer.Tests
{
    [TestClass, TestCategory("Phase2GType511Evidence")]
    public class Type511EvidenceCollectorTests
    {
        [TestMethod]
        public void NormalType511RemainsUnsupportedIndeterminateWithoutProductionKeyOrContract()
        {
            var source = Snapshot(Unsupported()); var target = Snapshot(Unsupported());
            Assert.IsTrue(new SolutionMembershipComparer().Compare(source, target).All(r => r.Presence == MembershipPresence.Indeterminate));
            Assert.IsTrue(source.Components.Concat(target.Components).All(c => c.Status == IdentityResolutionStatus.Unsupported && c.ComparisonKey == null));
            Assert.IsNull(D365SolutionComparer.Models.ComponentDetails.ComponentDefinitionContractCatalog.For(source.Components.Single().SemanticKind));
        }
        private static ComponentIdentity Unsupported() => new ComponentIdentity(new SolutionComponentRecord(Guid.NewGuid(), 511, Guid.NewGuid()),
            IdentityResolutionStatus.Unsupported, registeredDefinition: new SolutionComponentDefinitionIdentity(511, "TeamTemplate", "fixture_team_record"));
        [TestMethod]
        public void CollectorFormAndCapturePathAreAbsentFromRelease()
        {
            var assembly = typeof(SolutionComparerControl).Assembly;
#if DEBUG
            Assert.IsNotNull(assembly.GetType("D365SolutionComparer.Services.Membership.Type511EvidenceCollector"));
            Assert.IsNotNull(typeof(MembershipResultsForm).GetProperty("CaptureType511Evidence", BindingFlags.Instance | BindingFlags.NonPublic));
#else
            Assert.IsNull(assembly.GetType("D365SolutionComparer.Services.Membership.Type511EvidenceCollector"));
            Assert.IsNull(assembly.GetType("D365SolutionComparer.Type511EvidenceResultsForm"));
            Assert.IsNull(typeof(MembershipResultsForm).GetProperty("CaptureType511Evidence", BindingFlags.Instance | BindingFlags.NonPublic));
            Assert.IsNull(typeof(SolutionComparerControl).GetMethod("CaptureType511Evidence", BindingFlags.Instance | BindingFlags.NonPublic));
            Assert.IsFalse(typeof(MembershipCoverageDetailsForm).GetConstructors().SelectMany(c => c.GetParameters()).Any(p => p.Name == "captureType511Evidence"));
#endif
        }
#if DEBUG
        private const string Secret = "PRIVATE-EXPRESSION-CONTENT-AND-SERVICE-DETAILS";
        private const string EntityName = "fixture_team_record", Primary = "fixture_templateid";
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
                        new MembershipCoverageDiagnosticsBuilder().Build(target), presentation, captureType511Evidence: () => calls++))
                    {
                        form.StartPosition = FormStartPosition.Manual; form.Location = new System.Drawing.Point(-4000, -4000); form.Show(); Application.DoEvents();
                        var button = form.Controls.OfType<FlowLayoutPanel>().SelectMany(p => p.Controls.OfType<Button>()).Single(b => b.Text == "Capture Type 511 Team Template Evidence...");
                        Assert.IsTrue(button.Enabled); Assert.AreEqual(0, calls); button.PerformClick(); Assert.AreEqual(1, calls);
                    }
                    Assert.AreEqual(0, pair.Source.Service.Calls + pair.Target.Service.Calls + pair.Source.Service.ExecuteCalls + pair.Target.Service.ExecuteCalls);
                    using (var form = new Type511EvidenceResultsForm("hash-only evidence")) Assert.IsTrue(form.Controls.OfType<RichTextBox>().Single().ReadOnly);
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
        public void RegisteredMappingDiscoversActualEntityAndPrimaryWithoutAssumingTeamtemplateTable()
        {
            var pair = new Pair(); var left = pair.Source.Add(unique: " publisher_Template "); var right = pair.Target.Add(unique: "PUBLISHER_template"); right["ismanaged"] = true;
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
            pair.Source.Raw.Add(new ComponentIdentity(new SolutionComponentRecord(Guid.NewGuid(), 511, row.Id), IdentityResolutionStatus.Unsupported));
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
            if (ambiguous) Assert.IsNull(found.CandidateB); else Assert.IsNotNull(found.CandidateB);
            Assert.IsFalse(found.ParentComplete); Assert.AreEqual(ambiguous ? "Ambiguous" : "Incomplete", found.ParentStatus);
            Assert.IsFalse(report.Pairs.Any(p => p.Outcome == "SemanticPair"));
        }
        [TestMethod]
        public void BlankExposedParentStaysIncomplete()
        {
            var pair = new Pair(); var row = pair.Source.Add(); pair.Source.Scope(row); row.Attributes.Remove("parenttableid");
            var found = pair.Capture().Source.Rows[row.Id]; Assert.IsNull(found.CandidateA); Assert.AreEqual("Incomplete", found.ParentStatus);
        }
        [TestMethod]
        public void CandidateACollisionCannotBeRepairedByDistinctDisplayNamesHashesOrLocalIds()
        {
            var pair = new Pair(); pair.Source.Add(); var extra = pair.Source.Add(unique: "PUBLISHER_template"); extra["name"] = "Other display"; extra["expression"] = "Other expression"; pair.Target.Add();
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
            Assert.IsTrue(pair.Source.Service.Calls <= Type511EvidenceCollector.MaxIsolationGroups + 2);
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
            var pair = new Pair(); for (int i = 0; i < count; i++) pair.Source.Add(unique: "publisher_template_" + i);
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
            foreach (var heading in new[] { "RAW TYPE 511 MEMBERSHIP", "REGISTERED-DEFINITION / BACKING-ENTITY DISCOVERY", "ENTITY / TABLE SCOPE RESOLUTION", "BACKING-RECORD CORRELATION", "READABLE / UNAVAILABLE SCHEMA", "RUNTIME-READABLE / FAULTED COLUMNS",
                "PARENT / REFERENCE RELATIONSHIP DISCOVERY", "CANDIDATE A / B ANALYSIS", "DUPLICATE / COLLISION ANALYSIS", "SOURCE / TARGET FIELD COMPARISON",
                "LIFECYCLE CORRELATION MATRIX", "DIFFERING PRIMARY-ID SEMANTIC PAIRS", "PORTABILITY ASSESSMENT", "EXACT REQUEST LEDGER" }) StringAssert.Contains(text, heading);
            StringAssert.Contains(text, "TotalReads=3"); StringAssert.Contains(text, "NormalMembershipEvidenceRequests=0");
            StringAssert.Contains(text, "AdditionalWhoAmI=0"); StringAssert.Contains(text, "Writes=0");
        }
        [TestMethod]
        public void MetadataConfirmedPrimaryNameCanBeEvaluatedOnlyWithVerifiedTableScope()
        {
            var pair = new Pair(); foreach (var side in new[] { pair.Source, pair.Target }) {
                var row = side.Add(); side.Attributes.RemoveAll(a => a.LogicalName == "uniquename"); row.Attributes.Remove("uniquename");
            }
            var report = pair.Capture(); Assert.AreEqual("SemanticPair", report.Pairs.Single().Outcome);
            Assert.AreEqual("ScopedPrimaryNameHypothesis", report.Source.CandidateRole);
            Assert.AreNotEqual(report.Source.Rows.Values.Single().PrimaryId, report.Target.Rows.Values.Single().PrimaryId);
            pair.Source.Data.Single().Attributes.Remove("entitylogicalname");
            var incomplete = pair.Capture().Source.Rows.Values.Single(); Assert.IsNull(incomplete.CandidateA); Assert.IsNull(incomplete.CandidateB);
        }
        [TestMethod]
        public void DisplayNameAccessRightsAndHashesWithoutTableScopeDoNotCreateIdentity()
        {
            var pair = new Pair(); var row = pair.Source.Add(); pair.Source.Attributes.RemoveAll(a => a.LogicalName == "entitylogicalname");
            row["ruletype"] = new OptionSetValue(99); var report = pair.Capture();
            Assert.IsNull(report.Source.Rows[row.Id].CandidateA); Assert.IsNull(report.Source.Rows[row.Id].CandidateB);
            Assert.IsTrue(report.Pairs.All(p => p.Outcome == "Incomplete"));
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
        public void LargePrimaryNameIsHashOnlyAndCannotManufactureAOrB()
        {
            var pair = new Pair(); var row = pair.Source.Add(); row["name"] = new string('x', 600) + Secret;
            pair.Source.Attributes.RemoveAll(a => a.LogicalName == "uniquename"); row.Attributes.Remove("uniquename");
            var report = pair.Capture(); Assert.IsNull(report.Source.Rows[row.Id].CandidateA); Assert.IsNull(report.Source.Rows[row.Id].CandidateB);
            Assert.IsTrue(report.Source.Rows[row.Id].Content["name"].Known); Assert.IsFalse(report.Build().Contains(Secret));
        }
        [TestMethod]
        public void ObjectTypeCodesUseOneDeduplicatedFilteredMetadataRequestAcrossDifferentCodesAndGuids()
        {
            var pair = new Pair(); foreach (var side in new[] { pair.Source, pair.Target }) {
                var row = side.Add(); side.Attributes.RemoveAll(a => a.LogicalName == "entitylogicalname"); row.Attributes.Remove("entitylogicalname");
                side.Attributes.Add(Attribute("objecttypecode", new IntegerAttributeMetadata())); row["objecttypecode"] = side == pair.Source ? 10001 : 20002;
            }
            var report = pair.Capture(); Assert.AreEqual("SemanticPair", report.Pairs.Single().Outcome);
            Assert.AreEqual(2, pair.Source.Service.ExecuteCalls); Assert.AreEqual(2, pair.Target.Service.ExecuteCalls);
            StringAssert.Contains(report.Build(), "selected entity scopes=[code:10001]");
            StringAssert.Contains(report.Build(), "selected entity scopes=[code:20002]");
        }
        [DataTestMethod, DataRow(false), DataRow(true)]
        public void MissingOrDuplicateScopedMetadataCannotBeRepairedByDisplayName(bool duplicate)
        {
            var pair = new Pair(); var row = pair.Source.Add(); pair.Source.Context.Clear(); var normal = pair.Source.Service.ExecuteRequest;
            pair.Source.Service.ExecuteRequest = r => {
                if (!(r is RetrieveMetadataChangesRequest)) return normal(r);
                var metadata = new EntityMetadataCollection(); if (duplicate) {
                    metadata.Add(new EntityMetadata { LogicalName = "ava_case", MetadataId = Guid.NewGuid() });
                    metadata.Add(new EntityMetadata { LogicalName = "ava_case", MetadataId = Guid.NewGuid() });
                }
                var response = new RetrieveMetadataChangesResponse(); response.Results["EntityMetadata"] = metadata; return response;
            };
            var report = pair.Capture(); Assert.IsNull(report.Source.Rows[row.Id].CandidateA); Assert.IsNull(report.Source.Rows[row.Id].CandidateB);
            Assert.AreEqual(duplicate ? "Ambiguous" : "Incomplete", report.Source.Rows[row.Id].TableStatus);
        }
        [TestMethod]
        public void ConflictingLookupAndScalarTableScopeStaysAmbiguous()
        {
            var pair = new Pair(); var row = pair.Source.Add(); pair.Source.Scope(row);
            row["entitylogicalname"] = "ava_other";
            pair.Source.Context.Add(new ComponentIdentity(new SolutionComponentRecord(Guid.NewGuid(), 1, Guid.NewGuid()), IdentityResolutionStatus.Resolved, "ava_other"));
            var found = pair.Capture().Source.Rows[row.Id]; Assert.AreEqual("Ambiguous", found.TableStatus); Assert.IsNull(found.CandidateA);
        }
        [DataTestMethod, DataRow("entitylogicalname"), DataRow("metadata")]
        public void ScopeRuntimeFaultPreservesCorrelationButBlocksCandidate(string fault)
        {
            var pair = new Pair(); var row = pair.Source.Add();
            if (fault == "metadata") {
                pair.Source.Context.Clear(); var normal = pair.Source.Service.ExecuteRequest;
                pair.Source.Service.ExecuteRequest = r => r is RetrieveMetadataChangesRequest ? throw new InvalidOperationException(Secret) : normal(r);
            }
            else Fail(pair.Source, fault);
            var report = pair.Capture(); Assert.AreEqual("Unique", report.Source.Rows[row.Id].Status); Assert.IsNull(report.Source.Rows[row.Id].CandidateA);
            Assert.IsFalse(report.Build().Contains(Secret));
        }
        [TestMethod]
        public void CancellationDuringScopeMetadataStopsBeforeTarget()
        {
            var pair = new Pair(); pair.Source.Add(); pair.Source.Context.Clear(); pair.Target.Add();
            using (var cancellation = new CancellationTokenSource()) {
                var normal = pair.Source.Service.ExecuteRequest; pair.Source.Service.ExecuteRequest = r => {
                    var result = normal(r); if (r is RetrieveMetadataChangesRequest) cancellation.Cancel(); return result;
                };
                Assert.ThrowsException<OperationCanceledException>(() => pair.Capture(cancellation.Token));
                Assert.AreEqual(0, pair.Target.Service.ExecuteCalls + pair.Target.Service.Calls);
            }
        }
        [TestMethod]
        public void ScopedRegisteredDefinitionDiscoveryDoesNotRequireAssumedBackingEntity()
        {
            var pair = new Pair(); var row = pair.Source.Add(); pair.Source.Raw.Clear();
            pair.Source.Raw.Add(new ComponentIdentity(new SolutionComponentRecord(Guid.NewGuid(), 511, row.Id), IdentityResolutionStatus.Unsupported));
            var normal = pair.Source.Service.RetrievePage;
            pair.Source.Service.RetrievePage = q => q.EntityName == "solutioncomponentdefinition" ? new EntityCollection(new List<Entity> {
                new Entity("solutioncomponentdefinition", Guid.NewGuid()) { ["objecttypecode"] = 511, ["name"] = "TeamTemplate", ["primaryentityname"] = EntityName }
            }) { PagingCookie = "terminal" } : normal(q);
            var report = pair.Capture(); Assert.AreEqual(EntityName, report.Source.EntityName); Assert.AreEqual("Unique", report.Source.Rows[row.Id].Status);
            Assert.IsNotNull(report.Source.Rows[row.Id].CandidateA); Assert.AreEqual(3, pair.Source.Service.Calls);
            StringAssert.Contains(report.Build(), "Discovered registered definition");
        }
        [TestMethod]
        public void ConflictingRegisteredDefinitionRecordsStopBeforeBackingSchemaOrQuery()
        {
            var pair = new Pair(); var row = pair.Source.Add(); pair.Source.Raw.Clear();
            pair.Source.Raw.Add(new ComponentIdentity(new SolutionComponentRecord(Guid.NewGuid(), 511, row.Id), IdentityResolutionStatus.Unsupported));
            pair.Source.Service.RetrievePage = q => new EntityCollection(new List<Entity> {
                new Entity("solutioncomponentdefinition", Guid.NewGuid()) { ["objecttypecode"] = 511 },
                new Entity("solutioncomponentdefinition", Guid.NewGuid()) { ["objecttypecode"] = 511 }
            });
            var report = pair.Capture(); Assert.IsNull(report.Source.EntityName); Assert.AreEqual(0, pair.Source.Service.ExecuteCalls);
            Assert.AreEqual(1, pair.Source.Service.Calls); Assert.IsNull(report.Source.Rows[row.Id].CandidateA);
        }
        [TestMethod]
        public void LegacyCompletedBackingEvidenceIsReportedSeparatelyFromMissingRegisteredDefinition()
        {
            var pair = new Pair(); var row = pair.Source.Add(); pair.Source.BackingEntity = "teamtemplate"; row.LogicalName = "teamtemplate";
            pair.Source.Raw.Clear(); pair.Source.Raw.Add(new ComponentIdentity(new SolutionComponentRecord(Guid.NewGuid(), 511, row.Id),
                IdentityResolutionStatus.Unsupported, semanticKind: ComponentSemanticKinds.TeamTemplate,
                diagnosticEvidence: new[] { "TeamTemplate diagnostic lookup matched. teamtemplateid=" + row.Id.ToString("D") + "; verified local audit" }));
            var report = pair.Capture(); Assert.AreEqual("teamtemplate", report.Source.EntityName); Assert.AreEqual("Unique", report.Source.Rows[row.Id].Status);
            StringAssert.Contains(report.Build(), "No registered solutioncomponentdefinition"); StringAssert.Contains(report.Build(), "independently correlated by completed Type 511 diagnostics");
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
            internal TeamTemplateEvidenceReport Capture(CancellationToken token = default(CancellationToken)) => new Type511EvidenceCollector().Capture(Source.Service, Source.Snapshot(), "1", Target.Service, Target.Snapshot(), "2", token);
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
                            Assert.AreEqual(511, q.Criteria.Conditions.Single().Values.Single()); return new EntityCollection();
                        }
                        Assert.AreEqual(BackingEntity, q.EntityName); var filter = q.Criteria.Conditions.Single();
                        Assert.AreEqual(Primary, filter.AttributeName); Assert.AreEqual(ConditionOperator.In, filter.Operator); Assert.IsTrue(filter.Values.Count <= 200);
                        return new EntityCollection(Data.Where(row => filter.Values.Contains(row.Id)).Select(row => Clone(row, q)).ToList()); }
                };
            }
            internal Entity Add(Guid? id = null, string unique = "publisher_template")
            {
                var row = new Entity(EntityName, id ?? Guid.NewGuid()) { ["uniquename"] = unique, ["name"] = "Display template", ["entitylogicalname"] = "ava_case",
                    ["installationidunique"] = Guid.NewGuid().ToString("D"), ["ismanaged"] = false, ["expression"] = Secret, ["ruletype"] = new OptionSetValue(1) };
                row[Primary] = row.Id; Data.Add(row); Reference(row.Id);
                if (!Context.Any(c => c.Record.ComponentType == 1)) Context.Add(new ComponentIdentity(
                    new SolutionComponentRecord(Guid.NewGuid(), 1, Guid.NewGuid()), IdentityResolutionStatus.Resolved, "ava_case"));
                return row;
            }
            internal void Reference(Guid id) => Raw.Add(new ComponentIdentity(new SolutionComponentRecord(Guid.NewGuid(), 511, id), IdentityResolutionStatus.Unsupported,
                registeredDefinition: new SolutionComponentDefinitionIdentity(511, "TeamTemplate", EntityName)));
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
