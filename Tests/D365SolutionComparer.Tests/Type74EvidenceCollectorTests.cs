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
using Microsoft.Xrm.Sdk.Query;
using static D365SolutionComparer.Tests.MembershipTestData;

namespace D365SolutionComparer.Tests
{
    [TestClass, TestCategory("Phase2GType74Evidence")]
    public class Type74EvidenceCollectorTests
    {
        [TestMethod]
        public void NormalType74RemainsUnsupportedIndeterminateWithoutProductionKeyOrContract()
        {
            var source = Snapshot(Unsupported()); var target = Snapshot(Unsupported());
            Assert.IsTrue(new SolutionMembershipComparer().Compare(source, target).All(r => r.Presence == MembershipPresence.Indeterminate));
            Assert.IsTrue(source.Components.Concat(target.Components).All(c => c.Status == IdentityResolutionStatus.Unsupported && c.ComparisonKey == null));
            Assert.IsNull(D365SolutionComparer.Models.ComponentDetails.ComponentDefinitionContractCatalog.For(source.Components.Single().SemanticKind));
        }
        private static ComponentIdentity Unsupported() => new ComponentIdentity(new SolutionComponentRecord(Guid.NewGuid(), 74, Guid.NewGuid()),
            IdentityResolutionStatus.Unsupported, registeredDefinition: new SolutionComponentDefinitionIdentity(74, "MaskingRule", "fixture_masking_record"));
        [TestMethod]
        public void CollectorFormAndCapturePathAreAbsentFromRelease()
        {
            var assembly = typeof(SolutionComparerControl).Assembly;
#if DEBUG
            Assert.IsNotNull(assembly.GetType("D365SolutionComparer.Services.Membership.Type74EvidenceCollector"));
            Assert.IsNotNull(typeof(MembershipResultsForm).GetProperty("CaptureType74Evidence", BindingFlags.Instance | BindingFlags.NonPublic));
#else
            Assert.IsNull(assembly.GetType("D365SolutionComparer.Services.Membership.Type74EvidenceCollector"));
            Assert.IsNull(assembly.GetType("D365SolutionComparer.Type74EvidenceResultsForm"));
            Assert.IsNull(typeof(MembershipResultsForm).GetProperty("CaptureType74Evidence", BindingFlags.Instance | BindingFlags.NonPublic));
            Assert.IsNull(typeof(SolutionComparerControl).GetMethod("CaptureType74Evidence", BindingFlags.Instance | BindingFlags.NonPublic));
            Assert.IsFalse(typeof(MembershipCoverageDetailsForm).GetConstructors().SelectMany(c => c.GetParameters()).Any(p => p.Name == "captureType74Evidence"));
#endif
        }
#if DEBUG
        private const string Secret = "PRIVATE-EXPRESSION-CONTENT-AND-SERVICE-DETAILS";
        private const string EntityName = "fixture_masking_record", Primary = "fixture_ruleid";
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
                        new MembershipCoverageDiagnosticsBuilder().Build(target), presentation, captureType74Evidence: () => calls++))
                    {
                        form.StartPosition = FormStartPosition.Manual; form.Location = new System.Drawing.Point(-4000, -4000); form.Show(); Application.DoEvents();
                        var button = form.Controls.OfType<FlowLayoutPanel>().SelectMany(p => p.Controls.OfType<Button>()).Single(b => b.Text == "Capture Type 74 MaskingRule Evidence...");
                        Assert.IsTrue(button.Enabled); Assert.AreEqual(0, calls); button.PerformClick(); Assert.AreEqual(1, calls);
                    }
                    Assert.AreEqual(0, pair.Source.Service.Calls + pair.Target.Service.Calls + pair.Source.Service.ExecuteCalls + pair.Target.Service.ExecuteCalls);
                    using (var form = new Type74EvidenceResultsForm("hash-only evidence")) Assert.IsTrue(form.Controls.OfType<RichTextBox>().Single().ReadOnly);
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
        public void RegisteredMappingDiscoversActualEntityAndPrimaryWithoutAssumingMaskingruleTable()
        {
            var pair = new Pair(); var left = pair.Source.Add(unique: " publisher_Rule "); var right = pair.Target.Add(unique: "PUBLISHER_rule"); right["ismanaged"] = true;
            var report = pair.Capture(); var match = report.Pairs.Single(); Assert.AreEqual("SemanticPair", match.Outcome);
            Assert.AreEqual(EntityName, report.Source.EntityName); Assert.AreEqual(Primary, report.Source.PrimaryId);
            Assert.AreEqual(left.Id, match.Source.PrimaryId); Assert.AreNotEqual(left.Id, right.Id);
            Assert.IsTrue(match.Categories.Contains("DifferentPrimaryId") && match.Categories.Contains("UnmanagedToManaged"));
            Assert.AreEqual(2, pair.Source.Service.ExecuteCalls); Assert.AreEqual(2, pair.Source.Service.Calls);
            CollectionAssert.AreEquivalent(new[] { Primary, "uniquename" }, pair.Source.Queries.First().ColumnSet.Columns.ToArray());
        }
        [DataTestMethod, DataRow(false), DataRow(true)]
        public void MissingOrConflictingRegisteredMappingNeverGuessesBackingTable(bool conflict)
        {
            var pair = new Pair(); var row = pair.Source.Add(); pair.Source.Raw.Clear();
            pair.Source.Raw.Add(new ComponentIdentity(new SolutionComponentRecord(Guid.NewGuid(), 74, row.Id), IdentityResolutionStatus.Unsupported));
            if (conflict) pair.Source.Reference(Guid.NewGuid());
            var report = pair.Capture(); Assert.IsNull(report.Source.EntityName); Assert.AreEqual("Incomplete", report.Source.Rows[row.Id].Status);
            Assert.AreEqual(0, pair.Source.Service.ExecuteCalls + pair.Source.Service.Calls);
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
            var report = pair.Capture(); var found = report.Source.Rows[row.Id]; Assert.IsNull(found.CandidateA); Assert.IsNotNull(found.CandidateB);
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
            var pair = new Pair(); pair.Source.Add(); var extra = pair.Source.Add(unique: "PUBLISHER_rule"); extra["name"] = "Other display"; extra["expression"] = "Other expression"; pair.Target.Add();
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
            Assert.IsTrue(pair.Source.Service.Calls <= Type74EvidenceCollector.MaxIsolationGroups + 2);
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
            var pair = new Pair(); for (int i = 0; i < count; i++) pair.Source.Add(unique: "publisher_rule_" + i);
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
            foreach (var heading in new[] { "RAW TYPE 74 MEMBERSHIP", "BACKING-RECORD CORRELATION", "READABLE / UNAVAILABLE SCHEMA", "RUNTIME-READABLE / FAULTED COLUMNS",
                "PARENT / REFERENCE RELATIONSHIP DISCOVERY", "CANDIDATE IDENTITY ANALYSIS", "DUPLICATE / REPEATED ANALYSIS", "SOURCE / TARGET FIELD COMPARISON",
                "LIFECYCLE CORRELATION MATRIX", "DIFFERING PRIMARY-ID SEMANTIC PAIRS", "PORTABILITY ASSESSMENT", "EXACT REQUEST LEDGER" }) StringAssert.Contains(text, heading);
            StringAssert.Contains(text, "TotalReads=4"); StringAssert.Contains(text, "NormalMembershipEvidenceRequests=0");
            StringAssert.Contains(text, "AdditionalWhoAmI=0"); StringAssert.Contains(text, "Writes=0");
        }
        [TestMethod]
        public void OneRelatedAttributeProvidesDiagnosticScopeAcrossDifferentRuleAndAttributeGuids()
        {
            var pair = RelatedPair(); var report = pair.Capture(); var match = report.Pairs.Single();
            Assert.AreEqual("SemanticPair", match.Outcome); Assert.AreNotEqual(match.Source.PrimaryId, match.Target.PrimaryId);
            Assert.AreEqual("Verified", match.Source.AttributeScopeStatus); Assert.AreEqual(1, match.Source.AttributeRelationshipCount);
            CollectionAssert.AreEqual(new[] { "ava_case.ava_private" }, match.Source.AttributeKeys.ToArray());
            Assert.IsTrue(match.Source.CandidateA.StartsWith("maskingrule-attribute-scope-candidate-a:"));
            Assert.AreNotEqual(report.Source.AttributeRows.Values.Single().Get("column_reference"), report.Target.AttributeRows.Values.Single().Get("column_reference"));
            Assert.IsTrue(pair.Source.Queries.Where(q => q.EntityName == RelatedEntity).All(q => q.Criteria.Conditions.Single().AttributeName == RelatedForeign));
            Assert.AreEqual(3, report.Source.Requests.Skip(report.Source.AttributeRequestStart).Take(report.Source.ScopeInvestigationStart - report.Source.AttributeRequestStart).Count());
            foreach (var section in new[] { "ATTRIBUTEMASKINGRULE SCHEMA", "MASKINGRULE -> ATTRIBUTEMASKINGRULE CORRELATION", "RELATED TABLE / ATTRIBUTE PORTABLE SCOPE",
                "SCOPE MULTIPLICITY", "UPDATED CANDIDATE A COMPLETENESS", "EXACT ADDITIONAL ATTRIBUTE SCOPE REQUEST LEDGER" }) StringAssert.Contains(report.Build(), section);
            StringAssert.Contains(report.Build(), "CompleteA=True"); StringAssert.Contains(report.Build(), "AdditionalScopeReads=3");
        }
        [TestMethod]
        public void MultipleRelatedAttributesAreAnOrderIndependentCaseInsensitiveSet()
        {
            var pair = RelatedPair(); pair.Source.Related(pair.Source.Data.Single(), "ava_case.ava_second");
            pair.Target.Related(pair.Target.Data.Single(), "AVA_CASE.AVA_SECOND"); pair.Target.RelatedData.Reverse();
            var match = pair.Capture().Pairs.Single(); Assert.AreEqual("SemanticPair", match.Outcome);
            Assert.AreEqual(2, match.Source.AttributeKeys.Count); Assert.AreEqual(2, match.Source.AttributeRelationshipCount);
            CollectionAssert.AreEqual(match.Source.AttributeKeys.ToArray(), match.Target.AttributeKeys.ToArray());
        }
        [TestMethod]
        public void RelatedScopeBatches201SelectedRulesWithoutPerAttributeRequests()
        {
            var pair = new Pair();
            for (int i = 0; i < 201; i++) {
                var rule = pair.Source.Add(unique: "publisher_rule_" + i);
                pair.Source.Related(rule, "ava_case.ava_attribute_" + i);
                pair.Source.Reference(rule.Id);
            }
            var report = pair.Capture();
            Assert.AreEqual(201, report.Source.AttributeRows.Count);
            Assert.IsTrue(report.Source.Rows.Values.All(r => r.AttributeScopeStatus == "Verified" && r.CandidateA != null));
            var queries = pair.Source.Queries.Where(q => q.EntityName == RelatedEntity).ToArray();
            Assert.AreEqual(4, queries.Length);
            CollectionAssert.AreEqual(new[] { 200, 200, 1, 1 }, queries.Select(q => q.Criteria.Conditions.Single().Values.Count).ToArray());
            Assert.AreEqual(5, report.Source.Requests.Skip(report.Source.AttributeRequestStart).Take(report.Source.ScopeInvestigationStart - report.Source.AttributeRequestStart).Count());
            Assert.AreEqual(0, pair.Source.Service.WriteCalls);
        }
        [TestMethod]
        public void MissingRelatedScopeCannotBeRepairedByEqualCandidateBOrContentHashes()
        {
            var pair = RelatedPair(); pair.Target.RelatedData.Clear(); var report = pair.Capture(); var target = report.Target.Rows.Values.Single();
            Assert.AreEqual("Missing", target.AttributeScopeStatus); Assert.IsNull(target.CandidateA);
            Assert.AreEqual(report.Source.Rows.Values.Single().CandidateB, target.CandidateB);
            Assert.IsFalse(report.Pairs.Any(p => p.Outcome == "SemanticPair"));
        }
        [TestMethod]
        public void NameWithoutVerifiedAttributeScopeCannotCreateCandidateA()
        {
            var pair = RelatedPair(); pair.Source.Context.Clear(); var row = pair.Capture().Source.Rows.Values.Single();
            Assert.AreEqual("Incomplete", row.AttributeScopeStatus); Assert.IsNull(row.CandidateA); Assert.IsNotNull(row.CandidateB);
        }
        [TestMethod]
        public void ConflictingTableAndColumnParentScopeIsAmbiguous()
        {
            var pair = RelatedPair(); var table = Guid.NewGuid();
            pair.Source.RelatedAttributes.Add(Attribute("table_reference", new LookupAttributeMetadata { Targets = new[] { "entity" } }));
            pair.Source.RelatedData.Single()["table_reference"] = new EntityReference("entity", table);
            pair.Source.Context.Add(new ComponentIdentity(new SolutionComponentRecord(Guid.NewGuid(), 1, table), IdentityResolutionStatus.Resolved, "ava_other"));
            var report = pair.Capture(); Assert.AreEqual("Ambiguous", report.Source.Rows.Values.Single().AttributeScopeStatus);
            Assert.IsNull(report.Source.Rows.Values.Single().CandidateA); StringAssert.Contains(report.Build(), "Conflicting table versus resolved column parent scope");
        }
        [DataTestMethod, DataRow(false), DataRow(true)]
        public void DuplicateRelationshipRowsOrRepeatedPortableAttributeStayAmbiguous(bool samePrimary)
        {
            var pair = RelatedPair(); var original = pair.Source.RelatedData.Single(); var duplicate = new Entity(RelatedEntity, samePrimary ? original.Id : Guid.NewGuid());
            foreach (var field in original.Attributes) duplicate[field.Key] = field.Value; duplicate[RelatedPrimary] = duplicate.Id; pair.Source.RelatedData.Add(duplicate);
            var report = pair.Capture(); Assert.AreEqual("Ambiguous", report.Source.Rows.Values.Single().AttributeScopeStatus);
            Assert.IsNull(report.Source.Rows.Values.Single().CandidateA); Assert.IsFalse(report.Pairs.Any(p => p.Outcome == "SemanticPair"));
        }
        [DataTestMethod, DataRow("expression", true), DataRow("column_reference", false), DataRow(RelatedForeign, false)]
        public void RelatedRuntimeFaultIsolationPreservesBackingCorrelationAndCriticalGuards(string field, bool complete)
        {
            var pair = RelatedPair(); var normal = pair.Source.Service.RetrievePage;
            pair.Source.Service.RetrievePage = q => { if (q.EntityName == RelatedEntity && q.ColumnSet.Columns.Contains(field))
                throw new FaultException<OrganizationServiceFault>(new OrganizationServiceFault { ErrorCode = unchecked((int)0x8004023B), Message = Secret }, new FaultReason(Secret)); return normal(q); };
            var report = pair.Capture(); var root = report.Source.Rows.Values.Single(); Assert.AreEqual("Unique", root.Status);
            Assert.AreEqual(complete, root.CandidateA != null); Assert.IsFalse(report.Build().Contains(Secret)); StringAssert.Contains(report.Build(), "0x8004023B");
            Assert.IsTrue(pair.Source.Queries.Where(q => q.EntityName == RelatedEntity).All(q => q.Criteria.Conditions.Single().AttributeName == RelatedForeign));
            Assert.IsTrue(pair.Source.Service.Calls <= Type74EvidenceCollector.MaxIsolationGroups + 5);
        }
        [TestMethod]
        public void RelatedPagingAcceptsTerminalCookieAndDeduplicatesIdenticalCrossPageRows()
        {
            var pair = RelatedPair(); pair.Source.Related(pair.Source.Data.Single(), "ava_case.ava_second"); var normal = pair.Source.Service.RetrievePage;
            pair.Source.Service.RetrievePage = q => {
                if (q.EntityName != RelatedEntity) return normal(q);
                var first = Clone(pair.Source.RelatedData[0], q); var second = Clone(pair.Source.RelatedData[1], q);
                return q.PageInfo.PageNumber == 1 ? new EntityCollection(new List<Entity> { first }) { MoreRecords = true, PagingCookie = "next" }
                    : new EntityCollection(new List<Entity> { first, second }) { MoreRecords = false, PagingCookie = "terminal" };
            };
            var report = pair.Capture(); var root = report.Source.Rows.Values.Single(); Assert.AreEqual("Verified", root.AttributeScopeStatus);
            Assert.AreEqual(2, root.AttributeKeys.Count); StringAssert.Contains(report.Build(), "pageCount=2");
        }
        [TestMethod]
        public void ForeignAssociationOrStalledRelatedPagingCannotProveScope()
        {
            var pair = RelatedPair(); pair.Source.RelatedData.Single()[RelatedForeign] = new EntityReference(EntityName, Guid.NewGuid());
            var normal = pair.Source.Service.RetrievePage;
            pair.Source.Service.RetrievePage = q => q.EntityName == RelatedEntity ? new EntityCollection(new List<Entity> { Clone(pair.Source.RelatedData.Single(), q) }) : normal(q);
            var root = pair.Capture().Source.Rows.Values.Single(); Assert.AreEqual("Incomplete", root.AttributeScopeStatus); Assert.IsNull(root.CandidateA);
            pair = RelatedPair(); normal = pair.Source.Service.RetrievePage;
            pair.Source.Service.RetrievePage = q => { var result = normal(q); if (q.EntityName == RelatedEntity) { result.MoreRecords = true; result.PagingCookie = "same"; } return result; };
            root = pair.Capture().Source.Rows.Values.Single(); Assert.AreEqual("Incomplete", root.AttributeScopeStatus); Assert.IsNull(root.CandidateA);
        }
        [TestMethod]
        public void CancellationDuringRelatedReadStopsBeforeTargetAndChangesNoMembership()
        {
            var pair = RelatedPair(); using (var cancellation = new CancellationTokenSource()) {
                var normal = pair.Source.Service.RetrievePage; pair.Source.Service.RetrievePage = q => { if (q.EntityName == RelatedEntity) cancellation.Cancel(); return normal(q); };
                Assert.ThrowsException<OperationCanceledException>(() => pair.Capture(cancellation.Token)); Assert.AreEqual(0, pair.Target.Service.Calls + pair.Target.Service.ExecuteCalls);
            }
            Assert.IsTrue(pair.Source.Raw.Concat(pair.Target.Raw).All(c => c.Status == IdentityResolutionStatus.Unsupported && c.ComparisonKey == null));
            Assert.AreEqual(0, pair.Source.Service.WriteCalls + pair.Target.Service.WriteCalls);
        }
        [TestMethod]
        public void UnreadableAttributeReferenceDoesNotTriggerGuessedFieldOrParentQuery()
        {
            var pair = RelatedPair(); Set(pair.Source.RelatedAttributes.Single(a => a.LogicalName == "column_reference"), "IsValidForRead", false);
            var root = pair.Capture().Source.Rows.Values.Single(); Assert.AreEqual("Incomplete", root.AttributeScopeStatus); Assert.IsNull(root.CandidateA);
            Assert.IsTrue(pair.Source.Queries.All(q => q.EntityName != RelatedEntity || !q.ColumnSet.Columns.Contains("column_reference")));
        }
        [TestMethod]
        public void ScopeInventoryIndependentlyReusesPortableColumnAcrossDifferentMetadataGuids()
        {
            var pair = RelatedPair(); var report = pair.Capture();
            foreach (var side in new[] { report.Source, report.Target }) {
                var rule = side.Rows.Values.Single(); var related = side.AttributeRows.Values.Single();
                Assert.IsTrue(rule.RuleIdentifierComplete && rule.RecoveredScopeComplete); Assert.IsNotNull(rule.CandidateA);
                var column = related.ScopeSources.Single(p => p.Field == "column_reference");
                Assert.AreEqual("Both", column.Establishes); Assert.AreEqual("ava_case.ava_private", column.ColumnKey);
                Assert.AreEqual("VerifiedSnapshotScope", column.Status); Assert.IsTrue(column.RuntimeReadable);
                Assert.AreEqual(0, side.ScopeAttributeSubrequests);
            }
            Assert.AreNotEqual(report.Source.Rows.Values.Single().PrimaryId, report.Target.Rows.Values.Single().PrimaryId);
            Assert.AreEqual(report.Source.Rows.Values.Single().CandidateA, report.Target.Rows.Values.Single().CandidateA);
            StringAssert.Contains(report.Build(), "Portable scope independently established; continue readiness review");
        }

        [DataTestMethod, DataRow(false), DataRow(true)]
        public void MissingOrAmbiguousAttributeSnapshotScopeCannotBeRepairedByRuleNameOrHash(bool ambiguous)
        {
            var pair = RelatedPair(); var selected = pair.Source.Context.Single(c => c.Record.ComponentType == 2);
            if (ambiguous) pair.Source.Context.Add(new ComponentIdentity(new SolutionComponentRecord(Guid.NewGuid(), 2, selected.Record.ObjectId), IdentityResolutionStatus.Ambiguous));
            else pair.Source.Context.Remove(selected);
            var report = pair.Capture(); var rule = report.Source.Rows.Values.Single();
            Assert.IsTrue(rule.RuleIdentifierComplete); Assert.IsFalse(rule.RecoveredScopeComplete); Assert.IsNull(rule.CandidateA);
            Assert.IsNotNull(rule.CandidateB); Assert.AreEqual(0, report.Source.ScopeAttributeSubrequests);
            StringAssert.Contains(report.Build(), "No independent scope source found; Type 74 should remain unsupported");
        }

        [TestMethod]
        public void ZeroSelectedRelatedRowsAreNotGlobalScopeEvenWhenNamesIdsAndHashesAgree()
        {
            var pair = RelatedPair(); var root = pair.Source.Data.Single(); pair.Source.RelatedData.Clear(); pair.Target.RelatedData.Clear();
            var report = pair.Capture(); foreach (var side in new[] { report.Source, report.Target }) {
                var row = side.Rows.Values.Single(); Assert.AreEqual(0, row.AttributeRelationshipCount);
                Assert.AreEqual("Missing", row.AttributeScopeStatus); Assert.IsFalse(row.RecoveredScopeComplete); Assert.IsNull(row.CandidateA);
            }
            StringAssert.Contains(report.Build(), "No selected related rows returned != Rule is global");
            StringAssert.Contains(report.Build(), "RuntimeReadability=NotObservedNoSelectedRowsOrIncompleteRetrieval; ActualValue=NoSelectedRowValue; Establishes=Neither");
            StringAssert.Contains(report.Build(), "TerminalZeroRelatedRowsEvidence=True");
            StringAssert.Contains(report.Build(), "IndependentlyVerifiedGlobalScope=False");
            Assert.IsFalse(report.Pairs.Any(p => p.Outcome == "SemanticPair"));
            Assert.AreEqual(0, report.Source.ScopeAttributeSubrequests + report.Target.ScopeAttributeSubrequests);
        }

        [TestMethod]
        public void ApparentGlobalMarkerAndCatalogLabelDoNotProveASelectedRuleIsGlobal()
        {
            var pair = RelatedPair(); foreach (var side in new[] { pair.Source, pair.Target }) {
                side.RelatedData.Clear(); side.Attributes.Add(Attribute("isglobal", new BooleanAttributeMetadata())); side.Data.Single()["isglobal"] = true;
            }
            var report = pair.Capture(); Assert.IsTrue(report.Source.Rows.Values.All(r => !r.RecoveredScopeComplete && r.CandidateA == null));
            var probe = report.Source.Rows.Values.Single().ScopeSources.Single(p => p.Field == "isglobal");
            Assert.AreEqual("True", probe.Value); Assert.AreEqual("Neither", probe.Establishes); Assert.AreEqual("ObservedSemanticsUnproven", probe.Status);
            StringAssert.Contains(report.Build(), "not selected rule scope or a global-scope contract");
        }

        [TestMethod]
        public void TableLookupIsPortableScopeAndItsLookupNameShadowIsExcluded()
        {
            var pair = new Pair(); var backing = pair.Source.Add(); pair.Source.Scope(backing);
            pair.Source.Attributes.Add(Attribute("parenttableidname", new StringAttributeMetadata())); backing["parenttableidname"] = "Display scope";
            var report = pair.Capture(); var row = report.Source.Rows[backing.Id];
            Assert.IsTrue(row.RecoveredScopeComplete); CollectionAssert.AreEqual(new[] { "ava_case" }, row.RecoveredTables.ToArray());
            Assert.AreEqual("TableIdentity", row.ScopeSources.Single(p => p.Field == "parenttableid").Establishes);
            var shadow = row.ScopeSources.Single(p => p.Field == "parenttableidname"); Assert.AreEqual("ExcludedLookupShadow", shadow.Status); Assert.AreEqual("Neither", shadow.Establishes);
            Assert.IsFalse(pair.Source.Queries.Any(q => q.ColumnSet.Columns.Contains("parenttableidname")));
        }

        [TestMethod]
        public void SelectedAttributeMetadataCanRecoverScopeWithoutRewritingExistingCandidateA()
        {
            var pair = ScopedMetadataPair(); SupplyAttributeMetadata(pair.Source); SupplyAttributeMetadata(pair.Target);
            var report = pair.Capture(); foreach (var side in new[] { report.Source, report.Target }) {
                var rule = side.Rows.Values.Single(); Assert.IsTrue(rule.RecoveredScopeComplete); Assert.IsNull(rule.CandidateA); // Original snapshot-only A contract remains unchanged.
                CollectionAssert.AreEqual(new[] { "ava_case.ava_selected" }, rule.RecoveredColumns.ToArray());
                Assert.AreEqual(1, side.ScopeMetadataBatches); Assert.AreEqual(1, side.ScopeAttributeSubrequests);
                Assert.AreEqual("VerifiedSelectedAttributeMetadata", side.AttributeRows.Values.Single().ScopeSources.Single(p => p.Field == "column_reference").Status);
            }
            Assert.IsFalse(report.Pairs.Any(p => p.Outcome == "SemanticPair"));
            Assert.IsTrue(pair.Source.Raw.Concat(pair.Target.Raw).All(r => r.Status == IdentityResolutionStatus.Unsupported && r.ComparisonKey == null));
        }

        [DataTestMethod, DataRow("duplicate"), DataRow("wrongId"), DataRow("wrongTable"), DataRow("fault"), DataRow("missing")]
        public void ScopedAttributeMetadataFailureRemainsIncompleteOrAmbiguous(string failure)
        {
            var pair = ScopedMetadataPair(); SupplyAttributeMetadata(pair.Source, failure);
            var report = pair.Capture(); var row = report.Source.Rows.Values.Single();
            Assert.IsFalse(row.RecoveredScopeComplete); Assert.IsNull(row.CandidateA);
            Assert.IsFalse(report.Build().Contains(Secret)); Assert.AreEqual(1, report.Source.ScopeAttributeSubrequests);
        }

        [TestMethod]
        public void MetadataIdsWithoutAProvenAttributeRoleCannotCreateScopeOrTriggerGlobalMetadataReads()
        {
            var pair = RelatedPair(); pair.Source.RelatedData.Clear();
            pair.Source.Attributes.Add(Attribute("attributemetadataid", new UniqueIdentifierAttributeMetadata())); pair.Source.Data.Single()["attributemetadataid"] = Guid.NewGuid();
            var report = pair.Capture(); var probe = report.Source.Rows.Values.Single().ScopeSources.Single(p => p.Field == "attributemetadataid");
            Assert.AreEqual("Neither", probe.Establishes); Assert.IsFalse(report.Source.Rows.Values.Single().RecoveredScopeComplete);
            Assert.AreEqual(0, report.Source.ScopeAttributeSubrequests);
        }

        [TestMethod]
        public void SelectedTableEntityNameIsVerifiedByBoundedMetadataWithoutReadingAllAttributes()
        {
            var pair = RelatedPair(); pair.Source.RelatedData.Clear();
            pair.Source.Attributes.Add(Attribute("objecttypecode", new EntityNameAttributeMetadata())); pair.Source.Data.Single()["objecttypecode"] = "ava_case";
            var normal = pair.Source.Service.ExecuteRequest;
            pair.Source.Service.ExecuteRequest = request => {
                if (!(request is RetrieveMetadataChangesRequest)) return normal(request);
                var query = ((RetrieveMetadataChangesRequest)request).Query;
                Assert.IsNull(query.AttributeQuery); Assert.AreEqual(1, query.Criteria.Conditions.Count);
                Assert.AreEqual("LogicalName", query.Criteria.Conditions[0].PropertyName); Assert.AreEqual("ava_case", query.Criteria.Conditions[0].Value);
                var entity = new EntityMetadata { LogicalName = "ava_case", MetadataId = Guid.NewGuid() };
                var response = new RetrieveMetadataChangesResponse(); response.Results["EntityMetadata"] = new EntityMetadataCollection { entity }; return response;
            };
            var report = pair.Capture(); var probe = report.Source.Rows.Values.Single().ScopeSources.Single(p => p.Field == "objecttypecode");
            Assert.AreEqual("VerifiedSelectedTableMetadata", probe.Status); Assert.AreEqual("TableIdentity", probe.Establishes);
            Assert.IsTrue(report.Source.Rows.Values.Single().RecoveredScopeComplete); Assert.IsNull(report.Source.Rows.Values.Single().CandidateA);
        }

        [TestMethod]
        public void DefinitionOnlyChangeDoesNotChangeExistingCandidateOrScopeEvidence()
        {
            var pair = RelatedPair(); pair.Target.Data.Single()["expression"] = "PRIVATE CHANGED CONFIGURATION";
            var report = pair.Capture(); var matched = report.Pairs.Single();
            Assert.AreEqual(matched.Source.CandidateA, matched.Target.CandidateA); Assert.IsTrue(matched.Source.RecoveredScopeComplete && matched.Target.RecoveredScopeComplete);
            Assert.IsTrue(matched.Categories.Contains("DifferentDefinitionEvidence")); Assert.IsFalse(report.Build().Contains("PRIVATE CHANGED CONFIGURATION"));
        }

        [TestMethod]
        public void ScopeAttributeBatchesAreDeduplicatedAndBoundedAt200()
        {
            var pair = new Pair(); var root = pair.Source.Add(); pair.Source.Scope(root);
            for (int i = 0; i < 201; i++) pair.Source.Related(root, "ava_case.ava_attribute_" + i);
            pair.Source.Context.RemoveAll(c => c.Record.ComponentType == 2);
            SupplyAttributeMetadata(pair.Source, "uniqueColumns");
            var report = pair.Capture(); Assert.AreEqual(2, report.Source.ScopeMetadataBatches); Assert.AreEqual(201, report.Source.ScopeAttributeSubrequests);
            Assert.IsTrue(report.Source.AttributeRows.Values.All(r => r.RecoveredScopeComplete));
            Assert.AreEqual(0, pair.Target.Service.Calls + pair.Target.Service.ExecuteCalls);
        }

        [TestMethod]
        public void CancellationDuringScopedAttributeMetadataDoesNotContinueIntoTargetOrMutateMembership()
        {
            var pair = ScopedMetadataPair(); var normal = pair.Source.Service.ExecuteRequest;
            using (var cancel = new CancellationTokenSource()) {
                pair.Source.Service.ExecuteRequest = request => { if (request is ExecuteMultipleRequest) { cancel.Cancel(); cancel.Token.ThrowIfCancellationRequested(); } return normal(request); };
                Assert.ThrowsException<OperationCanceledException>(() => pair.Capture(cancel.Token));
                Assert.AreEqual(0, pair.Target.Service.Calls + pair.Target.Service.ExecuteCalls);
                Assert.IsTrue(pair.Source.Raw.All(r => r.Status == IdentityResolutionStatus.Unsupported));
                Assert.AreEqual(0, pair.Source.Service.WriteCalls + pair.Target.Service.WriteCalls);
            }
        }

        [TestMethod]
        public void DifferentSelectedAttributeIdsReturningSamePortableColumnRemainAmbiguous()
        {
            var pair = ScopedMetadataPair(); pair.Source.Related(pair.Source.Data.Single(), "ava_case.other"); pair.Source.Context.RemoveAll(c => c.Record.ComponentType == 2);
            SupplyAttributeMetadata(pair.Source); var report = pair.Capture();
            Assert.IsFalse(report.Source.Rows.Values.Single().RecoveredScopeComplete); Assert.IsNull(report.Source.Rows.Values.Single().CandidateA);
            Assert.IsTrue(report.Source.AttributeRows.Values.All(r => !r.RecoveredScopeComplete));
            Assert.IsTrue(report.Source.AttributeRows.Values.SelectMany(r => r.ScopeSources).Any(p => p.Status.StartsWith("AmbiguousCanonicalColumnScope")));
        }

        [TestMethod]
        public void ContradictoryRootTableAndRelatedPortableColumnDoNotProvideCompleteScope()
        {
            var pair = RelatedPair(); var root = pair.Source.Data.Single(); var id = pair.Source.Scope(root);
            pair.Source.Context.RemoveAll(c => c.Record.ComponentType == 1);
            pair.Source.Context.Add(new ComponentIdentity(new SolutionComponentRecord(Guid.NewGuid(), 1, id), IdentityResolutionStatus.Resolved, "other_table"));
            var report = pair.Capture(); Assert.IsFalse(report.Source.Rows[root.Id].RecoveredScopeComplete);
            StringAssert.Contains(report.Source.Rows[root.Id].ScopeRecoveryReason, "Conflicting");
        }

        [TestMethod]
        public void RegisteredGlobalMarkerIsAuditedWithPagingCookieButCannotSupplyRuleScope()
        {
            var pair = RelatedPair(); pair.Source.RelatedData.Clear(); var normal = pair.Source.Service.ExecuteRequest;
            pair.Source.Service.ExecuteRequest = request => {
                if (!(request is RetrieveEntityRequest) || ((RetrieveEntityRequest)request).LogicalName != "solutioncomponentdefinition") return normal(request);
                var metadata = new EntityMetadata { LogicalName = "solutioncomponentdefinition" }; Set(metadata, "PrimaryIdAttribute", "solutioncomponentdefinitionid");
                Set(metadata, "Attributes", new[] { Attribute("solutioncomponentdefinitionid", new UniqueIdentifierAttributeMetadata()), Attribute("objecttypecode", new IntegerAttributeMetadata()),
                    Attribute("name", new StringAttributeMetadata()), Attribute("primaryentityname", new StringAttributeMetadata()), Attribute("isglobal", new BooleanAttributeMetadata()) });
                var response = new RetrieveEntityResponse(); response.Results["EntityMetadata"] = metadata; return response;
            };
            var registration = new Entity("solutioncomponentdefinition", Guid.NewGuid()); registration["solutioncomponentdefinitionid"] = registration.Id;
            registration["objecttypecode"] = 74; registration["name"] = "MaskingRule"; registration["primaryentityname"] = EntityName; registration["isglobal"] = true;
            var normalPage = pair.Source.Service.RetrievePage;
            pair.Source.Service.RetrievePage = query => {
                if (query.EntityName != "solutioncomponentdefinition") return normalPage(query);
                Assert.AreEqual(ConditionOperator.Equal, query.Criteria.Conditions.Single().Operator); Assert.AreEqual(74, query.Criteria.Conditions.Single().Values.Single());
                Assert.AreEqual(200, query.PageInfo.Count); return new EntityCollection(new List<Entity> { registration }) { MoreRecords = false, PagingCookie = "private-cookie" };
            };
            var report = pair.Capture(); StringAssert.Contains(report.Build(), "Field=isglobal; ActualValue=True; Establishes=Neither");
            StringAssert.Contains(report.Build(), "distinctRows=1; pageCount=1"); Assert.IsFalse(report.Build().Contains("private-cookie"));
            Assert.IsFalse(report.Source.Rows.Values.Single().RecoveredScopeComplete); Assert.IsNull(report.Source.Rows.Values.Single().CandidateA);
            Assert.AreEqual(report.Source.Requests.Count, pair.Source.Service.Calls + pair.Source.Service.ExecuteCalls);
        }

        [TestMethod]
        public void ScopeSourceInventoryDoesNotChangeExistingCandidateAOrCandidateBConstruction()
        {
            var pair = RelatedPair(); var report = pair.Capture(); var left = report.Source.Rows.Values.Single(); var right = report.Target.Rows.Values.Single();
            Assert.AreEqual(left.CandidateA, right.CandidateA); Assert.IsTrue(left.CandidateA.StartsWith("maskingrule-attribute-scope-candidate-a:"));
            Assert.AreEqual(left.CandidateB, right.CandidateB); Assert.IsTrue(pair.Source.Raw.Concat(pair.Target.Raw).All(c => c.Status == IdentityResolutionStatus.Unsupported && c.ComparisonKey == null));
            Assert.IsFalse(report.Build().Contains("CandidateP=")); StringAssert.Contains(report.Build(), "no Candidate P");
            Assert.AreEqual(0, pair.Source.Service.WriteCalls + pair.Target.Service.WriteCalls);
        }

        [TestMethod]
        public void ScopeNamedMemoOrConfigurationRemainsHashOnlyEvenWhenTextLooksLikeALogicalName()
        {
            var pair = RelatedPair(); var root = pair.Source.Data.Single();
            pair.Source.Attributes.Add(Attribute("scopeconfiguration", new StringAttributeMetadata())); root["scopeconfiguration"] = "ShortSensitiveConfiguration";
            pair.Source.Attributes.Add(Attribute("attributescope", new MemoAttributeMetadata())); root["attributescope"] = "ShortSensitiveMemo";
            var report = pair.Capture(); Assert.IsFalse(report.Build().Contains("ShortSensitiveConfiguration")); Assert.IsFalse(report.Build().Contains("ShortSensitiveMemo"));
            Assert.IsTrue(report.Source.Rows[root.Id].ScopeSources.Where(p => p.Field == "scopeconfiguration" || p.Field == "attributescope").All(p => p.Value.StartsWith("HashOnly;")));
            Assert.IsTrue(report.Source.Rows[root.Id].RecoveredScopeComplete);
        }

        private static Pair ScopedMetadataPair()
        {
            var pair = new Pair(); foreach (var side in new[] { pair.Source, pair.Target }) {
                var root = side.Add(); side.Scope(root); side.Related(root, "ava_case.ava_selected"); side.Context.RemoveAll(c => c.Record.ComponentType == 2);
            } return pair;
        }
        private static void SupplyAttributeMetadata(Fixture side, string failure = null)
        {
            var normal = side.Service.ExecuteRequest;
            side.Service.ExecuteRequest = request => {
                var batch = request as ExecuteMultipleRequest; if (batch == null) return normal(request);
                Assert.IsTrue(batch.Requests.Count <= 200 && batch.Requests.Count > 0);
                var response = new ExecuteMultipleResponse(); var items = new ExecuteMultipleResponseItemCollection(); response.Results["Responses"] = items;
                for (int index = 0; index < batch.Requests.Count; index++) {
                    var selected = (RetrieveAttributeRequest)batch.Requests[index]; Assert.AreEqual("ava_case", selected.EntityLogicalName); Assert.IsFalse(selected.RetrieveAsIfPublished);
                    Assert.IsTrue(side.RelatedData.Any(r => ((EntityReference)r["column_reference"]).Id == selected.MetadataId));
                    if (failure == "missing") continue;
                    var attribute = new StringAttributeMetadata { LogicalName = failure == "uniqueColumns" ? "ava_column_" + side.RelatedData.FindIndex(r => ((EntityReference)r["column_reference"]).Id == selected.MetadataId) : "ava_selected", MetadataId = failure == "wrongId" ? Guid.NewGuid() : selected.MetadataId };
                    Set(attribute, "EntityLogicalName", failure == "wrongTable" ? "other_table" : "ava_case");
                    var result = new RetrieveAttributeResponse(); result.Results["AttributeMetadata"] = attribute;
                    items.Add(new ExecuteMultipleResponseItem { RequestIndex = index, Response = failure == "fault" ? null : result,
                        Fault = failure == "fault" ? new OrganizationServiceFault { ErrorCode = unchecked((int)0x8004023B), Message = Secret, TraceText = Secret } : null });
                    if (failure == "duplicate") items.Add(new ExecuteMultipleResponseItem { RequestIndex = index, Response = result });
                } return response;
            };
        }

        private const string RelatedEntity = "attributemaskingrule", RelatedPrimary = "fixture_associationid", RelatedForeign = "fixture_maskingrule_link";
        private static Pair RelatedPair()
        {
            var pair = new Pair(); foreach (var side in new[] { pair.Source, pair.Target }) {
                var root = side.Add(); side.Attributes.RemoveAll(a => a.LogicalName == "uniquename"); root.Attributes.Remove("uniquename");
                side.Related(root, "ava_case.ava_private");
            } return pair;
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
            internal MaskingRuleEvidenceReport Capture(CancellationToken token = default(CancellationToken)) => new Type74EvidenceCollector().Capture(Source.Service, Source.Snapshot(), "1", Target.Service, Target.Snapshot(), "2", token);
        }
        private sealed class Fixture
        {
            internal readonly List<Entity> Data = new List<Entity>();
            internal readonly List<ComponentIdentity> Raw = new List<ComponentIdentity>(), Context = new List<ComponentIdentity>();
            internal readonly List<AttributeMetadata> Attributes = new List<AttributeMetadata>();
            internal readonly List<QueryExpression> Queries = new List<QueryExpression>();
            internal readonly List<Entity> RelatedData = new List<Entity>();
            internal readonly List<AttributeMetadata> RelatedAttributes = new List<AttributeMetadata>();
            internal OneToManyRelationshipMetadata Incoming;
            internal readonly FakeOrganizationService Service;
            internal Fixture()
            {
                Attributes.Add(Attribute(Primary, new UniqueIdentifierAttributeMetadata()));
                foreach (var name in new[] { "name", "uniquename", "installationidunique" }) Attributes.Add(Attribute(name, new StringAttributeMetadata()));
                Attributes.Add(Attribute("expression", new MemoAttributeMetadata())); Attributes.Add(Attribute("ismanaged", new BooleanAttributeMetadata()));
                Attributes.Add(Attribute("ruletype", new PicklistAttributeMetadata()));
                Service = new FakeOrganizationService {
                    ExecuteRequest = r => { Assert.IsInstanceOfType(r, typeof(RetrieveEntityRequest)); var request = (RetrieveEntityRequest)r;
                        Assert.IsTrue(new[] { EntityName, RelatedEntity, "solutioncomponentdefinition" }.Contains(request.LogicalName)); Assert.AreEqual(EntityFilters.Attributes | EntityFilters.Relationships, request.EntityFilters);
                        Assert.IsFalse(request.RetrieveAsIfPublished); var metadata = new EntityMetadata { LogicalName = request.LogicalName };
                        Set(metadata, "PrimaryIdAttribute", request.LogicalName == EntityName ? Primary : request.LogicalName == RelatedEntity ? RelatedPrimary : "solutioncomponentdefinitionid");
                        Set(metadata, "Attributes", request.LogicalName == EntityName ? Attributes.ToArray() : request.LogicalName == RelatedEntity ? RelatedAttributes.ToArray() :
                            new[] { Attribute("solutioncomponentdefinitionid", new UniqueIdentifierAttributeMetadata()) });
                        if (Incoming != null) {
                            Set(metadata, "OneToManyRelationships", request.LogicalName == EntityName ? new[] { Incoming } : new OneToManyRelationshipMetadata[0]);
                            Set(metadata, "ManyToOneRelationships", request.LogicalName == RelatedEntity ? new[] { Incoming } : new OneToManyRelationshipMetadata[0]);
                        }
                        var result = new RetrieveEntityResponse(); result.Results["EntityMetadata"] = metadata; return result; },
                    RetrievePage = q => { Queries.Add(q); Assert.IsTrue(new[] { EntityName, RelatedEntity }.Contains(q.EntityName)); var filter = q.Criteria.Conditions.Single();
                        Assert.AreEqual(q.EntityName == EntityName ? Primary : RelatedForeign, filter.AttributeName); Assert.AreEqual(ConditionOperator.In, filter.Operator); Assert.IsTrue(filter.Values.Count <= 200);
                        return new EntityCollection((q.EntityName == EntityName ? Data.Where(row => filter.Values.Contains(row.Id)) :
                            RelatedData.Where(row => filter.Values.Contains(((EntityReference)row[RelatedForeign]).Id))).Select(row => Clone(row, q)).ToList()); }
                };
            }
            internal Entity Add(Guid? id = null, string unique = "publisher_rule")
            {
                var row = new Entity(EntityName, id ?? Guid.NewGuid()) { ["uniquename"] = unique, ["name"] = "Display rule",
                    ["installationidunique"] = Guid.NewGuid().ToString("D"), ["ismanaged"] = false, ["expression"] = Secret, ["ruletype"] = new OptionSetValue(1) };
                row[Primary] = row.Id; Data.Add(row); Reference(row.Id); return row;
            }
            internal void Reference(Guid id) => Raw.Add(new ComponentIdentity(new SolutionComponentRecord(Guid.NewGuid(), 74, id), IdentityResolutionStatus.Unsupported,
                registeredDefinition: new SolutionComponentDefinitionIdentity(74, "MaskingRule", EntityName)));
            internal Guid Scope(Entity row)
            {
                if (!Attributes.Any(a => a.LogicalName == "parenttableid")) Attributes.Add(Attribute("parenttableid", new LookupAttributeMetadata { Targets = new[] { "entity" } }));
                var parent = Guid.NewGuid(); row["parenttableid"] = new EntityReference("entity", parent);
                Context.Add(new ComponentIdentity(new SolutionComponentRecord(Guid.NewGuid(), 1, parent), IdentityResolutionStatus.Resolved, "ava_case")); return parent;
            }
            internal Entity Related(Entity rule, string portableColumn)
            {
                if (Incoming == null) {
                    Incoming = new OneToManyRelationshipMetadata { ReferencedEntity = EntityName, ReferencedAttribute = Primary,
                        ReferencingEntity = RelatedEntity, ReferencingAttribute = RelatedForeign, SchemaName = "fixture_incoming" };
                    RelatedAttributes.Add(Attribute(RelatedPrimary, new UniqueIdentifierAttributeMetadata()));
                    RelatedAttributes.Add(Attribute(RelatedForeign, new LookupAttributeMetadata { Targets = new[] { EntityName } }));
                    RelatedAttributes.Add(Attribute("column_reference", new LookupAttributeMetadata { Targets = new[] { "attribute" } }));
                    RelatedAttributes.Add(Attribute("expression", new MemoAttributeMetadata()));
                }
                var column = Guid.NewGuid(); var row = new Entity(RelatedEntity, Guid.NewGuid()) {
                    [RelatedForeign] = new EntityReference(EntityName, rule.Id), ["column_reference"] = new EntityReference("attribute", column), ["expression"] = Secret };
                row[RelatedPrimary] = row.Id; RelatedData.Add(row);
                Context.Add(new ComponentIdentity(new SolutionComponentRecord(Guid.NewGuid(), 2, column), IdentityResolutionStatus.Resolved, portableColumn)); return row;
            }
            internal MembershipSnapshot Snapshot() => MembershipSnapshot.Complete(Solution(), Raw.Concat(Context), DateTimeOffset.UtcNow);
        }
#endif
    }
}
