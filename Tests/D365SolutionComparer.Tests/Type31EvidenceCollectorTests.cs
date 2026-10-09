using System;
using System.Collections.Generic;
using System.Linq;
using System.ServiceModel;
using System.Threading;
using System.Windows.Forms;
using D365SolutionComparer.Models.Identity;
using D365SolutionComparer.Models.Membership;
using D365SolutionComparer.Models.ComponentDetails;
using D365SolutionComparer.Services.Membership;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Metadata.Query;
using Microsoft.Xrm.Sdk.Query;
using static D365SolutionComparer.Tests.MembershipTestData;

namespace D365SolutionComparer.Tests
{
    [TestClass, TestCategory("Phase2GType31Evidence")]
    public class Type31EvidenceCollectorTests
    {
        [TestMethod]
        public void ProductionType31RemainsUnsupportedAndCollectorUiIsExcludedFromRelease()
        {
            Assert.AreEqual("unsupported:componenttype:31", ComponentSemanticKinds.FromRawComponentType(31));
            Assert.IsNull(ComponentDefinitionContractCatalog.For(ComponentSemanticKinds.FromRawComponentType(31)));
            var solution = Solution();
            var id = Guid.NewGuid();
            var service = Service(solution, query => Rows(ProductionReport(id, null)));
            var identity = new DataverseComponentIdentityResolver().Resolve(service, solution.Environment,
                new SolutionComponentRecord(Guid.NewGuid(), 31, id), CancellationToken.None);
            Assert.AreEqual(IdentityResolutionStatus.Unsupported, identity.Status); Assert.IsNull(identity.ComparisonKey);
            Assert.AreEqual(0, service.WriteCalls);
#if !DEBUG
            var assembly = typeof(SolutionComparerControl).Assembly;
            Assert.IsNull(assembly.GetType("D365SolutionComparer.Services.Membership.Type31EvidenceCollector"));
            Assert.IsNull(assembly.GetType("D365SolutionComparer.Type31EvidenceResultsForm"));
            Assert.IsNull(assembly.GetType("D365SolutionComparer.Services.Membership.Type31RelatedScopeReader"));
            Assert.IsNull(typeof(MembershipResultsForm).GetProperty("CaptureType31Evidence",
                System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic));
            Assert.IsFalse(typeof(MembershipCoverageDetailsForm).GetConstructors().SelectMany(c => c.GetParameters()).Any(p => p.Name == "captureType31Evidence"));
            Assert.IsNull(typeof(SolutionComparerControl).GetMethod("CaptureType31Evidence", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic));
#endif
        }

        [TestMethod]
        public void ExistingSignedReportSubsetStillResolvesInBothConfigurations()
        {
            var signature = Guid.NewGuid(); var sourceSolution = Solution(); var targetSolution = Solution();
            var sourceId = Guid.NewGuid(); var targetId = Guid.NewGuid();
            var source = new DataverseComponentIdentityResolver().Resolve(Service(sourceSolution, q => Rows(ProductionReport(sourceId, signature))),
                sourceSolution.Environment, new SolutionComponentRecord(Guid.NewGuid(), 31, sourceId), CancellationToken.None);
            var target = new DataverseComponentIdentityResolver().Resolve(Service(targetSolution, q => Rows(ProductionReport(targetId, signature))),
                targetSolution.Environment, new SolutionComponentRecord(Guid.NewGuid(), 31, targetId), CancellationToken.None);
            Assert.AreEqual(IdentityResolutionStatus.Resolved, source.Status); Assert.AreEqual(ComponentSemanticKinds.Report, source.SemanticKind);
            Assert.AreEqual(signature.ToString("D"), source.ComparisonKey); Assert.AreEqual(source.ComparisonKey, target.ComparisonKey);
            var membership = new SolutionMembershipComparer().Compare(Snapshot(source), Snapshot(target));
            Assert.AreEqual(MembershipPresence.PresentInBoth, membership.Single().Presence);
        }

        private static Entity ProductionReport(Guid id, Guid? signature)
        {
            var row = new Entity("report", id) { ["reportid"] = id, ["name"] = "Report", ["filename"] = "report.rdl",
                ["reporttypecode"] = new OptionSetValue(1), ["reportidunique"] = Guid.NewGuid(),
                ["componentstate"] = new OptionSetValue(0), ["ismanaged"] = false, ["signaturelcid"] = 1033 };
            if (signature.HasValue) row["signatureid"] = signature.Value;
            return row;
        }

        [TestMethod]
        public void ExistingDuplicateSignedSignaturesRemainAmbiguousDespiteDifferentLcids()
        {
            var solution = Solution(); var signature = Guid.NewGuid(); var first = ProductionReport(Guid.NewGuid(), signature);
            var second = ProductionReport(Guid.NewGuid(), signature); second["signaturelcid"] = 1036;
            var snapshot = MembershipSnapshot.Complete(solution, new[] { first, second }.Select(row =>
                new ComponentIdentity(new SolutionComponentRecord(Guid.NewGuid(), 31, row.Id), IdentityResolutionStatus.Unresolved)), DateTimeOffset.UtcNow);
            var resolved = new DataverseComponentIdentityResolver().ResolveSnapshot(Service(solution, q => Rows(first, second)), snapshot, CancellationToken.None);
            Assert.IsTrue(resolved.Components.All(c => c.Status == IdentityResolutionStatus.Ambiguous && c.ComparisonKey == null));
            var compared = new SolutionMembershipComparer().Compare(resolved, MembershipSnapshot.Complete(Solution(), new ComponentIdentity[0], DateTimeOffset.UtcNow));
            Assert.IsTrue(compared.All(r => r.Presence == MembershipPresence.Indeterminate));
        }

#if DEBUG
        private const string Secret = "PRIVATE-BODY-CASE-123-DO-NOT-EXPORT";
        [TestMethod]
        public void RelatedScopeIsDiscoveredFromMetadataAndOnlySelectedReportIdsAreQueried()
        {
            var pair = new Pair();
            foreach (var side in new[] { pair.Source, pair.Target })
            {
                var first = side.Add(title: "First"); var second = side.Add(title: "Second");
                first.Attributes.Remove("relatedentities"); second.Attributes.Remove("relatedentities");
                var scope = new ScopeFixture(side, "ava_reportbindings");
                scope.Add(first.Id, "account"); scope.Add(second.Id, "incident"); scope.Add(Guid.NewGuid(), "unrelated");
            }
            var report = pair.Capture();
            Assert.AreEqual(2, report.Lifecycle.Count(e => e.Pair && e.Outcomes.Contains("SameSemanticCandidate")));
            foreach (var side in new[] { report.Source, report.Target })
            {
                Assert.AreEqual(1, side.ScopeMetadataRequests); Assert.AreEqual(1, side.ScopeQueries);
                Assert.IsTrue(side.Rows.Values.All(r => r.ScopeStatus == "Available" && r.CandidateA != null));
                Assert.AreEqual(1, side.Batches.Count); // Backing Report rows were reused, not read again.
                StringAssert.Contains(report.Build(), "Selected evidence path=ava_reportbindings.boundreport -> objecttypecode");
            }
            StringAssert.Contains(report.Build(), "AdditionalScopeRequests=2");
            Assert.AreEqual(2, pair.Source.Service.Calls); Assert.AreEqual(2, pair.Source.Service.ExecuteCalls);
            Assert.AreEqual(0, pair.Source.Service.WriteCalls + pair.Target.Service.WriteCalls);
        }

        [TestMethod]
        public void NoRelationshipOrNoScopeRowsLeavesCandidatesIncompleteWithoutGuessedTables()
        {
            var pair = new Pair(); var first = pair.Source.Add(); first.Attributes.Remove("relatedentities");
            var target = pair.Target.Add(); target.Attributes.Remove("relatedentities"); new ScopeFixture(pair.Target);
            var report = pair.Capture();
            Assert.AreEqual("Unavailable", report.Source.Rows[first.Id].ScopeStatus);
            Assert.AreEqual(0, report.Source.ScopeMetadataRequests + report.Source.ScopeQueries);
            Assert.AreEqual("Unavailable", report.Target.Rows[target.Id].ScopeStatus);
            Assert.IsTrue(report.Source.Rows.Values.Concat(report.Target.Rows.Values).All(r => r.CandidateA == null && r.CandidateB == null));
            Assert.IsFalse(report.Source.Complete || report.Target.Complete);
        }

        [DataTestMethod]
        [DataRow(false)] [DataRow(true)]
        public void MultipleOrDuplicateScopeAssociationsRemainAmbiguous(bool sameEntity)
        {
            var pair = new Pair(); var first = pair.Source.Add(); first.Attributes.Remove("relatedentities");
            var scope = new ScopeFixture(pair.Source); scope.Add(first.Id, "account"); scope.Add(first.Id, sameEntity ? "account" : "incident");
            var report = pair.Capture(); var evidence = report.Source.Rows[first.Id];
            Assert.AreEqual("Unique", evidence.Status); Assert.AreEqual("Ambiguous", evidence.ScopeStatus);
            Assert.IsNull(evidence.CandidateA); Assert.IsNull(evidence.CandidateB); Assert.IsNull(evidence.RelatedScope);
        }

        [TestMethod]
        public void DuplicateScopeReturnedRowIsUnsafeButIdenticalCrossPageOverlapIsDeduplicated()
        {
            var pair = new Pair(); var first = pair.Source.Add(); first.Attributes.Remove("relatedentities");
            var scope = new ScopeFixture(pair.Source); var association = scope.Add(first.Id, "account");
            scope.Page = q => Rows(association, association);
            var duplicate = pair.Capture(); Assert.AreEqual("Ambiguous", duplicate.Source.Rows[first.Id].ScopeStatus);
            scope.Page = q => { var page = Rows(association); page.MoreRecords = q.PageInfo.PageNumber == 1; page.PagingCookie = "private-cookie"; return page; };
            var paged = pair.Capture(); Assert.AreEqual("Available", paged.Source.Rows[first.Id].ScopeStatus);
            Assert.AreEqual(2, paged.Source.ScopeQueries); Assert.IsFalse(paged.Build().Contains("private-cookie"));
            StringAssert.Contains(paged.Build(), "MoreRecords=False; PagingCookieSupplied=True");
        }

        [TestMethod]
        public void NumericScopeCodesAreResolvedByBatchedMetadataAndNeverUsedAsPortableText()
        {
            var pair = new Pair();
            foreach (var side in new[] { pair.Source, pair.Target })
            {
                var first = side.Add(); first.Attributes.Remove("relatedentities");
                var scope = new ScopeFixture(side); scope.Add(first.Id, side == pair.Source ? 10001 : 20002);
                scope.CodeName = "ava_case";
            }
            var report = pair.Capture(); Assert.IsTrue(report.Lifecycle.Single().Outcomes.Contains("SameSemanticCandidate"));
            Assert.AreEqual("ava_case", report.Source.Rows.Single().Value.RelatedScope);
            Assert.AreEqual(2, report.Source.ScopeMetadataRequests); Assert.AreEqual(1, report.Source.ScopeQueries);
        }

        [TestMethod]
        public void IntersectPathUsesOnlyTheReportLinkPublishedByRelationshipMetadata()
        {
            var pair = new Pair(); var first = pair.Source.Add(); first.Attributes.Remove("relatedentities");
            var scope = new ScopeFixture(pair.Source); scope.Add(first.Id, "account");
            pair.Source.Relationships = new OneToManyRelationshipMetadata[0];
            pair.Source.Intersects = new[] { new ManyToManyRelationshipMetadata { SchemaName = "test_report_entity_intersect",
                Entity1LogicalName = "report", Entity2LogicalName = "entity", IntersectEntityName = scope.Table,
                Entity1IntersectAttribute = "boundreport", Entity2IntersectAttribute = "boundentity" } };
            var report = pair.Capture(); Assert.AreEqual("Available", report.Source.Rows[first.Id].ScopeStatus);
            StringAssert.Contains(report.Build(), "intersect=True");
        }

        [DataTestMethod]
        [DataRow("unreadableScope")] [DataRow("unreadableForeignKey")] [DataRow("wrongLookupTarget")]
        [DataRow("multipleColumns")] [DataRow("missingSchema")] [DataRow("unresolvedCode")]
        public void UnverifiedScopeSchemaOrValuesNeverBuildCandidates(string failure)
        {
            var pair = new Pair(); var first = pair.Source.Add(); first.Attributes.Remove("relatedentities");
            var scope = new ScopeFixture(pair.Source); scope.Add(first.Id, failure == "unresolvedCode" ? (object)999 : "account");
            if (failure == "unreadableScope") Set(scope.Attributes.Single(a => a.LogicalName == "objecttypecode"), "IsValidForRead", false);
            if (failure == "unreadableForeignKey") Set(scope.Attributes.Single(a => a.LogicalName == "boundreport"), "IsValidForRead", false);
            if (failure == "wrongLookupTarget") ((LookupAttributeMetadata)scope.Attributes.Single(a => a.LogicalName == "boundreport")).Targets = new[] { "account" };
            if (failure == "multipleColumns") scope.Attributes.Add(Attribute("entitylogicalname", new EntityNameAttributeMetadata()));
            scope.MissingSchema = failure == "missingSchema";
            var report = pair.Capture(); var row = report.Source.Rows[first.Id];
            Assert.IsNull(row.CandidateA); Assert.IsNull(row.CandidateB); Assert.IsNull(row.RelatedScope);
            Assert.AreNotEqual("Available", row.ScopeStatus);
            if (failure != "unresolvedCode") Assert.AreEqual(0, report.Source.ScopeQueries);
        }

        [DataTestMethod]
        [DataRow("schemaFault")] [DataRow("queryFault")] [DataRow("schemaCancel")]
        [DataRow("queryCancel")] [DataRow("stalledPaging")] [DataRow("foreignReport")]
        public void ScopeFaultsCancellationAndIncompletePagingAreConservative(string failure)
        {
            var pair = new Pair(); var first = pair.Source.Add(); first.Attributes.Remove("relatedentities");
            var scope = new ScopeFixture(pair.Source); var association = scope.Add(first.Id, "account");
            using (var cancellation = new CancellationTokenSource())
            {
                scope.BeforeSchema = () => { if (failure == "schemaFault") throw new FaultException(Secret); if (failure == "schemaCancel") cancellation.Cancel(); };
                scope.Page = q =>
                {
                    if (failure == "queryFault") throw new FaultException(Secret);
                    if (failure == "queryCancel") cancellation.Cancel();
                    if (failure == "foreignReport") association["boundreport"] = new EntityReference("report", Guid.NewGuid());
                    var page = Rows(association); page.MoreRecords = failure == "stalledPaging"; return page;
                };
                if (failure.EndsWith("Cancel", StringComparison.Ordinal))
                    Assert.ThrowsException<OperationCanceledException>(() => pair.Capture(cancellation.Token));
                else
                {
                    var report = pair.Capture(); var row = report.Source.Rows[first.Id];
                    Assert.AreEqual(failure.EndsWith("Fault", StringComparison.Ordinal) ? "Faulted" : "Incomplete", row.ScopeStatus);
                    Assert.IsNull(row.CandidateA); Assert.IsNull(row.RelatedScope); Assert.AreEqual("Unique", row.Status);
                    Assert.IsFalse(report.Build().Contains(Secret));
                }
                Assert.IsTrue(pair.Source.Raw.All(r => r.Status == IdentityResolutionStatus.Unsupported && r.ComparisonKey == null));
                Assert.AreEqual(0, pair.Source.Service.WriteCalls);
            }
        }
        [TestMethod]
        public void CaptureSeparatesExistingSignedSubsetUnsupportedAndUncertainIdentitiesWithoutMutatingThem()
        {
            var pair = new Pair(); var signed = pair.Source.Add(title: "Signed"); var unsigned = pair.Source.Add(title: "Unsigned");
            var signature = Guid.NewGuid(); signed["signatureid"] = signature; signed["signaturelcid"] = 1033;
            pair.Source.Raw[0] = new ComponentIdentity(pair.Source.Raw[0].Record, IdentityResolutionStatus.Resolved,
                signature.ToString("D"), componentTypeKey: ComponentSemanticKinds.Report, diagnostic: "Verified production signature");
            pair.Source.Raw.Add(new ComponentIdentity(new SolutionComponentRecord(Guid.NewGuid(), 31, Guid.NewGuid()),
                IdentityResolutionStatus.Unresolved, componentTypeKey: ComponentSemanticKinds.ReportCandidateTypeKey, diagnostic: "Backing unavailable"));
            var source = pair.Source.Snapshot(); var target = pair.Target.Snapshot();
            var before = new MembershipCoverageCsvExporter().CreateCsv(Present(source, target));
            var report = pair.Capture(); var text = report.Build();
            StringAssert.Contains(text, "VerifiedSignedSubset"); StringAssert.Contains(text, "OutsideVerifiedSignedSubset");
            StringAssert.Contains(text, "ProductionIdentityUnresolved"); StringAssert.Contains(text, signature.ToString("D"));
            Assert.AreEqual(signature.ToString("D"), pair.Source.Raw[0].ComparisonKey);
            Assert.AreEqual(IdentityResolutionStatus.Unsupported, pair.Source.Raw[1].Status);
            Assert.IsNull(pair.Source.Raw[1].ComparisonKey);
            Assert.AreEqual(before, new MembershipCoverageCsvExporter().CreateCsv(Present(source, target)));
            Assert.AreEqual("Missing", report.Source.Rows.Values.Single(r => r.ObjectId != signed.Id && r.ObjectId != unsigned.Id).Status);
        }

        [TestMethod]
        public void SignatureHypothesisIsIndependentAndLcidsCannotDisambiguateCollisions()
        {
            var pair = new Pair(); var first = pair.Source.Add(title: "First"); var second = pair.Source.Add(title: "Second");
            var signature = Guid.NewGuid(); first["signatureid"] = second["signatureid"] = signature;
            first["signaturelcid"] = 1033; second["signaturelcid"] = 1036;
            var target = pair.Target.Add(); target["signatureid"] = signature; target["signaturelcid"] = 1041;
            var report = pair.Capture();
            Assert.IsTrue(report.Source.Rows.Values.All(r => r.DuplicateSignature));
            Assert.IsFalse(report.Target.Rows.Single().Value.DuplicateSignature);
            StringAssert.Contains(report.Build(), "SignatureCandidateStatus=NonuniqueUnsafe");
            StringAssert.Contains(report.Build(), "SignatureIdOverlapAuditOnly\t" + signature);
            Assert.IsTrue(pair.Source.Raw.Concat(pair.Target.Raw).All(r => r.Status == IdentityResolutionStatus.Unsupported && r.ComparisonKey == null));
        }

        [TestMethod]
        public void SameSignatureWithDifferentLocaleIsEqualAuditEvidenceWithoutNewIdentityDecision()
        {
            var pair = new Pair(); var first = pair.Source.Add(); var second = pair.Target.Add(); var signature = Guid.NewGuid();
            first["signatureid"] = second["signatureid"] = signature; first["signaturelcid"] = 1033; second["signaturelcid"] = 1036;
            var report = pair.Capture(); var text = report.Build();
            StringAssert.Contains(text, "signatureid\t" + signature + "\t" + signature + "\tEqualObserved");
            StringAssert.Contains(text, "signaturelcid\t1033\t1036\tDifferentObserved");
            Assert.IsNull(pair.Source.Raw[0].ComparisonKey); Assert.IsNull(pair.Target.Raw[0].ComparisonKey);
        }

        [TestMethod]
        public void DisplayNameWithoutReadableScopeCannotEstablishSemanticCandidate()
        {
            var pair = new Pair(); var first = pair.Source.Add(); first.Attributes.Remove("relatedentities");
            var report = pair.Capture(); Assert.IsNull(report.Source.Rows[first.Id].CandidateA);
            Assert.IsNull(report.Source.Rows[first.Id].CandidateB); Assert.IsNull(pair.Source.Raw[0].ComparisonKey);
        }

        [TestMethod]
        public void EntityScopeUsesCanonicalSetAndSignatureLcidCanBeRecordedWhenLanguageFieldUnavailable()
        {
            var pair = new Pair(); var first = pair.Source.Add(); var second = pair.Target.Add();
            first["relatedentities"] = " account,contact,ACCOUNT "; second["relatedentities"] = "CONTACT,account";
            foreach (var side in new[] { pair.Source, pair.Target })
            {
                Set(side.Attributes.Single(a => a.LogicalName == "languagecode"), "IsValidForRead", false);
                side.Rows.Single()["signaturelcid"] = 1033;
            }
            var report = pair.Capture(); Assert.IsTrue(report.Lifecycle.Single().Outcomes.Contains("SameSemanticCandidate"));
            Assert.IsFalse(report.Source.Columns.Contains("languagecode"));
            Assert.IsNull(pair.Source.Raw[0].ComparisonKey);
        }

        [TestMethod]
        public void ParentAuditExcludesLookupDisplayNamesAndBinaryContentIsNeverRetrieved()
        {
            var pair = new Pair(); var first = pair.Source.Add();
            first["parentreportid"] = new EntityReference("report", Guid.NewGuid()) { Name = Secret };
            pair.Source.Attributes.Add(Attribute("bodybinary", new MemoAttributeMetadata())); first["bodybinary"] = Secret;
            pair.Source.Attributes.Add(Attribute("attachmentbody", new MemoAttributeMetadata())); first["attachmentbody"] = Secret;
            var report = pair.Capture();
            Assert.IsFalse(report.Source.Columns.Contains("bodybinary")); Assert.IsFalse(report.Source.Columns.Contains("attachmentbody"));
            StringAssert.Contains(report.Build(), "parentreportid\treport:"); Assert.IsFalse(report.Build().Contains(Secret));
            Assert.AreEqual(0, pair.Source.Service.WriteCalls); Assert.AreEqual(1, pair.Source.Service.ExecuteCalls);
        }

        [TestMethod]
        public void UniqueBackingRowsCorrelateExactlyAndHashContentWithoutRetainingEntityPayloads()
        {
            var pair = new Pair(); var row = pair.Source.Add(); pair.Target.Add(row.Id);
            var report = pair.Capture(); var evidence = report.Source.Rows.Single().Value;
            Assert.AreEqual("Unique", evidence.Status); Assert.AreEqual(1, evidence.BackingRowCount);
            Assert.AreEqual(row.Id, evidence.ReportId); Assert.AreEqual(row.Id, evidence.ObjectId);
            Assert.IsNotNull(evidence.CandidateA); Assert.IsNotNull(evidence.CandidateB);
            Assert.AreEqual(Secret.Length, evidence.Content["bodytext"].Length);
            Assert.AreEqual(64, evidence.Content["bodytext"].Sha256.Length);
            Assert.AreEqual(1, report.Lifecycle.Count(e => e.Pair));
            Assert.IsFalse(report.Build().Contains(Secret));
            Assert.IsFalse(typeof(Type31ReportEvidence).GetFields(System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic)
                .Any(f => f.FieldType == typeof(Entity)));
            StringAssert.Contains(report.Build(), "Exact ObjectId/reportid/Entity.Id correlation");
        }

        [TestMethod]
        public void MissingBackingRowsRemainEvidenceOnlyAndPreventCompleteInventory()
        {
            var pair = new Pair(); pair.Source.Add(); pair.Source.Rows.Clear();
            var report = pair.Capture();
            Assert.AreEqual("Missing", report.Source.Rows.Single().Value.Status);
            Assert.AreEqual(0, report.Source.Rows.Single().Value.BackingRowCount);
            Assert.IsFalse(report.Source.Complete); Assert.AreEqual(0, report.Lifecycle.Count(e => e.Pair));
            StringAssert.Contains(report.Build(), "not a membership absence finding");
        }

        [TestMethod]
        public void DuplicateReturnedPrimaryKeyRowsNeverEvaluateCandidates()
        {
            var pair = new Pair(); var row = pair.Source.Add(); pair.Source.Rows.Add(row);
            var report = pair.Capture(); var evidence = report.Source.Rows.Single().Value;
            Assert.AreEqual("Duplicate", evidence.Status); Assert.AreEqual(2, evidence.BackingRowCount);
            Assert.IsNull(evidence.CandidateA); Assert.AreEqual(0, report.Lifecycle.Count(e => e.Pair));
            Assert.IsTrue(report.Lifecycle.Single().Outcomes.Contains("Ambiguous"));
        }

        [DataTestMethod]
        [DataRow("missingPrimary")] [DataRow("emptyPrimary")] [DataRow("wrongPrimary")]
        [DataRow("stringPrimary")] [DataRow("wrongEntityId")] [DataRow("wrongEntity")]
        public void BlankOrConflictingPrimaryKeysStayIncomplete(string invalid)
        {
            var pair = new Pair(); var row = pair.Source.Add(); var objectId = row.Id;
            if (invalid == "missingPrimary") row.Attributes.Remove("reportid");
            if (invalid == "emptyPrimary") row["reportid"] = Guid.Empty;
            if (invalid == "wrongPrimary") row["reportid"] = Guid.NewGuid();
            if (invalid == "stringPrimary") row["reportid"] = row.Id.ToString();
            if (invalid == "wrongEntityId") row.Id = Guid.NewGuid();
            if (invalid == "wrongEntity") row.LogicalName = "other";
            pair.Source.Service.RetrievePage = query => Rows(row);
            var report = pair.Capture(); var evidence = report.Source.Rows[objectId];
            Assert.AreEqual("Incomplete", evidence.Status); Assert.IsNull(evidence.CandidateA);
            Assert.AreEqual(0, report.Lifecycle.Count(e => e.Pair));
            if (invalid == "wrongPrimary") StringAssert.Contains(report.Build(), "Incomplete\tFalse");
        }

        [TestMethod]
        public void BlankObjectIdsDoNotCauseSchemaOrBackingRequests()
        {
            var pair = new Pair(); pair.Source.Reference(null); pair.Source.Reference(Guid.Empty);
            var report = pair.Capture();
            Assert.AreEqual(2, report.Source.Raw.Count); Assert.AreEqual(0, report.Source.Rows.Count);
            Assert.AreEqual(0, pair.Source.Service.Calls + pair.Source.Service.ExecuteCalls);
            Assert.AreEqual(2, report.Lifecycle.Count(e => e.Outcomes.Contains("Incomplete")));
            StringAssert.Contains(report.Build(), "Blank\t0\tIncomplete\tNotProven");
        }

        [TestMethod]
        public void ZeroMembersOnBothSidesStopsBeforeAllDataverseReads()
        {
            var pair = new Pair(); var report = pair.Capture();
            Assert.AreEqual(0, pair.Source.Service.Calls + pair.Target.Service.Calls + pair.Source.Service.ExecuteCalls + pair.Target.Service.ExecuteCalls);
            Assert.AreEqual(0, report.Lifecycle.Count); StringAssert.Contains(report.Build(), "SchemaRequests=0\tReportQueries=0");
        }

        [DataTestMethod]
        [DataRow(200, 1)] [DataRow(201, 2)]
        public void QueriesDeduplicateAndUseTypedGuidsWithAtMost200Ids(int count, int requests)
        {
            var pair = new Pair();
            for (int i = 0; i < count; i++) { var row = pair.Source.Add(title: "Report " + i); pair.Source.Reference(row.Id); }
            var report = pair.Capture(); Assert.AreEqual(count, report.Source.Rows.Count);
            Assert.AreEqual(count * 2, report.Source.Raw.Count); Assert.AreEqual(1, pair.Source.Service.ExecuteCalls);
            Assert.AreEqual(requests, pair.Source.Service.Calls);
            foreach (var query in pair.Source.Queries)
            {
                Assert.AreEqual("report", query.EntityName); Assert.IsFalse(query.ColumnSet.AllColumns);
                Assert.AreEqual("reportid", query.Criteria.Conditions.Single().AttributeName);
                Assert.AreEqual(ConditionOperator.In, query.Criteria.Conditions.Single().Operator);
                Assert.IsTrue(query.Criteria.Conditions.Single().Values.All(v => v is Guid));
                Assert.IsTrue(query.Criteria.Conditions.Single().Values.Count <= 200);
                CollectionAssert.AreEqual(report.Source.Columns.ToArray(), query.ColumnSet.Columns.ToArray());
            }
        }

        [TestMethod]
        public void SamePrimaryIdDifferentInstallationIdDoesNotApproveIdentity()
        {
            var pair = new Pair(); var left = pair.Source.Add(); pair.Target.Add(left.Id);
            var report = pair.Capture(); var entry = report.Lifecycle.Single();
            Assert.IsTrue(entry.Outcomes.Contains("SamePrimaryId")); Assert.IsTrue(entry.Outcomes.Contains("DifferentUniqueId"));
            Assert.IsTrue(entry.Outcomes.Contains("SameSemanticCandidate")); Assert.IsTrue(entry.OnlyUniqueIdDifference);
            StringAssert.Contains(report.Build(), "OnlyInstallationSpecificIdDifferenceProven\t1");
            StringAssert.Contains(report.Build(), "do not prove general portability or approve production identity");
        }

        [TestMethod]
        public void DifferingPrimaryIdsWithUniqueCaseInsensitiveSemanticCandidateAreReportedAsEvidence()
        {
            var pair = new Pair(); var left = pair.Source.Add(title: " My Report "); var right = pair.Target.Add(title: "my report");
            right["relatedentities"] = "ACCOUNT";
            var report = pair.Capture(); var entry = report.Lifecycle.Single();
            Assert.IsTrue(entry.Pair); Assert.IsTrue(entry.Outcomes.Contains("DifferentPrimaryId"));
            Assert.IsTrue(entry.Outcomes.Contains("SameSemanticCandidate"));
            StringAssert.Contains(report.Build(), left.Id.ToString()); StringAssert.Contains(report.Build(), right.Id.ToString());
            StringAssert.Contains(report.Build(), "Unique differing-ID semantic observations support further semantic-identity investigation");
        }

        [TestMethod]
        public void ManagedTransitionAndUnmanagedToManagedAreAuditOnly()
        {
            var pair = new Pair(); var left = pair.Source.Add(); pair.Target.Add(left.Id)["ismanaged"] = true;
            var entry = pair.Capture().Lifecycle.Single();
            Assert.IsTrue(entry.Outcomes.Contains("UnmanagedToManaged")); Assert.IsTrue(entry.Outcomes.Contains("ManagedTransition"));
            Assert.IsTrue(entry.Outcomes.Contains("SameSemanticCandidate")); Assert.IsFalse(entry.OnlyUniqueIdDifference);
        }

        [DataTestMethod]
        [DataRow(true)] [DataRow(false)]
        public void SemanticPairContentDifferencesAreIndependentOfPrimaryIdentity(bool samePrimary)
        {
            var pair = new Pair(); var left = pair.Source.Add(); var right = pair.Target.Add(samePrimary ? (Guid?)left.Id : null);
            right["bodytext"] = "Changed private content";
            var report = pair.Capture(); var entry = report.Lifecycle.Single();
            Assert.IsTrue(entry.Outcomes.Contains("DifferentContent")); Assert.IsTrue(entry.Outcomes.Contains("SameSemanticCandidate"));
            Assert.IsTrue(entry.Outcomes.Contains(samePrimary ? "SamePrimaryId" : "DifferentPrimaryId"));
            Assert.IsFalse(report.Build().Contains("Changed private content"));
        }

        [TestMethod]
        public void SamePrimaryIdChangedSemanticCandidateIsNotForcedToSemanticEquality()
        {
            var pair = new Pair(); var left = pair.Source.Add(); pair.Target.Add(left.Id, "Renamed Report");
            var entry = pair.Capture().Lifecycle.Single();
            Assert.IsTrue(entry.Outcomes.Contains("SamePrimaryId")); Assert.IsTrue(entry.Outcomes.Contains("DifferentSemanticCandidate"));
        }

        [TestMethod]
        public void CandidateACollisionCannotBeRepairedByCandidateBOrGuidOverlap()
        {
            var pair = new Pair(); var left = pair.Source.Add(); pair.Source.Add(); pair.Target.Add(left.Id);
            var report = pair.Capture();
            Assert.AreEqual(0, report.Lifecycle.Count(e => e.Pair)); Assert.IsTrue(report.Lifecycle.Single().Outcomes.Contains("Ambiguous"));
            Assert.IsTrue(report.Source.Rows.Values.All(r => r.DuplicateA && r.DuplicateB));
            StringAssert.Contains(report.Build(), "B never repairs A ambiguity");
        }

        [TestMethod]
        public void CandidateBCollisionDoesNotInvalidateTwoUniqueLanguageScopedACandidates()
        {
            var pair = new Pair();
            foreach (var side in new[] { pair.Source, pair.Target }) { side.Add(); side.Add()["languagecode"] = 1036; }
            var report = pair.Capture(); Assert.AreEqual(2, report.Lifecycle.Count(e => e.Pair));
            Assert.IsTrue(report.Source.Rows.Values.All(r => r.DuplicateB && !r.DuplicateA));
        }

        [TestMethod]
        public void ContradictoryPrimaryAndSemanticRelationshipsRemainAmbiguous()
        {
            var pair = new Pair(); var left = pair.Source.Add(); pair.Target.Add(left.Id, "Other report"); pair.Target.Add();
            var report = pair.Capture(); Assert.AreEqual(0, report.Lifecycle.Count(e => e.Pair));
            Assert.IsTrue(report.Lifecycle.Single().Outcomes.Contains("Ambiguous"));
        }

        [TestMethod]
        public void RepeatedMembershipReferencesDoNotInflateBackingPairsOrCollisionGroups()
        {
            var pair = new Pair(); var left = pair.Source.Add(); pair.Source.Reference(left.Id); pair.Target.Add(left.Id); pair.Target.Reference(left.Id);
            var report = pair.Capture(); Assert.AreEqual(2, report.Source.Raw.Count); Assert.AreEqual(1, report.Source.Rows.Count);
            Assert.AreEqual(1, report.Lifecycle.Count(e => e.Pair)); Assert.IsFalse(report.Source.Rows.Values.Any(r => r.DuplicateA));
            StringAssert.Contains(report.Build(), "RepeatedMembershipOnly");
        }

        [TestMethod]
        public void OneSidedCandidateRemainsEvidenceOnlyAndDoesNotChangeMembership()
        {
            var pair = new Pair(); pair.Source.Add(); var source = pair.Source.Snapshot(); var target = pair.Target.Snapshot();
            var before = new SolutionMembershipComparer().Compare(source, target);
            var report = pair.Capture(); Assert.IsTrue(report.Lifecycle.Single().Outcomes.Contains("OneSidedEvidence"));
            var after = new SolutionMembershipComparer().Compare(source, target);
            Assert.AreEqual(1, after.Count); Assert.AreEqual(before.Single().Presence, after.Single().Presence);
            Assert.AreEqual(MembershipPresence.Indeterminate, after.Single().Presence); Assert.AreEqual(MembershipAbsenceEvidence.None, after.Single().AbsenceEvidence);
            Assert.IsNull(source.Components.Single().ComparisonKey);
        }

        [DataTestMethod]
        [DataRow("rdl")] [DataRow("bodytext")] [DataRow("languagecode")] [DataRow("reportidunique")]
        public void UnreadableOrMalformedOptionalSchemaFieldsAreNotRequested(string field)
        {
            var pair = new Pair(); pair.Source.Add();
            Set(pair.Source.Attributes.Single(a => a.LogicalName == field), "IsValidForRead", false);
            var report = pair.Capture(); Assert.IsFalse(report.Source.Columns.Contains(field));
            Assert.IsFalse(pair.Source.Queries.Single().ColumnSet.Columns.Contains(field));
            StringAssert.Contains(report.Build(), "UnavailableOrUnverified"); Assert.IsFalse(report.Build().Contains(Secret));
        }

        [TestMethod]
        public void MalformedOptionalValueNeverPrintsRawPayloadOrClaimsContentEquality()
        {
            var pair = new Pair(); var left = pair.Source.Add(); pair.Target.Add(left.Id);
            left["rdl"] = new byte[] { 1, 2, 3 }; left["languagecode"] = Secret;
            var report = pair.Capture(); Assert.IsNull(report.Source.Rows.Single().Value.CandidateA);
            Assert.IsFalse(report.Lifecycle.Single().Outcomes.Contains("SameContent"));
            Assert.IsTrue(report.Lifecycle.Single().Outcomes.Contains("Incomplete"));
            Assert.IsFalse(report.Build().Contains(Secret)); StringAssert.Contains(report.Build(), "presence=Malformed");
        }

        [TestMethod]
        public void SchemaDiscoveryHashOnlyTextAndUniqueIdsDoesNotRetrieveAttachmentsOrUnrelatedFields()
        {
            var pair = new Pair(); var left = pair.Source.Add();
            pair.Source.Attributes.Add(Attribute("custom_html", new MemoAttributeMetadata()));
            pair.Source.Attributes.Add(Attribute("installationuniqueid", new UniqueIdentifierAttributeMetadata()));
            pair.Source.Attributes.Add(Attribute("attachmentcontent", new MemoAttributeMetadata()));
            pair.Source.Attributes.Add(Attribute("unrelatedflag", new BooleanAttributeMetadata()));
            left["custom_html"] = Secret; left["attachmentcontent"] = Secret; left["installationuniqueid"] = Guid.NewGuid();
            var report = pair.Capture();
            Assert.IsTrue(report.Source.Columns.Contains("custom_html")); Assert.IsTrue(report.Source.Columns.Contains("installationuniqueid"));
            Assert.IsFalse(report.Source.Columns.Contains("attachmentcontent")); Assert.IsFalse(report.Source.Columns.Contains("unrelatedflag"));
            Assert.IsFalse(report.Build().Contains(Secret)); Assert.IsTrue(report.Source.Rows.Single().Value.Content.ContainsKey("custom_html"));
        }

        [DataTestMethod]
        [DataRow("null")] [DataRow("wrongPrimary")] [DataRow("duplicateAttribute")] [DataRow("fault")]
        public void MetadataFailuresNeverGuessColumnsOrExposeServerDetails(string failure)
        {
            var pair = new Pair(); pair.Source.Add();
            pair.Source.Service.ExecuteRequest = request =>
            {
                if (failure == "fault") throw new FaultException(Secret);
                if (failure == "null") return null;
                var response = pair.Source.Schema();
                if (failure == "wrongPrimary") Set(response.EntityMetadata, "PrimaryIdAttribute", "other");
                if (failure == "duplicateAttribute") Set(response.EntityMetadata, "Attributes", response.EntityMetadata.Attributes.Concat(new[] { response.EntityMetadata.Attributes[0] }).ToArray());
                return response;
            };
            var report = pair.Capture(); Assert.AreEqual(0, pair.Source.Service.Calls); Assert.IsFalse(report.Build().Contains(Secret));
            Assert.AreEqual(failure == "fault" ? "Faulted" : "Incomplete", report.Source.Rows.Single().Value.Status);
        }

        [DataTestMethod]
        [DataRow("moreRecords")] [DataRow("null")] [DataRow("fault")]
        public void FaultedAndIncompleteBatchesCannotProduceUniqueCandidates(string failure)
        {
            var pair = new Pair(); var row = pair.Source.Add();
            pair.Source.Service.RetrievePage = query =>
            {
                if (failure == "fault") throw new FaultException(Secret);
                if (failure == "null") return null;
                var result = Rows(row); result.MoreRecords = failure == "moreRecords"; result.PagingCookie = failure == "cookie" ? "cookie" : null; return result;
            };
            var report = pair.Capture(); Assert.IsNull(report.Source.Rows.Single().Value.CandidateA);
            Assert.AreEqual(failure == "fault" ? "Faulted" : "Incomplete", report.Source.Rows.Single().Value.Status);
            Assert.IsFalse(report.Build().Contains(Secret)); Assert.AreEqual(0, report.Lifecycle.Count(e => e.Pair));
        }

        [DataTestMethod]
        [DataRow(false)] [DataRow(true)]
        public void CompleteOnePageBatchAcceptsTerminalCookieAndReportsPagingEvidence(bool cookie)
        {
            var pair = new Pair();
            for (int i = 0; i < 4; i++) pair.Source.Add(title: "Report " + i);
            pair.Source.Service.RetrievePage = query =>
            {
                Assert.AreEqual(1, query.PageInfo.PageNumber); Assert.AreEqual(200, query.PageInfo.Count);
                Assert.IsNull(query.PageInfo.PagingCookie);
                var response = new EntityCollection(pair.Source.Rows); response.PagingCookie = cookie ? Secret : null;
                return response;
            };
            var report = pair.Capture();
            Assert.IsTrue(report.Source.Rows.Values.All(r => r.Status == "Unique" && r.BackingRowCount == 1));
            Assert.IsTrue(report.Source.Batches.Single().Complete); Assert.AreEqual(1, pair.Source.Service.Calls);
            StringAssert.Contains(report.Build(), "RequestedIdCount=4\tRowsReturned=4\tDistinctReturnedIds=4\tPageCount=1\tRetrievalComplete=True");
            StringAssert.Contains(report.Build(), "MoreRecords=False\tPagingCookieSupplied=" + cookie);
            Assert.IsFalse(report.Build().Contains(Secret));
        }

        [DataTestMethod]
        [DataRow(false)] [DataRow(true)]
        public void MultiPageBatchContinuesWithCookieOrSimplePagingUntilTerminalPage(bool cookie)
        {
            var pair = new Pair(); var first = pair.Source.Add(title: "First"); var second = pair.Source.Add(title: "Second");
            pair.Source.Reference(first.Id);
            pair.Source.Service.RetrievePage = query =>
            {
                Assert.AreEqual(2, query.Criteria.Conditions.Single().Values.Count);
                Assert.AreEqual("reportid", query.Orders.Single().AttributeName);
                if (query.PageInfo.PageNumber == 1)
                { var response = Rows(first); response.MoreRecords = true; response.PagingCookie = cookie ? "continuation" : null; return response; }
                Assert.AreEqual(2, query.PageInfo.PageNumber);
                Assert.AreEqual(cookie ? "continuation" : null, query.PageInfo.PagingCookie);
                // Identical page overlap is deduplicated by reportid, not another backing record.
                var terminal = Rows(first, second); terminal.PagingCookie = "terminal"; return terminal;
            };
            var report = pair.Capture(); var batch = report.Source.Batches.Single();
            Assert.AreEqual(2, pair.Source.Service.Calls); Assert.AreEqual(3, batch.ReturnedRows);
            Assert.AreEqual(2, batch.DistinctReturnedIds); Assert.IsTrue(batch.Complete);
            Assert.IsTrue(report.Source.Rows.Values.All(r => r.Status == "Unique" && r.BackingRowCount == 1));
            Assert.AreEqual(3, report.Source.Raw.Count); Assert.AreEqual(2, report.Source.Rows.Count);
            StringAssert.Contains(report.Build(), "PageCount=2\tRetrievalComplete=True");
            StringAssert.Contains(report.Build(), "MoreRecords=True");
        }

        [TestMethod]
        public void MissingRequestedIdIsClassifiedOnlyAfterTerminalPage()
        {
            var pair = new Pair(); var first = pair.Source.Add(title: "First"); var second = pair.Source.Add(title: "Second");
            var missing = pair.Source.Add(title: "Missing");
            pair.Source.Service.RetrievePage = query =>
            {
                if (query.PageInfo.PageNumber == 1) { var page = Rows(first); page.MoreRecords = true; return page; }
                return Rows(second);
            };
            var report = pair.Capture(); Assert.IsTrue(report.Source.Batches.Single().Complete);
            Assert.AreEqual("Missing", report.Source.Rows[missing.Id].Status);
            Assert.AreEqual(0, report.Source.Rows[missing.Id].BackingRowCount);
            Assert.IsFalse(report.Source.Complete); Assert.AreEqual(2, pair.Source.Service.Calls);
        }

        [TestMethod]
        public void ConflictingPayloadOnAnotherPageRemainsDuplicateAndCannotProduceCandidate()
        {
            var pair = new Pair(); var first = pair.Source.Add(); var second = pair.Source.Add(title: "Second");
            var conflict = new Entity("report", first.Id); foreach (var field in first.Attributes) conflict[field.Key] = field.Value;
            conflict["bodytext"] = "Changed private payload";
            pair.Source.Service.RetrievePage = query =>
            {
                if (query.PageInfo.PageNumber == 1) { var page = Rows(first); page.MoreRecords = true; page.PagingCookie = "next"; return page; }
                return Rows(conflict, second);
            };
            var report = pair.Capture(); Assert.IsTrue(report.Source.Batches.Single().Complete);
            Assert.AreEqual("Duplicate", report.Source.Rows[first.Id].Status);
            Assert.IsNull(report.Source.Rows[first.Id].CandidateA); Assert.AreEqual("Unique", report.Source.Rows[second.Id].Status);
            Assert.IsFalse(report.Build().Contains("Changed private payload"));
        }

        [DataTestMethod]
        [DataRow("fault")] [DataRow("null")] [DataRow("noProgress")] [DataRow("repeatedCookie")] [DataRow("conflictingPrimary")]
        public void FailedOrNonAdvancingContinuationNeverClaimsCompletion(string failure)
        {
            var pair = new Pair(); var first = pair.Source.Add(); var second = pair.Source.Add(title: "Second");
            pair.Source.Service.RetrievePage = query =>
            {
                if (query.PageInfo.PageNumber == 1) { var page = Rows(first); page.MoreRecords = true; page.PagingCookie = "next"; return page; }
                if (failure == "fault") throw new FaultException(Secret);
                if (failure == "null") return null;
                if (failure == "conflictingPrimary") { second["reportid"] = Guid.NewGuid(); return Rows(second); }
                var response = Rows(failure == "noProgress" ? first : second);
                response.MoreRecords = true; response.PagingCookie = failure == "repeatedCookie" ? "next" : "other";
                return response;
            };
            var report = pair.Capture(); Assert.IsFalse(report.Source.Batches.Single().Complete);
            Assert.AreEqual(2, pair.Source.Service.Calls);
            Assert.IsTrue(report.Source.Rows.Values.All(r => r.Status == (failure == "fault" ? "Faulted" : "Incomplete") && r.CandidateA == null));
            Assert.IsFalse(report.Build().Contains(Secret));
        }

        [TestMethod]
        public void CancellationOnContinuationPropagatesWithoutCompletingCapture()
        {
            var pair = new Pair(); var row = pair.Source.Add();
            using (var cancellation = new CancellationTokenSource())
            {
                pair.Source.Service.RetrievePage = query =>
                {
                    var page = Rows(row); page.MoreRecords = true;
                    if (query.PageInfo.PageNumber == 2) cancellation.Cancel();
                    return page;
                };
                Assert.ThrowsException<OperationCanceledException>(() => pair.Capture(cancellation.Token));
                Assert.AreEqual(2, pair.Source.Service.Calls); Assert.AreEqual(0, pair.Target.Service.Calls);
            }
        }

        [DataTestMethod]
        [DataRow(false)] [DataRow(true)]
        public void CancellationDuringSchemaOrReportQueryPropagates(bool backing)
        {
            var pair = new Pair(); pair.Source.Add();
            using (var cancellation = new CancellationTokenSource())
            {
                if (backing) pair.Source.Service.RetrievePage = query => { cancellation.Cancel(); return Rows(); };
                else pair.Source.Service.ExecuteRequest = request => { cancellation.Cancel(); return pair.Source.Schema(); };
                Assert.ThrowsException<OperationCanceledException>(() => pair.Capture(cancellation.Token));
            }
        }

        [TestMethod]
        public void CapturingEvidencePreservesCsvCoverageInventoryAndProductionStatus()
        {
            var pair = new Pair(); var left = pair.Source.Add(); pair.Target.Add(left.Id);
            var source = pair.Source.Snapshot(); var target = pair.Target.Snapshot();
            var presentation = Present(source, target); var csv = new MembershipCoverageCsvExporter().CreateCsv(presentation);
            var inventory = UnsupportedCoverageInventory.Build(new[] { new InventoryCheckpoint(source) }, new[] { new InventoryCheckpoint(target) }).Text;
            pair.Capture();
            Assert.AreEqual(csv, new MembershipCoverageCsvExporter().CreateCsv(Present(source, target)));
            Assert.AreEqual(inventory, UnsupportedCoverageInventory.Build(new[] { new InventoryCheckpoint(source) }, new[] { new InventoryCheckpoint(target) }).Text);
            Assert.IsTrue(source.Components.Concat(target.Components).All(c => c.Status == IdentityResolutionStatus.Unsupported && c.ComparisonKey == null));
            Assert.AreEqual(2, presentation.Summary.Unsupported); Assert.AreEqual(0, pair.Source.Service.WriteCalls + pair.Target.Service.WriteCalls);
            Assert.AreEqual(1, pair.Source.Service.ExecuteCalls); Assert.AreEqual(1, pair.Target.Service.ExecuteCalls);
        }

        [DataTestMethod]
        [DataRow(false)] [DataRow(true)]
        public void NormalApplicationCompositionKeepsExistingType31DiagnosticRequestCost(bool signed)
        {
            var solution = Solution(); var component = ComponentRow(solution, 31); var id = component.GetAttributeValue<Guid>("objectid");
            var signature = Guid.NewGuid();
            var service = Service(solution, query =>
            {
                if (query.EntityName == "solution") return Rows(SolutionRow(solution));
                if (query.EntityName == "solutioncomponent") return Rows(component);
                Assert.AreEqual("report", query.EntityName);
                CollectionAssert.AreEquivalent(new[] { "reportid", "name", "filename", "reporttypecode", "signatureid", "signaturelcid", "reportidunique", "componentstate", "ismanaged" }, query.ColumnSet.Columns.ToArray());
                return Rows(ProductionReport(id, signed ? (Guid?)signature : null));
            });
            var result = SolutionComparerControl.CreateMembershipComparisonOperation().ReadAndResolve(service, solution, CancellationToken.None);
            Assert.AreEqual(3, service.Calls); Assert.AreEqual(1, service.ExecuteCalls); Assert.AreEqual(0, service.WriteCalls);
            Assert.AreEqual(signed ? IdentityResolutionStatus.Resolved : IdentityResolutionStatus.Unsupported, result.Membership.Components.Single().Status);
            Assert.AreEqual(signed ? signature.ToString("D") : null, result.Membership.Components.Single().ComparisonKey);
        }

        [TestMethod]
        public void ReportContainsEveryRequiredSectionAndIsDeterministicOnRepeatedBuild()
        {
            var pair = new Pair(); pair.Source.Add(); pair.Target.Add(); var report = pair.Capture(); var text = report.Build();
            foreach (var section in new[] { "RAW TYPE 31 MEMBERSHIP", "BACKING REPORT CORRELATION", "READABLE / UNAVAILABLE SCHEMA", "REPEATED / DUPLICATE ANALYSIS", "CANDIDATE IDENTITY ANALYSIS",
                "SOURCE / TARGET FIELD COMPARISON", "LIFECYCLE CORRELATION MATRIX", "DIFFERING PRIMARY-ID SEMANTIC PAIRS", "PRIMARY-ID PORTABILITY ASSESSMENT", "REQUEST LEDGER" }) StringAssert.Contains(text, section);
            foreach (var category in Type31EvidenceReport.Categories) StringAssert.Contains(text, category + "\t");
            StringAssert.Contains(text, "SchemaRequests=1\tReportQueries=1\tWhoAmI=0\tWrites=0\tMembershipQueries=0");
            Assert.AreEqual(text, report.Build());
        }

        [TestMethod]
        public void DebugButtonRequiresCompletedSnapshotsAndDoesNoRetrievalUntilClicked()
        {
            Exception error = null;
            var thread = new Thread(() =>
            {
                try
                {
                    var snapshot = Snapshot(Enumerable.Range(0, 900).Select(_ => Identity(null, 31, IdentityResolutionStatus.Unsupported)).ToArray());
                    var coverage = new MembershipCoverageDiagnosticsBuilder().Build(snapshot); int calls = 0;
                    using (var form = new MembershipCoverageDetailsForm("Source", coverage, "Target", coverage,
                        Present(snapshot, snapshot), captureType31Evidence: () => calls++))
                    {
                        form.StartPosition = FormStartPosition.Manual; form.Location = new System.Drawing.Point(-4000, -4000); form.Show(); Application.DoEvents();
                        var button = form.Controls.OfType<FlowLayoutPanel>().SelectMany(p => p.Controls.OfType<Button>()).Single(b => b.Text == "Capture Type 31 Report Evidence...");
                        Assert.IsTrue(button.Enabled); Assert.AreEqual(0, calls); button.PerformClick(); Assert.AreEqual(1, calls);
                    }
                    using (var form = new Type31EvidenceResultsForm("redacted evidence"))
                    {
                        var body = form.Controls.OfType<RichTextBox>().Single(); Assert.IsTrue(body.ReadOnly); Assert.IsFalse(body.WordWrap);
                    }
                }
                catch (Exception ex) { error = ex; }
            }) { IsBackground = true };
            thread.SetApartmentState(ApartmentState.STA); thread.Start();
            Assert.IsTrue(thread.Join(TimeSpan.FromSeconds(30)), "Evidence action must not block Coverage Details initialization.");
            if (error != null) System.Runtime.ExceptionServices.ExceptionDispatchInfo.Capture(error).Throw();
        }

        [DataTestMethod]
        [DataRow(false)] [DataRow(true)]
        public void UnsignedRefinementDoesNotTurnBlankSignatureIntoSignedEligibility(bool emptyGuid)
        {
            var pair = new Pair(); var row = pair.Source.Add(); if (emptyGuid) row["signatureid"] = Guid.Empty;
            var report = pair.Capture(); var evidence = report.Source.Rows[row.Id];
            Assert.AreEqual("UnsignedOrUnverifiedRemainder", evidence.ReportSubset); Assert.IsFalse(evidence.ProductionSignedSnapshot);
            Assert.IsNull(evidence.CandidateP); StringAssert.Contains(evidence.ProposedBlockingReason, "Fixed internal uniquename");
            Assert.AreEqual(IdentityResolutionStatus.Unsupported, pair.Source.Raw.Single().Status); Assert.IsNull(pair.Source.Raw.Single().ComparisonKey);
            StringAssert.Contains(report.Build(), "blank signature or this diagnostic");
        }

        [TestMethod]
        public void VerifiedProductionSignedSnapshotIsExcludedFromUnsignedRefinementEvenWithInternalName()
        {
            var pair = new Pair(); var row = pair.Source.Add(); var signature = Guid.NewGuid(); row["signatureid"] = signature; row["uniquename"] = "ava_signed";
            pair.Source.Raw.Clear(); pair.Source.Raw.Add(new ComponentIdentity(new SolutionComponentRecord(Guid.NewGuid(), 31, row.Id),
                IdentityResolutionStatus.Resolved, comparisonKey: signature.ToString("D"), semanticKind: ComponentSemanticKinds.Report));
            var report = pair.Capture(); var evidence = report.Source.Rows[row.Id];
            Assert.AreEqual("VerifiedSignedSubset", evidence.ReportSubset); Assert.IsNull(evidence.CandidateP);
            StringAssert.Contains(evidence.ProposedBlockingReason, "signed subset excluded"); Assert.AreEqual(signature.ToString("D"), pair.Source.Raw.Single().ComparisonKey);
            row.Attributes.Remove("signatureid"); evidence = pair.Capture().Source.Rows[row.Id];
            Assert.AreEqual("UnsignedOrUnverifiedRemainder", evidence.ReportSubset); Assert.IsNull(evidence.CandidateP);
            Assert.AreEqual(IdentityResolutionStatus.Resolved, pair.Source.Raw.Single().Status); // Fresh evidence cannot mutate production snapshot.
        }

        [TestMethod]
        public void NonblankSignatureOutsideVerifiedSnapshotCannotGainSignedEligibilityOrUnsignedFallback()
        {
            var pair = UnsignedReadinessPair(); var row = pair.Source.Rows.Single(); row["signatureid"] = Guid.NewGuid();
            var report = pair.Capture(); var found = report.Source.Rows[row.Id];
            Assert.AreEqual("UnsignedOrUnverifiedRemainder", found.ReportSubset); Assert.IsNull(found.CandidateP);
            StringAssert.Contains(found.ProposedBlockingReason, "Unsigned signature state");
            Assert.AreEqual(IdentityResolutionStatus.Unsupported, pair.Source.Raw.Single().Status);
        }

        [TestMethod]
        public void UnsignedNameAndFilenameCollisionsRemainVisibleWithoutCreatingInternalIdentity()
        {
            var pair = new Pair(); pair.Source.Add(); pair.Source.Add(); pair.Target.Add(); pair.Target.Add(title: "Other name");
            var report = pair.Capture(); Assert.IsTrue(report.Source.Rows.Values.Concat(report.Target.Rows.Values).All(r => r.CandidateP == null));
            StringAssert.Contains(report.Build(), "Source\tNameOnly\tCollisionGroups=1"); StringAssert.Contains(report.Build(), "Source\tFilenameOnly\tCollisionGroups=1");
            StringAssert.Contains(report.Build(), "Target\tNameOnly\tCollisionGroups=0"); StringAssert.Contains(report.Build(), "Target\tFilenameOnly\tCollisionGroups=1");
            StringAssert.Contains(report.Build(), "recommend remaining unsupported");
        }

        [TestMethod]
        public void CompleteFixedInternalIdentifierHypothesisPairsDifferingPrimaryIdsOnlyDiagnostically()
        {
            var pair = UnsignedReadinessPair(); var left = pair.Source.Rows.Single(); var right = pair.Target.Rows.Single();
            right["uniquename"] = " AVA_UNSIGNED_REPORT "; right["filename"] = "different.rdl";
            var report = pair.Capture(); var first = report.Source.Rows[left.Id]; var second = report.Target.Rows[right.Id];
            Assert.AreNotEqual(left.Id, right.Id); Assert.IsTrue(StringComparer.OrdinalIgnoreCase.Equals(first.CandidateP, second.CandidateP)); Assert.IsNotNull(first.CandidateP);
            Assert.IsFalse(first.CandidateP.Contains(left.Id.ToString())); Assert.IsFalse(first.CandidateP.Contains("different.rdl"));
            StringAssert.Contains(report.Build(), "UniqueCandidatePPairs=1\tSamePrimaryId=0\tDifferentPrimaryId=1");
            Assert.IsTrue(pair.Source.Raw.Concat(pair.Target.Raw).All(r => r.Status == IdentityResolutionStatus.Unsupported && r.ComparisonKey == null));
            Assert.AreEqual(0, pair.Source.Service.WriteCalls + pair.Target.Service.WriteCalls);
        }

        [DataTestMethod]
        [DataRow("uniquename", "blank")] [DataRow("uniquename", "unavailable")]
        [DataRow("languagecode", "blank")] [DataRow("languagecode", "unavailable")]
        [DataRow("reporttypecode", "blank")] [DataRow("reporttypecode", "unavailable")]
        public void FixedIdentityDimensionUnavailableCannotUseDisplaySignatureLocaleOrHashFallback(string field, string state)
        {
            var pair = UnsignedReadinessPair(); var row = pair.Source.Rows.Single(); row["schemaname"] = "other_identifier"; row["signaturelcid"] = 1033;
            if (state == "blank") row.Attributes.Remove(field); else Set(pair.Source.Attributes.Single(a => a.LogicalName == field), "IsValidForRead", false);
            var report = pair.Capture(); Assert.IsNull(report.Source.Rows[row.Id].CandidateP); Assert.IsNotNull(report.Source.Rows[row.Id].Content["bodytext"].Sha256);
            Assert.IsFalse(string.IsNullOrWhiteSpace(report.Source.Rows[row.Id].ProposedBlockingReason));
            StringAssert.Contains(report.Build(), "No schemaname/name/filename/signaturelcid fallback");
        }

        [DataTestMethod]
        [DataRow(false)] [DataRow(true)]
        public void MissingOrAmbiguousRelatedScopeKeepsInternalHypothesisIncomplete(bool ambiguous)
        {
            var pair = new Pair(); var row = pair.Source.Add(); row["uniquename"] = "ava_unsigned_report"; row.Attributes.Remove("relatedentities");
            var scope = new ScopeFixture(pair.Source); if (ambiguous) { scope.Add(row.Id, "account"); scope.Add(row.Id, "incident"); }
            var report = pair.Capture(); Assert.IsNull(report.Source.Rows[row.Id].CandidateP);
            Assert.AreEqual(ambiguous ? "Ambiguous" : "Incomplete", report.Source.Rows[row.Id].ProposedScopeStatus);
            StringAssert.Contains(report.Source.Rows[row.Id].ProposedBlockingReason, "Verified entity/table scope");
        }

        [TestMethod]
        public void SamePrimaryIdAndMatchingContentNeverRepairsMissingUnsignedInternalIdentifier()
        {
            var pair = new Pair(); var left = pair.Source.Add(); pair.Target.Add(left.Id);
            var report = pair.Capture(); Assert.IsNull(report.Source.Rows[left.Id].CandidateP);
            StringAssert.Contains(report.Build(), "UniqueCandidatePPairs=0"); StringAssert.Contains(report.Build(), "Same reportid does not prove portability");
        }

        [DataTestMethod]
        [DataRow("bodytext")] [DataRow("signaturedate")] [DataRow("signaturelcid")] [DataRow("ismanaged")]
        public void ContentSignatureAuditAndManagementDifferencesCannotChangeUnsignedInternalCandidate(string field)
        {
            var pair = UnsignedReadinessPair(); var left = pair.Source.Rows.Single(); var right = pair.Target.Rows.Single();
            right[field] = field == "signaturedate" ? (object)new DateTime(2026, 10, 8, 12, 0, 0, DateTimeKind.Utc) :
                field == "signaturelcid" ? (object)1036 : field == "ismanaged" ? (object)true : "PRIVATE CHANGED REPORT BODY";
            var report = pair.Capture(); Assert.AreEqual(report.Source.Rows[left.Id].CandidateP, report.Target.Rows[right.Id].CandidateP);
            if (field == "bodytext") StringAssert.Contains(report.Build(), "Content=DifferentContent");
            if (field == "ismanaged") StringAssert.Contains(report.Build(), "ManagedTransition=True");
            Assert.IsFalse(report.Build().Contains("PRIVATE CHANGED REPORT BODY"));
        }

        [TestMethod]
        public void DuplicateInternalCandidateAndRepeatedMembershipAreDistinguished()
        {
            var pair = UnsignedReadinessPair(); var left = pair.Source.Rows.Single(); pair.Source.Reference(left.Id);
            var report = pair.Capture(); Assert.IsFalse(report.Source.Rows[left.Id].DuplicateP);
            var other = pair.Source.Add(); other["uniquename"] = "ava_unsigned_report"; other.Attributes.Remove("relatedentities");
            pair.Source.Raw.Add(Identity("account", 1)); // This second raw scope is verified independently without another backing read.
            other["relatedentities"] = "account";
            report = pair.Capture(); Assert.IsTrue(report.Source.Rows.Values.All(r => r.DuplicateP));
            StringAssert.Contains(report.Build(), "Source\tCandidateP\tCollisionGroups=1"); Assert.IsFalse(report.Build().Split('\n').Any(l => l.StartsWith("InternalIdentifierPairHypothesis\t")));
        }

        [TestMethod]
        public void BinaryAndBodyPayloadsRemainExcludedOrHashOnlyWithMetadataAnalysis()
        {
            var pair = new Pair(); var row = pair.Source.Add(); pair.Source.Attributes.Add(Attribute("bodybinary", new MemoAttributeMetadata())); row["bodybinary"] = new byte[] { 1, 2, 3 };
            var report = pair.Capture(); Assert.IsTrue(pair.Source.Queries.All(q => !q.ColumnSet.Columns.Contains("bodybinary")));
            StringAssert.Contains(report.Build(), "bodybinary"); StringAssert.Contains(report.Build(), "PrimaryNameAttribute=name");
            Assert.IsFalse(report.Build().Contains(Secret));
            Assert.IsNull(report.Source.Rows[row.Id].CandidateP); Assert.IsFalse(report.Source.Rows[row.Id].RuntimeColumns.Contains("bodybinary"));
        }

        private static Pair UnsignedReadinessPair()
        {
            var pair = new Pair(); foreach (var side in new[] { pair.Source, pair.Target }) {
                var row = side.Add(); row["uniquename"] = "ava_unsigned_report"; row.Attributes.Remove("relatedentities");
                var scope = new ScopeFixture(side); scope.Add(row.Id, "account");
            } return pair;
        }

        private static MembershipComparisonPresentation Present(MembershipSnapshot source, MembershipSnapshot target) => new MembershipResultPresenter().Create(
            MembershipEnvironmentResult.FromSnapshot("Source", source, 4, TimeSpan.Zero), MembershipEnvironmentResult.FromSnapshot("Target", target, 4, TimeSpan.Zero));
        private static void Set(object target, string property, object value) => target.GetType().GetProperty(property).SetValue(target, value, null);
        private static AttributeMetadata Attribute(string name, AttributeMetadata attribute)
        { attribute.LogicalName = name; Set(attribute, "IsValidForRead", true); return attribute; }
        private sealed class Pair
        {
            internal readonly Fixture Source = new Fixture("Source"), Target = new Fixture("Target");
            internal Type31EvidenceReport Capture(CancellationToken token = default(CancellationToken)) => new Type31EvidenceCollector().Capture(
                Source.Service, Source.Snapshot(), "1", Target.Service, Target.Snapshot(), "2", token);
        }

        private sealed class ScopeFixture
        {
            internal readonly string Table;
            internal readonly List<AttributeMetadata> Attributes = new List<AttributeMetadata>();
            internal readonly List<Entity> Rows = new List<Entity>();
            internal Func<QueryExpression, EntityCollection> Page;
            internal Action BeforeSchema;
            internal bool MissingSchema;
            internal string CodeName;
            internal ScopeFixture(Fixture side, string table = "test_reportbindings")
            {
                Table = table;
                side.Relationships = new[] { new OneToManyRelationshipMetadata { SchemaName = "report_to_" + table,
                    ReferencedEntity = "report", ReferencedAttribute = "reportid", ReferencingEntity = table, ReferencingAttribute = "boundreport" } };
                Attributes.Add(Attribute("associationid", new UniqueIdentifierAttributeMetadata()));
                Attributes.Add(Attribute("boundreport", new LookupAttributeMetadata { Targets = new[] { "report" } }));
                Attributes.Add(Attribute("objecttypecode", new EntityNameAttributeMetadata()));
                var execute = side.Service.ExecuteRequest;
                side.Service.ExecuteRequest = request =>
                {
                    if (request is RetrieveEntityRequest) return execute(request);
                    Assert.IsInstanceOfType(request, typeof(RetrieveMetadataChangesRequest)); BeforeSchema?.Invoke();
                    var query = ((RetrieveMetadataChangesRequest)request).Query;
                    var result = new RetrieveMetadataChangesResponse(); var metadata = new EntityMetadataCollection();
                    if (query.Criteria.Conditions[0].PropertyName == "LogicalName")
                    {
                        Assert.IsTrue(query.Criteria.Conditions.All(c => (string)c.Value == Table));
                        if (!MissingSchema)
                        {
                            var entity = new EntityMetadata { LogicalName = Table };
                            Set(entity, "PrimaryIdAttribute", "associationid"); Set(entity, "Attributes", Attributes.ToArray()); metadata.Add(entity);
                        }
                    }
                    else if (CodeName != null)
                    {
                        foreach (var condition in query.Criteria.Conditions)
                        {
                            Assert.AreEqual("ObjectTypeCode", condition.PropertyName);
                            var entity = new EntityMetadata { LogicalName = CodeName }; Set(entity, "ObjectTypeCode", (int)condition.Value); metadata.Add(entity);
                        }
                    }
                    result.Results["EntityMetadata"] = metadata; return result;
                };
                var retrieve = side.Service.RetrievePage;
                side.Service.RetrievePage = query =>
                {
                    if (query.EntityName == "report") return retrieve(query);
                    Assert.AreEqual(Table, query.EntityName);
                    CollectionAssert.AreEquivalent(new[] { "associationid", "boundreport", "objecttypecode" }, query.ColumnSet.Columns.ToArray());
                    Assert.AreEqual("boundreport", query.Criteria.Conditions.Single().AttributeName);
                    var ids = query.Criteria.Conditions.Single().Values.Cast<Guid>().ToArray();
                    Assert.IsTrue(ids.Length <= 200); Assert.IsTrue(ids.All(id => side.Raw.Any(r => r.Record.ObjectId == id)));
                    return Page != null ? Page(query) : new EntityCollection(Rows.Where(r => ids.Contains(((EntityReference)r["boundreport"]).Id)).ToList());
                };
            }
            internal Entity Add(Guid reportId, object scope)
            {
                var row = new Entity(Table, Guid.NewGuid()); row["associationid"] = row.Id;
                row["boundreport"] = new EntityReference("report", reportId); row["objecttypecode"] = scope; Rows.Add(row); return row;
            }
        }
        private sealed class Fixture
        {
            internal readonly SolutionIdentity Solution;
            internal readonly List<ComponentIdentity> Raw = new List<ComponentIdentity>();
            internal readonly List<Entity> Rows = new List<Entity>();
            internal readonly List<QueryExpression> Queries = new List<QueryExpression>();
            internal readonly List<AttributeMetadata> Attributes = new List<AttributeMetadata>();
            internal OneToManyRelationshipMetadata[] Relationships = new OneToManyRelationshipMetadata[0];
            internal ManyToManyRelationshipMetadata[] Intersects = new ManyToManyRelationshipMetadata[0];
            internal readonly FakeOrganizationService Service;
            internal Fixture(string environment)
            {
                Solution = new SolutionIdentity(new EnvironmentIdentity(Guid.NewGuid(), environment), Guid.NewGuid(), "EDU");
                foreach (var field in Type31EvidenceCollector.AuditFields)
                {
                    AttributeMetadata attribute = field == "reportid" || field == "reportidunique" || field == "signatureid" ? (AttributeMetadata)new UniqueIdentifierAttributeMetadata() :
                        new[] { "name", "filename", "categories", "relatedentities", "objecttypecode", "uniquename", "schemaname" }.Contains(field) ? (AttributeMetadata)new StringAttributeMetadata() :
                        field == "signaturedate" ? (AttributeMetadata)new DateTimeAttributeMetadata() :
                        field == "ismanaged" || field == "ispersonal" || field == "iscustomreport" ? (AttributeMetadata)new BooleanAttributeMetadata() :
                        new[] { "ownerid", "parentreportid", "originalreportid", "organizationid", "solutionid", "owninguser", "owningteam" }.Contains(field) ? (AttributeMetadata)new LookupAttributeMetadata() : new IntegerAttributeMetadata();
                    Attributes.Add(Attribute(field, attribute));
                }
                foreach (var field in Type31EvidenceCollector.ContentFields.Where(f => f != "bodybinary")) Attributes.Add(Attribute(field, new MemoAttributeMetadata()));
                Service = new FakeOrganizationService
                {
                    ExecuteRequest = request =>
                    {
                        Assert.IsInstanceOfType(request, typeof(RetrieveEntityRequest));
                        var query = (RetrieveEntityRequest)request; Assert.AreEqual("report", query.LogicalName);
                        Assert.AreEqual(EntityFilters.Attributes | EntityFilters.Relationships, query.EntityFilters); Assert.IsFalse(query.RetrieveAsIfPublished);
                        return Schema();
                    },
                    RetrievePage = query =>
                    {
                        Queries.Add(query); var ids = query.Criteria.Conditions.Single().Values.Cast<Guid>().ToArray();
                        var rows = Rows.Where(r => ids.Contains(r.Id)).Select(r =>
                        {
                            var copy = new Entity(r.LogicalName, r.Id);
                            foreach (var column in query.ColumnSet.Columns) if (r.Contains(column)) copy[column] = r[column];
                            return copy;
                        });
                        return new EntityCollection(rows.ToList());
                    }
                };
            }
            internal RetrieveEntityResponse Schema()
            {
                var metadata = new EntityMetadata { LogicalName = "report" };
                Set(metadata, "PrimaryIdAttribute", "reportid"); Set(metadata, "PrimaryNameAttribute", "name"); Set(metadata, "Attributes", Attributes.ToArray());
                Set(metadata, "OneToManyRelationships", Relationships); Set(metadata, "ManyToManyRelationships", Intersects);
                var result = new RetrieveEntityResponse(); result.Results["EntityMetadata"] = metadata; return result;
            }
            internal Entity Add(Guid? id = null, string title = "Customer follow-up")
            {
                var row = new Entity("report", id ?? Guid.NewGuid()) { ["reportidunique"] = Guid.NewGuid(), ["name"] = title,
                    ["reporttypecode"] = new OptionSetValue(1), ["relatedentities"] = "account", ["filename"] = "report.rdl", ["languagecode"] = 1033, ["ispersonal"] = false, ["statecode"] = 0,
                    ["statuscode"] = 1, ["componentstate"] = 0, ["ismanaged"] = false, ["bodytext"] = Secret,
                    ["rdl"] = "SUBJECT-SUBSTITUTION-SECRET", ["bodyxml"] = "<secret>" + Secret + "</secret>" };
                row["reportid"] = row.Id; Rows.Add(row); Reference(row.Id); return row;
            }
            internal void Reference(Guid? id) => Raw.Add(new ComponentIdentity(new SolutionComponentRecord(Guid.NewGuid(), 31, id),
                IdentityResolutionStatus.Unsupported, diagnostic: "No identity resolver supports this known component type."));
            internal MembershipSnapshot Snapshot() => MembershipSnapshot.Complete(Solution, Raw, new DateTimeOffset(2026, 10, 5, 12, 0, 0, TimeSpan.Zero));
        }
#endif
    }
}
