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
    [TestClass, TestCategory("Phase2GType36Evidence")]
    public class Type36EvidenceCollectorTests
    {
        [TestMethod]
        public void ProductionType36RemainsUnsupportedAndCollectorUiIsExcludedFromRelease()
        {
            Assert.AreEqual("unsupported:componenttype:36", ComponentSemanticKinds.FromRawComponentType(36));
            Assert.IsNull(ComponentDefinitionContractCatalog.For(ComponentSemanticKinds.FromRawComponentType(36)));
            var solution = Solution();
            var service = Service(solution, query => Rows());
            var identity = new DataverseComponentIdentityResolver().Resolve(service, solution.Environment,
                new SolutionComponentRecord(Guid.NewGuid(), 36, Guid.NewGuid()), CancellationToken.None);
            Assert.AreEqual(IdentityResolutionStatus.Unsupported, identity.Status); Assert.IsNull(identity.ComparisonKey);
            Assert.AreEqual(0, service.WriteCalls);
#if !DEBUG
            var assembly = typeof(SolutionComparerControl).Assembly;
            Assert.IsNull(assembly.GetType("D365SolutionComparer.Services.Membership.Type36EvidenceCollector"));
            Assert.IsNull(assembly.GetType("D365SolutionComparer.Type36EvidenceResultsForm"));
            Assert.IsNull(typeof(MembershipResultsForm).GetProperty("CaptureType36Evidence",
                System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic));
            Assert.IsFalse(typeof(MembershipCoverageDetailsForm).GetConstructors().SelectMany(c => c.GetParameters()).Any(p => p.Name == "captureType36Evidence"));
            Assert.IsNull(typeof(SolutionComparerControl).GetMethod("CaptureType36Evidence", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic));
#endif
        }

#if DEBUG
        private const string Secret = "PRIVATE-BODY-CASE-123-DO-NOT-EXPORT";
        [TestMethod]
        public void UniqueBackingRowsCorrelateExactlyAndHashContentWithoutRetainingEntityPayloads()
        {
            var pair = new Pair(); var row = pair.Source.Add(); pair.Target.Add(row.Id);
            var report = pair.Capture(); var evidence = report.Source.Rows.Single().Value;
            Assert.AreEqual("Unique", evidence.Status); Assert.AreEqual(1, evidence.BackingRowCount);
            Assert.AreEqual(row.Id, evidence.TemplateId); Assert.AreEqual(row.Id, evidence.ObjectId);
            Assert.IsNotNull(evidence.CandidateA); Assert.IsNotNull(evidence.CandidateB);
            Assert.AreEqual(Secret.Length, evidence.Content["body"].Length);
            Assert.AreEqual(64, evidence.Content["body"].Sha256.Length);
            Assert.AreEqual(1, report.Lifecycle.Count(e => e.Pair));
            Assert.IsFalse(report.Build().Contains(Secret));
            Assert.IsFalse(typeof(Type36TemplateEvidence).GetFields(System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic)
                .Any(f => f.FieldType == typeof(Entity)));
            StringAssert.Contains(report.Build(), "Exact ObjectId/templateid/Entity.Id correlation");
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
            if (invalid == "missingPrimary") row.Attributes.Remove("templateid");
            if (invalid == "emptyPrimary") row["templateid"] = Guid.Empty;
            if (invalid == "wrongPrimary") row["templateid"] = Guid.NewGuid();
            if (invalid == "stringPrimary") row["templateid"] = row.Id.ToString();
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
            Assert.AreEqual(0, report.Lifecycle.Count); StringAssert.Contains(report.Build(), "SchemaRequests=0\tTemplateQueries=0");
        }

        [DataTestMethod]
        [DataRow(200, 1)] [DataRow(201, 2)]
        public void QueriesDeduplicateAndUseTypedGuidsWithAtMost200Ids(int count, int requests)
        {
            var pair = new Pair();
            for (int i = 0; i < count; i++) { var row = pair.Source.Add(title: "Template " + i); pair.Source.Reference(row.Id); }
            var report = pair.Capture(); Assert.AreEqual(count, report.Source.Rows.Count);
            Assert.AreEqual(count * 2, report.Source.Raw.Count); Assert.AreEqual(2, pair.Source.Service.ExecuteCalls);
            Assert.AreEqual(requests, pair.Source.Service.Calls);
            foreach (var query in pair.Source.Queries)
            {
                Assert.AreEqual("template", query.EntityName); Assert.IsFalse(query.ColumnSet.AllColumns);
                Assert.AreEqual("templateid", query.Criteria.Conditions.Single().AttributeName);
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
            var pair = new Pair(); var left = pair.Source.Add(title: " My Template "); var right = pair.Target.Add(title: "my template");
            right["templatetypecode"] = "ACCOUNT";
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
            right["body"] = "Changed private content";
            var report = pair.Capture(); var entry = report.Lifecycle.Single();
            Assert.IsTrue(entry.Outcomes.Contains("DifferentContent")); Assert.IsTrue(entry.Outcomes.Contains("SameSemanticCandidate"));
            Assert.IsTrue(entry.Outcomes.Contains(samePrimary ? "SamePrimaryId" : "DifferentPrimaryId"));
            Assert.IsFalse(report.Build().Contains("Changed private content"));
        }

        [TestMethod]
        public void SamePrimaryIdChangedSemanticCandidateIsNotForcedToSemanticEquality()
        {
            var pair = new Pair(); var left = pair.Source.Add(); pair.Target.Add(left.Id, "Renamed Template");
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
            var pair = new Pair(); var left = pair.Source.Add(); pair.Target.Add(left.Id, "Other template"); pair.Target.Add();
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
        [DataRow("subject")] [DataRow("body")] [DataRow("languagecode")] [DataRow("templateidunique")]
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
            left["subject"] = new byte[] { 1, 2, 3 }; left["languagecode"] = Secret;
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
            for (int i = 0; i < 4; i++) pair.Source.Add(title: "Template " + i);
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
                Assert.AreEqual("templateid", query.Orders.Single().AttributeName);
                if (query.PageInfo.PageNumber == 1)
                { var response = Rows(first); response.MoreRecords = true; response.PagingCookie = cookie ? "continuation" : null; return response; }
                Assert.AreEqual(2, query.PageInfo.PageNumber);
                Assert.AreEqual(cookie ? "continuation" : null, query.PageInfo.PagingCookie);
                // Identical page overlap is deduplicated by templateid, not another backing record.
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
            var conflict = new Entity("template", first.Id); foreach (var field in first.Attributes) conflict[field.Key] = field.Value;
            conflict["body"] = "Changed private payload";
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
                if (failure == "conflictingPrimary") { second["templateid"] = Guid.NewGuid(); return Rows(second); }
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
        public void CancellationDuringSchemaOrTemplateQueryPropagates(bool backing)
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
            Assert.AreEqual(2, pair.Source.Service.ExecuteCalls); Assert.AreEqual(2, pair.Target.Service.ExecuteCalls);
        }

        [TestMethod]
        public void NormalApplicationCompositionKeepsExistingType36DiagnosticRequestCost()
        {
            var solution = Solution(); var component = ComponentRow(solution, 36); var id = component.GetAttributeValue<Guid>("objectid");
            var service = Service(solution, query =>
            {
                if (query.EntityName == "solution") return Rows(SolutionRow(solution));
                if (query.EntityName == "solutioncomponent") return Rows(component);
                Assert.AreEqual("template", query.EntityName);
                CollectionAssert.AreEquivalent(new[] { "templateid", "title", "templatetypecode", "templateidunique", "ispersonal", "languagecode", "componentstate", "ismanaged" }, query.ColumnSet.Columns.ToArray());
                return Rows(new Entity("template", id) { ["templateid"] = id, ["title"] = "Title", ["templatetypecode"] = "account" });
            });
            var result = SolutionComparerControl.CreateMembershipComparisonOperation().ReadAndResolve(service, solution, CancellationToken.None);
            Assert.AreEqual(3, service.Calls); Assert.AreEqual(1, service.ExecuteCalls); Assert.AreEqual(0, service.WriteCalls);
            Assert.AreEqual(IdentityResolutionStatus.Unsupported, result.Membership.Components.Single().Status);
            Assert.IsNull(result.Membership.Components.Single().ComparisonKey);
        }

        [TestMethod]
        public void ReportContainsEveryRequiredSectionAndIsDeterministicOnRepeatedBuild()
        {
            var pair = new Pair(); pair.Source.Add(); pair.Target.Add(); var report = pair.Capture(); var text = report.Build();
            foreach (var section in new[] { "RAW TYPE 36 MEMBERSHIP", "BACKING TEMPLATE CORRELATION", "REPEATED / DUPLICATE ANALYSIS",
                "SOURCE / TARGET FIELD COMPARISON", "LIFECYCLE CORRELATION MATRIX", "DIFFERING PRIMARY-ID SEMANTIC PAIRS", "PRIMARY-ID PORTABILITY ASSESSMENT", "REQUEST LEDGER" }) StringAssert.Contains(text, section);
            foreach (var category in Type36EvidenceReport.Categories) StringAssert.Contains(text, category + "\t");
            StringAssert.Contains(text, "SchemaRequests=1\tTemplateQueries=1\tWhoAmI=0\tWrites=0\tMembershipQueries=0");
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
                    var snapshot = Snapshot(Enumerable.Range(0, 900).Select(_ => Identity(null, 36, IdentityResolutionStatus.Unsupported)).ToArray());
                    var coverage = new MembershipCoverageDiagnosticsBuilder().Build(snapshot); int calls = 0;
                    using (var form = new MembershipCoverageDetailsForm("Source", coverage, "Target", coverage,
                        Present(snapshot, snapshot), captureType36Evidence: () => calls++))
                    {
                        form.StartPosition = FormStartPosition.Manual; form.Location = new System.Drawing.Point(-4000, -4000); form.Show(); Application.DoEvents();
                        var button = form.Controls.OfType<FlowLayoutPanel>().SelectMany(p => p.Controls.OfType<Button>()).Single(b => b.Text == "Capture Type 36 Email Template Evidence...");
                        Assert.IsTrue(button.Enabled); Assert.AreEqual(0, calls); button.PerformClick(); Assert.AreEqual(1, calls);
                    }
                    using (var form = new Type36EvidenceResultsForm("redacted evidence"))
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
        [DataRow("equal", true)] [DataRow("table", false)] [DataRow("language", false)] [DataRow("title", false)]
        public void FixedProposedCandidateUsesVerifiedTableTitleAndExactLanguage(string difference, bool equal)
        {
            var pair = new Pair(); var left = pair.Source.Add(title: " Fixed Title "); var right = pair.Target.Add(title: "fixed title");
            if (difference == "table") right["templatetypecode"] = "contact";
            if (difference == "language") right["languagecode"] = 1036;
            if (difference == "title") right["title"] = "Other title";
            var report = pair.Capture(); var source = report.Source.Rows[left.Id]; var target = report.Target.Rows[right.Id];
            Assert.AreNotEqual(left.Id, right.Id); Assert.AreEqual("Verified", source.Scope.Status); Assert.IsNotNull(source.CandidateP);
            Assert.AreEqual(equal, StringComparer.OrdinalIgnoreCase.Equals(source.CandidateP, target.CandidateP));
            Assert.AreEqual(equal ? 1 : 0, report.Build().Split(new[] { "\r\n", "\n" }, StringSplitOptions.None).Count(l => l.StartsWith("ScopedPrimaryNamePairHypothesis\t", StringComparison.Ordinal)));
            StringAssert.Contains(report.Build(), "UniqueCandidatePPairs=" + (equal ? "1" : "0"));
            if (equal) StringAssert.Contains(report.Build(), "DifferentPrimaryId=1");
            Assert.IsFalse(source.CandidateP.Contains(left.Id.ToString()));
        }

        [DataTestMethod]
        [DataRow("title", "blank")] [DataRow("title", "unavailable")]
        [DataRow("languagecode", "blank")] [DataRow("languagecode", "unavailable")]
        [DataRow("templatetypecode", "blank")] [DataRow("templatetypecode", "unavailable")]
        public void MissingFixedIdentityDimensionCannotUseFallbackOrContent(string field, string failure)
        {
            var pair = new Pair(); var row = pair.Source.Add(); row["uniquename"] = "stronger_identifier";
            row["schemaname"] = "AnotherIdentifier"; row["objecttypecode"] = "account";
            if (failure == "blank") row.Attributes.Remove(field);
            else Set(pair.Source.Attributes.Single(a => a.LogicalName == field), "IsValidForRead", false);
            var evidence = pair.Capture().Source.Rows[row.Id];
            Assert.IsNull(evidence.CandidateP); Assert.IsNotNull(evidence.Content["body"].Sha256);
            Assert.IsFalse(string.IsNullOrWhiteSpace(evidence.ProposedBlockingReason));
            if (field == "templatetypecode") Assert.IsNull(evidence.Scope); // Existing A may use objecttypecode; P never does.
        }

        [DataTestMethod]
        [DataRow("missing")] [DataRow("ambiguous")] [DataRow("fault")]
        public void NumericScopeRequiresUniquePublishedLogicalTableMetadata(string state)
        {
            var pair = new Pair(); var row = pair.Source.Add(); row["templatetypecode"] = "10000";
            var normal = pair.Source.Service.ExecuteRequest;
            pair.Source.Service.ExecuteRequest = request =>
            {
                if (!(request is RetrieveMetadataChangesRequest)) return normal(request);
                var query = ((RetrieveMetadataChangesRequest)request).Query;
                Assert.AreEqual("ObjectTypeCode", query.Criteria.Conditions.Single().PropertyName);
                Assert.AreEqual(10000, query.Criteria.Conditions.Single().Value);
                if (state == "fault") throw new FaultException(Secret);
                var entities = new EntityMetadataCollection();
                if (state == "ambiguous") for (int i = 0; i < 2; i++) { var entity = new EntityMetadata { LogicalName = "custom_table", MetadataId = Guid.NewGuid() }; Set(entity, "ObjectTypeCode", 10000); entities.Add(entity); }
                var response = new RetrieveMetadataChangesResponse(); response.Results["EntityMetadata"] = entities; return response;
            };
            var report = pair.Capture(); var evidence = report.Source.Rows[row.Id];
            Assert.AreEqual("Unique", evidence.Status); Assert.IsNotNull(evidence.CandidateA); Assert.IsNull(evidence.CandidateP);
            Assert.AreEqual(state == "ambiguous" ? "Ambiguous" : state == "fault" ? "Faulted" : "Incomplete", evidence.Scope.Status);
            Assert.IsFalse(report.Build().Contains(Secret));
        }

        [TestMethod]
        public void NumericScopeMapsToLogicalNameAndNeverBecomesPortableNumericKey()
        {
            var pair = new Pair(); var left = pair.Source.Add(); var right = pair.Target.Add(); left["templatetypecode"] = "10000"; right["templatetypecode"] = "10001";
            foreach (var side in new[] { pair.Source, pair.Target })
            {
                var normal = side.Service.ExecuteRequest;
                side.Service.ExecuteRequest = request => {
                    if (!(request is RetrieveMetadataChangesRequest)) return normal(request);
                    var entity = new EntityMetadata { LogicalName = "custom_table", MetadataId = Guid.NewGuid() };
                    Set(entity, "ObjectTypeCode", (int)((RetrieveMetadataChangesRequest)request).Query.Criteria.Conditions.Single().Value);
                    var response = new RetrieveMetadataChangesResponse(); response.Results["EntityMetadata"] = new EntityMetadataCollection { entity }; return response;
                };
            }
            var report = pair.Capture(); Assert.AreEqual(report.Source.Rows[left.Id].CandidateP, report.Target.Rows[right.Id].CandidateP);
            Assert.AreNotEqual(report.Source.Rows[left.Id].CandidateA, report.Target.Rows[right.Id].CandidateA);
            StringAssert.Contains(report.Source.Rows[left.Id].CandidateP, "custom_table"); Assert.IsFalse(report.Source.Rows[left.Id].CandidateP.Contains("10000"));
        }

        [TestMethod]
        public void CollisionMatrixRetainsLanguageAndDoesNotCountRepeatedMembershipAsBackingCollision()
        {
            var pair = new Pair(); var first = pair.Source.Add(); pair.Source.Reference(first.Id);
            pair.Source.Add()["languagecode"] = 1036;
            var report = pair.Capture(); Assert.IsTrue(report.Source.Rows.Values.All(r => !r.DuplicateP));
            StringAssert.Contains(report.Build(), "Source\tTitle\tCollisionGroups=1");
            StringAssert.Contains(report.Build(), "Source\tTitle+VerifiedTable\tCollisionGroups=1");
            StringAssert.Contains(report.Build(), "Source\tTitle+VerifiedTable+Language\tCollisionGroups=0");
            pair.Source.Add(); report = pair.Capture(); Assert.AreEqual(2, report.Source.Rows.Values.Count(r => r.DuplicateP));
            StringAssert.Contains(report.Build(), "CandidatePCollisionGroups=1");
            Assert.IsFalse(report.Build().Split('\n').Any(l => l.StartsWith("ScopedPrimaryNamePairHypothesis\t")));
        }

        [DataTestMethod]
        [DataRow("body")] [DataRow("subject")] [DataRow("description")] [DataRow("generationtypecode")]
        public void DefinitionEvidenceDifferenceDoesNotManufactureOrChangeProposedIdentity(string field)
        {
            var pair = new Pair(); var left = pair.Source.Add(); var right = pair.Target.Add(); right["ismanaged"] = true;
            right[field] = field == "generationtypecode" ? (object)2 : "Private changed content";
            var report = pair.Capture(); Assert.AreEqual(report.Source.Rows[left.Id].CandidateP, report.Target.Rows[right.Id].CandidateP);
            StringAssert.Contains(report.Build(), "ManagedTransition=True"); StringAssert.Contains(report.Build(), "DifferentPrimaryId=1");
            if (field != "generationtypecode") StringAssert.Contains(report.Build(), "DifferentDefinitionContent=1");
            else StringAssert.Contains(report.Build(), "Generation/template kind and personal scope remain Unresolved");
            Assert.IsFalse(report.Build().Contains("Private changed content"));
        }

        [TestMethod]
        public void SamePrimaryIdWithDifferentTitleCannotCreateSemanticPairOrProvePortability()
        {
            var pair = new Pair(); var row = pair.Source.Add(); pair.Target.Add(row.Id, "Different title");
            var report = pair.Capture(); Assert.AreEqual(0, report.Lifecycle.Count(e => e.Pair));
            StringAssert.Contains(report.Build(), "UniqueCandidatePPairs=0"); StringAssert.Contains(report.Build(), "Same templateid does not prove portability");
        }

        [DataTestMethod]
        [DataRow("body")] [DataRow("generationtypecode")]
        public void FullContentQueryFaultRetriesSelectedIdentityColumnsWithoutPretendingHashesAreBlank(string faultingField)
        {
            var pair = new Pair(); var row = pair.Source.Add(); var normal = pair.Source.Service.RetrievePage;
            pair.Source.Service.RetrievePage = query => {
                Assert.IsTrue(query.Criteria.Conditions.Single().Values.Cast<Guid>().SequenceEqual(new[] { row.Id }));
                if (query.ColumnSet.Columns.Contains(faultingField)) throw new FaultException(Secret);
                return normal(query);
            };
            var report = pair.Capture(); var evidence = report.Source.Rows[row.Id];
            Assert.AreEqual("Unique", evidence.Status); Assert.IsNotNull(evidence.CandidateA); Assert.IsNotNull(evidence.CandidateP);
            Assert.AreEqual(2, pair.Source.Service.Calls); Assert.IsFalse(evidence.Content.ContainsKey("body"));
            Assert.IsFalse(evidence.RuntimeColumns.Contains("generationtypecode"));
            Assert.IsTrue(report.Source.Batches.Last().Complete); Assert.IsFalse(report.Source.Batches.First().Complete);
            StringAssert.Contains(report.Build(), "Optional content unavailable"); Assert.IsFalse(report.Build().Contains(Secret));
        }

        [TestMethod]
        public void CompletedTableSnapshotIsReusedWithoutScopeMetadataRequest()
        {
            var pair = new Pair(); var row = pair.Source.Add(); pair.Source.Raw.Add(Identity("account", 1));
            var report = pair.Capture(); Assert.IsNotNull(report.Source.Rows[row.Id].CandidateP);
            Assert.AreEqual(1, pair.Source.Service.ExecuteCalls); Assert.AreEqual(1, pair.Source.Service.Calls);
            StringAssert.Contains(report.Build(), "Completed snapshot Table correlation");
        }

        [TestMethod]
        public void ScopeMetadataCancellationPropagatesAndDoesNotStartTargetCapture()
        {
            var pair = new Pair(); pair.Source.Add(); using (var cancel = new CancellationTokenSource()) {
                var normal = pair.Source.Service.ExecuteRequest;
                pair.Source.Service.ExecuteRequest = request => { if (request is RetrieveMetadataChangesRequest) cancel.Cancel(); return normal(request); };
                Assert.ThrowsException<OperationCanceledException>(() => pair.Capture(cancel.Token));
                Assert.AreEqual(0, pair.Target.Service.ExecuteCalls + pair.Target.Service.Calls);
            }
        }

        private static MembershipComparisonPresentation Present(MembershipSnapshot source, MembershipSnapshot target) => new MembershipResultPresenter().Create(
            MembershipEnvironmentResult.FromSnapshot("Source", source, 4, TimeSpan.Zero), MembershipEnvironmentResult.FromSnapshot("Target", target, 4, TimeSpan.Zero));
        private static void Set(object target, string property, object value) => target.GetType().GetProperty(property).SetValue(target, value, null);
        private static AttributeMetadata Attribute(string name, AttributeMetadata attribute)
        { attribute.LogicalName = name; Set(attribute, "IsValidForRead", true); return attribute; }
        private sealed class Pair
        {
            internal readonly Fixture Source = new Fixture("Source"), Target = new Fixture("Target");
            internal Type36EvidenceReport Capture(CancellationToken token = default(CancellationToken)) => new Type36EvidenceCollector().Capture(
                Source.Service, Source.Snapshot(), "1", Target.Service, Target.Snapshot(), "2", token);
        }
        private sealed class Fixture
        {
            internal readonly SolutionIdentity Solution;
            internal readonly List<ComponentIdentity> Raw = new List<ComponentIdentity>();
            internal readonly List<Entity> Rows = new List<Entity>();
            internal readonly List<QueryExpression> Queries = new List<QueryExpression>();
            internal readonly List<AttributeMetadata> Attributes = new List<AttributeMetadata>();
            internal readonly FakeOrganizationService Service;
            internal Fixture(string environment)
            {
                Solution = new SolutionIdentity(new EnvironmentIdentity(Guid.NewGuid(), environment), Guid.NewGuid(), "EDU");
                foreach (var field in Type36EvidenceCollector.AuditFields)
                {
                    AttributeMetadata attribute = field == "templateid" || field == "templateidunique" ? (AttributeMetadata)new UniqueIdentifierAttributeMetadata() :
                        field == "title" || field == "templatetypecode" || field == "objecttypecode" || field == "uniquename" || field == "schemaname" ? (AttributeMetadata)new StringAttributeMetadata() :
                        field == "ismanaged" || field == "ispersonal" ? (AttributeMetadata)new BooleanAttributeMetadata() :
                        new[] { "ownerid", "organizationid", "solutionid", "owninguser", "owningteam" }.Contains(field) ? (AttributeMetadata)new LookupAttributeMetadata() : new IntegerAttributeMetadata();
                    Attributes.Add(Attribute(field, attribute));
                }
                foreach (var field in Type36EvidenceCollector.ContentFields) Attributes.Add(Attribute(field, new MemoAttributeMetadata()));
                Service = new FakeOrganizationService
                {
                    ExecuteRequest = request =>
                    {
                        if (request is RetrieveMetadataChangesRequest)
                        {
                            var scoped = ((RetrieveMetadataChangesRequest)request).Query;
                            CollectionAssert.AreEquivalent(new[] { "MetadataId", "LogicalName", "ObjectTypeCode" }, scoped.Properties.PropertyNames.ToArray());
                            Assert.AreEqual(LogicalOperator.Or, scoped.Criteria.FilterOperator);
                            Assert.IsTrue(scoped.Criteria.Conditions.Count <= 200);
                            var entities = new EntityMetadataCollection();
                            foreach (var condition in scoped.Criteria.Conditions)
                            {
                                Assert.AreEqual("LogicalName", condition.PropertyName);
                                entities.Add(new EntityMetadata { MetadataId = Guid.NewGuid(), LogicalName = ((string)condition.Value).ToLowerInvariant() });
                            }
                            var response = new RetrieveMetadataChangesResponse(); response.Results["EntityMetadata"] = entities; return response;
                        }
                        Assert.IsInstanceOfType(request, typeof(RetrieveEntityRequest));
                        var query = (RetrieveEntityRequest)request; Assert.AreEqual("template", query.LogicalName);
                        Assert.AreEqual(EntityFilters.Attributes, query.EntityFilters); Assert.IsFalse(query.RetrieveAsIfPublished);
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
                var metadata = new EntityMetadata { LogicalName = "template" };
                Set(metadata, "PrimaryIdAttribute", "templateid"); Set(metadata, "PrimaryNameAttribute", "title"); Set(metadata, "Attributes", Attributes.ToArray());
                var result = new RetrieveEntityResponse(); result.Results["EntityMetadata"] = metadata; return result;
            }
            internal Entity Add(Guid? id = null, string title = "Customer follow-up")
            {
                var row = new Entity("template", id ?? Guid.NewGuid()) { ["templateidunique"] = Guid.NewGuid(), ["title"] = title,
                    ["templatetypecode"] = "account", ["languagecode"] = 1033, ["ispersonal"] = false, ["statecode"] = 0,
                    ["statuscode"] = 1, ["componentstate"] = 0, ["ismanaged"] = false, ["body"] = Secret,
                    ["subject"] = "SUBJECT-SUBSTITUTION-SECRET", ["presentationxml"] = "<secret>" + Secret + "</secret>" };
                row["templateid"] = row.Id; Rows.Add(row); Reference(row.Id); return row;
            }
            internal void Reference(Guid? id) => Raw.Add(new ComponentIdentity(new SolutionComponentRecord(Guid.NewGuid(), 36, id),
                IdentityResolutionStatus.Unsupported, diagnostic: "No identity resolver supports this known component type."));
            internal MembershipSnapshot Snapshot() => MembershipSnapshot.Complete(Solution, Raw, new DateTimeOffset(2026, 10, 5, 12, 0, 0, TimeSpan.Zero));
        }
#endif
    }
}
