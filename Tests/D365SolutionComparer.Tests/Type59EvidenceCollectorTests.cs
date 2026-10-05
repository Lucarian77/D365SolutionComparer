using System;
using System.Collections.Generic;
using System.Linq;
using System.ServiceModel;
using System.Threading;
using D365SolutionComparer.Models.Identity;
using D365SolutionComparer.Models.Membership;
using D365SolutionComparer.Models.ComponentDetails;
using D365SolutionComparer.Services.Membership;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Query;

namespace D365SolutionComparer.Tests
{
    [TestClass]
    public class Type59EvidenceCollectorTests
    {
        private const string Category = "Phase2G15B";

        [TestMethod, TestCategory(Category)]
        public void PublicType59RemainsUnsupportedInBothBuildsAndReleaseExcludesCollectorAndWindow()
        {
            Assert.AreEqual("unsupported:componenttype:59", ComponentSemanticKinds.FromRawComponentType(59));
            Assert.IsNull(ComponentDefinitionContractCatalog.For("unsupported:componenttype:59"));
            var solution = MembershipTestData.Solution();
            var service = MembershipTestData.Service(solution, query => new EntityCollection());
            var identity = new DataverseComponentIdentityResolver().Resolve(service, solution.Environment,
                new SolutionComponentRecord(Guid.NewGuid(), 59, Guid.NewGuid()), CancellationToken.None);
            Assert.AreEqual(IdentityResolutionStatus.Unsupported, identity.Status);
            Assert.IsNull(identity.ComparisonKey);
            Assert.AreEqual(0, service.WriteCalls);
#if !DEBUG
            Assert.IsNull(typeof(DataverseComponentIdentityResolver).Assembly.GetType("D365SolutionComparer.Services.Membership.Type59EvidenceCollector"));
            Assert.IsNull(typeof(DataverseComponentIdentityResolver).Assembly.GetType("D365SolutionComparer.Type59EvidenceResultsForm"));
            Assert.IsNull(typeof(MembershipResultsForm).GetProperty("CaptureType59Evidence",
                System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic));
#endif
        }

#if DEBUG
        [TestMethod, TestCategory(Category)]
        public void ZeroSnapshotInventoryUsesNoRequestsAndNoBackingReads()
        {
            var pair = new Pair(); var report = pair.Capture();
            Assert.AreEqual(0, pair.Source.Service.Calls + pair.Target.Service.Calls);
            Assert.AreEqual(0, pair.Source.Service.ExecuteCalls + pair.Target.Service.ExecuteCalls);
            Assert.AreEqual(0, report.Pairs.Count); Assert.AreEqual(0, report.SelectedPairs.Count);
            StringAssert.Contains(report.Build(), "raw=0");
        }

        [TestMethod, TestCategory(Category)]
        public void ZeroDiscoveryStopsAfterTwoMembershipQueries()
        {
            var pair = new Pair(); var report = pair.Capture(true);
            Assert.AreEqual(1, report.Source.Count("solutioncomponent")); Assert.AreEqual(1, report.Target.Count("solutioncomponent"));
            Assert.AreEqual(0, report.Source.Count(Type59EvidenceCollector.EntityName));
            Assert.AreEqual(0, report.Target.Count(Type59EvidenceCollector.EntityName));
            Assert.AreEqual(0, report.Source.Count("RetrieveMetadataChanges"));
            StringAssert.Contains(report.Build(), "Source-only solution unique names=[]");
        }

        [TestMethod, TestCategory(Category)]
        public void BlankRawObjectIdsAreRetainedAndPreventAbsenceEvidenceWithoutBackingReads()
        {
            var pair = new Pair(); pair.Source.Reference(null); pair.Source.Reference(Guid.Empty); pair.Target.Add("Chart");
            var report = pair.Capture(); Assert.AreEqual(2, report.Source.Raw.Count); Assert.AreEqual(0, report.Source.Charts.Count);
            Assert.AreEqual(0, pair.Source.Service.Calls);
            StringAssert.Contains(report.Build(), "blank=2"); StringAssert.Contains(report.Build(), "Indeterminate: incomplete inventory evidence");
            StringAssert.Contains(report.Build(), "datadescription audit: status=CanonicalXml");
        }

        [TestMethod, TestCategory(Category)]
        public void BroadDiscovery201IdsUsesTwoLightweightBatchesAndCachesRepeatedMembershipAcrossSolutions()
        {
            var pair = new Pair();
            for (int i = 0; i < 201; i++)
            {
                var id = Guid.NewGuid();
                foreach (var side in new[] { pair.Source, pair.Target })
                { side.Add("Chart" + i, id: id); side.Reference(id, "OtherSolution"); }
            }
            var report = pair.Capture(true);
            Assert.AreEqual(201, report.Source.Charts.Count); Assert.AreEqual(402, report.Source.Raw.Count);
            Assert.AreEqual(2, report.Source.Count(Type59EvidenceCollector.EntityName));
            Assert.AreEqual(2, report.Target.Count(Type59EvidenceCollector.EntityName));
            Assert.AreEqual(0, report.SelectedPairs.Count); Assert.AreEqual(0, report.Source.Details.Count);
            Assert.IsTrue(pair.Source.Queries.Where(item => item.EntityName == Type59EvidenceCollector.EntityName)
                .All(item => !item.ColumnSet.Columns.Contains("datadescription") && item.Criteria.Conditions.Single().Values.Count <= 200));
        }

        [TestMethod, TestCategory(Category)]
        public void UnrelatedDuplicateGroupsDoNotBlockUniqueOneSidedCandidateButIncompleteScopeDoes()
        {
            var pair = new Pair(); pair.Source.Add("Unique one-sided");
            foreach (var side in new[] { pair.Source, pair.Target }) { side.Add("Duplicate"); side.Add("Duplicate"); }
            StringAssert.Contains(pair.Capture().Build(), "status=Source-only evidence");
            pair.Target.Add("Missing scope").Attributes.Remove("primaryentitytypecode");
            var text = pair.Capture().Build(); StringAssert.Contains(text, "Indeterminate: incomplete inventory evidence");
            Assert.IsFalse(text.Contains("status=Source-only evidence"));
        }

        [DataTestMethod, TestCategory(Category)]
        [DataRow(200, 1)] [DataRow(201, 2)]
        public void SelectedSolutionDeduplicatesAndBatchesAt200WithExactQueryShape(int count, int requests)
        {
            var pair = new Pair();
            for (int i = 0; i < count; i++) pair.Source.Add("Chart " + i);
            pair.Source.Reference(pair.Source.Rows[0].Id);
            var report = pair.Capture();
            Assert.AreEqual(count + 1, report.Source.Raw.Count); Assert.AreEqual(count, report.Source.Charts.Count);
            Assert.AreEqual(requests, pair.Source.Queries.Count);
            var ids = new List<Guid>();
            foreach (var query in pair.Source.Queries)
            {
                Assert.AreEqual(Type59EvidenceCollector.EntityName, query.EntityName);
                CollectionAssert.AreEquivalent(Type59EvidenceCollector.DetailColumns, query.ColumnSet.Columns.ToArray());
                var condition = query.Criteria.Conditions.Single();
                Assert.AreEqual("savedqueryvisualizationid", condition.AttributeName); Assert.AreEqual(ConditionOperator.In, condition.Operator);
                Assert.IsTrue(condition.Values.Count <= 200); Assert.IsTrue(condition.Values.All(value => value is Guid));
                ids.AddRange(condition.Values.Cast<Guid>());
            }
            Assert.AreEqual(count, ids.Distinct().Count()); CollectionAssert.AreEqual(ids.OrderBy(item => item).ToList(), ids);
            StringAssert.Contains(report.Build(), "repeated raw membership objectid=");
            Assert.IsTrue(report.Source.Charts.Values.All(item => item.StatusA == "Unique"));
        }

        [TestMethod, TestCategory(Category)]
        public void UniqueCorrelationRetainsRawRootEvidenceAndDifferentLocalGuidsMatchCandidates()
        {
            var pair = new Pair(); var left = pair.Source.Add("  Revenue Chart  ", " account ", false);
            var right = pair.Target.Add("revenue chart", "ACCOUNT", true);
            var report = pair.Capture(); var match = report.Pairs.Single();
            Assert.IsTrue(match.UniqueA); Assert.IsTrue(match.DifferentIds); Assert.IsTrue(match.UnmanagedToManaged);
            Assert.AreNotEqual(left.Id, right.Id);
            Assert.AreNotEqual(left["savedqueryvisualizationidunique"], right["savedqueryvisualizationidunique"]);
            Assert.IsTrue(StringComparer.OrdinalIgnoreCase.Equals(match.Source.CandidateB, match.Target.CandidateB));
            StringAssert.Contains(report.Build(), "objectidEqualsPrimaryId=True");
            StringAssert.Contains(report.Build(), "rootsolutioncomponentid=");
            StringAssert.Contains(report.Build(), "rootcomponentbehavior=2; ismetadata=False");
            Assert.AreEqual(0, pair.Source.Service.WriteCalls + pair.Target.Service.WriteCalls);
            Assert.AreEqual(0, pair.Source.Service.ExecuteCalls + pair.Target.Service.ExecuteCalls);
        }

        [DataTestMethod, TestCategory(Category)]
        [DataRow("missing", "Missing")] [DataRow("duplicate", "Duplicate/Ambiguous")]
        [DataRow("conflict", "Incomplete")] [DataRow("entity", "Incomplete")]
        [DataRow("primaryType", "Incomplete")]
        [DataRow("paged", "Incomplete")] [DataRow("null", "Incomplete")]
        public void InvalidBackingCorrelationsNeverProduceCandidateOrFalseAbsence(string scenario, string correlation)
        {
            var pair = new Pair(); var left = pair.Source.Add("Chart"); pair.Target.Add("Chart");
            pair.Source.QueryOverride = query =>
            {
                if (scenario == "missing") return new EntityCollection();
                if (scenario == "null") return null;
                if (scenario == "conflict") left[Type59EvidenceCollector.PrimaryId] = Guid.NewGuid();
                if (scenario == "primaryType") left[Type59EvidenceCollector.PrimaryId] = new EntityReference(Type59EvidenceCollector.EntityName, left.Id);
                if (scenario == "entity") left.LogicalName = "savedquery";
                var result = MembershipTestData.Rows(left);
                if (scenario == "duplicate") result.Entities.Add(left);
                if (scenario == "paged") result.MoreRecords = true;
                return result;
            };
            var report = pair.Capture(); var chart = report.Source.Charts.Single().Value;
            Assert.AreEqual(correlation, chart.Correlation); Assert.IsNull(chart.CandidateA); Assert.IsNull(chart.CandidateB);
            Assert.AreEqual(0, report.Pairs.Count);
            StringAssert.Contains(report.Build(), "Indeterminate: incomplete inventory evidence");
        }

        [DataTestMethod, TestCategory(Category)]
        [DataRow("name")] [DataRow("scope")] [DataRow("numericScope")]
        [DataRow("invalidScope")] [DataRow("noneScope")]
        public void BlankNameOrUnresolvedScopeKeepsCandidatesIncomplete(string scenario)
        {
            var pair = new Pair(); var row = pair.Source.Add("Chart"); pair.Target.Add("Chart");
            if (scenario == "name") row["name"] = " \t ";
            if (scenario == "scope") row.Attributes.Remove("primaryentitytypecode");
            if (scenario == "numericScope") row["primaryentitytypecode"] = 123;
            if (scenario == "invalidScope") row["primaryentitytypecode"] = "not a logical name";
            if (scenario == "noneScope") row["primaryentitytypecode"] = "none";
            var report = pair.Capture();
            Assert.AreEqual("Incomplete", report.Source.Charts.Single().Value.StatusA);
            Assert.AreEqual(0, report.Pairs.Count);
        }

        [TestMethod, TestCategory(Category)]
        public void MissingNumericClassificationMakesOnlyCandidateBIncomplete()
        {
            var pair = new Pair(); var row = pair.Source.Add("Chart"); row.Attributes.Remove("type"); pair.Target.Add("Chart");
            var report = pair.Capture(); Assert.IsTrue(report.Pairs.Single().UniqueA);
            Assert.IsNull(report.Source.Charts.Single().Value.CandidateB);
            StringAssert.Contains(report.Build(), "CandidateBEqual=False");
        }

        [TestMethod, TestCategory(Category)]
        public void SameChartNameAcrossDifferentEntitiesStaysDistinct()
        {
            var pair = new Pair();
            foreach (var side in new[] { pair.Source, pair.Target }) { side.Add("Chart", "account"); side.Add("Chart", "contact"); }
            var report = pair.Capture(); Assert.AreEqual(2, report.Pairs.Count); Assert.IsTrue(report.Pairs.All(item => item.UniqueA));
        }

        [TestMethod, TestCategory(Category)]
        public void DuplicateSameEntityNamesRemainAmbiguousAndDoNotUseGuidsOrXmlToDisambiguate()
        {
            var pair = new Pair();
            foreach (var side in new[] { pair.Source, pair.Target })
            {
                side.Add("Chart", "account"); side.Add(" chart ", "ACCOUNT")["datadescription"] = "<Different/>";
                side.Add("Unique", "account");
            }
            var report = pair.Capture();
            Assert.AreEqual(2, report.Source.Charts.Values.Count(item => item.StatusA == "Ambiguous"));
            Assert.AreEqual(2, report.Source.Charts.Values.Count(item => item.StatusB == "Ambiguous"));
            Assert.AreEqual(1, report.Pairs.Count(item => item.UniqueA));
            Assert.AreEqual(4, report.Pairs.Count(item => !item.UniqueA));
        }

        [DataTestMethod, TestCategory(Category)]
        [DataRow("type")] [DataRow("charttype")]
        public void CandidateBCanDiscriminateACollisionButIsNotApprovedIdentity(string field)
        {
            var pair = new Pair();
            foreach (var side in new[] { pair.Source, pair.Target })
            { side.Add("Chart"); side.Add("Chart")[field] = new OptionSetValue(1); }
            var report = pair.Capture(true);
            Assert.IsTrue(report.Source.Charts.Values.All(item => item.StatusA == "Ambiguous" && item.StatusB == "Unique"));
            Assert.AreEqual(0, report.SelectedPairs.Count);
            StringAssert.Contains(report.Build(), "Unique candidate match (B not approved)");
            Assert.IsTrue(pair.Source.Queries.Where(item => item.EntityName == Type59EvidenceCollector.EntityName)
                .All(item => !item.ColumnSet.Columns.Contains("datadescription")));
        }

        [TestMethod, TestCategory(Category)]
        public void BroadDiscoveryDeduplicatesAcrossSolutionsAndNoDifferingIdsMeansNoXmlReads()
        {
            var pair = new Pair(); var sharedId = Guid.NewGuid();
            foreach (var side in new[] { pair.Source, pair.Target })
            {
                side.Add("Chart", id: sharedId); side.Reference(sharedId, "OtherSolution"); side.Reference(sharedId);
                side.Reference(Guid.NewGuid(), side == pair.Source ? "SourceExclusive" : "TargetExclusive");
            }
            var report = pair.Capture(true);
            Assert.AreEqual(2, report.SharedSolutions.Count); Assert.AreEqual(2, report.Pairs.Count);
            Assert.AreEqual(1, report.Source.Charts.Count); Assert.AreEqual(0, report.SelectedPairs.Count);
            foreach (var side in new[] { pair.Source, pair.Target })
            {
                Assert.AreEqual(2, side.Queries.Count); var query = side.Queries.Single(item => item.EntityName == Type59EvidenceCollector.EntityName);
                CollectionAssert.AreEquivalent(Type59EvidenceCollector.LightColumns, query.ColumnSet.Columns.ToArray());
                Assert.AreEqual(1, query.Criteria.Conditions.Single().Values.Count);
                Assert.AreEqual(0, side.Service.ExecuteCalls);
            }
            StringAssert.Contains(report.Build(), "Source-only solution unique names=[SourceExclusive]");
            StringAssert.Contains(report.Build(), "Target-only solution unique names=[TargetExclusive]");
        }

        [TestMethod, TestCategory(Category)]
        public void DiscoveryWithNoSharedSolutionsNeverReadsBackingRows()
        {
            var pair = new Pair(); pair.Source.Add("Chart", solution: "SourceOnly"); pair.Target.Add("Chart", solution: "TargetOnly");
            var report = pair.Capture(true); Assert.AreEqual(0, report.SharedSolutions.Count);
            Assert.AreEqual(1, pair.Source.Service.Calls); Assert.AreEqual(1, pair.Target.Service.Calls);
        }

        [TestMethod, TestCategory(Category)]
        public void DiscoverySelectsAtMostTwoDistinctDifferingIdPairsPrefersManagedTransportAndBatchesDetailReads()
        {
            var pair = new Pair();
            for (int i = 0; i < 4; i++)
            {
                var source = pair.Source.Add("Chart" + i, managed: false); var target = pair.Target.Add("Chart" + i, managed: i >= 2);
                pair.Source.Reference(source.Id, "OtherSolution"); pair.Target.Reference(target.Id, "OtherSolution");
            }
            var report = pair.Capture(true);
            Assert.AreEqual(2, report.SelectedPairs.Count); Assert.IsTrue(report.SelectedPairs.All(item => item.UnmanagedToManaged));
            Assert.AreEqual(2, report.Source.Details.Count); Assert.AreEqual(2, report.Target.Details.Count);
            foreach (var side in new[] { pair.Source, pair.Target })
            {
                var queries = side.Queries.Where(item => item.EntityName == Type59EvidenceCollector.EntityName).ToList();
                Assert.AreEqual(2, queries.Count); Assert.IsFalse(queries[0].ColumnSet.Columns.Contains("datadescription"));
                Assert.IsTrue(queries[1].ColumnSet.Columns.Contains("datadescription"));
                Assert.AreEqual(2, queries[1].Criteria.Conditions.Single().Values.Count);
            }
            StringAssert.Contains(report.Build(), "selected pairs=2");
        }

        [TestMethod, TestCategory(Category)]
        public void DiscoveryMembershipQueryUsesSupportedFieldsAndRejectsIncompletePaging()
        {
            var pair = new Pair(); pair.Source.Add("Chart"); pair.Target.Add("Chart"); pair.Capture(true);
            var query = pair.Source.Queries.First();
            CollectionAssert.AreEquivalent(Type59EvidenceCollector.MembershipColumns, query.ColumnSet.Columns.ToArray());
            Assert.AreEqual(59, query.Criteria.Conditions.Single().Values.Single());
            Assert.AreEqual("solution", query.LinkEntities.Single().LinkToEntityName);
            CollectionAssert.AreEquivalent(new[] { "uniquename", "version" }, query.LinkEntities.Single().Columns.Columns.ToArray());
            Assert.IsFalse(query.ColumnSet.Columns.Contains("componentidunique")); Assert.IsFalse(query.ColumnSet.Columns.Contains("ismanaged"));
            Assert.IsFalse(query.ColumnSet.Columns.Contains("componentstate"));
            pair.Source.QueryOverride = q => new EntityCollection { MoreRecords = true };
            Assert.ThrowsException<InvalidOperationException>(() => pair.Capture(true));
        }

        [TestMethod, TestCategory(Category)]
        public void DiscoveryMembershipPagingPreservesRowsAndCountsRequests()
        {
            var pair = new Pair(); pair.Source.Add("A"); pair.Source.Add("B"); pair.Target.Add("A");
            pair.Source.QueryOverride = query => query.EntityName != "solutioncomponent" ? pair.Source.DefaultQuery(query) :
                new EntityCollection(new[] { pair.Source.MembershipRow(pair.Source.Raw[query.PageInfo.PageNumber - 1]) })
                { MoreRecords = query.PageInfo.PageNumber == 1, PagingCookie = query.PageInfo.PageNumber == 1 ? "cookie" : null };
            var report = pair.Capture(true); Assert.AreEqual(2, report.Source.Raw.Count);
            Assert.AreEqual(2, report.Source.Count("solutioncomponent"));
        }

        [TestMethod, TestCategory(Category)]
        public void NumericScopesUseOneGroupedMetadataReadAndAreCachedAcrossSolutions()
        {
            var pair = new Pair();
            foreach (var side in new[] { pair.Source, pair.Target })
            {
                var a = side.Add("A"); a["primaryentitytypecode"] = 1;
                var b = side.Add("B"); b["primaryentitytypecode"] = "2";
                side.Reference(a.Id, "OtherSolution"); side.Reference(b.Id, "OtherSolution");
                side.Metadata.Add(MetadataEntity(1, "account"));
                side.Metadata.Add(MetadataEntity(2, "contact"));
            }
            var report = pair.Capture(true);
            Assert.AreEqual(1, report.Source.Count("RetrieveMetadataChanges")); Assert.AreEqual(1, report.Target.Count("RetrieveMetadataChanges"));
            Assert.IsTrue(report.Source.Charts.Values.All(item => item.StatusA == "Unique"));
            Assert.AreEqual(1, pair.Source.Service.ExecuteCalls); Assert.AreEqual(1, pair.Target.Service.ExecuteCalls);
        }

        [DataTestMethod, TestCategory(Category)]
        [DataRow("duplicate")] [DataRow("unexpected")] [DataRow("fault")]
        public void MetadataConflictsAndFaultsLeaveScopeUnresolved(string failure)
        {
            var pair = new Pair(); var row = pair.Source.Add("Chart"); row["primaryentitytypecode"] = 1;
            pair.Source.Metadata.Add(MetadataEntity(failure == "unexpected" ? 2 : 1, "account"));
            if (failure == "duplicate") pair.Source.Metadata.Add(MetadataEntity(1, "contact"));
            if (failure == "fault") pair.Source.ExecuteOverride = request => throw new FaultException("Denied");
            var report = pair.Capture(); Assert.IsNull(report.Source.Charts.Single().Value.CandidateA);
            Assert.AreEqual(1, report.Source.Count("RetrieveMetadataChanges"));
        }

        [TestMethod, TestCategory(Category)]
        public void BackingFaultRetainsConservativeEvidenceWithoutExposingServerDetails()
        {
            var pair = new Pair(); pair.Source.Add("Chart"); pair.Target.Add("Chart");
            pair.Source.QueryOverride = query => throw new FaultException("secret server details");
            var report = pair.Capture(); Assert.AreEqual("Incomplete", report.Source.Charts.Single().Value.Correlation);
            StringAssert.Contains(report.Build(), "Faulted backing retrieval"); Assert.IsFalse(report.Build().Contains("secret server details"));
            Assert.AreEqual(0, report.Pairs.Count);
        }

        [TestMethod, TestCategory(Category)]
        public void InventoryFaultDoesNotReturnCompletedPartialReport()
        {
            var pair = new Pair(); pair.Source.QueryOverride = query => throw new FaultException("Denied");
            Assert.ThrowsException<FaultException>(() => pair.Capture(true));
        }

        [DataTestMethod, TestCategory(Category)]
        [DataRow(false)] [DataRow(true)]
        public void CancellationDuringSnapshotAndDiscoveryReadsPropagates(bool discovery)
        {
            var pair = new Pair(); pair.Source.Add("Chart"); pair.Target.Add("Chart");
            using (var cancellation = new CancellationTokenSource())
            {
                pair.Source.QueryOverride = query => { cancellation.Cancel(); return pair.Source.DefaultQuery(query); };
                Assert.ThrowsException<OperationCanceledException>(() => pair.Capture(discovery, cancellation.Token));
            }
        }

        [TestMethod, TestCategory(Category)]
        public void UnavailableSnapshotIsRejectedBeforeQueries()
        {
            var pair = new Pair();
            var unavailable = MembershipSnapshot.Unavailable(pair.Source.Solution.Environment, "Shared", DateTimeOffset.UtcNow, "Faulted");
            Assert.ThrowsException<ArgumentException>(() => new Type59EvidenceCollector().Capture(pair.Source.Service, unavailable,
                "1", pair.Target.Service, pair.Target.Snapshot(), "1", CancellationToken.None));
            Assert.AreEqual(0, pair.Source.Service.Calls);
        }

        [TestMethod, TestCategory(Category)]
        public void XmlWhitespaceNormalizationPreservesOrderingAndProducesHashesAndBoundedDifference()
        {
            var row = new Entity(); row["xml"] = "<chart>\n <a/>\n <b/>\n</chart>";
            var a = Type59XmlEvidence.Create(row, "xml"); row["xml"] = "<chart><a/><b/></chart>";
            var b = Type59XmlEvidence.Create(row, "xml"); Assert.AreEqual(a.Canonical, b.Canonical); Assert.AreEqual(a.Hash, b.Hash);
            row["xml"] = "<chart><b/><a/></chart>"; var c = Type59XmlEvidence.Create(row, "xml");
            Assert.AreNotEqual(a.Hash, c.Hash); Assert.AreEqual(64, c.Hash.Length);
            var difference = Type59XmlEvidence.Difference(new string('a', 1000), new string('b', 1000));
            Assert.IsTrue(difference.Length < 350); StringAssert.Contains(difference, "First difference offset=0");
            row["xml"] = "<p xml:space='preserve'>  Text  </p>";
            StringAssert.Contains(Type59XmlEvidence.Create(row, "xml").Canonical, "  Text  ");
        }

        [DataTestMethod, TestCategory(Category)]
        [DataRow("<chart>")] [DataRow("<!DOCTYPE chart [<!ENTITY x SYSTEM 'file:///never-read'>]><chart>&x;</chart>")]
        public void MalformedAndExternalEntityXmlRemainUnavailable(string xml)
        {
            var row = new Entity(); row["xml"] = xml;
            var evidence = Type59XmlEvidence.Create(row, "xml"); Assert.IsNull(evidence.Canonical);
            Assert.IsNull(evidence.Hash); Assert.AreEqual("MalformedOrUnsafeXml", evidence.Status);
        }

        [TestMethod, TestCategory(Category)]
        public void DefinitionXmlAndDefaultDifferencesNeverPreventSemanticMatch()
        {
            var pair = new Pair(); var a = pair.Source.Add("Chart"); var b = pair.Target.Add("Chart");
            b["datadescription"] = "<Different/>"; b["presentationdescription"] = "<Different/>"; b["isdefault"] = true;
            var report = pair.Capture(); Assert.IsTrue(report.Pairs.Single().UniqueA);
            StringAssert.Contains(report.Build(), "First difference offset=");
            StringAssert.Contains(report.Build(), "isdefault Source=False; Target=True");
            Assert.AreNotEqual(a.Id, b.Id);
        }

        [TestMethod, TestCategory(Category)]
        public void DetailedFaultOrChangedNameDoesNotProduceUsableXmlPairEvidence()
        {
            foreach (var fault in new[] { true, false })
            {
                var pair = new Pair(); var row = pair.Source.Add("Chart"); pair.Target.Add("Chart");
                pair.Source.QueryOverride = query =>
                {
                    if (query.EntityName == Type59EvidenceCollector.EntityName && query.ColumnSet.Columns.Contains("datadescription"))
                    { if (fault) throw new FaultException("Denied"); row["name"] = "Renamed"; }
                    return pair.Source.DefaultQuery(query);
                };
                var report = pair.Capture(true);
                StringAssert.Contains(report.Build(), "Incomplete or contradictory detailed correlation/identity evidence");
                Assert.IsFalse(report.Build().Contains("First difference offset="));
            }
        }

        [TestMethod, TestCategory(Category)]
        public void EvidenceCollectionDoesNotMutateSnapshotsOrCreateMembershipMatchesAndReportUsesNeutralLabels()
        {
            var pair = new Pair(); pair.Source.Add("Chart"); pair.Target.Add("Chart");
            var source = pair.Source.Snapshot(); var target = pair.Target.Snapshot(); var comparer = new SolutionMembershipComparer();
            var before = comparer.Compare(source, target).Select(item => item.Presence).ToList();
            var report = new Type59EvidenceCollector().Capture(pair.Source.Service, source, "1", pair.Target.Service, target, "2", CancellationToken.None);
            CollectionAssert.AreEqual(before, comparer.Compare(source, target).Select(item => item.Presence).ToList());
            Assert.IsTrue(source.Components.Concat(target.Components).All(item => item.Status == IdentityResolutionStatus.Unsupported && item.ComparisonKey == null));
            Assert.IsTrue(before.All(item => item == MembershipPresence.Indeterminate));
            StringAssert.Contains(report.Build(), "SOURCE / TARGET RECONCILIATION");
            Assert.IsFalse(report.Build().Contains("DEV/UAT")); Assert.IsFalse(report.Build().Contains("DEV VS UAT"));
        }

        private static EntityMetadata MetadataEntity(int code, string name)
        {
            var entity = new EntityMetadata { LogicalName = name };
            typeof(EntityMetadata).GetProperty("ObjectTypeCode").SetValue(entity, code, null);
            return entity;
        }

        private sealed class Pair
        {
            internal readonly Fixture Source = new Fixture("Source managed session"), Target = new Fixture("Target managed session");
            internal Type59EvidenceReport Capture(bool discover = false, CancellationToken token = default(CancellationToken)) =>
                new Type59EvidenceCollector().Capture(Source.Service, Source.Snapshot(), "1", Target.Service, Target.Snapshot(), "2", token, discover);
        }
        private sealed class Fixture
        {
            internal readonly SolutionIdentity Solution;
            internal readonly List<Type59RawEvidence> Raw = new List<Type59RawEvidence>();
            internal readonly List<Entity> Rows = new List<Entity>();
            internal readonly List<EntityMetadata> Metadata = new List<EntityMetadata>();
            internal readonly List<QueryExpression> Queries = new List<QueryExpression>();
            internal readonly FakeOrganizationService Service;
            internal Func<QueryExpression, EntityCollection> QueryOverride;
            internal Func<OrganizationRequest, OrganizationResponse> ExecuteOverride;
            internal Fixture(string environment)
            {
                Solution = new SolutionIdentity(new EnvironmentIdentity(Guid.NewGuid(), environment), Guid.NewGuid(), "Shared");
                Service = new FakeOrganizationService { RetrievePage = query => { Queries.Add(query); return QueryOverride != null ? QueryOverride(query) : DefaultQuery(query); },
                    ExecuteRequest = request =>
                    {
                        if (ExecuteOverride != null) return ExecuteOverride(request);
                        Assert.IsInstanceOfType(request, typeof(RetrieveMetadataChangesRequest));
                        var query = ((RetrieveMetadataChangesRequest)request).Query;
                        CollectionAssert.AreEquivalent(new[] { "ObjectTypeCode", "LogicalName" }, query.Properties.PropertyNames.ToArray());
                        var response = new RetrieveMetadataChangesResponse(); var metadata = new EntityMetadataCollection();
                        metadata.AddRange(Metadata); response.Results["EntityMetadata"] = metadata;
                        return response;
                    } };
            }
            internal Entity Add(string name, string entity = "account", bool managed = false, Guid? id = null, string solution = "Shared")
            {
                var row = new Entity(Type59EvidenceCollector.EntityName, id ?? Guid.NewGuid())
                {
                    [Type59EvidenceCollector.PrimaryId] = id ?? Guid.Empty,
                    ["savedqueryvisualizationidunique"] = Guid.NewGuid(), ["name"] = name, ["primaryentitytypecode"] = entity,
                    ["type"] = new OptionSetValue(0), ["charttype"] = new OptionSetValue(0), ["componentstate"] = new OptionSetValue(0),
                    ["ismanaged"] = managed, ["isdefault"] = false, ["datadescription"] = "<data/>", ["presentationdescription"] = "<presentation/>"
                };
                row[Type59EvidenceCollector.PrimaryId] = row.Id; Rows.Add(row); Reference(row.Id, solution); return row;
            }
            internal void Reference(Guid? id, string solution = "Shared") => Raw.Add(new Type59RawEvidence { Solution = solution, Version = "1.0",
                Record = new SolutionComponentRecord(Guid.NewGuid(), 59, id, 2, Guid.NewGuid(), false) });
            internal MembershipSnapshot Snapshot() => MembershipSnapshot.Complete(Solution, Raw.Where(item => item.Solution == "Shared")
                .Select(item => new ComponentIdentity(item.Record, IdentityResolutionStatus.Unsupported,
                    componentTypeKey: "unsupported:componenttype:59")), DateTimeOffset.UtcNow);
            internal Entity MembershipRow(Type59RawEvidence raw) => new Entity("solutioncomponent", raw.Record.SolutionComponentId)
            {
                ["solutioncomponentid"] = raw.Record.SolutionComponentId, ["componenttype"] = new OptionSetValue(59),
                ["solutionid"] = new EntityReference("solution", Solution.SolutionId), ["objectid"] = raw.Record.ObjectId,
                ["rootcomponentbehavior"] = new OptionSetValue(2), ["rootsolutioncomponentid"] = raw.Record.RootSolutionComponentId,
                ["ismetadata"] = false, ["evidenceSolution.uniquename"] = new AliasedValue("solution", "uniquename", raw.Solution),
                ["evidenceSolution.version"] = new AliasedValue("solution", "version", raw.Version)
            };
            internal EntityCollection DefaultQuery(QueryExpression query)
            {
                if (query.EntityName == "solutioncomponent") return new EntityCollection(Raw.Select(MembershipRow).ToList());
                Assert.AreEqual(Type59EvidenceCollector.EntityName, query.EntityName);
                var ids = query.Criteria.Conditions.Single().Values.Cast<Guid>().ToList();
                return new EntityCollection(Rows.Where(item => ids.Contains(item.Id)).Select(item =>
                {
                    var projected = new Entity(item.LogicalName, item.Id);
                    foreach (var column in query.ColumnSet.Columns) if (item.Attributes.ContainsKey(column)) projected[column] = item[column];
                    return projected;
                }).ToList());
            }
        }
#endif
    }
}
