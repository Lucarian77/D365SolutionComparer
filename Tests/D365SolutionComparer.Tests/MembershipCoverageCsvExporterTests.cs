using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using D365SolutionComparer.Models.Identity;
using D365SolutionComparer.Models.Membership;
using D365SolutionComparer.Services.Membership;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace D365SolutionComparer.Tests
{
    [TestClass]
    public class MembershipCoverageCsvExporterTests
    {
        private static readonly Guid SourceEnvironmentId = Guid.Parse("10000000-0000-0000-0000-000000000001");
        private static readonly Guid TargetEnvironmentId = Guid.Parse("20000000-0000-0000-0000-000000000002");
        private static readonly Guid SolutionId = Guid.Parse("30000000-0000-0000-0000-000000000003");
        private static readonly Guid SolutionComponentId = Guid.Parse("40000000-0000-0000-0000-000000000004");
        private static readonly Guid ObjectId = Guid.Parse("50000000-0000-0000-0000-000000000005");
        private static readonly Guid CanvasAppId = Guid.Parse("60000000-0000-0000-0000-000000000006");
        private static readonly DateTimeOffset CapturedAt = new DateTimeOffset(2026, 9, 9, 14, 30, 0, TimeSpan.Zero);

        [TestMethod]
        public void CanvasAppComponentExportsLifecycleFieldsFromExistingEvidence()
        {
            var presentation = CreatePresentation(CanvasIdentity());

            var rows = Parse(new MembershipCoverageCsvExporter().CreateCsv(presentation, "1.2.3.4", "1.2.3.3",
                "Unmanaged DEV to managed UAT deployment"));
            var row = rows.Single(item => item["RowType"] == "Component" && item["Side"] == "Source");

            Assert.AreEqual(ObjectId.ToString("D"), row["ObjectId"]);
            Assert.AreEqual("300", row["RawComponentType"]);
            Assert.AreEqual("new_edu_canvas", row["CanvasAppName"]);
            Assert.AreEqual(CanvasAppId.ToString("D"), row["CanvasAppId"]);
            Assert.AreEqual("shared-canvas-token", row["UniqueCanvasAppId"]);
            Assert.AreEqual("EDU, Canvas", row["CanvasAppDisplayName"]);
            Assert.AreEqual("0 ('Published')", row["CanvasAppComponentState"]);
            Assert.AreEqual("True", row["CanvasAppIsManaged"]);
            Assert.AreEqual("CandidateValid", row["CanvasAppCandidateStatus"]);
            Assert.AreEqual("Unmanaged DEV to managed UAT deployment", row["LifecycleOperation"]);
            StringAssert.StartsWith(row["CanvasAppCandidateDiagnostic"],
                "Canvas App lifecycle candidate status=CandidateValid;");
            Assert.AreEqual("CSC-ICMS-DEV", row["Environment"]);
            Assert.AreEqual(SourceEnvironmentId.ToString("D"), row["EnvironmentOrganizationId"]);
            Assert.AreEqual("EDU", row["SolutionUniqueName"]);
            Assert.AreEqual(SolutionId.ToString("D"), row["SolutionId"]);
            Assert.AreEqual("1.2.3.4", row["SolutionVersion"]);
            Assert.AreEqual(CapturedAt.UtcDateTime.ToString("o"), row["CaptureTimestampUtc"]);
            Assert.AreEqual(IdentityResolutionStatus.Unsupported.ToString(), row["ResolutionStatus"]);
            Assert.AreEqual("unsupported:componenttype:300", row["SemanticKind"]);
            Assert.AreEqual(string.Empty, row["ComparisonKey"]);
        }

        [TestMethod]
        public void ExportIncludesOperationComparisonAndCoverageSummaries()
        {
            var presentation = CreatePresentation(CanvasIdentity());
            var rows = Parse(new MembershipCoverageCsvExporter().CreateCsv(presentation));

            var comparison = rows.Single(item => item["RowType"] == "ComparisonSummary");
            Assert.AreEqual("0", comparison["ComparisonPresentInBoth"]);
            Assert.AreEqual("0", comparison["ComparisonSourceOnly"]);
            Assert.AreEqual("0", comparison["ComparisonTargetOnly"]);
            Assert.AreEqual("1", comparison["ComparisonUnsupported"]);

            var operation = rows.Single(item => item["RowType"] == "OperationSummary" && item["Side"] == "Source");
            Assert.AreEqual("17", operation["RequestCount"]);
            Assert.AreEqual("1", operation["RawMembershipCount"]);
            Assert.AreEqual("0", operation["ResolvedCount"]);
            Assert.AreEqual("1", operation["UnsupportedCount"]);
            StringAssert.Contains(operation["DiagnosticEvidence"],
                "Canvas App diagnostic summary: RawType300MembershipCount=1");

            var coverage = rows.Single(item => item["RowType"] == "CoverageBucket" &&
                item["Side"] == "Source" && item["SemanticKind"] == "unsupported:componenttype:300");
            Assert.AreEqual("Isolated unsupported component type 300", coverage["CoverageDisplayName"]);
            Assert.AreEqual("KnownUnsupportedIsolatedType", coverage["CoverageBucketType"]);
            Assert.AreEqual("Incomplete", coverage["CoverageStatus"]);
            Assert.AreEqual("1", coverage["CoverageTotal"]);
            StringAssert.Contains(coverage["CoverageDiagnosticGroups"], "Unsupported x1:");
        }

        [TestMethod]
        public void LifecycleCsvSchemaIsStableAndContainsRequiredMatrixColumns()
        {
            var header = ParseRecords(new MembershipCoverageCsvExporter().CreateCsv(
                CreatePresentation(CanvasIdentity()), "1.2.3.4", "1.2.3.3", "Canvas App rename"))[0];
            CollectionAssert.AreEqual(new[]
            {
                "RowType", "LifecycleOperation", "Side", "Environment", "EnvironmentOrganizationId",
                "SolutionUniqueName", "SolutionId", "SolutionVersion", "CaptureTimestampUtc", "SnapshotState",
                "OperationDiagnostic", "RequestCount", "ElapsedMilliseconds", "RawMembershipCount",
                "ResolvedCount", "UnsupportedCount", "UnresolvedCount", "AmbiguousCount",
                "ComparisonPresentInBoth", "ComparisonSourceOnly", "ComparisonTargetOnly",
                "ComparisonUnsupported", "ComparisonUnresolved", "ComparisonAmbiguous", "SolutionComponentId",
                "ObjectId", "RawComponentType", "RootComponentBehavior", "RootSolutionComponentId", "IsMetadata",
                "SemanticKind", "ComponentTypeKey", "ResolutionStatus", "ComparisonKey", "StableDiagnostic",
                "DiagnosticEvidence", "CoverageDisplayName", "CoverageBucketType", "CoverageStatus", "CoverageTotal",
                "CoverageResolved", "CoverageUnsupported", "CoverageUnresolved", "CoverageAmbiguous",
                "CoverageDiagnosticGroups", "RegisteredDefinitionName", "RegisteredDefinitionPrimaryEntity",
                "CanvasAppName", "CanvasAppId", "UniqueCanvasAppId", "CanvasAppDisplayName",
                "CanvasAppComponentState", "CanvasAppIsManaged", "CanvasAppCandidateStatus",
                "CanvasAppCandidateDiagnostic"
            }, header);
        }

        [TestMethod]
        public void LifecycleOperationIsTrimmedAndAppliedToEveryExportRow()
        {
            var rows = Parse(new MembershipCoverageCsvExporter().CreateCsv(CreatePresentation(CanvasIdentity()),
                "1.2.3.4", "1.2.3.3", "  Managed Upgrade  "));

            Assert.IsTrue(rows.Count > 0);
            Assert.IsTrue(rows.All(item => item["LifecycleOperation"] == "Managed Upgrade"));
        }

        [TestMethod]
        public void LifecycleExportKeepsCaseVariantsUnsupportedAndIndeterminate()
        {
            var source = CanvasIdentity(diagnosticEvidence: new[]
            {
                LookupEvidence(CanvasAppId, "new_edu_canvas"),
                "Canvas App lifecycle candidate status=CandidateValid; candidate='new_edu_canvas'. " +
                "Diagnostic validation only; the candidate is not used for membership comparison."
            });
            var target = CanvasIdentityForTarget("NEW_EDU_CANVAS");
            var presentation = CreateTwoSidedPresentation(source, target);

            var rows = Parse(new MembershipCoverageCsvExporter().CreateCsv(presentation,
                "1.2.3.4", "1.2.3.4", "Repeated import of the same unmanaged solution"))
                .Where(item => item["RowType"] == "Component" && item["RawComponentType"] == "300").ToList();

            Assert.AreEqual(2, rows.Count);
            Assert.AreEqual(0, presentation.Summary.PresentInBoth);
            Assert.AreEqual(2, presentation.Summary.Unsupported);
            Assert.IsTrue(presentation.Rows.All(item => item.MembershipStatus == "Indeterminate - Unsupported"));
            Assert.IsTrue(rows.All(item => item["ResolutionStatus"] == "Unsupported" &&
                item["ComparisonKey"].Length == 0 && item["SemanticKind"] == "unsupported:componenttype:300"));
            Assert.IsTrue(string.Equals(rows[0]["CanvasAppName"], rows[1]["CanvasAppName"],
                StringComparison.OrdinalIgnoreCase));
        }

        [TestMethod]
        public void ExportPreservesExactDiagnosticEvidenceAndCsvEscaping()
        {
            var identity = CanvasIdentity("Stable diagnostic, with \"quotes\"",
                "Additional evidence with comma, quote \" and\r\nline break.");
            var rows = Parse(new MembershipCoverageCsvExporter().CreateCsv(CreatePresentation(identity)));
            var component = rows.Single(item => item["RowType"] == "Component" && item["Side"] == "Source");

            Assert.AreEqual(identity.Diagnostic, component["StableDiagnostic"]);
            Assert.AreEqual(string.Join(Environment.NewLine, identity.DiagnosticEvidence),
                component["DiagnosticEvidence"]);
        }

        [TestMethod]
        public void ExportDoesNotChangeRequestsResolutionCoverageOrComparison()
        {
            var identity = CanvasIdentity();
            var presentation = CreatePresentation(identity);
            var sourceSnapshot = presentation.Source.Snapshot;
            var coverageBuilder = new MembershipCoverageDiagnosticsBuilder();
            var beforeCoverage = coverageBuilder.Build(sourceSnapshot);
            var beforeCoverageState = CoverageState(beforeCoverage);
            var beforeRows = presentation.Rows.Select(item => item.MembershipStatus).ToArray();
            var beforeSummary = ComparisonState(presentation.Summary);
            int beforeRequests = presentation.Source.Diagnostics.RequestCount;
            var service = new FakeOrganizationService();

            new MembershipCoverageCsvExporter().CreateCsv(presentation, "1.2.3.4", "1.2.3.3");

            var afterCoverage = coverageBuilder.Build(sourceSnapshot);
            Assert.AreSame(identity, sourceSnapshot.Components.Single());
            Assert.AreEqual(IdentityResolutionStatus.Unsupported, identity.Status);
            Assert.IsNull(identity.ComparisonKey);
            Assert.AreEqual(beforeRequests, presentation.Source.Diagnostics.RequestCount);
            Assert.AreEqual(0, service.Calls);
            Assert.AreEqual(0, service.ExecuteCalls);
            CollectionAssert.AreEqual(beforeRows,
                presentation.Rows.Select(item => item.MembershipStatus).ToArray());
            CollectionAssert.AreEqual(beforeSummary, ComparisonState(presentation.Summary));
            CollectionAssert.AreEqual(beforeCoverageState, CoverageState(afterCoverage));
            Assert.AreEqual(beforeCoverage.TotalCandidates, afterCoverage.TotalCandidates);
            Assert.AreEqual(beforeCoverage.BroadUnclassifiable.TotalCandidates,
                afterCoverage.BroadUnclassifiable.TotalCandidates);
            Assert.AreEqual(beforeCoverage.SemanticKinds.Sum(item => item.Resolved),
                afterCoverage.SemanticKinds.Sum(item => item.Resolved));
            Assert.AreEqual(beforeCoverage.SemanticKinds.Sum(item => item.Unsupported),
                afterCoverage.SemanticKinds.Sum(item => item.Unsupported));
            Assert.AreEqual(0, presentation.Summary.PresentInBoth);
            Assert.AreEqual(0, presentation.Summary.SourceOnly);
            Assert.AreEqual(0, presentation.Summary.TargetOnly);
            Assert.AreEqual(1, presentation.Summary.Unsupported);
            Assert.AreEqual("Indeterminate - Unsupported", presentation.Rows.Single().MembershipStatus);
        }

        private static string[] CoverageState(MembershipCoverageDiagnostics diagnostics) =>
            diagnostics.SemanticKinds.Concat(new[] { diagnostics.BroadUnclassifiable })
                .Select(item => string.Join("|", item.SemanticKind ?? "(broad)", item.BucketType,
                    item.CoverageStatus, item.TotalCandidates, item.Resolved, item.Unsupported,
                    item.Unresolved, item.Ambiguous)).ToArray();

        private static int[] ComparisonState(MembershipPresentationSummary summary) => new[]
        {
            summary.PresentInBoth, summary.SourceOnly, summary.TargetOnly,
            summary.Unsupported, summary.Unresolved, summary.Ambiguous
        };

        [TestMethod]
        public void DuplicateCanvasLookupEvidenceRemainsAuditableWithoutSelectingOneResult()
        {
            var first = LookupEvidence(CanvasAppId, "new_first");
            var second = LookupEvidence(Guid.Parse("70000000-0000-0000-0000-000000000007"), "new_second");
            var identity = CanvasIdentity(diagnosticEvidence: new[]
            {
                first, second,
                "Canvas App lifecycle candidate status=CorrelationDuplicate; candidate=(unavailable). " +
                "Diagnostic validation only; the candidate is not used for membership comparison."
            });

            var row = Parse(new MembershipCoverageCsvExporter().CreateCsv(CreatePresentation(identity)))
                .Single(item => item["RowType"] == "Component" && item["Side"] == "Source");

            Assert.AreEqual(string.Empty, row["CanvasAppName"]);
            Assert.AreEqual(string.Empty, row["CanvasAppId"]);
            Assert.AreEqual("CorrelationDuplicate", row["CanvasAppCandidateStatus"]);
            StringAssert.Contains(row["DiagnosticEvidence"], first);
            StringAssert.Contains(row["DiagnosticEvidence"], second);
        }

        [TestMethod]
        public void AbsentAndUnavailableSidesExportWithoutFabricatingIdentityOrCounts()
        {
            var sourceEnvironment = new EnvironmentIdentity(SourceEnvironmentId, "CSC-ICMS-DEV");
            var source = MembershipEnvironmentResult.FromSnapshot("CSC-ICMS-DEV",
                MembershipSnapshot.Absent(sourceEnvironment, "EDU", CapturedAt), 2, TimeSpan.FromMilliseconds(125));
            var target = MembershipEnvironmentResult.Unavailable("CSC-ICMS-UAT", "EDU", 4,
                TimeSpan.FromMilliseconds(250), "Read failed");
            var presentation = new MembershipResultPresenter().Create(source, target);

            var rows = Parse(new MembershipCoverageCsvExporter().CreateCsv(presentation, null, "9.9.9.9"));
            var sourceOperation = rows.Single(item => item["RowType"] == "OperationSummary" && item["Side"] == "Source");
            var targetOperation = rows.Single(item => item["RowType"] == "OperationSummary" && item["Side"] == "Target");
            Assert.AreEqual("SolutionAbsent", sourceOperation["SnapshotState"]);
            Assert.AreEqual("0", sourceOperation["RawMembershipCount"]);
            Assert.AreEqual(string.Empty, sourceOperation["SolutionId"]);
            Assert.AreEqual("Unavailable", targetOperation["SnapshotState"]);
            Assert.AreEqual(string.Empty, targetOperation["RawMembershipCount"]);
            Assert.AreEqual(string.Empty, targetOperation["EnvironmentOrganizationId"]);
            Assert.AreEqual("Read failed", targetOperation["OperationDiagnostic"]);
            Assert.AreEqual("9.9.9.9", targetOperation["SolutionVersion"]);
        }

        [TestMethod]
        public void FileExportUsesUtf8Bom()
        {
            var path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".csv");
            try
            {
                new MembershipCoverageCsvExporter().WriteCsv(path, CreatePresentation(CanvasIdentity()));
                var bytes = File.ReadAllBytes(path);
                CollectionAssert.AreEqual(new byte[] { 0xEF, 0xBB, 0xBF }, bytes.Take(3).ToArray());
            }
            finally
            {
                if (File.Exists(path)) File.Delete(path);
            }
        }

        private static MembershipComparisonPresentation CreatePresentation(ComponentIdentity identity)
        {
            var sourceEnvironment = new EnvironmentIdentity(SourceEnvironmentId, "CSC-ICMS-DEV");
            var targetEnvironment = new EnvironmentIdentity(TargetEnvironmentId, "CSC-ICMS-UAT");
            var solution = new SolutionIdentity(sourceEnvironment, SolutionId, "EDU");
            var source = MembershipEnvironmentResult.FromSnapshot("CSC-ICMS-DEV",
                MembershipSnapshot.Complete(solution, new[] { identity }, CapturedAt), 17,
                TimeSpan.FromMilliseconds(1500));
            var target = MembershipEnvironmentResult.FromSnapshot("CSC-ICMS-UAT",
                MembershipSnapshot.Absent(targetEnvironment, "EDU", CapturedAt.AddMinutes(1)), 2,
                TimeSpan.FromMilliseconds(200));
            return new MembershipResultPresenter().Create(source, target);
        }

        private static MembershipComparisonPresentation CreateTwoSidedPresentation(ComponentIdentity sourceIdentity,
            ComponentIdentity targetIdentity)
        {
            var sourceEnvironment = new EnvironmentIdentity(SourceEnvironmentId, "CSC-ICMS-DEV");
            var targetEnvironment = new EnvironmentIdentity(TargetEnvironmentId, "CSC-ICMS-UAT");
            var sourceSolution = new SolutionIdentity(sourceEnvironment, SolutionId, "EDU");
            var targetSolution = new SolutionIdentity(targetEnvironment,
                Guid.Parse("90000000-0000-0000-0000-000000000009"), "EDU");
            return new MembershipResultPresenter().Create(
                MembershipEnvironmentResult.FromSnapshot("CSC-ICMS-DEV",
                    MembershipSnapshot.Complete(sourceSolution, new[] { sourceIdentity }, CapturedAt),
                    17, TimeSpan.FromMilliseconds(1500)),
                MembershipEnvironmentResult.FromSnapshot("CSC-ICMS-UAT",
                    MembershipSnapshot.Complete(targetSolution, new[] { targetIdentity }, CapturedAt.AddMinutes(1)),
                    18, TimeSpan.FromMilliseconds(1600)));
        }

        private static ComponentIdentity CanvasIdentityForTarget(string name)
        {
            var targetObjectId = Guid.Parse("A0000000-0000-0000-0000-00000000000A");
            var targetCanvasId = Guid.Parse("B0000000-0000-0000-0000-00000000000B");
            return new ComponentIdentity(new SolutionComponentRecord(
                Guid.Parse("C0000000-0000-0000-0000-00000000000C"), 300, targetObjectId),
                IdentityResolutionStatus.Unsupported, diagnostic: "Unsupported solution component type 300.",
                componentTypeKey: "unsupported:componenttype:300",
                semanticKind: "unsupported:componenttype:300", diagnosticEvidence: new[]
                {
                    LookupEvidence(targetCanvasId, name),
                    "Canvas App lifecycle candidate status=CandidateValid; candidate='" + name + "'. " +
                    "Diagnostic validation only; the candidate is not used for membership comparison."
                });
        }

        private static ComponentIdentity CanvasIdentity(string diagnostic = null,
            string additionalEvidence = null, IEnumerable<string> diagnosticEvidence = null)
        {
            var evidence = diagnosticEvidence == null ? new List<string>
            {
                LookupEvidence(CanvasAppId, "new_edu_canvas"),
                "Canvas App lifecycle candidate status=CandidateValid; candidate='new_edu_canvas'. " +
                "Diagnostic validation only; the candidate is not used for membership comparison.",
                "Canvas App diagnostic summary: RawType300MembershipCount=1; DistinctNonemptyObjectIdCount=1; " +
                "ReturnedCanvasAppRowCount=1; UniqueObjectIdCorrelationCount=1; MissingRequestedObjectIdCount=0; " +
                "BlankNameCount=0; ValidCandidateCount=1."
            } : diagnosticEvidence.ToList();
            if (additionalEvidence != null) evidence.Add(additionalEvidence);
            return new ComponentIdentity(new SolutionComponentRecord(SolutionComponentId, 300, ObjectId,
                rootComponentBehavior: 0, rootSolutionComponentId: Guid.Parse("80000000-0000-0000-0000-000000000008"),
                isMetadata: false), IdentityResolutionStatus.Unsupported,
                diagnostic: diagnostic ?? "Unsupported solution component type 300.",
                componentTypeKey: "unsupported:componenttype:300",
                semanticKind: "unsupported:componenttype:300", diagnosticEvidence: evidence);
        }

        private static string LookupEvidence(Guid canvasAppId, string name) =>
            "Canvas App diagnostic lookup matched. canvasappid=" + canvasAppId.ToString("D") +
            "; name='" + name + "'; displayname='EDU, Canvas'; uniquecanvasappid='shared-canvas-token'" +
            "; componentstate=0 ('Published'); ismanaged=True; candidateportableidentity='" + name + "'. " +
            "Diagnostic evidence only; the candidate is not used for membership comparison.";

        private static List<Dictionary<string, string>> Parse(string csv)
        {
            var records = ParseRecords(csv);
            var headers = records[0];
            return records.Skip(1).Where(item => item.Count == headers.Count)
                .Select(item => headers.Select((header, index) => new { header, value = item[index] })
                    .ToDictionary(pair => pair.header, pair => pair.value, StringComparer.Ordinal)).ToList();
        }

        private static List<List<string>> ParseRecords(string csv)
        {
            var records = new List<List<string>>();
            var record = new List<string>();
            var field = new StringBuilder();
            bool quoted = false;
            for (int index = 0; index < csv.Length; index++)
            {
                char character = csv[index];
                if (quoted)
                {
                    if (character == '"' && index + 1 < csv.Length && csv[index + 1] == '"')
                    {
                        field.Append('"'); index++;
                    }
                    else if (character == '"') quoted = false;
                    else field.Append(character);
                }
                else if (character == '"') quoted = true;
                else if (character == ',') { record.Add(field.ToString()); field.Clear(); }
                else if (character == '\r' && index + 1 < csv.Length && csv[index + 1] == '\n')
                {
                    record.Add(field.ToString()); field.Clear(); records.Add(record); record = new List<string>(); index++;
                }
                else field.Append(character);
            }
            return records;
        }
    }
}
