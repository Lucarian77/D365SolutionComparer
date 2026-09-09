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
    public class CanvasAppLifecycleCsvComparerTests
    {
        private const string Operation = "Managed Upgrade";
        private static readonly Guid EnvironmentId = Guid.Parse("11000000-0000-0000-0000-000000000001");
        private static readonly Guid SolutionId = Guid.Parse("22000000-0000-0000-0000-000000000002");
        private static readonly DateTimeOffset BaseCapture =
            new DateTimeOffset(2026, 9, 9, 12, 0, 0, TimeSpan.Zero);

        [TestMethod]
        public void ChronologyAllowsBeforeEarlierThanAfter()
        {
            var rows = Parse(new CanvasAppLifecycleCsvComparer().Compare(
                ExportAt(Operation, EnvironmentId, "CSC-ICMS-DEV", "EDU", BaseCapture,
                    App("new_canvas", 1)),
                ExportAt(Operation, EnvironmentId, "CSC-ICMS-DEV", "EDU", BaseCapture.AddHours(1),
                    App("NEW_CANVAS", 2))));

            Assert.AreEqual(1, rows.Count);
            Assert.AreEqual("CandidateStable", rows[0]["Outcome"]);
        }

        [TestMethod]
        public void ChronologyRejectsReversedDevTimestampsWithClearDiagnostic()
        {
            var exception = Assert.ThrowsException<CanvasAppLifecycleChronologyException>(() =>
                new CanvasAppLifecycleCsvComparer().Compare(
                    ExportAt(Operation, EnvironmentId, "CSC-ICMS-DEV", "EDU", BaseCapture.AddHours(1),
                        App("new_canvas", 1)),
                    ExportAt(Operation, EnvironmentId, "CSC-ICMS-DEV", "EDU", BaseCapture,
                        App("new_canvas", 2))));

            Assert.AreEqual("CSC-ICMS-DEV", exception.EnvironmentName);
            Assert.AreEqual(BaseCapture.AddHours(1), exception.BeforeCaptureUtc);
            Assert.AreEqual(BaseCapture, exception.AfterCaptureUtc);
            StringAssert.Contains(exception.Message, "selected lifecycle snapshots appear to be reversed");
            StringAssert.Contains(exception.Message, "Before snapshot was captured after the After snapshot");
            StringAssert.Contains(exception.Message, "Environment: CSC-ICMS-DEV");
            StringAssert.Contains(exception.Message, "Before: " + BaseCapture.AddHours(1).UtcDateTime.ToString("o"));
            StringAssert.Contains(exception.Message, "After: " + BaseCapture.UtcDateTime.ToString("o"));
        }

        [TestMethod]
        public void ChronologyRejectsReversedUatTimestamps()
        {
            var uatId = Guid.Parse("55000000-0000-0000-0000-000000000005");
            var exception = Assert.ThrowsException<CanvasAppLifecycleChronologyException>(() =>
                new CanvasAppLifecycleCsvComparer().Compare(
                    ExportAt(Operation, uatId, "CSC-ICMS-UAT", "EDU", BaseCapture.AddMinutes(30),
                        App("new_canvas", 1)),
                    ExportAt(Operation, uatId, "CSC-ICMS-UAT", "EDU", BaseCapture,
                        App("new_canvas", 2))));

            Assert.AreEqual("CSC-ICMS-UAT", exception.EnvironmentName);
        }

        [TestMethod]
        public void OneValidAndOneReversedEnvironmentRejectsEntireComparison()
        {
            var uatId = Guid.Parse("55000000-0000-0000-0000-000000000005");
            var before = MergeCoverageCsv(
                ExportAt(Operation, EnvironmentId, "CSC-ICMS-DEV", "EDU", BaseCapture,
                    App("new_dev", 1)),
                ExportAt(Operation, uatId, "CSC-ICMS-UAT", "EDU", BaseCapture.AddHours(2),
                    App("new_uat", 2)));
            var after = MergeCoverageCsv(
                ExportAt(Operation, EnvironmentId, "CSC-ICMS-DEV", "EDU", BaseCapture.AddHours(1),
                    App("new_dev", 3)),
                ExportAt(Operation, uatId, "CSC-ICMS-UAT", "EDU", BaseCapture.AddHours(1),
                    App("new_uat", 4)));

            var exception = Assert.ThrowsException<CanvasAppLifecycleChronologyException>(() =>
                new CanvasAppLifecycleCsvComparer().Compare(before, after));

            Assert.AreEqual("CSC-ICMS-UAT", exception.EnvironmentName);
        }

        [TestMethod]
        public void ChronologyUsesCsvCaptureTimestampsRatherThanFilenameOrder()
        {
            var directory = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(directory);
            var beforePath = Path.Combine(directory, "2000-before.csv");
            var afterPath = Path.Combine(directory, "2099-after.csv");
            var outputPath = Path.Combine(directory, "comparison.csv");
            try
            {
                File.WriteAllText(beforePath, ExportAt(Operation, EnvironmentId, "CSC-ICMS-DEV", "EDU",
                    BaseCapture.AddHours(1), App("new_canvas", 1)));
                File.WriteAllText(afterPath, ExportAt(Operation, EnvironmentId, "CSC-ICMS-DEV", "EDU",
                    BaseCapture, App("new_canvas", 2)));

                Assert.ThrowsException<CanvasAppLifecycleChronologyException>(() =>
                    new CanvasAppLifecycleCsvComparer().CompareFiles(beforePath, afterPath, outputPath));
                Assert.IsFalse(File.Exists(outputPath));
            }
            finally
            {
                if (Directory.Exists(directory)) Directory.Delete(directory, true);
            }
        }

        [TestMethod]
        public void FilenameTimestampsThatAppearReversedDoNotOverrideCorrectCsvChronology()
        {
            var directory = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(directory);
            var beforePath = Path.Combine(directory, "2099-before.csv");
            var afterPath = Path.Combine(directory, "2000-after.csv");
            var outputPath = Path.Combine(directory, "comparison.csv");
            try
            {
                File.WriteAllText(beforePath, ExportAt(Operation, EnvironmentId, "CSC-ICMS-DEV", "EDU",
                    BaseCapture, App("new_canvas", 1)));
                File.WriteAllText(afterPath, ExportAt(Operation, EnvironmentId, "CSC-ICMS-DEV", "EDU",
                    BaseCapture.AddHours(1), App("new_canvas", 2)));

                new CanvasAppLifecycleCsvComparer().CompareFiles(beforePath, afterPath, outputPath);

                Assert.IsTrue(File.Exists(outputPath));
                Assert.AreEqual("CandidateStable", Parse(File.ReadAllText(outputPath)).Single()["Outcome"]);
            }
            finally
            {
                if (Directory.Exists(directory)) Directory.Delete(directory, true);
            }
        }

        [TestMethod]
        public void EqualCaptureTimestampsPreserveExistingComparisonBehavior()
        {
            var rows = Parse(new CanvasAppLifecycleCsvComparer().Compare(
                ExportAt(Operation, EnvironmentId, "CSC-ICMS-DEV", "EDU", BaseCapture,
                    App("new_canvas", 1)),
                ExportAt(Operation, EnvironmentId, "CSC-ICMS-DEV", "EDU", BaseCapture,
                    App("NEW_CANVAS", 2))));

            Assert.AreEqual("CandidateStable", rows.Single()["Outcome"]);
        }

        [TestMethod]
        public void LifecycleOperationMismatchStillFailsBeforeChronologyComparison()
        {
            var exception = Assert.ThrowsException<InvalidDataException>(() =>
                new CanvasAppLifecycleCsvComparer().Compare(
                    ExportAt("Managed Update", EnvironmentId, "CSC-ICMS-DEV", "EDU",
                        BaseCapture.AddHours(2), App("new_canvas", 1)),
                    ExportAt("Managed Upgrade", EnvironmentId, "CSC-ICMS-DEV", "EDU",
                        BaseCapture, App("new_canvas", 2))));

            StringAssert.Contains(exception.Message, "lifecycle operations do not match");
        }

        [TestMethod]
        public void ValidChronologyPreservesEveryDeterministicOutcome()
        {
            var before = ExportAt(Operation, EnvironmentId, "CSC-ICMS-DEV", "EDU", BaseCapture,
                App("stable", 1), App("changed_before", 2), App("duplicate", 3), App("DUPLICATE", 4));
            var after = ExportAt(Operation, EnvironmentId, "CSC-ICMS-DEV", "EDU", BaseCapture.AddMinutes(1),
                App("STABLE", 5), App("changed_after", 6), App("duplicate", 7));
            var comparer = new CanvasAppLifecycleCsvComparer();

            var first = Parse(comparer.Compare(before, after));
            var second = Parse(comparer.Compare(before, after));

            CollectionAssert.AreEqual(first.Select(item => item["Outcome"]).ToArray(),
                second.Select(item => item["Outcome"]).ToArray());
            CollectionAssert.Contains(first.Select(item => item["Outcome"]).ToArray(), "CandidateStable");
            CollectionAssert.Contains(first.Select(item => item["Outcome"]).ToArray(), "CandidateChanged");
            CollectionAssert.Contains(first.Select(item => item["Outcome"]).ToArray(), "DuplicateCandidate");
        }

        [TestMethod]
        public void ReversedChronologyDoesNotCreateOrOverwriteComparisonOutput()
        {
            var directory = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(directory);
            var beforePath = Path.Combine(directory, "before.csv");
            var afterPath = Path.Combine(directory, "after.csv");
            var outputPath = Path.Combine(directory, "comparison.csv");
            try
            {
                File.WriteAllText(beforePath, ExportAt(Operation, EnvironmentId, "CSC-ICMS-DEV", "EDU",
                    BaseCapture.AddHours(1), App("new_canvas", 1)));
                File.WriteAllText(afterPath, ExportAt(Operation, EnvironmentId, "CSC-ICMS-DEV", "EDU",
                    BaseCapture, App("new_canvas", 2)));
                File.WriteAllText(outputPath, "existing comparison evidence");

                Assert.ThrowsException<CanvasAppLifecycleChronologyException>(() =>
                    new CanvasAppLifecycleCsvComparer().CompareFiles(beforePath, afterPath, outputPath));

                Assert.AreEqual("existing comparison evidence", File.ReadAllText(outputPath));
            }
            finally
            {
                if (Directory.Exists(directory)) Directory.Delete(directory, true);
            }
        }

        [TestMethod]
        public void CaseInsensitiveCandidatePairsAndReportsAuditFieldDifferences()
        {
            var before = Export(Operation, App("new_edu_canvas", 1, "Before name", false, 0));
            var after = Export(Operation, App("NEW_EDU_CANVAS", 2, "After name", true, 1));

            var row = Parse(new CanvasAppLifecycleCsvComparer().Compare(before, after)).Single();

            Assert.AreEqual("CandidateStable", row["Outcome"]);
            Assert.AreEqual("CaseInsensitiveCanvasAppName", row["CorrelationBasis"]);
            Assert.AreEqual("True", row["CandidateNameEqualOrdinalIgnoreCase"]);
            Assert.AreEqual("new_edu_canvas", row["BeforeCanvasAppName"]);
            Assert.AreEqual("NEW_EDU_CANVAS", row["AfterCanvasAppName"]);
            StringAssert.Contains(row["AuditDifferences"], "ObjectId");
            StringAssert.Contains(row["AuditDifferences"], "CanvasAppId");
            StringAssert.Contains(row["AuditDifferences"], "UniqueCanvasAppId");
            StringAssert.Contains(row["AuditDifferences"], "CanvasAppDisplayName");
            StringAssert.Contains(row["AuditDifferences"], "CanvasAppComponentState");
            StringAssert.Contains(row["AuditDifferences"], "CanvasAppIsManaged");
        }

        [TestMethod]
        public void SingleRemainingCandidatesWithDifferentNamesAreReportedAsChangedNotMatched()
        {
            var result = Parse(new CanvasAppLifecycleCsvComparer().Compare(
                Export(Operation, App("new_before", 1)), Export(Operation, App("new_after", 2)))).Single();

            Assert.AreEqual("CandidateChanged", result["Outcome"]);
            Assert.AreEqual("SingleRemainingCandidateWithinPartition", result["CorrelationBasis"]);
            Assert.AreEqual("False", result["CandidateNameEqualOrdinalIgnoreCase"]);
        }

        [TestMethod]
        public void DuplicateCaseInsensitiveNamesRemainExplicitAndUnpaired()
        {
            var rows = Parse(new CanvasAppLifecycleCsvComparer().Compare(
                Export(Operation, App("new_duplicate", 1), App("NEW_DUPLICATE", 2)),
                Export(Operation, App("new_duplicate", 3))));

            Assert.AreEqual(3, rows.Count);
            Assert.IsTrue(rows.All(item => item["Outcome"] == "DuplicateCandidate"));
            Assert.IsTrue(rows.All(item => item["CorrelationBasis"] ==
                "DuplicateCaseInsensitiveCanvasAppName"));
            Assert.AreEqual(2, rows.Count(item => item["BeforeObjectId"].Length > 0));
            Assert.AreEqual(1, rows.Count(item => item["AfterObjectId"].Length > 0));
        }

        [TestMethod]
        public void RepeatedRawMembershipForSameObjectDoesNotCreateDuplicateCandidate()
        {
            var first = App("new_repeated", 1);
            var repeated = App("NEW_REPEATED", 2);
            repeated.ObjectId = first.ObjectId;
            repeated.CanvasAppId = first.CanvasAppId;
            repeated.UniqueCanvasAppId = first.UniqueCanvasAppId;
            var rows = Parse(new CanvasAppLifecycleCsvComparer().Compare(
                Export(Operation, repeated, first), Export(Operation, App("new_repeated", 3))));

            Assert.AreEqual(2, rows.Count);
            Assert.AreEqual(1, rows.Count(item => item["Outcome"] == "CandidateStable"));
            Assert.AreEqual(1, rows.Count(item => item["Outcome"] == "RepeatedMembershipEvidence"));
            Assert.IsFalse(rows.Any(item => item["Outcome"] == "DuplicateCandidate"));
        }

        [TestMethod]
        public void MissingBeforeOrAfterRowsAreReportedWithoutFabricatedPairing()
        {
            var beforeOnly = Parse(new CanvasAppLifecycleCsvComparer().Compare(
                Export(Operation, App("new_removed", 1)), Export(Operation))).Single();
            var afterOnly = Parse(new CanvasAppLifecycleCsvComparer().Compare(
                Export(Operation), Export(Operation, App("new_added", 2)))).Single();

            Assert.AreEqual("BeforeOnly", beforeOnly["Outcome"]);
            Assert.AreEqual("NoAfterCandidate", beforeOnly["CorrelationBasis"]);
            Assert.AreEqual(string.Empty, beforeOnly["AfterCanvasAppName"]);
            Assert.AreEqual("AfterOnly", afterOnly["Outcome"]);
            Assert.AreEqual("NoBeforeCandidate", afterOnly["CorrelationBasis"]);
            Assert.AreEqual(string.Empty, afterOnly["BeforeCanvasAppName"]);
        }

        [DataTestMethod]
        [DataRow(true, false)]
        [DataRow(false, true)]
        public void CandidateNamesAreCorrelatedOnlyWithinTheSameEnvironmentAndSolution(
            bool changeEnvironment, bool changeSolution)
        {
            var afterEnvironment = changeEnvironment
                ? Guid.Parse("44000000-0000-0000-0000-000000000004") : EnvironmentId;
            var afterSolution = changeSolution ? "EDU_CLONE" : "EDU";
            var rows = Parse(new CanvasAppLifecycleCsvComparer().Compare(
                Export(Operation, App("new_canvas", 1)),
                ExportFor(Operation, afterEnvironment, afterSolution, App("NEW_CANVAS", 2))));

            Assert.AreEqual(2, rows.Count);
            CollectionAssert.AreEquivalent(new[] { "BeforeOnly", "AfterOnly" },
                rows.Select(item => item["Outcome"]).ToArray());
            Assert.IsFalse(rows.Any(item => item["Outcome"] == "CandidateStable"));
        }

        [TestMethod]
        public void BlankOrNonvalidCandidateEvidenceIsInsufficient()
        {
            var blank = App(string.Empty, 1);
            blank.CandidateStatus = "BlankName";
            var faulted = App("new_evidence_only", 2);
            faulted.CandidateStatus = "Faulted";
            var rows = Parse(new CanvasAppLifecycleCsvComparer().Compare(
                Export(Operation, blank), Export(Operation, faulted)));

            Assert.AreEqual(2, rows.Count);
            Assert.IsTrue(rows.All(item => item["Outcome"] == "InsufficientEvidence"));
            Assert.IsTrue(rows.All(item => item["CorrelationBasis"] == "CandidateEvidenceIncomplete"));
        }

        [TestMethod]
        public void MultipleUnmatchedCandidatesRemainAmbiguousAndOutputIsDeterministic()
        {
            var before = Export(Operation, App("new_b", 2), App("new_a", 1));
            var after = Export(Operation, App("new_d", 4), App("new_c", 3));
            var comparer = new CanvasAppLifecycleCsvComparer();

            var first = comparer.Compare(before, after);
            var second = comparer.Compare(before, after);
            var rows = Parse(first);

            Assert.AreEqual(first, second);
            Assert.AreEqual(4, rows.Count);
            Assert.IsTrue(rows.All(item => item["Outcome"] == "AmbiguousCorrelation"));
            CollectionAssert.AreEqual(new[] { "new_a", "new_b", string.Empty, string.Empty },
                rows.Select(item => item["BeforeCanvasAppName"]).ToArray());
        }

        [TestMethod]
        public void CurrentSchemaAndQuotedMultilineEvidenceAreReadExactly()
        {
            var app = App("new_canvas", 1);
            app.CandidateDiagnostic = "Canvas App lifecycle candidate status=CandidateValid; " +
                "candidate='new_canvas'. Diagnostic, with \"quotes\" and\r\na second line.";
            var csv = Export(Operation, app);

            var row = Parse(new CanvasAppLifecycleCsvComparer().Compare(csv, csv)).Single();

            Assert.AreEqual(app.CandidateDiagnostic, row["BeforeCanvasAppCandidateDiagnostic"]);
            Assert.AreEqual(app.CandidateDiagnostic, row["AfterCanvasAppCandidateDiagnostic"]);
        }

        [TestMethod]
        public void LifecycleComparisonOutputSchemaIsStable()
        {
            var csv = new CanvasAppLifecycleCsvComparer().Compare(
                Export(Operation, App("new_canvas", 1)), Export(Operation, App("NEW_CANVAS", 2)));
            var evidence = new[]
            {
                "Side", "Environment", "EnvironmentOrganizationId", "SolutionUniqueName", "SolutionId",
                "SolutionVersion", "CaptureTimestampUtc", "LifecycleOperation", "RawComponentType",
                "SolutionComponentId", "ObjectId", "CanvasAppId", "CanvasAppName", "CanvasAppDisplayName",
                "UniqueCanvasAppId", "CanvasAppComponentState", "CanvasAppIsManaged", "CanvasAppCandidateStatus",
                "CanvasAppCandidateDiagnostic", "SemanticKind", "ComponentTypeKey", "ResolutionStatus",
                "ComparisonKey", "StableDiagnostic", "DiagnosticEvidence"
            };
            var expected = new[]
            {
                "Outcome", "CorrelationBasis", "Diagnostic", "CandidateNameEqualOrdinalIgnoreCase",
                "AuditDifferences"
            }.Concat(evidence.Select(item => "Before" + item))
                .Concat(evidence.Select(item => "After" + item)).ToArray();

            CollectionAssert.AreEqual(expected,
                csv.Substring(0, csv.IndexOf("\r\n", StringComparison.Ordinal)).Split(','));
        }

        [TestMethod]
        public void MissingHeadersMalformedRowsAndLifecycleMismatchAreRejected()
        {
            var valid = Export(Operation, App("new_canvas", 1));
            var missingHeader = valid.Replace("CanvasAppCandidateDiagnostic", "RemovedCandidateDiagnostic");
            Assert.ThrowsException<InvalidDataException>(() =>
                new CanvasAppLifecycleCsvComparer().Compare(missingHeader, valid));
            Assert.ThrowsException<InvalidDataException>(() =>
                new CanvasAppLifecycleCsvComparer().Compare(valid + "\"unterminated", valid));
            Assert.ThrowsException<InvalidDataException>(() =>
                new CanvasAppLifecycleCsvComparer().Compare(valid,
                    Export("Delete and recreate", App("new_canvas", 1))));
        }

        [TestMethod]
        public void ComparisonDoesNotChangeRequestsResolutionCoverageOrMembershipSemantics()
        {
            var app = App("new_canvas", 1);
            var identity = Identity(app);
            var presentation = Presentation(identity);
            var beforeStatus = identity.Status;
            var beforeKey = identity.ComparisonKey;
            var beforeRequests = presentation.Source.Diagnostics.RequestCount;
            var beforeMembership = presentation.Rows.Single().MembershipStatus;
            var csv = new MembershipCoverageCsvExporter().CreateCsv(presentation, "1.0.0.0", null, Operation);
            var service = new FakeOrganizationService();

            new CanvasAppLifecycleCsvComparer().Compare(csv, csv);

            Assert.AreEqual(beforeStatus, identity.Status);
            Assert.AreEqual(IdentityResolutionStatus.Unsupported, identity.Status);
            Assert.AreEqual(beforeKey, identity.ComparisonKey);
            Assert.IsNull(identity.ComparisonKey);
            Assert.AreEqual(beforeRequests, presentation.Source.Diagnostics.RequestCount);
            Assert.AreEqual(0, service.Calls);
            Assert.AreEqual(0, service.ExecuteCalls);
            Assert.AreEqual(beforeMembership, presentation.Rows.Single().MembershipStatus);
            Assert.AreEqual("Indeterminate - Unsupported", presentation.Rows.Single().MembershipStatus);
            Assert.AreEqual(0, presentation.Summary.PresentInBoth);
            Assert.AreEqual(0, presentation.Summary.SourceOnly);
            Assert.AreEqual(0, presentation.Summary.TargetOnly);
        }

        [TestMethod]
        public void FileComparisonReadsUtf8BomAndWritesUtf8Bom()
        {
            var beforePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + "-before.csv");
            var afterPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + "-after.csv");
            var outputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + "-result.csv");
            try
            {
                File.WriteAllText(beforePath, Export(Operation, App("new_canvas", 1)), new UTF8Encoding(true));
                File.WriteAllText(afterPath, Export(Operation, App("NEW_CANVAS", 2)), new UTF8Encoding(true));
                new CanvasAppLifecycleCsvComparer().CompareFiles(beforePath, afterPath, outputPath);

                var bytes = File.ReadAllBytes(outputPath);
                CollectionAssert.AreEqual(new byte[] { 0xEF, 0xBB, 0xBF }, bytes.Take(3).ToArray());
                Assert.AreEqual("CandidateStable", Parse(File.ReadAllText(outputPath)).Single()["Outcome"]);
            }
            finally
            {
                foreach (var path in new[] { beforePath, afterPath, outputPath })
                    if (File.Exists(path)) File.Delete(path);
            }
        }

        private static string Export(string operation, params AppEvidence[] apps)
        {
            return ExportFor(operation, EnvironmentId, "EDU", apps);
        }

        private static string ExportAt(string operation, Guid environmentId, string environmentName,
            string solutionName, DateTimeOffset capture, params AppEvidence[] apps) =>
            new MembershipCoverageCsvExporter().CreateCsv(
                Presentation(environmentId, environmentName, solutionName, capture,
                    apps.Select(Identity).ToArray()), "1.0.0.0", null, operation);

        private static string MergeCoverageCsv(params string[] files)
        {
            var firstBreak = files[0].IndexOf("\r\n", StringComparison.Ordinal);
            var result = new StringBuilder(files[0]);
            foreach (var file in files.Skip(1))
            {
                int headerEnd = file.IndexOf("\r\n", StringComparison.Ordinal);
                result.Append(file.Substring(headerEnd + 2));
            }
            Assert.IsTrue(firstBreak > 0);
            return result.ToString();
        }

        private static string ExportFor(string operation, Guid environmentId, string solutionName,
            params AppEvidence[] apps) => new MembershipCoverageCsvExporter().CreateCsv(
                Presentation(environmentId, solutionName, apps.Select(Identity).ToArray()),
                "1.0.0.0", null, operation);

        private static MembershipComparisonPresentation Presentation(params ComponentIdentity[] identities)
        {
            return Presentation(EnvironmentId, "EDU", identities);
        }

        private static MembershipComparisonPresentation Presentation(Guid environmentId, string solutionName,
            params ComponentIdentity[] identities)
        {
            return Presentation(environmentId, "CSC-ICMS-DEV", solutionName, BaseCapture, identities);
        }

        private static MembershipComparisonPresentation Presentation(Guid environmentId, string environmentName,
            string solutionName, DateTimeOffset capture, params ComponentIdentity[] identities)
        {
            var environment = new EnvironmentIdentity(environmentId, environmentName);
            var solution = new SolutionIdentity(environment, SolutionId, solutionName);
            var source = MembershipEnvironmentResult.FromSnapshot(environmentName,
                MembershipSnapshot.Complete(solution, identities, capture), 21, TimeSpan.FromSeconds(2));
            var target = MembershipEnvironmentResult.FromSnapshot("CSC-ICMS-UAT",
                MembershipSnapshot.Absent(new EnvironmentIdentity(
                    Guid.Parse("33000000-0000-0000-0000-000000000003"), "CSC-ICMS-UAT"), solutionName,
                    capture.AddMinutes(1)), 2, TimeSpan.FromSeconds(1));
            return new MembershipResultPresenter().Create(source, target);
        }

        private static ComponentIdentity Identity(AppEvidence app)
        {
            var lookup = "Canvas App diagnostic lookup matched. canvasappid=" + app.CanvasAppId.ToString("D") +
                "; name='" + app.Name + "'; displayname='" + app.DisplayName + "'; uniquecanvasappid='" +
                app.UniqueCanvasAppId + "'; componentstate=" + app.ComponentState + "; ismanaged=" + app.IsManaged +
                "; candidateportableidentity=" + (app.Name.Length == 0 ? "(unavailable)" : "'" + app.Name + "'") +
                ". Diagnostic evidence only; the candidate is not used for membership comparison.";
            var candidate = app.CandidateDiagnostic ?? "Canvas App lifecycle candidate status=" +
                app.CandidateStatus + "; candidate=" + (app.Name.Length == 0 ? "(unavailable)" : "'" + app.Name +
                "'") + ". Diagnostic validation only; the candidate is not used for membership comparison.";
            return new ComponentIdentity(new SolutionComponentRecord(app.SolutionComponentId, 300, app.ObjectId),
                IdentityResolutionStatus.Unsupported, diagnostic: "Unsupported solution component type 300.",
                componentTypeKey: "unsupported:componenttype:300", semanticKind: "unsupported:componenttype:300",
                diagnosticEvidence: new[] { lookup, candidate });
        }

        private static AppEvidence App(string name, int seed, string displayName = "EDU Canvas",
            bool isManaged = false, int componentState = 0)
        {
            return new AppEvidence
            {
                Name = name,
                ObjectId = GuidFrom(seed, 1),
                CanvasAppId = GuidFrom(seed, 2),
                SolutionComponentId = GuidFrom(seed, 3),
                UniqueCanvasAppId = "unique-" + seed,
                DisplayName = displayName,
                IsManaged = isManaged.ToString(),
                ComponentState = componentState.ToString() + (componentState == 0 ? " ('Published')" : string.Empty),
                CandidateStatus = "CandidateValid"
            };
        }

        private static Guid GuidFrom(int seed, int suffix) =>
            Guid.Parse(seed.ToString("X8") + "-0000-0000-0000-" + suffix.ToString("D12"));

        private static List<Dictionary<string, string>> Parse(string csv)
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
                    { field.Append('"'); index++; }
                    else if (character == '"') quoted = false;
                    else field.Append(character);
                }
                else if (character == '"') quoted = true;
                else if (character == ',') { record.Add(field.ToString()); field.Clear(); }
                else if (character == '\r' && index + 1 < csv.Length && csv[index + 1] == '\n')
                { record.Add(field.ToString()); field.Clear(); records.Add(record); record = new List<string>(); index++; }
                else field.Append(character);
            }
            records[0][0] = records[0][0].TrimStart('\uFEFF');
            var headers = records[0];
            return records.Skip(1).Where(item => item.Count == headers.Count)
                .Select(item => headers.Select((header, index) => new { header, value = item[index] })
                    .ToDictionary(pair => pair.header, pair => pair.value, StringComparer.Ordinal)).ToList();
        }

        private sealed class AppEvidence
        {
            public string Name { get; set; }
            public Guid ObjectId { get; set; }
            public Guid CanvasAppId { get; set; }
            public Guid SolutionComponentId { get; set; }
            public string UniqueCanvasAppId { get; set; }
            public string DisplayName { get; set; }
            public string IsManaged { get; set; }
            public string ComponentState { get; set; }
            public string CandidateStatus { get; set; }
            public string CandidateDiagnostic { get; set; }
        }
    }
}
