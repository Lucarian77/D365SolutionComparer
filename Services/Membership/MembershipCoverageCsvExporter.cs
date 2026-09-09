using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using D365SolutionComparer.Models.Membership;

namespace D365SolutionComparer.Services.Membership
{
    /// <summary>Exports already-retrieved membership and coverage diagnostics without accessing Dataverse.</summary>
    public sealed class MembershipCoverageCsvExporter
    {
        private static readonly string[] Headers =
        {
            "RowType", "Side", "Environment", "EnvironmentOrganizationId", "SolutionUniqueName",
            "SolutionId", "SolutionVersion", "CaptureTimestampUtc", "SnapshotState", "OperationDiagnostic",
            "RequestCount", "ElapsedMilliseconds", "RawMembershipCount", "ResolvedCount", "UnsupportedCount",
            "UnresolvedCount", "AmbiguousCount", "ComparisonPresentInBoth", "ComparisonSourceOnly",
            "ComparisonTargetOnly", "ComparisonUnsupported", "ComparisonUnresolved", "ComparisonAmbiguous",
            "SolutionComponentId", "ObjectId", "RawComponentType", "RootComponentBehavior",
            "RootSolutionComponentId", "IsMetadata", "SemanticKind", "ComponentTypeKey", "ResolutionStatus",
            "ComparisonKey", "StableDiagnostic", "DiagnosticEvidence", "CoverageDisplayName",
            "CoverageBucketType", "CoverageStatus", "CoverageTotal", "CoverageResolved", "CoverageUnsupported",
            "CoverageUnresolved", "CoverageAmbiguous", "CoverageDiagnosticGroups", "RegisteredDefinitionName",
            "RegisteredDefinitionPrimaryEntity", "CanvasAppName", "CanvasAppId", "UniqueCanvasAppId",
            "CanvasAppDisplayName", "CanvasAppComponentState", "CanvasAppIsManaged", "CanvasAppCandidateStatus"
        };

        public string CreateCsv(MembershipComparisonPresentation presentation,
            string sourceSolutionVersion = null, string targetSolutionVersion = null)
        {
            if (presentation == null) throw new ArgumentNullException(nameof(presentation));
            var rows = new List<IDictionary<string, string>> { CreateComparisonSummary(presentation) };
            AppendEnvironment(rows, "Source", presentation.Source, sourceSolutionVersion);
            AppendEnvironment(rows, "Target", presentation.Target, targetSolutionVersion);

            var text = new StringBuilder();
            AppendCsvLine(text, Headers);
            foreach (var row in rows)
                AppendCsvLine(text, Headers.Select(header => Value(row, header)));
            return text.ToString();
        }

        public void WriteCsv(string path, MembershipComparisonPresentation presentation,
            string sourceSolutionVersion = null, string targetSolutionVersion = null)
        {
            if (string.IsNullOrWhiteSpace(path))
                throw new ArgumentException("An export path is required.", nameof(path));
            File.WriteAllText(path, CreateCsv(presentation, sourceSolutionVersion, targetSolutionVersion),
                new UTF8Encoding(true));
        }

        private static IDictionary<string, string> CreateComparisonSummary(
            MembershipComparisonPresentation presentation)
        {
            return Row(
                Pair("RowType", "ComparisonSummary"),
                Pair("SolutionUniqueName", presentation.SolutionUniqueName),
                Pair("ComparisonPresentInBoth", Number(presentation.Summary.PresentInBoth)),
                Pair("ComparisonSourceOnly", Number(presentation.Summary.SourceOnly)),
                Pair("ComparisonTargetOnly", Number(presentation.Summary.TargetOnly)),
                Pair("ComparisonUnsupported", Number(presentation.Summary.Unsupported)),
                Pair("ComparisonUnresolved", Number(presentation.Summary.Unresolved)),
                Pair("ComparisonAmbiguous", Number(presentation.Summary.Ambiguous)));
        }

        private static void AppendEnvironment(ICollection<IDictionary<string, string>> rows, string side,
            MembershipEnvironmentResult result, string solutionVersion)
        {
            var snapshot = result.Snapshot;
            var environmentName = result.Diagnostics.EnvironmentName;
            var common = new[]
            {
                Pair("Side", side), Pair("Environment", environmentName),
                Pair("EnvironmentOrganizationId", snapshot == null ? null : snapshot.Environment.OrganizationId.ToString("D")),
                Pair("SolutionUniqueName", result.SolutionUniqueName),
                Pair("SolutionId", snapshot?.Solution == null ? null : snapshot.Solution.SolutionId.ToString("D")),
                Pair("SolutionVersion", solutionVersion),
                Pair("CaptureTimestampUtc", snapshot == null ? null : snapshot.CapturedAt.UtcDateTime.ToString("o", CultureInfo.InvariantCulture)),
                Pair("SnapshotState", result.State.ToString())
            };

            var operation = Row(common.Concat(new[]
            {
                Pair("RowType", "OperationSummary"), Pair("OperationDiagnostic", result.Diagnostics.Diagnostic),
                Pair("RequestCount", Number(result.Diagnostics.RequestCount)),
                Pair("ElapsedMilliseconds", result.Diagnostics.Elapsed.TotalMilliseconds.ToString("0.###", CultureInfo.InvariantCulture)),
                Pair("RawMembershipCount", NullableNumber(result.Diagnostics.RawMembershipCount)),
                Pair("ResolvedCount", NullableNumber(result.Diagnostics.ResolvedCount)),
                Pair("UnsupportedCount", NullableNumber(result.Diagnostics.UnsupportedCount)),
                Pair("UnresolvedCount", NullableNumber(result.Diagnostics.UnresolvedCount)),
                Pair("AmbiguousCount", NullableNumber(result.Diagnostics.AmbiguousCount)),
                Pair("DiagnosticEvidence", snapshot == null ? null : string.Join(Environment.NewLine,
                    snapshot.Components.SelectMany(item => item.DiagnosticEvidence)
                        .Where(IsOperationSummaryEvidence).Distinct(StringComparer.Ordinal)))
            }).ToArray());
            rows.Add(operation);

            var coverage = snapshot == null
                ? new MembershipCoverageDiagnosticsBuilder().BuildUnavailable()
                : new MembershipCoverageDiagnosticsBuilder().Build(snapshot);
            foreach (var bucket in coverage.SemanticKinds.Concat(new[] { coverage.BroadUnclassifiable }))
            {
                rows.Add(Row(common.Concat(new[]
                {
                    Pair("RowType", "CoverageBucket"), Pair("SemanticKind", bucket.SemanticKind),
                    Pair("CoverageDisplayName", bucket.DisplayName),
                    Pair("CoverageBucketType", bucket.BucketType.ToString()),
                    Pair("CoverageStatus", bucket.CoverageStatus.ToString()),
                    Pair("CoverageTotal", Number(bucket.TotalCandidates)), Pair("CoverageResolved", Number(bucket.Resolved)),
                    Pair("CoverageUnsupported", Number(bucket.Unsupported)), Pair("CoverageUnresolved", Number(bucket.Unresolved)),
                    Pair("CoverageAmbiguous", Number(bucket.Ambiguous)),
                    Pair("CoverageDiagnosticGroups", FormatDiagnosticGroups(bucket.DiagnosticGroups))
                }).ToArray()));
            }

            if (snapshot == null) return;
            foreach (var identity in snapshot.Components.OrderBy(item => item.Record.ComponentType)
                .ThenBy(item => item.Record.SolutionComponentId))
            {
                var record = identity.Record;
                var canvas = record.ComponentType == 300
                    ? ReadCanvasAppEvidence(identity.DiagnosticEvidence) : new CanvasAppEvidence();
                rows.Add(Row(common.Concat(new[]
                {
                    Pair("RowType", "Component"), Pair("SolutionComponentId", record.SolutionComponentId.ToString("D")),
                    Pair("ObjectId", record.ObjectId?.ToString("D")), Pair("RawComponentType", Number(record.ComponentType)),
                    Pair("RootComponentBehavior", NullableNumber(record.RootComponentBehavior)),
                    Pair("RootSolutionComponentId", record.RootSolutionComponentId?.ToString("D")),
                    Pair("IsMetadata", NullableBoolean(record.IsMetadata)), Pair("SemanticKind", identity.SemanticKind),
                    Pair("ComponentTypeKey", identity.ComponentTypeKey), Pair("ResolutionStatus", identity.Status.ToString()),
                    Pair("ComparisonKey", identity.ComparisonKey), Pair("StableDiagnostic", identity.Diagnostic),
                    Pair("DiagnosticEvidence", string.Join(Environment.NewLine, identity.DiagnosticEvidence)),
                    Pair("RegisteredDefinitionName", identity.RegisteredDefinition?.Name),
                    Pair("RegisteredDefinitionPrimaryEntity", identity.RegisteredDefinition?.PrimaryEntityName),
                    Pair("CanvasAppName", canvas.Name), Pair("CanvasAppId", canvas.CanvasAppId),
                    Pair("UniqueCanvasAppId", canvas.UniqueCanvasAppId), Pair("CanvasAppDisplayName", canvas.DisplayName),
                    Pair("CanvasAppComponentState", canvas.ComponentState), Pair("CanvasAppIsManaged", canvas.IsManaged),
                    Pair("CanvasAppCandidateStatus", canvas.CandidateStatus)
                }).ToArray()));
            }
        }

        private static bool IsOperationSummaryEvidence(string value) => value != null &&
            value.IndexOf(" diagnostic summary:", StringComparison.Ordinal) >= 0;

        private static string FormatDiagnosticGroups(IEnumerable<MembershipCoverageDiagnosticGroup> groups) =>
            string.Join(Environment.NewLine, groups.Select(group => group.ResolutionStatus + " x" +
                group.Count.ToString(CultureInfo.InvariantCulture) + ": " + group.Diagnostic));

        private static CanvasAppEvidence ReadCanvasAppEvidence(IEnumerable<string> evidence)
        {
            var items = evidence == null ? new List<string>() : evidence.ToList();
            var lookup = items.Where(item => item != null &&
                (item.StartsWith("Canvas App diagnostic lookup matched. ", StringComparison.Ordinal) ||
                 item.StartsWith("Canvas App diagnostic lookup matched but returned incomplete data. ",
                     StringComparison.Ordinal))).ToList();
            var result = new CanvasAppEvidence();
            if (lookup.Count == 1)
            {
                result.CanvasAppId = ReadValue(lookup[0], "canvasappid=", "; name=");
                result.Name = ReadValue(lookup[0], "name=", "; displayname=");
                result.DisplayName = ReadValue(lookup[0], "displayname=", "; uniquecanvasappid=");
                result.UniqueCanvasAppId = ReadValue(lookup[0], "uniquecanvasappid=", "; componentstate=");
                result.ComponentState = ReadValue(lookup[0], "componentstate=", "; ismanaged=");
                result.IsManaged = ReadValue(lookup[0], "ismanaged=", "; candidateportableidentity=");
            }
            var candidate = items.FirstOrDefault(item => item != null &&
                item.StartsWith("Canvas App lifecycle candidate status=", StringComparison.Ordinal));
            if (candidate != null)
                result.CandidateStatus = ReadValue(candidate, "Canvas App lifecycle candidate status=", "; candidate=");
            return result;
        }

        private static string ReadValue(string text, string prefix, string suffix)
        {
            int start = text.IndexOf(prefix, StringComparison.Ordinal);
            if (start < 0) return null;
            start += prefix.Length;
            int end = text.IndexOf(suffix, start, StringComparison.Ordinal);
            if (end < 0) return null;
            var value = text.Substring(start, end - start).Trim();
            if (value.Length >= 2 && value[0] == '\'' && value[value.Length - 1] == '\'')
                value = UnescapeDiagnosticText(value.Substring(1, value.Length - 2));
            return value == "(not supplied)" || value == "(null)" || value == "(unavailable)" ? null : value;
        }

        private static string UnescapeDiagnosticText(string value)
        {
            var result = new StringBuilder();
            bool escaped = false;
            foreach (var character in value)
            {
                if (!escaped && character == '\\') { escaped = true; continue; }
                if (escaped)
                {
                    switch (character)
                    {
                        case 'r': result.Append('\r'); break;
                        case 'n': result.Append('\n'); break;
                        case 't': result.Append('\t'); break;
                        default: result.Append(character); break;
                    }
                    escaped = false;
                }
                else result.Append(character);
            }
            if (escaped) result.Append('\\');
            return result.ToString();
        }

        private static IDictionary<string, string> Row(params KeyValuePair<string, string>[] values) =>
            values.ToDictionary(item => item.Key, item => item.Value ?? string.Empty, StringComparer.Ordinal);

        private static KeyValuePair<string, string> Pair(string name, string value) =>
            new KeyValuePair<string, string>(name, value);

        private static string Value(IDictionary<string, string> row, string name)
        {
            string value;
            return row.TryGetValue(name, out value) ? value : string.Empty;
        }

        private static string Number(int value) => value.ToString(CultureInfo.InvariantCulture);
        private static string NullableNumber(int? value) => value.HasValue ? Number(value.Value) : null;
        private static string NullableBoolean(bool? value) => value.HasValue ? value.Value.ToString() : null;

        private static void AppendCsvLine(StringBuilder text, IEnumerable<string> values)
        {
            text.Append(string.Join(",", values.Select(EscapeCsv))).Append("\r\n");
        }

        private static string EscapeCsv(string value)
        {
            value = value ?? string.Empty;
            return value.IndexOfAny(new[] { ',', '"', '\r', '\n' }) < 0
                ? value : "\"" + value.Replace("\"", "\"\"") + "\"";
        }

        private sealed class CanvasAppEvidence
        {
            public string Name { get; set; }
            public string CanvasAppId { get; set; }
            public string UniqueCanvasAppId { get; set; }
            public string DisplayName { get; set; }
            public string ComponentState { get; set; }
            public string IsManaged { get; set; }
            public string CandidateStatus { get; set; }
        }
    }
}
