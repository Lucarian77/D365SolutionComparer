using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;

namespace D365SolutionComparer.Services.Membership
{
    public sealed class CanvasAppLifecycleChronologyException : InvalidOperationException
    {
        public CanvasAppLifecycleChronologyException(string environmentName,
            DateTimeOffset beforeCaptureUtc, DateTimeOffset afterCaptureUtc)
            : base(CreateMessage(environmentName, beforeCaptureUtc, afterCaptureUtc))
        {
            EnvironmentName = environmentName ?? string.Empty;
            BeforeCaptureUtc = beforeCaptureUtc.ToUniversalTime();
            AfterCaptureUtc = afterCaptureUtc.ToUniversalTime();
        }

        public string EnvironmentName { get; }
        public DateTimeOffset BeforeCaptureUtc { get; }
        public DateTimeOffset AfterCaptureUtc { get; }

        private static string CreateMessage(string environmentName, DateTimeOffset before, DateTimeOffset after) =>
            "The selected lifecycle snapshots appear to be reversed." + Environment.NewLine + Environment.NewLine +
            "The Before snapshot was captured after the After snapshot." + Environment.NewLine + Environment.NewLine +
            "Environment: " + (environmentName ?? string.Empty) + Environment.NewLine +
            "Before: " + FormatUtc(before) + Environment.NewLine +
            "After: " + FormatUtc(after) + Environment.NewLine + Environment.NewLine +
            "Select the earlier lifecycle export as Before and the later export as After.";

        private static string FormatUtc(DateTimeOffset value) =>
            value.UtcDateTime.ToString("o", CultureInfo.InvariantCulture);
    }

    /// <summary>Compares Phase 2F.9 Canvas App evidence without producing membership identities or results.</summary>
    public sealed class CanvasAppLifecycleCsvComparer
    {
        private static readonly string[] RequiredInputHeaders =
        {
            "RowType", "LifecycleOperation", "Side", "Environment", "EnvironmentOrganizationId",
            "SolutionUniqueName", "SolutionId", "SolutionVersion", "CaptureTimestampUtc",
            "SolutionComponentId", "ObjectId", "RawComponentType", "SemanticKind", "ComponentTypeKey",
            "ResolutionStatus", "ComparisonKey", "StableDiagnostic", "DiagnosticEvidence", "CanvasAppName",
            "CanvasAppId", "UniqueCanvasAppId", "CanvasAppDisplayName", "CanvasAppComponentState",
            "CanvasAppIsManaged", "CanvasAppCandidateStatus", "CanvasAppCandidateDiagnostic"
        };

        private static readonly string[] EvidenceFields =
        {
            "Side", "Environment", "EnvironmentOrganizationId", "SolutionUniqueName", "SolutionId",
            "SolutionVersion", "CaptureTimestampUtc", "LifecycleOperation", "RawComponentType",
            "SolutionComponentId", "ObjectId", "CanvasAppId", "CanvasAppName", "CanvasAppDisplayName",
            "UniqueCanvasAppId", "CanvasAppComponentState", "CanvasAppIsManaged", "CanvasAppCandidateStatus",
            "CanvasAppCandidateDiagnostic", "SemanticKind", "ComponentTypeKey", "ResolutionStatus",
            "ComparisonKey", "StableDiagnostic", "DiagnosticEvidence"
        };

        private static readonly string[] OutputHeaders = new[]
        {
            "Outcome", "CorrelationBasis", "Diagnostic", "CandidateNameEqualOrdinalIgnoreCase",
            "AuditDifferences"
        }.Concat(EvidenceFields.Select(field => "Before" + field))
            .Concat(EvidenceFields.Select(field => "After" + field)).ToArray();

        public string Compare(string beforeCsv, string afterCsv)
        {
            var before = Parse(beforeCsv, "Before");
            var after = Parse(afterCsv, "After");
            ValidateLifecycleOperation(before.AllRows, after.AllRows);
            ValidateChronology(before.CanvasApps, after.CanvasApps);

            var output = new List<ComparisonRow>();
            var partitionKeys = before.CanvasApps.Select(item => item.PartitionKey)
                .Union(after.CanvasApps.Select(item => item.PartitionKey), StringComparer.OrdinalIgnoreCase)
                .OrderBy(item => item, StringComparer.OrdinalIgnoreCase).ToList();
            foreach (var partitionKey in partitionKeys)
                ComparePartition(before.CanvasApps.Where(item => Same(item.PartitionKey, partitionKey)).ToList(),
                    after.CanvasApps.Where(item => Same(item.PartitionKey, partitionKey)).ToList(), output);

            var ordered = output.OrderBy(item => item.PartitionKey, StringComparer.OrdinalIgnoreCase)
                .ThenBy(item => OutcomeOrder(item.Outcome))
                .ThenBy(item => item.SortCandidate, StringComparer.OrdinalIgnoreCase)
                .ThenBy(item => item.Before?.SortId, StringComparer.OrdinalIgnoreCase)
                .ThenBy(item => item.After?.SortId, StringComparer.OrdinalIgnoreCase).ToList();
            var text = new StringBuilder();
            AppendCsvLine(text, OutputHeaders);
            foreach (var row in ordered) AppendCsvLine(text, OutputHeaders.Select(header => row.Value(header)));
            return text.ToString();
        }

        private static void ValidateChronology(IReadOnlyList<EvidenceRecord> before,
            IReadOnlyList<EvidenceRecord> after)
        {
            var matchingPartitions = before.Select(item => item.PartitionKey)
                .Intersect(after.Select(item => item.PartitionKey), StringComparer.OrdinalIgnoreCase)
                .OrderBy(item => item, StringComparer.OrdinalIgnoreCase).ToList();
            foreach (var partition in matchingPartitions)
            {
                var beforeRows = before.Where(item => Same(item.PartitionKey, partition)).ToList();
                var afterRows = after.Where(item => Same(item.PartitionKey, partition)).ToList();
                var beforeCapture = ReadCaptureTimestamp(beforeRows, "Before");
                var afterCapture = ReadCaptureTimestamp(afterRows, "After");
                if (beforeCapture > afterCapture)
                    throw new CanvasAppLifecycleChronologyException(
                        DisplayEnvironment(beforeRows, afterRows), beforeCapture, afterCapture);
            }
        }

        private static DateTimeOffset ReadCaptureTimestamp(IReadOnlyList<EvidenceRecord> records, string label)
        {
            var parsed = new List<DateTimeOffset>();
            foreach (var record in records)
            {
                DateTimeOffset timestamp;
                if (!DateTimeOffset.TryParseExact(record.Get("CaptureTimestampUtc"), "o",
                    CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind, out timestamp))
                    throw new InvalidDataException(label + " CSV has a missing or invalid CaptureTimestampUtc for " +
                        "environment '" + record.Get("Environment") + "' and solution '" +
                        record.Get("SolutionUniqueName") + "'.");
                parsed.Add(timestamp.ToUniversalTime());
            }
            var distinct = parsed.Distinct().ToList();
            if (distinct.Count != 1)
                throw new InvalidDataException(label + " CSV contains conflicting capture timestamps for " +
                    "environment '" + records[0].Get("Environment") + "' and solution '" +
                    records[0].Get("SolutionUniqueName") + "'.");
            return distinct[0];
        }

        private static string DisplayEnvironment(IReadOnlyList<EvidenceRecord> before,
            IReadOnlyList<EvidenceRecord> after)
        {
            var name = before.Concat(after).Select(item => item.Get("Environment"))
                .FirstOrDefault(item => !string.IsNullOrWhiteSpace(item));
            return name ?? string.Empty;
        }

        public void CompareFiles(string beforePath, string afterPath, string outputPath)
        {
            if (string.IsNullOrWhiteSpace(beforePath))
                throw new ArgumentException("A Before CSV path is required.", nameof(beforePath));
            if (string.IsNullOrWhiteSpace(afterPath))
                throw new ArgumentException("An After CSV path is required.", nameof(afterPath));
            if (string.IsNullOrWhiteSpace(outputPath))
                throw new ArgumentException("An output CSV path is required.", nameof(outputPath));
            var result = Compare(File.ReadAllText(beforePath), File.ReadAllText(afterPath));
            File.WriteAllText(outputPath, result, new UTF8Encoding(true));
        }

        private static void ComparePartition(IReadOnlyList<EvidenceRecord> before,
            IReadOnlyList<EvidenceRecord> after, ICollection<ComparisonRow> output)
        {
            var invalidBefore = before.Where(item => !item.HasUsableCandidate).ToList();
            var invalidAfter = after.Where(item => !item.HasUsableCandidate).ToList();
            foreach (var item in invalidBefore)
                output.Add(ComparisonRow.OneSided("InsufficientEvidence", "CandidateEvidenceIncomplete",
                    "The Before record does not contain a valid Canvas App candidate name.", item, null));
            foreach (var item in invalidAfter)
                output.Add(ComparisonRow.OneSided("InsufficientEvidence", "CandidateEvidenceIncomplete",
                    "The After record does not contain a valid Canvas App candidate name.", null, item));

            var beforeCandidates = before.Where(item => item.HasUsableCandidate).ToList();
            var afterCandidates = after.Where(item => item.HasUsableCandidate).ToList();
            var beforeGroups = beforeCandidates.GroupBy(item => item.CandidateName,
                StringComparer.OrdinalIgnoreCase).ToDictionary(group => group.Key, group => group.ToList(),
                StringComparer.OrdinalIgnoreCase);
            var afterGroups = afterCandidates.GroupBy(item => item.CandidateName,
                StringComparer.OrdinalIgnoreCase).ToDictionary(group => group.Key, group => group.ToList(),
                StringComparer.OrdinalIgnoreCase);
            var candidateNames = beforeGroups.Keys.Union(afterGroups.Keys, StringComparer.OrdinalIgnoreCase)
                .OrderBy(item => item, StringComparer.OrdinalIgnoreCase).ToList();
            var unmatchedBefore = new List<EvidenceRecord>();
            var unmatchedAfter = new List<EvidenceRecord>();

            foreach (var candidateName in candidateNames)
            {
                List<EvidenceRecord> beforeGroup;
                List<EvidenceRecord> afterGroup;
                beforeGroups.TryGetValue(candidateName, out beforeGroup);
                afterGroups.TryGetValue(candidateName, out afterGroup);
                beforeGroup = beforeGroup ?? new List<EvidenceRecord>();
                afterGroup = afterGroup ?? new List<EvidenceRecord>();
                if (DistinctCandidateObjects(beforeGroup) > 1 || DistinctCandidateObjects(afterGroup) > 1)
                {
                    foreach (var item in beforeGroup)
                        output.Add(ComparisonRow.OneSided("DuplicateCandidate",
                            "DuplicateCaseInsensitiveCanvasAppName",
                            "The candidate name is not unique within the Before or After evidence partition.",
                            item, null));
                    foreach (var item in afterGroup)
                        output.Add(ComparisonRow.OneSided("DuplicateCandidate",
                            "DuplicateCaseInsensitiveCanvasAppName",
                            "The candidate name is not unique within the Before or After evidence partition.",
                            null, item));
                }
                else
                {
                    var beforeRepresentative = Representative(beforeGroup);
                    var afterRepresentative = Representative(afterGroup);
                    AddRepeatedMembershipEvidence(beforeGroup, beforeRepresentative, true, output);
                    AddRepeatedMembershipEvidence(afterGroup, afterRepresentative, false, output);
                    if (beforeRepresentative != null && afterRepresentative != null)
                        output.Add(ComparisonRow.Pair("CandidateStable", "CaseInsensitiveCanvasAppName",
                            "The nonblank candidate name is equal using StringComparer.OrdinalIgnoreCase.",
                            beforeRepresentative, afterRepresentative));
                    else if (beforeRepresentative != null) unmatchedBefore.Add(beforeRepresentative);
                    else if (afterRepresentative != null) unmatchedAfter.Add(afterRepresentative);
                }
            }

            if (unmatchedBefore.Count == 1 && unmatchedAfter.Count == 1)
            {
                output.Add(ComparisonRow.Pair("CandidateChanged", "SingleRemainingCandidateWithinPartition",
                    "One usable candidate remains on each side, but their names differ case-insensitively.",
                    unmatchedBefore[0], unmatchedAfter[0]));
                return;
            }
            if (unmatchedBefore.Count > 0 && unmatchedAfter.Count > 0)
            {
                foreach (var item in unmatchedBefore)
                    output.Add(ComparisonRow.OneSided("AmbiguousCorrelation", "MultipleUnmatchedCandidates",
                        "Multiple unmatched candidates prevent a deterministic Before/After pairing.", item, null));
                foreach (var item in unmatchedAfter)
                    output.Add(ComparisonRow.OneSided("AmbiguousCorrelation", "MultipleUnmatchedCandidates",
                        "Multiple unmatched candidates prevent a deterministic Before/After pairing.", null, item));
                return;
            }
            foreach (var item in unmatchedBefore)
                output.Add(ComparisonRow.OneSided("BeforeOnly", "NoAfterCandidate",
                    "No After candidate with the same case-insensitive name is available.", item, null));
            foreach (var item in unmatchedAfter)
                output.Add(ComparisonRow.OneSided("AfterOnly", "NoBeforeCandidate",
                    "No Before candidate with the same case-insensitive name is available.", null, item));
        }

        private static int DistinctCandidateObjects(IEnumerable<EvidenceRecord> records) => records
            .Select(item => item.CandidateObjectKey).Distinct(StringComparer.OrdinalIgnoreCase).Count();

        private static EvidenceRecord Representative(IEnumerable<EvidenceRecord> records) => records
            .OrderBy(item => item.SortId, StringComparer.OrdinalIgnoreCase).FirstOrDefault();

        private static void AddRepeatedMembershipEvidence(IEnumerable<EvidenceRecord> records,
            EvidenceRecord representative, bool isBefore, ICollection<ComparisonRow> output)
        {
            if (representative == null) return;
            foreach (var repeated in records.Where(item => !ReferenceEquals(item, representative)))
                output.Add(ComparisonRow.OneSided("RepeatedMembershipEvidence", "RepeatedRawObjectId",
                    "This row repeats the same environment-local objectid and candidate evidence within one snapshot.",
                    isBefore ? repeated : null, isBefore ? null : repeated));
        }

        private static ParsedCsv Parse(string csv, string label)
        {
            if (string.IsNullOrWhiteSpace(csv)) throw new InvalidDataException(label + " CSV is empty.");
            var records = ParseRecords(csv, label);
            if (records.Count == 0) throw new InvalidDataException(label + " CSV has no header row.");
            var headers = records[0];
            if (headers.Count > 0) headers[0] = headers[0].TrimStart('\uFEFF');
            if (headers.Any(string.IsNullOrEmpty) || headers.Distinct(StringComparer.Ordinal).Count() != headers.Count)
                throw new InvalidDataException(label + " CSV contains blank or duplicate headers.");
            var missing = RequiredInputHeaders.Where(required => !headers.Contains(required,
                StringComparer.Ordinal)).ToList();
            if (missing.Count > 0)
                throw new InvalidDataException(label + " CSV is missing required headers: " +
                    string.Join(", ", missing) + ".");

            var rows = new List<IDictionary<string, string>>();
            for (int index = 1; index < records.Count; index++)
            {
                if (records[index].Count != headers.Count)
                    throw new InvalidDataException(label + " CSV row " + (index + 1).ToString(
                        CultureInfo.InvariantCulture) + " has an unexpected field count.");
                rows.Add(headers.Select((header, column) => new { header, value = records[index][column] })
                    .ToDictionary(item => item.header, item => item.value, StringComparer.Ordinal));
            }
            var canvasApps = rows.Where(row => Same(Value(row, "RowType"), "Component") &&
                    Same(Value(row, "RawComponentType"), "300"))
                .Select(row => new EvidenceRecord(row)).ToList();
            return new ParsedCsv(rows, canvasApps);
        }

        private static IReadOnlyList<List<string>> ParseRecords(string csv, string label)
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
                else if (character == '"')
                {
                    if (field.Length != 0) throw new InvalidDataException(label + " CSV contains an invalid quote.");
                    quoted = true;
                }
                else if (character == ',') { record.Add(field.ToString()); field.Clear(); }
                else if (character == '\r' || character == '\n')
                {
                    if (character == '\r' && index + 1 < csv.Length && csv[index + 1] == '\n') index++;
                    record.Add(field.ToString()); field.Clear(); records.Add(record); record = new List<string>();
                }
                else field.Append(character);
            }
            if (quoted) throw new InvalidDataException(label + " CSV contains an unterminated quoted field.");
            if (field.Length > 0 || record.Count > 0)
            {
                record.Add(field.ToString());
                records.Add(record);
            }
            return records;
        }

        private static void ValidateLifecycleOperation(IReadOnlyList<IDictionary<string, string>> before,
            IReadOnlyList<IDictionary<string, string>> after)
        {
            var beforeOperation = SingleOperation(before, "Before");
            var afterOperation = SingleOperation(after, "After");
            if (!Same(beforeOperation, afterOperation))
                throw new InvalidDataException("Before and After lifecycle operations do not match: '" +
                    beforeOperation + "' and '" + afterOperation + "'.");
        }

        private static string SingleOperation(IEnumerable<IDictionary<string, string>> rows, string label)
        {
            var values = rows.Select(row => Value(row, "LifecycleOperation").Trim())
                .Where(value => value.Length > 0).Distinct(StringComparer.OrdinalIgnoreCase).ToList();
            if (values.Count != 1)
                throw new InvalidDataException(label + " CSV must contain exactly one nonblank lifecycle operation.");
            return values[0];
        }

        private static int OutcomeOrder(string outcome)
        {
            switch (outcome)
            {
                case "CandidateStable": return 0;
                case "CandidateChanged": return 1;
                case "DuplicateCandidate": return 2;
                case "AmbiguousCorrelation": return 3;
                case "InsufficientEvidence": return 4;
                case "BeforeOnly": return 5;
                default: return 6;
            }
        }

        private static string Value(IDictionary<string, string> row, string name)
        {
            string value;
            return row.TryGetValue(name, out value) ? value ?? string.Empty : string.Empty;
        }

        private static bool Same(string left, string right) =>
            string.Equals(left, right, StringComparison.OrdinalIgnoreCase);

        private static void AppendCsvLine(StringBuilder text, IEnumerable<string> values) =>
            text.Append(string.Join(",", values.Select(EscapeCsv))).Append("\r\n");

        private static string EscapeCsv(string value)
        {
            value = value ?? string.Empty;
            return value.IndexOfAny(new[] { ',', '"', '\r', '\n' }) < 0
                ? value : "\"" + value.Replace("\"", "\"\"") + "\"";
        }

        private sealed class ParsedCsv
        {
            public ParsedCsv(IReadOnlyList<IDictionary<string, string>> allRows,
                IReadOnlyList<EvidenceRecord> canvasApps)
            {
                AllRows = allRows;
                CanvasApps = canvasApps;
            }
            public IReadOnlyList<IDictionary<string, string>> AllRows { get; }
            public IReadOnlyList<EvidenceRecord> CanvasApps { get; }
        }

        private sealed class EvidenceRecord
        {
            private readonly IDictionary<string, string> values;
            public EvidenceRecord(IDictionary<string, string> values)
            {
                this.values = values;
                CandidateName = Value(values, "CanvasAppName").Trim();
                var environmentId = Value(values, "EnvironmentOrganizationId").Trim();
                var environmentName = Value(values, "Environment").Trim();
                var solutionName = Value(values, "SolutionUniqueName").Trim();
                HasUsableCandidate = CandidateName.Length > 0 &&
                    Same(Value(values, "CanvasAppCandidateStatus"), "CandidateValid") &&
                    Value(values, "CanvasAppCandidateDiagnostic").Trim().Length > 0 &&
                    Value(values, "ObjectId").Trim().Length > 0 &&
                    Value(values, "CanvasAppId").Trim().Length > 0 &&
                    Value(values, "UniqueCanvasAppId").Trim().Length > 0 &&
                    Value(values, "CanvasAppDisplayName").Trim().Length > 0 &&
                    Value(values, "CanvasAppComponentState").Trim().Length > 0 &&
                    Value(values, "CanvasAppIsManaged").Trim().Length > 0 &&
                    Same(Value(values, "ResolutionStatus"), "Unsupported") &&
                    Value(values, "ComparisonKey").Length == 0 &&
                    Same(Value(values, "SemanticKind"), "unsupported:componenttype:300") &&
                    (environmentId.Length > 0 || environmentName.Length > 0) && solutionName.Length > 0;
                var environmentKey = environmentId;
                if (environmentKey.Length == 0) environmentKey = "name:" + environmentName;
                else environmentKey = "id:" + environmentKey;
                PartitionKey = environmentKey + "|solution:" + solutionName;
                SortId = Value(values, "ObjectId") + "|" + Value(values, "CanvasAppId") + "|" +
                    Value(values, "SolutionComponentId");
                CandidateObjectKey = Value(values, "ObjectId").Trim();
                if (CandidateObjectKey.Length == 0)
                    CandidateObjectKey = "solutioncomponent:" + Value(values, "SolutionComponentId").Trim();
            }
            public string CandidateName { get; }
            public bool HasUsableCandidate { get; }
            public string PartitionKey { get; }
            public string SortId { get; }
            public string CandidateObjectKey { get; }
            public string Get(string field) => Value(values, field);
        }

        private sealed class ComparisonRow
        {
            private ComparisonRow(string outcome, string basis, string diagnostic,
                EvidenceRecord before, EvidenceRecord after)
            {
                Outcome = outcome;
                CorrelationBasis = basis;
                Diagnostic = diagnostic;
                Before = before;
                After = after;
                PartitionKey = before?.PartitionKey ?? after?.PartitionKey ?? string.Empty;
                SortCandidate = before?.CandidateName ?? after?.CandidateName ?? string.Empty;
            }
            public string Outcome { get; }
            public string CorrelationBasis { get; }
            public string Diagnostic { get; }
            public EvidenceRecord Before { get; }
            public EvidenceRecord After { get; }
            public string PartitionKey { get; }
            public string SortCandidate { get; }

            public static ComparisonRow Pair(string outcome, string basis, string diagnostic,
                EvidenceRecord before, EvidenceRecord after) =>
                new ComparisonRow(outcome, basis, diagnostic, before, after);

            public static ComparisonRow OneSided(string outcome, string basis, string diagnostic,
                EvidenceRecord before, EvidenceRecord after) =>
                new ComparisonRow(outcome, basis, diagnostic, before, after);

            public string Value(string header)
            {
                switch (header)
                {
                    case "Outcome": return Outcome;
                    case "CorrelationBasis": return CorrelationBasis;
                    case "Diagnostic": return Diagnostic;
                    case "CandidateNameEqualOrdinalIgnoreCase":
                        return Before == null || After == null ? string.Empty :
                            Same(Before.CandidateName, After.CandidateName).ToString();
                    case "AuditDifferences": return AuditDifferences();
                }
                const string beforePrefix = "Before";
                const string afterPrefix = "After";
                if (header.StartsWith(beforePrefix, StringComparison.Ordinal))
                    return Before?.Get(header.Substring(beforePrefix.Length)) ?? string.Empty;
                if (header.StartsWith(afterPrefix, StringComparison.Ordinal))
                    return After?.Get(header.Substring(afterPrefix.Length)) ?? string.Empty;
                return string.Empty;
            }

            private string AuditDifferences()
            {
                if (Before == null || After == null) return string.Empty;
                var fields = new[] { "ObjectId", "CanvasAppId", "UniqueCanvasAppId", "CanvasAppDisplayName",
                    "CanvasAppComponentState", "CanvasAppIsManaged", "SolutionId", "SolutionVersion" };
                return string.Join(";", fields.Where(field => !string.Equals(Before.Get(field), After.Get(field),
                    StringComparison.Ordinal)).ToArray());
            }
        }
    }
}
