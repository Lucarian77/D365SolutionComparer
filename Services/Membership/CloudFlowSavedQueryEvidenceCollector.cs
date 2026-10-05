#if DEBUG
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Security.Cryptography;
using System.ServiceModel;
using System.Text;
using System.Threading;
using System.Xml;
using D365SolutionComparer.Models.Membership;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Metadata.Query;
using Microsoft.Xrm.Sdk.Query;

namespace D365SolutionComparer.Services.Membership
{
    /// <summary>On-demand evidence only. Does not supply identities, definitions or absence proof.</summary>
    internal sealed class CloudFlowSavedQueryEvidenceCollector
    {
        internal static readonly string[] WorkflowFields = { "workflowid", "workflowidunique", "name", "uniquename",
            "category", "type", "mode", "modernflowtype", "primaryentity", "statecode", "statuscode",
            "componentstate", "ismanaged", "parentworkflowid", "activeworkflowid", "processorder",
            "clientdata", "clientdataiscompressed", "xaml", "resourceid", "solutionid", "supportingsolutionid" };
        internal static readonly string[] QueryFields = { "savedqueryid", "savedqueryidunique", "name",
            "returnedtypecode", "querytype", "componentstate", "ismanaged", "fetchxml", "layoutxml", "columnsetxml" };

        internal ProcessQueryEvidenceReport Capture(IOrganizationService sourceService, MembershipSnapshot source,
            string sourceVersion, IOrganizationService targetService, MembershipSnapshot target, string targetVersion,
            bool savedQuery, CancellationToken token)
        {
            if (source?.State != MembershipSnapshotState.Complete || target?.State != MembershipSnapshotState.Complete ||
                !StringComparer.OrdinalIgnoreCase.Equals(source.SolutionUniqueName, target.SolutionUniqueName))
                throw new ArgumentException("Complete Source and Target snapshots of the same solution are required.");
            token.ThrowIfCancellationRequested();
            var report = new ProcessQueryEvidenceReport { SavedQuery = savedQuery };
            report.Source = Read(sourceService, source, sourceVersion, savedQuery, token);
            report.Target = Read(targetService, target, targetVersion, savedQuery, token);
            token.ThrowIfCancellationRequested();
            return report;
        }

        private static ProcessQuerySideEvidence Read(IOrganizationService service, MembershipSnapshot snapshot,
            string version, bool savedQuery, CancellationToken token)
        {
            if (service == null) throw new ArgumentNullException(nameof(service));
            var side = new ProcessQuerySideEvidence { Snapshot = snapshot, Version = version,
                CapturedUtc = DateTimeOffset.UtcNow, SavedQuery = savedQuery };
            // Production resolution is an annotation, never a filter for diagnostic capture.
            // Resolved target flows and business-process context must remain visible.
            side.Raw.AddRange(snapshot.Components.Where(c => c.Record.ComponentType == (savedQuery ? 26 : 29)));
            var ids = side.Raw.Where(c => c.Record.ObjectId.HasValue && c.Record.ObjectId != Guid.Empty)
                .Select(c => c.Record.ObjectId.Value).Distinct().OrderBy(id => id).ToList();
            foreach (var id in ids) side.Rows.Add(id, new ProcessQueryRowEvidence { ObjectId = id, Status = "Unavailable" });
            if (ids.Count == 0) return side;
            string entity = savedQuery ? "savedquery" : "workflow", primary = entity + "id";
            var desired = savedQuery ? QueryFields : WorkflowFields;
            side.Fields.UnionWith(desired);
            try
            {
                side.Requests.Add("RetrieveEntity(" + entity + ", Attributes, RetrieveAsIfPublished=False)");
                var schema = service.Execute(new RetrieveEntityRequest { LogicalName = entity,
                    EntityFilters = EntityFilters.Attributes, RetrieveAsIfPublished = false }) as RetrieveEntityResponse;
                token.ThrowIfCancellationRequested();
                var metadata = schema?.EntityMetadata;
                if (metadata?.Attributes == null || metadata.PrimaryIdAttribute != primary || metadata.LogicalName != entity ||
                    metadata.Attributes.Any(a => a == null || string.IsNullOrWhiteSpace(a.LogicalName)) ||
                    metadata.Attributes.GroupBy(a => a.LogicalName, StringComparer.OrdinalIgnoreCase).Any(g => g.Count() != 1))
                    throw new InvalidOperationException("Incomplete or conflicting entity metadata.");
                var readable = metadata.Attributes.Where(a => a.IsValidForRead == true).ToList();
                if (!savedQuery)
                    side.Fields.UnionWith(readable.Where(a => a.AttributeType == AttributeTypeCode.Uniqueidentifier ||
                        (a.AttributeType == AttributeTypeCode.Lookup &&
                            (a.LogicalName.Contains("workflow") || a.LogicalName.Contains("solution"))) ||
                        (a.LogicalName.Contains("flow") && (a.LogicalName.EndsWith("id", StringComparison.Ordinal) ||
                            a.LogicalName.EndsWith("identifier", StringComparison.Ordinal)))).Select(a => a.LogicalName));
                side.Columns.AddRange(readable.Select(a => a.LogicalName).Where(side.Fields.Contains)
                    .OrderBy(n => n, StringComparer.Ordinal));
                if (!side.Columns.Contains(primary)) throw new InvalidOperationException("Primary ID is not readable.");
            }
            catch (OperationCanceledException) { throw; }
            catch (Exception ex) when (ex is FaultException || ex is InvalidOperationException)
            {
                token.ThrowIfCancellationRequested();
                side.Diagnostic = "Metadata unavailable/faulted/incomplete; server details withheld. No guessed columns were queried.";
                return side;
            }
            for (int offset = 0; offset < ids.Count; offset += 200)
            {
                token.ThrowIfCancellationRequested();
                var batch = ids.Skip(offset).Take(200).ToList();
                var query = new QueryExpression(entity) { ColumnSet = new ColumnSet(side.Columns.ToArray()) };
                query.Criteria.AddCondition(new ConditionExpression(primary, ConditionOperator.In,
                    batch.Select(id => (object)id).ToArray()));
                side.Requests.Add("RetrieveMultiple(" + entity + ", columns=[" + string.Join(",", side.Columns) +
                    "], " + primary + " IN Guid[" + batch.Count + "])");
                try
                {
                    var response = service.RetrieveMultiple(query);
                    token.ThrowIfCancellationRequested();
                    if (response == null || response.MoreRecords)
                    { Fail(side, batch, "IncompleteBatch"); continue; }
                    side.ReturnedCount += response.Entities.Count;
                    if (response.Entities.Any(row => row == null || row.LogicalName != entity ||
                        !(row.GetAttributeValue<object>(primary) is Guid) || (Guid)row[primary] != row.Id || !batch.Contains(row.Id)))
                    { Fail(side, batch, "ConflictingPrimaryKey"); continue; }
                    foreach (var id in batch)
                    {
                        var matches = response.Entities.Where(r => r.Id == id).ToList();
                        var evidence = side.Rows[id];
                        evidence.Status = matches.Count == 1 ? "Unique" : matches.Count == 0 ? "Missing" : "DuplicateReturnedRow";
                        if (matches.Count == 1) evidence.Row = matches[0];
                    }
                }
                catch (OperationCanceledException) { throw; }
                catch (Exception ex) when (ex is FaultException || ex is InvalidOperationException)
                { token.ThrowIfCancellationRequested(); Fail(side, batch, "FaultedBatch (server details withheld)"); }
            }
            if (savedQuery) ResolveQueryScopes(service, side, token);
            token.ThrowIfCancellationRequested();
            return side;
        }

        private static void Fail(ProcessQuerySideEvidence side, IEnumerable<Guid> ids, string status)
        { foreach (var id in ids) side.Rows[id].Status = status; }

        private static void ResolveQueryScopes(IOrganizationService service, ProcessQuerySideEvidence side, CancellationToken token)
        {
            var numeric = new Dictionary<Guid, int>();
            foreach (var evidence in side.Rows.Values.Where(r => r.Row != null))
            {
                var raw = evidence.Row.GetAttributeValue<object>("returnedtypecode");
                var text = raw as string;
                int code;
                if (raw is int) numeric[evidence.ObjectId] = (int)raw;
                else if (int.TryParse(text, NumberStyles.Integer, CultureInfo.InvariantCulture, out code)) numeric[evidence.ObjectId] = code;
                else if (LogicalName(text?.Trim())) evidence.Scope = text.Trim();
            }
            var codes = numeric.Values.Distinct().OrderBy(c => c).ToList();
            for (int offset = 0; offset < codes.Count; offset += 200)
            {
                var batch = codes.Skip(offset).Take(200).ToList();
                var query = new EntityQueryExpression { Properties = new MetadataPropertiesExpression("ObjectTypeCode", "LogicalName"),
                    Criteria = new MetadataFilterExpression(LogicalOperator.Or) };
                foreach (var code in batch) query.Criteria.Conditions.Add(new MetadataConditionExpression("ObjectTypeCode", MetadataConditionOperator.Equals, code));
                token.ThrowIfCancellationRequested();
                side.Requests.Add("RetrieveMetadataChanges(ObjectTypeCode IN [" + string.Join(",", batch) + "], ObjectTypeCode,LogicalName)");
                try
                {
                    var response = service.Execute(new RetrieveMetadataChangesRequest { Query = query }) as RetrieveMetadataChangesResponse;
                    token.ThrowIfCancellationRequested();
                    if (response?.EntityMetadata == null || response.EntityMetadata.Any(m => m == null ||
                        !m.ObjectTypeCode.HasValue || !batch.Contains(m.ObjectTypeCode.Value))) continue;
                    foreach (var code in batch)
                    {
                        var matches = response.EntityMetadata.Where(m => m.ObjectTypeCode == code).ToList();
                        if (matches.Count == 1 && LogicalName(matches[0].LogicalName?.Trim()))
                            foreach (var id in numeric.Where(n => n.Value == code).Select(n => n.Key)) side.Rows[id].Scope = matches[0].LogicalName.Trim();
                    }
                }
                catch (OperationCanceledException) { throw; }
                catch (FaultException) { token.ThrowIfCancellationRequested(); side.Diagnostic = "Entity scope metadata faulted; affected candidates remain incomplete."; }
            }
        }

        private static bool LogicalName(string value) => !string.IsNullOrWhiteSpace(value) &&
            !value.Equals("none", StringComparison.OrdinalIgnoreCase) && (char.IsLetter(value[0]) || value[0] == '_') &&
            value.All(c => char.IsLetterOrDigit(c) || c == '_');

        internal static string Text(Entity row, string field) => (row?.GetAttributeValue<object>(field) as string)?.Trim();
        internal static string Frame(params string[] fields) => string.Concat(fields.Select(f => f.Length.ToString(CultureInfo.InvariantCulture) + ":" + f + ":"));
        internal static string Safe(string value) => (value ?? "(not returned/null)").Replace("\r", "\\r").Replace("\n", "\\n").Replace("\t", "\\t");
        internal static string Value(Entity row, string field)
        {
            var value = row?.GetAttributeValue<object>(field);
            if (value == null) return null;
            if (field == "clientdata" || field == "xaml" || field.EndsWith("xml", StringComparison.Ordinal))
            {
                var text = value as string;
                if (text == null) return "UnexpectedValueType";
                string canonical = null;
                if (field.EndsWith("xml", StringComparison.Ordinal) && !string.IsNullOrWhiteSpace(text))
                {
                    try
                    {
                        var doc = new XmlDocument { XmlResolver = null };
                        using (var input = new System.IO.StringReader(text))
                        using (var reader = XmlReader.Create(input, new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null })) doc.Load(reader);
                        canonical = doc.OuterXml;
                    }
                    catch (XmlException) { }
                }
                return "present=" + !string.IsNullOrWhiteSpace(text) + "; length=" + text.Length + "; rawSha256=" + Hash(text) +
                    (field.EndsWith("xml", StringComparison.Ordinal) ? "; canonicalXmlSha256=" + (canonical == null ? "Unavailable" : Hash(canonical)) : "");
            }
            if (value is OptionSetValue) return ((OptionSetValue)value).Value.ToString(CultureInfo.InvariantCulture);
            if (value is EntityReference) return ((EntityReference)value).LogicalName + ":" + ((EntityReference)value).Id.ToString("D");
            if (value is Guid) return ((Guid)value).ToString("D");
            if (value is string || value is int || value is bool) return Convert.ToString(value, CultureInfo.InvariantCulture);
            return "UnexpectedValueType";
        }
        private static string Hash(string value)
        { using (var sha = SHA256.Create()) return BitConverter.ToString(sha.ComputeHash(Encoding.UTF8.GetBytes(value))).Replace("-", ""); }
    }

    internal sealed class ProcessQueryRowEvidence
    {
        internal Guid ObjectId;
        internal string Status, Scope;
        internal Entity Row;
        internal string Candidate(bool a)
        {
            var name = CloudFlowSavedQueryEvidenceCollector.Text(Row, "name");
            if (Status != "Unique" || string.IsNullOrWhiteSpace(Scope) || string.IsNullOrWhiteSpace(name)) return null;
            var queryType = Row.GetAttributeValue<object>("querytype");
            var number = queryType is OptionSetValue ? (int?)((OptionSetValue)queryType).Value : queryType as int?;
            return a ? (number.HasValue ? "savedquery:evidence:A:" + CloudFlowSavedQueryEvidenceCollector.Frame(Scope,
                number.Value.ToString(CultureInfo.InvariantCulture), name) : null) :
                "savedquery:evidence:B:" + CloudFlowSavedQueryEvidenceCollector.Frame(Scope, name);
        }
    }

    internal sealed class ProcessQuerySideEvidence
    {
        internal MembershipSnapshot Snapshot;
        internal string Version, Diagnostic;
        internal bool SavedQuery;
        internal DateTimeOffset CapturedUtc;
        internal int ReturnedCount;
        internal readonly List<ComponentIdentity> Raw = new List<ComponentIdentity>();
        internal readonly Dictionary<Guid, ProcessQueryRowEvidence> Rows = new Dictionary<Guid, ProcessQueryRowEvidence>();
        internal readonly HashSet<string> Fields = new HashSet<string>(StringComparer.Ordinal);
        internal readonly List<string> Columns = new List<string>();
        internal readonly List<string> Requests = new List<string>();
    }

    internal sealed class ProcessQueryEvidenceReport
    {
        internal bool SavedQuery;
        internal ProcessQuerySideEvidence Source, Target;
        internal string Build()
        {
            var output = new StringBuilder();
            output.AppendLine(SavedQuery ? "TYPE 26 SAVED QUERY EVIDENCE ONLY" : "ALL TYPE 29 / CLOUD FLOW EVIDENCE ONLY");
            output.AppendLine("No production identity, membership, definition or absence decisions are changed. Display name alone is not portable identity.");
            Append(output, "Source", Source); Append(output, "Target", Target);
            output.AppendLine("SOURCE / TARGET FIELD COMPARISON (observations, not production matches)");
            if (SavedQuery)
            {
                foreach (bool a in new[] { true, false })
                {
                    var left = Groups(Source, a); var right = Groups(Target, a);
                    foreach (var key in left.Keys.Union(right.Keys, StringComparer.OrdinalIgnoreCase).OrderBy(k => k, StringComparer.OrdinalIgnoreCase))
                    {
                        List<ProcessQueryRowEvidence> l, r;
                        left.TryGetValue(key, out l); right.TryGetValue(key, out r);
                        output.AppendLine("Candidate " + (a ? "A" : "B") + "=" + CloudFlowSavedQueryEvidenceCollector.Safe(key) +
                            "; Source distinct rows=" + (l?.Count ?? 0) + "; Target distinct rows=" + (r?.Count ?? 0) +
                            "; evidenceStatus=" + ((l?.Count > 1 || r?.Count > 1) ? "Ambiguous" :
                                (l?.Count == 1 && r?.Count == 1) ? "UniqueCandidatePair" : "OneSidedCandidate (no absence inference)"));
                        if (l?.Count == 1 && r?.Count == 1) Compare(output, l[0], r[0]);
                    }
                }
            }
            else
            {
                var left = Source.Rows.Values.Where(Cloud).ToList(); var right = Target.Rows.Values.Where(Cloud).ToList();
                output.AppendLine("category=5 candidate counts: Source=" + left.Count + "; Target=" + right.Count);
                if (left.Count > 1 || right.Count > 1)
                    output.AppendLine("CloudFlowContext=Ambiguous; more than one category=5 candidate exists; no automatic pairing.");
                else if (left.Count == 1 && right.Count == 1 && CompleteWorkflowContext(Source) && CompleteWorkflowContext(Target))
                {
                    output.AppendLine("CloudFlowContext=Unique; one uniquely correlated category=5 row per side. Other categories are contextual evidence only. Continued identity is not inferred from display name or production resolution status.");
                    Compare(output, left[0], right[0]);
                }
                else output.AppendLine("Insufficient/ambiguous Cloud Flow context: no automatic pairing. Review raw correlations and category=5 evidence.");
                output.AppendLine("Identity hypotheses: verified uniquename when populated; workflowid for the documented solution-aware Modern Flow import lineage. Neither is promoted here.");
                output.AppendLine("Microsoft documents workflowid as the identifier across cloud-flow imports and workflowidunique as installation-specific: https://learn.microsoft.com/en-us/power-automate/manage-flows-with-code");
                output.AppendLine("Equal workflowidunique/resourceid/other IDs are audit observations, not proof of portability; resourceid is documented for internal use. Clone/recreate/collision/lifecycle scope remains to be assessed.");
            }
            output.AppendLine("Incomplete identity coverage continues to block unsafe membership absence findings. Hashes and GUID equality never force a semantic match.");
            return output.ToString();
        }
        private static bool Cloud(ProcessQueryRowEvidence row) => row.Status == "Unique" &&
            row.Row?.GetAttributeValue<object>("category") is OptionSetValue && ((OptionSetValue)row.Row["category"]).Value == 5;
        private static bool CompleteWorkflowContext(ProcessQuerySideEvidence side) =>
            side.Raw.All(c => c.Record.ObjectId.HasValue && c.Record.ObjectId != Guid.Empty) &&
            side.Rows.Values.All(r => r.Status == "Unique" && r.Row.GetAttributeValue<object>("category") is OptionSetValue &&
                ((OptionSetValue)r.Row["category"]).Value >= 0 && ((OptionSetValue)r.Row["category"]).Value <= 7);
        private static Dictionary<string, List<ProcessQueryRowEvidence>> Groups(ProcessQuerySideEvidence side, bool a) =>
            side.Rows.Values.Where(r => r.Candidate(a) != null).GroupBy(r => r.Candidate(a), StringComparer.OrdinalIgnoreCase)
                .ToDictionary(g => g.Key, g => g.OrderBy(r => r.ObjectId).ToList(), StringComparer.OrdinalIgnoreCase);
        private void Compare(StringBuilder output, ProcessQueryRowEvidence left, ProcessQueryRowEvidence right)
        {
            foreach (var field in Source.Fields.Union(Target.Fields).OrderBy(f => f, StringComparer.Ordinal))
            {
                var l = Source.Columns.Contains(field) ? CloudFlowSavedQueryEvidenceCollector.Value(left.Row, field) : null;
                var r = Target.Columns.Contains(field) ? CloudFlowSavedQueryEvidenceCollector.Value(right.Row, field) : null;
                output.AppendLine("  " + field + ": Source=" + CloudFlowSavedQueryEvidenceCollector.Safe(l) +
                    "; Target=" + CloudFlowSavedQueryEvidenceCollector.Safe(r) + "; observation=" +
                    (l == null || r == null || l == "UnexpectedValueType" || r == "UnexpectedValueType" ? "Unavailable" :
                        l == r ? "EqualObserved" : "DifferentObserved") +
                    (l != null && r != null && l != r && StringComparer.OrdinalIgnoreCase.Equals(l, r) ? " (case-only)" : ""));
            }
        }
        private static void Append(StringBuilder output, string label, ProcessQuerySideEvidence side)
        {
            output.AppendLine(label + " environment=" + CloudFlowSavedQueryEvidenceCollector.Safe(side.Snapshot.Environment.DisplayName) +
                "; solution=" + CloudFlowSavedQueryEvidenceCollector.Safe(side.Snapshot.SolutionUniqueName) + "; version=" +
                CloudFlowSavedQueryEvidenceCollector.Safe(side.Version) + "; snapshotUtc=" + side.Snapshot.CapturedAt.ToString("O") +
                "; capturedUtc=" + side.CapturedUtc.ToString("O"));
            output.AppendLine("raw=" + side.Raw.Count + "; distinctObjectIds=" + side.Rows.Count + "; returnedRows=" + side.ReturnedCount +
                "; unique=" + side.Rows.Values.Count(r => r.Status == "Unique") + "; diagnostic=" + side.Diagnostic);
            if (!side.SavedQuery) output.AppendLine("Selected solution Type29 raw count=" + side.Snapshot.Components.Count(c => c.Record.ComponentType == 29) +
                "; resolved rows retained=" + side.Snapshot.Components.Count(c => c.Record.ComponentType == 29 && c.Status == IdentityResolutionStatus.Resolved));
            foreach (var raw in side.Raw.OrderBy(r => r.Record.SolutionComponentId))
                output.AppendLine("  solutioncomponentid=" + raw.Record.SolutionComponentId + "; objectid=" + raw.Record.ObjectId +
                    "; priorIdentityStatus=" + raw.Status + "; priorIdentityDiagnostic=" + CloudFlowSavedQueryEvidenceCollector.Safe(raw.Diagnostic) +
                    "; productionPortableKey=" + CloudFlowSavedQueryEvidenceCollector.Safe(raw.ComparisonKey));
            foreach (var row in side.Rows.Values.OrderBy(r => r.ObjectId))
            {
                output.AppendLine("  objectid=" + row.ObjectId + "; correlation=" + row.Status + "; objectidEqualsPrimaryId=" +
                    (row.Row != null && row.Row.Id == row.ObjectId) + "; entityScope=" + row.Scope);
                if (!side.SavedQuery) output.AppendLine("    analysisRole=" + (Cloud(row) ? "CloudFlowCandidate" :
                    row.Status == "Unique" && row.Row.GetAttributeValue<object>("category") is OptionSetValue ? "NonCloudContext" : "UnavailableCategoryOrCorrelation"));
                foreach (var field in side.Fields.OrderBy(f => f, StringComparer.Ordinal))
                    output.AppendLine("    " + field + "=" + (side.Columns.Contains(field) ?
                        CloudFlowSavedQueryEvidenceCollector.Safe(CloudFlowSavedQueryEvidenceCollector.Value(row.Row, field)) : "UnavailableInReadableMetadata"));
                if (side.SavedQuery) output.AppendLine("    CandidateA=" + row.Candidate(true) + "; CandidateB=" + row.Candidate(false));
            }
            output.AppendLine(label + " exact request ledger: total=" + side.Requests.Count + "; WhoAmI=0; writes=0; solutioncomponent=0 (membership snapshot reused)");
            output.AppendLine("  workflow=" + side.Requests.Count(r => r.StartsWith("RetrieveMultiple(workflow,", StringComparison.Ordinal)) +
                "; savedquery=" + side.Requests.Count(r => r.StartsWith("RetrieveMultiple(savedquery,", StringComparison.Ordinal)) +
                "; RetrieveEntity=" + side.Requests.Count(r => r.StartsWith("RetrieveEntity(", StringComparison.Ordinal)) +
                "; RetrieveMetadataChanges=" + side.Requests.Count(r => r.StartsWith("RetrieveMetadataChanges(", StringComparison.Ordinal)));
            foreach (var request in side.Requests) output.AppendLine("  " + request);
        }
    }
}
#endif
