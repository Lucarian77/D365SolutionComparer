#if DEBUG
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.ServiceModel;
using System.Text;
using System.Threading;
using D365SolutionComparer.Models.Membership;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Query;

namespace D365SolutionComparer.Services.Membership
{
    /// <summary>Explicit Debug evidence only. Does not supply production identities or absence proof.</summary>
    internal sealed class Type300EvidenceCollector
    {
        internal const int BatchSize = 200, MaxIsolationGroups = 64;
        private static readonly string[] InternalNames = { "uniquename", "schemaname", "name" };
        private static readonly string[] AuditText = { "name", "uniquename", "schemaname", "displayname", "uniquecanvasappid", "canvasappidunique", "componentidunique", "version", "appversion", "appversionnumber" };
        private static readonly string[] UnsafePayload = { "binary", "attachment", "thumbnail", "image", "media", "package", "base64", "encoded", "secret", "secure", "credential", "token", "msapp" };

        internal CanvasAppEvidenceReport Capture(IOrganizationService sourceService, MembershipSnapshot source, string sourceVersion,
            IOrganizationService targetService, MembershipSnapshot target, string targetVersion, CancellationToken token, Action<string> progress = null)
        {
            if (source?.State != MembershipSnapshotState.Complete || target?.State != MembershipSnapshotState.Complete ||
                !StringComparer.OrdinalIgnoreCase.Equals(source.SolutionUniqueName, target.SolutionUniqueName))
                throw new ArgumentException("Completed snapshots of the same solution are required.");
            token.ThrowIfCancellationRequested();
            var report = new CanvasAppEvidenceReport { Source = Read(sourceService, source, sourceVersion, token, progress),
                Target = Read(targetService, target, targetVersion, token, progress) };
            report.Analyze(token); return report;
        }

        private static CanvasAppSideEvidence Read(IOrganizationService service, MembershipSnapshot snapshot, string version,
            CancellationToken token, Action<string> progress)
        {
            if (service == null) throw new ArgumentNullException(nameof(service));
            token.ThrowIfCancellationRequested();
            var side = new CanvasAppSideEvidence { Snapshot = snapshot, Version = version };
            side.Raw.AddRange(snapshot.Components.Where(c => c.Record.ComponentType == 300));
            var ids = side.Raw.Where(c => c.Record.ObjectId.HasValue && c.Record.ObjectId != Guid.Empty)
                .Select(c => c.Record.ObjectId.Value).Distinct().OrderBy(id => id).ToArray();
            foreach (var id in ids) side.Rows[id] = new CanvasAppRecordEvidence { ObjectId = id, Status = "Incomplete", Reason = "Schema not verified" };
            if (ids.Length == 0) return side;
            var metadata = Schema(service, "canvasapp", side, token);
            if (metadata == null)
            { foreach (var row in side.Rows.Values) { row.Status = side.SchemaFailure; row.Reason = "Canvas App schema unavailable; no guessed columns"; } return side; }
            side.PrimaryId = metadata.PrimaryIdAttribute;
            var columns = new List<string>(); var hashes = new HashSet<string>(StringComparer.Ordinal);
            foreach (var a in metadata.Attributes.OrderBy(a => a.LogicalName, StringComparer.Ordinal))
            {
                string name = a.LogicalName;
                bool shadow = metadata.Attributes.Where(b => b.AttributeType == AttributeTypeCode.Lookup).Any(b =>
                    StringComparer.OrdinalIgnoreCase.Equals(name, b.LogicalName + "name") ||
                    StringComparer.OrdinalIgnoreCase.Equals(name, b.LogicalName + "yominame")) ||
                    !string.IsNullOrWhiteSpace(a.AttributeOf) && (name.EndsWith("name", StringComparison.OrdinalIgnoreCase));
                // Document/package strings may contain an encoded app package. Never fetch them to calculate a hash.
                bool unsafePayload = UnsafePayload.Any(p => name.IndexOf(p, StringComparison.OrdinalIgnoreCase) >= 0) ||
                    new[] { "document", "documentbody", "solutiondocument", "content" }.Contains(name);
                bool text = a.AttributeType == AttributeTypeCode.String || a.AttributeType == AttributeTypeCode.Memo;
                bool shape = text || a.AttributeType == AttributeTypeCode.Lookup || a.AttributeType == AttributeTypeCode.Uniqueidentifier ||
                    a.AttributeType == AttributeTypeCode.Integer || a.AttributeType == AttributeTypeCode.BigInt ||
                    a.AttributeType == AttributeTypeCode.Boolean || a.AttributeType == AttributeTypeCode.Picklist ||
                    a.AttributeType == AttributeTypeCode.State || a.AttributeType == AttributeTypeCode.Status ||
                    a.AttributeType == AttributeTypeCode.DateTime || a.AttributeType == AttributeTypeCode.EntityName;
                bool readable = a.IsValidForRead == true && shape && !shadow && !unsafePayload;
                bool hash = text && (a.AttributeType == AttributeTypeCode.Memo || !AuditText.Contains(name));
                side.Schema.Add("canvasapp." + name + "; type=" + a.AttributeType + "; readable=" + a.IsValidForRead +
                    "; capture=" + (shadow ? "ExcludedLookupShadow" : unsafePayload ? "ExcludedPayload" : readable ? hash ? "HashOnly" : "Audit" : "UnavailableOrNotQueried"));
                if (!readable) continue;
                columns.Add(name); if (hash) hashes.Add(name);
            }
            if (!columns.Contains(side.PrimaryId) || metadata.Attributes.Single(a => a.LogicalName == side.PrimaryId).AttributeType != AttributeTypeCode.Uniqueidentifier)
            { foreach (var row in side.Rows.Values) row.Reason = "Readable GUID primary key not verified; no backing query"; return side; }
            // Explicit hypotheses, not a declaration of portable identity. A blank strongest field never falls back to a weaker field.
            side.CandidateField = InternalNames.FirstOrDefault(n => columns.Contains(n) && !hashes.Contains(n) &&
                metadata.Attributes.Single(a => a.LogicalName == n).AttributeType == AttributeTypeCode.String);
            var critical = new[] { side.PrimaryId }.Concat(side.CandidateField == null ? new string[0] : new[] { side.CandidateField }).ToArray();
            var backing = ReadFields(service, "canvasapp", side.PrimaryId, critical, columns.Except(critical).ToArray(), ids, side, token, progress);
            foreach (var id in ids)
            {
                var row = side.Rows[id]; var found = backing[id]; row.Status = found.Status; row.Reason = found.Reason;
                if (found.Status != "Unique") continue;
                row.PrimaryId = found.Row.Id; row.CriticalComplete = found.CriticalComplete;
                row.RuntimeColumns.AddRange(found.Columns.OrderBy(c => c, StringComparer.Ordinal));
                foreach (var column in columns)
                {
                    if (hashes.Contains(column)) row.Content[column] = found.Columns.Contains(column) ?
                        Type31ContentFingerprint.Create(found.Row.GetAttributeValue<object>(column)) : new Type31ContentFingerprint { Presence = "Unavailable" };
                    else row.Fields[column] = found.Columns.Contains(column) ? Format(found.Row.GetAttributeValue<object>(column)) : null;
                }
                row.Managed = found.Columns.Contains("ismanaged") && found.Row.GetAttributeValue<object>("ismanaged") is bool ? (bool?)found.Row.GetAttributeValue<bool>("ismanaged") : null;
                row.CandidateField = side.CandidateField;
                string identifier = side.CandidateField == null ? null : row.Get(side.CandidateField);
                bool internalText = side.CandidateField != null && found.Row.GetAttributeValue<object>(side.CandidateField) is string;
                if (found.CriticalComplete && internalText && !string.IsNullOrWhiteSpace(identifier) && !Guid.TryParse(identifier, out var ignored))
                    row.CandidateA = "canvasapp-candidate-a:" + Type31EvidenceCollector.Frame(side.CandidateField, identifier.Trim());
                // B is deliberately weak descriptive context. It cannot pair rows, repair A, or prove absence.
                if (!string.IsNullOrWhiteSpace(row.Get("displayname")) && !string.IsNullOrWhiteSpace(row.Get("appcomponenttype")))
                    row.CandidateB = "canvasapp-candidate-b:" + Type31EvidenceCollector.Frame(row.Get("displayname"), row.Get("appcomponenttype"));
            }
            CaptureAppElementLinks(service, side, token, progress);
            return side;
        }

        private static void CaptureAppElementLinks(IOrganizationService service, CanvasAppSideEvidence side, CancellationToken token, Action<string> progress)
        {
            var records = side.Snapshot.Components.Where(c => c.Record.ComponentType == 10072).ToArray();
            var ids = records.Where(c => c.Record.ObjectId.HasValue && c.Record.ObjectId != Guid.Empty)
                .Select(c => c.Record.ObjectId.Value).Distinct().OrderBy(id => id).ToArray();
            foreach (var record in records.Where(c => !c.Record.ObjectId.HasValue || c.Record.ObjectId == Guid.Empty))
                side.Links.Add(new CanvasAppDependencyEvidence { Status = "Incomplete", Reason = "Blank AppElement ObjectId" });
            if (ids.Length == 0) return;
            var metadata = Schema(service, "appelement", side, token);
            var field = metadata?.Attributes.SingleOrDefault(a => a.LogicalName == "canvasappid");
            var pk = metadata?.Attributes.SingleOrDefault(a => a.LogicalName == metadata.PrimaryIdAttribute);
            if (metadata == null || field?.IsValidForRead != true || pk?.IsValidForRead != true || pk.AttributeType != AttributeTypeCode.Uniqueidentifier ||
                !(field.AttributeType == AttributeTypeCode.Lookup || field.AttributeType == AttributeTypeCode.Uniqueidentifier))
            {
                foreach (var id in ids) side.Links.Add(new CanvasAppDependencyEvidence { AppElementId = id, Status = "Incomplete", Reason = "AppElement primary/reference schema unavailable; no guessed query" });
                return;
            }
            bool targetVerified = field is LookupAttributeMetadata && ((LookupAttributeMetadata)field).Targets?.Contains("canvasapp") == true ||
                (metadata.ManyToOneRelationships ?? new OneToManyRelationshipMetadata[0]).Any(r => r.ReferencingAttribute == "canvasappid" &&
                    r.ReferencedEntity == "canvasapp" && r.ReferencedAttribute == side.PrimaryId);
            side.Schema.Add("appelement." + metadata.PrimaryIdAttribute + "; capture=ScopedPrimaryCorrelation");
            side.Schema.Add("appelement.canvasappid; capture=ScopedReferenceAudit; targetSchemaVerified=" + targetVerified);
            var rows = ReadFields(service, "appelement", metadata.PrimaryIdAttribute, new[] { metadata.PrimaryIdAttribute, "canvasappid" },
                new string[0], ids, side, token, progress);
            foreach (var id in ids)
            {
                var found = rows[id]; var link = new CanvasAppDependencyEvidence { AppElementId = id, Status = found.Status, Reason = found.Reason, TargetSchemaVerified = targetVerified };
                side.Links.Add(link); if (found.Status != "Unique") continue;
                if (!found.CriticalComplete) { link.Status = "Incomplete"; continue; }
                object raw = found.Row.GetAttributeValue<object>("canvasappid");
                var reference = raw as EntityReference;
                link.CanvasId = raw is Guid ? (Guid?)raw : reference?.LogicalName == "canvasapp" ? (Guid?)reference.Id : null;
                if (link.CanvasId == null || link.CanvasId == Guid.Empty) { link.Status = "Incomplete"; link.Reason = "Canvas App reference blank or wrong target; no fallback"; }
            }
        }

        private static EntityMetadata Schema(IOrganizationService service, string entity, CanvasAppSideEvidence side, CancellationToken token)
        {
            token.ThrowIfCancellationRequested();
            side.Requests.Add("Execute RetrieveEntity(" + entity + ", Attributes|Relationships, RetrieveAsIfPublished=False)");
            try
            {
                var response = service.Execute(new RetrieveEntityRequest { LogicalName = entity,
                    EntityFilters = EntityFilters.Attributes | EntityFilters.Relationships, RetrieveAsIfPublished = false }) as RetrieveEntityResponse;
                token.ThrowIfCancellationRequested(); var metadata = response?.EntityMetadata;
                if (metadata == null || metadata.LogicalName != entity || !ValidName(metadata.PrimaryIdAttribute) || metadata.Attributes == null ||
                    metadata.Attributes.Any(a => a == null || !ValidName(a.LogicalName)) ||
                    metadata.Attributes.GroupBy(a => a.LogicalName, StringComparer.OrdinalIgnoreCase).Any(g => g.Count() != 1))
                { side.SchemaFailure = "Incomplete"; side.Schema.Add(entity + ": incomplete/conflicting schema"); return null; }
                foreach (var relationship in metadata.ManyToOneRelationships ?? new OneToManyRelationshipMetadata[0])
                    if (relationship != null) side.Relationships.Add(entity + "." + relationship.ReferencingAttribute + " -> " + relationship.ReferencedEntity +
                        "." + relationship.ReferencedAttribute + "; relationship=" + relationship.SchemaName);
                foreach (var relationship in metadata.OneToManyRelationships ?? new OneToManyRelationshipMetadata[0])
                    if (relationship != null) side.Relationships.Add("Incoming " + relationship.ReferencingEntity + "." + relationship.ReferencingAttribute + " -> " +
                        entity + "." + relationship.ReferencedAttribute + "; relationship=" + relationship.SchemaName + "; not scanned");
                foreach (var relationship in metadata.ManyToManyRelationships ?? new ManyToManyRelationshipMetadata[0])
                    if (relationship != null) side.Relationships.Add("Intersect " + relationship.IntersectEntityName + "; relationship=" + relationship.SchemaName + "; not scanned");
                foreach (var attribute in metadata.Attributes.OfType<LookupAttributeMetadata>())
                    side.Relationships.Add(entity + "." + attribute.LogicalName + "; metadata lookup targets=[" + string.Join(",", attribute.Targets ?? new string[0]) + "]");
                return metadata;
            }
            catch (OperationCanceledException) { throw; }
            catch (Exception error) { token.ThrowIfCancellationRequested(); side.RetrievalDiagnostics.Add(SafeFault(error)); side.SchemaFailure = "Faulted"; side.Schema.Add(entity + ": schema retrieval failed; server details withheld"); return null; }
        }

        private static Dictionary<Guid, LookupEvidence> ReadFields(IOrganizationService service, string entity, string primary,
            string[] critical, string[] optional, Guid[] ids, CanvasAppSideEvidence side, CancellationToken token, Action<string> progress)
        {
            var result = new Dictionary<Guid, LookupEvidence>();
            for (int offset = 0; offset < ids.Length; offset += BatchSize)
            {
                var batch = ids.Skip(offset).Take(BatchSize).ToArray();
                side.RetrievalDiagnostics.Add(entity + ": minimal identity-critical first; columns=[" + string.Join(",", critical) + "]");
                var current = Retrieve(service, entity, primary, critical, batch, side, token, progress);
                int remaining = MaxIsolationGroups;
                if (current.Values.Any(r => r.Status == "Faulted"))
                {
                    side.RetrievalDiagnostics.Add(entity + ": critical query faulted; primary-only retry before bounded isolation");
                    current = Retrieve(service, entity, primary, new[] { primary }, batch, side, token, progress);
                    if (current.Values.Any(r => r.Status == "Unique"))
                        Isolate(service, entity, primary, critical.Where(c => c != primary).ToArray(), batch, current, true, side, token, progress, ref remaining);
                }
                if (current.Values.Any(r => r.Status == "Unique" && r.CriticalComplete))
                    Isolate(service, entity, primary, optional.OrderBy(c => c, StringComparer.Ordinal).ToArray(), batch, current, false, side, token, progress, ref remaining);
                else side.RetrievalDiagnostics.Add(entity + ": optional fields excluded/deferred because critical correlation is incomplete");
                foreach (var id in batch) result[id] = current[id];
            }
            return result;
        }

        private static void Isolate(IOrganizationService service, string entity, string primary, string[] fields, Guid[] batch,
            Dictionary<Guid, LookupEvidence> current, bool critical, CanvasAppSideEvidence side, CancellationToken token,
            Action<string> progress, ref int remaining)
        {
            if (fields.Length == 0) return;
            var pending = new Queue<string[]>(); pending.Enqueue(fields);
            while (pending.Count > 0)
            {
                token.ThrowIfCancellationRequested(); var group = pending.Dequeue();
                if (remaining == 0)
                {
                    side.RetrievalDiagnostics.Add("Excluded unproven " + (critical ? "identity-critical" : "optional") + " columns after isolation limit=" + MaxIsolationGroups + ": [" + string.Join(",", group) + "]");
                    if (critical) foreach (var value in current.Values) value.CriticalComplete = false;
                    continue;
                }
                remaining--;
                side.RetrievalDiagnostics.Add("Isolating " + (critical ? "identity-critical" : "optional/hash-only") + " group=[" + string.Join(",", group) + "]; requestedIds=" + batch.Length);
                var attempt = Retrieve(service, entity, primary, new[] { primary }.Concat(group).ToArray(), batch, side, token, progress);
                if (attempt.Values.Any(r => r.Status == "Faulted"))
                {
                    if (group.Length > 1)
                    { int half = group.Length / 2; pending.Enqueue(group.Take(half).ToArray()); pending.Enqueue(group.Skip(half).ToArray()); }
                    else
                    {
                        side.RetrievalDiagnostics.Add("Attribute faulted: " + group[0] + "; class=" + (critical ? "IdentityCriticalUnavailable" : "OptionalUnavailable") + "; no value retained");
                        if (critical) foreach (var value in current.Values) value.CriticalComplete = false;
                    }
                    continue;
                }
                side.RetrievalDiagnostics.Add("Group retrieval completed: [" + string.Join(",", group) + "]");
                foreach (var id in batch)
                {
                    var existing = current[id]; var found = attempt[id];
                    if (existing.Status != "Unique") continue;
                    if (found.Status != "Unique")
                    {
                        if (critical) existing.Status = found.Status == "Duplicate" ? "Duplicate" : "Incomplete";
                        else existing.CriticalComplete = false; // Preserve primary correlation, but contradictory optional reads block pairing.
                        existing.Reason = "Contradictory or incomplete retry correlation: " + found.Status + "; no candidate";
                        continue;
                    }
                    foreach (var field in group)
                    {
                        existing.Columns.Add(field);
                        if (found.Row.Attributes.TryGetValue(field, out var raw)) existing.Row[field] = raw;
                    }
                }
            }
        }

        private static string SafeFault(Exception error)
        {
            var sdk = error as FaultException<OrganizationServiceFault>;
            string type = sdk != null ? "OrganizationServiceFault" : error is FaultException ? "FaultException" :
                error is TimeoutException ? "TimeoutException" : error is InvalidOperationException ? "InvalidOperationException" : "Exception";
            string code = sdk?.Detail != null ? "0x" + unchecked((uint)sdk.Detail.ErrorCode).ToString("X8", CultureInfo.InvariantCulture) : "Unavailable";
            // No exception message, stack, trace, inner fault, URL, token, or payload is copied from the service.
            return "FaultType=" + type + "; SDKErrorCode=" + code + "; Message=Read request failed; service details withheld";
        }

        private static Dictionary<Guid, LookupEvidence> Retrieve(IOrganizationService service, string entity, string primary, string[] columns,
            Guid[] ids, CanvasAppSideEvidence side, CancellationToken token, Action<string> progress)
        {
            var result = ids.ToDictionary(id => id, id => new LookupEvidence { Status = "Incomplete", Reason = "Terminal retrieval not proven" });
            for (int offset = 0; offset < ids.Length; offset += BatchSize)
            {
                var batch = ids.Skip(offset).Take(BatchSize).ToArray(); var rows = new Dictionary<Guid, Entity>(); var duplicates = new HashSet<Guid>();
                int page = 1, returned = 0; string cookie = null;
                try
                {
                    while (true)
                    {
                        token.ThrowIfCancellationRequested();
                        var query = new QueryExpression(entity) { ColumnSet = new ColumnSet(columns),
                            PageInfo = new PagingInfo { Count = BatchSize, PageNumber = page, PagingCookie = cookie } };
                        query.Criteria.AddCondition(primary, ConditionOperator.In, batch.Select(id => (object)id).ToArray()); query.AddOrder(primary, OrderType.Ascending);
                        side.Requests.Add("RetrieveMultiple " + entity + "; columns=[" + string.Join(",", columns) + "]; " + primary + " IN Guid[" + batch.Length + "]; page=" + page);
                        progress?.Invoke(side.Snapshot.Environment.DisplayName + ": reading scoped " + entity + " page " + page);
                        var response = service.RetrieveMultiple(query); token.ThrowIfCancellationRequested();
                        if (response == null) break;
                        returned += response.Entities.Count;
                        side.Pages.Add(entity + "; requestedIds=" + batch.Length + "; returnedRows=" + response.Entities.Count + "; totalRows=" + returned +
                            "; page=" + page + "; MoreRecords=" + response.MoreRecords + "; PagingCookieSupplied=" + !string.IsNullOrEmpty(response.PagingCookie));
                        if (response.Entities.Any(r => r == null || r.LogicalName != entity || r.Id == Guid.Empty ||
                            Type31EvidenceCollector.Id(r, primary) != r.Id || !batch.Contains(r.Id)))
                        { foreach (var id in batch) result[id].Reason = "Blank/conflicting/foreign primary key; no candidates"; break; }
                        int before = rows.Count;
                        foreach (var group in response.Entities.GroupBy(r => r.Id))
                        {
                            if (group.Count() > 1) duplicates.Add(group.Key);
                            foreach (var row in group)
                                if (rows.TryGetValue(row.Id, out var previous)) { if (!Type31EvidenceCollector.SameReturnedRow(previous, row)) duplicates.Add(row.Id); }
                                else rows.Add(row.Id, row);
                        }
                        if (!response.MoreRecords)
                        {
                            foreach (var id in batch) result[id] = new LookupEvidence { Row = rows.TryGetValue(id, out var row) ? row : null,
                                Status = !rows.ContainsKey(id) ? "Missing" : duplicates.Contains(id) ? "Duplicate" : "Unique",
                                Columns = new HashSet<string>(columns, StringComparer.Ordinal),
                                Reason = "Terminal page received; distinctReturnedIds=" + rows.Count + "; pageCount=" + page + "; not membership absence evidence" };
                            break;
                        }
                        if (rows.Count == before || !string.IsNullOrEmpty(response.PagingCookie) && response.PagingCookie == cookie)
                        { foreach (var id in batch) result[id].Reason = "Stalled paging; terminal retrieval not proven"; break; }
                        cookie = response.PagingCookie; page++;
                    }
                }
                catch (OperationCanceledException) { throw; }
                catch (Exception error)
                {
                    token.ThrowIfCancellationRequested(); var fault = SafeFault(error);
                    side.RetrievalDiagnostics.Add("Faulted query; entity=" + entity + "; requestedIds=" + batch.Length + "; columns=[" + string.Join(",", columns) + "]; " + fault);
                    foreach (var id in batch) result[id] = new LookupEvidence { Status = "Faulted", Reason = fault };
                }
            }
            return result;
        }

        private sealed class LookupEvidence
        {
            internal string Status, Reason; internal Entity Row; internal bool CriticalComplete = true;
            internal HashSet<string> Columns = new HashSet<string>(StringComparer.Ordinal);
        }
        private static bool ValidName(string value) => !string.IsNullOrWhiteSpace(value) && value.All(c => char.IsLetterOrDigit(c) || c == '_');
        private static string Format(object raw)
        {
            if (raw == null) return "";
            if (raw is EntityReference) { var r = (EntityReference)raw; return ValidName(r.LogicalName) ? r.LogicalName + ":" + r.Id.ToString("D") : null; }
            if (raw is OptionSetValue) return ((OptionSetValue)raw).Value.ToString(CultureInfo.InvariantCulture);
            if (raw is DateTime) return ((DateTime)raw).ToUniversalTime().ToString("O", CultureInfo.InvariantCulture);
            if (raw is string || raw is int || raw is long || raw is bool || raw is Guid) return Convert.ToString(raw, CultureInfo.InvariantCulture)?.Trim();
            return null;
        }
    }

    internal sealed class CanvasAppSideEvidence
    {
        internal MembershipSnapshot Snapshot;
        internal string Version, PrimaryId, CandidateField, SchemaFailure;
        internal readonly List<ComponentIdentity> Raw = new List<ComponentIdentity>();
        internal readonly SortedDictionary<Guid, CanvasAppRecordEvidence> Rows = new SortedDictionary<Guid, CanvasAppRecordEvidence>();
        internal readonly List<CanvasAppDependencyEvidence> Links = new List<CanvasAppDependencyEvidence>();
        internal readonly List<string> Schema = new List<string>(), Relationships = new List<string>(), Requests = new List<string>(), Pages = new List<string>(), RetrievalDiagnostics = new List<string>();
    }
    internal sealed class CanvasAppRecordEvidence
    {
        internal Guid ObjectId; internal Guid? PrimaryId; internal bool? Managed;
        internal string Status, Reason, CandidateField, CandidateA, CandidateB;
        internal bool CriticalComplete, DuplicateA, DuplicateB;
        internal readonly SortedDictionary<string, string> Fields = new SortedDictionary<string, string>(StringComparer.Ordinal);
        internal readonly SortedDictionary<string, Type31ContentFingerprint> Content = new SortedDictionary<string, Type31ContentFingerprint>(StringComparer.Ordinal);
        internal readonly List<string> RuntimeColumns = new List<string>();
        internal string Get(string field) => Fields.TryGetValue(field, out var value) ? value : null;
        internal string Evidence(string field) => Content.TryGetValue(field, out var hash) ? hash.Evidence : Get(field);
    }
    internal sealed class CanvasAppDependencyEvidence
    {
        internal Guid? AppElementId, CanvasId; internal string Status, Reason; internal bool TargetSchemaVerified;
    }
    internal sealed class CanvasAppPairEvidence
    {
        internal CanvasAppRecordEvidence Source, Target; internal string Outcome;
        internal readonly HashSet<string> Categories = new HashSet<string>(StringComparer.Ordinal);
    }
    internal sealed class CanvasAppEvidenceReport
    {
        internal CanvasAppSideEvidence Source, Target;
        internal readonly List<CanvasAppPairEvidence> Pairs = new List<CanvasAppPairEvidence>();
        internal void Analyze(CancellationToken token)
        {
            foreach (var side in new[] { Source, Target })
            {
                foreach (var group in side.Rows.Values.Where(r => r.CandidateA != null).GroupBy(r => r.CandidateA, StringComparer.OrdinalIgnoreCase).Where(g => g.Count() > 1))
                    foreach (var row in group) row.DuplicateA = true;
                foreach (var group in side.Rows.Values.Where(r => r.CandidateB != null).GroupBy(r => r.CandidateB, StringComparer.OrdinalIgnoreCase).Where(g => g.Count() > 1))
                    foreach (var row in group) row.DuplicateB = true;
            }
            foreach (var key in Source.Rows.Values.Concat(Target.Rows.Values).Where(r => r.CandidateA != null).Select(r => r.CandidateA)
                .Distinct(StringComparer.OrdinalIgnoreCase).OrderBy(k => k, StringComparer.OrdinalIgnoreCase))
            {
                token.ThrowIfCancellationRequested();
                var left = Source.Rows.Values.Where(r => StringComparer.OrdinalIgnoreCase.Equals(r.CandidateA, key)).ToArray();
                var right = Target.Rows.Values.Where(r => StringComparer.OrdinalIgnoreCase.Equals(r.CandidateA, key)).ToArray();
                if (left.Length > 1 || right.Length > 1)
                {
                    foreach (var row in left) Pairs.Add(new CanvasAppPairEvidence { Source = row, Outcome = "Ambiguous" });
                    foreach (var row in right) Pairs.Add(new CanvasAppPairEvidence { Target = row, Outcome = "Ambiguous" });
                }
                else Pairs.Add(new CanvasAppPairEvidence { Source = left.SingleOrDefault(), Target = right.SingleOrDefault(),
                    Outcome = left.Length == 1 && right.Length == 1 ? "SemanticPair" : "OneSidedEvidence" });
            }
            foreach (var side in new[] { Source, Target })
            {
                foreach (var row in side.Rows.Values.Where(r => r.CandidateA == null))
                    Pairs.Add(new CanvasAppPairEvidence { Source = side == Source ? row : null, Target = side == Target ? row : null,
                        Outcome = row.Status == "Duplicate" ? "Ambiguous" : "Incomplete" });
                foreach (var raw in side.Raw.Where(r => !r.Record.ObjectId.HasValue || r.Record.ObjectId == Guid.Empty))
                    Pairs.Add(new CanvasAppPairEvidence { Outcome = "Incomplete" });
            }
            foreach (var pair in Pairs)
            {
                token.ThrowIfCancellationRequested(); pair.Categories.Add(pair.Outcome);
                if (pair.Outcome != "SemanticPair") continue;
                var left = pair.Source; var right = pair.Target;
                pair.Categories.Add(left.PrimaryId == right.PrimaryId ? "SamePrimaryId" : "DifferentPrimaryId");
                var uniqueFields = left.Fields.Keys.Union(right.Fields.Keys).Where(f => f.IndexOf("unique", StringComparison.OrdinalIgnoreCase) >= 0 &&
                    !string.IsNullOrWhiteSpace(left.Get(f)) && !string.IsNullOrWhiteSpace(right.Get(f))).ToArray();
                if (uniqueFields.Length > 0) pair.Categories.Add(uniqueFields.Any(f => !StringComparer.OrdinalIgnoreCase.Equals(left.Get(f), right.Get(f))) ? "DifferentUniqueId" : "SameUniqueId");
                var content = left.Content.Keys.Union(right.Content.Keys).ToArray();
                if (content.Any(f => left.Content.ContainsKey(f) && right.Content.ContainsKey(f) && left.Content[f].Known && right.Content[f].Known && left.Content[f].Sha256 != right.Content[f].Sha256))
                    pair.Categories.Add("DifferentDefinitionEvidence");
                else if (content.Length > 0 && content.All(f => left.Content.ContainsKey(f) && right.Content.ContainsKey(f) && left.Content[f].Known && right.Content[f].Known)) pair.Categories.Add("SameDefinitionEvidence");
                if (left.Managed.HasValue && right.Managed.HasValue && left.Managed != right.Managed)
                { pair.Categories.Add("ManagedTransition"); if (left.Managed == false && right.Managed == true) pair.Categories.Add("UnmanagedToManaged"); }
            }
        }
        private static void Line(StringBuilder text, params object[] values) => text.AppendLine(string.Join("\t", values.Select(Type31EvidenceCollector.Safe)));
        internal string Build()
        {
            var text = new StringBuilder(); text.AppendLine("TYPE 300 CANVAS APP EVIDENCE - DEBUG ONLY");
            text.AppendLine("Evidence only. Type 300 remains Unsupported / Indeterminate; no portable comparison key or absence proof is supplied.");
            text.AppendLine("Uniqueness is scoped to selected solution members; rename/recreate portability remains unproven. Candidate B and hashes never repair Candidate A.");
            text.AppendLine("\nRAW TYPE 300 MEMBERSHIP");
            foreach (var side in new[] { Source, Target })
            {
                string label = side == Source ? "Source" : "Target";
                Line(text, label, side.Snapshot.Environment.DisplayName, side.Snapshot.SolutionUniqueName, "version=" + side.Version,
                    "snapshotUtc=" + side.Snapshot.CapturedAt.ToUniversalTime().ToString("O"), "raw=" + side.Raw.Count, "distinct=" + side.Rows.Count,
                    "blank=" + side.Raw.Count(r => !r.Record.ObjectId.HasValue || r.Record.ObjectId == Guid.Empty));
                foreach (var raw in side.Raw.OrderBy(r => r.Record.SolutionComponentId))
                    Line(text, label, "solutioncomponentid=" + raw.Record.SolutionComponentId, "objectid=" + raw.Record.ObjectId,
                        "productionStatus=" + raw.Status, "productionDiagnostic=" + raw.Diagnostic);
            }
            text.AppendLine("\nBACKING CANVAS APP CORRELATION");
            EachSide(text, (s, label) => { foreach (var row in s.Rows.Values) Line(text, label, "objectid=" + row.ObjectId, "primaryId=" + row.PrimaryId,
                "objectid==primaryId=" + (row.PrimaryId.HasValue ? (row.ObjectId == row.PrimaryId.Value).ToString() : "Unknown"), row.Status, row.Reason); });
            text.AppendLine("\nREADABLE / UNAVAILABLE SCHEMA"); EachSide(text, (s, label) => { foreach (var item in s.Schema) Line(text, label, item); });
            text.AppendLine("\nRUNTIME-READABLE / FAULTED COLUMNS"); EachSide(text, (s, label) => {
                foreach (var item in s.RetrievalDiagnostics) Line(text, label, item);
                foreach (var row in s.Rows.Values) Line(text, label, row.ObjectId, "runtimeSucceeded=[" + string.Join(",", row.RuntimeColumns) + "]", "criticalComplete=" + row.CriticalComplete);
                foreach (var item in s.Pages) Line(text, label, item);
            });
            text.AppendLine("\nPARENT / REFERENCE RELATIONSHIP DISCOVERY"); EachSide(text, (s, label) => { foreach (var item in s.Relationships) Line(text, label, item); });
            text.AppendLine("\nCANDIDATE IDENTITY ANALYSIS"); EachSide(text, (s, label) => {
                Line(text, label, "Candidate A hypothesis: strongest metadata-readable internal text field=" + (s.CandidateField ?? "Unavailable"),
                    "priority=uniquename,schemaname,name; blank/GUID-only strongest field has no fallback",
                    "Candidate B hypothesis: displayname+appcomponenttype, descriptive evidence only");
                foreach (var row in s.Rows.Values) Line(text, label, row.ObjectId, "CandidateA=" + (row.CandidateA ?? "Incomplete"),
                    "CandidateB=" + (row.CandidateB ?? "NotAvailable"), "DuplicateA=" + row.DuplicateA, "DuplicateB=" + row.DuplicateB);
            });
            text.AppendLine("\nDUPLICATE / REPEATED ANALYSIS"); EachSide(text, (s, label) => {
                foreach (var group in s.Raw.Where(r => r.Record.ObjectId.HasValue).GroupBy(r => r.Record.ObjectId).Where(g => g.Count() > 1))
                    Line(text, label, "RepeatedRawObjectId=" + group.Key, "references=" + group.Count(), "not distinct backing identities");
                Line(text, label, "CandidateACollisionGroups=" + s.Rows.Values.Where(r => r.DuplicateA).Select(r => r.CandidateA).Distinct(StringComparer.OrdinalIgnoreCase).Count(),
                    "CandidateBCollisionGroups=" + s.Rows.Values.Where(r => r.DuplicateB).Select(r => r.CandidateB).Distinct(StringComparer.OrdinalIgnoreCase).Count());
            });
            text.AppendLine("\nSOURCE / TARGET FIELD COMPARISON");
            foreach (var pair in Pairs)
            {
                Line(text, pair.Outcome, "Source=" + pair.Source?.ObjectId, "Target=" + pair.Target?.ObjectId);
                foreach (var field in (pair.Source?.Fields.Keys.AsEnumerable() ?? Enumerable.Empty<string>()).Union(pair.Target?.Fields.Keys.AsEnumerable() ?? Enumerable.Empty<string>())
                    .Union(pair.Source?.Content.Keys.AsEnumerable() ?? Enumerable.Empty<string>()).Union(pair.Target?.Content.Keys.AsEnumerable() ?? Enumerable.Empty<string>()).OrderBy(f => f, StringComparer.Ordinal))
                {
                    string left = pair.Source?.Evidence(field), right = pair.Target?.Evidence(field);
                    Line(text, field, "Source=" + (left ?? "Unavailable"), "Target=" + (right ?? "Unavailable"), pair.Outcome != "SemanticPair" ? "NotPaired" :
                        left == null || right == null ? "Unknown" : StringComparer.Ordinal.Equals(left, right) ? "EqualObserved" : "DifferentObserved");
                }
            }
            text.AppendLine("\nLIFECYCLE CORRELATION MATRIX");
            foreach (var category in new[] { "SemanticPair", "SamePrimaryId", "DifferentPrimaryId", "SameUniqueId", "DifferentUniqueId", "SameDefinitionEvidence", "DifferentDefinitionEvidence", "ManagedTransition", "UnmanagedToManaged", "OneSidedEvidence", "Ambiguous", "Incomplete" })
                Line(text, category, Pairs.Count(p => p.Categories.Contains(category)));
            text.AppendLine("\nDIFFERING PRIMARY-ID SEMANTIC PAIRS");
            foreach (var pair in Pairs.Where(p => p.Categories.Contains("DifferentPrimaryId"))) Line(text, pair.Source.CandidateA, pair.Source.PrimaryId, pair.Target.PrimaryId, "Candidate A unique on both sides; no GUID/hash pairing");
            if (!Pairs.Any(p => p.Categories.Contains("DifferentPrimaryId"))) text.AppendLine("No differing-primary-ID semantic pair was observed.");
            text.AppendLine("\nTYPE 10072 DEPENDENCY ASSESSMENT"); EachSide(text, (s, label) => {
                foreach (var link in s.Links)
                {
                    CanvasAppRecordEvidence app = null; if (link.CanvasId.HasValue) s.Rows.TryGetValue(link.CanvasId.Value, out app);
                    bool paired = app != null && app.CandidateA != null && !app.DuplicateA && link.TargetSchemaVerified &&
                        Pairs.Any(p => p.Outcome == "SemanticPair" && (p.Source == app || p.Target == app));
                    Line(text, label, "AppElementId=" + link.AppElementId, "appelement.canvasappid=" + link.CanvasId, link.Status, link.Reason,
                        "Type300BackingId=" + app?.PrimaryId, "localReferenceEqualsBacking=" + (app?.PrimaryId.HasValue == true && link.CanvasId == app.PrimaryId),
                        "targetSchemaVerified=" + link.TargetSchemaVerified, "CandidateA=" + app?.CandidateA,
                        "dependency=" + (paired ? "SemanticPairObserved; potential future reuse only; production identity not approved" : "IncompleteOrAmbiguous; no independent identity"));
                }
                if (s.Links.Count == 0) Line(text, label, "No selected AppElement dependency was captured.");
            });
            text.AppendLine("\nPORTABILITY ASSESSMENT");
            Line(text, "Unique semantic pairs=" + Pairs.Count(p => p.Outcome == "SemanticPair"), "differing primary IDs=" + Pairs.Count(p => p.Categories.Contains("DifferentPrimaryId")));
            text.AppendLine("Observed evidence does not establish lifecycle portability or safe absence semantics. Additional live/lifecycle review required before production promotion or Type 10072 reuse.");
            text.AppendLine("\nEXACT REQUEST LEDGER"); EachSide(text, (s, label) => {
                foreach (var request in s.Requests) Line(text, label, request);
                foreach (var group in s.Requests.GroupBy(r => r.StartsWith("Execute", StringComparison.Ordinal) ? r.Split('(')[0] + "(" + r.Split('(')[1].Split(',')[0] + ")" : r.Split(';')[0])) Line(text, label, group.Key, "reads=" + group.Count());
                Line(text, label, "TotalReads=" + s.Requests.Count, "AdditionalWhoAmI=0", "Writes=0", "NormalMembershipEvidenceRequests=0");
            }); return text.ToString();
        }
        private void EachSide(StringBuilder text, Action<CanvasAppSideEvidence, string> action) { action(Source, "Source"); action(Target, "Target"); }
    }
}
#endif
