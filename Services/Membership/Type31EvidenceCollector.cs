#if DEBUG
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading;
using D365SolutionComparer.Models.Membership;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Metadata.Query;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Query;

namespace D365SolutionComparer.Services.Membership
{
    /// <summary>Explicit Debug evidence capture only. Never supplies a production identity or absence proof.</summary>
    internal sealed class Type31EvidenceCollector
    {
        internal const int BatchSize = 200;
        internal const string PrimaryId = "reportid";
        internal const string ProposedIdentifierField = "uniquename";
        internal static readonly string[] AuditFields = { PrimaryId, "reportidunique", "signatureid", "signaturelcid",
            "name", "filename", "reporttypecode", "categories", "relatedentities", "objecttypecode", "languagecode",
            "ispersonal", "iscustomreport", "statecode", "statuscode", "ismanaged", "componentstate", "ownerid",
            "parentreportid", "originalreportid", "reportcategorycode", "signaturedate", "uniquename", "schemaname",
            "organizationid", "solutionid", "owninguser", "owningteam" };
        internal static readonly string[] ContentFields = { "bodytext", "bodyxml", "rdl", "reportdefinition",
            "query", "defaultfilter", "description", "presentationdescription", "datadescription", "bodybinary" };

        internal Type31EvidenceReport Capture(IOrganizationService sourceService, MembershipSnapshot source,
            string sourceVersion, IOrganizationService targetService, MembershipSnapshot target,
            string targetVersion, CancellationToken token, Action<string> progress = null)
        {
            if (source?.State != MembershipSnapshotState.Complete || target?.State != MembershipSnapshotState.Complete ||
                !StringComparer.OrdinalIgnoreCase.Equals(source.SolutionUniqueName, target.SolutionUniqueName))
                throw new ArgumentException("Completed snapshots of the same solution are required.");
            token.ThrowIfCancellationRequested();
            var report = new Type31EvidenceReport
            {
                Source = Read(sourceService, source, sourceVersion, token, progress),
                Target = Read(targetService, target, targetVersion, token, progress)
            };
            foreach (var side in new[] { report.Source, report.Target })
            {
                foreach (var row in side.Rows.Values.Where(r => r.Status == "Unique")) row.ConstructCandidates();
                MarkDuplicates(side.Rows.Values, false); MarkDuplicates(side.Rows.Values, true);
                foreach (var group in side.Rows.Values.Where(r => r.Status == "Unique" && r.SignatureId.HasValue)
                    .GroupBy(r => r.SignatureId.Value).Where(g => g.Count() > 1))
                    foreach (var row in group) row.DuplicateSignature = true;
                RefineUnsignedEvidence(side);
            }
            report.Analyze(token);
            token.ThrowIfCancellationRequested();
            return report;
        }

        private static void MarkDuplicates(IEnumerable<Type31ReportEvidence> rows, bool weaker)
        {
            foreach (var group in rows.Where(r => r.Status == "Unique" && (weaker ? r.CandidateB : r.CandidateA) != null)
                .GroupBy(r => weaker ? r.CandidateB : r.CandidateA, StringComparer.OrdinalIgnoreCase).Where(g => g.Count() > 1))
                foreach (var row in group) { if (weaker) row.DuplicateB = true; else row.DuplicateA = true; }
        }

        private static void RefineUnsignedEvidence(Type31SideEvidence side)
        {
            foreach (var row in side.Rows.Values)
            {
                var production = side.Raw.Where(c => c.Record.ObjectId == row.ObjectId).ToArray();
                row.ProductionSignedSnapshot = production.Any(c => c.Status == IdentityResolutionStatus.Resolved &&
                    StringComparer.OrdinalIgnoreCase.Equals(c.SemanticKind, ComponentSemanticKinds.Report));
                bool verified = row.ProductionSignedSnapshot && row.Status == "Unique" && row.SignatureId.HasValue && !row.DuplicateSignature &&
                    production.All(c => c.Status == IdentityResolutionStatus.Resolved && StringComparer.OrdinalIgnoreCase.Equals(c.SemanticKind, ComponentSemanticKinds.Report) &&
                        StringComparer.OrdinalIgnoreCase.Equals(c.ComparisonKey, row.SignatureId.Value.ToString("D")));
                row.ReportSubset = verified ? "VerifiedSignedSubset" : "UnsignedOrUnverifiedRemainder";
                row.BoundaryReason = verified ? "Existing production signed identity agrees with uniquely correlated nonblank signature; production gate unchanged" :
                    row.ProductionSignedSnapshot ? "Current evidence does not reverify the signed snapshot; no production result changed and no unsigned fallback applied" :
                    "Outside the existing verified production signed subset; blank/missing/nonblank signature observation alone cannot establish signed eligibility";
                if (row.Status != "Unique") { row.ProposedBlockingReason = "Backing correlation " + row.Status; continue; }
                // Refine only the unsupported remainder. A formerly resolved signed snapshot is
                // never offered an alternative unsigned identity, even if fresh evidence is incomplete.
                if (row.ProductionSignedSnapshot) { row.ProposedBlockingReason = "Existing signed subset excluded from unsigned Candidate P"; continue; }
                var blockers = new List<string>();
                if (!row.RuntimeColumns.Contains("signatureid") || row.Get("signatureid") != "")
                    blockers.Add("Unsigned signature state unavailable/malformed/nonblank; no signed-gate bypass");
                string identifier = row.Get(ProposedIdentifierField);
                if (!side.FixedIdentifierVerified || !row.RuntimeColumns.Contains(ProposedIdentifierField) ||
                    string.IsNullOrWhiteSpace(identifier) || Guid.TryParse(identifier, out var ignored))
                    blockers.Add("Fixed internal uniquename unavailable/blank/unverified; name/filename/schemaname cannot repair identity");
                if (row.ScopeStatus == "Available" && LogicalScope(row.RelatedScope))
                {
                    row.ProposedScope = row.RelatedScope.Trim().ToLowerInvariant(); row.ProposedScopeStatus = "Verified";
                    row.ProposedScopeReason = "Reused metadata-proven report-related entity scope";
                }
                else if (row.ScopeStatus == "Ambiguous" || row.ScopeStatus == "Faulted" || row.ScopeStatus == "Incomplete")
                { row.ProposedScopeStatus = row.ScopeStatus; row.ProposedScopeReason = row.ScopeReason; }
                else
                {
                    string input = row.Get("relatedentities"); if (string.IsNullOrWhiteSpace(input)) input = row.Get("objecttypecode");
                    var tables = LogicalScope(input) ? side.Snapshot.Components.Where(c => c.Record.ComponentType == 1 &&
                        StringComparer.OrdinalIgnoreCase.Equals(c.ComparisonKey, input.Trim())).ToArray() : new ComponentIdentity[0];
                    bool unique = tables.Length > 0 && tables.All(c => c.Status == IdentityResolutionStatus.Resolved && c.InventoryAbsencePolicy == InventoryAbsencePolicy.CompleteInventory) &&
                        tables.Select(c => c.Record.ObjectId).Distinct().Count() == 1 && tables[0].Record.ObjectId.HasValue && tables[0].Record.ObjectId != Guid.Empty;
                    row.ProposedScopeStatus = unique ? "Verified" : tables.Length > 0 ? "Ambiguous" : "Incomplete";
                    row.ProposedScope = unique ? input.Trim().ToLowerInvariant() : null;
                    row.ProposedScopeReason = unique ? "Reused unique completed snapshot Table identity" : "Raw scope/name/numeric code is not independently verified; no global/tableless assumption";
                }
                if (row.ProposedScopeStatus != "Verified") blockers.Add("Verified entity/table scope " + row.ProposedScopeStatus);
                int type = 0, language = 0;
                if (!row.RuntimeColumns.Contains("reporttypecode") || !int.TryParse(row.Get("reporttypecode"), NumberStyles.Integer, CultureInfo.InvariantCulture, out type))
                    blockers.Add("Exact reporttypecode unavailable/malformed");
                if (!row.RuntimeColumns.Contains("languagecode") || !int.TryParse(row.Get("languagecode"), NumberStyles.Integer, CultureInfo.InvariantCulture, out language) || language <= 0)
                    blockers.Add("Exact languagecode unavailable/blank/malformed; no signaturelcid fallback");
                row.ProposedBlockingReason = blockers.Count == 0 ? "None; InternalIdentifierHypothesis only, field semantics and lifecycle portability require live validation" : string.Join("; ", blockers);
                if (blockers.Count == 0) row.CandidateP = "report-candidate-p:v1:" + Frame(ProposedIdentifierField, identifier.Trim(), row.ProposedScope,
                    type.ToString(CultureInfo.InvariantCulture), language.ToString(CultureInfo.InvariantCulture));
            }
            foreach (var group in side.Rows.Values.Where(r => r.CandidateP != null).GroupBy(r => r.CandidateP, StringComparer.OrdinalIgnoreCase).Where(g => g.Count() > 1))
                foreach (var row in group) row.DuplicateP = true;
        }

        private static bool LogicalScope(string value) => !string.IsNullOrWhiteSpace(value) && value.Trim().Length <= 128 &&
            !StringComparer.OrdinalIgnoreCase.Equals(value.Trim(), "none") && (char.IsLetter(value.Trim()[0]) || value.Trim()[0] == '_') &&
            value.Trim().All(c => char.IsLetterOrDigit(c) || c == '_');

        private static Type31SideEvidence Read(IOrganizationService service, MembershipSnapshot snapshot,
            string version, CancellationToken token, Action<string> progress)
        {
            if (service == null) throw new ArgumentNullException(nameof(service));
            var side = new Type31SideEvidence { Snapshot = snapshot, Version = version, CapturedUtc = DateTimeOffset.UtcNow };
            side.Raw.AddRange(snapshot.Components.Where(c => c.Record.ComponentType == 31));
            var ids = side.Raw.Where(c => c.Record.ObjectId.HasValue && c.Record.ObjectId.Value != Guid.Empty)
                .Select(c => c.Record.ObjectId.Value).Distinct().OrderBy(id => id).ToArray();
            foreach (var id in ids) side.Rows.Add(id, new Type31ReportEvidence { ObjectId = id, Status = "Incomplete", Reason = "Schema not verified" });
            if (ids.Length == 0) return side;
            EntityMetadata reportMetadata = null;
            token.ThrowIfCancellationRequested();
            side.Requests.Add("Execute RetrieveEntity(report, Attributes|Relationships, RetrieveAsIfPublished=False)");
            try
            {
                progress?.Invoke(snapshot.Environment.DisplayName + ": validating readable report fields");
                var response = service.Execute(new RetrieveEntityRequest
                { LogicalName = "report", EntityFilters = EntityFilters.Attributes | EntityFilters.Relationships, RetrieveAsIfPublished = false }) as RetrieveEntityResponse;
                token.ThrowIfCancellationRequested();
                var metadata = response?.EntityMetadata;
                if (metadata?.Attributes == null || metadata.LogicalName != "report" || metadata.PrimaryIdAttribute != PrimaryId ||
                    metadata.Attributes.Any(a => a == null || string.IsNullOrWhiteSpace(a.LogicalName)) ||
                    metadata.Attributes.GroupBy(a => a.LogicalName, StringComparer.OrdinalIgnoreCase).Any(g => g.Count() != 1))
                    throw new InvalidOperationException("Incomplete/conflicting schema");
                side.PrimaryName = metadata.PrimaryNameAttribute;
                foreach (var attribute in metadata.Attributes.OrderBy(a => a.LogicalName, StringComparer.Ordinal))
                {
                    string name = attribute.LogicalName;
                    bool text = attribute.AttributeType == AttributeTypeCode.String || attribute.AttributeType == AttributeTypeCode.Memo;
                    bool uniqueId = attribute.AttributeType == AttributeTypeCode.Uniqueidentifier && name.Contains("unique");
                    // Never retrieve attachments/binary payloads. Unknown text content is hash-only.
                    bool content = text && (ContentFields.Contains(name) || attribute.AttributeType == AttributeTypeCode.Memo ||
                        new[] { "body", "subject", "content", "xml", "html", "text", "rdl", "query", "layout", "definition" }.Any(part => name.Contains(part)));
                    bool desired = AuditFields.Contains(name) || uniqueId || content;
                    side.MetadataFields.Add(name + ";type=" + attribute.AttributeType + ";metadataReadable=" + attribute.IsValidForRead + ";selected=" + desired);
                    if (!desired) continue;
                    bool validShape = name == PrimaryId || name == "signatureid" || uniqueId ? attribute.AttributeType == AttributeTypeCode.Uniqueidentifier :
                        content || name == "name" || name == "filename" || name == "categories" || name == "uniquename" || name == "schemaname" ? text :
                        name == "signaturedate" ? attribute.AttributeType == AttributeTypeCode.DateTime :
                        new[] { "ownerid", "parentreportid", "originalreportid", "organizationid", "solutionid", "owninguser", "owningteam" }.Contains(name) ? attribute.AttributeType == AttributeTypeCode.Lookup || attribute.AttributeType == AttributeTypeCode.Owner :
                        name == "ismanaged" || name == "ispersonal" || name == "iscustomreport" ? attribute.AttributeType == AttributeTypeCode.Boolean :
                        attribute.AttributeType == AttributeTypeCode.Integer || attribute.AttributeType == AttributeTypeCode.Picklist ||
                        attribute.AttributeType == AttributeTypeCode.State || attribute.AttributeType == AttributeTypeCode.Status ||
                        (name == "relatedentities" || name == "objecttypecode") && (text || attribute.AttributeType == AttributeTypeCode.EntityName);
                    bool allowed = attribute.IsValidForRead == true && validShape &&
                        !new[] { "attachment", "binary", "base64", "encoded" }.Any(part => name.Contains(part));
                    side.Schema[name] = allowed ? "Readable" : "UnavailableOrUnverified";
                    if (!allowed) continue;
                    side.Columns.Add(name);
                    if (content) side.HashFields.Add(name);
                    if (uniqueId) side.UniqueFields.Add(name);
                    if (name == ProposedIdentifierField) side.FixedIdentifierVerified = attribute.AttributeType == AttributeTypeCode.String;
                }
                if (!side.Columns.Contains(PrimaryId)) throw new InvalidOperationException("Primary ID is not readable");
                reportMetadata = metadata;
            }
            catch (OperationCanceledException) { throw; }
            catch (Exception ex)
            {
                token.ThrowIfCancellationRequested();
                foreach (var row in side.Rows.Values) { row.Status = ex is InvalidOperationException ? "Incomplete" : "Faulted"; row.Reason = "Schema unavailable/faulted/incomplete; server details withheld. No guessed columns requested."; }
                return side;
            }

            for (int offset = 0; offset < ids.Length; offset += BatchSize)
            {
                token.ThrowIfCancellationRequested();
                var batch = ids.Skip(offset).Take(BatchSize).ToArray();
                var audit = new Type31BatchEvidence { BatchNumber = offset / BatchSize + 1, RequestedIdCount = batch.Length };
                side.Batches.Add(audit);
                var uniqueRows = new Dictionary<Guid, Entity>();
                var duplicateIds = new HashSet<Guid>();
                string cookie = null;
                int pageNumber = 1;
                try
                {
                    while (true)
                    {
                        token.ThrowIfCancellationRequested();
                        var query = new QueryExpression("report")
                        {
                            ColumnSet = new ColumnSet(side.Columns.ToArray()),
                            PageInfo = new PagingInfo { Count = BatchSize, PageNumber = pageNumber, PagingCookie = cookie }
                        };
                        query.Criteria.AddCondition(new ConditionExpression(PrimaryId, ConditionOperator.In, batch.Select(id => (object)id).ToArray()));
                        query.AddOrder(PrimaryId, OrderType.Ascending);
                        side.Requests.Add("RetrieveMultiple report; columns=[" + string.Join(",", side.Columns) + "]; reportid IN Guid[" + batch.Length + "]; page=" + pageNumber);
                        var page = new Type31PageEvidence { PageNumber = pageNumber };
                        audit.Pages.Add(page);
                        progress?.Invoke(snapshot.Environment.DisplayName + ": reading report batch " + audit.BatchNumber + ", page " + pageNumber);
                        var response = service.RetrieveMultiple(query);
                        token.ThrowIfCancellationRequested();
                        if (response == null) { Fail(side, batch, "Incomplete", "Null result set; batch completion not proven"); break; }
                        page.RowsReturned = response.Entities.Count;
                        page.MoreRecords = response.MoreRecords;
                        page.PagingCookieSupplied = !string.IsNullOrEmpty(response.PagingCookie);
                        audit.ReturnedRows += response.Entities.Count;
                        side.ReturnedRows += response.Entities.Count;
                        foreach (var id in batch)
                        {
                            var observed = response.Entities.Where(r => r != null && (r.Id == id || Id(r, PrimaryId) == id)).ToArray();
                            var observedIds = observed.Select(r => Id(r, PrimaryId)).Distinct().ToArray();
                            if (observedIds.Length == 1) side.Rows[id].ReportId = observedIds[0];
                        }
                        if (response.Entities.Any(r => r == null || r.LogicalName != "report" || !Id(r, PrimaryId).HasValue ||
                            r.Id == Guid.Empty || Id(r, PrimaryId) != r.Id || !batch.Contains(r.Id)))
                        { Fail(side, batch, "Incomplete", "Blank/conflicting primary ID or unexpected entity/result; ObjectId equality not proven"); break; }

                        int previousCount = uniqueRows.Count;
                        foreach (var group in response.Entities.GroupBy(r => Id(r, PrimaryId).Value))
                        {
                            // Preserve duplicate-returned-row safeguards within a page. An identical
                            // overlap on another page is paging duplication only; conflicts never collapse.
                            if (group.Count() > 1) duplicateIds.Add(group.Key);
                            foreach (var row in group)
                            {
                                if (uniqueRows.TryGetValue(group.Key, out var previous))
                                {
                                    if (!SameReturnedRow(previous, row)) duplicateIds.Add(group.Key);
                                }
                                else uniqueRows.Add(group.Key, row);
                            }
                            var evidence = side.Rows[group.Key];
                            evidence.BackingRowCount = duplicateIds.Contains(group.Key) ? Math.Max(2, evidence.BackingRowCount + group.Count()) : 1;
                        }
                        audit.DistinctReturnedIds = uniqueRows.Count;
                        // A cookie alone does not mean more data exists. Only MoreRecords asks us
                        // to continue; no-cookie responses use SDK simple paging with PageNumber.
                        if (!response.MoreRecords)
                        {
                            foreach (var id in batch)
                            {
                                var row = side.Rows[id];
                                row.Status = !uniqueRows.ContainsKey(id) ? "Missing" : duplicateIds.Contains(id) ? "Duplicate" : "Unique";
                                row.Reason = row.Status == "Missing" ? "Backing record not returned; not a membership absence finding" :
                                    row.Status == "Duplicate" ? "Duplicate/conflicting returned primary-key rows; no candidate evaluated" : "Exact ObjectId/reportid/Entity.Id correlation; terminal page received";
                                if (row.Status == "Unique") row.Capture(uniqueRows[id], side);
                            }
                            audit.Complete = true;
                            break;
                        }
                        if (uniqueRows.Count == previousCount || (!string.IsNullOrEmpty(response.PagingCookie) && response.PagingCookie == cookie))
                        { Fail(side, batch, "Incomplete", "Paging made no distinct-ID progress or repeated its continuation cookie; batch completion not proven"); break; }
                        cookie = response.PagingCookie;
                        pageNumber++;
                    }
                }
                catch (OperationCanceledException) { throw; }
                catch (Exception)
                {
                    token.ThrowIfCancellationRequested();
                    Fail(side, batch, "Faulted", "Report retrieval failed; server details withheld");
                }
            }
            Type31RelatedScopeReader.Read(service, side, reportMetadata, token, progress);
            return side;
        }

        internal static bool SameReturnedRow(Entity left, Entity right) => left.Attributes.Count == right.Attributes.Count &&
            left.Attributes.All(a => right.Attributes.TryGetValue(a.Key, out var value) && SameValue(a.Value, value));
        private static bool SameValue(object left, object right)
        {
            if (left is OptionSetValue && right is OptionSetValue) return ((OptionSetValue)left).Value == ((OptionSetValue)right).Value;
            if (left is EntityReference && right is EntityReference)
                return ((EntityReference)left).Id == ((EntityReference)right).Id &&
                    StringComparer.OrdinalIgnoreCase.Equals(((EntityReference)left).LogicalName, ((EntityReference)right).LogicalName);
            return Equals(left, right);
        }

        private static void Fail(Type31SideEvidence side, IEnumerable<Guid> ids, string status, string reason)
        { foreach (var id in ids) { side.Rows[id].Status = status; side.Rows[id].Reason = reason; } }
        internal static Guid? Id(Entity row, string field) => row.Attributes.TryGetValue(field, out var raw) && raw is Guid && (Guid)raw != Guid.Empty ? (Guid?)raw : null;
        internal static string Frame(params string[] values) => string.Concat(values.Select(value => value.Length.ToString(CultureInfo.InvariantCulture) + ":" + value + ":"));
        internal static string Safe(object value) => (Convert.ToString(value, CultureInfo.InvariantCulture) ?? "Unknown")
            .Replace("\r", "\\r").Replace("\n", "\\n").Replace("\t", "\\t");
    }

    internal sealed class Type31SideEvidence
    {
        internal MembershipSnapshot Snapshot;
        internal string Version;
        internal DateTimeOffset CapturedUtc;
        internal int ReturnedRows;
        internal string PrimaryName;
        internal bool FixedIdentifierVerified;
        internal readonly List<string> MetadataFields = new List<string>();
        internal readonly List<ComponentIdentity> Raw = new List<ComponentIdentity>();
        internal readonly SortedDictionary<Guid, Type31ReportEvidence> Rows = new SortedDictionary<Guid, Type31ReportEvidence>();
        internal readonly List<string> Columns = new List<string>();
        internal readonly HashSet<string> HashFields = new HashSet<string>(StringComparer.Ordinal);
        internal readonly HashSet<string> UniqueFields = new HashSet<string>(StringComparer.Ordinal);
        internal readonly SortedDictionary<string, string> Schema = new SortedDictionary<string, string>(StringComparer.Ordinal);
        internal readonly List<string> Requests = new List<string>();
        internal readonly List<Type31BatchEvidence> Batches = new List<Type31BatchEvidence>();
        internal readonly List<string> ScopeSchema = new List<string>();
        internal int ScopeMetadataRequests, ScopeQueries;
        internal bool Complete => Raw.All(r => r.Record.ObjectId.HasValue && r.Record.ObjectId.Value != Guid.Empty) && Rows.Values.All(r => r.Status == "Unique" && r.CandidateA != null);
    }

    internal sealed class Type31BatchEvidence
    {
        internal int BatchNumber, RequestedIdCount, ReturnedRows, DistinctReturnedIds;
        internal bool Complete;
        internal readonly List<Type31PageEvidence> Pages = new List<Type31PageEvidence>();
    }

    internal sealed class Type31PageEvidence
    {
        internal int PageNumber;
        internal int? RowsReturned;
        internal bool? MoreRecords, PagingCookieSupplied;
    }

    internal sealed class Type31ContentFingerprint
    {
        internal string Presence;
        internal int? Length;
        internal string Sha256;
        internal static Type31ContentFingerprint Create(object raw)
        {
            if (raw != null && !(raw is string)) return new Type31ContentFingerprint { Presence = "Malformed" };
            string text = raw as string ?? string.Empty;
            using (var sha = SHA256.Create())
                return new Type31ContentFingerprint { Presence = text.Length == 0 ? "Blank" : "Present", Length = text.Length,
                    Sha256 = BitConverter.ToString(sha.ComputeHash(Encoding.UTF8.GetBytes(text))).Replace("-", string.Empty) };
        }
        internal bool Known => Sha256 != null;
        internal string Evidence => "presence=" + Presence + ";length=" + (Length?.ToString(CultureInfo.InvariantCulture) ?? "Unknown") + ";sha256=" + (Sha256 ?? "Unknown");
    }

    internal sealed class Type31ReportEvidence
    {
        internal Guid ObjectId;
        internal Guid? ReportId;
        internal Guid? SignatureId;
        internal int BackingRowCount;
        internal string Status, Reason, CandidateA, CandidateB;
        internal bool DuplicateA, DuplicateB;
        internal bool DuplicateSignature;
        internal string RelatedScope, ScopeStatus = "NotRequired", ScopeReason;
        internal string ReportSubset, BoundaryReason, CandidateP, ProposedBlockingReason, ProposedScope,
            ProposedScopeStatus = "Incomplete", ProposedScopeReason;
        internal bool ProductionSignedSnapshot, DuplicateP;
        internal readonly List<string> RuntimeColumns = new List<string>();
        // No Entity/raw subject/body is retained in the report model.
        internal readonly SortedDictionary<string, string> Fields = new SortedDictionary<string, string>(StringComparer.Ordinal);
        internal readonly SortedDictionary<string, Type31ContentFingerprint> Content = new SortedDictionary<string, Type31ContentFingerprint>(StringComparer.Ordinal);
        internal readonly SortedDictionary<string, Guid?> UniqueIds = new SortedDictionary<string, Guid?>(StringComparer.Ordinal);
        internal bool? Managed;
        internal string Get(string field) => Fields.TryGetValue(field, out var value) ? value : null;

        internal void Capture(Entity row, Type31SideEvidence side)
        {
            ReportId = Type31EvidenceCollector.Id(row, Type31EvidenceCollector.PrimaryId);
            SignatureId = Type31EvidenceCollector.Id(row, "signatureid");
            RuntimeColumns.AddRange(side.Columns);
            foreach (var field in side.Columns)
            {
                row.Attributes.TryGetValue(field, out var raw);
                if (side.HashFields.Contains(field)) { Content[field] = Type31ContentFingerprint.Create(raw); continue; }
                if (side.UniqueFields.Contains(field)) { UniqueIds[field] = Type31EvidenceCollector.Id(row, field); continue; }
                string value = null;
                if (field == "name" || field == "uniquename" || field == "schemaname") value = raw is string ? ((string)raw).Trim() : raw == null ? "" : null;
                else if (field == Type31EvidenceCollector.PrimaryId) value = ReportId?.ToString("D");
                else if (field == "signatureid") value = SignatureId?.ToString("D") ?? (raw == null || raw is Guid && (Guid)raw == Guid.Empty ? "" : null);
                else if ((field == "filename" || field == "categories") && raw is string) value = (string)raw;
                else if (raw == null) value = "";
                else if (raw is OptionSetValue) value = ((OptionSetValue)raw).Value.ToString(CultureInfo.InvariantCulture);
                else if (raw is int) value = ((int)raw).ToString(CultureInfo.InvariantCulture);
                else if (raw is bool) value = (bool)raw ? "True" : "False";
                else if (field == "signaturedate" && raw is DateTime) value = ((DateTime)raw).ToString("O", CultureInfo.InvariantCulture);
                else if ((field == "relatedentities" || field == "objecttypecode") && raw is string)
                {
                    var scope = ((string)raw).Trim();
                    if (scope.Length <= 2048 && scope.All(c => char.IsLetterOrDigit(c) || c == '_' || c == ',' || char.IsWhiteSpace(c)))
                        value = string.Join(",", scope.Split(',').Select(s => s.Trim()).Where(s => s.Length > 0)
                            .Distinct(StringComparer.OrdinalIgnoreCase).OrderBy(s => s, StringComparer.OrdinalIgnoreCase));
                }
                else if (raw is EntityReference)
                {
                    var lookup = (EntityReference)raw;
                    if (lookup.Id != Guid.Empty && !string.IsNullOrWhiteSpace(lookup.LogicalName) &&
                        lookup.LogicalName.All(c => char.IsLetterOrDigit(c) || c == '_')) value = lookup.LogicalName + ":" + lookup.Id.ToString("D");
                }
                Fields[field] = value;
            }
            Managed = row.Attributes.TryGetValue("ismanaged", out var managed) && managed is bool ? (bool?)managed : null;
        }

        internal void ConstructCandidates()
        {
            string scope = RelatedScope ?? Get("relatedentities"), title = Get("name"), language = Get("languagecode"), type = Get("reporttypecode");
            if (string.IsNullOrWhiteSpace(scope)) scope = Get("objecttypecode");
            if (string.IsNullOrWhiteSpace(language)) language = Get("signaturelcid");
            if (string.IsNullOrWhiteSpace(scope) || string.IsNullOrWhiteSpace(title)) return;
            CandidateB = "report-candidate-b:" + Type31EvidenceCollector.Frame(scope, title);
            if (string.IsNullOrWhiteSpace(language) || string.IsNullOrWhiteSpace(type)) return;
            CandidateA = "report-candidate-a:" + Type31EvidenceCollector.Frame(scope, type, language, title);
        }
    }

    internal sealed class Type31LifecycleEntry
    {
        internal readonly List<Type31ReportEvidence> Source = new List<Type31ReportEvidence>();
        internal readonly List<Type31ReportEvidence> Target = new List<Type31ReportEvidence>();
        internal readonly HashSet<string> Outcomes = new HashSet<string>(StringComparer.Ordinal);
        internal string Basis;
        internal bool Pair => Source.Count == 1 && Target.Count == 1 && Source[0].Status == "Unique" && Target[0].Status == "Unique" && !Outcomes.Contains("Ambiguous");
        internal bool OnlyUniqueIdDifference;
    }

    internal sealed class Type31EvidenceReport
    {
        internal Type31SideEvidence Source, Target;
        internal readonly List<Type31LifecycleEntry> Lifecycle = new List<Type31LifecycleEntry>();
        internal static readonly string[] Categories = { "SamePrimaryId", "DifferentPrimaryId", "SameUniqueId", "DifferentUniqueId",
            "SameSemanticCandidate", "DifferentSemanticCandidate", "SameContent", "DifferentContent", "UnmanagedToManaged", "ManagedTransition", "OneSidedEvidence", "Ambiguous", "Incomplete" };

        internal void Analyze(CancellationToken token)
        {
            var nodes = Source.Rows.Values.Where(r => r.Status == "Unique").Select(r => new Node { Row = r, Source = true })
                .Concat(Target.Rows.Values.Where(r => r.Status == "Unique").Select(r => new Node { Row = r })).ToArray();
            var edges = nodes.ToDictionary(n => n, n => new HashSet<Node>());
            // Two evidence relationships only: observed primary-ID overlap and Candidate A.
            // Connected conflicts cannot be repaired with GUIDs or weaker Candidate B.
            foreach (var group in nodes.Where(n => n.Row.CandidateA != null).GroupBy(n => n.Row.CandidateA, StringComparer.OrdinalIgnoreCase))
                Connect(group.ToArray(), edges);
            foreach (var group in nodes.GroupBy(n => n.Row.ReportId).Where(g => g.Any(n => n.Source) && g.Any(n => !n.Source)))
                Connect(group.ToArray(), edges);
            var visited = new HashSet<Node>();
            foreach (var node in nodes)
            {
                token.ThrowIfCancellationRequested();
                if (!visited.Add(node)) continue;
                var pending = new Queue<Node>(); pending.Enqueue(node);
                var entry = new Type31LifecycleEntry();
                while (pending.Count > 0)
                {
                    token.ThrowIfCancellationRequested();
                    var current = pending.Dequeue();
                    (current.Source ? entry.Source : entry.Target).Add(current.Row);
                    foreach (var next in edges[current]) if (visited.Add(next)) pending.Enqueue(next);
                }
                if (entry.Source.Count > 1 || entry.Target.Count > 1 || entry.Source.Concat(entry.Target).Any(r => r.DuplicateA))
                { entry.Outcomes.Add("Ambiguous"); entry.Basis = "Candidate A collision or contradictory primary/semantic relationships; no automatic pairing"; }
                else if (entry.Source.Count == 0 || entry.Target.Count == 0)
                {
                    entry.Outcomes.Add("OneSidedEvidence"); entry.Basis = "Observed on one snapshot only; not Source Only/Target Only membership";
                    if (!Source.Complete || !Target.Complete) entry.Outcomes.Add("Incomplete");
                }
                else AnalyzePair(entry);
                Lifecycle.Add(entry);
            }
            foreach (var side in new[] { Source, Target })
            {
                foreach (var row in side.Rows.Values.Where(r => r.Status != "Unique"))
                {
                    var entry = new Type31LifecycleEntry { Basis = "Backing correlation " + row.Status };
                    (side == Source ? entry.Source : entry.Target).Add(row);
                    entry.Outcomes.Add(row.Status == "Duplicate" ? "Ambiguous" : "Incomplete"); Lifecycle.Add(entry);
                }
                foreach (var raw in side.Raw.Where(c => !c.Record.ObjectId.HasValue || c.Record.ObjectId.Value == Guid.Empty))
                    Lifecycle.Add(new Type31LifecycleEntry { Basis = (side == Source ? "Source" : "Target") + " blank ObjectId", Outcomes = { "Incomplete" } });
            }
        }

        private sealed class Node { internal Type31ReportEvidence Row; internal bool Source; }
        private static void Connect(Node[] group, Dictionary<Node, HashSet<Node>> edges)
        {
            // Star topology keeps duplicate-group analysis linear instead of N squared.
            for (int i = 1; i < group.Length; i++) { edges[group[0]].Add(group[i]); edges[group[i]].Add(group[0]); }
        }

        private void AnalyzePair(Type31LifecycleEntry entry)
        {
            var left = entry.Source[0]; var right = entry.Target[0];
            bool samePrimary = left.ReportId == right.ReportId;
            entry.Basis = samePrimary ? "Primary-ID audit observation (not identity approval)" : "Unique Candidate A evidence pair";
            entry.Outcomes.Add(samePrimary ? "SamePrimaryId" : "DifferentPrimaryId");
            if (left.CandidateA == null || right.CandidateA == null) entry.Outcomes.Add("Incomplete");
            else entry.Outcomes.Add(StringComparer.OrdinalIgnoreCase.Equals(left.CandidateA, right.CandidateA) ? "SameSemanticCandidate" : "DifferentSemanticCandidate");
            var uniqueFields = left.UniqueIds.Keys.Union(right.UniqueIds.Keys).ToArray();
            bool uniqueKnown = uniqueFields.Length > 0 && uniqueFields.All(f => left.UniqueIds.ContainsKey(f) && right.UniqueIds.ContainsKey(f) && left.UniqueIds[f].HasValue && right.UniqueIds[f].HasValue);
            bool uniqueDifferent = uniqueFields.Any(f => left.UniqueIds.ContainsKey(f) && right.UniqueIds.ContainsKey(f) && left.UniqueIds[f].HasValue && right.UniqueIds[f].HasValue && left.UniqueIds[f] != right.UniqueIds[f]);
            if (uniqueDifferent) entry.Outcomes.Add("DifferentUniqueId");
            else if (uniqueKnown) entry.Outcomes.Add("SameUniqueId");
            string content = ContentComparison(left, right);
            if (content != "Unknown") entry.Outcomes.Add(content); // Missing content is never SameContent.
            if (left.Managed.HasValue && right.Managed.HasValue && left.Managed != right.Managed)
            {
                entry.Outcomes.Add("ManagedTransition");
                if (left.Managed == false && right.Managed == true) entry.Outcomes.Add("UnmanagedToManaged");
            }
            var fields = left.Fields.Keys.Union(right.Fields.Keys).ToArray();
            entry.OnlyUniqueIdDifference = samePrimary && uniqueDifferent && uniqueKnown && content == "SameContent" &&
                fields.All(f => left.Fields.ContainsKey(f) && right.Fields.ContainsKey(f) && left.Fields[f] != null && right.Fields[f] != null &&
                    StringComparer.OrdinalIgnoreCase.Equals(left.Fields[f], right.Fields[f]));
        }

        internal static string ContentComparison(Type31ReportEvidence left, Type31ReportEvidence right)
        {
            var fields = left.Content.Keys.Union(right.Content.Keys).ToArray();
            if (fields.Any(f => left.Content.ContainsKey(f) && right.Content.ContainsKey(f) && left.Content[f].Known && right.Content[f].Known && left.Content[f].Sha256 != right.Content[f].Sha256)) return "DifferentContent";
            return fields.Length > 0 && fields.All(f => left.Content.ContainsKey(f) && right.Content.ContainsKey(f) && left.Content[f].Known && right.Content[f].Known)
                ? "SameContent" : "Unknown";
        }

        private static void Line(StringBuilder text, params object[] cells) => text.AppendLine(string.Join("\t", cells.Select(Type31EvidenceCollector.Safe)));
        private static string Ids(IEnumerable<Type31ReportEvidence> rows) => string.Join(",", rows.Select(r => r.ObjectId).OrderBy(id => id).Select(id => id.ToString("D")));
        internal string Build()
        {
            var text = new StringBuilder();
            text.AppendLine("TYPE 31 REPORT EVIDENCE - DEBUG ONLY");
            text.AppendLine("Evidence only. Type 31 outside the verified signed-report subset remains Unsupported / Indeterminate. No new production identity, membership matching, definition contract or absence inference.");
            text.AppendLine("Selected completed solution snapshots only. Raw GUID overlap is not proof of portability. Primary-ID pairs are audit observations.");
            text.AppendLine("Evidence only. Existing verified signed-report production resolution is preserved; unsupported Type 31 reports receive no identity or absence proof from this capture.");
            text.AppendLine("Candidate A: related entity scope + report type + language (languagecode, otherwise signaturelcid) + name. Candidate B: scope + name. Trim + ordinal case-insensitive; B never repairs A ambiguity. Names/localized labels and numeric scope portability remain unproven.");
            text.AppendLine("Signature hypothesis: signatureid alone, evaluated independently. signaturelcid is locale audit/context only and never repairs duplicate signatures. These semantic candidates do not change the verified signed-report resolver.");
            text.AppendLine("Report definition/content: only presence, UTF-16 character length and exact UTF-8 SHA-256; no raw XML/RDL/query/body/layout, attachments, binary or encoded payloads.");
            text.AppendLine("RAW TYPE 31 MEMBERSHIP");
            Line(text, "Side", "Environment", "Solution", "Version", "SnapshotUtc", "SolutionComponentId", "ObjectId", "ProductionResolutionStatus", "ProductionDiagnostic", "ExistingPortableKey", "ProductionKind", "ReportSubset");
            foreach (var side in new[] { Source, Target })
            {
                string label = side == Source ? "Source" : "Target";
                foreach (var raw in side.Raw.OrderBy(r => r.Record.SolutionComponentId))
                    Line(text, label, side.Snapshot.Environment.DisplayName, side.Snapshot.SolutionUniqueName, side.Version,
                        side.Snapshot.CapturedAt.UtcDateTime.ToString("O", CultureInfo.InvariantCulture), raw.Record.SolutionComponentId,
                        raw.Record.ObjectId?.ToString("D") ?? "Blank", raw.Status, raw.Diagnostic, raw.ComparisonKey ?? "(none)", raw.SemanticKind, Subset(raw));
                Line(text, label, "raw=" + side.Raw.Count, "distinctNonblankIds=" + side.Rows.Count,
                    "blankObjectIds=" + side.Raw.Count(r => !r.Record.ObjectId.HasValue || r.Record.ObjectId.Value == Guid.Empty));
                foreach (var group in side.Raw.GroupBy(Subset).OrderBy(g => g.Key, StringComparer.Ordinal))
                    Line(text, label, "ProductionReportSubset", group.Key, "RawMembershipCount=" + group.Count());
            }
            text.AppendLine("BACKING REPORT CORRELATION");
            foreach (var side in new[] { Source, Target })
            {
                string label = side == Source ? "Source" : "Target";
                Line(text, label, "CapturedUtc=" + side.CapturedUtc.UtcDateTime.ToString("O", CultureInfo.InvariantCulture), "ReadableColumns=[" + string.Join(",", side.Columns) + "]", "ReturnedRows=" + side.ReturnedRows);
                foreach (var batch in side.Batches)
                {
                    Line(text, label, "Batch=" + batch.BatchNumber, "RequestedIdCount=" + batch.RequestedIdCount,
                        "RowsReturned=" + batch.ReturnedRows, "DistinctReturnedIds=" + batch.DistinctReturnedIds,
                        "PageCount=" + batch.Pages.Count, "RetrievalComplete=" + batch.Complete);
                    foreach (var page in batch.Pages)
                        Line(text, label, "Batch=" + batch.BatchNumber, "Page=" + page.PageNumber,
                            "RowsReturned=" + (page.RowsReturned?.ToString(CultureInfo.InvariantCulture) ?? "Unknown"),
                            "MoreRecords=" + (page.MoreRecords?.ToString() ?? "Unknown"),
                            "PagingCookieSupplied=" + (page.PagingCookieSupplied?.ToString() ?? "Unknown"));
                }
                Line(text, "Side", "ObjectId", "BackingRowCount", "CorrelationStatus", "ObjectIdEqualsReportId", "ReportId", "Reason", "CandidateA", "CandidateB", "DuplicateA", "DuplicateB");
                foreach (var row in side.Rows.Values)
                {
                    Line(text, label, row.ObjectId, row.BackingRowCount, row.Status,
                        row.ReportId.HasValue ? (row.ReportId == row.ObjectId ? "True" : "False") : "NotProven",
                        row.ReportId?.ToString("D") ?? "Unknown", row.Reason, row.CandidateA ?? "Incomplete", row.CandidateB ?? "Incomplete", row.DuplicateA, row.DuplicateB);
                    foreach (var field in row.Fields) Line(text, label, row.ObjectId, field.Key, field.Value ?? "Malformed");
                    foreach (var field in row.UniqueIds) Line(text, label, row.ObjectId, field.Key, field.Value?.ToString("D") ?? "BlankOrMalformed");
                    foreach (var field in row.Content) Line(text, label, row.ObjectId, field.Key, field.Value.Evidence);
                    Line(text, label, row.ObjectId, "RelatedScopeStatus=" + row.ScopeStatus,
                        "RelatedEntityScope=" + (row.RelatedScope ?? "Unavailable"), "ScopeReason=" + row.ScopeReason);
                }
                foreach (var raw in side.Raw.Where(r => !r.Record.ObjectId.HasValue || r.Record.ObjectId.Value == Guid.Empty))
                    Line(text, label, "Blank", 0, "Incomplete", "NotProven", "Unknown", "Blank ObjectId; no backing query",
                        "Incomplete", "Incomplete", false, false);
            }
            text.AppendLine("READABLE / UNAVAILABLE SCHEMA");
            foreach (var side in new[] { Source, Target })
            {
                string label = side == Source ? "Source" : "Target";
                foreach (var attribute in side.Schema) Line(text, label, "Schema", attribute.Key, attribute.Value);
                foreach (var field in Type31EvidenceCollector.AuditFields.Concat(Type31EvidenceCollector.ContentFields).Except(side.Schema.Keys).OrderBy(f => f, StringComparer.Ordinal))
                    Line(text, label, "Schema", field, "UnavailableInMetadataOrNotQueried");
                foreach (var evidence in side.ScopeSchema) Line(text, label, "RelatedScopeSchema", evidence);
            }
            text.AppendLine("REPEATED / DUPLICATE ANALYSIS");
            foreach (var side in new[] { Source, Target })
            {
                string label = side == Source ? "Source" : "Target";
                foreach (var group in side.Raw.Where(r => r.Record.ObjectId.HasValue && r.Record.ObjectId.Value != Guid.Empty)
                    .GroupBy(r => r.Record.ObjectId.Value).Where(g => g.Count() > 1).OrderBy(g => g.Key))
                    Line(text, label, "RepeatedMembershipOnly", group.Key, "rawReferences=" + group.Count());
                Line(text, label, "DuplicateBackingIds=" + side.Rows.Values.Count(r => r.Status == "Duplicate"),
                    "CandidateACollisionGroups=" + side.Rows.Values.Where(r => r.DuplicateA).Select(r => r.CandidateA).Distinct(StringComparer.OrdinalIgnoreCase).Count(),
                    "CandidateBCollisionGroups=" + side.Rows.Values.Where(r => r.DuplicateB).Select(r => r.CandidateB).Distinct(StringComparer.OrdinalIgnoreCase).Count());
            }
            text.AppendLine("CANDIDATE IDENTITY ANALYSIS");
            foreach (var side in new[] { Source, Target })
            {
                string label = side == Source ? "Source" : "Target";
                foreach (var row in side.Rows.Values)
                    Line(text, label, row.ObjectId, "Correlation=" + row.Status,
                        "CandidateA=" + (row.CandidateA ?? "Incomplete"), "CandidateB=" + (row.CandidateB ?? "Incomplete"),
                        "CandidateADuplicate=" + row.DuplicateA, "CandidateBDuplicate=" + row.DuplicateB,
                        "CandidateSignatureId=" + (row.SignatureId?.ToString("D") ?? "BlankOrUnavailable"),
                        "SignatureCandidateStatus=" + (!row.SignatureId.HasValue ? "Noncandidate" : row.DuplicateSignature ? "NonuniqueUnsafe" : "UniqueObserved"),
                        "SignatureLcid=" + (row.Get("signaturelcid") ?? "Unknown"));
            }
            foreach (var signature in Source.Rows.Values.Where(r => r.Status == "Unique" && r.SignatureId.HasValue)
                .Select(r => r.SignatureId.Value).Intersect(Target.Rows.Values.Where(r => r.Status == "Unique" && r.SignatureId.HasValue).Select(r => r.SignatureId.Value)).OrderBy(id => id))
                Line(text, "SignatureIdOverlapAuditOnly", signature, "SourceCount=" + Source.Rows.Values.Count(r => r.SignatureId == signature),
                    "TargetCount=" + Target.Rows.Values.Count(r => r.SignatureId == signature));
            text.AppendLine("SOURCE / TARGET FIELD COMPARISON");
            Line(text, "SourceId", "TargetId", "PairBasis", "Field", "SourceEvidence", "TargetEvidence", "ObservedComparison");
            foreach (var entry in Lifecycle.Where(e => e.Pair))
            {
                var left = entry.Source[0]; var right = entry.Target[0];
                foreach (var field in left.Fields.Keys.Union(right.Fields.Keys).Union(left.UniqueIds.Keys).Union(right.UniqueIds.Keys).Union(left.Content.Keys).Union(right.Content.Keys).OrderBy(f => f, StringComparer.Ordinal))
                {
                    string l = Field(left, field), r = Field(right, field);
                    bool known = l != null && r != null;
                    Line(text, left.ObjectId, right.ObjectId, entry.Basis, field, l ?? "Unknown", r ?? "Unknown",
                        !known ? "InsufficientEvidence" : StringComparer.OrdinalIgnoreCase.Equals(l, r) ? "EqualObserved" : "DifferentObserved");
                }
            }
            text.AppendLine("LIFECYCLE CORRELATION MATRIX");
            text.AppendLine("Counts are diagnostic entries: a unique backing pair, a conflict group, or a one-sided/incomplete backing observation. Categories overlap; raw repeated references do not inflate counts.");
            foreach (var category in Categories) Line(text, category, Lifecycle.Count(e => e.Outcomes.Contains(category)));
            Line(text, "SourceIds", "TargetIds", "Basis", "Outcomes");
            foreach (var entry in Lifecycle) Line(text, Ids(entry.Source), Ids(entry.Target), entry.Basis, string.Join(",", entry.Outcomes.OrderBy(s => s, StringComparer.Ordinal)));
            text.AppendLine("DIFFERING PRIMARY-ID SEMANTIC PAIRS");
            var differing = Lifecycle.Where(e => e.Pair && e.Outcomes.Contains("DifferentPrimaryId") && e.Outcomes.Contains("SameSemanticCandidate")).ToArray();
            if (differing.Length == 0) text.AppendLine("No unique differing-primary-ID semantic pair observed in these snapshots.");
            foreach (var entry in differing)
                Line(text, Source.Snapshot.SolutionUniqueName, entry.Source[0].ObjectId, entry.Target[0].ObjectId,
                    entry.Source[0].CandidateA, "CandidateAUnique=True", "Content=" + ContentComparison(entry.Source[0], entry.Target[0]),
                    "SourceManaged=" + entry.Source[0].Managed, "TargetManaged=" + entry.Target[0].Managed);
            text.AppendLine("PRIMARY-ID PORTABILITY ASSESSMENT");
            Line(text, "DistinctUniqueBackingPairsObserved", Lifecycle.Count(e => e.Pair));
            Line(text, "RetainedSameReportId", Lifecycle.Count(e => e.Outcomes.Contains("SamePrimaryId")));
            Line(text, "DifferentReportId", Lifecycle.Count(e => e.Outcomes.Contains("DifferentPrimaryId")));
            Line(text, "OnlyInstallationSpecificIdDifferenceProven", Lifecycle.Count(e => e.OnlyUniqueIdDifference));
            Line(text, "ContentHashDifferencesObserved", Lifecycle.Count(e => e.Outcomes.Contains("DifferentContent")));
            Line(text, "SemanticCollisionsObserved", Source.Rows.Values.Any(r => r.DuplicateA) || Target.Rows.Values.Any(r => r.DuplicateA));
            Line(text, "AmbiguousBackingCorrelationsObserved", Source.Rows.Values.Concat(Target.Rows.Values).Any(r => r.Status == "Duplicate"));
            if (Lifecycle.Any(e => e.Outcomes.Contains("Ambiguous") || e.Outcomes.Contains("Incomplete")) || !Source.Complete || !Target.Complete)
                text.AppendLine("Neither identity is established. Resolve incomplete/colliding evidence before proposing a guarded resolver; no production approval.");
            else if (differing.Length > 0)
                text.AppendLine("Unique differing-ID semantic observations support further semantic-identity investigation; do not establish lifecycle stability or approve production identity.");
            else if (Lifecycle.Any(e => e.Pair))
                text.AppendLine("Observed primary-ID preservation supports further guarded primary-ID investigation only. Raw GUID overlap and one deployment sample do not prove general portability or approve production identity.");
            else text.AppendLine("Neither identity can be assessed: no complete unique backing pairs observed.");
            BuildUnsignedReadiness(text);
            text.AppendLine("REQUEST LEDGER");
            foreach (var side in new[] { Source, Target })
            {
                string label = side == Source ? "Source" : "Target";
                Line(text, label, "SchemaRequests=" + side.Requests.Count(r => r.StartsWith("Execute RetrieveEntity(report,", StringComparison.Ordinal)),
                    "ReportQueries=" + side.Requests.Count(r => r.StartsWith("RetrieveMultiple report;", StringComparison.Ordinal)), "WhoAmI=0", "Writes=0", "MembershipQueries=0");
                Line(text, label, "RelatedScopeMetadataRequests=" + side.ScopeMetadataRequests, "RelatedScopeQueries=" + side.ScopeQueries,
                    "AdditionalScopeRequests=" + (side.ScopeMetadataRequests + side.ScopeQueries));
                for (int i = 0; i < side.Requests.Count; i++) Line(text, label, i + 1, side.Requests[i]);
            }
            return text.ToString();
        }

        private void BuildUnsignedReadiness(StringBuilder text)
        {
            text.AppendLine("SIGNED SUBSET / UNSIGNED REMAINDER BOUNDARY");
            text.AppendLine("VerifiedSignedSubset requires agreement with the existing resolved signed production snapshot and uniquely correlated nonblank signature evidence. UnsignedOrUnverifiedRemainder never gains signed eligibility from a blank signature or this diagnostic. Production resolution is never changed by capture.");
            foreach (var side in new[] { Source, Target })
            {
                string label = side == Source ? "Source" : "Target";
                foreach (var row in side.Rows.Values) Line(text, label, row.ObjectId, "ReportSubset=" + row.ReportSubset,
                    "ProductionSignedSnapshot=" + row.ProductionSignedSnapshot, "BoundaryReason=" + row.BoundaryReason,
                    "signatureid=" + (row.Get("signatureid") ?? "UnavailableOrMalformed"));
                foreach (var raw in side.Raw.Where(c => !c.Record.ObjectId.HasValue || c.Record.ObjectId == Guid.Empty))
                    Line(text, label, raw.Record.SolutionComponentId, "ReportSubset=UnsignedOrUnverifiedRemainder", "Blank ObjectId; no signature/backing inference");
            }
            text.AppendLine("UNSIGNED IDENTIFIER / SCOPE / LANGUAGE ANALYSIS");
            text.AppendLine("Candidate P is diagnostic-only InternalIdentifierHypothesis: fixed readable string uniquename + independently verified logical table scope + exact reporttypecode + exact languagecode. No schemaname/name/filename/signaturelcid fallback. Metadata availability alone does not establish internal-identifier portability. Name/filename remain context, not independently portable identifiers; no name-only Candidate P is proposed.");
            foreach (var side in new[] { Source, Target })
            {
                string label = side == Source ? "Source" : "Target";
                Line(text, label, "PrimaryNameAttribute=" + (side.PrimaryName ?? "Unavailable"));
                foreach (var field in side.MetadataFields) Line(text, label, "Metadata", field);
                foreach (var row in side.Rows.Values.Where(r => !r.ProductionSignedSnapshot))
                {
                    Line(text, label, row.ObjectId, "Correlation=" + row.Status, "TableStatus=" + row.ProposedScopeStatus,
                        "TableKey=" + row.ProposedScope, "ScopeReason=" + row.ProposedScopeReason,
                        "CandidateP=" + (row.CandidateP ?? "Incomplete"), "CompleteP=" + (row.CandidateP != null && !row.DuplicateP),
                        "DuplicateP=" + row.DuplicateP, "BlockingReason=" + (row.DuplicateP ? "Candidate P collision; no automatic pairing" : row.ProposedBlockingReason));
                    foreach (var field in new[] { "name", "filename", "uniquename", "schemaname", "reporttypecode", "reportcategorycode", "languagecode", "ispersonal", "signatureid", "signaturedate", "signaturelcid" })
                        Line(text, label, row.ObjectId, field, "Schema=" + (side.Schema.TryGetValue(field, out var schema) ? schema : "NotExposed"),
                            "RuntimeRead=" + row.RuntimeColumns.Contains(field), "Value=" + (row.Get(field) ?? "UnavailableOrMalformed"));
                    foreach (var field in Type31EvidenceCollector.ContentFields)
                        Line(text, label, row.ObjectId, field, row.Content.TryGetValue(field, out var hash) ? hash.Evidence : "NotRetrievedOrUnavailable; presence/length/hash unknown; binary never requested");
                }
            }
            text.AppendLine("UNSIGNED CRITICAL COLLISION MATRIX");
            foreach (var side in new[] { Source, Target })
            {
                string label = side == Source ? "Source" : "Target";
                var rows = side.Rows.Values.Where(r => !r.ProductionSignedSnapshot && r.Status == "Unique").ToArray();
                Collision(text, label, "NameOnly", rows, r => Key(r.Get("name")));
                Collision(text, label, "FilenameOnly", rows, r => Key(r.Get("filename")));
                Collision(text, label, "Name+ReportType", rows, r => Key(r.Get("name"), r.Get("reporttypecode")));
                Collision(text, label, "Name+Language", rows, r => Key(r.Get("name"), r.Get("languagecode")));
                Collision(text, label, "Name+VerifiedScope", rows, r => r.ProposedScopeStatus == "Verified" ? Key(r.Get("name"), r.ProposedScope) : null);
                Collision(text, label, "CandidateP", rows, r => r.CandidateP);
            }
            text.AppendLine("UNSIGNED CANDIDATE P LIFECYCLE COMPARISON");
            var sourceGroups = Source.Rows.Values.Where(r => r.CandidateP != null).GroupBy(r => r.CandidateP, StringComparer.OrdinalIgnoreCase)
                .ToDictionary(g => g.Key, g => g.ToArray(), StringComparer.OrdinalIgnoreCase);
            var targetGroups = Target.Rows.Values.Where(r => r.CandidateP != null).GroupBy(r => r.CandidateP, StringComparer.OrdinalIgnoreCase)
                .ToDictionary(g => g.Key, g => g.ToArray(), StringComparer.OrdinalIgnoreCase);
            int pairs = 0, different = 0;
            foreach (var key in sourceGroups.Keys.Union(targetGroups.Keys, StringComparer.OrdinalIgnoreCase).OrderBy(k => k, StringComparer.OrdinalIgnoreCase))
            {
                var left = sourceGroups.TryGetValue(key, out var l) ? l : new Type31ReportEvidence[0];
                var right = targetGroups.TryGetValue(key, out var r) ? r : new Type31ReportEvidence[0];
                string outcome = left.Length > 1 || right.Length > 1 ? "Ambiguous" : left.Length == 1 && right.Length == 1 ? "InternalIdentifierPairHypothesis" : "OneSidedEvidence";
                Line(text, outcome, key, "SourceIds=" + Ids(left), "TargetIds=" + Ids(right));
                if (outcome != "InternalIdentifierPairHypothesis") continue;
                pairs++; bool differing = left[0].ReportId != right[0].ReportId; if (differing) different++;
                Line(text, "reportid=" + (differing ? "DifferentObserved" : "EqualObserved"),
                    "SourceSubset=" + left[0].ReportSubset, "TargetSubset=" + right[0].ReportSubset,
                    "SourceManaged=" + left[0].Managed, "TargetManaged=" + right[0].Managed,
                    "ManagedTransition=" + (left[0].Managed.HasValue && right[0].Managed.HasValue && left[0].Managed != right[0].Managed),
                    "Content=" + ContentComparison(left[0], right[0]));
                foreach (var field in left[0].Fields.Keys.Union(right[0].Fields.Keys).Union(left[0].Content.Keys).Union(right[0].Content.Keys).Union(left[0].UniqueIds.Keys).Union(right[0].UniqueIds.Keys).OrderBy(f => f, StringComparer.Ordinal))
                {
                    string first = Field(left[0], field), second = Field(right[0], field);
                    Line(text, field, "Source=" + (first ?? "Unavailable"), "Target=" + (second ?? "Unavailable"), first == null || second == null ? "Incomplete" :
                        StringComparer.OrdinalIgnoreCase.Equals(first, second) ? "EqualObserved" : "DifferentObserved");
                }
            }
            Line(text, "UniqueCandidatePPairs=" + pairs, "SamePrimaryId=" + (pairs - different), "DifferentPrimaryId=" + different);
            text.AppendLine("UNSIGNED DEFINITION EVIDENCE / RESOLVER READINESS");
            text.AppendLine("Name and filename are context/mutable definition hypotheses. Report type/category and language require semantic classification; type/language are retained conservatively in hypothetical Candidate P and cannot also become definition properties without a revised contract. Content hashes and signature fields/date/locale are definition/audit evidence only; none create unsigned identity. bodybinary/attachments are not retrieved.");
            text.AppendLine("Same reportid does not prove portability. An absent readable fixed internal identifier, unverified scope, missing language/type or collisions make Candidate P incomplete/ambiguous: recommend remaining unsupported. Even a complete diagnostic Candidate P needs independently deployed differing-primary-ID lifecycle evidence and semantic review; metadata/readability, matching hashes and managed transition do not alone make the unsigned remainder production-design-ready. No absence inference.");
        }

        private static string Key(params string[] fields) => fields.Any(string.IsNullOrWhiteSpace) ? null : Type31EvidenceCollector.Frame(fields.Select(f => f.Trim()).ToArray());
        private static void Collision(StringBuilder text, string side, string dimension, Type31ReportEvidence[] rows, Func<Type31ReportEvidence, string> select)
        {
            var keys = rows.Select(select).Where(k => k != null).ToArray();
            var collisions = keys.GroupBy(k => k, StringComparer.OrdinalIgnoreCase).Where(g => g.Count() > 1).ToArray();
            Line(text, side, dimension, "CollisionGroups=" + collisions.Length, "IncompleteDimension=" + (rows.Length - keys.Length));
            foreach (var group in collisions) Line(text, side, dimension, "Candidate=" + group.Key, "DistinctBackingRecords=" + group.Count());
        }

        private static string Field(Type31ReportEvidence row, string field)
        {
            if (row.Content.TryGetValue(field, out var content)) return content.Known ? content.Evidence : null;
            if (row.UniqueIds.TryGetValue(field, out var id)) return id?.ToString("D");
            return row.Get(field);
        }

        private static string Subset(ComponentIdentity identity) => identity.Status == IdentityResolutionStatus.Resolved &&
            StringComparer.OrdinalIgnoreCase.Equals(identity.SemanticKind, ComponentSemanticKinds.Report)
                ? "VerifiedSignedSubset" : identity.Status == IdentityResolutionStatus.Unsupported ? "OutsideVerifiedSignedSubset" : "ProductionIdentity" + identity.Status;
    }

    /// <summary>Discovers scope associations from report relationships; never guesses a backing table.</summary>
    internal static class Type31RelatedScopeReader
    {
        private sealed class Binding
        {
            internal string Entity, ForeignKey, Relationship, PrimaryId, ScopeField;
            internal bool Intersect;
        }

        internal static void Read(IOrganizationService service, Type31SideEvidence side, EntityMetadata reportMetadata,
            CancellationToken token, Action<string> progress)
        {
            var rows = side.Rows.Values.Where(r => r.Status == "Unique" &&
                string.IsNullOrWhiteSpace(r.Get("relatedentities")) && string.IsNullOrWhiteSpace(r.Get("objecttypecode"))).ToArray();
            if (rows.Length == 0) return;
            Set(rows, "Unavailable", "No metadata-proven related entity scope found; semantic candidates remain incomplete.");
            if (reportMetadata.OneToManyRelationships == null || reportMetadata.ManyToManyRelationships == null)
            {
                Set(rows, "Incomplete", "Report relationship metadata was not returned; no related table guessed.");
                side.ScopeSchema.Add("Report relationship metadata unavailable/incomplete.");
                return;
            }
            var bindings = new List<Binding>();
            foreach (var relation in reportMetadata.OneToManyRelationships.Where(r => r != null &&
                r.ReferencedEntity == "report" && r.ReferencedAttribute == Type31EvidenceCollector.PrimaryId &&
                LogicalName(r.ReferencingEntity) && r.ReferencingEntity != "report" && LogicalName(r.ReferencingAttribute)))
                bindings.Add(new Binding { Entity = relation.ReferencingEntity, ForeignKey = relation.ReferencingAttribute, Relationship = relation.SchemaName });
            foreach (var relation in reportMetadata.ManyToManyRelationships.Where(r => r != null &&
                (r.Entity1LogicalName == "report" || r.Entity2LogicalName == "report") && LogicalName(r.IntersectEntityName)))
            {
                string foreignKey = relation.Entity1LogicalName == "report" ? relation.Entity1IntersectAttribute : relation.Entity2IntersectAttribute;
                if (LogicalName(foreignKey)) bindings.Add(new Binding { Entity = relation.IntersectEntityName, ForeignKey = foreignKey,
                    Relationship = relation.SchemaName, Intersect = true });
            }
            bindings = bindings.GroupBy(b => b.Entity + ":" + b.ForeignKey, StringComparer.Ordinal).Select(g => g.First())
                .OrderBy(b => b.Entity, StringComparer.Ordinal).ThenBy(b => b.ForeignKey, StringComparer.Ordinal).ToList();
            foreach (var binding in bindings)
                side.ScopeSchema.Add("Discovered relationship=" + binding.Relationship + "; relatedEntity=" + binding.Entity + "; reportForeignKey=" + binding.ForeignKey + "; intersect=" + binding.Intersect);
            if (bindings.Count == 0) { side.ScopeSchema.Add("No report-scoped relationship/intersect candidate exposed by metadata."); return; }

            var schemas = new Dictionary<string, EntityMetadata>(StringComparer.Ordinal);
            var names = bindings.Select(b => b.Entity).Distinct().ToArray();
            for (int offset = 0; offset < names.Length; offset += Type31EvidenceCollector.BatchSize)
            {
                var batch = names.Skip(offset).Take(Type31EvidenceCollector.BatchSize).ToArray();
                var query = new EntityQueryExpression
                {
                    Properties = new MetadataPropertiesExpression("LogicalName", "PrimaryIdAttribute", "Attributes"),
                    AttributeQuery = new AttributeQueryExpression { Properties = new MetadataPropertiesExpression { AllProperties = true } },
                    Criteria = new MetadataFilterExpression(LogicalOperator.Or)
                };
                foreach (var name in batch) query.Criteria.Conditions.Add(new MetadataConditionExpression("LogicalName", MetadataConditionOperator.Equals, name));
                token.ThrowIfCancellationRequested();
                side.ScopeMetadataRequests++;
                side.Requests.Add("Execute ScopeSchema RetrieveMetadataChanges(LogicalName IN [" + string.Join(",", batch) + "], LogicalName,PrimaryIdAttribute,Attributes; attribute properties=All)");
                try
                {
                    var response = service.Execute(new RetrieveMetadataChangesRequest { Query = query }) as RetrieveMetadataChangesResponse;
                    token.ThrowIfCancellationRequested();
                    if (response?.EntityMetadata == null || response.EntityMetadata.Any(m => m == null || !batch.Contains(m.LogicalName)) ||
                        batch.Any(name => response.EntityMetadata.Count(m => m.LogicalName == name) != 1))
                    { Set(rows, "Incomplete", "Related entity schema missing/conflicting; no scope query issued."); return; }
                    foreach (var metadata in response.EntityMetadata) schemas.Add(metadata.LogicalName, metadata);
                }
                catch (OperationCanceledException) { throw; }
                catch (Exception)
                { token.ThrowIfCancellationRequested(); Set(rows, "Faulted", "Related scope schema retrieval failed; server details withheld."); return; }
            }
            var qualified = new List<Binding>();
            foreach (var binding in bindings)
            {
                var metadata = schemas[binding.Entity]; var attributes = metadata.Attributes;
                if (attributes == null || attributes.Any(a => a == null || !LogicalName(a.LogicalName)) ||
                    attributes.GroupBy(a => a.LogicalName).Any(g => g.Count() != 1))
                { Set(rows, "Incomplete", "Related entity attributes incomplete/conflicting; no guessed columns."); return; }
                var primary = attributes.SingleOrDefault(a => a.LogicalName == metadata.PrimaryIdAttribute);
                var foreign = attributes.SingleOrDefault(a => a.LogicalName == binding.ForeignKey);
                var lookup = foreign as LookupAttributeMetadata;
                bool validForeign = foreign?.IsValidForRead == true &&
                    (lookup != null && lookup.Targets?.Length == 1 && lookup.Targets[0] == "report" ||
                     binding.Intersect && foreign.AttributeType == AttributeTypeCode.Uniqueidentifier);
                var scopes = attributes.Where(a => a.IsValidForRead == true &&
                    (a.AttributeType == AttributeTypeCode.EntityName &&
                     (a.LogicalName == "objecttypecode" || a.LogicalName == "entitylogicalname" || a.LogicalName == "entityname") ||
                     a.LogicalName == "objecttypecode" && a.AttributeType == AttributeTypeCode.Integer)).ToArray();
                side.ScopeSchema.Add("Validated entity=" + binding.Entity + "; primary=" + metadata.PrimaryIdAttribute +
                    "; primaryReadable=" + (primary?.IsValidForRead == true && primary.AttributeType == AttributeTypeCode.Uniqueidentifier) +
                    "; reportForeignKeyReadable=" + validForeign + "; entityScopeColumns=[" + string.Join(",", scopes.Select(a => a.LogicalName)) + "]");
                if (primary?.IsValidForRead != true || primary.AttributeType != AttributeTypeCode.Uniqueidentifier || !validForeign) continue;
                if (scopes.Length > 1) { Set(rows, "Ambiguous", "Multiple metadata entity-scope columns; no role guessed."); return; }
                if (scopes.Length == 0) continue;
                binding.PrimaryId = metadata.PrimaryIdAttribute; binding.ScopeField = scopes[0].LogicalName;
                qualified.Add(binding);
            }
            if (qualified.Count != 1)
            {
                if (qualified.Count > 1) Set(rows, "Ambiguous", "Multiple metadata-proven scope paths; no automatic path selection.");
                else side.ScopeSchema.Add("No readable association with a metadata entity-name/object-type scope column. Lookup/local IDs are not scope identities.");
                return;
            }
            var selected = qualified[0];
            side.ScopeSchema.Add("Selected evidence path=" + selected.Entity + "." + selected.ForeignKey + " -> " + selected.ScopeField);
            for (int offset = 0; offset < rows.Length; offset += Type31EvidenceCollector.BatchSize)
            {
                token.ThrowIfCancellationRequested();
                var batch = rows.Skip(offset).Take(Type31EvidenceCollector.BatchSize).ToArray();
                var ids = batch.Select(r => r.ObjectId).ToArray();
                var associations = new Dictionary<Guid, Entity>(); var duplicateReports = new HashSet<Guid>();
                string cookie = null; int pageNumber = 1; bool complete = false;
                try
                {
                    while (true)
                    {
                        token.ThrowIfCancellationRequested();
                        var query = new QueryExpression(selected.Entity)
                        {
                            ColumnSet = new ColumnSet(selected.PrimaryId, selected.ForeignKey, selected.ScopeField),
                            PageInfo = new PagingInfo { Count = 200, PageNumber = pageNumber, PagingCookie = cookie }
                        };
                        query.Criteria.AddCondition(selected.ForeignKey, ConditionOperator.In, ids.Select(id => (object)id).ToArray());
                        query.AddOrder(selected.PrimaryId, OrderType.Ascending);
                        side.ScopeQueries++;
                        side.Requests.Add("RetrieveMultiple Scope " + selected.Entity + "; columns=[" + string.Join(",", query.ColumnSet.Columns) +
                            "]; " + selected.ForeignKey + " IN Guid[" + ids.Length + "]; page=" + pageNumber);
                        progress?.Invoke(side.Snapshot.Environment.DisplayName + ": reading related Report scope page " + pageNumber);
                        var response = service.RetrieveMultiple(query); token.ThrowIfCancellationRequested();
                        if (response == null) { Set(batch, "Incomplete", "Null related scope result; terminal retrieval not proven."); break; }
                        side.ScopeSchema.Add("Scope batchRequestedIds=" + ids.Length + "; page=" + pageNumber + "; returnedRows=" + response.Entities.Count +
                            "; MoreRecords=" + response.MoreRecords + "; PagingCookieSupplied=" + !string.IsNullOrEmpty(response.PagingCookie));
                        int previousCount = associations.Count;
                        if (response.Entities.Any(r => r == null || r.LogicalName != selected.Entity || r.Id == Guid.Empty ||
                            Type31EvidenceCollector.Id(r, selected.PrimaryId) != r.Id || !ids.Contains(ReportId(r, selected.ForeignKey))))
                        { Set(batch, "Incomplete", "Conflicting/foreign association primary or report ID; scope not proven."); break; }
                        foreach (var group in response.Entities.GroupBy(r => r.Id))
                        {
                            if (group.Count() > 1) foreach (var r in group) duplicateReports.Add(ReportId(r, selected.ForeignKey));
                            foreach (var r in group)
                            {
                                if (associations.TryGetValue(r.Id, out var previous))
                                {
                                    if (!Type31EvidenceCollector.SameReturnedRow(previous, r))
                                    { duplicateReports.Add(ReportId(previous, selected.ForeignKey)); duplicateReports.Add(ReportId(r, selected.ForeignKey)); }
                                }
                                else associations.Add(r.Id, r);
                            }
                        }
                        if (!response.MoreRecords) { complete = true; break; }
                        if (associations.Count == previousCount || !string.IsNullOrEmpty(response.PagingCookie) && response.PagingCookie == cookie)
                        { Set(batch, "Incomplete", "Related scope paging stalled; terminal retrieval not proven."); break; }
                        cookie = response.PagingCookie; pageNumber++;
                    }
                    if (!complete) continue;
                    var values = associations.Values.ToArray();
                    var codes = values.Select(r => ScopeCode(r.GetAttributeValue<object>(selected.ScopeField))).Where(c => c.HasValue)
                        .Select(c => c.Value).Distinct().OrderBy(c => c).ToArray();
                    var logicalNames = ResolveCodes(service, side, codes, token);
                    foreach (var row in batch)
                    {
                        var matches = values.Where(r => ReportId(r, selected.ForeignKey) == row.ObjectId).ToArray();
                        if (duplicateReports.Contains(row.ObjectId) || matches.Length > 1)
                        { row.ScopeStatus = "Ambiguous"; row.ScopeReason = "Multiple/duplicate scope association rows; no single entity scope selected."; continue; }
                        if (matches.Length == 0) { row.ScopeStatus = "Unavailable"; row.ScopeReason = "No report-scoped association row returned."; continue; }
                        var raw = matches[0].GetAttributeValue<object>(selected.ScopeField);
                        var code = ScopeCode(raw);
                        string scope = code.HasValue ? (logicalNames.TryGetValue(code.Value, out var value) ? value : null) : raw as string;
                        if (!LogicalName(scope) || scope.Equals("none", StringComparison.OrdinalIgnoreCase))
                        { row.ScopeStatus = "Incomplete"; row.ScopeReason = "Blank/unresolved metadata entity scope; local codes/IDs never substitute for logical names."; continue; }
                        row.RelatedScope = scope.Trim(); row.ScopeStatus = "Available";
                        row.ScopeReason = "One uniquely correlated association; metadata entity logical-name scope via " + selected.Entity + "." + selected.ScopeField;
                    }
                }
                catch (OperationCanceledException) { throw; }
                catch (Exception) { token.ThrowIfCancellationRequested(); Set(batch, "Faulted", "Related scope retrieval failed; server details withheld."); }
            }
        }

        private static Dictionary<int, string> ResolveCodes(IOrganizationService service, Type31SideEvidence side, int[] codes, CancellationToken token)
        {
            var names = new Dictionary<int, string>();
            for (int offset = 0; offset < codes.Length; offset += 200)
            {
                var batch = codes.Skip(offset).Take(200).ToArray();
                var query = new EntityQueryExpression { Properties = new MetadataPropertiesExpression("LogicalName", "ObjectTypeCode"), Criteria = new MetadataFilterExpression(LogicalOperator.Or) };
                foreach (var code in batch) query.Criteria.Conditions.Add(new MetadataConditionExpression("ObjectTypeCode", MetadataConditionOperator.Equals, code));
                token.ThrowIfCancellationRequested(); side.ScopeMetadataRequests++;
                side.Requests.Add("Execute ScopeCodes RetrieveMetadataChanges(ObjectTypeCode IN [" + string.Join(",", batch) + "], LogicalName,ObjectTypeCode)");
                var response = service.Execute(new RetrieveMetadataChangesRequest { Query = query }) as RetrieveMetadataChangesResponse;
                token.ThrowIfCancellationRequested();
                if (response?.EntityMetadata == null || response.EntityMetadata.Any(m => m == null || !m.ObjectTypeCode.HasValue || !batch.Contains(m.ObjectTypeCode.Value))) continue;
                foreach (var code in batch)
                {
                    var matches = response.EntityMetadata.Where(m => m.ObjectTypeCode == code).ToArray();
                    if (matches.Length == 1 && LogicalName(matches[0].LogicalName)) names[code] = matches[0].LogicalName;
                }
            }
            return names;
        }

        private static int? ScopeCode(object raw)
        {
            if (raw is int) return (int)raw;
            if (raw is string && int.TryParse(((string)raw).Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out var code)) return code;
            return null;
        }
        private static Guid ReportId(Entity row, string attribute)
        {
            var raw = row.GetAttributeValue<object>(attribute);
            if (raw is EntityReference && ((EntityReference)raw).LogicalName == "report") return ((EntityReference)raw).Id;
            return raw is Guid ? (Guid)raw : Guid.Empty;
        }
        private static bool LogicalName(string value) => !string.IsNullOrWhiteSpace(value) &&
            (char.IsLetter(value.Trim()[0]) || value.Trim()[0] == '_') && value.Trim().All(c => char.IsLetterOrDigit(c) || c == '_');
        private static void Set(IEnumerable<Type31ReportEvidence> rows, string status, string reason)
        { foreach (var row in rows) { row.RelatedScope = null; row.ScopeStatus = status; row.ScopeReason = reason; } }
    }
}
#endif
