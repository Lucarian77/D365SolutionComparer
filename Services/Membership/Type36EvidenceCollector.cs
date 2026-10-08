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
using Microsoft.Xrm.Sdk.Query;

namespace D365SolutionComparer.Services.Membership
{
    /// <summary>Explicit Debug evidence capture only. Never supplies a production identity or absence proof.</summary>
    internal sealed class Type36EvidenceCollector
    {
        internal const int BatchSize = 200;
        internal const string PrimaryId = "templateid";
        internal static readonly string[] AuditFields = { PrimaryId, "templateidunique", "title", "templatetypecode",
            "objecttypecode", "languagecode", "ispersonal", "statecode", "statuscode", "ismanaged", "componentstate", "ownerid",
            "generationtypecode", "uniquename", "schemaname", "organizationid", "solutionid", "owninguser", "owningteam" };
        internal static readonly string[] ContentFields = { "subject", "body", "presentationxml", "subjectpresentationxml", "description" };

        internal Type36EvidenceReport Capture(IOrganizationService sourceService, MembershipSnapshot source,
            string sourceVersion, IOrganizationService targetService, MembershipSnapshot target,
            string targetVersion, CancellationToken token, Action<string> progress = null)
        {
            if (source?.State != MembershipSnapshotState.Complete || target?.State != MembershipSnapshotState.Complete ||
                !StringComparer.OrdinalIgnoreCase.Equals(source.SolutionUniqueName, target.SolutionUniqueName))
                throw new ArgumentException("Completed snapshots of the same solution are required.");
            token.ThrowIfCancellationRequested();
            var report = new Type36EvidenceReport
            {
                Source = Read(sourceService, source, sourceVersion, token, progress),
                Target = Read(targetService, target, targetVersion, token, progress)
            };
            // Same evidence contract on both populated sides. Personal scope strengthens A only
            // when readable on every populated side; B never repairs an A collision.
            var populated = new[] { report.Source, report.Target }.Where(s => s.Rows.Count > 0).ToArray();
            report.IncludePersonalScope = populated.Length > 0 && populated.All(s => s.Columns.Contains("ispersonal"));
            foreach (var side in new[] { report.Source, report.Target })
            {
                foreach (var row in side.Rows.Values.Where(r => r.Status == "Unique")) row.ConstructCandidates(report.IncludePersonalScope);
                MarkDuplicates(side.Rows.Values, false); MarkDuplicates(side.Rows.Values, true);
                foreach (var group in side.Rows.Values.Where(r => r.CandidateP != null)
                    .GroupBy(r => r.CandidateP, StringComparer.OrdinalIgnoreCase).Where(g => g.Count() > 1))
                    foreach (var row in group) row.DuplicateP = true;
            }
            report.Analyze(token);
            token.ThrowIfCancellationRequested();
            return report;
        }

        private static void MarkDuplicates(IEnumerable<Type36TemplateEvidence> rows, bool weaker)
        {
            foreach (var group in rows.Where(r => r.Status == "Unique" && (weaker ? r.CandidateB : r.CandidateA) != null)
                .GroupBy(r => weaker ? r.CandidateB : r.CandidateA, StringComparer.OrdinalIgnoreCase).Where(g => g.Count() > 1))
                foreach (var row in group) { if (weaker) row.DuplicateB = true; else row.DuplicateA = true; }
        }

        private static Type36SideEvidence Read(IOrganizationService service, MembershipSnapshot snapshot,
            string version, CancellationToken token, Action<string> progress)
        {
            if (service == null) throw new ArgumentNullException(nameof(service));
            var side = new Type36SideEvidence { Snapshot = snapshot, Version = version, CapturedUtc = DateTimeOffset.UtcNow };
            side.Raw.AddRange(snapshot.Components.Where(c => c.Record.ComponentType == 36));
            var ids = side.Raw.Where(c => c.Record.ObjectId.HasValue && c.Record.ObjectId.Value != Guid.Empty)
                .Select(c => c.Record.ObjectId.Value).Distinct().OrderBy(id => id).ToArray();
            foreach (var id in ids) side.Rows.Add(id, new Type36TemplateEvidence { ObjectId = id, Status = "Incomplete", Reason = "Schema not verified" });
            if (ids.Length == 0) return side;
            token.ThrowIfCancellationRequested();
            side.Requests.Add("Execute RetrieveEntity(template, Attributes, RetrieveAsIfPublished=False)");
            try
            {
                progress?.Invoke(snapshot.Environment.DisplayName + ": validating readable template fields");
                var response = service.Execute(new RetrieveEntityRequest
                { LogicalName = "template", EntityFilters = EntityFilters.Attributes, RetrieveAsIfPublished = false }) as RetrieveEntityResponse;
                token.ThrowIfCancellationRequested();
                var metadata = response?.EntityMetadata;
                if (metadata?.Attributes == null || metadata.LogicalName != "template" || metadata.PrimaryIdAttribute != PrimaryId ||
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
                        new[] { "body", "subject", "content", "xml", "html", "text" }.Any(part => name.Contains(part)));
                    bool desired = AuditFields.Contains(name) || uniqueId || content;
                    side.MetadataFields.Add(name + ";type=" + attribute.AttributeType + ";metadataReadable=" + attribute.IsValidForRead +
                        ";selected=" + desired);
                    if (!desired) continue;
                    bool validShape = name == PrimaryId || uniqueId ? attribute.AttributeType == AttributeTypeCode.Uniqueidentifier :
                        content || name == "title" || name == "uniquename" || name == "schemaname" ? text :
                        new[] { "ownerid", "organizationid", "solutionid", "owninguser", "owningteam" }.Contains(name) ? attribute.AttributeType == AttributeTypeCode.Lookup || attribute.AttributeType == AttributeTypeCode.Owner :
                        name == "ismanaged" || name == "ispersonal" ? attribute.AttributeType == AttributeTypeCode.Boolean :
                        attribute.AttributeType == AttributeTypeCode.Integer || attribute.AttributeType == AttributeTypeCode.Picklist ||
                        attribute.AttributeType == AttributeTypeCode.State || attribute.AttributeType == AttributeTypeCode.Status ||
                        (name == "templatetypecode" || name == "objecttypecode") && (text || attribute.AttributeType == AttributeTypeCode.EntityName);
                    bool allowed = attribute.IsValidForRead == true && validShape && !name.Contains("attachment");
                    side.Schema[name] = allowed ? "Readable" : "UnavailableOrUnverified";
                    if (!allowed) continue;
                    side.Columns.Add(name);
                    if (content) side.HashFields.Add(name);
                    if (uniqueId) side.UniqueFields.Add(name);
                    if (name == "title") side.FixedTitleVerified = attribute.AttributeType == AttributeTypeCode.String;
                }
                if (!side.Columns.Contains(PrimaryId)) throw new InvalidOperationException("Primary ID is not readable");
            }
            catch (OperationCanceledException) { throw; }
            catch (Exception ex)
            {
                token.ThrowIfCancellationRequested();
                foreach (var row in side.Rows.Values) { row.Status = ex is InvalidOperationException ? "Incomplete" : "Faulted"; row.Reason = "Schema unavailable/faulted/incomplete; server details withheld. No guessed columns requested."; }
                return side;
            }

            ReadBatches(service, side, ids, side.Columns.ToArray(), token, progress, true);
            ResolveProposedScopes(service, side, token);
            return side;
        }

        private static void ReadBatches(IOrganizationService service, Type36SideEvidence side, Guid[] ids, string[] columns,
            CancellationToken token, Action<string> progress, bool allowContentRetry)
        {
            for (int offset = 0; offset < ids.Length; offset += BatchSize)
            {
                token.ThrowIfCancellationRequested();
                var batch = ids.Skip(offset).Take(BatchSize).ToArray();
                var audit = new Type36BatchEvidence { BatchNumber = offset / BatchSize + 1, RequestedIdCount = batch.Length };
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
                        var query = new QueryExpression("template")
                        {
                            ColumnSet = new ColumnSet(columns),
                            PageInfo = new PagingInfo { Count = BatchSize, PageNumber = pageNumber, PagingCookie = cookie }
                        };
                        query.Criteria.AddCondition(new ConditionExpression(PrimaryId, ConditionOperator.In, batch.Select(id => (object)id).ToArray()));
                        query.AddOrder(PrimaryId, OrderType.Ascending);
                        side.Requests.Add("RetrieveMultiple template; columns=[" + string.Join(",", columns) + "]; templateid IN Guid[" + batch.Length + "]; page=" + pageNumber);
                        var page = new Type36PageEvidence { PageNumber = pageNumber };
                        audit.Pages.Add(page);
                        progress?.Invoke(side.Snapshot.Environment.DisplayName + ": reading template batch " + audit.BatchNumber + ", page " + pageNumber);
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
                            if (observedIds.Length == 1) side.Rows[id].TemplateId = observedIds[0];
                        }
                        if (response.Entities.Any(r => r == null || r.LogicalName != "template" || !Id(r, PrimaryId).HasValue ||
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
                                    row.Status == "Duplicate" ? "Duplicate/conflicting returned primary-key rows; no candidate evaluated" : "Exact ObjectId/templateid/Entity.Id correlation; terminal page received";
                                if (row.Status == "Unique") row.Capture(uniqueRows[id], side, columns);
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
                    Fail(side, batch, "Faulted", "Template retrieval failed; server details withheld");
                    // A failed continuation remains incomplete/faulted. Only a full-query initial
                    // fault gets one selected-ID identity/audit retry, with content unavailable.
                    if (allowContentRetry && pageNumber == 1 && columns.Any(c => side.HashFields.Contains(c) || c == "generationtypecode"))
                    {
                        side.RuntimeEvidence.Add("Full query faulted; retrying identity/audit columns only. Optional content unavailable; generationtypecode unavailable. Server details withheld.");
                        ReadBatches(service, side, batch, columns.Where(c => !side.HashFields.Contains(c) && c != "generationtypecode").ToArray(), token, progress, false);
                    }
                }
            }
        }

        private static void ResolveProposedScopes(IOrganizationService service, Type36SideEvidence side, CancellationToken token)
        {
            var pending = new Dictionary<string, Type36ScopeEvidence>(StringComparer.OrdinalIgnoreCase);
            foreach (var row in side.Rows.Values.Where(r => r.Status == "Unique"))
            {
                string input = row.Get("templatetypecode");
                if (!string.IsNullOrWhiteSpace(input))
                {
                    bool numeric = int.TryParse(input, NumberStyles.Integer, CultureInfo.InvariantCulture, out var number);
                    bool logical = input.Length <= 128 && (char.IsLetter(input[0]) || input[0] == '_') && input.All(c => char.IsLetterOrDigit(c) || c == '_') &&
                        !StringComparer.OrdinalIgnoreCase.Equals(input, "none");
                    if (numeric || logical)
                    {
                        string key = numeric ? "otc:" + number.ToString(CultureInfo.InvariantCulture) : "logical:" + input;
                        if (!pending.TryGetValue(key, out var scope))
                        {
                            scope = new Type36ScopeEvidence { Number = numeric ? (int?)number : null, Name = numeric ? null : input };
                            pending.Add(key, scope);
                            if (!numeric)
                            {
                                var matches = side.Snapshot.Components.Where(c => c.Record.ComponentType == 1 &&
                                    StringComparer.OrdinalIgnoreCase.Equals(c.ComparisonKey, input)).ToArray();
                                if (matches.Length > 0)
                                {
                                    bool unique = matches.All(c => c.Status == IdentityResolutionStatus.Resolved && c.InventoryAbsencePolicy == InventoryAbsencePolicy.CompleteInventory) &&
                                        matches.Select(c => c.Record.ObjectId).Distinct().Count() == 1 && matches[0].Record.ObjectId.HasValue && matches[0].Record.ObjectId != Guid.Empty;
                                    scope.Status = unique ? "Verified" : "Ambiguous"; scope.Key = unique ? input.ToLowerInvariant() : null;
                                    scope.Reason = "Completed snapshot Table correlation";
                                }
                            }
                        }
                        row.Scope = scope;
                    }
                }
            }
            var unresolved = pending.Values.Where(s => s.Status == "Incomplete").ToArray();
            for (int offset = 0; offset < unresolved.Length; offset += BatchSize)
            {
                token.ThrowIfCancellationRequested(); var batch = unresolved.Skip(offset).Take(BatchSize).ToArray();
                var query = new EntityQueryExpression { Properties = new MetadataPropertiesExpression("MetadataId", "LogicalName", "ObjectTypeCode"),
                    Criteria = new MetadataFilterExpression(LogicalOperator.Or) };
                foreach (var scope in batch) query.Criteria.Conditions.Add(new MetadataConditionExpression(scope.Number.HasValue ? "ObjectTypeCode" : "LogicalName",
                    MetadataConditionOperator.Equals, scope.Number.HasValue ? (object)scope.Number.Value : scope.Name));
                side.Requests.Add("Execute RetrieveMetadataChanges; selected template scopes=" + batch.Length + "; properties=[MetadataId,LogicalName,ObjectTypeCode]");
                try
                {
                    var response = service.Execute(new RetrieveMetadataChangesRequest { Query = query }) as RetrieveMetadataChangesResponse;
                    token.ThrowIfCancellationRequested();
                    if (response?.EntityMetadata == null || response.EntityMetadata.Any(m => m == null || !batch.Any(s => s.Matches(m))))
                    { foreach (var scope in batch) scope.Reason = "Incomplete/foreign scoped metadata response"; continue; }
                    foreach (var scope in batch)
                    {
                        var matches = response.EntityMetadata.Where(scope.Matches).ToArray();
                        if (matches.Length > 1) { scope.Status = "Ambiguous"; scope.Reason = "Multiple table metadata candidates"; }
                        else if (matches.Length == 1 && matches[0].MetadataId.HasValue && matches[0].MetadataId != Guid.Empty &&
                            !string.IsNullOrWhiteSpace(matches[0].LogicalName) && matches[0].LogicalName.Length <= 128 &&
                            (char.IsLetter(matches[0].LogicalName[0]) || matches[0].LogicalName[0] == '_') &&
                            matches[0].LogicalName.All(c => char.IsLetterOrDigit(c) || c == '_') && !StringComparer.OrdinalIgnoreCase.Equals(matches[0].LogicalName, "none"))
                        {
                            var local = side.Snapshot.Components.Where(c => c.Record.ComponentType == 1 && c.Record.ObjectId == matches[0].MetadataId).ToArray();
                            if (local.Any(c => c.Status != IdentityResolutionStatus.Resolved || !StringComparer.OrdinalIgnoreCase.Equals(c.ComparisonKey, matches[0].LogicalName)))
                            { scope.Status = "Ambiguous"; scope.Reason = "Conflicting completed Table snapshot evidence"; continue; }
                            scope.Status = "Verified"; scope.Key = matches[0].LogicalName.ToLowerInvariant(); scope.Reason = "Unique selected-scope table metadata; GUID local correlation only";
                        }
                        else scope.Reason = "Missing/incomplete selected table metadata";
                    }
                }
                catch (OperationCanceledException) { throw; }
                catch (Exception) { token.ThrowIfCancellationRequested(); foreach (var scope in batch) { scope.Status = "Faulted"; scope.Reason = "Scope metadata fault; server details withheld"; } }
            }
            foreach (var row in side.Rows.Values.Where(r => r.Status == "Unique"))
            {
                var blockers = new List<string>();
                string title = row.Get("title"), language = row.Get("languagecode");
                if (!side.FixedTitleVerified || !row.RuntimeColumns.Contains("title") || string.IsNullOrWhiteSpace(title) || Guid.TryParse(title, out var ignored))
                    blockers.Add("Fixed title unavailable/blank/invalid; no identifier fallback");
                if (row.Scope?.Status != "Verified") blockers.Add("templatetypecode table mapping " + (row.Scope?.Status ?? "Incomplete") + "; no numeric/global fallback");
                int languageValue = 0;
                if (!row.RuntimeColumns.Contains("languagecode") || !int.TryParse(language, NumberStyles.Integer, CultureInfo.InvariantCulture, out languageValue) || languageValue <= 0)
                    blockers.Add("Exact languagecode unavailable/blank/malformed");
                row.ProposedBlockingReason = blockers.Count == 0 ? "None; ScopedPrimaryNameHypothesis, generation/personal-scope semantics and lifecycle portability need live review" : string.Join("; ", blockers);
                if (blockers.Count == 0) row.CandidateP = "template-candidate-p:v1:" + Frame(row.Scope.Key, languageValue.ToString(CultureInfo.InvariantCulture), title.Trim());
            }
        }

        private static bool SameReturnedRow(Entity left, Entity right) => left.Attributes.Count == right.Attributes.Count &&
            left.Attributes.All(a => right.Attributes.TryGetValue(a.Key, out var value) && SameValue(a.Value, value));
        private static bool SameValue(object left, object right)
        {
            if (left is OptionSetValue && right is OptionSetValue) return ((OptionSetValue)left).Value == ((OptionSetValue)right).Value;
            if (left is EntityReference && right is EntityReference)
                return ((EntityReference)left).Id == ((EntityReference)right).Id &&
                    StringComparer.OrdinalIgnoreCase.Equals(((EntityReference)left).LogicalName, ((EntityReference)right).LogicalName);
            return Equals(left, right);
        }

        private static void Fail(Type36SideEvidence side, IEnumerable<Guid> ids, string status, string reason)
        { foreach (var id in ids) { side.Rows[id].Status = status; side.Rows[id].Reason = reason; } }
        internal static Guid? Id(Entity row, string field) => row.Attributes.TryGetValue(field, out var raw) && raw is Guid && (Guid)raw != Guid.Empty ? (Guid?)raw : null;
        internal static string Frame(params string[] values) => string.Concat(values.Select(value => value.Length.ToString(CultureInfo.InvariantCulture) + ":" + value + ":"));
        internal static string Safe(object value) => (Convert.ToString(value, CultureInfo.InvariantCulture) ?? "Unknown")
            .Replace("\r", "\\r").Replace("\n", "\\n").Replace("\t", "\\t");
    }

    internal sealed class Type36SideEvidence
    {
        internal MembershipSnapshot Snapshot;
        internal string Version;
        internal DateTimeOffset CapturedUtc;
        internal int ReturnedRows;
        internal string PrimaryName;
        internal bool FixedTitleVerified;
        internal readonly List<string> MetadataFields = new List<string>(), RuntimeEvidence = new List<string>();
        internal readonly List<ComponentIdentity> Raw = new List<ComponentIdentity>();
        internal readonly SortedDictionary<Guid, Type36TemplateEvidence> Rows = new SortedDictionary<Guid, Type36TemplateEvidence>();
        internal readonly List<string> Columns = new List<string>();
        internal readonly HashSet<string> HashFields = new HashSet<string>(StringComparer.Ordinal);
        internal readonly HashSet<string> UniqueFields = new HashSet<string>(StringComparer.Ordinal);
        internal readonly SortedDictionary<string, string> Schema = new SortedDictionary<string, string>(StringComparer.Ordinal);
        internal readonly List<string> Requests = new List<string>();
        internal readonly List<Type36BatchEvidence> Batches = new List<Type36BatchEvidence>();
        internal bool Complete => Raw.All(r => r.Record.ObjectId.HasValue && r.Record.ObjectId.Value != Guid.Empty) && Rows.Values.All(r => r.Status == "Unique" && r.CandidateA != null);
    }

    internal sealed class Type36ScopeEvidence
    {
        internal string Status = "Incomplete", Name, Key, Reason = "Table scope unavailable";
        internal int? Number;
        internal bool Matches(EntityMetadata metadata) => Number.HasValue ? metadata.ObjectTypeCode == Number : StringComparer.OrdinalIgnoreCase.Equals(metadata.LogicalName, Name);
    }

    internal sealed class Type36BatchEvidence
    {
        internal int BatchNumber, RequestedIdCount, ReturnedRows, DistinctReturnedIds;
        internal bool Complete;
        internal readonly List<Type36PageEvidence> Pages = new List<Type36PageEvidence>();
    }

    internal sealed class Type36PageEvidence
    {
        internal int PageNumber;
        internal int? RowsReturned;
        internal bool? MoreRecords, PagingCookieSupplied;
    }

    internal sealed class Type36ContentFingerprint
    {
        internal string Presence;
        internal int? Length;
        internal string Sha256;
        internal static Type36ContentFingerprint Create(object raw)
        {
            if (raw != null && !(raw is string)) return new Type36ContentFingerprint { Presence = "Malformed" };
            string text = raw as string ?? string.Empty;
            using (var sha = SHA256.Create())
                return new Type36ContentFingerprint { Presence = text.Length == 0 ? "Blank" : "Present", Length = text.Length,
                    Sha256 = BitConverter.ToString(sha.ComputeHash(Encoding.UTF8.GetBytes(text))).Replace("-", string.Empty) };
        }
        internal bool Known => Sha256 != null;
        internal string Evidence => "presence=" + Presence + ";length=" + (Length?.ToString(CultureInfo.InvariantCulture) ?? "Unknown") + ";sha256=" + (Sha256 ?? "Unknown");
    }

    internal sealed class Type36TemplateEvidence
    {
        internal Guid ObjectId;
        internal Guid? TemplateId;
        internal int BackingRowCount;
        internal string Status, Reason, CandidateA, CandidateB;
        internal bool DuplicateA, DuplicateB;
        internal string CandidateP, ProposedBlockingReason;
        internal bool DuplicateP;
        internal Type36ScopeEvidence Scope;
        internal readonly List<string> RuntimeColumns = new List<string>();
        // No Entity/raw subject/body is retained in the report model.
        internal readonly SortedDictionary<string, string> Fields = new SortedDictionary<string, string>(StringComparer.Ordinal);
        internal readonly SortedDictionary<string, Type36ContentFingerprint> Content = new SortedDictionary<string, Type36ContentFingerprint>(StringComparer.Ordinal);
        internal readonly SortedDictionary<string, Guid?> UniqueIds = new SortedDictionary<string, Guid?>(StringComparer.Ordinal);
        internal bool? Managed;
        internal string Get(string field) => Fields.TryGetValue(field, out var value) ? value : null;

        internal void Capture(Entity row, Type36SideEvidence side, string[] columns)
        {
            TemplateId = Type36EvidenceCollector.Id(row, Type36EvidenceCollector.PrimaryId);
            RuntimeColumns.AddRange(columns);
            foreach (var field in columns)
            {
                row.Attributes.TryGetValue(field, out var raw);
                if (side.HashFields.Contains(field)) { Content[field] = Type36ContentFingerprint.Create(raw); continue; }
                if (side.UniqueFields.Contains(field)) { UniqueIds[field] = Type36EvidenceCollector.Id(row, field); continue; }
                string value = null;
                if (field == "title" || field == "uniquename" || field == "schemaname") value = raw is string ? ((string)raw).Trim() : raw == null ? "" : null;
                else if (field == Type36EvidenceCollector.PrimaryId) value = TemplateId?.ToString("D");
                else if (raw == null) value = "";
                else if (raw is OptionSetValue) value = ((OptionSetValue)raw).Value.ToString(CultureInfo.InvariantCulture);
                else if (raw is int) value = ((int)raw).ToString(CultureInfo.InvariantCulture);
                else if (raw is bool) value = (bool)raw ? "True" : "False";
                else if ((field == "templatetypecode" || field == "objecttypecode") && raw is string)
                {
                    var scope = ((string)raw).Trim();
                    if (scope.Length <= 128 && scope.All(c => char.IsLetterOrDigit(c) || c == '_')) value = scope;
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

        internal void ConstructCandidates(bool personal)
        {
            string scope = Get("templatetypecode"), title = Get("title"), language = Get("languagecode");
            if (string.IsNullOrWhiteSpace(scope)) scope = Get("objecttypecode");
            if (string.IsNullOrWhiteSpace(scope) || string.IsNullOrWhiteSpace(title)) return;
            CandidateB = "template-candidate-b:" + Type36EvidenceCollector.Frame(scope, title);
            if (string.IsNullOrWhiteSpace(language) || personal && string.IsNullOrWhiteSpace(Get("ispersonal"))) return;
            CandidateA = "template-candidate-a:" + Type36EvidenceCollector.Frame(scope, language, title,
                personal ? Get("ispersonal") : "PersonalScopeUnavailable");
        }
    }

    internal sealed class Type36LifecycleEntry
    {
        internal readonly List<Type36TemplateEvidence> Source = new List<Type36TemplateEvidence>();
        internal readonly List<Type36TemplateEvidence> Target = new List<Type36TemplateEvidence>();
        internal readonly HashSet<string> Outcomes = new HashSet<string>(StringComparer.Ordinal);
        internal string Basis;
        internal bool Pair => Source.Count == 1 && Target.Count == 1 && Source[0].Status == "Unique" && Target[0].Status == "Unique" &&
            Source[0].CandidateA != null && StringComparer.OrdinalIgnoreCase.Equals(Source[0].CandidateA, Target[0].CandidateA) && !Outcomes.Contains("Ambiguous");
        internal bool OnlyUniqueIdDifference;
    }

    internal sealed class Type36EvidenceReport
    {
        internal Type36SideEvidence Source, Target;
        internal bool IncludePersonalScope;
        internal readonly List<Type36LifecycleEntry> Lifecycle = new List<Type36LifecycleEntry>();
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
            foreach (var group in nodes.GroupBy(n => n.Row.TemplateId).Where(g => g.Any(n => n.Source) && g.Any(n => !n.Source)))
                Connect(group.ToArray(), edges);
            var visited = new HashSet<Node>();
            foreach (var node in nodes)
            {
                token.ThrowIfCancellationRequested();
                if (!visited.Add(node)) continue;
                var pending = new Queue<Node>(); pending.Enqueue(node);
                var entry = new Type36LifecycleEntry();
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
                    var entry = new Type36LifecycleEntry { Basis = "Backing correlation " + row.Status };
                    (side == Source ? entry.Source : entry.Target).Add(row);
                    entry.Outcomes.Add(row.Status == "Duplicate" ? "Ambiguous" : "Incomplete"); Lifecycle.Add(entry);
                }
                foreach (var raw in side.Raw.Where(c => !c.Record.ObjectId.HasValue || c.Record.ObjectId.Value == Guid.Empty))
                    Lifecycle.Add(new Type36LifecycleEntry { Basis = (side == Source ? "Source" : "Target") + " blank ObjectId", Outcomes = { "Incomplete" } });
            }
        }

        private sealed class Node { internal Type36TemplateEvidence Row; internal bool Source; }
        private static void Connect(Node[] group, Dictionary<Node, HashSet<Node>> edges)
        {
            // Star topology keeps duplicate-group analysis linear instead of N squared.
            for (int i = 1; i < group.Length; i++) { edges[group[0]].Add(group[i]); edges[group[i]].Add(group[0]); }
        }

        private void AnalyzePair(Type36LifecycleEntry entry)
        {
            var left = entry.Source[0]; var right = entry.Target[0];
            bool samePrimary = left.TemplateId == right.TemplateId;
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

        internal static string ContentComparison(Type36TemplateEvidence left, Type36TemplateEvidence right)
        {
            var fields = left.Content.Keys.Union(right.Content.Keys).ToArray();
            if (fields.Any(f => left.Content.ContainsKey(f) && right.Content.ContainsKey(f) && left.Content[f].Known && right.Content[f].Known && left.Content[f].Sha256 != right.Content[f].Sha256)) return "DifferentContent";
            return fields.Length > 0 && fields.All(f => left.Content.ContainsKey(f) && right.Content.ContainsKey(f) && left.Content[f].Known && right.Content[f].Known)
                ? "SameContent" : "Unknown";
        }

        private static void Line(StringBuilder text, params object[] cells) => text.AppendLine(string.Join("\t", cells.Select(Type36EvidenceCollector.Safe)));
        private static string Ids(IEnumerable<Type36TemplateEvidence> rows) => string.Join(",", rows.Select(r => r.ObjectId).OrderBy(id => id).Select(id => id.ToString("D")));
        internal string Build()
        {
            var text = new StringBuilder();
            text.AppendLine("TYPE 36 EMAIL TEMPLATE EVIDENCE - DEBUG ONLY");
            text.AppendLine("Evidence only. Type 36 remains Unsupported / Indeterminate. No production identity, membership matching, definition contract or absence inference.");
            text.AppendLine("Selected completed solution snapshots only. Raw GUID overlap is not proof of portability. Primary-ID pairs are audit observations.");
            text.AppendLine("Candidate A: template scope/type + language + title" + (IncludePersonalScope ? " + personal scope" : "; personal scope unavailable on a populated side") + ". Candidate B: scope/type + title. Trim + ordinal case-insensitive; B never repairs A ambiguity.");
            text.AppendLine("Subject/content: only presence, UTF-16 character length and exact UTF-8 SHA-256; no content normalization or raw content. Candidate A retains its original raw scope; Candidate P verifies portable table scope independently. Title localization/rename and lifecycle stability remain unproven.");
            text.AppendLine("RAW TYPE 36 MEMBERSHIP");
            Line(text, "Side", "Environment", "Solution", "Version", "SnapshotUtc", "SolutionComponentId", "ObjectId", "ProductionResolutionStatus", "ProductionDiagnostic", "ExistingPortableKey");
            foreach (var side in new[] { Source, Target })
            {
                string label = side == Source ? "Source" : "Target";
                foreach (var raw in side.Raw.OrderBy(r => r.Record.SolutionComponentId))
                    Line(text, label, side.Snapshot.Environment.DisplayName, side.Snapshot.SolutionUniqueName, side.Version,
                        side.Snapshot.CapturedAt.UtcDateTime.ToString("O", CultureInfo.InvariantCulture), raw.Record.SolutionComponentId,
                        raw.Record.ObjectId?.ToString("D") ?? "Blank", raw.Status, raw.Diagnostic, raw.ComparisonKey ?? "(none)");
                Line(text, label, "raw=" + side.Raw.Count, "distinctNonblankIds=" + side.Rows.Count,
                    "blankObjectIds=" + side.Raw.Count(r => !r.Record.ObjectId.HasValue || r.Record.ObjectId.Value == Guid.Empty));
            }
            text.AppendLine("BACKING TEMPLATE CORRELATION");
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
                foreach (var attribute in side.Schema) Line(text, label, "Schema", attribute.Key, attribute.Value);
                foreach (var field in Type36EvidenceCollector.AuditFields.Concat(Type36EvidenceCollector.ContentFields).Except(side.Schema.Keys).OrderBy(f => f, StringComparer.Ordinal))
                    Line(text, label, "Schema", field, "UnavailableInMetadataOrNotQueried");
                Line(text, "Side", "ObjectId", "BackingRowCount", "CorrelationStatus", "ObjectIdEqualsTemplateId", "TemplateId", "Reason", "CandidateA", "CandidateB", "DuplicateA", "DuplicateB");
                foreach (var row in side.Rows.Values)
                {
                    Line(text, label, row.ObjectId, row.BackingRowCount, row.Status,
                        row.TemplateId.HasValue ? (row.TemplateId == row.ObjectId ? "True" : "False") : "NotProven",
                        row.TemplateId?.ToString("D") ?? "Unknown", row.Reason, row.CandidateA ?? "Incomplete", row.CandidateB ?? "Incomplete", row.DuplicateA, row.DuplicateB);
                    foreach (var field in row.Fields) Line(text, label, row.ObjectId, field.Key, field.Value ?? "Malformed");
                    foreach (var field in row.UniqueIds) Line(text, label, row.ObjectId, field.Key, field.Value?.ToString("D") ?? "BlankOrMalformed");
                    foreach (var field in row.Content) Line(text, label, row.ObjectId, field.Key, field.Value.Evidence);
                }
                foreach (var raw in side.Raw.Where(r => !r.Record.ObjectId.HasValue || r.Record.ObjectId.Value == Guid.Empty))
                    Line(text, label, "Blank", 0, "Incomplete", "NotProven", "Unknown", "Blank ObjectId; no backing query",
                        "Incomplete", "Incomplete", false, false);
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
            Line(text, "RetainedSameTemplateId", Lifecycle.Count(e => e.Outcomes.Contains("SamePrimaryId")));
            Line(text, "DifferentTemplateId", Lifecycle.Count(e => e.Outcomes.Contains("DifferentPrimaryId")));
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
            BuildReadiness(text);
            text.AppendLine("REQUEST LEDGER");
            foreach (var side in new[] { Source, Target })
            {
                string label = side == Source ? "Source" : "Target";
                Line(text, label, "SchemaRequests=" + side.Requests.Count(r => r.StartsWith("Execute RetrieveEntity", StringComparison.Ordinal)),
                    "TemplateQueries=" + side.Requests.Count(r => r.StartsWith("RetrieveMultiple ", StringComparison.Ordinal)), "WhoAmI=0", "Writes=0", "MembershipQueries=0");
                Line(text, label, "ScopeMetadataRequests=" + side.Requests.Count(r => r.StartsWith("Execute RetrieveMetadataChanges", StringComparison.Ordinal)), "TotalReads=" + side.Requests.Count);
                for (int i = 0; i < side.Requests.Count; i++) Line(text, label, i + 1, side.Requests[i]);
            }
            return text.ToString();
        }

        private void BuildReadiness(StringBuilder text)
        {
            text.AppendLine("IDENTIFIER / SCOPE / LANGUAGE ANALYSIS");
            text.AppendLine("Candidate P is a ScopedPrimaryNameHypothesis only: fixed title + independently verified templatetypecode logical table + exact positive languagecode. No uniquename/schemaname/objecttypecode fallback. Language is retained as a proposed structural discriminator, not proven structural semantics; selected-sample collisions cannot establish global uniqueness.");
            text.AppendLine("Generation/template kind and personal scope remain Unresolved in identity classification pending live semantic review; they may require a stronger key. Internal identifiers, if exposed, are separate evidence and never silently replace title. Ownership/organization/solution/lifecycle IDs are audit-only. No global/tableless scope inferred.");
            foreach (var side in new[] { Source, Target })
            {
                string label = side == Source ? "Source" : "Target";
                Line(text, label, "PrimaryNameAttribute=" + (side.PrimaryName ?? "Unavailable"));
                foreach (var field in side.MetadataFields) Line(text, label, "Metadata", field);
                foreach (var evidence in side.RuntimeEvidence) Line(text, label, "Runtime", evidence);
                foreach (var row in side.Rows.Values)
                {
                    Line(text, label, row.ObjectId, "RuntimeColumns=[" + string.Join(",", row.RuntimeColumns) + "]",
                        "TemplateScope=" + row.Get("templatetypecode"), "TableStatus=" + (row.Scope?.Status ?? "Incomplete"),
                        "TableKey=" + row.Scope?.Key, "ScopeReason=" + row.Scope?.Reason);
                    Line(text, label, row.ObjectId, "CandidateA=" + row.CandidateA, "CandidateP=" + (row.CandidateP ?? "Incomplete"),
                        "CompleteP=" + (row.CandidateP != null && !row.DuplicateP), "DuplicateP=" + row.DuplicateP,
                        "BlockingReason=" + (row.DuplicateP ? "Candidate P collision; no automatic pairing" : row.ProposedBlockingReason ?? "Backing correlation " + row.Status));
                    foreach (var field in new[] { "title", "uniquename", "schemaname", "languagecode", "generationtypecode", "ispersonal" })
                        Line(text, label, row.ObjectId, field, "Schema=" + (side.Schema.TryGetValue(field, out var schema) ? schema : "NotExposed"),
                            "RuntimeRead=" + row.RuntimeColumns.Contains(field), "Value=" + (row.Get(field) ?? "UnavailableOrMalformed"));
                }
            }
            text.AppendLine("CRITICAL COLLISION MATRIX");
            foreach (var side in new[] { Source, Target })
            {
                string label = side == Source ? "Source" : "Target";
                var rows = side.Rows.Values.Where(r => r.Status == "Unique" && !string.IsNullOrWhiteSpace(r.Get("title"))).ToArray();
                Collision(text, label, "Title", rows.Select(r => Type36EvidenceCollector.Frame(r.Get("title"))));
                Collision(text, label, "Title+VerifiedTable", rows.Where(r => r.Scope?.Status == "Verified").Select(r => Type36EvidenceCollector.Frame(r.Scope.Key, r.Get("title"))));
                Collision(text, label, "Title+VerifiedTable+Language", rows.Where(r => r.CandidateP != null).Select(r => r.CandidateP));
                Line(text, label, "CandidateACollisionGroups=" + side.Rows.Values.Where(r => r.DuplicateA).Select(r => r.CandidateA).Distinct(StringComparer.OrdinalIgnoreCase).Count(),
                    "CandidatePCollisionGroups=" + side.Rows.Values.Where(r => r.DuplicateP).Select(r => r.CandidateP).Distinct(StringComparer.OrdinalIgnoreCase).Count(),
                    "IncompleteP=" + side.Rows.Values.Count(r => r.CandidateP == null));
            }
            text.AppendLine("CANDIDATE P LIFECYCLE COMPARISON");
            int pairs = 0, differing = 0, managed = 0, definitionDifferent = 0, ambiguous = 0, oneSided = 0;
            foreach (var key in Source.Rows.Values.Concat(Target.Rows.Values).Where(r => r.CandidateP != null).Select(r => r.CandidateP)
                .Distinct(StringComparer.OrdinalIgnoreCase).OrderBy(k => k, StringComparer.OrdinalIgnoreCase))
            {
                var left = Source.Rows.Values.Where(r => StringComparer.OrdinalIgnoreCase.Equals(r.CandidateP, key)).ToArray();
                var right = Target.Rows.Values.Where(r => StringComparer.OrdinalIgnoreCase.Equals(r.CandidateP, key)).ToArray();
                string outcome = left.Length > 1 || right.Length > 1 ? "Ambiguous" : left.Length == 1 && right.Length == 1 ? "ScopedPrimaryNamePairHypothesis" : "OneSidedEvidence";
                if (outcome == "Ambiguous") ambiguous++; else if (outcome == "OneSidedEvidence") oneSided++;
                Line(text, outcome, key, "SourceIds=" + Ids(left), "TargetIds=" + Ids(right));
                if (outcome != "ScopedPrimaryNamePairHypothesis") continue;
                pairs++; bool different = left[0].TemplateId != right[0].TemplateId; if (different) differing++;
                bool transition = left[0].Managed.HasValue && right[0].Managed.HasValue && left[0].Managed != right[0].Managed;
                if (transition) managed++;
                string content = ContentComparison(left[0], right[0]); if (content == "DifferentContent") definitionDifferent++;
                Line(text, "templateid=" + (different ? "DifferentObserved" : "EqualObserved"), "SourceManaged=" + left[0].Managed,
                    "TargetManaged=" + right[0].Managed, "ManagedTransition=" + transition, "Content=" + content);
                foreach (var field in left[0].Fields.Keys.Union(right[0].Fields.Keys).Union(left[0].Content.Keys).Union(right[0].Content.Keys).Union(left[0].UniqueIds.Keys).Union(right[0].UniqueIds.Keys).OrderBy(f => f, StringComparer.Ordinal))
                {
                    string l = Field(left[0], field), r = Field(right[0], field);
                    Line(text, field, "Source=" + (l ?? "Unavailable"), "Target=" + (r ?? "Unavailable"),
                        l == null || r == null ? "Incomplete" : StringComparer.OrdinalIgnoreCase.Equals(l, r) ? "EqualObserved" : "DifferentObserved");
                }
            }
            Line(text, "UniqueCandidatePPairs=" + pairs, "SamePrimaryId=" + (pairs - differing), "DifferentPrimaryId=" + differing,
                "ManagedTransition=" + managed, "DifferentDefinitionContent=" + definitionDifferent, "Ambiguous=" + ambiguous, "OneSidedEvidence=" + oneSided);
            text.AppendLine("INITIAL DEFINITION EVIDENCE / PORTABILITY GATE");
            text.AppendLine("Subject/body/presentation/layout/description hashes are definition evidence only; unknown content is never equal content. Generation type is a potential definition property only if live review excludes it from structural identity. Language remains in Candidate P until its role is established. No production definition contract is approved.");
            text.AppendLine("Same templateid does not prove portability. Even differing-ID Candidate P pairs remain hypotheses until independent deployment, title rename/localization, generation/personal scope and collision semantics are reviewed. Candidate B, GUID overlap and content hashes cannot repair incomplete/colliding identity. OneSidedEvidence is not absence proof. Type 36 remains Unsupported / Indeterminate.");
        }

        private static void Collision(StringBuilder text, string side, string dimension, IEnumerable<string> values)
        {
            var groups = values.GroupBy(v => v, StringComparer.OrdinalIgnoreCase).Where(g => g.Count() > 1).ToArray();
            Line(text, side, dimension, "CollisionGroups=" + groups.Length);
            foreach (var group in groups) Line(text, side, dimension, "Candidate=" + group.Key, "DistinctBackingRecords=" + group.Count());
        }

        private static string Field(Type36TemplateEvidence row, string field)
        {
            if (row.Content.TryGetValue(field, out var content)) return content.Known ? content.Evidence : null;
            if (row.UniqueIds.TryGetValue(field, out var id)) return id?.ToString("D");
            return row.Get(field);
        }
    }
}
#endif
