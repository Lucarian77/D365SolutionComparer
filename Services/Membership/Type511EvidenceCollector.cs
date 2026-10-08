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
using Microsoft.Xrm.Sdk.Metadata.Query;
using Microsoft.Xrm.Sdk.Query;

namespace D365SolutionComparer.Services.Membership
{
    /// <summary>Explicit metadata-first Debug investigation; never supplies production keys or absence evidence.</summary>
    internal sealed class Type511EvidenceCollector
    {
        internal const int BatchSize = 200, MaxIsolationGroups = 64;
        internal const string ProposedIdentifierField = "teamtemplatename";
        internal static readonly string[] DefinitionFields = { "defaultaccessrightsmask", "issystem" };
        private static readonly string[] InternalNames = { "uniquename", "schemaname", "logicalname" };
        private static readonly string[] AuditText = { "name", "displayname", "uniquename", "schemaname", "logicalname", "version" };
        private static readonly string[] UnsafePayload = { "binary", "attachment", "thumbnail", "image", "media", "package", "base64", "encoded", "secret", "secure", "credential", "token", "document" };
        private static readonly HashSet<string> AuditReferences = new HashSet<string>(new[] { "organizationid", "ownerid", "owninguser", "owningteam",
            "owningbusinessunit", "createdby", "createdonbehalfby", "modifiedby", "modifiedonbehalfby", "solutionid", "supportingsolutionid",
            "rootsolutionid" }, StringComparer.OrdinalIgnoreCase);
        private static bool AuditReference(string field) => AuditReferences.Contains(field) || field.EndsWith("idunique", StringComparison.OrdinalIgnoreCase);

        internal TeamTemplateEvidenceReport Capture(IOrganizationService sourceService, MembershipSnapshot source, string sourceVersion,
            IOrganizationService targetService, MembershipSnapshot target, string targetVersion, CancellationToken token, Action<string> progress = null)
        {
            if (source?.State != MembershipSnapshotState.Complete || target?.State != MembershipSnapshotState.Complete ||
                !StringComparer.OrdinalIgnoreCase.Equals(source.SolutionUniqueName, target.SolutionUniqueName))
                throw new ArgumentException("Completed snapshots of the same solution are required.");
            token.ThrowIfCancellationRequested();
            var report = new TeamTemplateEvidenceReport { Source = Read(sourceService, source, sourceVersion, token, progress),
                Target = Read(targetService, target, targetVersion, token, progress) };
            report.Analyze(token); return report;
        }

        private static TeamTemplateSideEvidence Read(IOrganizationService service, MembershipSnapshot snapshot, string version,
            CancellationToken token, Action<string> progress)
        {
            if (service == null) throw new ArgumentNullException(nameof(service));
            token.ThrowIfCancellationRequested();
            var side = new TeamTemplateSideEvidence { Snapshot = snapshot, Version = version };
            side.Raw.AddRange(snapshot.Components.Where(c => c.Record.ComponentType == 511));
            var ids = side.Raw.Where(c => c.Record.ObjectId.HasValue && c.Record.ObjectId != Guid.Empty)
                .Select(c => c.Record.ObjectId.Value).Distinct().OrderBy(id => id).ToArray();
            foreach (var id in ids) side.Rows[id] = new TeamTemplateRecordEvidence { ObjectId = id, Status = "Incomplete", Reason = "Backing entity/schema not verified" };
            if (ids.Length == 0) return side;

            if (!DiscoverBacking(service, side, token)) return side;
            var metadata = Schema(service, side.EntityName, side, token);
            if (metadata == null)
            { foreach (var row in side.Rows.Values) { row.Status = side.SchemaFailure; row.Reason = "Backing schema unavailable; no guessed columns"; } return side; }
            side.PrimaryId = metadata.PrimaryIdAttribute;
            side.Metadata = metadata;
            var columns = new List<string>(); var hashes = new HashSet<string>(StringComparer.Ordinal);
            foreach (var attribute in metadata.Attributes.OrderBy(a => a.LogicalName, StringComparer.Ordinal))
            {
                string name = attribute.LogicalName;
                bool shadow = metadata.Attributes.Where(a => a.AttributeType == AttributeTypeCode.Lookup).Any(a =>
                    StringComparer.OrdinalIgnoreCase.Equals(name, a.LogicalName + "name") || StringComparer.OrdinalIgnoreCase.Equals(name, a.LogicalName + "yominame")) ||
                    !string.IsNullOrWhiteSpace(attribute.AttributeOf) && name.EndsWith("name", StringComparison.OrdinalIgnoreCase);
                bool payload = UnsafePayload.Any(p => name.IndexOf(p, StringComparison.OrdinalIgnoreCase) >= 0) || name == "content";
                bool text = attribute.AttributeType == AttributeTypeCode.String || attribute.AttributeType == AttributeTypeCode.Memo;
                bool shape = text || attribute.AttributeType == AttributeTypeCode.Uniqueidentifier || attribute.AttributeType == AttributeTypeCode.Lookup ||
                    attribute.AttributeType == AttributeTypeCode.EntityName || attribute.AttributeType == AttributeTypeCode.Picklist ||
                    attribute.AttributeType == AttributeTypeCode.State || attribute.AttributeType == AttributeTypeCode.Status ||
                    attribute.AttributeType == AttributeTypeCode.Integer || attribute.AttributeType == AttributeTypeCode.BigInt || attribute.AttributeType == AttributeTypeCode.Owner ||
                    attribute.AttributeType == AttributeTypeCode.Boolean || attribute.AttributeType == AttributeTypeCode.DateTime;
                bool readable = attribute.IsValidForRead == true && shape && !shadow && !payload;
                bool scope = new[] { "objecttypecode", "entitylogicalname", "entityname", "primaryentity" }.Contains(name) &&
                    (attribute.AttributeType == AttributeTypeCode.Integer || attribute.AttributeType == AttributeTypeCode.EntityName || attribute.AttributeType == AttributeTypeCode.String);
                if (scope) side.ScopeFields.Add(name);
                bool hash = text && (attribute.AttributeType == AttributeTypeCode.Memo ||
                    !AuditText.Contains(name) && name != metadata.PrimaryNameAttribute && name != ProposedIdentifierField && !scope);
                side.Schema.Add(side.EntityName + "." + name + "; type=" + attribute.AttributeType + "; readable=" + attribute.IsValidForRead +
                    "; capture=" + (shadow ? "ExcludedLookupShadow" : payload ? "ExcludedPayload" : readable ? hash ? "HashOnly" : "Audit" : "UnavailableOrNotQueried"));
                if (readable) { columns.Add(name); if (hash) hashes.Add(name); }
            }
            if (!columns.Contains(side.PrimaryId) || metadata.Attributes.Single(a => a.LogicalName == side.PrimaryId).AttributeType != AttributeTypeCode.Uniqueidentifier)
            { foreach (var row in side.Rows.Values) row.Reason = "Readable GUID primary key not verified"; return side; }
            // An exposed but unreadable strongest identifier remains incomplete; the display
            // hypothesis cannot bypass it. Field presence/type is proven by metadata first.
            side.CandidateField = InternalNames.FirstOrDefault(field =>
                metadata.Attributes.Any(a => a.LogicalName == field && a.AttributeType == AttributeTypeCode.String));
            side.PrimaryNameField = columns.Contains(metadata.PrimaryNameAttribute) && !hashes.Contains(metadata.PrimaryNameAttribute) &&
                metadata.Attributes.Single(a => a.LogicalName == metadata.PrimaryNameAttribute).AttributeType == AttributeTypeCode.String
                ? metadata.PrimaryNameAttribute : null;
            side.CandidateRole = side.CandidateField == null ? "ScopedPrimaryNameHypothesis" : "ScopedInternalNameHypothesis";
            side.CandidateField = side.CandidateField ?? side.PrimaryNameField;
            var referenceFields = metadata.Attributes.OfType<LookupAttributeMetadata>().Where(a => !AuditReference(a.LogicalName)).Select(a => a.LogicalName)
                .Concat((metadata.ManyToOneRelationships ?? new OneToManyRelationshipMetadata[0]).Where(r => r != null &&
                    r.ReferencingEntity == side.EntityName && !AuditReference(r.ReferencingAttribute)).Select(r => r.ReferencingAttribute))
                .Distinct(StringComparer.Ordinal).OrderBy(f => f, StringComparer.Ordinal).ToArray();
            var critical = new[] { side.PrimaryId }.Concat(side.CandidateField == null || !columns.Contains(side.CandidateField) ? new string[0] : new[] { side.CandidateField })
                .Concat(referenceFields.Where(columns.Contains)).Concat(side.ScopeFields.Where(columns.Contains)).Distinct(StringComparer.Ordinal).ToArray();
            var backing = ReadFields(service, side.EntityName, side.PrimaryId, critical, columns.Except(critical).ToArray(), ids, side, token, progress);
            foreach (var id in ids)
            {
                token.ThrowIfCancellationRequested(); var row = side.Rows[id]; var found = backing[id]; row.Status = found.Status; row.Reason = found.Reason;
                if (found.Status != "Unique") continue;
                row.PrimaryId = found.Row.Id; row.CriticalComplete = found.CriticalComplete;
                row.RuntimeColumns.AddRange(found.Columns.OrderBy(c => c, StringComparer.Ordinal));
                foreach (var field in columns)
                    if (hashes.Contains(field)) row.Content[field] = found.Columns.Contains(field) ? Type31ContentFingerprint.Create(found.Row.GetAttributeValue<object>(field))
                        : new Type31ContentFingerprint { Presence = "Unavailable" };
                    else if (found.Row.GetAttributeValue<object>(field) is string text && text.Length > 512)
                        row.Content[field] = Type31ContentFingerprint.Create(text);
                    else row.Fields[field] = found.Columns.Contains(field) ? Format(found.Row.GetAttributeValue<object>(field)) : null;
                row.Managed = found.Row.GetAttributeValue<object>("ismanaged") is bool ? (bool?)found.Row.GetAttributeValue<bool>("ismanaged") : null;
                ResolveScope(snapshot, side, metadata, referenceFields, found, row);
                row.CandidateField = side.CandidateField;
                foreach (var field in InternalNames.Concat(new[] { ProposedIdentifierField }))
                {
                    var attribute = metadata.Attributes.SingleOrDefault(a => a.LogicalName == field);
                    row.IdentifierEvidence[field] = "metadataExposed=" + (attribute != null) + "; metadataReadable=" + attribute?.IsValidForRead +
                        "; runtimeRead=" + found.Columns.Contains(field) + "; value=" + (row.Get(field) ?? "Unavailable") +
                        "; hashOnly=" + row.Content.ContainsKey(field);
                }
                row.ProposedNameVerified = metadata.LogicalName == "teamtemplate" &&
                    metadata.Attributes.SingleOrDefault(a => a.LogicalName == ProposedIdentifierField)?.AttributeType == AttributeTypeCode.String &&
                    found.Columns.Contains(ProposedIdentifierField) && found.Row.GetAttributeValue<object>(ProposedIdentifierField) is string;
                foreach (var field in DefinitionFields)
                {
                    var attribute = metadata.Attributes.SingleOrDefault(a => a.LogicalName == field);
                    object value = found.Row.GetAttributeValue<object>(field);
                    bool valid = field == "issystem" ? attribute?.AttributeType == AttributeTypeCode.Boolean && (value == null || value is bool) :
                        (attribute?.AttributeType == AttributeTypeCode.Integer || attribute?.AttributeType == AttributeTypeCode.BigInt) && (value == null || value is int || value is long);
                    row.DefinitionValues[field] = valid && found.Columns.Contains(field) ? value == null ? "Null" : Format(value) : "Unavailable";
                }
            }
            ResolveEntityScopes(service, side, backing, token);
            foreach (var row in side.Rows.Values.Where(r => r.Status == "Unique"))
            {
                var value = side.CandidateField == null ? null : row.Get(side.CandidateField);
                if (row.CriticalComplete && row.ParentComplete && row.TableStatus == "Verified" &&
                    !string.IsNullOrWhiteSpace(value) && !Guid.TryParse(value, out var ignored))
                    row.CandidateA = "teamtemplate-candidate-a:" + Type31EvidenceCollector.Frame(side.EntityName, side.CandidateRole,
                        side.CandidateField, value.Trim(), row.TableKey, row.ParentKey);
                // B is independent context only; it also requires a verified table scope.
                var display = side.PrimaryNameField == null ? row.Get("name") : row.Get(side.PrimaryNameField);
                if (row.TableStatus == "Verified" && !string.IsNullOrWhiteSpace(display))
                    row.CandidateB = "teamtemplate-candidate-b:" + Type31EvidenceCollector.Frame(side.EntityName, row.TableKey, display.Trim());
                var blockers = new List<string>();
                string fixedName = row.Get(ProposedIdentifierField);
                if (!row.CriticalComplete) blockers.Add("Primary/critical backing evidence incomplete");
                if (!row.ProposedNameVerified || string.IsNullOrWhiteSpace(fixedName) || Guid.TryParse(fixedName, out var invalidName))
                    blockers.Add("Fixed teamtemplatename unavailable/blank/invalid/hash-only; no fallback");
                if (row.ObjectTypeCodeScopeStatus != "Verified" || row.TableStatus != "Verified" ||
                    !StringComparer.OrdinalIgnoreCase.Equals(row.ObjectTypeCodeScopeKey, row.TableKey))
                    blockers.Add("objecttypecode portable table mapping unresolved/ambiguous/conflicting");
                row.ProposedBlockingReason = blockers.Count == 0 ? "None; ScopedPrimaryNameHypothesis, lifecycle portability not proven" : string.Join("; ", blockers);
                if (blockers.Count == 0) row.CandidateP = "teamtemplate-candidate-p:v1:" +
                    Type31EvidenceCollector.Frame(row.ObjectTypeCodeScopeKey.Trim(), fixedName.Trim());
            }
            return side;
        }

        private static bool DiscoverBacking(IOrganizationService service, TeamTemplateSideEvidence side, CancellationToken token)
        {
            var definitions = side.Raw.Select(r => r.RegisteredDefinition).ToArray();
            var definition = definitions.FirstOrDefault(d => d != null);
            if (definition != null)
            {
                if (definitions.Any(d => d == null || d.ObjectTypeCode != 511 || !ValidName(d.PrimaryEntityName) ||
                    !StringComparer.OrdinalIgnoreCase.Equals(d.Name.Replace(" ", ""), "TeamTemplate") ||
                    !StringComparer.OrdinalIgnoreCase.Equals(d.PrimaryEntityName, definition.PrimaryEntityName)))
                { side.Discovery.Add("Incomplete/conflicting completed registered definitions; no guessed backing query"); return false; }
                side.EntityName = definition.PrimaryEntityName;
                side.Discovery.Add("Reused completed registered definition: name=" + definition.Name + "; primaryentityname=" + side.EntityName);
                return true;
            }
            // Existing Type 511 classification may use legacy diagnostics without a registered definition.
            // Discover that absence explicitly; never infer a backing mapping from the numeric type alone.
            try
            {
                var rows = new Dictionary<Guid, Entity>(); int page = 1; string cookie = null;
                while (true)
                {
                    token.ThrowIfCancellationRequested();
                    var query = new QueryExpression("solutioncomponentdefinition") { ColumnSet = new ColumnSet("objecttypecode", "name", "primaryentityname"),
                        PageInfo = new PagingInfo { Count = BatchSize, PageNumber = page, PagingCookie = cookie } };
                    query.Criteria.AddCondition("objecttypecode", ConditionOperator.Equal, 511);
                    query.AddOrder("solutioncomponentdefinitionid", OrderType.Ascending);
                    side.Requests.Add("RetrieveMultiple solutioncomponentdefinition; objecttypecode=511; columns=[objecttypecode,name,primaryentityname]; page=" + page);
                    var response = service.RetrieveMultiple(query); token.ThrowIfCancellationRequested();
                    if (response == null) { side.Discovery.Add("Incomplete registered-definition response"); return false; }
                    side.Pages.Add("solutioncomponentdefinition; rows=" + response.Entities.Count + "; page=" + page +
                        "; MoreRecords=" + response.MoreRecords + "; PagingCookieSupplied=" + !string.IsNullOrEmpty(response.PagingCookie));
                    int before = rows.Count;
                    foreach (var row in response.Entities)
                    {
                        var code = row?.GetAttributeValue<object>("objecttypecode");
                        if (row == null || row.LogicalName != "solutioncomponentdefinition" || row.Id == Guid.Empty ||
                            !(code is int || code is OptionSetValue) || (code is int ? (int)code : ((OptionSetValue)code).Value) != 511 ||
                            rows.ContainsKey(row.Id) && (response.Entities.Count(r => r.Id == row.Id) > 1 || !Type31EvidenceCollector.SameReturnedRow(rows[row.Id], row)))
                        { side.Discovery.Add("Conflicting/duplicate registered-definition rows; no backing query"); return false; }
                        rows[row.Id] = row;
                    }
                    if (!response.MoreRecords) break;
                    if (rows.Count == before || !string.IsNullOrEmpty(response.PagingCookie) && cookie == response.PagingCookie)
                    { side.Discovery.Add("Stalled registered-definition paging; no backing query"); return false; }
                    cookie = response.PagingCookie; page++;
                }
                if (rows.Count > 1) { side.Discovery.Add("Multiple registered-definition records; no backing query"); return false; }
                if (rows.Count == 1)
                {
                    var row = rows.Values.Single(); var name = row.GetAttributeValue<string>("name"); var backing = row.GetAttributeValue<string>("primaryentityname");
                    if (string.IsNullOrWhiteSpace(name) || !StringComparer.OrdinalIgnoreCase.Equals(name.Replace(" ", ""), "TeamTemplate") || !ValidName(backing))
                    { side.Discovery.Add("Registered definition has incomplete/conflicting family/backing fields"); return false; }
                    side.EntityName = backing; side.Discovery.Add("Discovered registered definition: name=" + name + "; primaryentityname=" + backing); return true;
                }
                side.Discovery.Add("No registered solutioncomponentdefinition for Type 511. This is not absence of Team Templates.");
                if (side.Raw.All(r => r.SemanticKind == ComponentSemanticKinds.TeamTemplate && r.DiagnosticEvidence.Any(e =>
                    e.StartsWith("TeamTemplate diagnostic lookup matched. teamtemplateid=" + r.Record.ObjectId?.ToString("D") + ";", StringComparison.Ordinal))))
                {
                    side.EntityName = "teamtemplate";
                    side.Discovery.Add("Backing hypothesis teamtemplate independently correlated by completed Type 511 diagnostics; schema/correlation rechecked here, no production key inferred");
                    return true;
                }
                side.Discovery.Add("Backing mapping unavailable: no consistent registered definition or verified completed backing correlation; no assumed scan");
                return false;
            }
            catch (OperationCanceledException) { throw; }
            catch (Exception error) { token.ThrowIfCancellationRequested(); side.Discovery.Add(SafeFault(error)); return false; }
        }

        private static void ResolveEntityScopes(IOrganizationService service, TeamTemplateSideEvidence side,
            Dictionary<Guid, LookupEvidence> backing, CancellationToken token)
        {
            var pending = new Dictionary<string, ScopeEvidence>(StringComparer.OrdinalIgnoreCase);
            var rowScopes = new Dictionary<Guid, List<ScopeEvidence>>();
            var objectTypeScopes = new Dictionary<Guid, ScopeEvidence>();
            foreach (var row in side.Rows.Values.Where(r => r.Status == "Unique"))
            {
                var scopes = new List<ScopeEvidence>(); rowScopes[row.ObjectId] = scopes;
                foreach (var key in row.TableKeys) scopes.Add(new ScopeEvidence { Status = "Verified", Key = key, Reason = "Reused verified snapshot Table identity" });
                foreach (var field in side.ScopeFields)
                {
                    var found = backing[row.ObjectId]; var raw = found.Row.GetAttributeValue<object>(field); var text = (raw as string)?.Trim();
                    int number;
                    bool numeric = raw is int || int.TryParse(text, NumberStyles.Integer, CultureInfo.InvariantCulture, out number);
                    var input = numeric ? "code:" + (raw is int ? (int)raw : int.Parse(text, CultureInfo.InvariantCulture)).ToString(CultureInfo.InvariantCulture) : "name:" + text;
                    if (!found.Columns.Contains(field) || raw == null || !numeric && !LogicalName(text))
                    {
                        var invalid = new ScopeEvidence { Reason = field + ": missing/faulted/invalid entity scope" };
                        scopes.Add(invalid); if (field == "objecttypecode") objectTypeScopes[row.ObjectId] = invalid;
                        continue;
                    }
                    if (!pending.TryGetValue(input, out var scope))
                    {
                        scope = new ScopeEvidence { Input = input, Number = numeric ? (int?)(raw is int ? (int)raw : int.Parse(text, CultureInfo.InvariantCulture)) : null,
                            Name = numeric ? null : text, Reason = "Metadata scope not yet verified" };
                        pending.Add(input, scope);
                        if (!numeric) ApplySnapshotScope(side.Snapshot, text, scope);
                    }
                    scopes.Add(scope);
                    if (field == "objecttypecode") objectTypeScopes[row.ObjectId] = scope;
                }
            }
            var unresolved = pending.Values.Where(s => s.Status == "Incomplete").OrderBy(s => s.Input, StringComparer.Ordinal).ToArray();
            for (int offset = 0; offset < unresolved.Length; offset += BatchSize)
            {
                token.ThrowIfCancellationRequested(); var batch = unresolved.Skip(offset).Take(BatchSize).ToArray();
                var query = new EntityQueryExpression { Properties = new MetadataPropertiesExpression("MetadataId", "LogicalName", "ObjectTypeCode"),
                    Criteria = new MetadataFilterExpression(LogicalOperator.Or) };
                foreach (var scope in batch) query.Criteria.Conditions.Add(new MetadataConditionExpression(scope.Number.HasValue ? "ObjectTypeCode" : "LogicalName",
                    MetadataConditionOperator.Equals, scope.Number.HasValue ? (object)scope.Number.Value : scope.Name));
                side.Requests.Add("Execute RetrieveMetadataChanges; selected entity scopes=[" + string.Join(",", batch.Select(s => s.Input)) + "]; properties=[MetadataId,LogicalName,ObjectTypeCode]");
                try
                {
                    var response = service.Execute(new RetrieveMetadataChangesRequest { Query = query }) as RetrieveMetadataChangesResponse;
                    token.ThrowIfCancellationRequested();
                    if (response?.EntityMetadata == null || response.EntityMetadata.Any(m => m == null ||
                        !batch.Any(s => s.Number.HasValue ? m.ObjectTypeCode == s.Number : StringComparer.OrdinalIgnoreCase.Equals(m.LogicalName, s.Name))))
                    { foreach (var scope in batch) scope.Reason = "Incomplete/foreign scoped metadata response"; continue; }
                    foreach (var scope in batch)
                    {
                        var matches = response.EntityMetadata.Where(m => scope.Number.HasValue ? m.ObjectTypeCode == scope.Number :
                            StringComparer.OrdinalIgnoreCase.Equals(m.LogicalName, scope.Name)).ToArray();
                        if (matches.Length > 1) { scope.Status = "Ambiguous"; scope.Reason = "Multiple entity metadata candidates for selected scope"; }
                        else if (matches.Length == 1 && matches[0].MetadataId.HasValue && matches[0].MetadataId != Guid.Empty && LogicalName(matches[0].LogicalName))
                        {
                            var local = side.Snapshot.Components.Where(c => c.Record.ComponentType == 1 && c.Record.ObjectId == matches[0].MetadataId).ToArray();
                            if (local.Any(c => c.Status != IdentityResolutionStatus.Resolved || !StringComparer.OrdinalIgnoreCase.Equals(c.ComparisonKey, matches[0].LogicalName)))
                            { scope.Status = local.Any(c => c.Status == IdentityResolutionStatus.Ambiguous) ? "Ambiguous" : "Incomplete"; scope.Reason = "Conflicting/incomplete snapshot Table correlation"; continue; }
                            if (ApplySnapshotScope(side.Snapshot, matches[0].LogicalName, scope) && scope.Status != "Verified") continue;
                            scope.Status = "Verified"; scope.Key = matches[0].LogicalName.Trim().ToLowerInvariant(); scope.Reason = "Unique scoped published entity metadata; GUID local correlation only";
                        }
                        else scope.Reason = "Missing/incomplete entity metadata for selected scope";
                    }
                }
                catch (OperationCanceledException) { throw; }
                catch (Exception error) { token.ThrowIfCancellationRequested(); foreach (var scope in batch) { scope.Status = "Faulted"; scope.Reason = SafeFault(error); } }
            }
            foreach (var row in side.Rows.Values.Where(r => r.Status == "Unique"))
            {
                var scopes = rowScopes[row.ObjectId]; var keys = scopes.Where(s => s.Status == "Verified").Select(s => s.Key).Distinct(StringComparer.OrdinalIgnoreCase).ToArray();
                row.TableStatus = scopes.Any(s => s.Status == "Ambiguous") || keys.Length > 1 ? "Ambiguous" :
                    scopes.Count > 0 && scopes.All(s => s.Status == "Verified") && keys.Length == 1 ? "Verified" : "Incomplete";
                row.TableKey = row.TableStatus == "Verified" ? keys[0] : null;
                objectTypeScopes.TryGetValue(row.ObjectId, out var objectTypeScope);
                row.ObjectTypeCodeScopeStatus = objectTypeScope?.Status ?? "Incomplete";
                row.ObjectTypeCodeScopeKey = objectTypeScope?.Status == "Verified" ? objectTypeScope.Key : null;
                foreach (var scope in scopes) row.Context.Add("Entity scope: " + scope.Input + "; " + scope.Status + "; logicalName=" + scope.Key + "; " + scope.Reason);
                if (scopes.Count == 0) row.Context.Add("No independently verified entity/table scope exposed; no global/tableless assumption");
                if (keys.Length > 1) row.Context.Add("Conflicting independently resolved parent table scopes");
            }
        }

        private static bool ApplySnapshotScope(MembershipSnapshot snapshot, string name, ScopeEvidence scope)
        {
            var matches = snapshot.Components.Where(c => c.Record.ComponentType == 1 && StringComparer.OrdinalIgnoreCase.Equals(c.ComparisonKey, name)).ToArray();
            if (matches.Length == 0) return false;
            bool unique = matches.All(c => c.Status == IdentityResolutionStatus.Resolved && c.InventoryAbsencePolicy == InventoryAbsencePolicy.CompleteInventory) &&
                matches.Select(c => c.Record.ObjectId).Distinct().Count() == 1;
            scope.Status = unique ? "Verified" : "Ambiguous"; scope.Key = unique ? name.Trim().ToLowerInvariant() : null;
            scope.Reason = unique ? "Reused verified snapshot Table identity" : "Ambiguous snapshot portable Table identity"; return true;
        }
        private static bool LogicalName(string value) => ValidName(value) && value.Length <= 128 && !StringComparer.OrdinalIgnoreCase.Equals(value, "none") &&
            (char.IsLetter(value[0]) || value[0] == '_');
        private sealed class ScopeEvidence
        { internal string Status = "Incomplete", Input, Name, Key, Reason; internal int? Number; }

        private static void ResolveScope(MembershipSnapshot snapshot, TeamTemplateSideEvidence side, EntityMetadata metadata,
            string[] referenceFields, LookupEvidence found, TeamTemplateRecordEvidence row)
        {
            var keys = new SortedSet<string>(StringComparer.OrdinalIgnoreCase); var reasons = new List<string>(); bool ambiguous = false;
            foreach (var field in referenceFields)
            {
                var raw = found.Row.GetAttributeValue<object>(field); var reference = raw as EntityReference;
                string entity = reference?.LogicalName; Guid? id = reference?.Id;
                var links = (metadata.ManyToOneRelationships ?? new OneToManyRelationshipMetadata[0]).Where(r => r != null &&
                    r.ReferencingEntity == side.EntityName && r.ReferencingAttribute == field).ToArray();
                if (links.Length > 1) { ambiguous = true; reasons.Add(field + ": multiple metadata parent/reference relationships"); continue; }
                if (raw is Guid && links.Length == 1 && PrimaryFor(links[0].ReferencedEntity) == links[0].ReferencedAttribute)
                { id = (Guid)raw; entity = links[0].ReferencedEntity; }
                var lookup = metadata.Attributes.OfType<LookupAttributeMetadata>().SingleOrDefault(a => a.LogicalName == field);
                bool targetValid = reference != null ? lookup?.Targets?.Contains(entity) == true &&
                    (links.Length == 0 || links[0].ReferencedEntity == entity && links[0].ReferencedAttribute == PrimaryFor(entity)) : raw is Guid && links.Length == 1;
                if (!found.Columns.Contains(field) || !targetValid || !id.HasValue || id == Guid.Empty || !RawTypeFor(entity).HasValue)
                { reasons.Add(field + ": parent/reference identity unavailable or unsupported; no GUID fallback"); continue; }
                int type = RawTypeFor(entity).Value;
                var matches = snapshot.Components.Where(c => c.Record.ComponentType == type && c.Record.ObjectId == id).ToArray();
                var identities = matches.Where(c => c.Status == IdentityResolutionStatus.Resolved && !string.IsNullOrWhiteSpace(c.ComparisonKey) &&
                    c.InventoryAbsencePolicy == InventoryAbsencePolicy.CompleteInventory).Select(c => Type31EvidenceCollector.Frame(c.SemanticKind, c.ComparisonKey))
                    .Distinct(StringComparer.OrdinalIgnoreCase).ToArray();
                bool unique = identities.Length == 1 && snapshot.Components.Where(c => c.Status == IdentityResolutionStatus.Resolved &&
                    StringComparer.OrdinalIgnoreCase.Equals(Type31EvidenceCollector.Frame(c.SemanticKind, c.ComparisonKey), identities[0]))
                    .Select(c => c.Record.ObjectId).Distinct().Count() == 1;
                if (matches.Any(c => c.Status == IdentityResolutionStatus.Ambiguous) || identities.Length > 1 || identities.Length == 1 && !unique)
                { ambiguous = true; reasons.Add(field + ": ambiguous parent identity in completed snapshot"); }
                else if (matches.Length == 0 || matches.Any(c => c.Status != IdentityResolutionStatus.Resolved || c.InventoryAbsencePolicy != InventoryAbsencePolicy.CompleteInventory) || !unique)
                    reasons.Add(field + ": missing/incomplete verified parent identity in completed snapshot");
                else
                {
                    keys.Add(Type31EvidenceCollector.Frame(field, identities[0])); row.Context.Add(field + ": reused verified snapshot identity " + identities[0]);
                    if (type == 1) row.TableKeys.Add(matches[0].ComparisonKey.Trim().ToLowerInvariant());
                }
            }
            row.ParentComplete = reasons.Count == 0;
            row.ParentStatus = ambiguous ? "Ambiguous" : row.ParentComplete ? referenceFields.Length == 0 ? "NoParentRelationshipExposed" : "VerifiedSnapshotScope" : "Incomplete";
            row.ParentKey = row.ParentComplete ? referenceFields.Length == 0 ? "NoParentRelationshipExposed" : Type31EvidenceCollector.Frame(keys.ToArray()) : null;
            row.Context.AddRange(reasons);
            row.Context.Add("Scope interpretation is diagnostic only. Unexposed relationships are not proof of global uniqueness or lifecycle portability.");
        }

        private static int? RawTypeFor(string entity)
        {
            switch (entity) { case "entity": return 1; case "attribute": return 2; case "relationship": return 10;
                case "webresource": return 61; case "appmodule": return 80; case "savedquery": return 26; case "workflow": return 29;
                case "savedqueryvisualization": return 59; case "systemform": return 60; case "sitemap": return 62;
                case "pluginassembly": return 91; case "sdkmessageprocessingstep": return 92; case "environmentvariabledefinition": return 380; default: return null; }
        }
        private static string PrimaryFor(string entity) => RawTypeFor(entity).HasValue ? entity == "systemform" ? "formid" : entity + "id" : null;
        private static EntityMetadata Schema(IOrganizationService service, string entity, TeamTemplateSideEvidence side, CancellationToken token)
        {
            token.ThrowIfCancellationRequested();
            side.Requests.Add("Execute RetrieveEntity(" + entity + ", Entity|Attributes|Relationships, RetrieveAsIfPublished=False)");
            try
            {
                var response = service.Execute(new RetrieveEntityRequest { LogicalName = entity,
                    EntityFilters = EntityFilters.Entity | EntityFilters.Attributes | EntityFilters.Relationships, RetrieveAsIfPublished = false }) as RetrieveEntityResponse;
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
            string[] critical, string[] optional, Guid[] ids, TeamTemplateSideEvidence side, CancellationToken token, Action<string> progress)
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
            Dictionary<Guid, LookupEvidence> current, bool critical, TeamTemplateSideEvidence side, CancellationToken token,
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
            Guid[] ids, TeamTemplateSideEvidence side, CancellationToken token, Action<string> progress)
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

    internal sealed class TeamTemplateSideEvidence
    {
        internal MembershipSnapshot Snapshot;
        internal string Version, EntityName, PrimaryId, CandidateField, PrimaryNameField, CandidateRole, SchemaFailure;
        internal EntityMetadata Metadata;
        internal readonly List<string> Discovery = new List<string>(), ScopeFields = new List<string>();
        internal readonly List<ComponentIdentity> Raw = new List<ComponentIdentity>();
        internal readonly SortedDictionary<Guid, TeamTemplateRecordEvidence> Rows = new SortedDictionary<Guid, TeamTemplateRecordEvidence>();
        internal readonly List<string> Schema = new List<string>(), Relationships = new List<string>(), Requests = new List<string>(), Pages = new List<string>(), RetrievalDiagnostics = new List<string>();
    }
    internal sealed class TeamTemplateRecordEvidence
    {
        internal Guid ObjectId; internal Guid? PrimaryId; internal bool? Managed;
        internal string Status, Reason, CandidateField, CandidateA, CandidateB;
        internal bool CriticalComplete, DuplicateA, DuplicateB, ParentComplete;
        internal string ParentKey, ParentStatus, TableKey, TableStatus;
        internal string CandidateP, ProposedBlockingReason, ObjectTypeCodeScopeKey, ObjectTypeCodeScopeStatus;
        internal bool ProposedNameVerified, DuplicateP;
        internal bool CompleteP => CandidateP != null && !DuplicateP;
        internal readonly SortedDictionary<string, string> IdentifierEvidence = new SortedDictionary<string, string>(StringComparer.Ordinal);
        internal readonly SortedDictionary<string, string> DefinitionValues = new SortedDictionary<string, string>(StringComparer.Ordinal);
        internal readonly SortedSet<string> TableKeys = new SortedSet<string>(StringComparer.OrdinalIgnoreCase);
        internal readonly List<string> Context = new List<string>();
        internal readonly SortedDictionary<string, string> Fields = new SortedDictionary<string, string>(StringComparer.Ordinal);
        internal readonly SortedDictionary<string, Type31ContentFingerprint> Content = new SortedDictionary<string, Type31ContentFingerprint>(StringComparer.Ordinal);
        internal readonly List<string> RuntimeColumns = new List<string>();
        internal string Get(string field) => Fields.TryGetValue(field, out var value) ? value : null;
        internal string Evidence(string field) => Content.TryGetValue(field, out var hash) ? hash.Evidence : Get(field);
    }
    internal sealed class TeamTemplatePairEvidence
    {
        internal TeamTemplateRecordEvidence Source, Target; internal string Outcome;
        internal readonly HashSet<string> Categories = new HashSet<string>(StringComparer.Ordinal);
    }
    internal sealed class TeamTemplateEvidenceReport
    {
        internal TeamTemplateSideEvidence Source, Target;
        internal readonly List<TeamTemplatePairEvidence> Pairs = new List<TeamTemplatePairEvidence>();
        internal readonly List<TeamTemplatePairEvidence> ProposedPairs = new List<TeamTemplatePairEvidence>();
        internal void Analyze(CancellationToken token)
        {
            foreach (var side in new[] { Source, Target })
            {
                foreach (var group in side.Rows.Values.Where(r => r.CandidateA != null).GroupBy(r => r.CandidateA, StringComparer.OrdinalIgnoreCase).Where(g => g.Count() > 1))
                    foreach (var row in group) row.DuplicateA = true;
                foreach (var group in side.Rows.Values.Where(r => r.CandidateB != null).GroupBy(r => r.CandidateB, StringComparer.OrdinalIgnoreCase).Where(g => g.Count() > 1))
                    foreach (var row in group) row.DuplicateB = true;
                foreach (var group in side.Rows.Values.Where(r => r.CandidateP != null).GroupBy(r => r.CandidateP, StringComparer.OrdinalIgnoreCase).Where(g => g.Count() > 1))
                    foreach (var row in group) row.DuplicateP = true;
            }
            foreach (var key in Source.Rows.Values.Concat(Target.Rows.Values).Where(r => r.CandidateP != null).Select(r => r.CandidateP)
                .Distinct(StringComparer.OrdinalIgnoreCase).OrderBy(k => k, StringComparer.OrdinalIgnoreCase))
            {
                token.ThrowIfCancellationRequested();
                var left = Source.Rows.Values.Where(r => StringComparer.OrdinalIgnoreCase.Equals(r.CandidateP, key)).ToArray();
                var right = Target.Rows.Values.Where(r => StringComparer.OrdinalIgnoreCase.Equals(r.CandidateP, key)).ToArray();
                if (left.Length > 1 || right.Length > 1)
                {
                    foreach (var row in left) ProposedPairs.Add(new TeamTemplatePairEvidence { Source = row, Outcome = "Ambiguous" });
                    foreach (var row in right) ProposedPairs.Add(new TeamTemplatePairEvidence { Target = row, Outcome = "Ambiguous" });
                }
                else ProposedPairs.Add(new TeamTemplatePairEvidence { Source = left.SingleOrDefault(), Target = right.SingleOrDefault(),
                    Outcome = left.Length == 1 && right.Length == 1 ? "ScopedPrimaryNamePairHypothesis" : "OneSidedEvidence" });
            }
            foreach (var side in new[] { Source, Target }) foreach (var row in side.Rows.Values.Where(r => r.CandidateP == null))
                ProposedPairs.Add(new TeamTemplatePairEvidence { Source = side == Source ? row : null, Target = side == Target ? row : null,
                    Outcome = row.Status == "Duplicate" || row.ObjectTypeCodeScopeStatus == "Ambiguous" || row.TableStatus == "Ambiguous" ? "Ambiguous" : "Incomplete" });
            foreach (var pair in ProposedPairs)
            {
                pair.Categories.Add(pair.Outcome);
                if (pair.Outcome != "ScopedPrimaryNamePairHypothesis") continue;
                pair.Categories.Add(pair.Source.PrimaryId == pair.Target.PrimaryId ? "SamePrimaryId" : "DifferentPrimaryId");
                string leftUnique = pair.Source.Get("componentidunique"), rightUnique = pair.Target.Get("componentidunique");
                pair.Categories.Add(string.IsNullOrWhiteSpace(leftUnique) || string.IsNullOrWhiteSpace(rightUnique) ? "UniqueIdUnavailable" :
                    StringComparer.OrdinalIgnoreCase.Equals(leftUnique, rightUnique) ? "SameUniqueId" : "DifferentUniqueId");
                if (pair.Source.Managed.HasValue && pair.Target.Managed.HasValue && pair.Source.Managed != pair.Target.Managed)
                { pair.Categories.Add("ManagedTransition"); if (pair.Source.Managed == false && pair.Target.Managed == true) pair.Categories.Add("UnmanagedToManaged"); }
            }
            foreach (var key in Source.Rows.Values.Concat(Target.Rows.Values).Where(r => r.CandidateA != null).Select(r => r.CandidateA)
                .Distinct(StringComparer.OrdinalIgnoreCase).OrderBy(k => k, StringComparer.OrdinalIgnoreCase))
            {
                token.ThrowIfCancellationRequested();
                var left = Source.Rows.Values.Where(r => StringComparer.OrdinalIgnoreCase.Equals(r.CandidateA, key)).ToArray();
                var right = Target.Rows.Values.Where(r => StringComparer.OrdinalIgnoreCase.Equals(r.CandidateA, key)).ToArray();
                if (left.Length > 1 || right.Length > 1)
                {
                    foreach (var row in left) Pairs.Add(new TeamTemplatePairEvidence { Source = row, Outcome = "Ambiguous" });
                    foreach (var row in right) Pairs.Add(new TeamTemplatePairEvidence { Target = row, Outcome = "Ambiguous" });
                }
                else Pairs.Add(new TeamTemplatePairEvidence { Source = left.SingleOrDefault(), Target = right.SingleOrDefault(),
                    Outcome = left.Length == 1 && right.Length == 1 ? "SemanticPair" : "OneSidedEvidence" });
            }
            foreach (var side in new[] { Source, Target })
            {
                foreach (var row in side.Rows.Values.Where(r => r.CandidateA == null))
                    Pairs.Add(new TeamTemplatePairEvidence { Source = side == Source ? row : null, Target = side == Target ? row : null,
                        Outcome = row.Status == "Duplicate" || row.ParentStatus == "Ambiguous" || row.TableStatus == "Ambiguous" ? "Ambiguous" : "Incomplete" });
                foreach (var raw in side.Raw.Where(r => !r.Record.ObjectId.HasValue || r.Record.ObjectId == Guid.Empty))
                    Pairs.Add(new TeamTemplatePairEvidence { Outcome = "Incomplete" });
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
            var text = new StringBuilder(); text.AppendLine("TYPE 511 TEAM TEMPLATE EVIDENCE - DEBUG ONLY");
            text.AppendLine("Evidence only. Type 511 remains Unsupported / Indeterminate; no portable comparison key or absence proof is supplied.");
            text.AppendLine("Uniqueness is scoped to selected solution members; rename/recreate portability remains unproven. Candidate B and hashes never repair Candidate A.");
            text.AppendLine("\nRAW TYPE 511 MEMBERSHIP");
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
            text.AppendLine("\nREGISTERED-DEFINITION / BACKING-ENTITY DISCOVERY");
            EachSide(text, (s, label) => { foreach (var item in s.Discovery) Line(text, label, item); });
            text.AppendLine("\nBACKING-RECORD CORRELATION");
            EachSide(text, (s, label) => { foreach (var row in s.Rows.Values) Line(text, label, "objectid=" + row.ObjectId, "primaryId=" + row.PrimaryId, "backingEntity=" + s.EntityName,
                "objectid==primaryId=" + (row.PrimaryId.HasValue ? (row.ObjectId == row.PrimaryId.Value).ToString() : "Unknown"), row.Status, row.Reason); });
            text.AppendLine("\nREADABLE / UNAVAILABLE SCHEMA"); EachSide(text, (s, label) => { foreach (var item in s.Schema) Line(text, label, item); });
            text.AppendLine("\nRUNTIME-READABLE / FAULTED COLUMNS"); EachSide(text, (s, label) => {
                foreach (var item in s.RetrievalDiagnostics) Line(text, label, item);
                foreach (var row in s.Rows.Values) Line(text, label, row.ObjectId, "runtimeSucceeded=[" + string.Join(",", row.RuntimeColumns) + "]", "criticalComplete=" + row.CriticalComplete);
                foreach (var item in s.Pages) Line(text, label, item);
            });
            text.AppendLine("\nPARENT / REFERENCE RELATIONSHIP DISCOVERY"); EachSide(text, (s, label) => { foreach (var item in s.Relationships) Line(text, label, item); });
            text.AppendLine("\nENTITY / TABLE SCOPE RESOLUTION"); EachSide(text, (s, label) => { foreach (var row in s.Rows.Values)
                Line(text, label, row.ObjectId, "EntityScopeStatus=" + row.TableStatus, "PortableTable=" + row.TableKey,
                    "ParentStatus=" + row.ParentStatus, "MetadataConfirmedScopeFields=[" + string.Join(",", s.ScopeFields) + "]"); });
            text.AppendLine("\nCANDIDATE A / B ANALYSIS"); EachSide(text, (s, label) => {
                Line(text, label, "BackingEntity=" + s.EntityName, "Candidate A hypothesis: strongest internal text field=" + (s.CandidateField ?? "Unavailable"),
                    "CandidateRole=" + s.CandidateRole,
                    "A: independently verified table scope + metadata-confirmed internal identifier, or scoped primary-name hypothesis when no internal identifier exists; never name alone",
                    "B: verified table scope + descriptive/primary name; context only; B/access-rights/content hashes/local IDs never repair A");
                foreach (var row in s.Rows.Values) {
                    Line(text, label, row.ObjectId, "CandidateA=" + (row.CandidateA ?? "Incomplete"), "CandidateB=" + (row.CandidateB ?? "NotAvailable"),
                        "DuplicateA=" + row.DuplicateA, "DuplicateB=" + row.DuplicateB, "CompleteA=" + (row.CandidateA != null && !row.DuplicateA),
                        "ParentStatus=" + row.ParentStatus, "ParentIdentityComplete=" + row.ParentComplete, "TableScopeStatus=" + row.TableStatus,
                        "BlockingReason=" + (row.CandidateA != null && !row.DuplicateA ? "None" : row.DuplicateA ? "Candidate A collision" :
                        "Incomplete/faulted/ambiguous primary correlation, critical identifier, independent table scope or parent reference"));
                    foreach (var item in row.Context) Line(text, label, row.ObjectId, item);
                }
            });
            text.AppendLine("\nDUPLICATE / COLLISION ANALYSIS"); EachSide(text, (s, label) => {
                foreach (var group in s.Raw.Where(r => r.Record.ObjectId.HasValue).GroupBy(r => r.Record.ObjectId).Where(g => g.Count() > 1))
                    Line(text, label, "RepeatedRawObjectId=" + group.Key, "references=" + group.Count(), "not distinct backing identities");
                Line(text, label, "CandidateACollisionGroups=" + s.Rows.Values.Where(r => r.DuplicateA).Select(r => r.CandidateA).Distinct(StringComparer.OrdinalIgnoreCase).Count(),
                    "CandidateBCollisionGroups=" + s.Rows.Values.Where(r => r.DuplicateB).Select(r => r.CandidateB).Distinct(StringComparer.OrdinalIgnoreCase).Count());
            });
            text.AppendLine("\nFIXED IDENTIFIER / CANDIDATE P READINESS");
            text.AppendLine("Candidate P = independently verified objecttypecode portable table logical name + fixed teamtemplatename. Trim + ordinal case-insensitive. Classification=ScopedPrimaryNameHypothesis, not an internal semantic identifier.");
            text.AppendLine("No GUID/display/hash/access-rights fallback. Candidate A is unchanged; A/B/content/deployment observations cannot repair P.");
            EachSide(text, (s, label) => {
                Line(text, label, "PrimaryNameAttribute=" + (s.Metadata?.PrimaryNameAttribute ?? "Unavailable"),
                    "StrongerInternalIdentifiersExposed=[" + string.Join(",", (s.Metadata?.Attributes ?? new AttributeMetadata[0])
                        .Where(a => new[] { "uniquename", "schemaname", "logicalname" }.Contains(a.LogicalName)).Select(a => a.LogicalName)) + "]");
                foreach (var row in s.Rows.Values)
                {
                    foreach (var item in row.IdentifierEvidence) Line(text, label, row.ObjectId, item.Key, item.Value);
                    Line(text, label, row.ObjectId, "ObjectTypeCode=" + row.Get("objecttypecode"),
                        "ObjectTypeCodeScopeStatus=" + row.ObjectTypeCodeScopeStatus, "PortableTable=" + row.ObjectTypeCodeScopeKey,
                        "CandidateP=" + (row.CandidateP ?? "Incomplete"), "CompleteP=" + row.CompleteP,
                        "BlockingReason=" + (row.DuplicateP ? "Candidate P collision" : row.ProposedBlockingReason ?? row.Reason));
                }
                Line(text, label, "CandidateACollisionGroups=" + s.Rows.Values.Where(r => r.DuplicateA).Select(r => r.CandidateA).Distinct(StringComparer.OrdinalIgnoreCase).Count(),
                    "CandidatePCollisionGroups=" + s.Rows.Values.Where(r => r.DuplicateP).Select(r => r.CandidateP).Distinct(StringComparer.OrdinalIgnoreCase).Count());
            });
            text.AppendLine("\nINITIAL DEFINITION EVIDENCE");
            text.AppendLine("defaultaccessrightsmask: potential initial behavioral definition property controlling access granted by template-created teams; exact numeric comparison, never identity. issystem: system/custom classification evidence; lifecycle semantics need review before any definition contract. Description: presence/length/SHA-256 only.");
            EachSide(text, (s, label) => {
                foreach (var row in s.Rows.Values)
                {
                    foreach (var item in row.DefinitionValues) Line(text, label, row.ObjectId, item.Key + "=" + item.Value);
                    Line(text, label, row.ObjectId, "description=" + DescriptionEvidence(row));
                }
            });
            text.AppendLine("\nCANDIDATE P LIFECYCLE COMPARISON");
            foreach (var pair in ProposedPairs)
            {
                Line(text, pair.Outcome, "Source=" + pair.Source?.PrimaryId, "Target=" + pair.Target?.PrimaryId,
                    "CandidateP=" + (pair.Source?.CandidateP ?? pair.Target?.CandidateP ?? "Incomplete"));
                if (pair.Outcome != "ScopedPrimaryNamePairHypothesis") continue;
                Line(text, "teamtemplateid=" + (pair.Source.PrimaryId == pair.Target.PrimaryId ? "EqualObserved" : "DifferentObserved"),
                    "componentidunique=" + EqualEvidence(pair.Source.Get("componentidunique"), pair.Target.Get("componentidunique")),
                    "CandidateA=" + EqualEvidence(pair.Source.CandidateA, pair.Target.CandidateA),
                    "CandidateP=" + EqualEvidence(pair.Source.CandidateP, pair.Target.CandidateP),
                    "SourceManaged=" + pair.Source.Managed, "TargetManaged=" + pair.Target.Managed,
                    "ManagedTransition=" + pair.Categories.Contains("ManagedTransition"));
                foreach (var field in Type511EvidenceCollector.DefinitionFields)
                    Line(text, field, "Source=" + DefinitionValue(pair.Source, field), "Target=" + DefinitionValue(pair.Target, field),
                        "Comparison=" + EqualEvidence(DefinitionValue(pair.Source, field), DefinitionValue(pair.Target, field)));
                Line(text, "description", "Source=" + DescriptionEvidence(pair.Source), "Target=" + DescriptionEvidence(pair.Target),
                    "Comparison=" + (KnownDescription(pair.Source) && KnownDescription(pair.Target) ?
                        StringComparer.Ordinal.Equals(DescriptionEvidence(pair.Source), DescriptionEvidence(pair.Target)) ? "EqualObserved" : "DifferentObserved" : "Incomplete"));
            }
            text.AppendLine("\nCANDIDATE P LIFECYCLE MATRIX");
            foreach (var category in new[] { "ScopedPrimaryNamePairHypothesis", "SamePrimaryId", "DifferentPrimaryId", "SameUniqueId", "DifferentUniqueId", "UniqueIdUnavailable", "ManagedTransition", "UnmanagedToManaged", "OneSidedEvidence", "Ambiguous", "Incomplete" })
                Line(text, category, ProposedPairs.Count(p => p.Categories.Contains(category)));
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
            text.AppendLine("\nPORTABILITY ASSESSMENT");
            Line(text, "Unique semantic pairs=" + Pairs.Count(p => p.Outcome == "SemanticPair"), "differing primary IDs=" + Pairs.Count(p => p.Categories.Contains("DifferentPrimaryId")));
            text.AppendLine("Observed evidence does not establish lifecycle portability or safe absence semantics. Additional live/lifecycle review required before production promotion or any production use.");
            text.AppendLine("\nCANDIDATE P PORTABILITY GATE");
            Line(text, "UniqueCandidatePPairs=" + ProposedPairs.Count(p => p.Outcome == "ScopedPrimaryNamePairHypothesis"),
                "DifferentPrimaryIdCandidatePPairs=" + ProposedPairs.Count(p => p.Categories.Contains("DifferentPrimaryId")));
            text.AppendLine("Same teamtemplateid is audit evidence only and never proves portability. Candidate P cannot be promoted solely from the current EDU same-primary-ID pair, even with an unmanaged-to-managed transition or matching definition hashes.");
            text.AppendLine("Minimum additional evidence: at least one independently deployed semantic Team Template pair with different teamtemplateid values, the same independently verified table scope, equal Candidate P, and zero ambiguity/collisions. DifferentPrimaryId counts alone do not establish independent deployment; that provenance requires live review. No environment-wide scan or absence inference is authorized.");
            text.AppendLine("\nEXACT REQUEST LEDGER"); EachSide(text, (s, label) => {
                foreach (var request in s.Requests) Line(text, label, request);
                foreach (var group in s.Requests.GroupBy(r => r.StartsWith("Execute RetrieveEntity(", StringComparison.Ordinal) ? r.Split(',')[0] + ")" :
                    r.StartsWith("Execute RetrieveMetadataChanges", StringComparison.Ordinal) ? "Execute RetrieveMetadataChanges" : r.Split(';')[0]))
                    Line(text, label, group.Key, "reads=" + group.Count());
                Line(text, label, "TotalReads=" + s.Requests.Count, "AdditionalWhoAmI=0", "Writes=0", "NormalMembershipEvidenceRequests=0");
            }); return text.ToString();
        }
        private void EachSide(StringBuilder text, Action<TeamTemplateSideEvidence, string> action) { action(Source, "Source"); action(Target, "Target"); }
        private static string EqualEvidence(string source, string target) => string.IsNullOrWhiteSpace(source) || string.IsNullOrWhiteSpace(target) || source == "Unavailable" || target == "Unavailable"
            ? "Incomplete" : StringComparer.OrdinalIgnoreCase.Equals(source, target) ? "EqualObserved" : "DifferentObserved";
        private static string DefinitionValue(TeamTemplateRecordEvidence row, string field) => row.DefinitionValues.TryGetValue(field, out var value) ? value : "Unavailable";
        private static bool KnownDescription(TeamTemplateRecordEvidence row) => row.Content.TryGetValue("description", out var hash) && hash.Known;
        private static string DescriptionEvidence(TeamTemplateRecordEvidence row) => row.Content.TryGetValue("description", out var hash) ? hash.Evidence : "Unavailable";
    }
}
#endif
