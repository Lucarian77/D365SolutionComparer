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
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Metadata.Query;
using Microsoft.Xrm.Sdk.Query;

namespace D365SolutionComparer.Services.Membership
{
    /// <summary>Explicit metadata-first Debug investigation; never supplies production keys or absence evidence.</summary>
    internal sealed class Type10276EvidenceCollector
    {
        internal const int BatchSize = 200, MaxIsolationGroups = 64;
        private static readonly string[] InternalNames = { "uniquename", "schemaname", "logicalname" };
        private static readonly string[] AuditText = { "name", "displayname", "uniquename", "schemaname", "logicalname", "version" };
        private static readonly string[] UnsafePayload = { "binary", "attachment", "thumbnail", "image", "media", "package", "base64", "encoded", "secret", "secure", "credential", "token", "document" };
        private static readonly HashSet<string> AuditReferences = new HashSet<string>(new[] { "organizationid", "ownerid", "owninguser", "owningteam",
            "owningbusinessunit", "createdby", "createdonbehalfby", "modifiedby", "modifiedonbehalfby", "solutionid", "supportingsolutionid",
            "rootsolutionid" }, StringComparer.OrdinalIgnoreCase);
        private static bool AuditReference(string field) => AuditReferences.Contains(field) || field.EndsWith("idunique", StringComparison.OrdinalIgnoreCase);

        internal AiSkillConfigEvidenceReport Capture(IOrganizationService sourceService, MembershipSnapshot source, string sourceVersion,
            IOrganizationService targetService, MembershipSnapshot target, string targetVersion, CancellationToken token, Action<string> progress = null)
        {
            if (source?.State != MembershipSnapshotState.Complete || target?.State != MembershipSnapshotState.Complete ||
                !StringComparer.OrdinalIgnoreCase.Equals(source.SolutionUniqueName, target.SolutionUniqueName))
                throw new ArgumentException("Completed snapshots of the same solution are required.");
            token.ThrowIfCancellationRequested();
            var report = new AiSkillConfigEvidenceReport { Source = Read(sourceService, source, sourceVersion, token, progress, true),
                Target = Read(targetService, target, targetVersion, token, progress, false) };
            report.Analyze(token); return report;
        }

        private static AiSkillConfigSideEvidence Read(IOrganizationService service, MembershipSnapshot snapshot, string version,
            CancellationToken token, Action<string> progress, bool allowAttributeMetadata)
        {
            if (service == null) throw new ArgumentNullException(nameof(service));
            token.ThrowIfCancellationRequested();
            var side = new AiSkillConfigSideEvidence { Snapshot = snapshot, Version = version };
            side.Raw.AddRange(snapshot.Components.Where(c => c.Record.ComponentType == 10276));
            var ids = side.Raw.Where(c => c.Record.ObjectId.HasValue && c.Record.ObjectId != Guid.Empty)
                .Select(c => c.Record.ObjectId.Value).Distinct().OrderBy(id => id).ToArray();
            foreach (var id in ids) side.Rows[id] = new AiSkillConfigRecordEvidence { ObjectId = id, Status = "Incomplete", Reason = "Backing entity/schema not verified" };
            if (ids.Length == 0) return side;

            if (!DiscoverBacking(service, side, token)) return side;
            var metadata = Schema(service, side.EntityName, side, token);
            if (metadata == null)
            { foreach (var row in side.Rows.Values) { row.Status = side.SchemaFailure; row.Reason = "Backing schema unavailable; no guessed columns"; } return side; }
            side.PrimaryId = metadata.PrimaryIdAttribute;
            var columns = new List<string>(); var hashes = new HashSet<string>(StringComparer.Ordinal);
            var excludedShadows = new HashSet<string>(StringComparer.Ordinal);
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
                // Exposed scope fields remain required evidence even when their type/readability is unusable.
                if (shadow) excludedShadows.Add(name);
                bool scope = !shadow && attribute.AttributeType != AttributeTypeCode.Lookup && new[] { "objecttypecode", "entitylogicalname", "entityname", "primaryentity" }.Contains(name);
                if (scope) side.ScopeFields.Add(name);
                bool hash = text && (attribute.AttributeType == AttributeTypeCode.Memo ||
                    !AuditText.Contains(name) && !scope);
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
            side.CandidateRole = "InternalSemanticIdentifierHypothesis";
            var referenceFields = metadata.Attributes.OfType<LookupAttributeMetadata>().Where(a => !AuditReference(a.LogicalName)).Select(a => a.LogicalName)
                .Concat((metadata.ManyToOneRelationships ?? new OneToManyRelationshipMetadata[0]).Where(r => r != null &&
                    r.ReferencingEntity == side.EntityName && !AuditReference(r.ReferencingAttribute)).Select(r => r.ReferencingAttribute))
                .Where(f => !excludedShadows.Contains(f)).Distinct(StringComparer.Ordinal).OrderBy(f => f, StringComparer.Ordinal).ToArray();
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
            }
            ResolveEntityScopes(service, side, backing, token);
            ResolveAttributes(service, side, token, allowAttributeMetadata);
            foreach (var row in side.Rows.Values.Where(r => r.Status == "Unique"))
            {
                FinalizeReferences(row);
                var value = side.CandidateField == null ? null : row.Get(side.CandidateField);
                if (row.CriticalComplete && row.ParentComplete && (row.TableStatus == "Verified" || row.TableStatus == "NotExposed") &&
                    !string.IsNullOrWhiteSpace(value) && !Guid.TryParse(value, out var ignored))
                    row.CandidateA = "aiskillconfig-candidate-a:" + Type31EvidenceCollector.Frame(side.EntityName, side.CandidateRole,
                        side.CandidateField, value.Trim(), row.TableKey ?? "NoTableScopeExposed", row.ParentKey);
                row.BlockingReason = !row.CriticalComplete ? "Critical field retrieval incomplete/faulted" :
                    !row.ParentComplete ? "Parent/reference identity " + row.ParentStatus :
                    row.TableStatus != "Verified" && row.TableStatus != "NotExposed" ? "Table scope " + row.TableStatus :
                    side.CandidateField == null ? "No independently semantic internal/unique/schema/logical identifier exposed" :
                    string.IsNullOrWhiteSpace(value) ? "Internal semantic identifier blank/unavailable" :
                    Guid.TryParse(value, out var localId) ? "Internal field is a GUID; portability unproven" : "None";
                // B is descriptive context only. It cannot supply or repair A or create a pair.
                var display = side.PrimaryNameField == null ? row.Get("name") : row.Get(side.PrimaryNameField);
                if (!string.IsNullOrWhiteSpace(display))
                    row.CandidateB = "aiskillconfig-candidate-b:" + Type31EvidenceCollector.Frame(side.EntityName, row.TableKey ?? "ScopeUnavailable", display.Trim());
            }
            return side;
        }

        private static bool DiscoverBacking(IOrganizationService service, AiSkillConfigSideEvidence side, CancellationToken token)
        {
            var definitions = side.Raw.Select(r => r.RegisteredDefinition).ToArray();
            var definition = definitions.FirstOrDefault(d => d != null);
            if (definition != null)
            {
                if (definitions.Any(d => d == null || d.ObjectTypeCode != 10276 || !ValidName(d.PrimaryEntityName) ||
                    !StringComparer.OrdinalIgnoreCase.Equals(d.Name.Replace(" ", ""), "AISkillConfig") ||
                    !StringComparer.OrdinalIgnoreCase.Equals(d.PrimaryEntityName, definition.PrimaryEntityName)))
                { side.Discovery.Add("Incomplete/conflicting completed registered definitions; no guessed backing query"); return false; }
                side.EntityName = definition.PrimaryEntityName;
                side.Discovery.Add("Reused completed registered definition: name=" + definition.Name + "; primaryentityname=" + side.EntityName);
                return true;
            }
            // Missing completed registered evidence triggers one scoped definition lookup; never guess a backing entity.
            try
            {
                var rows = new Dictionary<Guid, Entity>(); int page = 1; string cookie = null;
                while (true)
                {
                    token.ThrowIfCancellationRequested();
                    var query = new QueryExpression("solutioncomponentdefinition") { ColumnSet = new ColumnSet("objecttypecode", "name", "primaryentityname"),
                        PageInfo = new PagingInfo { Count = BatchSize, PageNumber = page, PagingCookie = cookie } };
                    query.Criteria.AddCondition("objecttypecode", ConditionOperator.Equal, 10276);
                    query.AddOrder("solutioncomponentdefinitionid", OrderType.Ascending);
                    side.Requests.Add("RetrieveMultiple solutioncomponentdefinition; objecttypecode=10276; columns=[objecttypecode,name,primaryentityname]; page=" + page);
                    var response = service.RetrieveMultiple(query); token.ThrowIfCancellationRequested();
                    if (response == null) { side.Discovery.Add("Incomplete registered-definition response"); return false; }
                    side.Pages.Add("solutioncomponentdefinition; rows=" + response.Entities.Count + "; page=" + page +
                        "; MoreRecords=" + response.MoreRecords + "; PagingCookieSupplied=" + !string.IsNullOrEmpty(response.PagingCookie));
                    int before = rows.Count;
                    foreach (var row in response.Entities)
                    {
                        var code = row?.GetAttributeValue<object>("objecttypecode");
                        if (row == null || row.LogicalName != "solutioncomponentdefinition" || row.Id == Guid.Empty ||
                            !(code is int || code is OptionSetValue) || (code is int ? (int)code : ((OptionSetValue)code).Value) != 10276 ||
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
                    if (string.IsNullOrWhiteSpace(name) || !StringComparer.OrdinalIgnoreCase.Equals(name.Replace(" ", ""), "AISkillConfig") || !ValidName(backing))
                    { side.Discovery.Add("Registered definition has incomplete/conflicting family/backing fields"); return false; }
                    side.EntityName = backing; side.Discovery.Add("Discovered registered definition: name=" + name + "; primaryentityname=" + backing); return true;
                }
                side.Discovery.Add("No registered solutioncomponentdefinition for Type 10276. This is not absence of AI Skill Config records.");
                side.Discovery.Add("Backing mapping unavailable: no consistent registered definition or verified completed backing correlation; no assumed scan");
                return false;
            }
            catch (OperationCanceledException) { throw; }
            catch (Exception error) { token.ThrowIfCancellationRequested(); var fault = SafeFault(error); side.Discovery.Add(fault);
                foreach (var row in side.Rows.Values) { row.Status = "Faulted"; row.Reason = "Registered-definition read faulted; " + fault; } return false; }
        }

        private static void ResolveEntityScopes(IOrganizationService service, AiSkillConfigSideEvidence side,
            Dictionary<Guid, LookupEvidence> backing, CancellationToken token)
        {
            var pending = new Dictionary<string, ScopeEvidence>(StringComparer.OrdinalIgnoreCase);
            var rowScopes = new Dictionary<Guid, List<ScopeEvidence>>();
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
                    { scopes.Add(new ScopeEvidence { Reason = field + ": missing/faulted/invalid entity scope" }); continue; }
                    if (!pending.TryGetValue(input, out var scope))
                    {
                        scope = new ScopeEvidence { Input = input, Number = numeric ? (int?)(raw is int ? (int)raw : int.Parse(text, CultureInfo.InvariantCulture)) : null,
                            Name = numeric ? null : text, Reason = "Metadata scope not yet verified" };
                        pending.Add(input, scope);
                        if (!numeric) ApplySnapshotScope(side.Snapshot, text, scope);
                    }
                    scopes.Add(scope);
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
                    scopes.Count == 0 ? "NotExposed" : scopes.All(s => s.Status == "Verified") && keys.Length == 1 ? "Verified" : "Incomplete";
                row.TableKey = row.TableStatus == "Verified" ? keys[0] : null;
                foreach (var scope in scopes) row.Context.Add("Entity scope: " + scope.Input + "; " + scope.Status + "; logicalName=" + scope.Key + "; " + scope.Reason);
                if (scopes.Count == 0) row.Context.Add("No table scope exposed; internal identifier remains a selected-snapshot hypothesis, not global uniqueness or absence proof");
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

        private static void ResolveScope(MembershipSnapshot snapshot, AiSkillConfigSideEvidence side, EntityMetadata metadata,
            string[] referenceFields, LookupEvidence found, AiSkillConfigRecordEvidence row)
        {
            foreach (var field in referenceFields)
            {
                var evidence = new AiSkillReferenceEvidence { Field = field, Reason = "Reference not yet verified" };
                row.References.Add(field, evidence);
                // Null means NotPresent only when the column was successfully retrieved. An unreadable/faulted
                // column cannot be assumed blank. Shadows never enter referenceFields.
                if (!found.Columns.Contains(field)) { evidence.Reason = "Reference column unavailable/faulted"; continue; }
                var raw = found.Row.GetAttributeValue<object>(field);
                if (raw == null) { evidence.Status = "NotPresent"; evidence.Reason = "Successfully read blank optional relationship"; continue; }
                var reference = raw as EntityReference;
                string entity = reference?.LogicalName; Guid? id = reference?.Id;
                var links = (metadata.ManyToOneRelationships ?? new OneToManyRelationshipMetadata[0]).Where(r => r != null &&
                    r.ReferencingEntity == side.EntityName && r.ReferencingAttribute == field).ToArray();
                if (links.Length > 1) { evidence.Status = "Ambiguous"; evidence.Reason = "Multiple metadata reference relationships"; continue; }
                if (raw is Guid && links.Length == 1 && PrimaryFor(links[0].ReferencedEntity) == links[0].ReferencedAttribute)
                { id = (Guid)raw; entity = links[0].ReferencedEntity; }
                var lookup = metadata.Attributes.OfType<LookupAttributeMetadata>().SingleOrDefault(a => a.LogicalName == field);
                bool targetValid = reference != null ? lookup?.Targets?.Contains(entity) == true &&
                    (links.Length == 0 || links[0].ReferencedEntity == entity && links[0].ReferencedAttribute == PrimaryFor(entity)) : raw is Guid && links.Length == 1;
                if (!targetValid || !id.HasValue || id == Guid.Empty || !RawTypeFor(entity).HasValue)
                { evidence.Reason = "Populated reference unavailable/invalid/unsupported; no GUID fallback"; continue; }
                evidence.Entity = entity; evidence.Id = id.Value;
                int type = RawTypeFor(entity).Value;
                var matches = snapshot.Components.Where(c => c.Record.ComponentType == type && c.Record.ObjectId == id).ToArray();
                if (matches.Length == 0 && type == 2)
                { evidence.Status = "PendingAttributeMetadata"; evidence.Reason = "No completed snapshot Column identity for populated attribute"; continue; }
                var identities = matches.Where(c => c.Status == IdentityResolutionStatus.Resolved && !string.IsNullOrWhiteSpace(c.ComparisonKey) &&
                    c.InventoryAbsencePolicy == InventoryAbsencePolicy.CompleteInventory).Select(c => Type31EvidenceCollector.Frame(c.SemanticKind, c.ComparisonKey))
                    .Distinct(StringComparer.OrdinalIgnoreCase).ToArray();
                bool unique = identities.Length == 1 && snapshot.Components.Where(c => c.Status == IdentityResolutionStatus.Resolved &&
                    StringComparer.OrdinalIgnoreCase.Equals(Type31EvidenceCollector.Frame(c.SemanticKind, c.ComparisonKey), identities[0]))
                    .Select(c => c.Record.ObjectId).Distinct().Count() == 1;
                if (matches.Any(c => c.Status == IdentityResolutionStatus.Ambiguous) || identities.Length > 1 || identities.Length == 1 && !unique)
                { evidence.Status = "Ambiguous"; evidence.Reason = "Ambiguous parent/reference identity in completed snapshot"; }
                else if (type == 2 && identities.Length == 0 && matches.All(c => c.Status != IdentityResolutionStatus.Ambiguous))
                { evidence.Status = "PendingAttributeMetadata"; evidence.Reason = "Completed snapshot has no verified Column key; independent scoped metadata required"; }
                else if (matches.Length == 0 || matches.Any(c => c.Status != IdentityResolutionStatus.Resolved || c.InventoryAbsencePolicy != InventoryAbsencePolicy.CompleteInventory) || !unique)
                    evidence.Reason = "Missing/incomplete verified reference identity in completed snapshot";
                else
                {
                    evidence.Status = "VerifiedSnapshot"; evidence.Key = identities[0]; evidence.PortableKey = matches[0].ComparisonKey;
                    evidence.Reason = "reused verified snapshot identity";
                    if (type == 1) row.TableKeys.Add(matches[0].ComparisonKey.Trim().ToLowerInvariant());
                }
            }
        }

        private static void ResolveAttributes(IOrganizationService service, AiSkillConfigSideEvidence side,
            CancellationToken token, bool allowMetadata)
        {
            var pending = new List<Tuple<AiSkillConfigRecordEvidence, AiSkillReferenceEvidence>>();
            foreach (var row in side.Rows.Values.Where(r => r.Status == "Unique"))
                foreach (var reference in row.References.Values.Where(r => r.Entity == "attribute"))
                {
                    token.ThrowIfCancellationRequested();
                    if (reference.Status == "VerifiedSnapshot" && !ColumnInTable(reference.PortableKey, row.TableKey, row.TableStatus))
                    { reference.Status = "Incomplete"; reference.Key = null; reference.Reason = "Snapshot Column scope conflicts with or lacks verified parent table"; }
                    if (reference.Status != "PendingAttributeMetadata") continue;
                    if (!allowMetadata || row.TableStatus != "Verified" || !LogicalName(row.TableKey))
                    { reference.Status = "Incomplete"; reference.Reason = !allowMetadata ? "Source-only scoped attribute lookup is disabled on Target" : "Verified parent table required for scoped attribute lookup"; continue; }
                    pending.Add(Tuple.Create(row, reference));
                }
            // The same metadata ID cannot be correlated to contradictory selected table scopes.
            foreach (var group in pending.GroupBy(p => p.Item2.Id).Where(g => g.Select(p => p.Item1.TableKey).Distinct(StringComparer.OrdinalIgnoreCase).Count() > 1))
                foreach (var item in group) { item.Item2.Status = "Ambiguous"; item.Item2.Reason = "Attribute ID has conflicting selected parent table scopes"; }
            foreach (var table in pending.Where(p => p.Item2.Status == "PendingAttributeMetadata").GroupBy(p => p.Item1.TableKey, StringComparer.OrdinalIgnoreCase).OrderBy(g => g.Key, StringComparer.Ordinal))
            {
                var ids = table.Select(p => p.Item2.Id).Distinct().OrderBy(id => id).ToArray();
                for (int offset = 0; offset < ids.Length; offset += BatchSize)
                {
                    token.ThrowIfCancellationRequested(); var batch = ids.Skip(offset).Take(BatchSize).ToArray();
                    var references = table.Where(p => batch.Contains(p.Item2.Id)).Select(p => p.Item2).ToArray();
                    // Child filters in RetrieveMetadataChanges are not supported by the existing SDK path.
                    // ExecuteMultiple batches only exact selected IDs, with explicit verified table scope.
                    // No whole-table Attributes inventory, per-ID network loop, or Target lookup is issued.
                    var request = new ExecuteMultipleRequest { Settings = new ExecuteMultipleSettings { ContinueOnError = true, ReturnResponses = true },
                        Requests = new OrganizationRequestCollection() };
                    foreach (var id in batch) request.Requests.Add(new RetrieveAttributeRequest {
                        EntityLogicalName = table.Key, MetadataId = id, RetrieveAsIfPublished = false });
                    side.AttributeMetadataBatches++; side.AttributeMetadataSubrequests += batch.Length;
                    side.Requests.Add("Execute ExecuteMultiple(RetrieveAttribute); table=" + table.Key + "; selectedMetadataIds=[" +
                        string.Join(",", batch) + "]; subrequests=" + batch.Length + "; RetrieveAsIfPublished=False");
                    try
                    {
                        var response = service.Execute(request) as ExecuteMultipleResponse; token.ThrowIfCancellationRequested();
                        if (response == null || response.Responses == null || response.Responses.Any(r => r == null || r.RequestIndex < 0 || r.RequestIndex >= batch.Length))
                        { foreach (var reference in references) { reference.Status = "Incomplete"; reference.Reason = "Incomplete/foreign batch metadata response"; } continue; }
                        for (int index = 0; index < batch.Length; index++)
                        {
                            var items = response.Responses.Where(r => r.RequestIndex == index).ToArray();
                            string status = "Incomplete", key = null, reason = "No unique selected attribute metadata response";
                            if (items.Length > 1) { status = "Ambiguous"; reason = "Multiple metadata responses for one selected attribute"; }
                            else if (items.Length == 1 && items[0].Fault != null)
                            { status = "Faulted"; reason = SafeFault(new FaultException<OrganizationServiceFault>(items[0].Fault)); }
                            else if (items.Length == 1)
                            {
                                var attribute = (items[0].Response as RetrieveAttributeResponse)?.AttributeMetadata;
                                if (attribute != null && attribute.MetadataId == batch[index] && LogicalName(attribute.LogicalName) &&
                                    StringComparer.OrdinalIgnoreCase.Equals(attribute.EntityLogicalName, table.Key))
                                { key = table.Key.ToLowerInvariant() + "." + attribute.LogicalName.Trim().ToLowerInvariant();
                                    var competing = side.Snapshot.Components.Where(c => c.Record.ComponentType == 2 &&
                                        StringComparer.OrdinalIgnoreCase.Equals(c.ComparisonKey, key)).ToArray();
                                    if (competing.Any(c => c.Record.ObjectId != batch[index] || c.Status != IdentityResolutionStatus.Resolved || c.SemanticKind != ComponentSemanticKinds.Column))
                                    { status = "Ambiguous"; key = null; reason = "Scoped metadata Column identity conflicts with completed snapshot correlation"; }
                                    else { status = "VerifiedScopedMetadata"; reason = "Exact selected MetadataId and independently verified table/attribute logical scope"; } }
                                else reason = "Missing/conflicting selected attribute MetadataId or table/logical name";
                            }
                            foreach (var reference in references.Where(r => r.Id == batch[index]))
                            { reference.Status = status; reference.PortableKey = key; reference.Key = key == null ? null : Type31EvidenceCollector.Frame(ComponentSemanticKinds.Column, key); reference.Reason = reason; }
                        }
                    }
                    catch (OperationCanceledException) { throw; }
                    catch (Exception error) { token.ThrowIfCancellationRequested(); foreach (var reference in references) { reference.Status = "Faulted"; reference.Reason = SafeFault(error); } }
                }
            }
            // Distinct backing attribute IDs producing one canonical Column identity remain ambiguous.
            var resolved = side.Rows.Values.SelectMany(r => r.References.Values).Where(r => r.Entity == "attribute" && r.Key != null &&
                (r.Status == "VerifiedSnapshot" || r.Status == "VerifiedScopedMetadata"));
            foreach (var collision in resolved.GroupBy(r => r.Key, StringComparer.OrdinalIgnoreCase).Where(g => g.Select(r => r.Id).Distinct().Count() > 1))
                foreach (var reference in collision) { reference.Status = "Ambiguous"; reference.Key = null; reference.Reason = "Distinct attribute IDs share one canonical Column identity"; }
        }

        private static bool ColumnInTable(string key, string table, string tableStatus) => tableStatus == "Verified" &&
            !string.IsNullOrWhiteSpace(key) && key.StartsWith(table + ".", StringComparison.OrdinalIgnoreCase) && LogicalName(key.Substring(table.Length + 1));

        private static void FinalizeReferences(AiSkillConfigRecordEvidence row)
        {
            var populated = row.References.Values.Where(r => r.Status != "NotPresent").ToArray();
            row.ParentComplete = populated.All(r => r.Status == "VerifiedSnapshot" || r.Status == "VerifiedScopedMetadata");
            row.ParentStatus = populated.Any(r => r.Status == "Ambiguous") ? "Ambiguous" : !row.ParentComplete ? "Incomplete" :
                populated.Length == 0 ? "NoPopulatedParentReference" : populated.Any(r => r.Status == "VerifiedScopedMetadata") ? "VerifiedScopedMetadata" : "VerifiedSnapshotScope";
            row.ParentKey = !row.ParentComplete ? null : populated.Length == 0 ? "NoPopulatedParentReference" :
                Type31EvidenceCollector.Frame(populated.Select(r => Type31EvidenceCollector.Frame(r.Field, r.Key)).OrderBy(k => k, StringComparer.OrdinalIgnoreCase).ToArray());
            foreach (var reference in row.References.Values) row.Context.Add(reference.Field + ": " + reference.Status +
                "; PortableKey=" + reference.PortableKey + "; " + reference.Reason);
            row.Context.Add("Only populated verified references contribute to Candidate A. Blank successfully read relationships are NotPresent; shadow fields and audit IDs never supply scope.");
        }

        private static int? RawTypeFor(string entity)
        {
            switch (entity) { case "entity": return 1; case "attribute": return 2; case "relationship": return 10;
                case "webresource": return 61; case "appmodule": return 80; case "savedquery": return 26; case "workflow": return 29;
                case "savedqueryvisualization": return 59; case "systemform": return 60; case "sitemap": return 62;
                case "pluginassembly": return 91; case "sdkmessageprocessingstep": return 92; case "environmentvariabledefinition": return 380; default: return null; }
        }
        private static string PrimaryFor(string entity) => RawTypeFor(entity).HasValue ? entity == "systemform" ? "formid" : entity + "id" : null;
        private static EntityMetadata Schema(IOrganizationService service, string entity, AiSkillConfigSideEvidence side, CancellationToken token)
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
            string[] critical, string[] optional, Guid[] ids, AiSkillConfigSideEvidence side, CancellationToken token, Action<string> progress)
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
            Dictionary<Guid, LookupEvidence> current, bool critical, AiSkillConfigSideEvidence side, CancellationToken token,
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
                        else { side.RetrievalDiagnostics.Add("Optional enrichment correlation unavailable for selected ID=" + id + "; no optional values merged; primary evidence retained"); continue; }
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
            Guid[] ids, AiSkillConfigSideEvidence side, CancellationToken token, Action<string> progress)
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

    internal sealed class AiSkillConfigSideEvidence
    {
        internal MembershipSnapshot Snapshot;
        internal int AttributeMetadataBatches, AttributeMetadataSubrequests;
        internal string Version, EntityName, PrimaryId, CandidateField, PrimaryNameField, CandidateRole, SchemaFailure;
        internal readonly List<string> Discovery = new List<string>(), ScopeFields = new List<string>();
        internal readonly List<ComponentIdentity> Raw = new List<ComponentIdentity>();
        internal readonly SortedDictionary<Guid, AiSkillConfigRecordEvidence> Rows = new SortedDictionary<Guid, AiSkillConfigRecordEvidence>();
        internal readonly List<string> Schema = new List<string>(), Relationships = new List<string>(), Requests = new List<string>(), Pages = new List<string>(), RetrievalDiagnostics = new List<string>();
    }
    internal sealed class AiSkillReferenceEvidence
    {
        internal string Field, Entity, Status = "Incomplete", Key, PortableKey, Reason; internal Guid Id;
    }
    internal sealed class AiSkillConfigRecordEvidence
    {
        internal Guid ObjectId; internal Guid? PrimaryId; internal bool? Managed;
        internal string Status, Reason, CandidateField, CandidateA, CandidateB, BlockingReason;
        internal bool CriticalComplete, DuplicateA, DuplicateB, ParentComplete;
        internal string ParentKey, ParentStatus, TableKey, TableStatus;
        internal readonly SortedDictionary<string, AiSkillReferenceEvidence> References = new SortedDictionary<string, AiSkillReferenceEvidence>(StringComparer.Ordinal);
        internal readonly SortedSet<string> TableKeys = new SortedSet<string>(StringComparer.OrdinalIgnoreCase);
        internal readonly List<string> Context = new List<string>();
        internal readonly SortedDictionary<string, string> Fields = new SortedDictionary<string, string>(StringComparer.Ordinal);
        internal readonly SortedDictionary<string, Type31ContentFingerprint> Content = new SortedDictionary<string, Type31ContentFingerprint>(StringComparer.Ordinal);
        internal readonly List<string> RuntimeColumns = new List<string>();
        internal string Get(string field) => Fields.TryGetValue(field, out var value) ? value : null;
        internal string Evidence(string field) => Content.TryGetValue(field, out var hash) ? hash.Evidence : Get(field);
    }
    internal sealed class AiSkillConfigPairEvidence
    {
        internal AiSkillConfigRecordEvidence Source, Target; internal string Outcome;
        internal readonly HashSet<string> Categories = new HashSet<string>(StringComparer.Ordinal);
    }
    internal sealed class AiSkillConfigEvidenceReport
    {
        internal AiSkillConfigSideEvidence Source, Target;
        internal readonly List<AiSkillConfigPairEvidence> Pairs = new List<AiSkillConfigPairEvidence>();
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
                    foreach (var row in left) Pairs.Add(new AiSkillConfigPairEvidence { Source = row, Outcome = "Ambiguous" });
                    foreach (var row in right) Pairs.Add(new AiSkillConfigPairEvidence { Target = row, Outcome = "Ambiguous" });
                }
                else Pairs.Add(new AiSkillConfigPairEvidence { Source = left.SingleOrDefault(), Target = right.SingleOrDefault(),
                    Outcome = left.Length == 1 && right.Length == 1 ? "SemanticPair" : "OneSidedEvidence" });
            }
            foreach (var side in new[] { Source, Target })
            {
                foreach (var row in side.Rows.Values.Where(r => r.CandidateA == null))
                    Pairs.Add(new AiSkillConfigPairEvidence { Source = side == Source ? row : null, Target = side == Target ? row : null,
                        Outcome = row.Status == "Duplicate" || row.ParentStatus == "Ambiguous" || row.TableStatus == "Ambiguous" ? "Ambiguous" : "Incomplete" });
                foreach (var raw in side.Raw.Where(r => !r.Record.ObjectId.HasValue || r.Record.ObjectId == Guid.Empty))
                    Pairs.Add(new AiSkillConfigPairEvidence { Outcome = "Incomplete" });
            }
            foreach (var pair in Pairs)
            {
                token.ThrowIfCancellationRequested(); pair.Categories.Add(pair.Outcome);
                if (pair.Outcome != "SemanticPair") continue;
                var left = pair.Source; var right = pair.Target;
                pair.Categories.Add(left.PrimaryId == right.PrimaryId ? "SamePrimaryId" : "DifferentPrimaryId");
                var uniqueFields = left.Fields.Keys.Union(right.Fields.Keys).Where(f => f.EndsWith("idunique", StringComparison.OrdinalIgnoreCase) &&
                    Guid.TryParse(left.Get(f), out var sourceId) && Guid.TryParse(right.Get(f), out var targetId)).ToArray();
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
            var text = new StringBuilder(); text.AppendLine("TYPE 10276 AI SKILL CONFIG EVIDENCE - DEBUG ONLY");
            text.AppendLine("Evidence only. Type 10276 remains Unsupported / Indeterminate; no production key, definition contract or absence proof.");
            text.AppendLine("Candidate A is an internal semantic identifier hypothesis, scoped by any verified exposed references; selected-snapshot uniqueness is not lifecycle proof. Candidate B, display names, GUIDs and hashes never repair A.");
            text.AppendLine("\nRAW TYPE 10276 MEMBERSHIP"); EachSide(text, (s, label) => {
                Line(text, label, s.Snapshot.Environment.DisplayName, s.Snapshot.SolutionUniqueName, "version=" + s.Version,
                    "snapshotUtc=" + s.Snapshot.CapturedAt.ToUniversalTime().ToString("O"), "raw=" + s.Raw.Count,
                    "distinctNonblankIds=" + s.Rows.Count, "blankIds=" + s.Raw.Count(r => !r.Record.ObjectId.HasValue || r.Record.ObjectId == Guid.Empty));
                if (s.Raw.Count == 0) Line(text, label, "No selected Type 10276 membership references on this side; no backing lookup performed; this does not prove backing-record absence.");
                else if (s.Rows.Count == 0) Line(text, label, "Selected membership references have no usable nonblank object IDs; no backing lookup performed; evidence remains incomplete.");
                foreach (var raw in s.Raw) Line(text, label, "solutioncomponentid=" + raw.Record.SolutionComponentId,
                    "objectid=" + raw.Record.ObjectId, "productionStatus=" + raw.Status, raw.Diagnostic);
            });
            text.AppendLine("\nREGISTERED-DEFINITION / BACKING-ENTITY DISCOVERY"); EachSide(text, (s, label) => {
                Line(text, label, "backingEntity=" + (s.EntityName ?? "NotEstablished"), "IsAiskillconfigObserved=" + StringComparer.OrdinalIgnoreCase.Equals(s.EntityName, "aiskillconfig"));
                foreach (var item in s.Discovery) Line(text, label, item);
            });
            text.AppendLine("\nBACKING-RECORD CORRELATION"); EachSide(text, (s, label) => {
                foreach (var row in s.Rows.Values) Line(text, label, "objectid=" + row.ObjectId, "primaryAttribute=" + s.PrimaryId,
                    "backingPrimaryId=" + row.PrimaryId, "ExactLocalCorrelation=" + (row.PrimaryId == row.ObjectId), row.Status, row.Reason);
            });
            text.AppendLine("\nREADABLE / UNAVAILABLE SCHEMA"); EachSide(text, (s, label) => {
                foreach (var item in s.Schema) Line(text, label, item);
            });
            text.AppendLine("\nRUNTIME-READABLE / FAULTED COLUMNS"); EachSide(text, (s, label) => {
                foreach (var item in s.RetrievalDiagnostics.Concat(s.Pages)) Line(text, label, item);
                foreach (var row in s.Rows.Values) Line(text, label, row.ObjectId, "RuntimeReadable=[" + string.Join(",", row.RuntimeColumns) + "]");
            });
            text.AppendLine("\nPARENT / REFERENCE RELATIONSHIP DISCOVERY"); EachSide(text, (s, label) => {
                foreach (var item in s.Relationships) Line(text, label, item);
            });
            text.AppendLine("\nPORTABLE PARENT / REFERENCE RESOLUTION"); EachSide(text, (s, label) => {
                foreach (var row in s.Rows.Values) { Line(text, label, row.ObjectId, "ParentStatus=" + row.ParentStatus,
                    "ParentIdentityComplete=" + row.ParentComplete, "ParentKey=" + row.ParentKey, "TableStatus=" + row.TableStatus, "TableKey=" + row.TableKey);
                    foreach (var item in row.Context) Line(text, label, row.ObjectId, item); }
            });
            text.AppendLine("\nCANDIDATE A / B ANALYSIS"); EachSide(text, (s, label) => {
                foreach (var row in s.Rows.Values) Line(text, label, row.ObjectId, "CandidateRole=" + s.CandidateRole,
                    "CandidateField=" + row.CandidateField, "CandidateA=" + (row.CandidateA ?? "Incomplete"), "CandidateB=" + (row.CandidateB ?? "NotAvailable"),
                    "CompleteA=" + (row.CandidateA != null && !row.DuplicateA), "BlockingReason=" + (row.DuplicateA ? "Candidate A collision" : row.BlockingReason ?? row.Reason));
            });
            text.AppendLine("\nDUPLICATE / COLLISION ANALYSIS"); EachSide(text, (s, label) => {
                foreach (var group in s.Raw.Where(r => r.Record.ObjectId.HasValue && r.Record.ObjectId != Guid.Empty).GroupBy(r => r.Record.ObjectId).Where(g => g.Count() > 1))
                    Line(text, label, "RepeatedRawObjectId=" + group.Key, "references=" + group.Count(), "membership repetition, not distinct backing identity");
                Line(text, label, "CandidateACollisionGroups=" + s.Rows.Values.Where(r => r.DuplicateA).Select(r => r.CandidateA).Distinct(StringComparer.OrdinalIgnoreCase).Count(),
                    "CandidateBCollisionGroups=" + s.Rows.Values.Where(r => r.DuplicateB).Select(r => r.CandidateB).Distinct(StringComparer.OrdinalIgnoreCase).Count());
            });
            text.AppendLine("\nSELECTED SOURCE / TARGET RECORD FIELD COMPARISON"); EachSide(text, (s, label) => {
                foreach (var row in s.Rows.Values) foreach (var field in row.Fields.Keys.Union(row.Content.Keys).OrderBy(f => f, StringComparer.Ordinal))
                    Line(text, label, row.ObjectId, field, row.Evidence(field) ?? "Unavailable");
            });
            foreach (var pair in Pairs.Where(p => p.Outcome == "SemanticPair")) {
                Line(text, "Observed semantic pair", pair.Source.ObjectId, pair.Target.ObjectId);
                foreach (var field in pair.Source.Fields.Keys.Union(pair.Target.Fields.Keys).Union(pair.Source.Content.Keys).Union(pair.Target.Content.Keys).OrderBy(f => f, StringComparer.Ordinal)) {
                    var left = pair.Source.Evidence(field); var right = pair.Target.Evidence(field);
                    Line(text, field, "Source=" + (left ?? "Unavailable"), "Target=" + (right ?? "Unavailable"),
                        left == null || right == null ? "Unknown" : StringComparer.Ordinal.Equals(left, right) ? "EqualObserved" : "DifferentObserved");
                }
            }
            text.AppendLine("\nONE-SIDED LIFECYCLE EVIDENCE");
            Line(text, "Selected Source semantic candidates=" + Source.Rows.Values.Count(r => r.CandidateA != null && !r.DuplicateA),
                "Selected Target semantic candidates=" + Target.Rows.Values.Count(r => r.CandidateA != null && !r.DuplicateA),
                "Zero selected Source membership references=" + (Source.Raw.Count == 0), "Zero selected Target membership references=" + (Target.Raw.Count == 0));
            foreach (var category in new[] { "SemanticPair", "SamePrimaryId", "DifferentPrimaryId", "SameUniqueId", "DifferentUniqueId", "SameDefinitionEvidence", "DifferentDefinitionEvidence", "ManagedTransition", "UnmanagedToManaged", "OneSidedEvidence", "Ambiguous", "Incomplete" })
                Line(text, category, Pairs.Count(p => p.Categories.Contains(category)));
            text.AppendLine("One-sided selected membership evidence is not backing-record absence proof; lifecycle portability is not inferred.");
            text.AppendLine("\nCROSS-ENVIRONMENT PAIRING STATUS");
            Line(text, "Actual cross-environment semantic pairs=" + Pairs.Count(p => p.Outcome == "SemanticPair"));
            foreach (var pair in Pairs) Line(text, pair.Outcome, "Source=" + pair.Source?.ObjectId, "Target=" + pair.Target?.ObjectId);
            text.AppendLine("\nPORTABILITY ASSESSMENT");
            text.AppendLine("No production identity or absence inference is supplied. Zero selected references do not establish backing absence. Live schema, candidate collisions, transport and lifecycle portability require review; no GUID, display-name or hash fallback.");
            text.AppendLine("\nEXACT REQUEST LEDGER"); EachSide(text, (s, label) => {
                foreach (var request in s.Requests) Line(text, label, request);
                foreach (var group in s.Requests.GroupBy(r => r.StartsWith("Execute RetrieveEntity(", StringComparison.Ordinal) ? r.Split(',')[0] + ")" :
                    r.StartsWith("Execute RetrieveMetadataChanges", StringComparison.Ordinal) ? "Execute RetrieveMetadataChanges" : r.Split(';')[0]))
                    Line(text, label, group.Key, "reads=" + group.Count());
                Line(text, label, "TotalReads=" + s.Requests.Count, "ReadServiceCalls=" + s.Requests.Count,
                    "AttributeMetadataSubrequests=" + s.AttributeMetadataSubrequests,
                    "TotalReadOperations=" + (s.Requests.Count - s.AttributeMetadataBatches + s.AttributeMetadataSubrequests), "AdditionalWhoAmI=0", "Writes=0", "NormalMembershipEvidenceRequests=0");
            }); return text.ToString();
        }
        private void EachSide(StringBuilder text, Action<AiSkillConfigSideEvidence, string> action) { action(Source, "Source"); action(Target, "Target"); }
    }
}
#endif
