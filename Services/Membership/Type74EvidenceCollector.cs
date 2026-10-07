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
    /// <summary>Explicit metadata-first Debug investigation; never supplies production keys or absence evidence.</summary>
    internal sealed class Type74EvidenceCollector
    {
        internal const int BatchSize = 200, MaxIsolationGroups = 64;
        private static readonly string[] InternalNames = { "uniquename", "schemaname", "logicalname" };
        private static readonly string[] AuditText = { "name", "displayname", "uniquename", "schemaname", "logicalname", "version" };
        private static readonly string[] UnsafePayload = { "binary", "attachment", "thumbnail", "image", "media", "package", "base64", "encoded", "secret", "secure", "credential", "token", "document" };
        private static readonly HashSet<string> AuditReferences = new HashSet<string>(new[] { "organizationid", "ownerid", "owninguser", "owningteam",
            "owningbusinessunit", "createdby", "createdonbehalfby", "modifiedby", "modifiedonbehalfby" }, StringComparer.OrdinalIgnoreCase);

        internal MaskingRuleEvidenceReport Capture(IOrganizationService sourceService, MembershipSnapshot source, string sourceVersion,
            IOrganizationService targetService, MembershipSnapshot target, string targetVersion, CancellationToken token, Action<string> progress = null)
        {
            if (source?.State != MembershipSnapshotState.Complete || target?.State != MembershipSnapshotState.Complete ||
                !StringComparer.OrdinalIgnoreCase.Equals(source.SolutionUniqueName, target.SolutionUniqueName))
                throw new ArgumentException("Completed snapshots of the same solution are required.");
            token.ThrowIfCancellationRequested();
            var report = new MaskingRuleEvidenceReport { Source = Read(sourceService, source, sourceVersion, token, progress),
                Target = Read(targetService, target, targetVersion, token, progress) };
            report.Analyze(token); return report;
        }

        private static MaskingRuleSideEvidence Read(IOrganizationService service, MembershipSnapshot snapshot, string version,
            CancellationToken token, Action<string> progress)
        {
            if (service == null) throw new ArgumentNullException(nameof(service));
            token.ThrowIfCancellationRequested();
            var side = new MaskingRuleSideEvidence { Snapshot = snapshot, Version = version };
            side.Raw.AddRange(snapshot.Components.Where(c => c.Record.ComponentType == 74));
            var ids = side.Raw.Where(c => c.Record.ObjectId.HasValue && c.Record.ObjectId != Guid.Empty)
                .Select(c => c.Record.ObjectId.Value).Distinct().OrderBy(id => id).ToArray();
            foreach (var id in ids) side.Rows[id] = new MaskingRuleRecordEvidence { ObjectId = id, Status = "Incomplete", Reason = "Backing entity/schema not verified" };
            if (ids.Length == 0) return side;

            // The family label alone is not a backing-table mapping. Require complete, consistent registered evidence.
            var definitions = side.Raw.Select(c => c.RegisteredDefinition).ToArray();
            var definition = definitions.FirstOrDefault(d => d != null);
            if (definition == null || definitions.Any(d => d == null || d.ObjectTypeCode != 74 ||
                !StringComparer.OrdinalIgnoreCase.Equals(d.Name, "MaskingRule") || !ValidName(d.PrimaryEntityName) ||
                !StringComparer.OrdinalIgnoreCase.Equals(d.PrimaryEntityName, definition.PrimaryEntityName)))
            {
                side.Schema.Add("Type 74 registered MaskingRule backing entity is missing/conflicting; no assumed table or query");
                foreach (var row in side.Rows.Values) row.Reason = "Registered component definition does not establish one backing entity";
                return side;
            }
            side.EntityName = definition.PrimaryEntityName;
            side.Schema.Add("Backing entity discovered from completed registered definition: Name=" + definition.Name + "; PrimaryEntity=" + side.EntityName);
            var metadata = Schema(service, side.EntityName, side, token);
            if (metadata == null)
            { foreach (var row in side.Rows.Values) { row.Status = side.SchemaFailure; row.Reason = "Backing schema unavailable; no guessed columns"; } return side; }
            side.PrimaryId = metadata.PrimaryIdAttribute;
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
                    attribute.AttributeType == AttributeTypeCode.Integer || attribute.AttributeType == AttributeTypeCode.BigInt ||
                    attribute.AttributeType == AttributeTypeCode.Boolean || attribute.AttributeType == AttributeTypeCode.DateTime;
                bool readable = attribute.IsValidForRead == true && shape && !shadow && !payload;
                bool hash = text && (attribute.AttributeType == AttributeTypeCode.Memo || !AuditText.Contains(name));
                side.Schema.Add(side.EntityName + "." + name + "; type=" + attribute.AttributeType + "; readable=" + attribute.IsValidForRead +
                    "; capture=" + (shadow ? "ExcludedLookupShadow" : payload ? "ExcludedPayload" : readable ? hash ? "HashOnly" : "Audit" : "UnavailableOrNotQueried"));
                if (readable) { columns.Add(name); if (hash) hashes.Add(name); }
            }
            if (!columns.Contains(side.PrimaryId) || metadata.Attributes.Single(a => a.LogicalName == side.PrimaryId).AttributeType != AttributeTypeCode.Uniqueidentifier)
            { foreach (var row in side.Rows.Values) row.Reason = "Readable GUID primary key not verified"; return side; }
            side.CandidateField = InternalNames.FirstOrDefault(field => columns.Contains(field) && !hashes.Contains(field) &&
                metadata.Attributes.Single(a => a.LogicalName == field).AttributeType == AttributeTypeCode.String);
            var referenceFields = metadata.Attributes.OfType<LookupAttributeMetadata>().Where(a => !AuditReferences.Contains(a.LogicalName)).Select(a => a.LogicalName)
                .Concat((metadata.ManyToOneRelationships ?? new OneToManyRelationshipMetadata[0]).Where(r => r != null &&
                    r.ReferencingEntity == side.EntityName && !AuditReferences.Contains(r.ReferencingAttribute)).Select(r => r.ReferencingAttribute))
                .Distinct(StringComparer.Ordinal).OrderBy(f => f, StringComparer.Ordinal).ToArray();
            var critical = new[] { side.PrimaryId }.Concat(side.CandidateField == null ? new string[0] : new[] { side.CandidateField })
                .Concat(referenceFields.Where(columns.Contains)).Distinct(StringComparer.Ordinal).ToArray();
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
                    else row.Fields[field] = found.Columns.Contains(field) ? Format(found.Row.GetAttributeValue<object>(field)) : null;
                row.Managed = found.Row.GetAttributeValue<object>("ismanaged") is bool ? (bool?)found.Row.GetAttributeValue<bool>("ismanaged") : null;
                ResolveScope(snapshot, side, metadata, referenceFields, found, row);
                string value = side.CandidateField == null ? null : row.Get(side.CandidateField);
                row.CandidateField = side.CandidateField;
                if (found.CriticalComplete && row.ParentComplete && !string.IsNullOrWhiteSpace(value) && !Guid.TryParse(value, out var ignored))
                    row.CandidateA = "maskingrule-candidate-a:" + Type31EvidenceCollector.Frame(side.EntityName, side.CandidateField, value.Trim(), row.ParentKey);
                // Name is descriptive B only. It never repairs missing A, missing scope or collisions.
                if (!string.IsNullOrWhiteSpace(row.Get("name"))) row.CandidateB = "maskingrule-candidate-b:" +
                    Type31EvidenceCollector.Frame(side.EntityName, row.Get("name"), row.ParentKey ?? "ScopeIncomplete");
            }
            CaptureAttributeScope(service, side, metadata, token, progress);
            return side;
        }

        private static void ResolveScope(MembershipSnapshot snapshot, MaskingRuleSideEvidence side, EntityMetadata metadata,
            string[] referenceFields, LookupEvidence found, MaskingRuleRecordEvidence row, string referencingEntity = null,
            ISet<string> attributeKeys = null, ISet<string> tableKeys = null)
        {
            var keys = new SortedSet<string>(StringComparer.OrdinalIgnoreCase); var reasons = new List<string>(); bool ambiguous = false;
            foreach (var field in referenceFields)
            {
                var raw = found.Row.GetAttributeValue<object>(field); var reference = raw as EntityReference;
                string entity = reference?.LogicalName; Guid? id = reference?.Id;
                var links = (metadata.ManyToOneRelationships ?? new OneToManyRelationshipMetadata[0]).Where(r => r != null &&
                    r.ReferencingEntity == (referencingEntity ?? side.EntityName) && r.ReferencingAttribute == field).ToArray();
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
                    if (type == 2) attributeKeys?.Add(matches[0].ComparisonKey.Trim().ToLowerInvariant());
                    if (type == 1) tableKeys?.Add(matches[0].ComparisonKey.Trim().ToLowerInvariant());
                }
            }
            row.ParentComplete = reasons.Count == 0;
            row.ParentStatus = ambiguous ? "Ambiguous" : row.ParentComplete ? referenceFields.Length == 0 ? "NoParentRelationshipExposed" : "VerifiedSnapshotScope" : "Incomplete";
            row.ParentKey = row.ParentComplete ? referenceFields.Length == 0 ? "NoParentRelationshipExposed" : Type31EvidenceCollector.Frame(keys.ToArray()) : null;
            row.Context.AddRange(reasons);
            row.Context.Add("Scope interpretation is diagnostic only. Unexposed relationships are not proof of global uniqueness or lifecycle portability.");
        }

        private static void CaptureAttributeScope(IOrganizationService service, MaskingRuleSideEvidence side, EntityMetadata ruleMetadata,
            CancellationToken token, Action<string> progress)
        {
            side.AttributeRequestStart = side.Requests.Count;
            var rules = side.Rows.Values.Where(r => r.Status == "Unique").ToArray();
            if (rules.Length == 0) return;
            var links = (ruleMetadata.OneToManyRelationships ?? new OneToManyRelationshipMetadata[0]).Where(r => r != null &&
                r.ReferencingEntity == "attributemaskingrule" && r.ReferencedEntity == side.EntityName && r.ReferencedAttribute == side.PrimaryId).ToArray();
            if (links.Length == 0) { side.AttributeSchema.Add("No incoming AttributeMaskingRule relationship exposed; no assumed query"); return; }
            foreach (var row in rules) { row.CandidateA = null; row.AttributeScopeStatus = "Incomplete"; }
            if (links.Length != 1 || !ValidName(links[0].ReferencingAttribute))
            { SetAttributeScope(rules, "Ambiguous", "Conflicting incoming AttributeMaskingRule metadata relationships"); return; }
            var link = links[0]; int schemaStart = side.Schema.Count;
            var metadata = Schema(service, link.ReferencingEntity, side, token);
            side.AttributeSchema.AddRange(side.Schema.Skip(schemaStart));
            if (metadata == null) { SetAttributeScope(rules, side.SchemaFailure, "AttributeMaskingRule metadata unavailable; no guessed fields"); return; }
            var primary = metadata.Attributes.SingleOrDefault(a => a.LogicalName == metadata.PrimaryIdAttribute);
            var foreign = metadata.Attributes.SingleOrDefault(a => a.LogicalName == link.ReferencingAttribute);
            var reverse = (metadata.ManyToOneRelationships ?? new OneToManyRelationshipMetadata[0]).Where(r => r != null &&
                r.ReferencingEntity == metadata.LogicalName && r.ReferencingAttribute == link.ReferencingAttribute).ToArray();
            bool foreignValid = foreign?.IsValidForRead == true && (foreign.AttributeType == AttributeTypeCode.Uniqueidentifier ||
                foreign is LookupAttributeMetadata && ((LookupAttributeMetadata)foreign).Targets?.Contains(side.EntityName) == true) &&
                reverse.Length == 1 && reverse[0].ReferencedEntity == side.EntityName && reverse[0].ReferencedAttribute == side.PrimaryId;
            if (primary?.IsValidForRead != true || primary.AttributeType != AttributeTypeCode.Uniqueidentifier || !foreignValid)
            { SetAttributeScope(rules, "Incomplete", "Readable primary ID and exact MaskingRule relationship not verified on related metadata"); return; }
            side.AttributeEntity = metadata.LogicalName; side.AttributeForeignKey = link.ReferencingAttribute;
            var referenceFields = metadata.Attributes.OfType<LookupAttributeMetadata>().Where(a => a.LogicalName != link.ReferencingAttribute && !AuditReferences.Contains(a.LogicalName))
                .Select(a => a.LogicalName).Concat((metadata.ManyToOneRelationships ?? new OneToManyRelationshipMetadata[0]).Where(r => r != null &&
                    r.ReferencingEntity == metadata.LogicalName && r.ReferencingAttribute != link.ReferencingAttribute && !AuditReferences.Contains(r.ReferencingAttribute))
                    .Select(r => r.ReferencingAttribute)).Distinct(StringComparer.Ordinal).OrderBy(f => f, StringComparer.Ordinal).ToArray();
            var columns = new List<string>(); var hashes = new HashSet<string>(StringComparer.Ordinal);
            foreach (var attribute in metadata.Attributes.OrderBy(a => a.LogicalName, StringComparer.Ordinal))
            {
                var field = attribute.LogicalName;
                bool shadow = metadata.Attributes.OfType<LookupAttributeMetadata>().Any(a => field == a.LogicalName + "name" || field == a.LogicalName + "yominame") ||
                    !string.IsNullOrWhiteSpace(attribute.AttributeOf) && field.EndsWith("name", StringComparison.OrdinalIgnoreCase);
                bool payload = UnsafePayload.Any(p => field.IndexOf(p, StringComparison.OrdinalIgnoreCase) >= 0) || field == "content";
                bool text = attribute.AttributeType == AttributeTypeCode.String || attribute.AttributeType == AttributeTypeCode.Memo;
                bool shape = text || attribute.AttributeType == AttributeTypeCode.Uniqueidentifier || attribute.AttributeType == AttributeTypeCode.Lookup ||
                    attribute.AttributeType == AttributeTypeCode.EntityName || attribute.AttributeType == AttributeTypeCode.Picklist ||
                    attribute.AttributeType == AttributeTypeCode.Integer || attribute.AttributeType == AttributeTypeCode.Boolean ||
                    attribute.AttributeType == AttributeTypeCode.State || attribute.AttributeType == AttributeTypeCode.Status || attribute.AttributeType == AttributeTypeCode.DateTime;
                bool readable = attribute.IsValidForRead == true && shape && !shadow && !payload;
                bool hash = text && (attribute.AttributeType == AttributeTypeCode.Memo || !AuditText.Contains(field));
                side.AttributeSchema.Add(metadata.LogicalName + "." + field + "; type=" + attribute.AttributeType + "; capture=" +
                    (shadow ? "ExcludedLookupShadow" : payload ? "ExcludedPayload" : readable ? hash ? "HashOnly" : "Audit" : "Unavailable"));
                if (readable) { columns.Add(field); if (hash) hashes.Add(field); }
            }
            var critical = new[] { metadata.PrimaryIdAttribute, link.ReferencingAttribute }.Concat(referenceFields.Where(columns.Contains)).Distinct(StringComparer.Ordinal).ToArray();
            foreach (var batch in Batches(rules.Select(r => r.ObjectId).OrderBy(id => id).ToArray()))
            {
                var selected = rules.Where(r => batch.Contains(r.ObjectId)).ToArray();
                var result = ReadAssociations(service, metadata, link.ReferencingAttribute, side.EntityName, critical,
                    columns.Except(critical).OrderBy(f => f, StringComparer.Ordinal).ToArray(), batch, side, token, progress);
                side.AttributeSchema.Add("Related batch selectedRuleIds=" + batch.Length + "; Status=" + result.Status +
                    "; " + result.Reason + "; CriticalComplete=" + result.CriticalComplete);
                if (result.Status != "Complete") { SetAttributeScope(selected, result.Status, result.Reason); continue; }
                foreach (var item in result.Rows.OrderBy(r => r.Key))
                {
                    var association = new MaskingRuleRecordEvidence { ObjectId = item.Key, PrimaryId = item.Key, Status = "Unique",
                        CriticalComplete = result.CriticalComplete, RelatedMaskingRuleId = AssociationParent(item.Value, link.ReferencingAttribute, side.EntityName) };
                    if (result.Duplicates.Contains(item.Key)) association.Status = "Duplicate";
                    foreach (var field in columns)
                        if (hashes.Contains(field)) association.Content[field] = result.Columns.Contains(field) ? Type31ContentFingerprint.Create(item.Value.GetAttributeValue<object>(field))
                            : new Type31ContentFingerprint { Presence = "Unavailable" };
                        else association.Fields[field] = result.Columns.Contains(field) ? Format(item.Value.GetAttributeValue<object>(field)) : null;
                    association.RuntimeColumns.AddRange(result.Columns.OrderBy(f => f, StringComparer.Ordinal));
                    var tables = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                    ResolveScope(side.Snapshot, side, metadata, referenceFields, new LookupEvidence { Row = item.Value, Columns = result.Columns },
                        association, metadata.LogicalName, association.AttributeKeys, tables);
                    if (tables.Any(table => association.AttributeKeys.Any(attribute => !attribute.StartsWith(table + ".", StringComparison.OrdinalIgnoreCase))))
                    { association.ParentComplete = false; association.ParentStatus = "Ambiguous"; association.Context.Add("Conflicting table versus resolved column parent scope"); }
                    side.AttributeRows[item.Key] = association;
                }
                foreach (var rule in selected)
                {
                    var associations = side.AttributeRows.Values.Where(r => r.RelatedMaskingRuleId == rule.ObjectId).ToArray();
                    rule.AttributeRelationshipCount = associations.Length;
                    if (associations.Length == 0) { SetAttributeScope(new[] { rule }, "Missing", "No selected-rule association rows returned on terminal retrieval"); continue; }
                    if (associations.Any(r => r.Status == "Duplicate" || r.ParentStatus == "Ambiguous") ||
                        associations.SelectMany(r => r.AttributeKeys).GroupBy(k => k, StringComparer.OrdinalIgnoreCase).Any(g => g.Count() > 1))
                    { SetAttributeScope(new[] { rule }, "Ambiguous", "Duplicate/conflicting relationship rows or repeated portable attribute scope"); continue; }
                    if (!result.CriticalComplete || associations.Any(r => !r.ParentComplete || r.AttributeKeys.Count == 0))
                    { SetAttributeScope(new[] { rule }, "Incomplete", "Attribute identity unavailable: readable relationships must resolve through complete unique snapshot Column identities; no textual-field role guessed"); continue; }
                    rule.AttributeKeys.UnionWith(associations.SelectMany(r => r.AttributeKeys));
                    rule.AttributeScopeStatus = "Verified"; rule.AttributeScopeReason = "All related attributes resolved independently; normalized set; no GUID/name/hash scope fallback";
                    if (rule.CriticalComplete && rule.ParentComplete)
                    {
                        string internalValue = side.CandidateField == null ? null : rule.Get(side.CandidateField);
                        string scope = Type31EvidenceCollector.Frame(rule.AttributeKeys.OrderBy(k => k, StringComparer.Ordinal).ToArray());
                        if (!string.IsNullOrWhiteSpace(internalValue) && !Guid.TryParse(internalValue, out var ignored))
                            rule.CandidateA = "maskingrule-candidate-a:" + Type31EvidenceCollector.Frame(side.EntityName, side.CandidateField, internalValue.Trim(), rule.ParentKey, scope);
                        else if (side.CandidateField == null && rule.RuntimeColumns.Contains("name") && !string.IsNullOrWhiteSpace(rule.Get("name")) &&
                            !Guid.TryParse(rule.Get("name"), out var ignoredName))
                            rule.CandidateA = "maskingrule-attribute-scope-candidate-a:" + Type31EvidenceCollector.Frame(side.EntityName, rule.Get("name").Trim(), rule.ParentKey, scope);
                    }
                }
            }
        }

        private static void SetAttributeScope(IEnumerable<MaskingRuleRecordEvidence> rules, string status, string reason)
        { foreach (var rule in rules) { rule.AttributeScopeStatus = status; rule.AttributeScopeReason = reason; rule.AttributeKeys.Clear(); rule.CandidateA = null; } }
        private static IEnumerable<Guid[]> Batches(Guid[] ids)
        { for (int offset = 0; offset < ids.Length; offset += BatchSize) yield return ids.Skip(offset).Take(BatchSize).ToArray(); }
        private static Guid AssociationParent(Entity row, string field, string ruleEntity)
        {
            var value = row.GetAttributeValue<object>(field); var reference = value as EntityReference;
            return value is Guid ? (Guid)value : reference?.LogicalName == ruleEntity ? reference.Id : Guid.Empty;
        }

        private static AssociationBatch ReadAssociations(IOrganizationService service, EntityMetadata metadata, string foreign, string ruleEntity,
            string[] critical, string[] optional, Guid[] ids, MaskingRuleSideEvidence side, CancellationToken token, Action<string> progress)
        {
            string primary = metadata.PrimaryIdAttribute;
            var current = RetrieveAssociations(service, metadata.LogicalName, primary, foreign, ruleEntity, critical, ids, side, token, progress);
            int remaining = MaxIsolationGroups;
            if (current.Status == "Faulted")
            {
                current = RetrieveAssociations(service, metadata.LogicalName, primary, foreign, ruleEntity, new[] { primary, foreign }, ids, side, token, progress);
                if (current.Status == "Complete") IsolateAssociations(service, metadata, foreign, ruleEntity, critical.Except(new[] { primary, foreign }).ToArray(),
                    ids, current, true, side, token, progress, ref remaining);
            }
            if (current.Status == "Complete" && current.CriticalComplete)
                IsolateAssociations(service, metadata, foreign, ruleEntity, optional, ids, current, false, side, token, progress, ref remaining);
            return current;
        }

        private static void IsolateAssociations(IOrganizationService service, EntityMetadata metadata, string foreign, string ruleEntity,
            string[] fields, Guid[] ids, AssociationBatch current, bool critical, MaskingRuleSideEvidence side, CancellationToken token,
            Action<string> progress, ref int remaining)
        {
            if (fields.Length == 0) return;
            var pending = new Queue<string[]>(); pending.Enqueue(fields);
            while (pending.Count > 0)
            {
                token.ThrowIfCancellationRequested(); var group = pending.Dequeue();
                if (remaining == 0)
                {
                    side.RetrievalDiagnostics.Add("Attribute scope isolation limit reached; excluded [" + string.Join(",", group) + "]");
                    if (critical) current.CriticalComplete = false;
                    continue;
                }
                remaining--;
                var result = RetrieveAssociations(service, metadata.LogicalName, metadata.PrimaryIdAttribute, foreign, ruleEntity,
                    new[] { metadata.PrimaryIdAttribute, foreign }.Concat(group).Distinct(StringComparer.Ordinal).ToArray(), ids, side, token, progress);
                if (result.Status == "Faulted")
                {
                    if (group.Length > 1) { int half = group.Length / 2; pending.Enqueue(group.Take(half).ToArray()); pending.Enqueue(group.Skip(half).ToArray()); }
                    else { side.RetrievalDiagnostics.Add("AttributeMaskingRule column faulted: " + group[0] + "; " + (critical ? "CriticalUnavailable" : "OptionalUnavailable")); if (critical) current.CriticalComplete = false; }
                    continue;
                }
                if (result.Status != "Complete" || !current.Rows.Keys.OrderBy(id => id).SequenceEqual(result.Rows.Keys.OrderBy(id => id)) ||
                    current.Rows.Any(r => AssociationParent(r.Value, foreign, ruleEntity) != AssociationParent(result.Rows[r.Key], foreign, ruleEntity)))
                { current.CriticalComplete = false; current.Reason = "Conflicting/incomplete related scope across column retrievals"; continue; }
                current.Duplicates.UnionWith(result.Duplicates);
                foreach (var field in group)
                {
                    current.Columns.Add(field);
                    foreach (var row in current.Rows) if (result.Rows[row.Key].Attributes.TryGetValue(field, out var value)) row.Value[field] = value;
                }
            }
        }

        private static AssociationBatch RetrieveAssociations(IOrganizationService service, string entity, string primary, string foreign, string ruleEntity,
            string[] columns, Guid[] ids, MaskingRuleSideEvidence side, CancellationToken token, Action<string> progress)
        {
            var result = new AssociationBatch(); int page = 1, returned = 0; string cookie = null;
            try
            {
                while (true)
                {
                    token.ThrowIfCancellationRequested();
                    var query = new QueryExpression(entity) { ColumnSet = new ColumnSet(columns), PageInfo = new PagingInfo { Count = BatchSize, PageNumber = page, PagingCookie = cookie } };
                    query.Criteria.AddCondition(foreign, ConditionOperator.In, ids.Select(id => (object)id).ToArray()); query.AddOrder(primary, OrderType.Ascending);
                    side.Requests.Add("RetrieveMultiple " + entity + "; columns=[" + string.Join(",", columns) + "]; " + foreign + " IN selectedMaskingRuleGuid[" + ids.Length + "]; page=" + page);
                    progress?.Invoke(side.Snapshot.Environment.DisplayName + ": reading selected MaskingRule attribute scope page " + page);
                    var response = service.RetrieveMultiple(query); token.ThrowIfCancellationRequested();
                    if (response == null) { result.Reason = "Null related result; terminal retrieval not proven"; return result; }
                    returned += response.Entities.Count;
                    side.Pages.Add(entity + "; requestedRuleIds=" + ids.Length + "; returnedRows=" + response.Entities.Count + "; totalRows=" + returned +
                        "; page=" + page + "; MoreRecords=" + response.MoreRecords + "; PagingCookieSupplied=" + !string.IsNullOrEmpty(response.PagingCookie));
                    if (response.Entities.Any(r => r == null || r.LogicalName != entity || r.Id == Guid.Empty || Type31EvidenceCollector.Id(r, primary) != r.Id ||
                        !ids.Contains(AssociationParent(r, foreign, ruleEntity))))
                    { result.Reason = "Foreign/conflicting association primary or MaskingRule key; no scope proof"; return result; }
                    int before = result.Rows.Count;
                    foreach (var group in response.Entities.GroupBy(r => r.Id))
                    {
                        if (group.Count() > 1) result.Duplicates.Add(group.Key);
                        foreach (var row in group)
                            if (result.Rows.TryGetValue(row.Id, out var prior)) { if (!Type31EvidenceCollector.SameReturnedRow(prior, row)) result.Duplicates.Add(row.Id); }
                            else result.Rows.Add(row.Id, row);
                    }
                    if (!response.MoreRecords)
                    { result.Status = "Complete"; result.Reason = "Terminal related page; distinctRows=" + result.Rows.Count + "; pageCount=" + page; result.Columns.UnionWith(columns); return result; }
                    if (before == result.Rows.Count || !string.IsNullOrEmpty(response.PagingCookie) && response.PagingCookie == cookie)
                    { result.Reason = "Stalled related paging; terminal retrieval not proven"; return result; }
                    cookie = response.PagingCookie; page++;
                }
            }
            catch (OperationCanceledException) { throw; }
            catch (Exception error) { token.ThrowIfCancellationRequested(); result.Status = "Faulted"; result.Reason = SafeFault(error); side.RetrievalDiagnostics.Add("Attribute scope query fault; " + result.Reason); return result; }
        }
        private sealed class AssociationBatch
        {
            internal string Status = "Incomplete", Reason;
            internal bool CriticalComplete = true;
            internal readonly SortedDictionary<Guid, Entity> Rows = new SortedDictionary<Guid, Entity>();
            internal readonly HashSet<Guid> Duplicates = new HashSet<Guid>();
            internal readonly HashSet<string> Columns = new HashSet<string>(StringComparer.Ordinal);
        }

        private static int? RawTypeFor(string entity)
        {
            switch (entity) { case "entity": return 1; case "attribute": return 2; case "relationship": return 10;
                case "webresource": return 61; case "appmodule": return 80; case "savedquery": return 26; case "workflow": return 29;
                case "savedqueryvisualization": return 59; case "systemform": return 60; case "sitemap": return 62;
                case "pluginassembly": return 91; case "sdkmessageprocessingstep": return 92; case "environmentvariabledefinition": return 380; default: return null; }
        }
        private static string PrimaryFor(string entity) => RawTypeFor(entity).HasValue ? entity == "systemform" ? "formid" : entity + "id" : null;
        private static EntityMetadata Schema(IOrganizationService service, string entity, MaskingRuleSideEvidence side, CancellationToken token)
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
            string[] critical, string[] optional, Guid[] ids, MaskingRuleSideEvidence side, CancellationToken token, Action<string> progress)
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
            Dictionary<Guid, LookupEvidence> current, bool critical, MaskingRuleSideEvidence side, CancellationToken token,
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
            Guid[] ids, MaskingRuleSideEvidence side, CancellationToken token, Action<string> progress)
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

    internal sealed class MaskingRuleSideEvidence
    {
        internal MembershipSnapshot Snapshot;
        internal string Version, EntityName, PrimaryId, CandidateField, SchemaFailure;
        internal readonly List<ComponentIdentity> Raw = new List<ComponentIdentity>();
        internal readonly SortedDictionary<Guid, MaskingRuleRecordEvidence> Rows = new SortedDictionary<Guid, MaskingRuleRecordEvidence>();
        internal int AttributeRequestStart = -1;
        internal string AttributeEntity, AttributeForeignKey;
        internal readonly List<string> AttributeSchema = new List<string>();
        internal readonly SortedDictionary<Guid, MaskingRuleRecordEvidence> AttributeRows = new SortedDictionary<Guid, MaskingRuleRecordEvidence>();
        internal readonly List<string> Schema = new List<string>(), Relationships = new List<string>(), Requests = new List<string>(), Pages = new List<string>(), RetrievalDiagnostics = new List<string>();
    }
    internal sealed class MaskingRuleRecordEvidence
    {
        internal Guid ObjectId; internal Guid? PrimaryId; internal bool? Managed;
        internal string Status, Reason, CandidateField, CandidateA, CandidateB;
        internal bool CriticalComplete, DuplicateA, DuplicateB, ParentComplete;
        internal string ParentKey, ParentStatus;
        internal Guid RelatedMaskingRuleId;
        internal string AttributeScopeStatus = "NotExposed", AttributeScopeReason;
        internal int AttributeRelationshipCount;
        internal readonly SortedSet<string> AttributeKeys = new SortedSet<string>(StringComparer.OrdinalIgnoreCase);
        internal readonly List<string> Context = new List<string>();
        internal readonly SortedDictionary<string, string> Fields = new SortedDictionary<string, string>(StringComparer.Ordinal);
        internal readonly SortedDictionary<string, Type31ContentFingerprint> Content = new SortedDictionary<string, Type31ContentFingerprint>(StringComparer.Ordinal);
        internal readonly List<string> RuntimeColumns = new List<string>();
        internal string Get(string field) => Fields.TryGetValue(field, out var value) ? value : null;
        internal string Evidence(string field) => Content.TryGetValue(field, out var hash) ? hash.Evidence : Get(field);
    }
    internal sealed class MaskingRulePairEvidence
    {
        internal MaskingRuleRecordEvidence Source, Target; internal string Outcome;
        internal readonly HashSet<string> Categories = new HashSet<string>(StringComparer.Ordinal);
    }
    internal sealed class MaskingRuleEvidenceReport
    {
        internal MaskingRuleSideEvidence Source, Target;
        internal readonly List<MaskingRulePairEvidence> Pairs = new List<MaskingRulePairEvidence>();
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
                    foreach (var row in left) Pairs.Add(new MaskingRulePairEvidence { Source = row, Outcome = "Ambiguous" });
                    foreach (var row in right) Pairs.Add(new MaskingRulePairEvidence { Target = row, Outcome = "Ambiguous" });
                }
                else Pairs.Add(new MaskingRulePairEvidence { Source = left.SingleOrDefault(), Target = right.SingleOrDefault(),
                    Outcome = left.Length == 1 && right.Length == 1 ? "SemanticPair" : "OneSidedEvidence" });
            }
            foreach (var side in new[] { Source, Target })
            {
                foreach (var row in side.Rows.Values.Where(r => r.CandidateA == null))
                    Pairs.Add(new MaskingRulePairEvidence { Source = side == Source ? row : null, Target = side == Target ? row : null,
                        Outcome = row.Status == "Duplicate" || row.ParentStatus == "Ambiguous" || row.AttributeScopeStatus == "Ambiguous" ? "Ambiguous" : "Incomplete" });
                foreach (var raw in side.Raw.Where(r => !r.Record.ObjectId.HasValue || r.Record.ObjectId == Guid.Empty))
                    Pairs.Add(new MaskingRulePairEvidence { Outcome = "Incomplete" });
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
            var text = new StringBuilder(); text.AppendLine("TYPE 74 MASKINGRULE EVIDENCE - DEBUG ONLY");
            text.AppendLine("Evidence only. Type 74 remains Unsupported / Indeterminate; no portable comparison key or absence proof is supplied.");
            text.AppendLine("Uniqueness is scoped to selected solution members; rename/recreate portability remains unproven. Candidate B and hashes never repair Candidate A.");
            text.AppendLine("\nRAW TYPE 74 MEMBERSHIP");
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
            text.AppendLine("\nCANDIDATE IDENTITY ANALYSIS"); EachSide(text, (s, label) => {
                Line(text, label, "BackingEntity=" + s.EntityName, "Candidate A hypothesis: strongest internal text field=" + (s.CandidateField ?? "Unavailable"),
                    "priority=uniquename,schemaname,logicalname + all metadata-exposed non-audit references using verified snapshot identities; hypotheses only",
                    "Additional hypothesis where no internal identifier exists: rule name + verified normalized attribute scope set. Name alone never supplies A; B/hash/GUID overlap never repairs scope",
                    "Candidate B: descriptive name + scope; never repairs A; no display-name-only A");
                foreach (var row in s.Rows.Values) {
                    Line(text, label, row.ObjectId, "CandidateA=" + (row.CandidateA ?? "Incomplete"), "CandidateB=" + (row.CandidateB ?? "NotAvailable"),
                        "DuplicateA=" + row.DuplicateA, "DuplicateB=" + row.DuplicateB, "ParentStatus=" + row.ParentStatus, "ParentIdentityComplete=" + row.ParentComplete);
                    foreach (var item in row.Context) Line(text, label, row.ObjectId, item);
                }
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
            text.AppendLine("\nATTRIBUTEMASKINGRULE SCHEMA");
            EachSide(text, (s, label) => { foreach (var item in s.AttributeSchema) Line(text, label, item); });
            text.AppendLine("\nMASKINGRULE -> ATTRIBUTEMASKINGRULE CORRELATION");
            EachSide(text, (s, label) => { foreach (var row in s.AttributeRows.Values) {
                Line(text, label, "MaskingRuleId=" + row.RelatedMaskingRuleId, "AssociationId=" + row.PrimaryId, row.Status,
                    "ForeignKey=" + s.AttributeForeignKey, "CriticalComplete=" + row.CriticalComplete);
                foreach (var field in row.Fields.Keys.Union(row.Content.Keys).OrderBy(f => f, StringComparer.Ordinal)) Line(text, label, row.ObjectId, field, row.Evidence(field) ?? "Unavailable");
            } });
            text.AppendLine("\nRELATED TABLE / ATTRIBUTE PORTABLE SCOPE");
            EachSide(text, (s, label) => { foreach (var row in s.AttributeRows.Values) {
                Line(text, label, row.RelatedMaskingRuleId, row.ObjectId, row.ParentStatus, "Attributes=[" + string.Join(",", row.AttributeKeys) + "]");
                foreach (var item in row.Context) Line(text, label, row.ObjectId, item);
            } });
            text.AppendLine("\nSCOPE MULTIPLICITY");
            EachSide(text, (s, label) => { foreach (var row in s.Rows.Values) Line(text, label, row.ObjectId, "AssociationRows=" + row.AttributeRelationshipCount,
                "DistinctPortableAttributes=" + row.AttributeKeys.Count, "NormalizedAttributeSet=[" + string.Join(",", row.AttributeKeys) + "]", row.AttributeScopeStatus, row.AttributeScopeReason); });
            text.AppendLine("\nUPDATED CANDIDATE A COMPLETENESS");
            EachSide(text, (s, label) => { foreach (var row in s.Rows.Values) Line(text, label, row.ObjectId, "CompleteA=" + (row.CandidateA != null && !row.DuplicateA),
                "ParentIdentityComplete=" + row.ParentComplete, "RelatedAttributeScope=" + row.AttributeScopeStatus, "CandidateA=" + (row.CandidateA ?? "Incomplete"),
                "BlockingReason=" + (row.CandidateA != null && !row.DuplicateA ? "None" : row.DuplicateA ? "Duplicate canonical Candidate A" :
                    row.AttributeScopeStatus != "NotExposed" && row.AttributeScopeStatus != "Verified" ? row.AttributeScopeReason :
                    "Incomplete primary/critical evidence, parent identity or explicit identifier; name alone cannot create A")); });
            text.AppendLine("\nEXACT ADDITIONAL ATTRIBUTE SCOPE REQUEST LEDGER");
            EachSide(text, (s, label) => {
                var requests = s.AttributeRequestStart < 0 ? new string[0] : s.Requests.Skip(s.AttributeRequestStart).ToArray();
                foreach (var request in requests) Line(text, label, request);
                Line(text, label, "AdditionalScopeReads=" + requests.Length, "ParentSnapshotReuse=True", "AdditionalWhoAmI=0", "Writes=0", "NormalMembershipEvidenceRequests=0");
            });
            text.AppendLine("\nPORTABILITY ASSESSMENT");
            Line(text, "Unique semantic pairs=" + Pairs.Count(p => p.Outcome == "SemanticPair"), "differing primary IDs=" + Pairs.Count(p => p.Categories.Contains("DifferentPrimaryId")));
            text.AppendLine("Observed evidence does not establish lifecycle portability or safe absence semantics. Additional live/lifecycle review required before production promotion or any production use.");
            text.AppendLine("\nEXACT REQUEST LEDGER"); EachSide(text, (s, label) => {
                foreach (var request in s.Requests) Line(text, label, request);
                foreach (var group in s.Requests.GroupBy(r => r.StartsWith("Execute", StringComparison.Ordinal) ? r.Split('(')[0] + "(" + r.Split('(')[1].Split(',')[0] + ")" : r.Split(';')[0])) Line(text, label, group.Key, "reads=" + group.Count());
                Line(text, label, "TotalReads=" + s.Requests.Count, "AdditionalWhoAmI=0", "Writes=0", "NormalMembershipEvidenceRequests=0");
            }); return text.ToString();
        }
        private void EachSide(StringBuilder text, Action<MaskingRuleSideEvidence, string> action) { action(Source, "Source"); action(Target, "Target"); }
    }
}
#endif
