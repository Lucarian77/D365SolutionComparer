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
    /// <summary>Explicit Debug capture. Supplies neither production identities nor absence evidence.</summary>
    internal sealed class Type10072EvidenceCollector
    {
        internal const int BatchSize = 200;
        internal const int MaxIsolationGroups = 64;
        private static readonly string[] MinimalColumns = { "appelementid", "parentappmoduleid", "objectid", "objectidtype", "name",
            "uniquename", "componentidunique", "componentstate", "ismanaged", "canvasappid" };
        private static readonly string[] TextEvidence = { "name", "uniquename", "schemaname", "logicalname", "displayname",
            "entitylogicalname", "entityname", "primaryentity", "objecttypecode", "objectidtype" };
        private static readonly string[] ScalarEvidence = { "type", "elementtype", "componenttype", "objecttypecode", "objectidtype", "ismanaged", "componentstate", "statecode", "statuscode" };
        // Lookup display shadows can be marked readable by metadata yet rejected by RetrieveMultiple (0x8004023B).
        // Retain the actual lookups; these virtual display values provide no identity evidence.
        private static readonly HashSet<string> LookupShadows = new HashSet<string>(new[] {
            "canvasappidname", "createdbyname", "createdbyyominame", "createdonbehalfbyname", "createdonbehalfbyyominame",
            "modifiedbyname", "modifiedbyyominame", "modifiedonbehalfbyname", "modifiedonbehalfbyyominame",
            "organizationidname", "parentappmoduleidname" }, StringComparer.OrdinalIgnoreCase);
        // Installation/ownership references are retained as audit evidence, not Candidate A component dependencies.
        private static readonly HashSet<string> AuditReferences = new HashSet<string>(new[] {
            "createdby", "modifiedby", "createdonbehalfby", "modifiedonbehalfby", "ownerid", "owninguser",
            "owningteam", "owningbusinessunit", "organizationid" }, StringComparer.OrdinalIgnoreCase);

        internal AppElementEvidenceReport Capture(IOrganizationService sourceService, MembershipSnapshot source, string sourceVersion,
            IOrganizationService targetService, MembershipSnapshot target, string targetVersion, CancellationToken token, Action<string> progress = null,
            CanvasAppEvidenceReport completedType300Evidence = null)
        {
            if (source?.State != MembershipSnapshotState.Complete || target?.State != MembershipSnapshotState.Complete ||
                !StringComparer.OrdinalIgnoreCase.Equals(source.SolutionUniqueName, target.SolutionUniqueName))
                throw new ArgumentException("Completed snapshots of the same solution are required.");
            token.ThrowIfCancellationRequested();
            var report = new AppElementEvidenceReport
            {
                Source = Read(sourceService, source, sourceVersion, token, progress, ReferenceEquals(completedType300Evidence?.Source.Snapshot, source) ? completedType300Evidence.Source.Metadata : null),
                Target = Read(targetService, target, targetVersion, token, progress, ReferenceEquals(completedType300Evidence?.Target.Snapshot, target) ? completedType300Evidence.Target.Metadata : null)
            };
            CompleteCanvasReferences(sourceService, targetService, report, completedType300Evidence, token, progress);
            foreach (var side in new[] { report.Source, report.Target })
                foreach (var group in side.Rows.Values.Where(r => r.Status == "Unique" && r.CandidateA != null)
                    .GroupBy(r => r.CandidateA, StringComparer.OrdinalIgnoreCase).Where(g => g.Count() > 1))
                    foreach (var row in group) row.DuplicateA = true;
            foreach (var side in new[] { report.Source, report.Target })
                foreach (var group in side.Rows.Values.Where(r => r.Status == "Unique" && r.CandidateB != null)
                    .GroupBy(r => r.CandidateB, StringComparer.OrdinalIgnoreCase).Where(g => g.Count() > 1))
                    foreach (var row in group) row.DuplicateB = true;
            report.Analyze(token);
            return report;
        }

        private static AppElementSideEvidence Read(IOrganizationService service, MembershipSnapshot snapshot, string version,
            CancellationToken token, Action<string> progress, EntityMetadata cachedCanvasMetadata = null)
        {
            if (service == null) throw new ArgumentNullException(nameof(service));
            var side = new AppElementSideEvidence { Snapshot = snapshot, Version = version };
            if (cachedCanvasMetadata?.LogicalName == "canvasapp") side.TypeMetadataCache["canvasapp"] = cachedCanvasMetadata;
            side.Raw.AddRange(snapshot.Components.Where(c => c.Record.ComponentType == 10072));
            var ids = side.Raw.Where(c => c.Record.ObjectId.HasValue && c.Record.ObjectId != Guid.Empty)
                .Select(c => c.Record.ObjectId.Value).Distinct().OrderBy(id => id).ToArray();
            foreach (var id in ids) side.Rows[id] = new AppElementRecordEvidence { ObjectId = id, Status = "Incomplete", Reason = "Metadata not verified" };
            if (ids.Length == 0) return side;
            var metadata = Schema(service, "appelement", side, token);
            if (metadata == null)
            { foreach (var row in side.Rows.Values) { row.Status = side.SchemaFailure; row.Reason = "AppElement schema unavailable; no guessed columns"; } return side; }
            side.PrimaryId = metadata.PrimaryIdAttribute;
            var columns = new List<string>(); var hashes = new HashSet<string>(StringComparer.Ordinal);
            foreach (var attribute in metadata.Attributes.OrderBy(a => a.LogicalName, StringComparer.Ordinal))
            {
                var name = attribute.LogicalName;
                bool lookupShadow = LookupShadows.Contains(name) || !string.IsNullOrWhiteSpace(attribute.AttributeOf) &&
                    metadata.Attributes.Any(a => a.AttributeType == AttributeTypeCode.Lookup &&
                        StringComparer.OrdinalIgnoreCase.Equals(a.LogicalName, attribute.AttributeOf)) &&
                    (StringComparer.OrdinalIgnoreCase.Equals(name, attribute.AttributeOf + "name") ||
                     StringComparer.OrdinalIgnoreCase.Equals(name, attribute.AttributeOf + "yominame"));
                bool excluded = new[] { "attachment", "binary", "base64", "encoded", "secure", "secret", "credential" }.Any(p => name.IndexOf(p, StringComparison.OrdinalIgnoreCase) >= 0);
                bool text = attribute.AttributeType == AttributeTypeCode.String || attribute.AttributeType == AttributeTypeCode.Memo;
                bool typeProbe = PlausibleTypeField(attribute);
                bool content = text && (!TextEvidence.Contains(name) || attribute.AttributeType == AttributeTypeCode.Memo);
                bool allowedShape = name == side.PrimaryId ? attribute.AttributeType == AttributeTypeCode.Uniqueidentifier :
                    text || attribute.AttributeType == AttributeTypeCode.EntityName || attribute.AttributeType == AttributeTypeCode.Uniqueidentifier ||
                    attribute.AttributeType == AttributeTypeCode.Lookup || (ScalarEvidence.Contains(name) || typeProbe) &&
                    (attribute.AttributeType == AttributeTypeCode.Picklist || attribute.AttributeType == AttributeTypeCode.Integer ||
                     attribute.AttributeType == AttributeTypeCode.Boolean || attribute.AttributeType == AttributeTypeCode.State || attribute.AttributeType == AttributeTypeCode.Status);
                bool readable = attribute.IsValidForRead == true && !excluded && !lookupShadow && allowedShape;
                side.Schema.Add("appelement." + name + "; type=" + attribute.AttributeType + "; readable=" + attribute.IsValidForRead +
                    "; capture=" + (lookupShadow ? "ExcludedLookupShadow" : readable ? content ? "HashOnly" : "Audit" : "UnavailableOrNotQueried"));
                if (!readable) continue;
                columns.Add(name); if (typeProbe) side.TypeRequestedColumns.Add(name); if (content) hashes.Add(name);
            }
            if (!columns.Contains(side.PrimaryId))
            { foreach (var row in side.Rows.Values) { row.Reason = "Primary ID is not readable; no backing query"; } return side; }
            var backing = RetrieveAppElements(service, side.PrimaryId, columns.ToArray(), ids, side, token, progress);
            foreach (var id in ids)
            {
                var row = side.Rows[id]; var found = backing[id]; row.Status = found.Status; row.Reason = found.Reason;
                if (found.Status != "Unique") continue;
                row.PrimaryId = found.Row.Id;
                row.CriticalComplete = found.CriticalComplete;
                row.RuntimeColumns.UnionWith(found.Columns);
                foreach (var attribute in metadata.Attributes.Where(PlausibleTypeField))
                    row.TypeValues[attribute.LogicalName] = found.Columns.Contains(attribute.LogicalName)
                        ? SafeTypeValue(found.Row.GetAttributeValue<object>(attribute.LogicalName)) : null;
                row.Context.Add("Successfully retrieved columns=[" + string.Join(",", found.Columns.OrderBy(c => c, StringComparer.Ordinal)) + "]");
                if (!found.CriticalComplete) row.Context.Add("Identity-critical evidence unavailable; neither candidate may be evaluated");
                foreach (var column in columns.Except(found.Columns))
                    if (hashes.Contains(column)) row.Content[column] = new Type31ContentFingerprint { Presence = "Unavailable" };
                    else row.Fields[column] = null;
                foreach (var column in found.Columns)
                {
                    var raw = found.Row.GetAttributeValue<object>(column);
                    if (hashes.Contains(column)) { row.Content[column] = Type31ContentFingerprint.Create(raw); continue; }
                    row.Fields[column] = Format(raw);
                    if (raw is EntityReference)
                    {
                        var reference = (EntityReference)raw;
                        if (reference.Id != Guid.Empty && ValidName(reference.LogicalName)) row.References[column] = reference;
                    }
                    // Untyped GUIDs need a metadata-proven PK relationship before reference interpretation.
                    else if (raw is Guid && (Guid)raw != Guid.Empty && !AuditReferences.Contains(column))
                    {
                        var links = (metadata.ManyToOneRelationships ?? new OneToManyRelationshipMetadata[0])
                            .Where(r => r != null && r.ReferencingEntity == "appelement" && r.ReferencingAttribute == column &&
                                ValidName(r.ReferencedEntity) && ValidName(r.ReferencedAttribute)).ToArray();
                        if (links.Length == 1) row.GuidLinks[column] = Tuple.Create(links[0].ReferencedEntity, links[0].ReferencedAttribute, (Guid)raw);
                        else if (links.Length > 1) row.RelationshipAmbiguous = true;
                    }
                }
                row.Managed = found.Row.GetAttributeValue<object>("ismanaged") is bool ? (bool?)found.Row.GetAttributeValue<bool>("ismanaged") : null;
            }
            ResolveContext(service, side, metadata, token, progress);
            InvestigateTypeSources(service, side, metadata, token);
            return side;
        }

        // Evidence only: this analysis never writes ElementType, CandidateA or CandidateB.
        private static bool PlausibleTypeField(AttributeMetadata attribute)
        {
            var name = attribute.LogicalName;
            bool shape = attribute.AttributeType == AttributeTypeCode.EntityName || attribute.AttributeType == AttributeTypeCode.Picklist ||
                attribute.AttributeType == AttributeTypeCode.Integer || attribute.AttributeType == AttributeTypeCode.String;
            bool named = new[] { "objectidtype", "componenttype", "type", "elementtype", "appelementtype", "objecttypecode" }.Contains(name) ||
                name.IndexOf("componentkind", StringComparison.OrdinalIgnoreCase) >= 0 || name.IndexOf("elementkind", StringComparison.OrdinalIgnoreCase) >= 0 ||
                name.IndexOf("componenttype", StringComparison.OrdinalIgnoreCase) >= 0 || name.IndexOf("elementtype", StringComparison.OrdinalIgnoreCase) >= 0;
            // Descriptions are used only to locate schema candidates; labels/descriptions never establish a discriminator contract.
            var descriptions = attribute.Description?.LocalizedLabels.Select(l => l.Label) ?? Enumerable.Empty<string>();
            bool described = descriptions.Any(d => d != null && (d.IndexOf("component kind", StringComparison.OrdinalIgnoreCase) >= 0 ||
                d.IndexOf("element type", StringComparison.OrdinalIgnoreCase) >= 0 || d.IndexOf("referenced entity", StringComparison.OrdinalIgnoreCase) >= 0));
            return shape && (named || described || attribute.AttributeType == AttributeTypeCode.EntityName);
        }

        private static string SafeTypeValue(object raw)
        {
            if (raw == null) return "";
            if (raw is OptionSetValue) return ((OptionSetValue)raw).Value.ToString(CultureInfo.InvariantCulture);
            if (raw is int) return ((int)raw).ToString(CultureInfo.InvariantCulture);
            if (raw is string && (((string)raw).Trim().Length <= 128 && ValidName(((string)raw).Trim()))) return ((string)raw).Trim();
            return "MalformedOrRedacted"; // Never export an arbitrary string payload as a type discriminator.
        }

        private static void InvestigateTypeSources(IOrganizationService service, AppElementSideEvidence side, EntityMetadata metadata, CancellationToken token)
        {
            var attributes = metadata.Attributes.Where(PlausibleTypeField).OrderBy(a => a.LogicalName, StringComparer.Ordinal).ToArray();
            foreach (var attribute in attributes)
            {
                bool selected = side.TypeRequestedColumns.Contains(attribute.LogicalName);
                side.TypeSchema.Add("appelement." + attribute.LogicalName + "; AttributeType=" + attribute.AttributeType +
                    "; MetadataReadable=" + attribute.IsValidForRead + "; Selected=" + selected + "; AttributeMetadataId=" + attribute.MetadataId);
                var enumeration = attribute as EnumAttributeMetadata;
                foreach (var option in (enumeration?.OptionSet?.Options ?? new OptionMetadataCollection()).Where(o => o != null))
                    side.TypeSchema.Add("OptionProvenance=RetrieveEntity(appelement,Attributes|Relationships); field=" + attribute.LogicalName +
                        "; OptionSetName=" + enumeration.OptionSet.Name + "; OptionSetMetadataId=" + enumeration.OptionSet.MetadataId +
                        "; IsGlobal=" + enumeration.OptionSet.IsGlobal + "; NumericValue=" + option.Value + "; LabelsAuditOnly=[" +
                        string.Join(" | ", option.Label?.LocalizedLabels.Select(l => "LCID=" + l.LanguageCode + ":" + l.Label) ?? Enumerable.Empty<string>()) + "]");
            }
            foreach (var name in new[] { "objectidtype", "componenttype", "type", "elementtype", "appelementtype", "objecttypecode" }.Except(attributes.Select(a => a.LogicalName)))
                side.TypeSchema.Add("appelement." + name + "; NotExposedOrNotTypeBearing=True; Selected=False");
            // EntityName strings are checked independently against selected entity metadata, never inferred from a lookup/GUID.
            var entityNames = new Dictionary<string, bool>(StringComparer.OrdinalIgnoreCase);
            foreach (var row in side.Rows.Values.Where(r => r.Status == "Unique"))
            {
                token.ThrowIfCancellationRequested();
                var verified = new List<string>();
                foreach (var attribute in attributes)
                {
                    string field = attribute.LogicalName;
                    row.TypeValues.TryGetValue(field, out var value);
                    string state = !row.RuntimeColumns.Contains(field) ? "Unavailable" : string.IsNullOrEmpty(value) ? "NotPresent" : "ObservedSemanticsUnproven";
                    var enumeration = attribute as EnumAttributeMetadata;
                    if (!string.IsNullOrEmpty(value) && value != "MalformedOrRedacted" && row.RuntimeColumns.Contains(field))
                    {
                        int number;
                        var options = enumeration?.OptionSet?.Options.Where(o => o != null && o.Value.HasValue && int.TryParse(value, NumberStyles.Integer,
                            CultureInfo.InvariantCulture, out number) && o.Value.Value == number).ToArray() ?? new OptionMetadata[0];
                        if (attribute.AttributeType == AttributeTypeCode.Picklist)
                        {
                            state = options.Length == 1 ? "VerifiedOptionMetadataValue" : options.Length > 1 ? "AmbiguousOptionMetadata" : "OptionMappingUnavailable";
                            if (options.Length == 1) verified.Add(field + ":picklist:" + value);
                        }
                        else if (attribute.AttributeType == AttributeTypeCode.EntityName)
                        {
                            if (!ValidName(value) || int.TryParse(value, out number)) state = "EntityNameInterpretationUnavailable";
                            else
                            {
                                if (!entityNames.ContainsKey(value))
                                {
                                    var entitySchema = StringComparer.OrdinalIgnoreCase.Equals(value, "appelement") ? metadata : Schema(service, value, side, token);
                                    entityNames[value] = entitySchema != null && StringComparer.OrdinalIgnoreCase.Equals(entitySchema.LogicalName, value);
                                }
                                state = entityNames[value] ? "VerifiedEntityNameTarget; referenced-component discriminator semantics unproven" : "EntityNameTargetUnavailable";
                                // Target existence alone does not prove which AppElement kind the field represents.
                            }
                        }
                    }
                    row.TypeAnalysis.Add("Field=" + field + "; RuntimeReadable=" + row.RuntimeColumns.Contains(field) +
                        "; ActualValue=" + (value ?? "Unavailable") + "; DiscriminatorStatus=" + state + "; Provenance=selected appelement row + RetrieveEntity attribute metadata");
                }
                // Values from different domains/fields are not silently equated; multiple populated sources need an explicit mapping contract.
                var populated = attributes.Where(a => row.TypeValues.TryGetValue(a.LogicalName, out var v) && !string.IsNullOrEmpty(v)).ToArray();
                row.IndependentTypeStatus = populated.Length > 1 ? "AmbiguousMultipleDiscriminators; no inter-field mapping contract" :
                    verified.Count == 1 && populated.Length == 1 ? "VerifiedExplicitOptionDiscriminatorEvidence" : "Incomplete; no independently verified discriminator";
                row.IndependentTypeComplete = populated.Length == 1 && verified.Count == 1;
                row.TypeAnalysis.Add("CandidateAUnchanged=True; IndependentElementTypeComplete=" + row.IndependentTypeComplete +
                    "; IndependentTypeStatus=" + row.IndependentTypeStatus + "; Canvas App/parent/GUID/B/hash evidence cannot supply type");
            }
            if (side.Rows.Values.Any(r => r.Status == "Unique")) ReadTypeRegistration(service, side, token);
        }

        private static void ReadTypeRegistration(IOrganizationService service, AppElementSideEvidence side, CancellationToken token)
        {
            foreach (var definition in side.Raw.Select(r => r.RegisteredDefinition).Where(d => d != null).Distinct())
                side.TypeRegistration.Add("CompletedSnapshotRegistration; ObjectTypeCode=" + definition.ObjectTypeCode +
                    "; Name=" + definition.Name + "; PrimaryEntityName=" + definition.PrimaryEntityName + "; catalog evidence only; no child-type mapping inferred");
            var metadata = Schema(service, "solutioncomponentdefinition", side, token);
            if (metadata == null) { side.TypeRegistration.Add("Registered-definition metadata unavailable; no discriminator contract inferred"); return; }
            var code = metadata.Attributes.SingleOrDefault(a => a.LogicalName == "objecttypecode");
            if (code?.IsValidForRead != true || code.AttributeType != AttributeTypeCode.Integer && code.AttributeType != AttributeTypeCode.Picklist)
            { side.TypeRegistration.Add("Readable registered objecttypecode unavailable; no query or inferred child type"); return; }
            var columns = metadata.Attributes.Where(a => a.IsValidForRead == true && (a.LogicalName == metadata.PrimaryIdAttribute ||
                a.LogicalName == "objecttypecode" || a.LogicalName == "name" || a.LogicalName == "primaryentityname" || PlausibleTypeField(a)))
                .Where(a => a.AttributeType == AttributeTypeCode.Uniqueidentifier || a.AttributeType == AttributeTypeCode.String ||
                    a.AttributeType == AttributeTypeCode.EntityName || a.AttributeType == AttributeTypeCode.Integer || a.AttributeType == AttributeTypeCode.Picklist)
                .Select(a => a.LogicalName).Distinct().OrderBy(f => f, StringComparer.Ordinal).ToArray();
            foreach (var attribute in metadata.Attributes.Where(PlausibleTypeField))
                side.TypeRegistration.Add("RegistrationSchema; Field=" + attribute.LogicalName + "; AttributeType=" + attribute.AttributeType +
                    "; MetadataReadable=" + attribute.IsValidForRead + "; Selected=" + columns.Contains(attribute.LogicalName));
            if (!columns.Contains(metadata.PrimaryIdAttribute)) { side.TypeRegistration.Add("Registration primary ID unavailable; no query"); return; }
            var seen = new Dictionary<Guid, Entity>(); int page = 1; string cookie = null;
            try
            {
                while (true)
                {
                    token.ThrowIfCancellationRequested();
                    var query = new QueryExpression("solutioncomponentdefinition") { ColumnSet = new ColumnSet(columns),
                        PageInfo = new PagingInfo { Count = BatchSize, PageNumber = page, PagingCookie = cookie } };
                    query.Criteria.AddCondition("objecttypecode", ConditionOperator.Equal, 10072); query.AddOrder(metadata.PrimaryIdAttribute, OrderType.Ascending);
                    side.Requests.Add("RetrieveMultiple solutioncomponentdefinition; objecttypecode=10072; page=" + page + "; selected registration only");
                    var response = service.RetrieveMultiple(query); token.ThrowIfCancellationRequested();
                    if (response == null) { side.TypeRegistration.Add("Registration incomplete response"); return; }
                    side.Pages.Add("solutioncomponentdefinition; page=" + page + "; ReturnedRows=" + response.Entities.Count +
                        "; MoreRecords=" + response.MoreRecords + "; PagingCookieSupplied=" + !string.IsNullOrEmpty(response.PagingCookie));
                    int before = seen.Count;
                    foreach (var row in response.Entities)
                    {
                        if (row == null || row.LogicalName != "solutioncomponentdefinition" || row.Id == Guid.Empty ||
                            Type31EvidenceCollector.Id(row, metadata.PrimaryIdAttribute) != row.Id || SafeTypeValue(row.GetAttributeValue<object>("objecttypecode")) != "10072")
                        { side.TypeRegistration.Add("Registration incomplete/conflicting primary key or raw code; no inference"); return; }
                        if (seen.TryGetValue(row.Id, out var prior) && !Type31EvidenceCollector.SameReturnedRow(prior, row))
                        { side.TypeRegistration.Add("Conflicting registration rows; no inference"); return; }
                        seen[row.Id] = row;
                    }
                    if (!response.MoreRecords) break;
                    if (seen.Count == before || !string.IsNullOrEmpty(response.PagingCookie) && response.PagingCookie == cookie)
                    { side.TypeRegistration.Add("Registration stalled paging; terminal retrieval not proven"); return; }
                    cookie = response.PagingCookie; page++;
                }
                side.TypeRegistration.Add("Terminal registration retrieval; distinctRows=" + seen.Count + "; pageCount=" + page + "; catalog labels do not establish selected-row subtype");
                foreach (var row in seen.Values) foreach (var column in columns)
                    side.TypeRegistration.Add("RegistrationRow=" + row.Id + "; Field=" + column + "; Value=" + SafeTypeValue(row.GetAttributeValue<object>(column)) +
                        "; No explicit selected-AppElement discriminator mapping established by numeric agreement/catalog label alone");
            }
            catch (OperationCanceledException) { throw; }
            catch (Exception error) { token.ThrowIfCancellationRequested(); side.TypeRegistration.Add("Registration fault; " + SafeFault(error) + "; no type inference"); }
        }

        private static EntityMetadata Schema(IOrganizationService service, string entity, AppElementSideEvidence side, CancellationToken token)
        {
            token.ThrowIfCancellationRequested();
            if (side.TypeMetadataCache.TryGetValue(entity, out var cachedMetadata))
            { side.TypeSchema.Add("EntityName target metadata reused within the same snapshot/operation: " + entity); return cachedMetadata; }
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
                side.TypeMetadataCache[entity] = metadata;
                return metadata;
            }
            catch (OperationCanceledException) { throw; }
            catch (Exception) { token.ThrowIfCancellationRequested(); side.SchemaFailure = "Faulted"; side.Schema.Add(entity + ": schema retrieval failed; server details withheld"); return null; }
        }

        private static void ResolveContext(IOrganizationService service, AppElementSideEvidence side, EntityMetadata metadata,
            CancellationToken token, Action<string> progress)
        {
            var rows = side.Rows.Values.Where(r => r.Status == "Unique").ToArray();
            var parents = new Dictionary<Guid, LookupEvidence>();
            // References to appmodule are discovered from validated lookup metadata/relationships.
            foreach (var row in rows)
            {
                foreach (var reference in row.References.Where(r => r.Value.LogicalName == "appmodule"))
                {
                    var attribute = metadata.Attributes.OfType<LookupAttributeMetadata>().SingleOrDefault(a => a.LogicalName == reference.Key);
                    if (attribute?.Targets?.Contains("appmodule") == true) row.ParentIds.Add(reference.Value.Id);
                    else row.RelationshipAmbiguous = true;
                }
            }
            var ids = rows.SelectMany(r => r.ParentIds).Distinct().OrderBy(id => id).ToArray();
            foreach (var id in ids)
            {
                var matches = side.Snapshot.Components.Where(c => c.Record.ComponentType == 80 && c.Record.ObjectId == id).ToArray();
                var keys = matches.Where(c => c.Status == IdentityResolutionStatus.Resolved).Select(c => c.ComparisonKey).Distinct(StringComparer.OrdinalIgnoreCase).ToArray();
                bool globallyUnique = keys.Length == 1 && side.Snapshot.Components.Where(c => c.Record.ComponentType == 80 &&
                    StringComparer.OrdinalIgnoreCase.Equals(c.ComparisonKey, keys[0])).Select(c => c.Record.ObjectId).Distinct().Count() == 1;
                if (matches.Length > 0) parents[id] = matches.All(c => c.Status == IdentityResolutionStatus.Resolved) && globallyUnique
                    ? new LookupEvidence { Status = "Unique", Key = keys[0], Reason = "Reused completed AppModule membership identity" }
                    : new LookupEvidence { Status = "Ambiguous", Reason = "Completed parent identity is uncertain/conflicting; not bypassed" };
            }
            // GUID-valued relationship fields are accepted only when their referenced attribute is the validated primary key.
            bool needsSchema = ids.Any(id => !parents.ContainsKey(id)) || rows.Any(r => r.GuidLinks.Values.Any(v => v.Item1 == "appmodule"));
            EntityMetadata parentSchema = needsSchema ? Schema(service, "appmodule", side, token) : null;
            if (parentSchema != null)
            {
                foreach (var row in rows)
                    foreach (var link in row.GuidLinks.Values.Where(v => v.Item1 == "appmodule"))
                        if (link.Item2 == parentSchema.PrimaryIdAttribute) row.ParentIds.Add(link.Item3);
                        else row.Context.Add("Parent link references " + link.Item2 + "; alternate/local ID not guessed as primary ID");
                var missing = rows.SelectMany(r => r.ParentIds).Distinct().Where(id => !parents.ContainsKey(id)).OrderBy(id => id).ToArray();
                bool valid = parentSchema.Attributes.Any(a => a.LogicalName == parentSchema.PrimaryIdAttribute && a.IsValidForRead == true && a.AttributeType == AttributeTypeCode.Uniqueidentifier) &&
                    parentSchema.Attributes.Any(a => a.LogicalName == "uniquename" && a.IsValidForRead == true && a.AttributeType == AttributeTypeCode.String);
                if (valid && missing.Length > 0)
                {
                    var data = Retrieve(service, "appmodule", parentSchema.PrimaryIdAttribute, new[] { parentSchema.PrimaryIdAttribute, "uniquename" }, missing, side, token, progress);
                    foreach (var id in missing)
                    {
                        var result = data[id]; var key = result.Row?.GetAttributeValue<object>("uniquename") as string;
                        parents[id] = result.Status != "Unique" ? result : string.IsNullOrWhiteSpace(key)
                            ? new LookupEvidence { Status = "Incomplete", Reason = "Blank parent appmodule.uniquename" }
                            : new LookupEvidence { Status = "Unique", Key = key.Trim(), Reason = "Exact parent primary-ID correlation; appmodule.uniquename" };
                    }
                }
            }
            foreach (var group in parents.Where(p => p.Value.Key != null).GroupBy(p => p.Value.Key, StringComparer.OrdinalIgnoreCase).Where(g => g.Count() > 1))
                foreach (var parent in group) { parent.Value.Status = "Ambiguous"; parent.Value.Key = null; parent.Value.Reason = "Distinct parent records share uniquename"; }
            foreach (var row in rows)
            {
                token.ThrowIfCancellationRequested();
                row.ParentStatus = row.RelationshipAmbiguous || row.ParentIds.Count > 1 ? "Ambiguous" : "Incomplete";
                if (!row.RelationshipAmbiguous && row.ParentIds.Count == 1 && parents.TryGetValue(row.ParentIds.Single(), out var parent))
                { row.ParentStatus = parent.Status; row.ParentKey = parent.Key; row.Context.Add(parent.Reason); }
                var referenceKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                foreach (var audit in row.References.Keys.Where(AuditReferences.Contains))
                    row.Context.Add(audit + " retained as audit evidence only; excluded from Candidate A referenced-component identity");
                foreach (var reference in row.References.Where(r => r.Value.LogicalName != "appmodule" &&
                    !AuditReferences.Contains(r.Key)))
                {
                    var lookup = metadata.Attributes.OfType<LookupAttributeMetadata>().SingleOrDefault(a => a.LogicalName == reference.Key);
                    if (lookup?.Targets?.Contains(reference.Value.LogicalName) != true) { row.ReferenceUncertain = true; continue; }
                    if (reference.Key == "canvasappid" && reference.Value.LogicalName == "canvasapp")
                    { row.CanvasReferenceId = reference.Value.Id; continue; }
                    ResolveSnapshotReference(side, row, reference.Value.Id, reference.Value.LogicalName, null, referenceKeys);
                }
                foreach (var entry in row.GuidLinks.Where(v => v.Value.Item1 != "appmodule" && !AuditReferences.Contains(v.Key)))
                {
                    var link = entry.Value;
                    if (entry.Key == "canvasappid" && link.Item1 == "canvasapp" && link.Item2 == "canvasappid")
                    { row.CanvasReferenceId = link.Item3; continue; }
                    // Only snapshot identities with the corresponding known PK are eligible; no local alternate-ID guesses.
                    if (PrimaryFor(link.Item1) == link.Item2) ResolveSnapshotReference(side, row, link.Item3, link.Item1, null, referenceKeys);
                    else { row.ReferenceUncertain = true; row.Context.Add("Reference alternate key unavailable as verified portable identity"); }
                }
                if (row.Get("objectid") != null && Guid.TryParse(row.Get("objectid"), out var objectId) && objectId != Guid.Empty &&
                    !row.GuidLinks.ContainsKey("objectid"))
                {
                    row.ReferenceUncertain = true;
                    row.Context.Add("Untyped objectid/componenttype values are audit evidence only; no referenced table/identity relationship assumed");
                }
                row.OtherReferenceKeys.UnionWith(referenceKeys);
                row.ReferenceStatus = row.CanvasReferenceId.HasValue || row.ReferenceUncertain || referenceKeys.Count > 1 ? "AmbiguousOrIncomplete" : referenceKeys.Count == 1 ? "Unique" : "Incomplete";
                row.ReferenceKey = row.ReferenceStatus == "Unique" ? referenceKeys.Single() : null;
                string unique = row.Get("uniquename") ?? row.Get("schemaname");
                UpdateCandidateA(row);
                if (row.CriticalComplete && row.ParentStatus == "Unique" && !string.IsNullOrWhiteSpace(unique))
                    row.CandidateB = "appelement-candidate-b:" + Type31EvidenceCollector.Frame(row.ParentKey, unique.Trim());
            }
        }

        private static void ResolveSnapshotReference(AppElementSideEvidence side, AppElementRecordEvidence row, Guid id, string entity,
            int? rawType, HashSet<string> keys)
        {
            int? type = rawType ?? RawTypeFor(entity);
            if (!type.HasValue && entity != "connectionreference") { row.ReferenceUncertain = true; row.Context.Add("Reference " + entity + " has no already-verified snapshot identity mapping"); return; }
            var matches = side.Snapshot.Components.Where(c => c.Record.ObjectId == id && (type.HasValue ? c.Record.ComponentType == type :
                c.SemanticKind == ComponentSemanticKinds.ConnectionReference)).ToArray();
            var known = matches.Where(c => c.Status == IdentityResolutionStatus.Resolved && !string.IsNullOrWhiteSpace(c.SemanticKind) && c.InventoryAbsencePolicy == InventoryAbsencePolicy.CompleteInventory)
                .Select(c => Type31EvidenceCollector.Frame(c.SemanticKind, c.ComparisonKey)).Distinct(StringComparer.OrdinalIgnoreCase).ToArray();
            bool globallyUnique = known.Length == 1 && side.Snapshot.Components.Where(c => c.Status == IdentityResolutionStatus.Resolved && !string.IsNullOrWhiteSpace(c.SemanticKind) &&
                StringComparer.OrdinalIgnoreCase.Equals(Type31EvidenceCollector.Frame(c.SemanticKind, c.ComparisonKey), known[0]))
                .Select(c => c.Record.ObjectId).Distinct().Count() == 1;
            if (matches.Length == 0 || matches.Any(c => c.Status != IdentityResolutionStatus.Resolved || c.InventoryAbsencePolicy != InventoryAbsencePolicy.CompleteInventory) || !globallyUnique)
            { row.ReferenceUncertain = true; row.Context.Add("Referenced component identity missing/ambiguous/incomplete in completed snapshot; no fallback"); }
            else { keys.Add(known[0]); row.Context.Add("Reference identity reused from completed snapshot; rawType=" + type); }
        }

        private static int? RawTypeFor(string entity)
        {
            switch (entity) { case "webresource": return 61; case "appmodule": return 80; case "savedquery": return 26;
                case "workflow": return 29; case "savedqueryvisualization": return 59; case "systemform": return 60;
                case "sitemap": return 62; case "pluginassembly": return 91; case "sdkmessageprocessingstep": return 92;
                case "environmentvariabledefinition": return 380; default: return null; }
        }
        private static string PrimaryFor(string entity) => RawTypeFor(entity).HasValue ? entity == "systemform" ? "formid" : entity + "id" : null;

        private static void UpdateCandidateA(AppElementRecordEvidence row)
        {
            // A proven Canvas App reference satisfies only the component-reference gate; it does not invent an element type.
            row.CandidateA = row.Status == "Unique" && row.CriticalComplete && row.ParentIdentityComplete &&
                row.ElementTypeComplete && row.ReferencedComponentIdentityComplete
                ? "appelement-candidate-a:" + Type31EvidenceCollector.Frame(row.ParentKey, row.ElementType, row.ReferenceKey) : null;
        }

        private static void CompleteCanvasReferences(IOrganizationService sourceService, IOrganizationService targetService,
            AppElementEvidenceReport report, CanvasAppEvidenceReport cached, CancellationToken token, Action<string> progress)
        {
            var sides = new[] { report.Source, report.Target }; var services = new[] { sourceService, targetService };
            var cachedSides = new[] { cached?.Source, cached?.Target };
            for (int i = 0; i < sides.Length; i++)
            {
                token.ThrowIfCancellationRequested(); var side = sides[i]; var prior = cachedSides[i];
                // Cache lifetime is one completed membership presentation. Do not reuse data across organizations or snapshots.
                if (prior != null && !ReferenceEquals(prior.Snapshot, side.Snapshot)) prior = null;
                var refs = side.Rows.Values.Where(r => r.Status == "Unique" && r.CanvasReferenceId.HasValue)
                    .Select(r => r.CanvasReferenceId.Value).Distinct().OrderBy(id => id).ToArray();
                if (refs.Length == 0) continue;
                var memberIds = side.Snapshot.Components.Where(c => c.Record.ComponentType == 300 && c.Record.ObjectId.HasValue)
                    .Select(c => c.Record.ObjectId.Value).ToHashSet();
                foreach (var id in refs.Where(memberIds.Contains))
                {
                    CanvasAppRecordEvidence row = null; prior?.Rows.TryGetValue(id, out row);
                    if (row != null) side.CanvasReferenceRows[id] = row; // Read-only reuse; Analyze never mutates this cached report.
                    else side.CanvasReferenceRows[id] = new CanvasAppRecordEvidence { ObjectId = id, Status = "Incomplete",
                        Reason = "Capture Type 300 evidence first for this completed comparison; member evidence is not guessed or re-queried" };
                }
                var dependencyIds = refs.Where(id => !memberIds.Contains(id)).ToArray();
                side.CanvasDependencies = Type300EvidenceCollector.ReadReferenced(services[i], side.Snapshot, side.Version,
                    dependencyIds, prior?.Metadata, token, progress);
                foreach (var row in side.CanvasDependencies.Rows) side.CanvasReferenceRows[row.Key] = row.Value;
                side.Requests.AddRange(side.CanvasDependencies.Requests);
                side.RetrievalDiagnostics.AddRange(side.CanvasDependencies.RetrievalDiagnostics);
                side.Pages.AddRange(side.CanvasDependencies.Pages);
                side.Schema.AddRange(side.CanvasDependencies.Schema); side.Relationships.AddRange(side.CanvasDependencies.Relationships);
                side.CanvasMemberIds.UnionWith(memberIds);
            }
            foreach (var side in sides)
            {
                var other = side == report.Source ? report.Target : report.Source;
                foreach (var row in side.Rows.Values.Where(r => r.CanvasReferenceId.HasValue))
                {
                    token.ThrowIfCancellationRequested();
                    var state = "Incomplete"; CanvasAppRecordEvidence canvas = null;
                    side.CanvasReferenceRows.TryGetValue(row.CanvasReferenceId.Value, out canvas);
                    if (canvas?.Status == "Missing") state = "Missing";
                    else if (canvas?.Status == "Faulted") state = "Faulted";
                    else if (canvas?.Status == "Duplicate" || canvas?.DuplicateA == true) state = "Ambiguous";
                    else if (canvas?.Status == "Unique" && canvas.PrimaryId == row.CanvasReferenceId && canvas.ObjectId == row.CanvasReferenceId &&
                        canvas.CriticalComplete && canvas.CandidateA != null)
                    {
                        var local = side.CanvasReferenceRows.Values.Where(c => StringComparer.OrdinalIgnoreCase.Equals(c.CandidateA, canvas.CandidateA)).ToArray();
                        var opposite = other.CanvasReferenceRows.Values.Where(c => StringComparer.OrdinalIgnoreCase.Equals(c.CandidateA, canvas.CandidateA)).ToArray();
                        if (local.Length != 1 || opposite.Length > 1 || opposite.Any(c => c.DuplicateA || c.Status == "Duplicate")) state = "Ambiguous";
                        else if (opposite.Length == 1 && opposite[0].Status == "Unique" && opposite[0].PrimaryId == opposite[0].ObjectId && opposite[0].CriticalComplete)
                            state = side.CanvasMemberIds.Contains(row.CanvasReferenceId.Value) ? "VerifiedType300SemanticIdentity" : "VerifiedDependencyOnlyCanvasAppIdentity";
                    }
                    row.CanvasDependencyState = state;
                    row.Context.Add("Canvas App dependency=" + state + "; localReference=" + row.CanvasReferenceId + "; backingPrimaryId=" + canvas?.PrimaryId +
                        "; CandidateA=" + canvas?.CandidateA + "; CandidateB=" + canvas?.CandidateB + "; evidence only; no production identity/absence proof");
                    row.CanvasCandidateA = canvas?.CandidateA; row.CanvasCandidateB = canvas?.CandidateB;
                    if (!state.StartsWith("Verified", StringComparison.Ordinal)) continue;
                    var keys = new HashSet<string>(row.OtherReferenceKeys, StringComparer.OrdinalIgnoreCase) { Type31EvidenceCollector.Frame("canvasapp-evidence-only", canvas.CandidateA) };
                    row.ReferenceStatus = row.ReferenceUncertain || keys.Count != 1 ? "AmbiguousOrIncomplete" : "Unique";
                    row.ReferenceKey = row.ReferenceStatus == "Unique" ? keys.Single() : null;
                    UpdateCandidateA(row);
                }
            }
        }

        private static Dictionary<Guid, LookupEvidence> RetrieveAppElements(IOrganizationService service, string primary, string[] columns,
            Guid[] ids, AppElementSideEvidence side, CancellationToken token, Action<string> progress)
        {
            var results = new Dictionary<Guid, LookupEvidence>();
            // Retries never widen the ID filter. Only service faults trigger isolation, not uncertain correlations.
            for (int offset = 0; offset < ids.Length; offset += BatchSize)
            {
                var batch = ids.Skip(offset).Take(BatchSize).ToArray();
                side.RetrievalDiagnostics.Add("Full AppElement attempt; requestedIds=" + batch.Length + "; columns=[" + string.Join(",", columns) + "]");
                var full = Retrieve(service, "appelement", primary, columns, batch, side, token, progress);
                if (!full.Values.Any(r => r.Status == "Faulted"))
                { foreach (var id in batch) results[id] = full[id]; continue; }

                var minimal = new[] { primary }.Concat(MinimalColumns.Where(columns.Contains)).Distinct(StringComparer.Ordinal)
                    .OrderBy(c => c, StringComparer.Ordinal).ToArray();
                foreach (var excluded in MinimalColumns.Except(minimal))
                    side.RetrievalDiagnostics.Add("Minimal column excluded by existing readable/shape metadata validation: " + excluded);
                side.RetrievalDiagnostics.Add("Minimal AppElement retry; requestedIds=" + batch.Length + "; columns=[" + string.Join(",", minimal) + "]");
                var current = Retrieve(service, "appelement", primary, minimal, batch, side, token, progress);
                int remaining = MaxIsolationGroups;
                if (current.Values.Any(r => r.Status == "Faulted"))
                {
                    side.RetrievalDiagnostics.Add("Minimal group faulted; proving primary correlation before bounded column isolation");
                    current = Retrieve(service, "appelement", primary, new[] { primary }, batch, side, token, progress);
                    if (current.Values.Any(r => r.Status == "Unique"))
                        Isolate(service, primary, minimal.Where(c => c != primary).ToArray(), batch, current, true, side, token, progress, ref remaining);
                    else side.RetrievalDiagnostics.Add("Primary-only correlation did not succeed; no optional or parent fallback");
                }
                // Optional/hash-only data, including publishconfiguration, cannot block primary correlation.
                if (current.Values.Any(r => r.Status == "Unique" && r.CriticalComplete))
                    Isolate(service, primary, columns.Except(minimal).OrderBy(c => c, StringComparer.Ordinal).ToArray(), batch, current, false, side, token, progress, ref remaining);
                else side.RetrievalDiagnostics.Add("Optional fields deferred/excluded=[" + string.Join(",", columns.Except(minimal)) +
                    "]: identity-critical evidence did not fully succeed; primary/partial critical evidence retained");
                foreach (var id in batch) results[id] = current[id];
            }
            return results;
        }

        private static void Isolate(IOrganizationService service, string primary, string[] fields, Guid[] batch,
            Dictionary<Guid, LookupEvidence> current, bool critical, AppElementSideEvidence side, CancellationToken token,
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
                var attempt = Retrieve(service, "appelement", primary, new[] { primary }.Concat(group).ToArray(), batch, side, token, progress);
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
                        existing.Status = found.Status == "Duplicate" ? "Duplicate" : "Incomplete";
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
            Guid[] ids, AppElementSideEvidence side, CancellationToken token, Action<string> progress)
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
            internal string Status, Reason, Key; internal Entity Row; internal bool CriticalComplete = true;
            internal HashSet<string> Columns = new HashSet<string>(StringComparer.Ordinal);
        }
        private static bool ValidName(string value) => !string.IsNullOrWhiteSpace(value) && value.All(c => char.IsLetterOrDigit(c) || c == '_');
        private static string Format(object raw)
        {
            if (raw == null) return "";
            if (raw is EntityReference) { var r = (EntityReference)raw; return ValidName(r.LogicalName) ? r.LogicalName + ":" + r.Id.ToString("D") : null; }
            if (raw is OptionSetValue) return ((OptionSetValue)raw).Value.ToString(CultureInfo.InvariantCulture);
            if (raw is string || raw is int || raw is bool || raw is Guid) return Convert.ToString(raw, CultureInfo.InvariantCulture)?.Trim();
            return null;
        }
    }

    internal sealed class AppElementSideEvidence
    {
        internal MembershipSnapshot Snapshot; internal string Version, PrimaryId, SchemaFailure;
        internal readonly List<ComponentIdentity> Raw = new List<ComponentIdentity>();
        internal readonly SortedDictionary<Guid, AppElementRecordEvidence> Rows = new SortedDictionary<Guid, AppElementRecordEvidence>();
        internal readonly List<string> Schema = new List<string>(), Relationships = new List<string>(), Requests = new List<string>(), Pages = new List<string>();
        internal readonly List<string> RetrievalDiagnostics = new List<string>();
        internal readonly List<string> TypeSchema = new List<string>(), TypeRegistration = new List<string>();
        internal readonly HashSet<string> TypeRequestedColumns = new HashSet<string>(StringComparer.Ordinal);
        internal readonly Dictionary<string, EntityMetadata> TypeMetadataCache = new Dictionary<string, EntityMetadata>(StringComparer.OrdinalIgnoreCase);
        internal CanvasAppSideEvidence CanvasDependencies;
        internal readonly SortedDictionary<Guid, CanvasAppRecordEvidence> CanvasReferenceRows = new SortedDictionary<Guid, CanvasAppRecordEvidence>();
        internal readonly HashSet<Guid> CanvasMemberIds = new HashSet<Guid>();
        internal bool Complete => Raw.All(r => r.Record.ObjectId.HasValue && r.Record.ObjectId != Guid.Empty) && Rows.Values.All(r => r.Status == "Unique" && r.CandidateA != null && !r.DuplicateA);
    }
    internal sealed class AppElementRecordEvidence
    {
        internal Guid ObjectId; internal Guid? PrimaryId; internal bool? Managed;
        internal string Status, Reason, ParentKey, ParentStatus, ReferenceKey, ReferenceStatus, CandidateA, CandidateB;
        internal bool DuplicateA, DuplicateB, RelationshipAmbiguous, ReferenceUncertain;
        internal bool CriticalComplete = true;
        internal Guid? CanvasReferenceId;
        internal string CanvasDependencyState = "Incomplete", CanvasCandidateA, CanvasCandidateB;
        internal readonly HashSet<string> OtherReferenceKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        internal readonly HashSet<Guid> ParentIds = new HashSet<Guid>();
        internal readonly SortedDictionary<string, string> Fields = new SortedDictionary<string, string>(StringComparer.Ordinal);
        internal readonly SortedDictionary<string, Type31ContentFingerprint> Content = new SortedDictionary<string, Type31ContentFingerprint>(StringComparer.Ordinal);
        internal readonly Dictionary<string, EntityReference> References = new Dictionary<string, EntityReference>();
        internal readonly Dictionary<string, Tuple<string, string, Guid>> GuidLinks = new Dictionary<string, Tuple<string, string, Guid>>();
        internal readonly List<string> Context = new List<string>();
        internal readonly HashSet<string> RuntimeColumns = new HashSet<string>(StringComparer.Ordinal);
        internal readonly Dictionary<string, string> TypeValues = new Dictionary<string, string>(StringComparer.Ordinal);
        internal readonly List<string> TypeAnalysis = new List<string>();
        internal bool IndependentTypeComplete;
        internal string IndependentTypeStatus = "Incomplete";
        internal string Get(string field) => Fields.TryGetValue(field, out var value) ? value : null;
        internal string ElementTypeField => new[] { "elementtype", "componenttype", "type", "objectidtype" }
            .FirstOrDefault(field => !string.IsNullOrWhiteSpace(Get(field)));
        internal string ElementType => ElementTypeField == null ? null : Get(ElementTypeField).Trim();
        internal bool ParentIdentityComplete => ParentStatus == "Unique" && !string.IsNullOrWhiteSpace(ParentKey);
        internal bool ElementTypeComplete => ElementType != null;
        internal bool ReferencedComponentIdentityComplete => ReferenceStatus == "Unique" && !string.IsNullOrWhiteSpace(ReferenceKey);
        internal bool CompleteA => CandidateA != null && !DuplicateA;
        internal string CandidateABlockingReason
        {
            get
            {
                var reasons = new List<string>();
                if (Status != "Unique") reasons.Add("Backing AppElement correlation=" + Status + ": " + Reason);
                if (!CriticalComplete) reasons.Add("Identity-critical AppElement fields unavailable");
                if (!ParentIdentityComplete) reasons.Add("Parent AppModule identity incomplete; ParentStatus=" + ParentStatus);
                if (!ElementTypeComplete) reasons.Add("Element/component type unavailable: elementtype, componenttype, type and objectidtype are blank or unavailable; Canvas App relationship does not supply an element type");
                if (!ReferencedComponentIdentityComplete) reasons.Add("Referenced component identity incomplete; ReferenceStatus=" + ReferenceStatus +
                    "; CanvasDependencyState=" + CanvasDependencyState + "; AdditionalReferenceUncertain=" + ReferenceUncertain);
                if (DuplicateA) reasons.Add("Duplicate canonical Candidate A");
                return reasons.Count == 0 ? "None" : string.Join("; ", reasons);
            }
        }
        internal string Evidence(string field) => Content.TryGetValue(field, out var hash) ? hash.Known ? hash.Evidence : null : Get(field);
    }
    internal sealed class AppElementPairEvidence
    {
        internal AppElementRecordEvidence Source, Target; internal string Outcome, Basis;
        internal readonly HashSet<string> Categories = new HashSet<string>(StringComparer.Ordinal);
    }
    internal sealed class AppElementEvidenceReport
    {
        internal AppElementSideEvidence Source, Target;
        internal readonly List<AppElementPairEvidence> Pairs = new List<AppElementPairEvidence>();
        internal void Analyze(CancellationToken token)
        {
            // Pair exclusively through complete unique Candidate A. Neither GUID overlap nor B repairs A.
            var keys = Source.Rows.Values.Concat(Target.Rows.Values).Where(r => r.CandidateA != null)
                .Select(r => r.CandidateA).Distinct(StringComparer.OrdinalIgnoreCase).OrderBy(k => k, StringComparer.OrdinalIgnoreCase);
            foreach (var key in keys)
            {
                token.ThrowIfCancellationRequested();
                var left = Source.Rows.Values.Where(r => StringComparer.OrdinalIgnoreCase.Equals(r.CandidateA, key)).ToArray();
                var right = Target.Rows.Values.Where(r => StringComparer.OrdinalIgnoreCase.Equals(r.CandidateA, key)).ToArray();
                if (left.Length > 1 || right.Length > 1)
                {
                    foreach (var row in left) Pairs.Add(new AppElementPairEvidence { Source = row, Outcome = "Ambiguous", Basis = "Candidate A collision; Candidate B cannot repair it" });
                    foreach (var row in right) Pairs.Add(new AppElementPairEvidence { Target = row, Outcome = "Ambiguous", Basis = "Candidate A collision; Candidate B cannot repair it" });
                }
                else Pairs.Add(new AppElementPairEvidence { Source = left.SingleOrDefault(), Target = right.SingleOrDefault(),
                    Outcome = left.Length == 1 && right.Length == 1 ? "SemanticPair" : "OneSidedEvidence",
                    Basis = "Unique Candidate A evidence only; no membership/absence finding" });
            }
            foreach (var side in new[] { Source, Target })
            {
                foreach (var row in side.Rows.Values.Where(r => r.CandidateA == null))
                    Pairs.Add(new AppElementPairEvidence { Source = side == Source ? row : null, Target = side == Target ? row : null,
                        Outcome = row.Status == "Duplicate" || row.ParentStatus == "Ambiguous" || row.CanvasDependencyState == "Ambiguous" ? "Ambiguous" : "Incomplete", Basis = row.Reason });
                foreach (var raw in side.Raw.Where(r => !r.Record.ObjectId.HasValue || r.Record.ObjectId == Guid.Empty))
                    Pairs.Add(new AppElementPairEvidence { Outcome = "Incomplete", Basis = (side == Source ? "Source" : "Target") + " blank ObjectId; no query/candidate" });
            }
            foreach (var pair in Pairs)
            {
                token.ThrowIfCancellationRequested(); pair.Categories.Add(pair.Outcome);
                if (pair.Outcome != "SemanticPair") continue;
                var left = pair.Source; var right = pair.Target;
                pair.Categories.Add(left.PrimaryId == right.PrimaryId ? "SamePrimaryId" : "DifferentPrimaryId");
                var uniqueFields = left.Fields.Keys.Union(right.Fields.Keys).Where(f => f.IndexOf("unique", StringComparison.OrdinalIgnoreCase) >= 0 &&
                    Guid.TryParse(left.Get(f), out var l) && l != Guid.Empty && Guid.TryParse(right.Get(f), out var r) && r != Guid.Empty).ToArray();
                if (uniqueFields.Length > 0) pair.Categories.Add(uniqueFields.Any(f => !StringComparer.OrdinalIgnoreCase.Equals(left.Get(f), right.Get(f))) ? "DifferentUniqueId" : "SameUniqueId");
                var contentFields = left.Content.Keys.Union(right.Content.Keys).ToArray();
                if (contentFields.Any(f => left.Content.ContainsKey(f) && right.Content.ContainsKey(f) && left.Content[f].Known && right.Content[f].Known && left.Content[f].Sha256 != right.Content[f].Sha256))
                    pair.Categories.Add("DifferentDefinition");
                else if (contentFields.Length > 0 && contentFields.All(f => left.Content.ContainsKey(f) && right.Content.ContainsKey(f) && left.Content[f].Known && right.Content[f].Known)) pair.Categories.Add("SameDefinition");
                if (left.Managed.HasValue && right.Managed.HasValue && left.Managed != right.Managed)
                { pair.Categories.Add("ManagedTransition"); if (left.Managed == false && right.Managed == true) pair.Categories.Add("UnmanagedToManaged"); }
            }
        }
        private static void Line(StringBuilder text, params object[] values) => text.AppendLine(string.Join("\t", values.Select(Type31EvidenceCollector.Safe)));
        internal string Build()
        {
            var text = new StringBuilder();
            text.AppendLine("TYPE 10072 APPELEMENT EVIDENCE - DEBUG ONLY");
            text.AppendLine("Type 10072 remains Unsupported / Indeterminate. No production identity, matching, definition contract or absence inference. No display-name/GUID/content-hash fallback.");
            text.AppendLine("Candidate A hypothesis: parent appmodule.uniquename + element/component type + verified snapshot identity or independently paired Canvas App diagnostic Candidate A. Candidate B hypothesis: parent appmodule.uniquename + element uniquename/schemaname. Trim + ordinal case-insensitive; B never repairs A ambiguity. Neither is approved portable identity.");
            text.AppendLine("RAW TYPE 10072 MEMBERSHIP");
            foreach (var side in new[] { Source, Target })
            {
                var label = side == Source ? "Source" : "Target";
                Line(text, label, side.Snapshot.Environment.DisplayName, side.Snapshot.SolutionUniqueName, side.Version,
                    "SnapshotUtc=" + side.Snapshot.CapturedAt.UtcDateTime.ToString("O"), "raw=" + side.Raw.Count,
                    "distinctNonblankIds=" + side.Rows.Count, "blankObjectIds=" + side.Raw.Count(r => !r.Record.ObjectId.HasValue || r.Record.ObjectId == Guid.Empty));
                foreach (var raw in side.Raw.OrderBy(r => r.Record.SolutionComponentId))
                    Line(text, label, raw.Record.SolutionComponentId, raw.Record.ObjectId, raw.Status, raw.SemanticKind, raw.Diagnostic, "ExistingPortableKey=" + raw.ComparisonKey);
            }
            text.AppendLine("BACKING APPELEMENT CORRELATION");
            foreach (var side in new[] { Source, Target })
            {
                var label = side == Source ? "Source" : "Target";
                Line(text, label, "MetadataPrimaryId=" + side.PrimaryId);
                foreach (var evidence in side.RetrievalDiagnostics) Line(text, label, "RetrievalDiagnostic", evidence);
                foreach (var page in side.Pages) Line(text, label, page);
                foreach (var row in side.Rows.Values) Line(text, label, row.ObjectId, row.PrimaryId,
                    "ObjectIdEqualsPrimaryId=" + (row.PrimaryId.HasValue && row.PrimaryId == row.ObjectId), row.Status, row.Reason);
            }
            text.AppendLine("READABLE / UNAVAILABLE SCHEMA");
            foreach (var side in new[] { Source, Target }) foreach (var schema in side.Schema) Line(text, side == Source ? "Source" : "Target", schema);
            text.AppendLine("PARENT / REFERENCE RELATIONSHIP DISCOVERY");
            foreach (var side in new[] { Source, Target })
            {
                var label = side == Source ? "Source" : "Target";
                foreach (var relationship in side.Relationships) Line(text, label, relationship);
                foreach (var row in side.Rows.Values)
                {
                    Line(text, label, row.ObjectId, "ParentStatus=" + row.ParentStatus, "ParentPortableEvidence=" + row.ParentKey,
                        "ReferenceStatus=" + row.ReferenceStatus, "ReferencePortableEvidence=" + row.ReferenceKey);
                    foreach (var evidence in row.Context) Line(text, label, row.ObjectId, evidence);
                }
            }
            text.AppendLine("ELEMENT-TYPE SOURCE METADATA / SELECTED-ROW VALUES");
            foreach (var side in new[] { Source, Target })
            {
                var label = side == Source ? "Source" : "Target";
                foreach (var item in side.TypeSchema) Line(text, label, item);
                foreach (var row in side.Rows.Values) foreach (var item in row.TypeAnalysis) Line(text, label, row.ObjectId, item);
            }
            text.AppendLine("REGISTERED-DEFINITION DISCRIMINATOR INVESTIGATION");
            foreach (var side in new[] { Source, Target }) foreach (var item in side.TypeRegistration) Line(text, side == Source ? "Source" : "Target", item);
            text.AppendLine("Relationship endpoints, catalog labels, Canvas App dependency identity and ObjectTypeCode agreement are not selected-row discriminator contracts. Incoming relationships are not scanned; no relationship-only type inference.");
            text.AppendLine("CANDIDATE A BEFORE / AFTER TYPE INVESTIGATION");
            foreach (var side in new[] { Source, Target }) foreach (var row in side.Rows.Values)
                Line(text, side == Source ? "Source" : "Target", row.ObjectId, "CandidateAConstructionUnchanged=True",
                    "ParentIdentityComplete=" + row.ParentIdentityComplete, "ElementTypeComplete=" + row.ElementTypeComplete,
                    "ReferencedComponentIdentityComplete=" + row.ReferencedComponentIdentityComplete, "CompleteA=" + row.CompleteA,
                    "IndependentElementTypeComplete=" + row.IndependentTypeComplete, "IndependentTypeStatus=" + row.IndependentTypeStatus);
            var typeRows = new[] { Source, Target }.SelectMany(s => s.Rows.Values).ToArray();
            bool typeEstablished = typeRows.Length > 0 && typeRows.All(r => r.Status == "Unique" && r.IndependentTypeComplete);
            text.AppendLine(typeEstablished ? "Element type independently established; candidate can advance to live validation. Option-metadata value evidence only; Candidate A/B are unchanged and subtype portability still requires review." :
                "No independent element-type source found; Type 10072 should remain unsupported. Blank/unavailable/ambiguous type evidence cannot be repaired by dependencies or a weaker candidate.");
            text.AppendLine("CANDIDATE IDENTITY ANALYSIS");
            foreach (var side in new[] { Source, Target }) foreach (var row in side.Rows.Values)
            {
                Line(text, side == Source ? "Source" : "Target", row.ObjectId, "CandidateA=" + (row.CandidateA ?? "Incomplete"),
                    "CandidateB=" + (row.CandidateB ?? "Incomplete"), "DuplicateA=" + row.DuplicateA, "DuplicateB=" + row.DuplicateB);
                foreach (var field in row.Fields.Keys.Union(row.Content.Keys).OrderBy(f => f, StringComparer.Ordinal))
                    Line(text, side == Source ? "Source" : "Target", row.ObjectId, field, row.Evidence(field) ?? "Unknown");
            }
            text.AppendLine("DUPLICATE / REPEATED ANALYSIS");
            foreach (var side in new[] { Source, Target })
            {
                var label = side == Source ? "Source" : "Target";
                foreach (var group in side.Raw.Where(r => r.Record.ObjectId.HasValue).GroupBy(r => r.Record.ObjectId).Where(g => g.Count() > 1))
                    Line(text, label, "RepeatedMembershipOnly", group.Key, "rawReferences=" + group.Count());
                Line(text, label, "CandidateACollisionGroups=" + side.Rows.Values.Where(r => r.DuplicateA).Select(r => r.CandidateA).Distinct(StringComparer.OrdinalIgnoreCase).Count(),
                    "CandidateBCollisionGroups=" + side.Rows.Values.Where(r => r.DuplicateB).Select(r => r.CandidateB).Distinct(StringComparer.OrdinalIgnoreCase).Count());
            }
            text.AppendLine("SOURCE / TARGET FIELD COMPARISON");
            foreach (var pair in Pairs.Where(p => p.Outcome == "SemanticPair"))
                foreach (var field in pair.Source.Fields.Keys.Union(pair.Target.Fields.Keys).Union(pair.Source.Content.Keys).Union(pair.Target.Content.Keys).OrderBy(f => f, StringComparer.Ordinal))
                {
                    string left = pair.Source.Evidence(field), right = pair.Target.Evidence(field);
                    Line(text, pair.Source.PrimaryId, pair.Target.PrimaryId, field, left ?? "Unknown", right ?? "Unknown",
                        left == null || right == null ? "InsufficientEvidence" : StringComparer.OrdinalIgnoreCase.Equals(left, right) ? "EqualObserved" : "DifferentObserved");
                }
            text.AppendLine("LIFECYCLE CORRELATION MATRIX");
            foreach (var pair in Pairs) Line(text, pair.Source?.PrimaryId, pair.Target?.PrimaryId, pair.Outcome, pair.Basis,
                "CandidateA=" + (pair.Source?.CandidateA ?? pair.Target?.CandidateA),
                string.Join(",", pair.Categories.OrderBy(c => c, StringComparer.Ordinal)));
            text.AppendLine("Counts are diagnostic backing pairs/observations; categories overlap. Repeated raw memberships do not inflate semantic pairs.");
            foreach (var outcome in new[] { "SemanticPair", "SamePrimaryId", "DifferentPrimaryId", "SameUniqueId", "DifferentUniqueId", "SameDefinition", "DifferentDefinition", "ManagedTransition", "UnmanagedToManaged", "OneSidedEvidence", "Ambiguous", "Incomplete" })
                Line(text, outcome, Pairs.Count(p => p.Categories.Contains(outcome)));
            text.AppendLine("DIFFERING PRIMARY-ID SEMANTIC PAIRS");
            var differing = Pairs.Where(p => p.Outcome == "SemanticPair" && p.Source.PrimaryId != p.Target.PrimaryId).ToArray();
            if (differing.Length == 0) text.AppendLine("No unique differing-primary-ID Candidate A pair observed.");
            foreach (var pair in differing) Line(text, Source.Snapshot.SolutionUniqueName, pair.Source.PrimaryId, pair.Target.PrimaryId,
                "Basis=Unique Candidate A only", pair.Source.CandidateA, "SourceManaged=" + pair.Source.Managed, "TargetManaged=" + pair.Target.Managed);
            text.AppendLine("DEPENDENCY-ONLY CANVAS APP BACKING CORRELATION");
            foreach (var side in new[] { Source, Target })
            {
                var label = side == Source ? "Source" : "Target";
                foreach (var app in side.CanvasReferenceRows.Values)
                    Line(text, label, app.ObjectId, "BackingPrimaryId=" + app.PrimaryId, "ExactLocalCorrelation=" + (app.PrimaryId == app.ObjectId),
                        app.Status, app.Reason, "EvidenceOrigin=" + (side.CanvasMemberIds.Contains(app.ObjectId) ? "CompletedType300Capture" : "DependencyOnlySelectedIds"));
            }
            text.AppendLine("DEPENDENCY CANVAS APP RUNTIME-READABLE / FAULTED COLUMNS");
            foreach (var side in new[] { Source, Target })
            {
                foreach (var item in side.CanvasDependencies?.RetrievalDiagnostics ?? Enumerable.Empty<string>()) Line(text, side == Source ? "Source" : "Target", item);
                foreach (var app in side.CanvasReferenceRows.Values) Line(text, side == Source ? "Source" : "Target", app.ObjectId,
                    "SucceededColumns=[" + string.Join(",", app.RuntimeColumns) + "]", "CriticalComplete=" + app.CriticalComplete);
            }
            text.AppendLine("DEPENDENCY CANVAS APP CANDIDATE A/B");
            foreach (var side in new[] { Source, Target }) foreach (var app in side.CanvasReferenceRows.Values)
            {
                Line(text, side == Source ? "Source" : "Target", app.ObjectId, "CandidateA=" + (app.CandidateA ?? "Incomplete"), "CandidateB=" + (app.CandidateB ?? "NotAvailable"));
                foreach (var field in app.Fields.Keys.Union(app.Content.Keys).OrderBy(f => f, StringComparer.Ordinal))
                    Line(text, side == Source ? "Source" : "Target", app.ObjectId, field, app.Evidence(field) ?? "Unavailable");
            }
            text.AppendLine("CROSS-ENVIRONMENT DEPENDENCY SEMANTIC PAIRING");
            foreach (var side in new[] { Source, Target }) foreach (var row in side.Rows.Values.Where(r => r.CanvasReferenceId.HasValue))
                Line(text, side == Source ? "Source" : "Target", row.ObjectId, row.CanvasReferenceId, row.CanvasDependencyState,
                    "CandidateA=" + (row.CanvasCandidateA ?? "Incomplete"), "CandidateB=" + (row.CanvasCandidateB ?? "NotAvailable"), "Evidence only; not production membership");
            text.AppendLine("UPDATED TYPE 10072 CANDIDATE A COMPLETENESS");
            foreach (var side in new[] { Source, Target }) foreach (var row in side.Rows.Values)
                Line(text, side == Source ? "Source" : "Target", row.ObjectId, "ParentIdentityComplete=" + row.ParentIdentityComplete,
                    "ElementTypeComplete=" + row.ElementTypeComplete, "ElementTypeField=" + row.ElementTypeField, "ElementType=" + row.ElementType,
                    "ReferencedComponentIdentityComplete=" + row.ReferencedComponentIdentityComplete, "CompleteA=" + row.CompleteA,
                    "BlockingReason=" + row.CandidateABlockingReason,
                    "DuplicateA=" + row.DuplicateA, "ParentStatus=" + row.ParentStatus, "ReferenceStatus=" + row.ReferenceStatus, row.CanvasDependencyState);
            text.AppendLine("UPDATED TYPE 10072 DIFFERING-PRIMARY-ID SEMANTIC PAIRS");
            foreach (var pair in differing) Line(text, pair.Source.ObjectId, pair.Target.ObjectId, pair.Source.CandidateA,
                "BothCompleteAndUnique=True; DifferentPrimaryIds=True; no GUID/B/hash pairing");
            Line(text, "UniqueDifferingPrimaryIdPairs=" + differing.Length, "CanvasCandidateACollisionGroups=" +
                new[] { Source, Target }.Sum(side => side.CanvasReferenceRows.Values.Where(r => r.CandidateA != null)
                    .GroupBy(r => r.CandidateA, StringComparer.OrdinalIgnoreCase).Count(g => g.Count() > 1 || g.Any(r => r.DuplicateA))));
            text.AppendLine("PORTABILITY ASSESSMENT");
            Line(text, "UniqueSemanticPairs=" + Pairs.Count(p => p.Outcome == "SemanticPair"), "DifferingPrimaryIdSemanticPairs=" + differing.Length,
                "SourceEvidenceComplete=" + Source.Complete, "TargetEvidenceComplete=" + Target.Complete);
            text.AppendLine("Observations only. Differing-ID semantic pairs, if present, support further investigation; collisions, uncertain parents/references and missing metadata cannot establish portable identity or absence. No production promotion.");
            text.AppendLine("EXACT REQUEST LEDGER");
            foreach (var side in new[] { Source, Target })
            {
                var label = side == Source ? "Source" : "Target";
                Line(text, label, "TotalReads=" + side.Requests.Count, "AdditionalWhoAmI=0", "Writes=0", "NormalMembershipEvidenceRequests=0");
                Line(text, label, "AppElementSchemaRequests=" + side.Requests.Count(r => r.StartsWith("Execute RetrieveEntity(appelement,", StringComparison.Ordinal)),
                    "TypeSourceMetadataQueries=" + side.Requests.Count(r => r.StartsWith("Execute RetrieveEntity(") && !r.StartsWith("Execute RetrieveEntity(appelement,") && !r.StartsWith("Execute RetrieveEntity(appmodule,") && !r.StartsWith("Execute RetrieveEntity(canvasapp,")),
                    "RegistrationQueries=" + side.Requests.Count(r => r.StartsWith("RetrieveMultiple solutioncomponentdefinition;")),
                    "AppElementQueries=" + side.Requests.Count(r => r.StartsWith("RetrieveMultiple appelement;", StringComparison.Ordinal)),
                    "ParentAppModuleSchemaRequests=" + side.Requests.Count(r => r.StartsWith("Execute RetrieveEntity(appmodule,", StringComparison.Ordinal)),
                    "ParentAppModuleQueries=" + side.Requests.Count(r => r.StartsWith("RetrieveMultiple appmodule;", StringComparison.Ordinal)),
                    "CanvasAppDependencySchemaRequests=" + side.Requests.Count(r => r.StartsWith("Execute RetrieveEntity(canvasapp,", StringComparison.Ordinal)),
                    "CanvasAppDependencyQueries=" + side.Requests.Count(r => r.StartsWith("RetrieveMultiple canvasapp;", StringComparison.Ordinal)),
                    "CanvasApp Type300 member rows/metadata reuse requires a completed capture from this exact snapshot; no member re-query");
                for (int i = 0; i < side.Requests.Count; i++) Line(text, label, i + 1, side.Requests[i]);
            }
            return text.ToString();
        }
    }
}
#endif
