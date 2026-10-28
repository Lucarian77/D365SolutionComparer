using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Text.RegularExpressions;
using D365SolutionComparer.Models.Membership;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Metadata.Query;
using Microsoft.Xrm.Sdk.Query;

namespace D365SolutionComparer.Services.Membership
{
    // One verified environment/operation only. Called by the existing worker resolver, never the UI.
    internal sealed class MetadataForArchivalResolver
    {
        internal const int ComponentType = 10139;
        private const string Kind = ComponentSemanticKinds.MetadataForArchival;
        private const string Table = "metadataforarchival";
        private const string Primary = "metadataforarchivalid";
        private readonly Dictionary<SolutionComponentRecord, Observation> observations =
            new Dictionary<SolutionComponentRecord, Observation>();
        private MembershipSnapshot owner;

        internal ComponentIdentity Resolve(DataverseReadContext context, SolutionComponentRecord record, CancellationToken token)
        {
            Read(context, new[] { record }, token); return observations[record].Identity;
        }

        internal MembershipSnapshot Apply(DataverseReadContext context, MembershipSnapshot snapshot, CancellationToken token)
        {
            token.ThrowIfCancellationRequested();
            if (snapshot.State != MembershipSnapshotState.Complete || !snapshot.Components.Any(i => i.Record.ComponentType == ComponentType)) return snapshot;
            if (snapshot.Environment.OrganizationId != context.Environment.OrganizationId)
                throw new InvalidOperationException("Archival metadata context belongs to another environment.");
            Read(context, snapshot.Components.Where(i => i.Record.ComponentType == ComponentType).Select(i => i.Record).ToArray(), token);
            owner = MembershipSnapshot.Complete(snapshot.Solution, snapshot.Components.Select(i => i.Record.ComponentType == ComponentType ? observations[i.Record].Identity : i), snapshot.CapturedAt);
            context.MetadataCache.MetadataForArchival = this;
            return owner;
        }

        internal bool TryInformation(MembershipSnapshot snapshot, ComponentIdentity identity, DisplayEnvironmentSide side,
            out MetadataForArchivalInformation information)
        {
            information = null; Observation value;
            if (!ReferenceEquals(owner, snapshot) || identity == null || identity.Record.ComponentType != ComponentType ||
                !observations.TryGetValue(identity.Record, out value) || !ReferenceEquals(value.Identity, identity)) return false;
            information = new MetadataForArchivalInformation(side, identity.Record, value.Fields); return true;
        }

        private void Read(DataverseReadContext context, IReadOnlyList<SolutionComponentRecord> records, CancellationToken token)
        {
            observations.Clear(); owner = null;
            foreach (var record in records) observations[record] = new Observation(record);
            if (records.Count == 0) return;
            try
            {
                token.ThrowIfCancellationRequested();
                var response = context.Execute(new RetrieveEntityRequest { LogicalName = Table,
                    EntityFilters = EntityFilters.Entity | EntityFilters.Attributes, RetrieveAsIfPublished = false }) as RetrieveEntityResponse;
                var schema = response?.EntityMetadata;
                var columns = Columns(schema);
                if (columns == null) { FailAll("Archival ID/name readable schema is not verified."); return; }
                var rows = new BatchedDiagnosticQueryReader(context).Read(records, Table, Primary, columns,
                    "Archival metadata batch is incomplete.", "Archival primary-key correlation is conflicting.",
                    error => "Archival metadata batch failed.", token);
                foreach (var item in observations.Values)
                {
                    token.ThrowIfCancellationRequested();
                    if (!item.Record.ObjectId.HasValue || item.Record.ObjectId == Guid.Empty) continue;
                    var correlation = rows.GetCorrelation(item.Record.ObjectId.Value);
                    if (correlation.Status != DiagnosticCorrelationStatus.Unique || correlation.Rows.Count != 1)
                    {
                        var availability = correlation.Status == DiagnosticCorrelationStatus.Duplicate ? DisplayAvailability.Ambiguous :
                            correlation.Failure == "Archival metadata batch failed." ? DisplayAvailability.RetrievalFailed : DisplayAvailability.Unavailable;
                        var fields = new[] { "Name", "Managed", "Archival availability", "Archival readiness", "Backing solution ID", "Component unique ID", "Related Table" };
                        var attributes = new[] { "name", "ismanaged", "isavailableforarchival", "isreadyforarchival", "solutionid", "componentidunique", "extensionofrecordid" };
                        for (int index = 0; index < fields.Length; index++)
                            if (columns.Contains(attributes[index])) item.Fields[fields[index]] = new ComponentDisplayValue(availability);
                        item.Identity = Unresolved(item.Record, "Archival backing record is not uniquely correlated.",
                            correlation.Status == DiagnosticCorrelationStatus.Duplicate);
                        continue;
                    }
                    var row = correlation.Rows[0];
                    if (row.Id != item.Record.ObjectId.Value || row.GetAttributeValue<Guid>(Primary) != row.Id || row.LogicalName != Table)
                    { item.Identity = Unresolved(item.Record, "Archival returned identity is inconsistent."); continue; }
                    item.Fields["Name"] = StringField(row, "name", columns);
                    item.Fields["Managed"] = BoolField(row, "ismanaged", columns);
                    item.Fields["Archival availability"] = BoolField(row, "isavailableforarchival", columns);
                    item.Fields["Archival readiness"] = BoolField(row, "isreadyforarchival", columns);
                    item.Fields["Backing solution ID"] = IdField(row, "solutionid", columns);
                    item.Fields["Component unique ID"] = IdField(row, "componentidunique", columns);
                    item.TableId = columns.Contains("extensionofrecordid") ? ReadId(row, "extensionofrecordid") : null;
                    item.Fields["Related Table"] = IdField(row, "extensionofrecordid", columns);
                }
                var tables = ReadTables(context, observations.Values.Where(i => i.TableId.HasValue).Select(i => i.TableId.Value).Distinct().ToArray(), token);
                foreach (var item in observations.Values)
                {
                    var name = item.Fields["Name"];
                    EntityMetadata table;
                    if (name.Availability != DisplayAvailability.Available || string.IsNullOrWhiteSpace(name.Value) ||
                        !Regex.IsMatch(name.Value, "^[A-Za-z][A-Za-z0-9_]*$") || !item.TableId.HasValue ||
                        !tables.TryGetValue(item.TableId.Value, out table) ||
                        !StringComparer.OrdinalIgnoreCase.Equals(name.Value, table.LogicalName)) continue;
                    item.Fields["Related Table"] = Observed(table.LogicalName + " (" + item.TableId + ")", "EntityMetadata.MetadataId/LogicalName + metadataforarchival.extensionofrecordid");
                    string own = name.Value.ToLowerInvariant(), related = table.LogicalName.ToLowerInvariant();
                    string key = "metadataforarchival:v1:" + own.Length + ":" + own + ":" + related.Length + ":" + related;
                    item.Identity = new ComponentIdentity(item.Record, IdentityResolutionStatus.Resolved, key,
                        "Archival own logical name and independently verified related Table identity; archival flags are not compared.", Kind, Kind);
                }
                // Different local records with the same own name are never paired arbitrarily.
                foreach (var group in observations.Values.Where(i => i.Fields["Name"].Availability == DisplayAvailability.Available &&
                    !string.IsNullOrWhiteSpace(i.Fields["Name"].Value)).GroupBy(i => i.Fields["Name"].Value, StringComparer.OrdinalIgnoreCase))
                    if (group.Select(i => i.Record.ObjectId).Distinct().Count() > 1)
                        foreach (var item in group) item.Identity = Unresolved(item.Record, "Archival logical name is ambiguous within this solution inventory.", true);
            }
            catch (OperationCanceledException) { throw; }
            catch (Exception) { token.ThrowIfCancellationRequested(); FailAll("Archival metadata could not be verified; no absence inference is established."); }
        }

        private static IReadOnlyList<string> Columns(EntityMetadata schema)
        {
            if (schema?.LogicalName != Table || schema.PrimaryIdAttribute != Primary || schema.PrimaryNameAttribute != "name") return null;
            var attributes = (schema.Attributes ?? new AttributeMetadata[0]).ToDictionary(a => a.LogicalName, StringComparer.OrdinalIgnoreCase);
            Func<string, AttributeTypeCode[], bool> readable = (name, types) => attributes.ContainsKey(name) &&
                attributes[name].IsValidForRead == true && attributes[name].IsSecured != true &&
                attributes[name].AttributeType.HasValue && types.Contains(attributes[name].AttributeType.Value);
            if (!readable(Primary, new[] { AttributeTypeCode.Uniqueidentifier }) || !readable("name", new[] { AttributeTypeCode.String })) return null;
            var columns = new List<string> { Primary, "name" };
            if (readable("extensionofrecordid", new[] { AttributeTypeCode.Uniqueidentifier })) columns.Add("extensionofrecordid");
            foreach (var name in new[] { "solutionid", "componentidunique" })
                if (readable(name, new[] { AttributeTypeCode.Uniqueidentifier, AttributeTypeCode.Lookup })) columns.Add(name);
            foreach (var name in new[] { "ismanaged", "isavailableforarchival", "isreadyforarchival" })
                if (readable(name, new[] { AttributeTypeCode.Boolean })) columns.Add(name);
            return columns.AsReadOnly();
        }

        private static Dictionary<Guid, EntityMetadata> ReadTables(DataverseReadContext context, Guid[] ids, CancellationToken token)
        {
            var result = new Dictionary<Guid, EntityMetadata>(); var pending = new List<Guid>();
            foreach (Guid id in ids)
            {
                EntityMetadata value;
                if (context.MetadataCache.TryGetEntity(id, out value) && value.MetadataId == id && !string.IsNullOrWhiteSpace(value.LogicalName)) result[id] = value;
                else pending.Add(id);
            }
            for (int offset = 0; offset < pending.Count; offset += 200)
            {
                token.ThrowIfCancellationRequested(); var batch = pending.Skip(offset).Take(200).ToArray();
                var filter = new MetadataFilterExpression(LogicalOperator.Or);
                foreach (var id in batch) filter.Conditions.Add(new MetadataConditionExpression("MetadataId", MetadataConditionOperator.Equals, id));
                var response = context.Execute(new RetrieveMetadataChangesRequest { Query = new EntityQueryExpression {
                    Properties = new MetadataPropertiesExpression("MetadataId", "LogicalName"), Criteria = filter } }) as RetrieveMetadataChangesResponse;
                var entities = response?.EntityMetadata;
                if (entities == null || entities.Any(e => !e.MetadataId.HasValue || !batch.Contains(e.MetadataId.Value) || string.IsNullOrWhiteSpace(e.LogicalName)) ||
                    entities.GroupBy(e => e.MetadataId).Any(g => g.Count() != 1) || entities.GroupBy(e => e.LogicalName, StringComparer.OrdinalIgnoreCase).Any(g => g.Count() != 1))
                    return new Dictionary<Guid, EntityMetadata>();
                foreach (var entity in entities) { result[entity.MetadataId.Value] = entity; context.MetadataCache.Store(entity); }
            }
            return result;
        }

        private static ComponentIdentity Unresolved(SolutionComponentRecord record, string reason, bool ambiguous = false) =>
            new ComponentIdentity(record, ambiguous ? IdentityResolutionStatus.Ambiguous : IdentityResolutionStatus.Unresolved,
                diagnostic: reason, componentTypeKey: Kind, semanticKind: Kind, blockerScope: ResolutionBlockerScope.SemanticKind);
        private void FailAll(string reason) { foreach (var value in observations.Values) value.Identity = Unresolved(value.Record, reason); }
        private static ComponentDisplayValue Observed(string value, string source) => new ComponentDisplayValue(DisplayAvailability.Available, value, source);
        private static ComponentDisplayValue Field(Entity row, string field, IReadOnlyList<string> columns, Func<object, string> convert)
        {
            if (!columns.Contains(field)) return new ComponentDisplayValue(DisplayAvailability.NotRetrieved);
            object raw;
            if (!row.Attributes.TryGetValue(field, out raw)) return new ComponentDisplayValue(DisplayAvailability.AttributeOmitted);
            if (raw == null) return Observed(null, Table + "." + field);
            var value = convert(raw);
            return value == null ? new ComponentDisplayValue(DisplayAvailability.Unavailable) : Observed(value, Table + "." + field);
        }
        private static ComponentDisplayValue StringField(Entity row, string field, IReadOnlyList<string> columns) => Field(row, field, columns, x => x as string);
        private static ComponentDisplayValue BoolField(Entity row, string field, IReadOnlyList<string> columns) => Field(row, field, columns, x => x is bool ? ((bool)x).ToString() : null);
        private static ComponentDisplayValue IdField(Entity row, string field, IReadOnlyList<string> columns) => Field(row, field, columns, x => x is Guid ? ((Guid)x).ToString("D") : x is EntityReference ? ((EntityReference)x).Id.ToString("D") : null);
        private static Guid? ReadId(Entity row, string field)
        {
            object value; if (!row.Attributes.TryGetValue(field, out value)) return null;
            Guid? id = value is Guid ? (Guid?)value : value is EntityReference ? ((EntityReference)value).Id : (Guid?)null;
            return id == Guid.Empty ? null : id;
        }
        private sealed class Observation
        {
            internal readonly SolutionComponentRecord Record;
            internal ComponentIdentity Identity;
            internal Guid? TableId;
            internal readonly Dictionary<string, ComponentDisplayValue> Fields = new Dictionary<string, ComponentDisplayValue>();
            internal Observation(SolutionComponentRecord record)
            {
                Record = record; Identity = Unresolved(record, "Archival logical identity is not established.");
                foreach (var name in new[] { "Name", "Related Table", "Managed", "Archival availability", "Archival readiness", "Backing solution ID", "Component unique ID" })
                    Fields[name] = new ComponentDisplayValue(DisplayAvailability.NotRetrieved);
            }
        }
    }
}
