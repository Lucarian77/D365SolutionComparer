using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Text.RegularExpressions;
using D365SolutionComparer.Models.Membership;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;

namespace D365SolutionComparer.Services.Membership
{
    // Optional worker-only local display evidence. Never changes identity, membership or definitions.
    internal sealed class MakerObjectDisplayReader
    {
        internal void Read(DataverseReadContext context, MembershipSnapshot owner, CancellationToken token)
        {
            if (context == null) throw new ArgumentNullException(nameof(context));
            context.MetadataCache.ClearMakerObjectDisplay();
            token.ThrowIfCancellationRequested();
            if (owner == null || owner.State != MembershipSnapshotState.Complete ||
                owner.Environment.OrganizationId != context.Environment.OrganizationId) return;
            var cache = context.MetadataCache;
            var values = new Dictionary<SolutionComponentRecord, MakerObjectDisplayEvidence>();
            var limitations = new List<string>();
            foreach (var identity in owner.Components)
            {
                token.ThrowIfCancellationRequested();
                if (!identity.Record.ObjectId.HasValue || identity.Record.ObjectId == Guid.Empty) continue;
                Guid id = identity.Record.ObjectId.Value;
                Entity row;
                if (identity.Record.ComponentType == 60 && cache.TryGetEntityRow("systemform", id, out row) &&
                    Exact(row, "systemform", "formid", id) && new[] { 0, 10, 13, 103 }.Contains(Option(row, "type") ?? -1))
                    values[identity.Record] = Value(identity, MakerObjectKind.Dashboard, id, row, "name", "uniquename",
                        dashboardFormType: Option(row, "type"));
                if (identity.Record.ComponentType == 29 && cache.WorkflowDefinitions.TryGetValue(id, out row) &&
                    row != null && StringComparer.OrdinalIgnoreCase.Equals(row.LogicalName, "workflow") && row.Id != Guid.Empty && Exact(row, "workflow", "workflowid", row.Id) &&
                    Option(row, "type") == 1 && new[] { 0, 1, 3, 4 }.Contains(Option(row, "category") ?? -1))
                {
                    Guid definition;
                    if (row.Id == id || cache.WorkflowDisplayAliases.TryGetValue(id, out definition) && definition == row.Id)
                        values[identity.Record] = Value(identity, MakerObjectKind.Process, row.Id, row, "name", "uniquename",
                            row.Id == id ? DisplayBackingCorrelationStatus.Unique : DisplayBackingCorrelationStatus.VerifiedDefinition,
                            processCategory: Option(row, "category"));
                }
                if (identity.Record.ComponentType == 300)
                {
                    DataverseComponentMetadataCache.NormalDisplayNameObservation canvas;
                    if (cache.TryGetNormalDisplayName(owner, identity, out canvas) &&
                        canvas.Correlation == DisplayBackingCorrelationStatus.Unique && canvas.CanvasAppType == 1)
                        values[identity.Record] = new MakerObjectDisplayEvidence(identity, MakerObjectKind.ComponentLibrary, id,
                            canvas.FriendlyName, canvas.InternalName, canvas.Correlation);
                }
            }
            ReadRegistered(context, owner, "aiskillconfig", MakerObjectKind.AiSkillConfig, values, limitations, token);
            ReadRegistered(context, owner, "maskingrule", MakerObjectKind.SecuredMaskingRule, values, limitations, token);
            token.ThrowIfCancellationRequested();
            cache.SealMakerObjectDisplay(owner, values, limitations);
        }

        private static void ReadRegistered(DataverseReadContext context, MembershipSnapshot owner, string family,
            MakerObjectKind kind, IDictionary<SolutionComponentRecord, MakerObjectDisplayEvidence> values,
            IList<string> limitations, CancellationToken token)
        {
            var members = owner.Components.Where(i => i.RegisteredDefinition != null &&
                Normalize(i.RegisteredDefinition.Name) == family).ToList();
            if (members.Count == 0) return;
            try
            {
                var tables = members.Select(i => i.RegisteredDefinition.PrimaryEntityName).Distinct(StringComparer.OrdinalIgnoreCase).ToArray();
                if (tables.Length != 1 || !Symbol(tables[0]) || members.Any(i =>
                    i.RegisteredDefinition.ObjectTypeCode != i.Record.ComponentType ||
                    i.SemanticKind != i.RegisteredDefinition.SemanticKind))
                { limitations.Add(family + ": registered backing mapping is not uniquely verified."); return; }
                var records = members.Where(i => i.Record.ObjectId.HasValue && i.Record.ObjectId != Guid.Empty).Select(i => i.Record).ToList();
                if (records.Count == 0) return;
                string table = tables[0];
                var retained = context.MetadataCache.RetainedEntities.Where(m =>
                    StringComparer.OrdinalIgnoreCase.Equals(m.LogicalName, table)).ToList();
                if (retained.Count > 1) { limitations.Add(family + ": backing schema is ambiguous."); return; }
                EntityMetadata schema = retained.SingleOrDefault();
                string primary, name;
                if (!ReadableSchema(schema, table, out primary, out name))
                {
                    token.ThrowIfCancellationRequested();
                    var response = context.Execute(new RetrieveEntityRequest { LogicalName = table,
                        EntityFilters = EntityFilters.Entity | EntityFilters.Attributes, RetrieveAsIfPublished = false }) as RetrieveEntityResponse;
                    schema = response?.EntityMetadata;
                }
                if (!ReadableSchema(schema, table, out primary, out name))
                { limitations.Add(family + ": readable primary-ID/primary-name schema is not verified."); return; }
                var found = new BatchedDiagnosticQueryReader(context).Read(records, table, primary, new[] { primary, name },
                    "Maker display ID/name batch is incomplete.", "Maker display ID/name correlation is conflicting.",
                    error => "Maker display ID/name batch failed.", token);
                foreach (var identity in members)
                {
                    token.ThrowIfCancellationRequested();
                    if (!identity.Record.ObjectId.HasValue || identity.Record.ObjectId == Guid.Empty) continue;
                    Guid id = identity.Record.ObjectId.Value;
                    var correlation = found.GetCorrelation(id);
                    if (correlation.Status != DiagnosticCorrelationStatus.Unique || correlation.Rows.Count != 1 ||
                        !Exact(correlation.Rows[0], table, primary, id)) continue;
                    var row = correlation.Rows[0];
                    if (!(row.GetAttributeValue<object>(name) is string) || string.IsNullOrWhiteSpace((string)row[name])) continue;
                    values[identity.Record] = Value(identity, kind, id, row, name, null);
                }
                if (members.Any(i => !values.ContainsKey(i.Record)))
                    limitations.Add(family + ": one or more local ID/name observations are unavailable or ambiguous.");
            }
            catch (OperationCanceledException) { throw; }
            catch (Exception)
            {
                token.ThrowIfCancellationRequested();
                limitations.Add(family + ": optional schema/ID/name projection is unavailable.");
            }
        }

        private static bool ReadableSchema(EntityMetadata schema, string table, out string primary, out string name)
        {
            primary = schema?.PrimaryIdAttribute; name = schema?.PrimaryNameAttribute;
            if (schema == null || !StringComparer.OrdinalIgnoreCase.Equals(schema.LogicalName, table) ||
                !Symbol(primary) || !Symbol(name) || primary == name || schema.Attributes == null) return false;
            string idName = primary, displayName = name;
            var ids = schema.Attributes.Where(a => a != null && a.LogicalName == idName).ToArray();
            var names = schema.Attributes.Where(a => a != null && a.LogicalName == displayName).ToArray();
            return ids.Length == 1 && names.Length == 1 && ids[0].AttributeType == AttributeTypeCode.Uniqueidentifier &&
                names[0].AttributeType == AttributeTypeCode.String && ids[0].IsValidForRead == true &&
                names[0].IsValidForRead == true && names[0].IsSecured != true;
        }
        private static bool Symbol(string value) => value != null && Regex.IsMatch(value, "^[a-z][a-z0-9_]*$", RegexOptions.CultureInvariant);
        private static string Normalize(string value) => value?.Replace(" ", "").Trim().ToLowerInvariant();
        private static int? Option(Entity row, string field) => (row.GetAttributeValue<object>(field) as OptionSetValue)?.Value;
        private static bool Exact(Entity row, string table, string primary, Guid id) => row != null && row.Id == id &&
            StringComparer.OrdinalIgnoreCase.Equals(row.LogicalName, table) && row.GetAttributeValue<object>(primary) is Guid && (Guid)row[primary] == id;
        private static ComponentDisplayValue Field(Entity row, string field)
        {
            if (field == null) return new ComponentDisplayValue(DisplayAvailability.NotRetrieved);
            object value;
            if (!row.Attributes.TryGetValue(field, out value)) return new ComponentDisplayValue(DisplayAvailability.AttributeOmitted);
            return value == null || value is string ? new ComponentDisplayValue(DisplayAvailability.Available, (string)value, row.LogicalName + "." + field)
                : new ComponentDisplayValue(DisplayAvailability.Unavailable);
        }
        private static MakerObjectDisplayEvidence Value(ComponentIdentity identity, MakerObjectKind kind, Guid canonical,
            Entity row, string name, string internalName, DisplayBackingCorrelationStatus correlation = DisplayBackingCorrelationStatus.Unique,
            int? dashboardFormType = null, int? processCategory = null) =>
            new MakerObjectDisplayEvidence(identity, kind, canonical, Field(row, name), Field(row, internalName), correlation, dashboardFormType, processCategory);
    }
}
