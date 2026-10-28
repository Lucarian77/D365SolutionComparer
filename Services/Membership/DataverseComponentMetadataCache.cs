using System;
using System.Collections.Generic;
using System.Linq;
using D365SolutionComparer.Models.Membership;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Metadata;

namespace D365SolutionComparer.Services.Membership
{
    /// <summary>Operation-scoped published metadata shared by identity and definition reads.</summary>
    internal sealed class DataverseComponentMetadataCache
    {
        internal MetadataForArchivalResolver MetadataForArchival { get; set; }
        private readonly Dictionary<Guid, EntityMetadata> entities = new Dictionary<Guid, EntityMetadata>();
        private readonly Dictionary<Guid, AttributeMetadata> attributes = new Dictionary<Guid, AttributeMetadata>();
        private readonly Dictionary<Guid, RelationshipMetadataBase> relationships =
            new Dictionary<Guid, RelationshipMetadataBase>();
        private readonly HashSet<Guid> completeEntityDefinitions = new HashSet<Guid>();
        private readonly Dictionary<string, Entity> entityRows =
            new Dictionary<string, Entity>(StringComparer.OrdinalIgnoreCase);

        // Optional name observations belong to this read only, never to a portable identity.
        private readonly Dictionary<string, NormalDisplayNameObservation> normalDisplayNames =
            new Dictionary<string, NormalDisplayNameObservation>(StringComparer.OrdinalIgnoreCase);
        private MembershipSnapshot normalDisplayRawOwner, normalDisplayFinalOwner;
        private readonly HashSet<SolutionComponentRecord> normalDisplayRecords = new HashSet<SolutionComponentRecord>();
        private bool normalDisplayDisabled;

        internal void BeginNormalDisplayNames(MembershipSnapshot rawOwner)
        {
            // A cache cannot be rebound to a later read, even for the same organization/IDs.
            if (normalDisplayRawOwner != null || rawOwner == null || rawOwner.Solution == null ||
                rawOwner.State != MembershipSnapshotState.Complete)
            { DisableNormalDisplayNames(); return; }
            normalDisplayRawOwner = rawOwner;
        }

        internal void SealNormalDisplayNames(MembershipSnapshot finalOwner)
        {
            try
            {
                var raw = normalDisplayRawOwner;
                if (normalDisplayDisabled || raw == null || normalDisplayFinalOwner != null ||
                    finalOwner == null || finalOwner.Solution == null ||
                    finalOwner.State != MembershipSnapshotState.Complete ||
                    raw.Environment.OrganizationId != finalOwner.Environment.OrganizationId ||
                    raw.Solution.SolutionId != finalOwner.Solution.SolutionId ||
                    !StringComparer.OrdinalIgnoreCase.Equals(raw.SolutionUniqueName, finalOwner.SolutionUniqueName) ||
                    raw.Components.Count != finalOwner.Components.Count)
                { DisableNormalDisplayNames(); return; }
                for (int index = 0; index < raw.Components.Count; index++)
                {
                    var record = raw.Components[index].Record;
                    if (!ReferenceEquals(record, finalOwner.Components[index].Record))
                    { DisableNormalDisplayNames(); return; }
                    if (record.ComponentType == 31 || record.ComponentType == 36 || record.ComponentType == 300 || record.ComponentType == 511 || record.ComponentType == 59) normalDisplayRecords.Add(record);
                }
                normalDisplayFinalOwner = finalOwner;
            }
            catch (Exception) { DisableNormalDisplayNames(); }
        }

        internal void ObserveNormalDisplayNames(int rawType, DiagnosticQueryResult retrieval)
        {
            if (normalDisplayDisabled || normalDisplayRawOwner == null || normalDisplayFinalOwner != null ||
                retrieval == null || rawType != 31 && rawType != 36 && rawType != 511 && rawType != 59) return;
            try
            {
                // Both fixed fields are selected by the existing normal diagnostic reads.
                string table = rawType == 31 ? "report" : rawType == 36 ? "template" : rawType == 511 ? "teamtemplate" : "savedqueryvisualization";
                string primary = rawType == 31 ? "reportid" : rawType == 36 ? "templateid" : rawType == 511 ? "teamtemplateid" : "savedqueryvisualizationid";
                string field = rawType == 36 ? "title" : rawType == 511 ? "teamtemplatename" : "name";
                foreach (var objectId in retrieval.ObjectIds)
                {
                    string key = EntityKey(table, objectId);
                    if (normalDisplayNames.ContainsKey(key))
                    { DisableNormalDisplayNames(); return; }
                    var correlation = retrieval.GetCorrelation(objectId);
                    var backing = DisplayBackingCorrelationStatus.Unavailable;
                    var availability = DisplayAvailability.Unavailable;
                    string value = null;
                    var parent = new ComponentDisplayValue(DisplayAvailability.NotRetrieved);
                    if (correlation.Status == DiagnosticCorrelationStatus.Missing)
                        backing = DisplayBackingCorrelationStatus.Missing;
                    else if (correlation.Status == DiagnosticCorrelationStatus.Duplicate)
                    { backing = DisplayBackingCorrelationStatus.Ambiguous; availability = DisplayAvailability.Ambiguous; }
                    else if (correlation.Status == DiagnosticCorrelationStatus.Unique && correlation.Rows.Count == 1)
                    {
                        var row = correlation.Rows[0]; object primaryValue, supplied;
                        if (!StringComparer.OrdinalIgnoreCase.Equals(row.LogicalName, table) ||
                            !row.Attributes.TryGetValue(primary, out primaryValue) || !(primaryValue is Guid) ||
                            (Guid)primaryValue != objectId || row.Id != objectId)
                        { backing = DisplayBackingCorrelationStatus.Ambiguous; availability = DisplayAvailability.Ambiguous; }
                        else
                        {
                            backing = DisplayBackingCorrelationStatus.Unique;
                            if (!row.Attributes.TryGetValue(field, out supplied)) availability = DisplayAvailability.AttributeOmitted;
                            else if (supplied == null || supplied is string)
                            { availability = DisplayAvailability.Available; value = (string)supplied; }
                            if (rawType == 59)
                            {
                                object parentValue;
                                parent = row.Attributes.TryGetValue("primaryentitytypecode", out parentValue) && parentValue is string
                                    ? new ComponentDisplayValue(DisplayAvailability.Available, (string)parentValue,
                                        "NormalMembershipDiagnostic.savedqueryvisualization.primaryentitytypecode")
                                    : new ComponentDisplayValue(DisplayAvailability.Unavailable);
                            }
                        }
                    }
                    // Failed combines faulted and incomplete batches in the existing reader.
                    // Do not parse its diagnostic text or invent a more specific failure state.
                    normalDisplayNames.Add(key, new NormalDisplayNameObservation(rawType, objectId, backing,
                        new ComponentDisplayValue(availability, value, "NormalMembershipDiagnostic." + table + "." + field), parentContext: parent));
                }
            }
            catch (Exception) { DisableNormalDisplayNames(); }
        }

        // Type 300 uses only scalars from its existing normal diagnostic query. This does
        // not participate in candidate analysis, matching, or absence classification.
        internal void ObserveCanvasAppDisplayNames(DiagnosticQueryResult retrieval)
        {
            if (normalDisplayDisabled || normalDisplayRawOwner == null || normalDisplayFinalOwner != null ||
                retrieval == null) return;
            try
            {
                foreach (var objectId in retrieval.ObjectIds)
                {
                    string key = EntityKey("canvasapp", objectId);
                    if (normalDisplayNames.ContainsKey(key))
                    { DisableNormalDisplayNames(); return; }
                    var correlation = retrieval.GetCorrelation(objectId);
                    var backing = DisplayBackingCorrelationStatus.Unavailable;
                    var availability = DisplayAvailability.Unavailable;
                    Entity verified = null;
                    if (correlation.Status == DiagnosticCorrelationStatus.Missing)
                        backing = DisplayBackingCorrelationStatus.Missing;
                    else if (correlation.Status == DiagnosticCorrelationStatus.Duplicate)
                    { backing = DisplayBackingCorrelationStatus.Ambiguous; availability = DisplayAvailability.Ambiguous; }
                    else if (correlation.Status == DiagnosticCorrelationStatus.Unique && correlation.Rows.Count == 1)
                    {
                        var row = correlation.Rows[0]; object primary;
                        if (StringComparer.OrdinalIgnoreCase.Equals(row.LogicalName, "canvasapp") &&
                            row.Attributes.TryGetValue("canvasappid", out primary) && primary is Guid &&
                            (Guid)primary == objectId && row.Id == objectId)
                        { backing = DisplayBackingCorrelationStatus.Unique; verified = row; }
                        else
                        { backing = DisplayBackingCorrelationStatus.Ambiguous; availability = DisplayAvailability.Ambiguous; }
                    }
                    normalDisplayNames.Add(key, new NormalDisplayNameObservation(300, objectId, backing,
                        CanvasAppDisplayValue(verified, "displayname", availability),
                        CanvasAppDisplayValue(verified, "name", availability), canvasAppType:
                            verified != null && verified.GetAttributeValue<object>("canvasapptype") is OptionSetValue
                                ? (int?)((OptionSetValue)verified["canvasapptype"]).Value : null));
                }
            }
            catch (Exception) { DisableNormalDisplayNames(); }
        }

        private static ComponentDisplayValue CanvasAppDisplayValue(Entity verified, string field, DisplayAvailability unavailable)
        {
            string provenance = "NormalMembershipDiagnostic.canvasapp." + field;
            if (verified == null) return new ComponentDisplayValue(unavailable, displayEvidenceSource: provenance);
            object value;
            if (!verified.Attributes.TryGetValue(field, out value))
                return new ComponentDisplayValue(DisplayAvailability.AttributeOmitted, displayEvidenceSource: provenance);
            return value == null || value is string
                ? new ComponentDisplayValue(DisplayAvailability.Available, (string)value, provenance)
                : new ComponentDisplayValue(DisplayAvailability.Unavailable, displayEvidenceSource: provenance);
        }

        internal bool TryGetNormalDisplayName(MembershipSnapshot expectedOwner, ComponentIdentity identity,
            out NormalDisplayNameObservation observation)
        {
            observation = null;
            if (normalDisplayDisabled || expectedOwner == null ||
                !ReferenceEquals(expectedOwner, normalDisplayFinalOwner) || identity == null ||
                !normalDisplayRecords.Contains(identity.Record) || !identity.Record.ObjectId.HasValue ||
                identity.Record.ObjectId.Value == Guid.Empty) return false;
            string table = identity.Record.ComponentType == 31 ? "report" : identity.Record.ComponentType == 36 ? "template" :
                identity.Record.ComponentType == 300 ? "canvasapp" : identity.Record.ComponentType == 511 ? "teamtemplate" :
                identity.Record.ComponentType == 59 ? "savedqueryvisualization" : null;
            return table != null && normalDisplayNames.TryGetValue(EntityKey(table, identity.Record.ObjectId.Value), out observation) &&
                observation.RawType == identity.Record.ComponentType && observation.ObjectId == identity.Record.ObjectId.Value;
        }

        private void DisableNormalDisplayNames()
        { normalDisplayDisabled = true; normalDisplayNames.Clear(); normalDisplayRecords.Clear(); normalDisplayFinalOwner = null; }

        internal sealed class NormalDisplayNameObservation
        {
            internal NormalDisplayNameObservation(int rawType, Guid objectId, DisplayBackingCorrelationStatus correlation,
                ComponentDisplayValue friendlyName, ComponentDisplayValue internalName = null, ComponentDisplayValue parentContext = null, int? canvasAppType = null)
            { RawType = rawType; ObjectId = objectId; Correlation = correlation; FriendlyName = friendlyName; InternalName = internalName; ParentContext = parentContext; CanvasAppType = canvasAppType; }
            internal int RawType { get; }
            internal Guid ObjectId { get; }
            internal DisplayBackingCorrelationStatus Correlation { get; }
            internal ComponentDisplayValue FriendlyName { get; }
            internal ComponentDisplayValue InternalName { get; }
            internal ComponentDisplayValue ParentContext { get; }
            internal int? CanvasAppType { get; }
        }

        // Raw membership ID -> uniquely correlated definition (possibly an activation's parent).
        internal readonly Dictionary<Guid, Entity> WorkflowDefinitions = new Dictionary<Guid, Entity>();
        // Set only at the existing verified activation-to-definition handoff.
        internal readonly Dictionary<Guid, Guid> WorkflowDisplayAliases = new Dictionary<Guid, Guid>();
        private MembershipSnapshot makerOwner;
        private readonly Dictionary<SolutionComponentRecord, MakerObjectDisplayEvidence> makerObjects =
            new Dictionary<SolutionComponentRecord, MakerObjectDisplayEvidence>();
        internal IReadOnlyList<string> MakerDisplayLimitations { get; private set; } = new string[0];
        internal void ClearMakerObjectDisplay()
        { makerOwner = null; makerObjects.Clear(); MakerDisplayLimitations = new string[0]; }
        internal void SealMakerObjectDisplay(MembershipSnapshot owner,
            IDictionary<SolutionComponentRecord, MakerObjectDisplayEvidence> values, IEnumerable<string> limitations)
        {
            ClearMakerObjectDisplay();
            if (owner == null || owner.State != MembershipSnapshotState.Complete || !ReferenceEquals(owner, normalDisplayFinalOwner)) return;
            var records = owner.Components.GroupBy(i => i.Record).ToDictionary(g => g.Key, g => g.First());
            if (values.Any(pair => !records.TryGetValue(pair.Key, out var identity) || pair.Value == null || !pair.Value.IsFor(identity))) return;
            foreach (var pair in values) makerObjects.Add(pair.Key, pair.Value);
            makerOwner = owner; MakerDisplayLimitations = limitations.ToList().AsReadOnly();
        }
        internal bool TryGetMakerObjectDisplay(MembershipSnapshot owner, ComponentIdentity identity, out MakerObjectDisplayEvidence value)
        {
            value = null;
            return owner != null && ReferenceEquals(owner, makerOwner) && ReferenceEquals(owner, normalDisplayFinalOwner) &&
                identity != null && makerObjects.TryGetValue(identity.Record, out value) && value.IsFor(identity);
        }

        internal ParentEntityMetadataInventory ParentMetadata { get; set; }
        internal AppSettingResolutionOperation AppSettings { get; set; }
        // Shared with the coordinated definition operation; no repeat Type 92 reads.
        internal Type92EnvironmentEvidence Type92Evidence { get; set; }

        public bool OptionSetCatalogAttempted { get; private set; }
        public OptionSetMetadataBase[] OptionSetCatalog { get; private set; }
        public string OptionSetCatalogFailure { get; private set; }

        public void Store(EntityMetadata metadata, bool completeDefinition = false)
        {
            if (metadata?.MetadataId == null) return;
            entities[metadata.MetadataId.Value] = metadata;
            if (completeDefinition) completeEntityDefinitions.Add(metadata.MetadataId.Value);
        }

        public void Store(AttributeMetadata metadata)
        {
            if (metadata?.MetadataId != null) attributes[metadata.MetadataId.Value] = metadata;
        }

        public void Store(RelationshipMetadataBase metadata)
        {
            if (metadata?.MetadataId != null) relationships[metadata.MetadataId.Value] = metadata;
        }

        public bool TryGetEntity(Guid metadataId, out EntityMetadata metadata) =>
            entities.TryGetValue(metadataId, out metadata);

        // Read-only traversal for optional display projection. No retrieval or cache mutation.
        internal IEnumerable<EntityMetadata> RetainedEntities => entities.Values;

        public bool HasCompleteEntityDefinition(Guid metadataId) =>
            completeEntityDefinitions.Contains(metadataId);

        public bool TryGetAttribute(Guid metadataId, out AttributeMetadata metadata) =>
            attributes.TryGetValue(metadataId, out metadata);

        public bool TryGetRelationship(Guid metadataId, out RelationshipMetadataBase metadata) =>
            relationships.TryGetValue(metadataId, out metadata);

        public void StoreOptionSetCatalog(OptionSetMetadataBase[] catalog, string failure = null)
        {
            OptionSetCatalogAttempted = true;
            OptionSetCatalog = catalog;
            OptionSetCatalogFailure = failure;
        }

        public void StoreEntityRow(Entity entity, Guid primaryId)
        {
            if (entity == null || string.IsNullOrWhiteSpace(entity.LogicalName) || primaryId == Guid.Empty)
                return;
            entityRows[EntityKey(entity.LogicalName, primaryId)] = entity;
        }

        public bool TryGetEntityRow(string logicalName, Guid primaryId, out Entity entity) =>
            entityRows.TryGetValue(EntityKey(logicalName, primaryId), out entity);

        private static string EntityKey(string logicalName, Guid primaryId) =>
            logicalName + ":" + primaryId.ToString("D");
    }
}
