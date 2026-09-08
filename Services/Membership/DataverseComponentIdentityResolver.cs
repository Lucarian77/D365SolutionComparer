using System;
using System.Collections.Generic;
using System.Linq;
using System.ServiceModel;
using System.Threading;
using D365SolutionComparer.Infrastructure;
using D365SolutionComparer.Models.Identity;
using D365SolutionComparer.Models.Membership;
using D365SolutionComparer.Services.Contracts;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Metadata.Query;
using Microsoft.Xrm.Sdk.Query;

namespace D365SolutionComparer.Services.Membership
{
    /// <summary>Published identity metadata only. No display-name or environment-local GUID fallback.</summary>
    public sealed class DataverseComponentIdentityResolver : IComponentIdentityResolver
    {
        private const int BatchSize = 200;

        public ComponentIdentity Resolve(IOrganizationService service, EnvironmentIdentity environment,
            SolutionComponentRecord component, CancellationToken cancellationToken)
        {
            if (component == null) throw new ArgumentNullException(nameof(component));
            return new ResolutionContext(new DataverseReadContext(service, environment, cancellationToken))
                .Resolve(component, cancellationToken);
        }

        /// <summary>Bulk resolution caches identities and groups safe lookups within this snapshot operation.</summary>
        public MembershipSnapshot ResolveSnapshot(IOrganizationService service, MembershipSnapshot snapshot,
            CancellationToken cancellationToken)
        {
            return ResolveSnapshot(service, snapshot, cancellationToken, null);
        }

        public MembershipSnapshot ResolveSnapshot(IOrganizationService service, MembershipSnapshot snapshot,
            CancellationToken cancellationToken, DataverseRequestCounter requestCounter)
        {
            if (snapshot == null) throw new ArgumentNullException(nameof(snapshot));
            cancellationToken.ThrowIfCancellationRequested();
            if (snapshot.State != MembershipSnapshotState.Complete) return snapshot;
            return ResolveSnapshot(new DataverseReadContext(service, snapshot.Environment, cancellationToken, requestCounter),
                snapshot, cancellationToken);
        }

        internal MembershipSnapshot ResolveSnapshot(DataverseReadContext context, MembershipSnapshot snapshot,
            CancellationToken cancellationToken)
        {
            if (context == null) throw new ArgumentNullException(nameof(context));
            if (snapshot == null) throw new ArgumentNullException(nameof(snapshot));
            cancellationToken.ThrowIfCancellationRequested();
            if (snapshot.State != MembershipSnapshotState.Complete) return snapshot;
            if (context.Environment.OrganizationId != snapshot.Environment.OrganizationId)
                throw new InvalidOperationException("The verified context belongs to a different environment.");
            var resolved = new ResolutionContext(context).ResolveAll(snapshot.Components, cancellationToken);
            cancellationToken.ThrowIfCancellationRequested();
            return MembershipSnapshot.Complete(snapshot.Solution, resolved, snapshot.CapturedAt);
        }

        private sealed class ResolutionContext
        {
            private readonly DataverseReadContext context;
            private readonly Dictionary<LookupKey, ResolutionValue> identityCache =
                new Dictionary<LookupKey, ResolutionValue>();
            private readonly Dictionary<Guid, ResolutionValue> parentWorkflowCache =
                new Dictionary<Guid, ResolutionValue>();
            private readonly Dictionary<int, DefinitionMapping> definitionMappings =
                new Dictionary<int, DefinitionMapping>();
            private readonly HashSet<int> metadataDiagnosticsLoaded = new HashSet<int>();
            private bool connectionMappingLoaded;
            private int? connectionTypeCode;
            private string connectionMappingDiagnostic;
            private IdentityResolutionStatus connectionMappingStatus = IdentityResolutionStatus.Unresolved;

            public ResolutionContext(DataverseReadContext context) { this.context = context; }

            public ComponentIdentity Resolve(SolutionComponentRecord record, CancellationToken cancellationToken)
            {
                cancellationToken.ThrowIfCancellationRequested();
                PrepareClassifications(new[] { record }, cancellationToken);
                string kind;
                var immediate = Classify(record, cancellationToken, out kind);
                if (immediate != null) return immediate;
                var cacheKey = new LookupKey(kind, record.ObjectId.Value);
                ResolutionValue value;
                if (!identityCache.TryGetValue(cacheKey, out value))
                {
                    value = ResolveOne(cacheKey, cancellationToken);
                    identityCache[cacheKey] = value;
                }
                return value.ToIdentity(record);
            }

            public IReadOnlyList<ComponentIdentity> ResolveAll(IReadOnlyList<ComponentIdentity> components,
                CancellationToken cancellationToken)
            {
                PrepareClassifications(components.Select(item => item.Record), cancellationToken);
                var results = new ComponentIdentity[components.Count];
                var pending = new List<PendingRecord>();
                var unique = new HashSet<LookupKey>();
                for (int index = 0; index < components.Count; index++)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    var record = components[index].Record;
                    string kind;
                    var immediate = Classify(record, cancellationToken, out kind);
                    if (immediate != null)
                    {
                        results[index] = immediate;
                        continue;
                    }
                    var key = new LookupKey(kind, record.ObjectId.Value);
                    pending.Add(new PendingRecord(index, record, key));
                    unique.Add(key);
                }

                foreach (var group in unique.GroupBy(item => item.Kind, StringComparer.Ordinal))
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    var keys = group.ToList();
                    if (group.Key == "process") ResolveProcessBatches(keys, cancellationToken);
                    else if (IsEntityBacked(group.Key)) ResolveEntityBatches(group.Key, keys, cancellationToken);
                    else if (group.Key == "table") ResolveTableBatches(keys, cancellationToken);
                    else foreach (var key in keys) identityCache[key] = ResolveOne(key, cancellationToken);
                }

                foreach (var item in pending)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    results[item.Index] = identityCache[item.Key].ToIdentity(item.Record);
                }
                return Array.AsReadOnly(results);
            }

            private ComponentIdentity Classify(SolutionComponentRecord record, CancellationToken cancellationToken,
                out string kind)
            {
                kind = null;
                switch (record.ComponentType)
                {
                    case 1: kind = "table"; break;
                    case 2: kind = "column"; break;
                    case 10: kind = "relationship"; break;
                    case 61: kind = "webresource"; break;
                    case 29: kind = "process"; break;
                    case 20: kind = "securityrole"; break;
                    case 380: kind = "environmentvariabledefinition"; break;
                    case 3:
                    case 11:
                    case 12:
                        return Unknown(record, IdentityResolutionStatus.Unsupported,
                            "This relationship component type is not supported in Phase 2A.");
                    default:
                        if (ComponentSemanticKinds.IsKnownBuiltInType(record.ComponentType))
                            return Unknown(record, IdentityResolutionStatus.Unsupported,
                                "No identity resolver supports this known component type.");
                        cancellationToken.ThrowIfCancellationRequested();
                        if (connectionMappingDiagnostic == null && connectionTypeCode.HasValue &&
                            connectionTypeCode.Value == record.ComponentType)
                        {
                            kind = "connectionreference";
                            break;
                        }
                        DefinitionMapping mapping;
                        if (!definitionMappings.TryGetValue(record.ComponentType, out mapping))
                            return Unknown(record, IdentityResolutionStatus.Unsupported,
                                "No identity resolver supports this component type.");
                        if (mapping.Definition != null)
                            return new ComponentIdentity(record, IdentityResolutionStatus.Unsupported,
                                diagnostic: "No portable identity resolver supports registered solution-component family '" +
                                    mapping.Definition.Name + "'.",
                                semanticKind: mapping.Definition.SemanticKind,
                                registeredDefinition: mapping.Definition);
                        if (connectionMappingDiagnostic != null)
                            return Unknown(record, connectionMappingStatus, connectionMappingDiagnostic);
                        return Unknown(record, mapping.Status, mapping.Diagnostic);
                }
                if (!record.ObjectId.HasValue || record.ObjectId.Value == Guid.Empty)
                    return Unknown(record, IdentityResolutionStatus.Unresolved,
                        "The raw component has no usable object ID.", kind);
                return null;
            }

            private ResolutionValue ResolveOne(LookupKey key, CancellationToken cancellationToken)
            {
                if (key.Kind == "process")
                {
                    ResolveProcessBatches(new[] { key }, cancellationToken);
                    return identityCache[key];
                }
                try
                {
                    string value;
                    if (key.Kind == "table")
                    {
                        var response = context.Execute(new RetrieveEntityRequest
                        {
                            MetadataId = key.ObjectId,
                            EntityFilters = EntityFilters.Entity,
                            RetrieveAsIfPublished = false
                        }) as RetrieveEntityResponse;
                        value = response?.EntityMetadata?.LogicalName;
                    }
                    else if (key.Kind == "column")
                    {
                        var response = context.Execute(new RetrieveAttributeRequest
                        {
                            MetadataId = key.ObjectId,
                            RetrieveAsIfPublished = false
                        }) as RetrieveAttributeResponse;
                        var metadata = response?.AttributeMetadata;
                        value = metadata == null || string.IsNullOrWhiteSpace(metadata.EntityLogicalName) ||
                            string.IsNullOrWhiteSpace(metadata.LogicalName)
                            ? null : metadata.EntityLogicalName + "." + metadata.LogicalName;
                    }
                    else if (key.Kind == "relationship")
                    {
                        var response = context.Execute(new RetrieveRelationshipRequest
                        {
                            MetadataId = key.ObjectId,
                            RetrieveAsIfPublished = false
                        }) as RetrieveRelationshipResponse;
                        value = response?.RelationshipMetadata?.SchemaName;
                    }
                    else return ResolveEntityOne(key, cancellationToken);
                    cancellationToken.ThrowIfCancellationRequested();
                    return ResolutionValue.FromKey(key.Kind, value);
                }
                catch (OperationCanceledException) { throw; }
                catch (FaultException ex)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    return ResolutionValue.Unresolved(key.Kind, "Identity read failed: " + ex.Message);
                }
            }

            private ResolutionValue ResolveEntityOne(LookupKey key, CancellationToken cancellationToken)
            {
                string table, primaryId, identityAttribute;
                GetEntityConfiguration(key.Kind, out table, out primaryId, out identityAttribute);
                var query = new QueryExpression(table)
                {
                    ColumnSet = new ColumnSet(primaryId, identityAttribute),
                    TopCount = 2
                };
                query.Criteria.AddCondition(primaryId, ConditionOperator.Equal, key.ObjectId);
                var rows = context.Query(query);
                cancellationToken.ThrowIfCancellationRequested();
                if (rows.MoreRecords || rows.Entities.Count > 1)
                    return ResolutionValue.Ambiguous(key.Kind, "An object ID lookup returned multiple records.");
                if (rows.Entities.Count == 0) return ResolutionValue.FromKey(key.Kind, null);
                if (rows.Entities[0].Id != key.ObjectId)
                    throw new InvalidOperationException("An object ID lookup returned a different object.");
                return ResolutionValue.FromKey(key.Kind,
                    ReadEntityIdentity(rows.Entities[0], key.Kind, identityAttribute));
            }

            private void ResolveTableBatches(IReadOnlyList<LookupKey> keys, CancellationToken cancellationToken)
            {
                foreach (var batch in Batch(keys))
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    try
                    {
                        var query = new EntityQueryExpression
                        {
                            Properties = new MetadataPropertiesExpression("MetadataId", "LogicalName"),
                            Criteria = new MetadataFilterExpression(LogicalOperator.Or)
                        };
                        foreach (var key in batch)
                            query.Criteria.Conditions.Add(new MetadataConditionExpression("MetadataId",
                                MetadataConditionOperator.Equals, key.ObjectId));
                        var response = context.Execute(new RetrieveMetadataChangesRequest { Query = query })
                            as RetrieveMetadataChangesResponse;
                        var metadata = response?.EntityMetadata ?? new EntityMetadataCollection();
                        var requested = new HashSet<Guid>(batch.Select(item => item.ObjectId));
                        var grouped = metadata.Where(item => item.MetadataId.HasValue)
                            .GroupBy(item => item.MetadataId.Value).ToDictionary(item => item.Key, item => item.ToList());
                        if (grouped.Keys.Any(id => !requested.Contains(id)))
                            throw new InvalidOperationException("A grouped table metadata query returned an unrequested object.");
                        foreach (var key in batch)
                        {
                            List<EntityMetadata> matches;
                            if (!grouped.TryGetValue(key.ObjectId, out matches))
                                identityCache[key] = ResolutionValue.FromKey(key.Kind, null);
                            else if (matches.Count != 1)
                                identityCache[key] = ResolutionValue.Ambiguous(key.Kind,
                                    "A metadata lookup returned multiple records.");
                            else identityCache[key] = ResolutionValue.FromKey(key.Kind, matches[0].LogicalName);
                        }
                    }
                    catch (OperationCanceledException) { throw; }
                    catch (FaultException ex)
                    {
                        cancellationToken.ThrowIfCancellationRequested();
                        foreach (var key in batch)
                            identityCache[key] = ResolutionValue.Unresolved(key.Kind,
                                "Identity read failed: " + ex.Message);
                    }
                }
            }

            private void ResolveEntityBatches(string kind, IReadOnlyList<LookupKey> keys,
                CancellationToken cancellationToken)
            {
                foreach (var batch in Batch(keys))
                    foreach (var result in ResolveEntityBatch(kind, batch, cancellationToken))
                        identityCache[result.Key] = result.Value;
            }

            private IDictionary<LookupKey, ResolutionValue> ResolveEntityBatch(string kind,
                IReadOnlyList<LookupKey> keys, CancellationToken cancellationToken)
            {
                var results = new Dictionary<LookupKey, ResolutionValue>();
                string table, primaryId, identityAttribute;
                GetEntityConfiguration(kind, out table, out primaryId, out identityAttribute);
                try
                {
                    var query = new QueryExpression(table) { ColumnSet = new ColumnSet(primaryId, identityAttribute) };
                    query.Criteria.AddCondition(new ConditionExpression(primaryId, ConditionOperator.In,
                        keys.Select(item => (object)item.ObjectId).ToArray()));
                    var rows = context.Query(query);
                    cancellationToken.ThrowIfCancellationRequested();
                    if (rows.MoreRecords)
                        throw new InvalidOperationException("A bounded identity query unexpectedly returned more records.");
                    var requested = new HashSet<Guid>(keys.Select(item => item.ObjectId));
                    var grouped = rows.Entities.GroupBy(item => item.Id)
                        .ToDictionary(item => item.Key, item => item.ToList());
                    if (grouped.Keys.Any(id => !requested.Contains(id) || id == Guid.Empty))
                        throw new InvalidOperationException("A grouped identity query returned an unrequested object.");
                    foreach (var key in keys)
                    {
                        List<Entity> matches;
                        if (!grouped.TryGetValue(key.ObjectId, out matches))
                            results[key] = ResolutionValue.FromKey(kind, null);
                        else if (matches.Count != 1)
                            results[key] = ResolutionValue.Ambiguous(kind,
                                "An object ID lookup returned multiple records.");
                        else results[key] = ResolutionValue.FromKey(kind,
                            ReadEntityIdentity(matches[0], kind, identityAttribute));
                    }
                }
                catch (OperationCanceledException) { throw; }
                catch (FaultException ex)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    foreach (var key in keys)
                        results[key] = ResolutionValue.Unresolved(kind, "Identity read failed: " + ex.Message);
                }
                return results;
            }

            private void ResolveProcessBatches(IReadOnlyList<LookupKey> keys,
                CancellationToken cancellationToken)
            {
                var activations = new List<PendingWorkflowActivation>();
                foreach (var batch in Batch(keys))
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    try
                    {
                        var query = new QueryExpression("workflow")
                        {
                            ColumnSet = new ColumnSet("workflowid", "uniquename", "name", "type", "category",
                                "primaryentity", "mode", "parentworkflowid", "workflowidunique", "statecode",
                                "statuscode", "componentstate", "ismanaged", "subprocess", "businessprocesstype",
                                "modernflowtype", "uiflowtype")
                        };
                        query.Criteria.AddCondition(new ConditionExpression("workflowid", ConditionOperator.In,
                            batch.Select(item => (object)item.ObjectId).ToArray()));
                        var rows = context.Query(query);
                        cancellationToken.ThrowIfCancellationRequested();
                        var grouped = GroupWorkflowRows(rows, batch.Select(item => item.ObjectId),
                            "A grouped workflow lookup returned an unrequested object.");
                        foreach (var key in batch)
                        {
                            List<Entity> matches;
                            if (!grouped.TryGetValue(key.ObjectId, out matches))
                            {
                                identityCache[key] = ResolutionValue.Unresolved(key.Kind,
                                    "Raw workflow row was not found.");
                                continue;
                            }
                            if (matches.Count != 1)
                            {
                                identityCache[key] = ResolutionValue.Ambiguous(key.Kind,
                                    "A raw workflow lookup returned multiple records.");
                                continue;
                            }
                            var row = matches[0];
                            var uniqueName = row.GetAttributeValue<string>("uniquename");
                            if (!string.IsNullOrWhiteSpace(uniqueName))
                            {
                                identityCache[key] = ResolutionValue.FromKey(key.Kind, uniqueName);
                                continue;
                            }
                            var workflowType = ReadOptionValue(row, "type");
                            if (workflowType == 1)
                                identityCache[key] = ResolutionValue.Unresolved(key.Kind,
                                    BuildBlankWorkflowDefinitionDiagnostic(row));
                            else if (workflowType == 2)
                            {
                                var parent = row.GetAttributeValue<EntityReference>("parentworkflowid");
                                if (parent == null || parent.Id == Guid.Empty)
                                    identityCache[key] = ResolutionValue.Unresolved(key.Kind,
                                        "Workflow activation has no parent workflow definition.");
                                else activations.Add(new PendingWorkflowActivation(key, parent.Id));
                            }
                            else
                                identityCache[key] = ResolutionValue.Unresolved(key.Kind,
                                    "Unsupported workflow record type " +
                                    (workflowType.HasValue
                                        ? workflowType.Value.ToString(System.Globalization.CultureInfo.InvariantCulture)
                                        : "(missing)") +
                                    "; parent identity inheritance is limited to documented activation records (type 2).");
                        }
                    }
                    catch (OperationCanceledException) { throw; }
                    catch (FaultException ex)
                    {
                        cancellationToken.ThrowIfCancellationRequested();
                        foreach (var key in batch)
                            identityCache[key] = ResolutionValue.Unresolved(key.Kind,
                                "Workflow identity read failed: " + ex.Message);
                    }
                }

                ResolveParentWorkflowDefinitions(activations.Select(item => item.ParentWorkflowId)
                    .Distinct().ToList(), cancellationToken);
                foreach (var activation in activations)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    identityCache[activation.Key] = parentWorkflowCache[activation.ParentWorkflowId];
                }
            }

            private void ResolveParentWorkflowDefinitions(IReadOnlyList<Guid> parentIds,
                CancellationToken cancellationToken)
            {
                var missing = parentIds.Where(id => !parentWorkflowCache.ContainsKey(id)).ToList();
                foreach (var batch in Batch(missing))
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    try
                    {
                        var query = new QueryExpression("workflow")
                        {
                            ColumnSet = new ColumnSet("workflowid", "uniquename", "type")
                        };
                        query.Criteria.AddCondition(new ConditionExpression("workflowid", ConditionOperator.In,
                            batch.Select(id => (object)id).ToArray()));
                        var rows = context.Query(query);
                        cancellationToken.ThrowIfCancellationRequested();
                        var grouped = GroupWorkflowRows(rows, batch,
                            "A grouped parent workflow lookup returned an unrequested object.");
                        foreach (var parentId in batch)
                        {
                            List<Entity> matches;
                            if (!grouped.TryGetValue(parentId, out matches))
                                parentWorkflowCache[parentId] = ResolutionValue.Unresolved("process",
                                    "Parent workflow definition was not found.");
                            else if (matches.Count != 1)
                                parentWorkflowCache[parentId] = ResolutionValue.Ambiguous("process",
                                    "Parent workflow definition lookup returned multiple records.");
                            else if (ReadOptionValue(matches[0], "type") != 1)
                                parentWorkflowCache[parentId] = ResolutionValue.Unresolved("process",
                                    "Parent workflow record is not a confirmed definition (type 1).");
                            else
                            {
                                var uniqueName = matches[0].GetAttributeValue<string>("uniquename");
                                parentWorkflowCache[parentId] = string.IsNullOrWhiteSpace(uniqueName)
                                    ? ResolutionValue.Unresolved("process",
                                        "Parent workflow definition has a blank uniquename.")
                                    : ResolutionValue.FromKey("process", uniqueName,
                                        "Portable identity inherited from the parent workflow definition's uniquename.");
                            }
                        }
                    }
                    catch (OperationCanceledException) { throw; }
                    catch (FaultException ex)
                    {
                        cancellationToken.ThrowIfCancellationRequested();
                        foreach (var parentId in batch)
                            parentWorkflowCache[parentId] = ResolutionValue.Unresolved("process",
                                "Parent workflow identity read failed: " + ex.Message);
                    }
                }
            }

            private static IDictionary<Guid, List<Entity>> GroupWorkflowRows(EntityCollection rows,
                IEnumerable<Guid> requestedIds, string unexpectedDiagnostic)
            {
                if (rows.MoreRecords)
                    throw new InvalidOperationException("A bounded workflow identity query unexpectedly returned more records.");
                var requested = new HashSet<Guid>(requestedIds);
                var grouped = rows.Entities.GroupBy(item => item.Id)
                    .ToDictionary(item => item.Key, item => item.ToList());
                if (grouped.Keys.Any(id => id == Guid.Empty || !requested.Contains(id)))
                    throw new InvalidOperationException(unexpectedDiagnostic);
                return grouped;
            }

            private static int? ReadOptionValue(Entity row, string attributeName)
            {
                var option = row.GetAttributeValue<OptionSetValue>(attributeName);
                return option == null ? (int?)null : option.Value;
            }

            private static string BuildBlankWorkflowDefinitionDiagnostic(Entity row)
            {
                var evidence = new[]
                {
                    "workflowid=" + row.Id.ToString("D"),
                    "name=" + FormatWorkflowEvidence(row, "name"),
                    "uniquename=" + FormatWorkflowEvidence(row, "uniquename"),
                    "type=" + FormatWorkflowEvidence(row, "type"),
                    "category=" + FormatWorkflowEvidence(row, "category"),
                    "primaryentity=" + FormatWorkflowEvidence(row, "primaryentity"),
                    "mode=" + FormatWorkflowEvidence(row, "mode"),
                    "parentworkflowid=" + FormatWorkflowEvidence(row, "parentworkflowid"),
                    "workflowidunique=" + FormatWorkflowEvidence(row, "workflowidunique"),
                    "statecode=" + FormatWorkflowEvidence(row, "statecode"),
                    "statuscode=" + FormatWorkflowEvidence(row, "statuscode"),
                    "componentstate=" + FormatWorkflowEvidence(row, "componentstate"),
                    "ismanaged=" + FormatWorkflowEvidence(row, "ismanaged"),
                    "subprocess=" + FormatWorkflowEvidence(row, "subprocess"),
                    "businessprocesstype=" + FormatWorkflowEvidence(row, "businessprocesstype"),
                    "modernflowtype=" + FormatWorkflowEvidence(row, "modernflowtype"),
                    "uiflowtype=" + FormatWorkflowEvidence(row, "uiflowtype")
                };
                return "Workflow definition has a blank uniquename. Diagnostic evidence: " +
                    string.Join("; ", evidence) +
                    ". Diagnostic evidence only; no field listed above is used as a comparison identity.";
            }

            private static string FormatWorkflowEvidence(Entity row, string attributeName)
            {
                object value;
                if (!row.Attributes.TryGetValue(attributeName, out value)) return "(not supplied)";
                if (value == null) return "(null)";
                var option = value as OptionSetValue;
                if (option != null)
                    return AppendFormattedWorkflowEvidence(row, attributeName,
                        option.Value.ToString(System.Globalization.CultureInfo.InvariantCulture));
                var reference = value as EntityReference;
                if (reference != null) return reference.Id == Guid.Empty ? "(empty Guid)" : reference.Id.ToString("D");
                if (value is Guid) return ((Guid)value).ToString("D");
                if (value is bool) return AppendFormattedWorkflowEvidence(row, attributeName,
                    ((bool)value).ToString());
                var text = value as string;
                if (text != null) return "'" + EscapeDiagnosticText(text) + "'";
                return "(unexpected " + value.GetType().FullName + ")";
            }

            private static string AppendFormattedWorkflowEvidence(Entity row, string attributeName,
                string rawValue)
            {
                string formatted;
                return row.FormattedValues.TryGetValue(attributeName, out formatted) &&
                    !string.IsNullOrWhiteSpace(formatted)
                    ? rawValue + " ('" + EscapeDiagnosticText(formatted) + "')"
                    : rawValue;
            }

            private static string EscapeDiagnosticText(string value) => value.Replace("\\", "\\\\")
                .Replace("\r", "\\r").Replace("\n", "\\n").Replace("\t", "\\t").Replace("'", "\\'");

            private static string ReadEntityIdentity(Entity entity, string kind, string identityAttribute)
            {
                if (kind == "securityrole")
                {
                    var template = entity.GetAttributeValue<EntityReference>(identityAttribute);
                    return template == null || template.Id == Guid.Empty ? null : template.Id.ToString("D");
                }
                return entity.GetAttributeValue<string>(identityAttribute);
            }

            private static void GetEntityConfiguration(string kind, out string table, out string primaryId,
                out string identityAttribute)
            {
                switch (kind)
                {
                    case "webresource": table = "webresource"; primaryId = "webresourceid"; identityAttribute = "name"; break;
                    case "process": table = "workflow"; primaryId = "workflowid"; identityAttribute = "uniquename"; break;
                    case "securityrole": table = "role"; primaryId = "roleid"; identityAttribute = "roletemplateid"; break;
                    case "environmentvariabledefinition": table = "environmentvariabledefinition"; primaryId = "environmentvariabledefinitionid"; identityAttribute = "schemaname"; break;
                    case "connectionreference": table = "connectionreference"; primaryId = "connectionreferenceid"; identityAttribute = "connectionreferencelogicalname"; break;
                    default: throw new ArgumentOutOfRangeException(nameof(kind));
                }
            }

            private static bool IsEntityBacked(string kind) => kind == "webresource" || kind == "process" ||
                kind == "securityrole" || kind == "environmentvariabledefinition" || kind == "connectionreference";

            private static IEnumerable<IReadOnlyList<T>> Batch<T>(IReadOnlyList<T> items)
            {
                for (int offset = 0; offset < items.Count; offset += BatchSize)
                    yield return items.Skip(offset).Take(Math.Min(BatchSize, items.Count - offset)).ToList();
            }

            private void PrepareClassifications(IEnumerable<SolutionComponentRecord> records,
                CancellationToken cancellationToken)
            {
                var broadTypes = records.Select(item => item.ComponentType)
                    .Where(type => !ComponentSemanticKinds.IsKnownBuiltInType(type))
                    .Distinct().OrderBy(type => type).ToList();
                if (broadTypes.Count == 0) return;
                LoadConnectionMapping();
                cancellationToken.ThrowIfCancellationRequested();
                var definitionTypes = broadTypes.Where(type => !connectionTypeCode.HasValue ||
                    connectionTypeCode.Value != type).ToList();
                LoadDefinitionMappings(definitionTypes, cancellationToken);
                var unclassifiedTypes = definitionTypes.Where(type =>
                    definitionMappings[type].Definition == null &&
                    definitionMappings[type].Status == IdentityResolutionStatus.Unsupported).ToList();
                LoadEntityMetadataDiagnostics(unclassifiedTypes, cancellationToken);
                LoadComponentTypeChoiceDiagnostics(unclassifiedTypes, cancellationToken);
            }

            private void LoadDefinitionMappings(IReadOnlyList<int> componentTypes,
                CancellationToken cancellationToken)
            {
                var missing = componentTypes.Where(type => !definitionMappings.ContainsKey(type)).ToList();
                if (missing.Count == 0) return;
                foreach (var type in missing)
                    definitionMappings[type] = DefinitionMapping.Unsupported(
                        "No registered solution-component definition was found for this component type.");
                try
                {
                    var query = new QueryExpression("solutioncomponentdefinition")
                    {
                        ColumnSet = new ColumnSet("objecttypecode", "name", "primaryentityname")
                    };
                    query.Criteria.AddCondition(new ConditionExpression("objecttypecode", ConditionOperator.In,
                        missing.Select(type => (object)type).ToArray()));
                    var rows = context.Query(query);
                    cancellationToken.ThrowIfCancellationRequested();
                    if (rows.MoreRecords)
                    {
                        SetDefinitionMappings(missing, DefinitionMapping.Ambiguous(
                            "Registered solution-component definition discovery returned an incomplete result set."));
                        return;
                    }

                    var returned = new List<KeyValuePair<int, Entity>>();
                    foreach (var row in rows.Entities)
                    {
                        var type = ReadObjectTypeCode(row);
                        if (!type.HasValue || !missing.Contains(type.Value))
                        {
                            SetDefinitionMappings(missing, DefinitionMapping.Ambiguous(
                                "Registered solution-component definition discovery returned conflicting or incomplete data."));
                            return;
                        }
                        returned.Add(new KeyValuePair<int, Entity>(type.Value, row));
                    }

                    foreach (var type in missing)
                    {
                        var matches = returned.Where(item => item.Key == type).Select(item => item.Value).ToList();
                        if (matches.Count > 1)
                        {
                            definitionMappings[type] = DefinitionMapping.Ambiguous(
                                "Multiple registered solution-component definitions use this component type.");
                            continue;
                        }
                        if (matches.Count == 0) continue;
                        var name = matches[0].GetAttributeValue<string>("name");
                        var primaryEntityName = matches[0].GetAttributeValue<string>("primaryentityname");
                        if (string.IsNullOrWhiteSpace(name))
                        {
                            definitionMappings[type] = DefinitionMapping.Unresolved(
                                "The registered solution-component definition has no stable name.");
                            continue;
                        }
                        if (string.Equals(primaryEntityName, "connectionreference",
                            StringComparison.OrdinalIgnoreCase))
                        {
                            definitionMappings[type] = DefinitionMapping.Ambiguous(
                                "A registered definition conflicts with Connection Reference type discovery.");
                            continue;
                        }
                        definitionMappings[type] = DefinitionMapping.Registered(
                            new SolutionComponentDefinitionIdentity(type, name, primaryEntityName));
                    }
                }
                catch (OperationCanceledException) { throw; }
                catch (FaultException ex)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    SetDefinitionMappings(missing, DefinitionMapping.Unresolved(
                        "Registered solution-component definition discovery failed: " + ex.Message));
                }
            }

            private void LoadEntityMetadataDiagnostics(IReadOnlyList<int> componentTypes,
                CancellationToken cancellationToken)
            {
                var missing = componentTypes.Where(type => !metadataDiagnosticsLoaded.Contains(type))
                    .Distinct().OrderBy(type => type).ToList();
                foreach (var type in missing) metadataDiagnosticsLoaded.Add(type);
                for (int offset = 0; offset < missing.Count; offset += BatchSize)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    var batch = missing.Skip(offset).Take(Math.Min(BatchSize, missing.Count - offset)).ToList();
                    try
                    {
                        var query = new EntityQueryExpression
                        {
                            Properties = new MetadataPropertiesExpression(
                                "ObjectTypeCode", "LogicalName", "SchemaName"),
                            Criteria = new MetadataFilterExpression(LogicalOperator.Or)
                        };
                        foreach (var type in batch)
                            query.Criteria.Conditions.Add(new MetadataConditionExpression("ObjectTypeCode",
                                MetadataConditionOperator.Equals, type));
                        var response = context.Execute(new RetrieveMetadataChangesRequest { Query = query })
                            as RetrieveMetadataChangesResponse;
                        var metadata = response?.EntityMetadata ?? new EntityMetadataCollection();
                        cancellationToken.ThrowIfCancellationRequested();
                        if (metadata.Any(item => !item.ObjectTypeCode.HasValue ||
                            !batch.Contains(item.ObjectTypeCode.Value)))
                        {
                            AppendMetadataDiagnostic(batch,
                                "Entity metadata diagnostic discovery returned conflicting or incomplete ObjectTypeCode data.");
                            continue;
                        }
                        foreach (var type in batch)
                        {
                            var matches = metadata.Where(item => item.ObjectTypeCode == type).ToList();
                            if (matches.Count == 0)
                                AppendMetadataDiagnostic(type,
                                    "No entity metadata candidate was found for ObjectTypeCode " +
                                    type.ToString(System.Globalization.CultureInfo.InvariantCulture) + ".");
                            else if (matches.Count == 1)
                                AppendMetadataDiagnostic(type, "Entity metadata candidate: " +
                                    DescribeMetadata(matches[0]) +
                                    ". Diagnostic evidence only; no semantic classification was assigned.");
                            else
                                AppendMetadataDiagnostic(type,
                                    "Multiple entity metadata candidates use this ObjectTypeCode: " +
                                    string.Join(" | ", matches.Select(DescribeMetadata)) + ".");
                        }
                    }
                    catch (OperationCanceledException) { throw; }
                    catch (FaultException ex)
                    {
                        cancellationToken.ThrowIfCancellationRequested();
                        AppendMetadataDiagnostic(batch,
                            "Entity metadata diagnostic discovery failed: " + ex.Message);
                    }
                }
            }

            private void LoadComponentTypeChoiceDiagnostics(IReadOnlyList<int> componentTypes,
                CancellationToken cancellationToken)
            {
                if (componentTypes.Count == 0) return;
                cancellationToken.ThrowIfCancellationRequested();
                try
                {
                    var response = context.Execute(new RetrieveAttributeRequest
                    {
                        EntityLogicalName = "solutioncomponent",
                        LogicalName = "componenttype",
                        RetrieveAsIfPublished = false
                    }) as RetrieveAttributeResponse;
                    var metadata = response?.AttributeMetadata as EnumAttributeMetadata;
                    var options = metadata?.OptionSet?.Options;
                    cancellationToken.ThrowIfCancellationRequested();
                    if (options == null)
                    {
                        AppendMetadataDiagnostic(componentTypes,
                            "Published solutioncomponent.componenttype choice metadata was unavailable.");
                        return;
                    }

                    foreach (var type in componentTypes)
                    {
                        var matches = options.Where(option => option.Value == type).ToList();
                        if (matches.Count == 0)
                            AppendMetadataDiagnostic(type,
                                "No published solutioncomponent.componenttype choice label was found for value " +
                                type.ToString(System.Globalization.CultureInfo.InvariantCulture) + ".");
                        else if (matches.Count == 1)
                            AppendMetadataDiagnostic(type,
                                "Published solutioncomponent.componenttype choice evidence: value=" +
                                type.ToString(System.Globalization.CultureInfo.InvariantCulture) +
                                ", labels=[" + DescribeOptionLabels(matches[0]) +
                                "]. Diagnostic evidence only; no semantic classification was assigned.");
                        else
                            AppendMetadataDiagnostic(type,
                                "Multiple published solutioncomponent.componenttype choices use value " +
                                type.ToString(System.Globalization.CultureInfo.InvariantCulture) + ".");
                    }
                }
                catch (OperationCanceledException) { throw; }
                catch (FaultException ex)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    AppendMetadataDiagnostic(componentTypes,
                        "Published solutioncomponent.componenttype choice discovery failed: " + ex.Message);
                }
            }

            private static string DescribeOptionLabels(OptionMetadata option)
            {
                var labels = option.Label?.LocalizedLabels
                    .Where(label => label != null && !string.IsNullOrWhiteSpace(label.Label))
                    .OrderBy(label => label.LanguageCode)
                    .Select(label => label.LanguageCode.ToString(System.Globalization.CultureInfo.InvariantCulture) +
                        ":'" + label.Label + "'")
                    .ToList() ?? new List<string>();
                return labels.Count == 0 ? "(none)" : string.Join(", ", labels);
            }

            private void AppendMetadataDiagnostic(IEnumerable<int> componentTypes, string diagnostic)
            {
                foreach (var type in componentTypes) AppendMetadataDiagnostic(type, diagnostic);
            }

            private void AppendMetadataDiagnostic(int componentType, string diagnostic)
            {
                definitionMappings[componentType] = definitionMappings[componentType]
                    .AppendDiagnostic(diagnostic);
            }

            private static string DescribeMetadata(EntityMetadata metadata)
            {
                return "ObjectTypeCode=" + (metadata.ObjectTypeCode.HasValue
                        ? metadata.ObjectTypeCode.Value.ToString(System.Globalization.CultureInfo.InvariantCulture)
                        : "(missing)") +
                    ", LogicalName='" + (metadata.LogicalName ?? "(missing)") +
                    "', SchemaName='" + (metadata.SchemaName ?? "(missing)") + "'";
            }

            private void SetDefinitionMappings(IEnumerable<int> componentTypes, DefinitionMapping mapping)
            {
                foreach (var type in componentTypes) definitionMappings[type] = mapping;
            }

            private static int? ReadObjectTypeCode(Entity entity)
            {
                object value;
                if (!entity.Attributes.TryGetValue("objecttypecode", out value) || value == null) return null;
                var option = value as OptionSetValue;
                if (option != null) return option.Value;
                return value is int ? (int?)value : null;
            }

            private void LoadConnectionMapping()
            {
                if (connectionMappingLoaded) return;
                connectionMappingLoaded = true;
                try
                {
                    var query = new QueryExpression("solutioncomponentdefinition")
                    {
                        ColumnSet = new ColumnSet("objecttypecode"), TopCount = 2
                    };
                    query.Criteria.AddCondition("primaryentityname", ConditionOperator.Equal, "connectionreference");
                    var rows = context.Query(query);
                    if (rows.MoreRecords || rows.Entities.Count > 1)
                    {
                        connectionMappingStatus = IdentityResolutionStatus.Ambiguous;
                        connectionMappingDiagnostic = "Connection-reference component type mapping is ambiguous.";
                    }
                    else if (rows.Entities.Count == 1)
                    {
                        connectionTypeCode = rows.Entities[0].GetAttributeValue<int?>("objecttypecode");
                        if (!connectionTypeCode.HasValue)
                            connectionMappingDiagnostic = "Connection-reference component type mapping is incomplete.";
                        else if (ComponentSemanticKinds.IsKnownBuiltInType(connectionTypeCode.Value))
                        {
                            connectionMappingStatus = IdentityResolutionStatus.Ambiguous;
                            connectionMappingDiagnostic = "Connection-reference component type mapping conflicts with a known non-Connection-Reference component type.";
                            connectionTypeCode = null;
                        }
                    }
                }
                catch (FaultException ex)
                {
                    connectionMappingDiagnostic = "Connection-reference component type mapping is unavailable: " + ex.Message;
                }
            }

            private static ComponentIdentity Unknown(SolutionComponentRecord record, IdentityResolutionStatus status,
                string diagnostic, string kind = null) =>
                new ComponentIdentity(record, status, diagnostic: diagnostic, componentTypeKey: kind);

            private sealed class DefinitionMapping
            {
                private DefinitionMapping(IdentityResolutionStatus status, string diagnostic,
                    SolutionComponentDefinitionIdentity definition)
                {
                    Status = status;
                    Diagnostic = diagnostic;
                    Definition = definition;
                }
                public IdentityResolutionStatus Status { get; }
                public string Diagnostic { get; }
                public SolutionComponentDefinitionIdentity Definition { get; }
                public static DefinitionMapping Registered(SolutionComponentDefinitionIdentity definition) =>
                    new DefinitionMapping(IdentityResolutionStatus.Unsupported, null, definition);
                public static DefinitionMapping Unsupported(string diagnostic) =>
                    new DefinitionMapping(IdentityResolutionStatus.Unsupported, diagnostic, null);
                public static DefinitionMapping Unresolved(string diagnostic) =>
                    new DefinitionMapping(IdentityResolutionStatus.Unresolved, diagnostic, null);
                public static DefinitionMapping Ambiguous(string diagnostic) =>
                    new DefinitionMapping(IdentityResolutionStatus.Ambiguous, diagnostic, null);
                public DefinitionMapping AppendDiagnostic(string diagnostic) =>
                    new DefinitionMapping(Status,
                        string.IsNullOrEmpty(Diagnostic) ? diagnostic : Diagnostic + " " + diagnostic,
                        Definition);
            }

            private sealed class PendingRecord
            {
                public PendingRecord(int index, SolutionComponentRecord record, LookupKey key)
                {
                    Index = index;
                    Record = record;
                    Key = key;
                }
                public int Index { get; }
                public SolutionComponentRecord Record { get; }
                public LookupKey Key { get; }
            }

            private sealed class ResolutionValue
            {
                private ResolutionValue(string kind, IdentityResolutionStatus status, string key, string diagnostic)
                {
                    Kind = kind;
                    Status = status;
                    Key = key;
                    Diagnostic = diagnostic;
                }
                public string Kind { get; }
                public IdentityResolutionStatus Status { get; }
                public string Key { get; }
                public string Diagnostic { get; }
                public ComponentIdentity ToIdentity(SolutionComponentRecord record) =>
                    new ComponentIdentity(record, Status, Key, Diagnostic, Kind);
                public static ResolutionValue FromKey(string kind, string key, string diagnostic = null) =>
                    string.IsNullOrWhiteSpace(key)
                    ? Unresolved(kind, "No strong portable identity was available; display names and local GUIDs are not used.")
                    : new ResolutionValue(kind, IdentityResolutionStatus.Resolved, key, diagnostic);
                public static ResolutionValue Unresolved(string kind, string diagnostic) =>
                    new ResolutionValue(kind, IdentityResolutionStatus.Unresolved, null, diagnostic);
                public static ResolutionValue Ambiguous(string kind, string diagnostic) =>
                    new ResolutionValue(kind, IdentityResolutionStatus.Ambiguous, null, diagnostic);
            }

            private sealed class PendingWorkflowActivation
            {
                public PendingWorkflowActivation(LookupKey key, Guid parentWorkflowId)
                {
                    Key = key;
                    ParentWorkflowId = parentWorkflowId;
                }
                public LookupKey Key { get; }
                public Guid ParentWorkflowId { get; }
            }

            private struct LookupKey : IEquatable<LookupKey>
            {
                public LookupKey(string kind, Guid objectId) { Kind = kind; ObjectId = objectId; }
                public string Kind { get; }
                public Guid ObjectId { get; }
                public bool Equals(LookupKey other) => ObjectId == other.ObjectId &&
                    string.Equals(Kind, other.Kind, StringComparison.Ordinal);
                public override bool Equals(object obj) => obj is LookupKey && Equals((LookupKey)obj);
                public override int GetHashCode() => unchecked((Kind.GetHashCode() * 397) ^ ObjectId.GetHashCode());
            }
        }
    }
}
