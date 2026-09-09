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
            private const int OptionSetComponentType = 9;
            private const int SavedQueryComponentType = 26;
            private const int ReportComponentType = 31;
            private const int EmailTemplateComponentType = 36;
            private const int SavedQueryVisualizationComponentType = 59;
            private const int SystemFormComponentType = 60;
            private const int SiteMapComponentType = 62;
            private const int AppModuleComponentType = 80;
            private const int CanvasAppComponentType = 300;
            private const int TeamTemplateComponentType = 511;
            private readonly DataverseReadContext context;
            private readonly Dictionary<LookupKey, ResolutionValue> identityCache =
                new Dictionary<LookupKey, ResolutionValue>();
            private readonly Dictionary<Guid, ResolutionValue> parentWorkflowCache =
                new Dictionary<Guid, ResolutionValue>();
            private readonly Dictionary<int, DefinitionMapping> definitionMappings =
                new Dictionary<int, DefinitionMapping>();
            private readonly HashSet<int> metadataDiagnosticsLoaded = new HashSet<int>();
            private readonly Dictionary<Guid, IReadOnlyList<string>> optionSetDiagnostics =
                new Dictionary<Guid, IReadOnlyList<string>>();
            private readonly Dictionary<Guid, IReadOnlyList<string>> reportDiagnostics =
                new Dictionary<Guid, IReadOnlyList<string>>();
            private readonly Dictionary<int, Dictionary<Guid, IReadOnlyList<string>>> weakIdentityDiagnostics =
                new Dictionary<int, Dictionary<Guid, IReadOnlyList<string>>>();
            private readonly HashSet<int> weakIdentityDiagnosticsLoaded = new HashSet<int>();
            private readonly Dictionary<int, Guid?> weakIdentitySummaryComponentIds =
                new Dictionary<int, Guid?>();
            private readonly Dictionary<int, string> weakIdentitySummaryEvidence =
                new Dictionary<int, string>();
            private readonly Dictionary<Guid, IReadOnlyList<string>> systemFormDiagnostics =
                new Dictionary<Guid, IReadOnlyList<string>>();
            private readonly Dictionary<Guid, IReadOnlyList<string>> siteMapDiagnostics =
                new Dictionary<Guid, IReadOnlyList<string>>();
            private readonly Dictionary<Guid, IReadOnlyList<string>> canvasAppDiagnostics =
                new Dictionary<Guid, IReadOnlyList<string>>();
            private readonly Dictionary<Guid, IReadOnlyList<string>> teamTemplateDiagnostics =
                new Dictionary<Guid, IReadOnlyList<string>>();
            private readonly HashSet<Guid> verifiedTeamTemplates = new HashSet<Guid>();
            private bool connectionMappingLoaded;
            private bool optionSetDiagnosticsLoaded;
            private bool reportDiagnosticsLoaded;
            private bool systemFormDiagnosticsLoaded;
            private bool siteMapDiagnosticsLoaded;
            private bool canvasAppDiagnosticsLoaded;
            private bool teamTemplateDiagnosticsLoaded;
            private Guid? optionSetSummaryComponentId;
            private string optionSetSummaryEvidence;
            private Guid? reportSummaryComponentId;
            private string reportSummaryEvidence;
            private Guid? systemFormSummaryComponentId;
            private string systemFormSummaryEvidence;
            private Guid? siteMapSummaryComponentId;
            private string siteMapSummaryEvidence;
            private Guid? canvasAppSummaryComponentId;
            private string canvasAppSummaryEvidence;
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
                    if (group.Key == ComponentSemanticKinds.Process) ResolveProcessBatches(keys, cancellationToken);
                    else if (group.Key == ComponentSemanticKinds.AppModule)
                        ResolveAppModuleBatches(keys, cancellationToken);
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
                    case AppModuleComponentType: kind = ComponentSemanticKinds.AppModule; break;
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
                                "No identity resolver supports this known component type.",
                                diagnosticEvidence: GetKnownComponentDiagnosticEvidence(record));
                        cancellationToken.ThrowIfCancellationRequested();
                        if (connectionMappingDiagnostic == null && connectionTypeCode.HasValue &&
                            connectionTypeCode.Value == record.ComponentType)
                        {
                            kind = "connectionreference";
                            break;
                        }
                        if (record.ComponentType == TeamTemplateComponentType && record.ObjectId.HasValue &&
                            verifiedTeamTemplates.Contains(record.ObjectId.Value))
                            return new ComponentIdentity(record, IdentityResolutionStatus.Unsupported,
                                diagnostic: "The component is a verified TeamTemplate, but no portable identity resolver is approved.",
                                componentTypeKey: ComponentSemanticKinds.TeamTemplate,
                                semanticKind: ComponentSemanticKinds.TeamTemplate,
                                diagnosticEvidence: GetTeamTemplateDiagnosticEvidence(record));
                        DefinitionMapping mapping;
                        if (!definitionMappings.TryGetValue(record.ComponentType, out mapping))
                            return Unknown(record, IdentityResolutionStatus.Unsupported,
                                "No identity resolver supports this component type.",
                                diagnosticEvidence: GetTeamTemplateDiagnosticEvidence(record));
                        if (mapping.Definition != null)
                            return new ComponentIdentity(record, IdentityResolutionStatus.Unsupported,
                                diagnostic: "No portable identity resolver supports registered solution-component family '" +
                                    mapping.Definition.Name + "'.",
                                semanticKind: mapping.Definition.SemanticKind,
                                registeredDefinition: mapping.Definition);
                        if (connectionMappingDiagnostic != null)
                            return Unknown(record, connectionMappingStatus, connectionMappingDiagnostic,
                                diagnosticEvidence: GetTeamTemplateDiagnosticEvidence(record));
                        return Unknown(record, mapping.Status, mapping.Diagnostic,
                            diagnosticEvidence: GetTeamTemplateDiagnosticEvidence(record));
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
                if (key.Kind == ComponentSemanticKinds.AppModule)
                {
                    ResolveAppModuleBatches(new[] { key }, cancellationToken);
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
                var recordList = records.ToList();
                LoadOptionSetDiagnostics(recordList.Where(item =>
                    item.ComponentType == OptionSetComponentType).ToList(), cancellationToken);
                LoadWeakIdentityDiagnostics(recordList.Where(item =>
                    item.ComponentType == SavedQueryComponentType).ToList(), cancellationToken);
                LoadReportDiagnostics(recordList.Where(item =>
                    item.ComponentType == ReportComponentType).ToList(), cancellationToken);
                LoadWeakIdentityDiagnostics(recordList.Where(item =>
                    item.ComponentType == EmailTemplateComponentType).ToList(), cancellationToken);
                LoadWeakIdentityDiagnostics(recordList.Where(item =>
                    item.ComponentType == SavedQueryVisualizationComponentType).ToList(), cancellationToken);
                LoadSystemFormDiagnostics(recordList.Where(item =>
                    item.ComponentType == SystemFormComponentType).ToList(), cancellationToken);
                LoadSiteMapDiagnostics(recordList.Where(item =>
                    item.ComponentType == SiteMapComponentType).ToList(), cancellationToken);
                LoadCanvasAppDiagnostics(recordList.Where(item =>
                    item.ComponentType == CanvasAppComponentType).ToList(), cancellationToken);
                var broadTypes = recordList.Select(item => item.ComponentType)
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
                bool type511RemainsBroad = definitionTypes.Contains(TeamTemplateComponentType) &&
                    definitionMappings[TeamTemplateComponentType].Definition == null;
                if (type511RemainsBroad)
                    LoadTeamTemplateDiagnostics(recordList.Where(item =>
                        item.ComponentType == TeamTemplateComponentType).ToList(), cancellationToken);
            }

            private void LoadOptionSetDiagnostics(IReadOnlyList<SolutionComponentRecord> records,
                CancellationToken cancellationToken)
            {
                if (records.Count == 0 || optionSetDiagnosticsLoaded) return;
                optionSetDiagnosticsLoaded = true;
                optionSetSummaryComponentId = records[0].SolutionComponentId;
                var objectIds = records.Where(item => item.ObjectId.HasValue &&
                        item.ObjectId.Value != Guid.Empty)
                    .Select(item => item.ObjectId.Value).Distinct().OrderBy(item => item).ToList();
                try
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    var response = context.Execute(new RetrieveAllOptionSetsRequest
                    {
                        RetrieveAsIfPublished = false
                    }) as RetrieveAllOptionSetsResponse;
                    cancellationToken.ThrowIfCancellationRequested();
                    var returned = response != null && response.Results.Contains("OptionSetMetadata")
                        ? response.Results["OptionSetMetadata"] as OptionSetMetadataBase[] : null;
                    if (returned == null)
                    {
                        SetOptionSetDiagnosticFailures(objectIds,
                            "RetrieveAllOptionSets diagnostic lookup returned no option-set catalog.");
                        optionSetSummaryEvidence = DescribeOptionSetSummary(records.Count, objectIds.Count,
                            null, null, null, records.Count(item => !item.ObjectId.HasValue ||
                                item.ObjectId.Value == Guid.Empty), null, null, null, null);
                        return;
                    }

                    var indexed = returned.Where(item => item != null && item.MetadataId.HasValue)
                        .GroupBy(item => item.MetadataId.Value)
                        .ToDictionary(group => group.Key, group => group.ToList());
                    int correlated = 0;
                    int missing = 0;
                    int nonUnique = 0;
                    int blankName = 0;
                    int nonGlobal = 0;
                    var candidateNames = new List<string>();
                    foreach (var objectId in objectIds)
                    {
                        List<OptionSetMetadataBase> matches;
                        if (!indexed.TryGetValue(objectId, out matches))
                        {
                            missing++;
                            optionSetDiagnostics[objectId] = new[]
                            {
                                "No option-set metadata in the retrieved catalog matched this solutioncomponent objectid."
                            };
                            continue;
                        }
                        if (matches.Count != 1)
                        {
                            nonUnique++;
                            var evidence = new List<string>
                            {
                                "Multiple option-set metadata definitions in the retrieved catalog matched this solutioncomponent objectid."
                            };
                            evidence.AddRange(matches.Select(item => DescribeOptionSet(item, objectId, false)));
                            optionSetDiagnostics[objectId] = evidence.AsReadOnly();
                            continue;
                        }

                        var metadata = matches[0];
                        correlated++;
                        if (string.IsNullOrWhiteSpace(metadata.Name)) blankName++;
                        else candidateNames.Add(metadata.Name);
                        if (metadata.IsGlobal != true) nonGlobal++;
                        optionSetDiagnostics[objectId] = new[] { DescribeOptionSet(metadata, objectId, true) };
                    }
                    int missingObjectIds = records.Count(item => !item.ObjectId.HasValue ||
                        item.ObjectId.Value == Guid.Empty);
                    var distinctNames = candidateNames.GroupBy(item => item, StringComparer.OrdinalIgnoreCase)
                        .Select(group => group.First()).OrderBy(item => item, StringComparer.OrdinalIgnoreCase)
                        .ToList();
                    optionSetSummaryEvidence = DescribeOptionSetSummary(records.Count, objectIds.Count,
                        returned.Length, correlated, missing, missingObjectIds, blankName, nonGlobal,
                        nonUnique, distinctNames);
                }
                catch (OperationCanceledException) { throw; }
                catch (FaultException ex)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    SetOptionSetDiagnosticFailures(objectIds,
                        "RetrieveAllOptionSets diagnostic lookup failed: " + ex.Message);
                    optionSetSummaryEvidence = DescribeOptionSetSummary(records.Count, objectIds.Count,
                        null, null, null, records.Count(item => !item.ObjectId.HasValue ||
                            item.ObjectId.Value == Guid.Empty), null, null, null, null);
                }
            }

            private IEnumerable<string> GetOptionSetDiagnosticEvidence(SolutionComponentRecord record)
            {
                if (record.ComponentType != OptionSetComponentType || !optionSetDiagnosticsLoaded)
                    return new string[0];
                var result = new List<string>();
                if (!record.ObjectId.HasValue || record.ObjectId.Value == Guid.Empty)
                    result.Add("OptionSet metadata diagnostic correlation was not attempted because objectid is unavailable.");
                else
                {
                    IReadOnlyList<string> evidence;
                    result.AddRange(optionSetDiagnostics.TryGetValue(record.ObjectId.Value, out evidence)
                        ? evidence : new[] { "OptionSet metadata diagnostic lookup produced no auditable result." });
                }
                if (record.SolutionComponentId == optionSetSummaryComponentId &&
                    !string.IsNullOrWhiteSpace(optionSetSummaryEvidence))
                    result.Add(optionSetSummaryEvidence);
                return result;
            }

            private IEnumerable<string> GetKnownComponentDiagnosticEvidence(SolutionComponentRecord record)
            {
                if (record.ComponentType == OptionSetComponentType)
                    return GetOptionSetDiagnosticEvidence(record);
                if (record.ComponentType == SavedQueryComponentType ||
                    record.ComponentType == EmailTemplateComponentType ||
                    record.ComponentType == SavedQueryVisualizationComponentType)
                    return GetWeakIdentityDiagnosticEvidence(record);
                if (record.ComponentType == ReportComponentType)
                    return GetReportDiagnosticEvidence(record);
                if (record.ComponentType == SystemFormComponentType)
                    return GetSystemFormDiagnosticEvidence(record);
                if (record.ComponentType == SiteMapComponentType)
                    return GetSiteMapDiagnosticEvidence(record);
                if (record.ComponentType == CanvasAppComponentType)
                    return GetCanvasAppDiagnosticEvidence(record);
                return new string[0];
            }

            private void SetOptionSetDiagnosticFailures(IEnumerable<Guid> objectIds, string diagnostic)
            {
                foreach (var objectId in objectIds)
                    optionSetDiagnostics[objectId] = new[] { diagnostic };
            }

            private static string DescribeOptionSet(OptionSetMetadataBase metadata, Guid requestedMetadataId,
                bool uniqueMatch)
            {
                var concerns = new List<string>();
                if (!metadata.MetadataId.HasValue)
                    concerns.Add("returned MetadataId is unavailable");
                else if (metadata.MetadataId.Value != requestedMetadataId)
                    concerns.Add("returned MetadataId does not match solutioncomponent.objectid");
                if (string.IsNullOrWhiteSpace(metadata.Name)) concerns.Add("Name is blank");
                if (metadata.IsGlobal != true) concerns.Add("IsGlobal is not true");

                var prefix = uniqueMatch
                    ? concerns.Count == 0
                        ? "OptionSet metadata diagnostic lookup matched uniquely. "
                        : "OptionSet metadata diagnostic lookup returned concerns: " +
                            string.Join("; ", concerns) + ". "
                    : "Duplicate OptionSet metadata candidate" + (concerns.Count == 0 ? ". " :
                        " with concerns: " + string.Join("; ", concerns) + ". ");
                return prefix +
                    "MetadataId=" + FormatOptionSetGuid(metadata.MetadataId) +
                    "; Name=" + FormatOptionSetText(metadata.Name) +
                    "; IsGlobal=" + FormatOptionSetBoolean(metadata.IsGlobal) +
                    "; OptionSetType=" + FormatOptionSetType(metadata.OptionSetType) +
                    "; IsManaged=" + FormatOptionSetBoolean(metadata.IsManaged) +
                    "; IsCustomOptionSet=" + FormatOptionSetBoolean(metadata.IsCustomOptionSet) +
                    ". Diagnostic evidence only; none of these values is used as a portable comparison identity.";
            }

            private static string DescribeOptionSetSummary(int rawCount, int distinctObjectIdCount,
                int? returnedCount, int? correlatedCount, int? missingCount, int? missingObjectIdCount,
                int? blankNameCount, int? nonGlobalCount, int? nonUniqueCount,
                IReadOnlyList<string> distinctNames)
            {
                return "OptionSet metadata diagnostic summary: RawType9MembershipCount=" + rawCount +
                    "; DistinctNonemptyObjectIdCount=" + distinctObjectIdCount +
                    "; ReturnedOptionSetCount=" + FormatOptionSetCount(returnedCount) +
                    "; CorrelatedMetadataIdCount=" + FormatOptionSetCount(correlatedCount) +
                    "; MissingRequestedMetadataIdCount=" + FormatOptionSetCount(missingCount) +
                    "; MissingObjectIdRecordCount=" + FormatOptionSetCount(missingObjectIdCount) +
                    "; BlankNameCount=" + FormatOptionSetCount(blankNameCount) +
                    "; IsGlobalNotTrueCount=" + FormatOptionSetCount(nonGlobalCount) +
                    "; NonUniqueMetadataIdCount=" + FormatOptionSetCount(nonUniqueCount) +
                    "; DistinctCandidateNames=" + (distinctNames == null ? "(unavailable)" :
                        "[" + string.Join(", ", distinctNames.Select(item => "'" +
                            EscapeDiagnosticText(item) + "'")) + "]") + ".";
            }

            private static string FormatOptionSetCount(int? value) => value.HasValue
                ? value.Value.ToString(System.Globalization.CultureInfo.InvariantCulture) : "(unavailable)";

            private static string FormatOptionSetGuid(Guid? value) =>
                value.HasValue ? value.Value.ToString("D") : "(null)";

            private static string FormatOptionSetText(string value) =>
                value == null ? "(null)" : "'" + EscapeDiagnosticText(value) + "'";

            private static string FormatOptionSetBoolean(bool? value) =>
                value.HasValue ? value.Value.ToString() : "(null)";

            private static string FormatOptionSetType(OptionSetType? value) =>
                value.HasValue
                    ? ((int)value.Value).ToString(System.Globalization.CultureInfo.InvariantCulture) +
                        " ('" + value.Value + "')"
                    : "(null)";

            private void LoadWeakIdentityDiagnostics(IReadOnlyList<SolutionComponentRecord> records,
                CancellationToken cancellationToken)
            {
                if (records.Count == 0) return;
                var componentType = records[0].ComponentType;
                if (weakIdentityDiagnosticsLoaded.Contains(componentType)) return;
                weakIdentityDiagnosticsLoaded.Add(componentType);
                weakIdentitySummaryComponentIds[componentType] = records[0].SolutionComponentId;
                var configuration = WeakIdentityConfiguration.For(componentType);
                var diagnostics = new Dictionary<Guid, IReadOnlyList<string>>();
                weakIdentityDiagnostics[componentType] = diagnostics;
                var objectIds = records.Where(item => item.ObjectId.HasValue &&
                        item.ObjectId.Value != Guid.Empty)
                    .Select(item => item.ObjectId.Value).Distinct().OrderBy(item => item).ToList();
                var returnedById = objectIds.ToDictionary(item => item, item => new List<Entity>());
                var failures = new Dictionary<Guid, string>();
                var unassociatedEvidence = new Dictionary<Guid, List<string>>();
                int returnedCount = 0;
                bool countUnavailable = false;

                foreach (var batch in Batch(objectIds))
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    try
                    {
                        var query = new QueryExpression(configuration.EntityName)
                        {
                            ColumnSet = new ColumnSet(configuration.Columns)
                        };
                        query.Criteria.AddCondition(new ConditionExpression(configuration.PrimaryIdAttribute,
                            ConditionOperator.In, batch.Select(item => (object)item).ToArray()));
                        var rows = context.Query(query);
                        cancellationToken.ThrowIfCancellationRequested();
                        returnedCount += rows.Entities.Count;
                        bool invalidResponse = false;
                        foreach (var row in rows.Entities)
                        {
                            Guid primaryId;
                            if (string.Equals(row.LogicalName, configuration.EntityName,
                                    StringComparison.OrdinalIgnoreCase) &&
                                TryReadGuid(row, configuration.PrimaryIdAttribute, out primaryId) &&
                                batch.Contains(primaryId) && (row.Id == Guid.Empty || row.Id == primaryId))
                            {
                                returnedById[primaryId].Add(row);
                                continue;
                            }

                            invalidResponse = true;
                            var detail = "Unassociated or conflicting returned " +
                                configuration.EntityName + " row: " + DescribeWeakIdentityRow(configuration, row);
                            Guid conflictingId;
                            var affectedIds = TryReadGuid(row, configuration.PrimaryIdAttribute,
                                out conflictingId) && batch.Contains(conflictingId)
                                ? new[] { conflictingId } : batch;
                            foreach (var objectId in affectedIds)
                            {
                                List<string> evidence;
                                if (!unassociatedEvidence.TryGetValue(objectId, out evidence))
                                    unassociatedEvidence[objectId] = evidence = new List<string>();
                                evidence.Add(detail);
                            }
                        }

                        if (rows.MoreRecords)
                        {
                            countUnavailable = true;
                            SetWeakIdentityFailures(batch, configuration.DisplayName +
                                " diagnostic lookup returned an incomplete result set.", failures);
                        }
                        else if (invalidResponse)
                        {
                            countUnavailable = true;
                            SetWeakIdentityFailures(batch, configuration.DisplayName +
                                " diagnostic lookup returned conflicting or incomplete primary-key data.",
                                failures);
                        }
                    }
                    catch (OperationCanceledException) { throw; }
                    catch (FaultException ex)
                    {
                        cancellationToken.ThrowIfCancellationRequested();
                        countUnavailable = true;
                        SetWeakIdentityFailures(batch, configuration.DisplayName +
                            " diagnostic lookup failed: " + ex.Message, failures);
                    }
                }

                int correlated = 0;
                int missing = 0;
                int nonUnique = 0;
                int incomplete = 0;
                int blankName = 0;
                var contexts = new List<string>();
                foreach (var objectId in objectIds)
                {
                    var evidence = new List<string>();
                    string failure;
                    var matches = returnedById[objectId];
                    if (failures.TryGetValue(objectId, out failure)) evidence.Add(failure);
                    else if (matches.Count == 0)
                    {
                        missing++;
                        evidence.Add("No " + configuration.EntityName +
                            " row matched this solutioncomponent objectid.");
                    }
                    else if (matches.Count > 1)
                    {
                        nonUnique++;
                        evidence.Add("Multiple " + configuration.EntityName +
                            " rows matched this solutioncomponent objectid.");
                    }

                    foreach (var row in matches) evidence.Add(DescribeWeakIdentityRow(configuration, row));
                    if (!failures.ContainsKey(objectId) && matches.Count == 1)
                    {
                        correlated++;
                        if (!IsCompleteWeakIdentityRow(configuration, matches[0])) incomplete++;
                        if (!HasText(matches[0], configuration.NameAttribute)) blankName++;
                        contexts.Add(DescribeWeakIdentityContext(configuration, matches[0]));
                    }
                    List<string> unassociated;
                    if (unassociatedEvidence.TryGetValue(objectId, out unassociated)) evidence.AddRange(unassociated);
                    diagnostics[objectId] = evidence.AsReadOnly();
                }

                int missingObjectIds = records.Count(item => !item.ObjectId.HasValue ||
                    item.ObjectId.Value == Guid.Empty);
                var distinctContexts = contexts.GroupBy(item => item, StringComparer.OrdinalIgnoreCase)
                    .Select(group => group.First()).OrderBy(item => item, StringComparer.OrdinalIgnoreCase)
                    .ToList();
                weakIdentitySummaryEvidence[componentType] = configuration.DisplayName +
                    " diagnostic summary: RawComponentCount=" + records.Count +
                    "; ComponentType=" + componentType +
                    "; DistinctObjectIdCount=" + objectIds.Count +
                    "; ReturnedRowCount=" + FormatOptionSetCount(countUnavailable ? (int?)null : returnedCount) +
                    "; CorrelatedCount=" + FormatOptionSetCount(countUnavailable ? (int?)null : correlated) +
                    "; MissingCount=" + FormatOptionSetCount(countUnavailable ? (int?)null : missing) +
                    "; MissingObjectIdRecordCount=" + missingObjectIds +
                    "; NonUniqueObjectIdCount=" + FormatOptionSetCount(countUnavailable ? (int?)null : nonUnique) +
                    "; IncompleteCorrelatedRowCount=" + FormatOptionSetCount(countUnavailable ? (int?)null : incomplete) +
                    "; BlankNameCount=" + FormatOptionSetCount(countUnavailable ? (int?)null : blankName) +
                    "; DistinctDiagnosticContexts=" + (countUnavailable ? "(unavailable)" :
                        "[" + string.Join(", ", distinctContexts.Select(item => "'" +
                            EscapeDiagnosticText(item) + "'")) + "]") +
                    ". Contexts are diagnostic evidence only and are not portable identities.";
            }

            private IEnumerable<string> GetWeakIdentityDiagnosticEvidence(SolutionComponentRecord record)
            {
                if (!weakIdentityDiagnosticsLoaded.Contains(record.ComponentType)) return new string[0];
                var configuration = WeakIdentityConfiguration.For(record.ComponentType);
                var result = new List<string>();
                if (!record.ObjectId.HasValue || record.ObjectId.Value == Guid.Empty)
                    result.Add(configuration.DisplayName +
                        " diagnostic lookup was not attempted because objectid is unavailable.");
                else
                {
                    IReadOnlyList<string> evidence;
                    Dictionary<Guid, IReadOnlyList<string>> diagnostics;
                    result.AddRange(weakIdentityDiagnostics.TryGetValue(record.ComponentType, out diagnostics) &&
                            diagnostics.TryGetValue(record.ObjectId.Value, out evidence)
                        ? evidence : new[] { configuration.DisplayName +
                            " diagnostic lookup produced no auditable result." });
                }
                Guid? summaryComponentId;
                string summary;
                if (weakIdentitySummaryComponentIds.TryGetValue(record.ComponentType, out summaryComponentId) &&
                    record.SolutionComponentId == summaryComponentId &&
                    weakIdentitySummaryEvidence.TryGetValue(record.ComponentType, out summary) &&
                    !string.IsNullOrWhiteSpace(summary))
                    result.Add(summary);
                return result;
            }

            private static void SetWeakIdentityFailures(IEnumerable<Guid> objectIds, string diagnostic,
                IDictionary<Guid, string> failures)
            {
                foreach (var objectId in objectIds) failures[objectId] = diagnostic;
            }

            private static string DescribeWeakIdentityRow(WeakIdentityConfiguration configuration, Entity row)
            {
                var complete = IsCompleteWeakIdentityRow(configuration, row);
                return (complete ? configuration.DisplayName + " diagnostic lookup matched. " :
                    configuration.DisplayName + " diagnostic lookup matched but returned incomplete data. ") +
                    string.Join("; ", configuration.Columns.Select(attribute => attribute + "=" +
                        FormatWeakIdentityValue(row, attribute))) +
                    "; diagnosticContext='" + EscapeDiagnosticText(
                        DescribeWeakIdentityContext(configuration, row)) +
                    "'. Diagnostic evidence only; no value is used for membership comparison.";
            }

            private static string DescribeWeakIdentityContext(WeakIdentityConfiguration configuration, Entity row)
            {
                if (configuration.ComponentType == SavedQueryComponentType)
                    return "returnedtypecode=" + FormatWeakIdentityValue(row, "returnedtypecode") +
                        ", querytype=" + FormatWeakIdentityValue(row, "querytype") +
                        ", name=" + FormatWeakIdentityValue(row, "name");
                if (configuration.ComponentType == EmailTemplateComponentType)
                    return "templatetypecode=" + FormatWeakIdentityValue(row, "templatetypecode") +
                        ", ispersonal=" + FormatWeakIdentityValue(row, "ispersonal") +
                        ", languagecode=" + FormatWeakIdentityValue(row, "languagecode") +
                        ", title=" + FormatWeakIdentityValue(row, "title");
                return "primaryentitytypecode=" + FormatWeakIdentityValue(row, "primaryentitytypecode") +
                    ", type=" + FormatWeakIdentityValue(row, "type") +
                    ", charttype=" + FormatWeakIdentityValue(row, "charttype") +
                    ", name=" + FormatWeakIdentityValue(row, "name");
            }

            private static bool IsCompleteWeakIdentityRow(WeakIdentityConfiguration configuration, Entity row)
            {
                if (!HasGuid(row, configuration.PrimaryIdAttribute) ||
                    !HasText(row, configuration.NameAttribute) ||
                    !HasGuid(row, configuration.UniqueIdAttribute) ||
                    !HasOption(row, "componentstate") || !HasBoolean(row, "ismanaged")) return false;
                if (configuration.ComponentType == SavedQueryComponentType)
                    return HasEntityNameValue(row, "returnedtypecode") && HasInteger(row, "querytype");
                if (configuration.ComponentType == EmailTemplateComponentType)
                    return HasEntityNameValue(row, "templatetypecode") && HasBoolean(row, "ispersonal") &&
                        HasInteger(row, "languagecode");
                return HasEntityNameValue(row, "primaryentitytypecode") && HasOption(row, "type") &&
                    HasOption(row, "charttype");
            }

            private static bool HasBoolean(Entity row, string attributeName)
            {
                object value;
                return row.Attributes.TryGetValue(attributeName, out value) && value is bool;
            }

            private static bool HasEntityNameValue(Entity row, string attributeName)
            {
                object value;
                if (!row.Attributes.TryGetValue(attributeName, out value) || value == null) return false;
                var text = value as string;
                return text != null ? !string.IsNullOrWhiteSpace(text) :
                    value is int || value is OptionSetValue;
            }

            private static string FormatWeakIdentityValue(Entity row, string attributeName)
            {
                object value;
                if (!row.Attributes.TryGetValue(attributeName, out value)) return "(not supplied)";
                if (value == null) return "(null)";
                var option = value as OptionSetValue;
                if (option != null)
                    return AppendFormattedWorkflowEvidence(row, attributeName,
                        option.Value.ToString(System.Globalization.CultureInfo.InvariantCulture));
                if (value is Guid) return ((Guid)value).ToString("D");
                if (value is bool) return ((bool)value).ToString();
                if (value is int) return ((int)value).ToString(System.Globalization.CultureInfo.InvariantCulture);
                var text = value as string;
                if (text != null) return "'" + EscapeDiagnosticText(text) + "'";
                return "(unexpected " + value.GetType().FullName + ")";
            }

            private sealed class WeakIdentityConfiguration
            {
                private WeakIdentityConfiguration(int componentType, string displayName, string entityName,
                    string primaryIdAttribute, string nameAttribute, string uniqueIdAttribute,
                    params string[] columns)
                {
                    ComponentType = componentType;
                    DisplayName = displayName;
                    EntityName = entityName;
                    PrimaryIdAttribute = primaryIdAttribute;
                    NameAttribute = nameAttribute;
                    UniqueIdAttribute = uniqueIdAttribute;
                    Columns = columns;
                }

                internal int ComponentType { get; }
                internal string DisplayName { get; }
                internal string EntityName { get; }
                internal string PrimaryIdAttribute { get; }
                internal string NameAttribute { get; }
                internal string UniqueIdAttribute { get; }
                internal string[] Columns { get; }

                internal static WeakIdentityConfiguration For(int componentType)
                {
                    if (componentType == SavedQueryComponentType)
                        return new WeakIdentityConfiguration(componentType, "Saved Query", "savedquery",
                            "savedqueryid", "name", "savedqueryidunique", "savedqueryid", "name",
                            "returnedtypecode", "querytype", "savedqueryidunique", "componentstate", "ismanaged");
                    if (componentType == EmailTemplateComponentType)
                        return new WeakIdentityConfiguration(componentType, "Email Template", "template",
                            "templateid", "title", "templateidunique", "templateid", "title",
                            "templatetypecode", "templateidunique", "ispersonal", "languagecode",
                            "componentstate", "ismanaged");
                    if (componentType == SavedQueryVisualizationComponentType)
                        return new WeakIdentityConfiguration(componentType, "System Chart",
                            "savedqueryvisualization", "savedqueryvisualizationid", "name",
                            "savedqueryvisualizationidunique", "savedqueryvisualizationid", "name",
                            "primaryentitytypecode", "type", "charttype",
                            "savedqueryvisualizationidunique", "componentstate", "ismanaged");
                    throw new ArgumentOutOfRangeException(nameof(componentType));
                }
            }

            private void LoadReportDiagnostics(IReadOnlyList<SolutionComponentRecord> records,
                CancellationToken cancellationToken)
            {
                if (records.Count == 0 || reportDiagnosticsLoaded) return;
                reportDiagnosticsLoaded = true;
                reportSummaryComponentId = records[0].SolutionComponentId;
                var objectIds = records.Where(item => item.ObjectId.HasValue &&
                        item.ObjectId.Value != Guid.Empty)
                    .Select(item => item.ObjectId.Value).Distinct().OrderBy(item => item).ToList();
                var returnedById = objectIds.ToDictionary(item => item, item => new List<Entity>());
                var failures = new Dictionary<Guid, string>();
                var unassociatedEvidence = new Dictionary<Guid, List<string>>();
                int returnedCount = 0;
                bool countUnavailable = false;

                foreach (var batch in Batch(objectIds))
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    try
                    {
                        var query = new QueryExpression("report")
                        {
                            ColumnSet = new ColumnSet("reportid", "name", "filename", "reporttypecode",
                                "signatureid", "signaturelcid", "reportidunique", "componentstate", "ismanaged")
                        };
                        query.Criteria.AddCondition(new ConditionExpression("reportid", ConditionOperator.In,
                            batch.Select(item => (object)item).ToArray()));
                        var rows = context.Query(query);
                        cancellationToken.ThrowIfCancellationRequested();
                        returnedCount += rows.Entities.Count;

                        bool invalidResponse = false;
                        foreach (var row in rows.Entities)
                        {
                            Guid reportId;
                            if (string.Equals(row.LogicalName, "report", StringComparison.OrdinalIgnoreCase) &&
                                TryReadGuid(row, "reportid", out reportId) && batch.Contains(reportId) &&
                                (row.Id == Guid.Empty || row.Id == reportId))
                            {
                                returnedById[reportId].Add(row);
                                continue;
                            }

                            invalidResponse = true;
                            var detail = "Unassociated or conflicting returned report row: " +
                                DescribeReport(row);
                            Guid conflictingId;
                            var affectedIds = TryReadGuid(row, "reportid", out conflictingId) &&
                                batch.Contains(conflictingId) ? new[] { conflictingId } : batch;
                            foreach (var objectId in affectedIds)
                            {
                                List<string> evidence;
                                if (!unassociatedEvidence.TryGetValue(objectId, out evidence))
                                    unassociatedEvidence[objectId] = evidence = new List<string>();
                                evidence.Add(detail);
                            }
                        }

                        if (rows.MoreRecords)
                        {
                            countUnavailable = true;
                            SetReportFailures(batch,
                                "Signed Report diagnostic lookup returned an incomplete result set.", failures);
                        }
                        else if (invalidResponse)
                        {
                            countUnavailable = true;
                            SetReportFailures(batch,
                                "Signed Report diagnostic lookup returned conflicting or incomplete primary-key data.",
                                failures);
                        }
                    }
                    catch (OperationCanceledException) { throw; }
                    catch (FaultException ex)
                    {
                        cancellationToken.ThrowIfCancellationRequested();
                        countUnavailable = true;
                        SetReportFailures(batch, "Signed Report diagnostic lookup failed: " + ex.Message,
                            failures);
                    }
                }

                var uniquelyCorrelated = objectIds.Where(objectId => !failures.ContainsKey(objectId) &&
                        returnedById[objectId].Count == 1)
                    .ToDictionary(objectId => objectId, objectId => returnedById[objectId][0]);
                var signatureGroups = uniquelyCorrelated.Values.Select(row =>
                    {
                        Guid signatureId;
                        return TryReadGuid(row, "signatureid", out signatureId) ? (Guid?)signatureId : null;
                    })
                    .Where(signatureId => signatureId.HasValue)
                    .GroupBy(signatureId => signatureId.Value)
                    .ToDictionary(group => group.Key, group => group.Count());
                int correlated = 0;
                int missing = 0;
                int nonUnique = 0;
                int blankSignature = 0;
                foreach (var objectId in objectIds)
                {
                    var evidence = new List<string>();
                    string failure;
                    var matches = returnedById[objectId];
                    if (failures.TryGetValue(objectId, out failure)) evidence.Add(failure);
                    else if (matches.Count == 0)
                    {
                        missing++;
                        evidence.Add("No report row matched this solutioncomponent objectid.");
                    }
                    else if (matches.Count > 1)
                    {
                        nonUnique++;
                        evidence.Add("Multiple report rows matched this solutioncomponent objectid.");
                    }

                    foreach (var row in matches) evidence.Add(DescribeReport(row));
                    if (uniquelyCorrelated.ContainsKey(objectId))
                    {
                        correlated++;
                        var row = uniquelyCorrelated[objectId];
                        Guid signatureId;
                        if (!TryReadGuid(row, "signatureid", out signatureId)) blankSignature++;
                        else
                        {
                            int duplicateCount;
                            if (signatureGroups.TryGetValue(signatureId, out duplicateCount) && duplicateCount > 1)
                                evidence.Add("The report signatureid " + signatureId.ToString("D") +
                                    " occurs on " + duplicateCount + " uniquely correlated Type-31 reports.");
                        }
                    }
                    List<string> unassociated;
                    if (unassociatedEvidence.TryGetValue(objectId, out unassociated)) evidence.AddRange(unassociated);
                    reportDiagnostics[objectId] = evidence.AsReadOnly();
                }

                int missingObjectIds = records.Count(item => !item.ObjectId.HasValue ||
                    item.ObjectId.Value == Guid.Empty);
                var distinctSignatureIds = signatureGroups.Keys.OrderBy(item => item).ToList();
                reportSummaryEvidence = DescribeReportSummary(records.Count, objectIds.Count,
                    countUnavailable ? (int?)null : returnedCount,
                    countUnavailable ? (int?)null : correlated,
                    countUnavailable ? (int?)null : missing, missingObjectIds,
                    countUnavailable ? (int?)null : blankSignature,
                    countUnavailable ? (int?)null : signatureGroups.Values.Sum(),
                    countUnavailable ? (int?)null : distinctSignatureIds.Count,
                    countUnavailable ? (int?)null : signatureGroups.Count(group => group.Value > 1),
                    countUnavailable ? (int?)null : nonUnique,
                    countUnavailable ? null : distinctSignatureIds);
            }

            private IEnumerable<string> GetReportDiagnosticEvidence(SolutionComponentRecord record)
            {
                if (record.ComponentType != ReportComponentType || !reportDiagnosticsLoaded)
                    return new string[0];
                var result = new List<string>();
                if (!record.ObjectId.HasValue || record.ObjectId.Value == Guid.Empty)
                    result.Add("Signed Report diagnostic lookup was not attempted because objectid is unavailable.");
                else
                {
                    IReadOnlyList<string> evidence;
                    result.AddRange(reportDiagnostics.TryGetValue(record.ObjectId.Value, out evidence)
                        ? evidence : new[] { "Signed Report diagnostic lookup produced no auditable result." });
                }
                if (record.SolutionComponentId == reportSummaryComponentId &&
                    !string.IsNullOrWhiteSpace(reportSummaryEvidence))
                    result.Add(reportSummaryEvidence);
                return result;
            }

            private static void SetReportFailures(IEnumerable<Guid> objectIds, string diagnostic,
                IDictionary<Guid, string> failures)
            {
                foreach (var objectId in objectIds) failures[objectId] = diagnostic;
            }

            private static string DescribeReport(Entity row)
            {
                Guid signatureId;
                bool signed = TryReadGuid(row, "signatureid", out signatureId);
                bool complete = HasGuid(row, "reportid") && HasText(row, "name") &&
                    HasText(row, "filename") && HasOption(row, "reporttypecode") &&
                    HasOptionalGuid(row, "signatureid") && (!signed || HasInteger(row, "signaturelcid")) &&
                    HasGuid(row, "reportidunique") && HasOption(row, "componentstate") &&
                    row.Attributes.ContainsKey("ismanaged") && row.Attributes["ismanaged"] is bool;
                return (complete ? "Signed Report diagnostic lookup matched. " :
                    "Signed Report diagnostic lookup matched but returned incomplete data. ") +
                    "reportid=" + FormatReportValue(row, "reportid") +
                    "; name=" + FormatReportValue(row, "name") +
                    "; filename=" + FormatReportValue(row, "filename") +
                    "; reporttypecode=" + FormatReportValue(row, "reporttypecode") +
                    "; signatureid=" + FormatReportValue(row, "signatureid") +
                    "; signaturelcid=" + FormatReportValue(row, "signaturelcid") +
                    "; reportidunique=" + FormatReportValue(row, "reportidunique") +
                    "; componentstate=" + FormatReportValue(row, "componentstate") +
                    "; ismanaged=" + FormatReportValue(row, "ismanaged") +
                    "; candidateSignatureId=" + (signed ? "'" + signatureId.ToString("D") + "'" :
                        "(unavailable)") +
                    "; signatureLcid=" + FormatReportValue(row, "signaturelcid") +
                    ". Diagnostic evidence only; no report value is used for membership comparison.";
            }

            private static string DescribeReportSummary(int rawCount, int distinctObjectIdCount,
                int? returnedCount, int? correlatedCount, int? missingCount, int missingObjectIdCount,
                int? blankSignatureCount, int? nonblankSignatureCount, int? distinctSignatureCount,
                int? duplicateSignatureCount, int? nonUniqueObjectIdCount, IReadOnlyList<Guid> signatureIds)
            {
                return "Signed Report diagnostic summary: RawType31Count=" + rawCount +
                    "; DistinctObjectIdCount=" + distinctObjectIdCount +
                    "; ReturnedRowCount=" + FormatOptionSetCount(returnedCount) +
                    "; CorrelatedCount=" + FormatOptionSetCount(correlatedCount) +
                    "; MissingCount=" + FormatOptionSetCount(missingCount) +
                    "; MissingObjectIdRecordCount=" + missingObjectIdCount +
                    "; BlankSignatureIdCount=" + FormatOptionSetCount(blankSignatureCount) +
                    "; NonblankSignatureIdCount=" + FormatOptionSetCount(nonblankSignatureCount) +
                    "; DistinctSignatureIdCount=" + FormatOptionSetCount(distinctSignatureCount) +
                    "; DuplicateSignatureIdCount=" + FormatOptionSetCount(duplicateSignatureCount) +
                    "; NonUniqueObjectIdCount=" + FormatOptionSetCount(nonUniqueObjectIdCount) +
                    "; DistinctSignatureIds=" + (signatureIds == null ? "(unavailable)" :
                        "[" + string.Join(", ", signatureIds.Select(item => "'" + item.ToString("D") + "'")) +
                        "]") +
                    "; DistinctCandidateSignatureIds=" + (signatureIds == null ? "(unavailable)" :
                        "[" + string.Join(", ", signatureIds.Select(item => "'" + item.ToString("D") + "'")) +
                        "]") +
                    ".";
            }

            private static bool HasOptionalGuid(Entity row, string attributeName)
            {
                object value;
                return !row.Attributes.TryGetValue(attributeName, out value) || value == null || value is Guid;
            }

            private static string FormatReportValue(Entity row, string attributeName)
            {
                object value;
                if (!row.Attributes.TryGetValue(attributeName, out value)) return "(not supplied)";
                if (value == null) return "(null)";
                var option = value as OptionSetValue;
                if (option != null)
                    return AppendFormattedWorkflowEvidence(row, attributeName,
                        option.Value.ToString(System.Globalization.CultureInfo.InvariantCulture));
                if (value is Guid) return ((Guid)value).ToString("D");
                if (value is bool) return ((bool)value).ToString();
                if (value is int) return ((int)value).ToString(System.Globalization.CultureInfo.InvariantCulture);
                var text = value as string;
                if (text != null) return "'" + EscapeDiagnosticText(text) + "'";
                return "(unexpected " + value.GetType().FullName + ")";
            }

            private void LoadSystemFormDiagnostics(IReadOnlyList<SolutionComponentRecord> records,
                CancellationToken cancellationToken)
            {
                if (records.Count == 0 || systemFormDiagnosticsLoaded) return;
                systemFormDiagnosticsLoaded = true;
                systemFormSummaryComponentId = records[0].SolutionComponentId;
                var objectIds = records.Where(item => item.ObjectId.HasValue &&
                        item.ObjectId.Value != Guid.Empty)
                    .Select(item => item.ObjectId.Value).Distinct().OrderBy(item => item).ToList();
                var returnedById = objectIds.ToDictionary(item => item, item => new List<Entity>());
                var failures = new Dictionary<Guid, string>();
                var unassociatedEvidence = new Dictionary<Guid, List<string>>();
                int returnedCount = 0;
                bool countUnavailable = false;

                foreach (var batch in Batch(objectIds))
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    try
                    {
                        var query = new QueryExpression("systemform")
                        {
                            ColumnSet = new ColumnSet("formid", "uniquename", "name", "objecttypecode",
                                "type", "formidunique", "componentstate", "ismanaged")
                        };
                        query.Criteria.AddCondition(new ConditionExpression("formid", ConditionOperator.In,
                            batch.Select(item => (object)item).ToArray()));
                        var rows = context.Query(query);
                        cancellationToken.ThrowIfCancellationRequested();
                        returnedCount += rows.Entities.Count;

                        bool invalidResponse = false;
                        foreach (var row in rows.Entities)
                        {
                            Guid formId;
                            if (string.Equals(row.LogicalName, "systemform", StringComparison.OrdinalIgnoreCase) &&
                                TryReadGuid(row, "formid", out formId) && batch.Contains(formId) &&
                                (row.Id == Guid.Empty || row.Id == formId))
                            {
                                returnedById[formId].Add(row);
                                continue;
                            }

                            invalidResponse = true;
                            var detail = "Unassociated or conflicting returned systemform row: " +
                                DescribeSystemForm(row, EntityLogicalNameDiagnostic.Unverified(
                                    "(not resolved for an unassociated row)"));
                            Guid conflictingId;
                            var affectedIds = TryReadGuid(row, "formid", out conflictingId) &&
                                batch.Contains(conflictingId) ? new[] { conflictingId } : batch;
                            foreach (var objectId in affectedIds)
                            {
                                List<string> evidence;
                                if (!unassociatedEvidence.TryGetValue(objectId, out evidence))
                                    unassociatedEvidence[objectId] = evidence = new List<string>();
                                evidence.Add(detail);
                            }
                        }

                        if (rows.MoreRecords)
                        {
                            countUnavailable = true;
                            SetSystemFormFailures(batch,
                                "System Form diagnostic lookup returned an incomplete result set.", failures);
                        }
                        else if (invalidResponse)
                        {
                            countUnavailable = true;
                            SetSystemFormFailures(batch,
                                "System Form diagnostic lookup returned conflicting or incomplete primary-key data.",
                                failures);
                        }
                    }
                    catch (OperationCanceledException) { throw; }
                    catch (FaultException ex)
                    {
                        cancellationToken.ThrowIfCancellationRequested();
                        countUnavailable = true;
                        SetSystemFormFailures(batch, "System Form diagnostic lookup failed: " + ex.Message,
                            failures);
                    }
                }

                var numericObjectTypeCodes = returnedById.Values.SelectMany(item => item)
                    .Select(ReadObjectTypeCode).Where(item => item.HasValue).Select(item => item.Value)
                    .Distinct().OrderBy(item => item).ToList();
                var entityLogicalNames = ResolveEntityLogicalNameDiagnostics(numericObjectTypeCodes,
                    cancellationToken);
                int correlated = 0;
                int missing = 0;
                int nonUnique = 0;
                int blankUniqueName = 0;
                int unresolvedEntityName = 0;
                var candidateIdentities = new List<string>();
                foreach (var objectId in objectIds)
                {
                    var evidence = new List<string>();
                    string failure;
                    var matches = returnedById[objectId];
                    if (failures.TryGetValue(objectId, out failure)) evidence.Add(failure);
                    else if (matches.Count == 0)
                    {
                        missing++;
                        evidence.Add("No systemform row matched this solutioncomponent objectid.");
                    }
                    else if (matches.Count > 1)
                    {
                        nonUnique++;
                        evidence.Add("Multiple systemform rows matched this solutioncomponent objectid.");
                    }

                    foreach (var row in matches)
                        evidence.Add(DescribeSystemForm(row,
                            ResolveSystemFormEntityLogicalName(row, entityLogicalNames)));
                    if (!failures.ContainsKey(objectId) && matches.Count == 1)
                    {
                        correlated++;
                        var row = matches[0];
                        var uniqueName = row.GetAttributeValue<string>("uniquename");
                        if (string.IsNullOrWhiteSpace(uniqueName)) blankUniqueName++;
                        var entityName = ResolveSystemFormEntityLogicalName(row, entityLogicalNames);
                        if (!entityName.IsVerified) unresolvedEntityName++;
                        else if (!string.IsNullOrWhiteSpace(uniqueName))
                            candidateIdentities.Add(entityName.DisplayValue + "." + uniqueName);
                    }
                    List<string> unassociated;
                    if (unassociatedEvidence.TryGetValue(objectId, out unassociated)) evidence.AddRange(unassociated);
                    systemFormDiagnostics[objectId] = evidence.AsReadOnly();
                }

                int missingObjectIds = records.Count(item => !item.ObjectId.HasValue ||
                    item.ObjectId.Value == Guid.Empty);
                var distinctCandidates = candidateIdentities
                    .GroupBy(item => item, StringComparer.OrdinalIgnoreCase)
                    .Select(group => group.First()).OrderBy(item => item, StringComparer.OrdinalIgnoreCase)
                    .ToList();
                systemFormSummaryEvidence = DescribeSystemFormSummary(records.Count, objectIds.Count,
                    countUnavailable ? (int?)null : returnedCount,
                    countUnavailable ? (int?)null : correlated,
                    countUnavailable ? (int?)null : missing, missingObjectIds,
                    countUnavailable ? (int?)null : blankUniqueName,
                    countUnavailable ? (int?)null : unresolvedEntityName,
                    countUnavailable ? (int?)null : nonUnique,
                    countUnavailable ? null : distinctCandidates);
            }

            private IEnumerable<string> GetSystemFormDiagnosticEvidence(SolutionComponentRecord record)
            {
                if (record.ComponentType != SystemFormComponentType || !systemFormDiagnosticsLoaded)
                    return new string[0];
                var result = new List<string>();
                if (!record.ObjectId.HasValue || record.ObjectId.Value == Guid.Empty)
                    result.Add("System Form diagnostic lookup was not attempted because objectid is unavailable.");
                else
                {
                    IReadOnlyList<string> evidence;
                    result.AddRange(systemFormDiagnostics.TryGetValue(record.ObjectId.Value, out evidence)
                        ? evidence : new[] { "System Form diagnostic lookup produced no auditable result." });
                }
                if (record.SolutionComponentId == systemFormSummaryComponentId &&
                    !string.IsNullOrWhiteSpace(systemFormSummaryEvidence))
                    result.Add(systemFormSummaryEvidence);
                return result;
            }

            private static void SetSystemFormFailures(IEnumerable<Guid> objectIds, string diagnostic,
                IDictionary<Guid, string> failures)
            {
                foreach (var objectId in objectIds) failures[objectId] = diagnostic;
            }

            private static EntityLogicalNameDiagnostic ResolveSystemFormEntityLogicalName(Entity row,
                IDictionary<int, EntityLogicalNameDiagnostic> metadataNames)
            {
                object value;
                if (!row.Attributes.TryGetValue("objecttypecode", out value) || value == null)
                    return EntityLogicalNameDiagnostic.Unverified("(objecttypecode unavailable)");
                var logicalName = value as string;
                if (logicalName != null)
                    return string.IsNullOrWhiteSpace(logicalName)
                        ? EntityLogicalNameDiagnostic.Unverified("(objecttypecode entity logical name is blank)")
                        : EntityLogicalNameDiagnostic.Verified(logicalName);
                var objectTypeCode = ReadObjectTypeCode(row);
                EntityLogicalNameDiagnostic result;
                return objectTypeCode.HasValue && metadataNames.TryGetValue(objectTypeCode.Value, out result)
                    ? result : EntityLogicalNameDiagnostic.Unverified(
                        "(objecttypecode has no supported entity-name representation)");
            }

            private static string DescribeSystemForm(Entity row, EntityLogicalNameDiagnostic entityLogicalName)
            {
                var uniqueName = row.GetAttributeValue<string>("uniquename");
                var candidate = entityLogicalName.IsVerified && !string.IsNullOrWhiteSpace(uniqueName)
                    ? "'" + EscapeDiagnosticText(entityLogicalName.DisplayValue + "." + uniqueName) + "'"
                    : "(unavailable)";
                bool complete = HasGuid(row, "formid") && HasText(row, "uniquename") &&
                    HasText(row, "name") && HasSystemFormObjectType(row) && HasOption(row, "type") &&
                    HasGuid(row, "formidunique") && HasOption(row, "componentstate") &&
                    row.Attributes.ContainsKey("ismanaged") && row.Attributes["ismanaged"] is bool;
                return (complete ? "System Form diagnostic lookup matched. " :
                    "System Form diagnostic lookup matched but returned incomplete data. ") +
                    "formid=" + FormatSystemFormValue(row, "formid") +
                    "; uniquename=" + FormatSystemFormValue(row, "uniquename") +
                    "; name=" + FormatSystemFormValue(row, "name") +
                    "; objecttypecode=" + FormatSystemFormValue(row, "objecttypecode") +
                    "; entitylogicalname=" + entityLogicalName.DisplayValue +
                    "; type=" + FormatSystemFormValue(row, "type") +
                    "; formidunique=" + FormatSystemFormValue(row, "formidunique") +
                    "; componentstate=" + FormatSystemFormValue(row, "componentstate") +
                    "; ismanaged=" + FormatSystemFormValue(row, "ismanaged") +
                    "; candidateportableidentity=" + candidate +
                    ". Diagnostic evidence only; the candidate is not used for membership comparison.";
            }

            private static string DescribeSystemFormSummary(int rawCount, int distinctObjectIdCount,
                int? returnedCount, int? correlatedCount, int? missingCount, int missingObjectIdCount,
                int? blankUniqueNameCount, int? unresolvedEntityNameCount, int? nonUniqueCount,
                IReadOnlyList<string> candidateIdentities)
            {
                return "System Form diagnostic summary: RawType60MembershipCount=" + rawCount +
                    "; DistinctNonemptyObjectIdCount=" + distinctObjectIdCount +
                    "; ReturnedSystemFormRowCount=" + FormatOptionSetCount(returnedCount) +
                    "; UniqueObjectIdCorrelationCount=" + FormatOptionSetCount(correlatedCount) +
                    "; MissingRequestedObjectIdCount=" + FormatOptionSetCount(missingCount) +
                    "; MissingObjectIdRecordCount=" + missingObjectIdCount +
                    "; BlankUniqueNameCount=" + FormatOptionSetCount(blankUniqueNameCount) +
                    "; UnresolvedEntityLogicalNameCount=" + FormatOptionSetCount(unresolvedEntityNameCount) +
                    "; NonUniqueObjectIdCount=" + FormatOptionSetCount(nonUniqueCount) +
                    "; DistinctCandidatePortableIdentities=" + (candidateIdentities == null ? "(unavailable)" :
                        "[" + string.Join(", ", candidateIdentities.Select(item => "'" +
                            EscapeDiagnosticText(item) + "'")) + "]") + ".";
            }

            private static bool HasSystemFormObjectType(Entity row)
            {
                object value;
                if (!row.Attributes.TryGetValue("objecttypecode", out value) || value == null) return false;
                var text = value as string;
                return text != null ? !string.IsNullOrWhiteSpace(text) : ReadObjectTypeCode(row).HasValue;
            }

            private static string FormatSystemFormValue(Entity row, string attributeName)
            {
                object value;
                if (!row.Attributes.TryGetValue(attributeName, out value)) return "(not supplied)";
                if (value == null) return "(null)";
                var option = value as OptionSetValue;
                if (option != null)
                    return AppendFormattedWorkflowEvidence(row, attributeName,
                        option.Value.ToString(System.Globalization.CultureInfo.InvariantCulture));
                if (value is Guid) return ((Guid)value).ToString("D");
                if (value is bool) return ((bool)value).ToString();
                if (value is int) return ((int)value).ToString(System.Globalization.CultureInfo.InvariantCulture);
                var text = value as string;
                if (text != null) return "'" + EscapeDiagnosticText(text) + "'";
                return "(unexpected " + value.GetType().FullName + ")";
            }

            private void LoadSiteMapDiagnostics(IReadOnlyList<SolutionComponentRecord> records,
                CancellationToken cancellationToken)
            {
                if (records.Count == 0 || siteMapDiagnosticsLoaded) return;
                siteMapDiagnosticsLoaded = true;
                siteMapSummaryComponentId = records[0].SolutionComponentId;
                var objectIds = records.Where(item => item.ObjectId.HasValue &&
                        item.ObjectId.Value != Guid.Empty)
                    .Select(item => item.ObjectId.Value).Distinct().OrderBy(item => item).ToList();
                var returnedById = objectIds.ToDictionary(item => item, item => new List<Entity>());
                var failures = new Dictionary<Guid, string>();
                var unassociatedEvidence = new Dictionary<Guid, List<string>>();
                int returnedCount = 0;
                bool countUnavailable = false;

                foreach (var batch in Batch(objectIds))
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    try
                    {
                        var query = new QueryExpression("sitemap")
                        {
                            ColumnSet = new ColumnSet("sitemapid", "sitemapnameunique", "sitemapname",
                                "sitemapidunique", "isappaware", "componentstate", "ismanaged")
                        };
                        query.Criteria.AddCondition(new ConditionExpression("sitemapid", ConditionOperator.In,
                            batch.Select(item => (object)item).ToArray()));
                        var rows = context.Query(query);
                        cancellationToken.ThrowIfCancellationRequested();
                        returnedCount += rows.Entities.Count;

                        bool invalidResponse = false;
                        foreach (var row in rows.Entities)
                        {
                            Guid siteMapId;
                            if (string.Equals(row.LogicalName, "sitemap", StringComparison.OrdinalIgnoreCase) &&
                                TryReadGuid(row, "sitemapid", out siteMapId) && batch.Contains(siteMapId) &&
                                (row.Id == Guid.Empty || row.Id == siteMapId))
                            {
                                returnedById[siteMapId].Add(row);
                                continue;
                            }

                            invalidResponse = true;
                            var detail = "Unassociated or conflicting returned sitemap row: " +
                                DescribeSiteMap(row);
                            Guid conflictingId;
                            var affectedIds = TryReadGuid(row, "sitemapid", out conflictingId) &&
                                batch.Contains(conflictingId) ? new[] { conflictingId } : batch;
                            foreach (var objectId in affectedIds)
                            {
                                List<string> evidence;
                                if (!unassociatedEvidence.TryGetValue(objectId, out evidence))
                                    unassociatedEvidence[objectId] = evidence = new List<string>();
                                evidence.Add(detail);
                            }
                        }

                        if (rows.MoreRecords)
                        {
                            countUnavailable = true;
                            SetSiteMapFailures(batch,
                                "Site Map diagnostic lookup returned an incomplete result set.", failures);
                        }
                        else if (invalidResponse)
                        {
                            countUnavailable = true;
                            SetSiteMapFailures(batch,
                                "Site Map diagnostic lookup returned conflicting or incomplete primary-key data.",
                                failures);
                        }
                    }
                    catch (OperationCanceledException) { throw; }
                    catch (FaultException ex)
                    {
                        cancellationToken.ThrowIfCancellationRequested();
                        countUnavailable = true;
                        SetSiteMapFailures(batch, "Site Map diagnostic lookup failed: " + ex.Message,
                            failures);
                    }
                }

                int correlated = 0;
                int missing = 0;
                int nonUnique = 0;
                int blankName = 0;
                int appAware = 0;
                var candidateNames = new List<string>();
                foreach (var objectId in objectIds)
                {
                    var evidence = new List<string>();
                    string failure;
                    var matches = returnedById[objectId];
                    if (failures.TryGetValue(objectId, out failure)) evidence.Add(failure);
                    else if (matches.Count == 0)
                    {
                        missing++;
                        evidence.Add("No sitemap row matched this solutioncomponent objectid.");
                    }
                    else if (matches.Count > 1)
                    {
                        nonUnique++;
                        evidence.Add("Multiple sitemap rows matched this solutioncomponent objectid.");
                    }

                    foreach (var row in matches) evidence.Add(DescribeSiteMap(row));
                    if (!failures.ContainsKey(objectId) && matches.Count == 1)
                    {
                        correlated++;
                        var row = matches[0];
                        var name = row.GetAttributeValue<string>("sitemapnameunique");
                        if (string.IsNullOrWhiteSpace(name)) blankName++;
                        else candidateNames.Add(name);
                        if (row.GetAttributeValue<bool?>("isappaware") == true) appAware++;
                    }
                    List<string> unassociated;
                    if (unassociatedEvidence.TryGetValue(objectId, out unassociated)) evidence.AddRange(unassociated);
                    siteMapDiagnostics[objectId] = evidence.AsReadOnly();
                }

                int missingObjectIds = records.Count(item => !item.ObjectId.HasValue ||
                    item.ObjectId.Value == Guid.Empty);
                var distinctNames = candidateNames.GroupBy(item => item, StringComparer.OrdinalIgnoreCase)
                    .Select(group => group.First()).OrderBy(item => item, StringComparer.OrdinalIgnoreCase)
                    .ToList();
                siteMapSummaryEvidence = DescribeSiteMapSummary(records.Count, objectIds.Count,
                    countUnavailable ? (int?)null : returnedCount,
                    countUnavailable ? (int?)null : correlated,
                    countUnavailable ? (int?)null : missing, missingObjectIds,
                    countUnavailable ? (int?)null : blankName,
                    countUnavailable ? (int?)null : appAware,
                    countUnavailable ? (int?)null : nonUnique,
                    countUnavailable ? null : distinctNames);
            }

            private IEnumerable<string> GetSiteMapDiagnosticEvidence(SolutionComponentRecord record)
            {
                if (record.ComponentType != SiteMapComponentType || !siteMapDiagnosticsLoaded)
                    return new string[0];
                var result = new List<string>();
                if (!record.ObjectId.HasValue || record.ObjectId.Value == Guid.Empty)
                    result.Add("Site Map diagnostic lookup was not attempted because objectid is unavailable.");
                else
                {
                    IReadOnlyList<string> evidence;
                    result.AddRange(siteMapDiagnostics.TryGetValue(record.ObjectId.Value, out evidence)
                        ? evidence : new[] { "Site Map diagnostic lookup produced no auditable result." });
                }
                if (record.SolutionComponentId == siteMapSummaryComponentId &&
                    !string.IsNullOrWhiteSpace(siteMapSummaryEvidence))
                    result.Add(siteMapSummaryEvidence);
                return result;
            }

            private static void SetSiteMapFailures(IEnumerable<Guid> objectIds, string diagnostic,
                IDictionary<Guid, string> failures)
            {
                foreach (var objectId in objectIds) failures[objectId] = diagnostic;
            }

            private static string DescribeSiteMap(Entity row)
            {
                var candidate = row.GetAttributeValue<string>("sitemapnameunique");
                bool complete = HasGuid(row, "sitemapid") && HasText(row, "sitemapnameunique") &&
                    HasText(row, "sitemapname") && HasGuid(row, "sitemapidunique") &&
                    row.Attributes.ContainsKey("isappaware") && row.Attributes["isappaware"] is bool &&
                    HasOption(row, "componentstate") && row.Attributes.ContainsKey("ismanaged") &&
                    row.Attributes["ismanaged"] is bool;
                return (complete ? "Site Map diagnostic lookup matched. " :
                    "Site Map diagnostic lookup matched but returned incomplete data. ") +
                    "sitemapid=" + FormatSiteMapValue(row, "sitemapid") +
                    "; sitemapnameunique=" + FormatSiteMapValue(row, "sitemapnameunique") +
                    "; sitemapname=" + FormatSiteMapValue(row, "sitemapname") +
                    "; sitemapidunique=" + FormatSiteMapValue(row, "sitemapidunique") +
                    "; isappaware=" + FormatSiteMapValue(row, "isappaware") +
                    "; componentstate=" + FormatSiteMapValue(row, "componentstate") +
                    "; ismanaged=" + FormatSiteMapValue(row, "ismanaged") +
                    "; candidatesitemapname=" + (string.IsNullOrWhiteSpace(candidate) ? "(unavailable)" :
                        "'" + EscapeDiagnosticText(candidate) + "'") +
                    ". Diagnostic evidence only; the candidate is not used for membership comparison.";
            }

            private static string DescribeSiteMapSummary(int rawCount, int distinctObjectIdCount,
                int? returnedCount, int? correlatedCount, int? missingCount, int missingObjectIdCount,
                int? blankNameCount, int? appAwareCount, int? nonUniqueCount,
                IReadOnlyList<string> distinctNames)
            {
                return "Site Map diagnostic summary: RawType62MembershipCount=" + rawCount +
                    "; DistinctNonemptyObjectIdCount=" + distinctObjectIdCount +
                    "; ReturnedSiteMapRowCount=" + FormatOptionSetCount(returnedCount) +
                    "; UniqueObjectIdCorrelationCount=" + FormatOptionSetCount(correlatedCount) +
                    "; MissingRequestedObjectIdCount=" + FormatOptionSetCount(missingCount) +
                    "; MissingObjectIdRecordCount=" + missingObjectIdCount +
                    "; BlankSiteMapNameUniqueCount=" + FormatOptionSetCount(blankNameCount) +
                    "; AppAwareCount=" + FormatOptionSetCount(appAwareCount) +
                    "; NonUniqueObjectIdCount=" + FormatOptionSetCount(nonUniqueCount) +
                    "; DistinctCandidateSiteMapNames=" + (distinctNames == null ? "(unavailable)" :
                        "[" + string.Join(", ", distinctNames.Select(item => "'" +
                            EscapeDiagnosticText(item) + "'")) + "]") + ".";
            }

            private static string FormatSiteMapValue(Entity row, string attributeName)
            {
                object value;
                if (!row.Attributes.TryGetValue(attributeName, out value)) return "(not supplied)";
                if (value == null) return "(null)";
                var option = value as OptionSetValue;
                if (option != null)
                    return AppendFormattedWorkflowEvidence(row, attributeName,
                        option.Value.ToString(System.Globalization.CultureInfo.InvariantCulture));
                if (value is Guid) return ((Guid)value).ToString("D");
                if (value is bool) return ((bool)value).ToString();
                var text = value as string;
                if (text != null) return "'" + EscapeDiagnosticText(text) + "'";
                return "(unexpected " + value.GetType().FullName + ")";
            }

            private void LoadCanvasAppDiagnostics(IReadOnlyList<SolutionComponentRecord> records,
                CancellationToken cancellationToken)
            {
                if (records.Count == 0 || canvasAppDiagnosticsLoaded) return;
                canvasAppDiagnosticsLoaded = true;
                canvasAppSummaryComponentId = records[0].SolutionComponentId;
                var objectIds = records.Where(item => item.ObjectId.HasValue &&
                        item.ObjectId.Value != Guid.Empty)
                    .Select(item => item.ObjectId.Value).Distinct().OrderBy(item => item).ToList();
                var returnedById = objectIds.ToDictionary(item => item, item => new List<Entity>());
                var failures = new Dictionary<Guid, string>();
                var unassociatedEvidence = new Dictionary<Guid, List<string>>();
                int returnedCount = 0;
                bool countUnavailable = false;

                foreach (var batch in Batch(objectIds))
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    try
                    {
                        var query = new QueryExpression("canvasapp")
                        {
                            ColumnSet = new ColumnSet("canvasappid", "name", "displayname",
                                "uniquecanvasappid", "componentstate", "ismanaged")
                        };
                        query.Criteria.AddCondition(new ConditionExpression("canvasappid", ConditionOperator.In,
                            batch.Select(item => (object)item).ToArray()));
                        var rows = context.Query(query);
                        cancellationToken.ThrowIfCancellationRequested();
                        returnedCount += rows.Entities.Count;

                        bool invalidResponse = false;
                        foreach (var row in rows.Entities)
                        {
                            Guid canvasAppId;
                            if (string.Equals(row.LogicalName, "canvasapp", StringComparison.OrdinalIgnoreCase) &&
                                TryReadGuid(row, "canvasappid", out canvasAppId) &&
                                batch.Contains(canvasAppId) &&
                                (row.Id == Guid.Empty || row.Id == canvasAppId))
                            {
                                returnedById[canvasAppId].Add(row);
                                continue;
                            }

                            invalidResponse = true;
                            var detail = "Unassociated or conflicting returned canvasapp row: " +
                                DescribeCanvasApp(row);
                            Guid conflictingId;
                            var affectedIds = TryReadGuid(row, "canvasappid", out conflictingId) &&
                                batch.Contains(conflictingId) ? new[] { conflictingId } : batch;
                            foreach (var objectId in affectedIds)
                            {
                                List<string> evidence;
                                if (!unassociatedEvidence.TryGetValue(objectId, out evidence))
                                    unassociatedEvidence[objectId] = evidence = new List<string>();
                                evidence.Add(detail);
                            }
                        }

                        if (rows.MoreRecords)
                        {
                            countUnavailable = true;
                            SetCanvasAppFailures(batch,
                                "Canvas App diagnostic lookup returned an incomplete result set.", failures);
                        }
                        else if (invalidResponse)
                        {
                            countUnavailable = true;
                            SetCanvasAppFailures(batch,
                                "Canvas App diagnostic lookup returned conflicting or incomplete primary-key data.",
                                failures);
                        }
                    }
                    catch (OperationCanceledException) { throw; }
                    catch (FaultException ex)
                    {
                        cancellationToken.ThrowIfCancellationRequested();
                        countUnavailable = true;
                        SetCanvasAppFailures(batch, "Canvas App diagnostic lookup failed: " + ex.Message,
                            failures);
                    }
                }

                int correlated = 0;
                int missing = 0;
                int nonUnique = 0;
                int blankName = 0;
                var candidateNames = new List<string>();
                foreach (var objectId in objectIds)
                {
                    var evidence = new List<string>();
                    string failure;
                    var matches = returnedById[objectId];
                    if (failures.TryGetValue(objectId, out failure)) evidence.Add(failure);
                    else if (matches.Count == 0)
                    {
                        missing++;
                        evidence.Add("No canvasapp row matched this solutioncomponent objectid.");
                    }
                    else if (matches.Count > 1)
                    {
                        nonUnique++;
                        evidence.Add("Multiple canvasapp rows matched this solutioncomponent objectid.");
                    }

                    foreach (var row in matches) evidence.Add(DescribeCanvasApp(row));
                    if (!failures.ContainsKey(objectId) && matches.Count == 1)
                    {
                        correlated++;
                        var name = matches[0].GetAttributeValue<string>("name");
                        if (string.IsNullOrWhiteSpace(name)) blankName++;
                        else candidateNames.Add(name);
                    }
                    List<string> unassociated;
                    if (unassociatedEvidence.TryGetValue(objectId, out unassociated)) evidence.AddRange(unassociated);
                    canvasAppDiagnostics[objectId] = evidence.AsReadOnly();
                }

                int missingObjectIds = records.Count(item => !item.ObjectId.HasValue ||
                    item.ObjectId.Value == Guid.Empty);
                var distinctNames = candidateNames.GroupBy(item => item, StringComparer.OrdinalIgnoreCase)
                    .Select(group => group.First()).OrderBy(item => item, StringComparer.OrdinalIgnoreCase)
                    .ToList();
                canvasAppSummaryEvidence = DescribeCanvasAppSummary(records.Count, objectIds.Count,
                    countUnavailable ? (int?)null : returnedCount,
                    countUnavailable ? (int?)null : correlated,
                    countUnavailable ? (int?)null : missing, missingObjectIds,
                    countUnavailable ? (int?)null : blankName,
                    countUnavailable ? (int?)null : nonUnique,
                    countUnavailable ? null : distinctNames);
            }

            private IEnumerable<string> GetCanvasAppDiagnosticEvidence(SolutionComponentRecord record)
            {
                if (record.ComponentType != CanvasAppComponentType || !canvasAppDiagnosticsLoaded)
                    return new string[0];
                var result = new List<string>();
                if (!record.ObjectId.HasValue || record.ObjectId.Value == Guid.Empty)
                    result.Add("Canvas App diagnostic lookup was not attempted because objectid is unavailable.");
                else
                {
                    IReadOnlyList<string> evidence;
                    result.AddRange(canvasAppDiagnostics.TryGetValue(record.ObjectId.Value, out evidence)
                        ? evidence : new[] { "Canvas App diagnostic lookup produced no auditable result." });
                }
                if (record.SolutionComponentId == canvasAppSummaryComponentId &&
                    !string.IsNullOrWhiteSpace(canvasAppSummaryEvidence))
                    result.Add(canvasAppSummaryEvidence);
                return result;
            }

            private static void SetCanvasAppFailures(IEnumerable<Guid> objectIds, string diagnostic,
                IDictionary<Guid, string> failures)
            {
                foreach (var objectId in objectIds) failures[objectId] = diagnostic;
            }

            private static string DescribeCanvasApp(Entity row)
            {
                bool complete = HasGuid(row, "canvasappid") && HasText(row, "name") &&
                    HasText(row, "displayname") && HasText(row, "uniquecanvasappid") &&
                    HasOption(row, "componentstate") && row.Attributes.ContainsKey("ismanaged") &&
                    row.Attributes["ismanaged"] is bool;
                return (complete ? "Canvas App diagnostic lookup matched. " :
                    "Canvas App diagnostic lookup matched but returned incomplete data. ") +
                    "canvasappid=" + FormatCanvasAppValue(row, "canvasappid") +
                    "; name=" + FormatCanvasAppValue(row, "name") +
                    "; displayname=" + FormatCanvasAppValue(row, "displayname") +
                    "; uniquecanvasappid=" + FormatCanvasAppValue(row, "uniquecanvasappid") +
                    "; componentstate=" + FormatCanvasAppValue(row, "componentstate") +
                    "; ismanaged=" + FormatCanvasAppValue(row, "ismanaged") +
                    ". Diagnostic evidence only; none of these values is used as a portable comparison identity.";
            }

            private static string DescribeCanvasAppSummary(int rawCount, int distinctObjectIdCount,
                int? returnedCount, int? correlatedCount, int? missingCount, int missingObjectIdCount,
                int? blankNameCount, int? nonUniqueCount, IReadOnlyList<string> distinctNames)
            {
                return "Canvas App diagnostic summary: RawType300MembershipCount=" + rawCount +
                    "; DistinctNonemptyObjectIdCount=" + distinctObjectIdCount +
                    "; ReturnedCanvasAppRowCount=" + FormatOptionSetCount(returnedCount) +
                    "; UniqueObjectIdCorrelationCount=" + FormatOptionSetCount(correlatedCount) +
                    "; MissingRequestedObjectIdCount=" + FormatOptionSetCount(missingCount) +
                    "; MissingObjectIdRecordCount=" + missingObjectIdCount +
                    "; BlankNameCount=" + FormatOptionSetCount(blankNameCount) +
                    "; NonUniqueObjectIdCount=" + FormatOptionSetCount(nonUniqueCount) +
                    "; DistinctCandidateNames=" + (distinctNames == null ? "(unavailable)" :
                        "[" + string.Join(", ", distinctNames.Select(item => "'" +
                            EscapeDiagnosticText(item) + "'")) + "]") + ".";
            }

            private static string FormatCanvasAppValue(Entity row, string attributeName)
            {
                object value;
                if (!row.Attributes.TryGetValue(attributeName, out value)) return "(not supplied)";
                if (value == null) return "(null)";
                var option = value as OptionSetValue;
                if (option != null)
                    return AppendFormattedWorkflowEvidence(row, attributeName,
                        option.Value.ToString(System.Globalization.CultureInfo.InvariantCulture));
                if (value is Guid) return ((Guid)value).ToString("D");
                if (value is bool) return ((bool)value).ToString();
                var text = value as string;
                if (text != null) return "'" + EscapeDiagnosticText(text) + "'";
                return "(unexpected " + value.GetType().FullName + ")";
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

            private void LoadTeamTemplateDiagnostics(IReadOnlyList<SolutionComponentRecord> records,
                CancellationToken cancellationToken)
            {
                teamTemplateDiagnosticsLoaded = true;
                var objectIds = records.Where(item => item.ObjectId.HasValue &&
                        item.ObjectId.Value != Guid.Empty)
                    .Select(item => item.ObjectId.Value).Distinct().OrderBy(item => item).ToList();
                var returnedById = objectIds.ToDictionary(item => item, item => new List<Entity>());
                var failures = new Dictionary<Guid, string>();
                var unassociatedEvidence = new Dictionary<Guid, List<string>>();

                foreach (var batch in Batch(objectIds))
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    try
                    {
                        var query = new QueryExpression("teamtemplate")
                        {
                            ColumnSet = new ColumnSet("teamtemplateid", "teamtemplatename", "objecttypecode",
                                "defaultaccessrightsmask", "componentidunique", "componentstate", "ismanaged")
                        };
                        query.Criteria.AddCondition(new ConditionExpression("teamtemplateid", ConditionOperator.In,
                            batch.Select(item => (object)item).ToArray()));
                        var rows = context.Query(query);
                        cancellationToken.ThrowIfCancellationRequested();

                        bool invalidResponse = false;
                        foreach (var row in rows.Entities)
                        {
                            Guid teamTemplateId;
                            if (string.Equals(row.LogicalName, "teamtemplate", StringComparison.OrdinalIgnoreCase) &&
                                TryReadGuid(row, "teamtemplateid", out teamTemplateId) &&
                                batch.Contains(teamTemplateId) &&
                                (row.Id == Guid.Empty || row.Id == teamTemplateId))
                            {
                                returnedById[teamTemplateId].Add(row);
                                continue;
                            }

                            invalidResponse = true;
                            var detail = "Unassociated or conflicting returned row: " +
                                DescribeTeamTemplate(row, "(not resolved for an unassociated row)");
                            Guid conflictingId;
                            var affectedIds = TryReadGuid(row, "teamtemplateid", out conflictingId) &&
                                batch.Contains(conflictingId) ? new[] { conflictingId } : batch;
                            foreach (var objectId in affectedIds)
                            {
                                List<string> evidence;
                                if (!unassociatedEvidence.TryGetValue(objectId, out evidence))
                                    unassociatedEvidence[objectId] = evidence = new List<string>();
                                evidence.Add(detail);
                            }
                        }

                        if (rows.MoreRecords)
                            SetTeamTemplateFailures(batch,
                                "TeamTemplate diagnostic lookup returned an incomplete result set.", failures);
                        else if (invalidResponse)
                            SetTeamTemplateFailures(batch,
                                "TeamTemplate diagnostic lookup returned conflicting or incomplete primary-key data.",
                                failures);
                    }
                    catch (OperationCanceledException) { throw; }
                    catch (FaultException ex)
                    {
                        cancellationToken.ThrowIfCancellationRequested();
                        SetTeamTemplateFailures(batch, "TeamTemplate diagnostic lookup failed: " + ex.Message,
                            failures);
                    }
                }

                var entityLogicalNames = ResolveEntityLogicalNameDiagnostics(returnedById.Values
                    .SelectMany(item => item).Select(ReadObjectTypeCode).Where(item => item.HasValue)
                    .Select(item => item.Value).Distinct().OrderBy(item => item).ToList(), cancellationToken);
                foreach (var objectId in objectIds)
                {
                    var evidence = new List<string>();
                    string failure;
                    var matches = returnedById[objectId];
                    if (failures.TryGetValue(objectId, out failure)) evidence.Add(failure);
                    else if (matches.Count == 0)
                        evidence.Add("No teamtemplate row matched this solutioncomponent objectid.");
                    else if (matches.Count > 1)
                        evidence.Add("Multiple teamtemplate rows matched this solutioncomponent objectid.");

                    foreach (var row in matches)
                    {
                        var objectTypeCode = ReadObjectTypeCode(row);
                        EntityLogicalNameDiagnostic entityLogicalName;
                        if (!objectTypeCode.HasValue ||
                            !entityLogicalNames.TryGetValue(objectTypeCode.Value, out entityLogicalName))
                            entityLogicalName = EntityLogicalNameDiagnostic.Unverified(
                                "(objecttypecode unavailable)");
                        evidence.Add(DescribeTeamTemplate(row, entityLogicalName.DisplayValue));
                        if (!failures.ContainsKey(objectId) && matches.Count == 1 &&
                            IsCompleteTeamTemplate(row) && entityLogicalName.IsVerified)
                            verifiedTeamTemplates.Add(objectId);
                    }
                    List<string> unassociated;
                    if (unassociatedEvidence.TryGetValue(objectId, out unassociated)) evidence.AddRange(unassociated);
                    teamTemplateDiagnostics[objectId] = evidence.AsReadOnly();
                }
            }

            private IDictionary<int, EntityLogicalNameDiagnostic> ResolveEntityLogicalNameDiagnostics(
                IReadOnlyList<int> objectTypeCodes, CancellationToken cancellationToken)
            {
                var results = new Dictionary<int, EntityLogicalNameDiagnostic>();
                foreach (var batch in Batch(objectTypeCodes))
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    try
                    {
                        var query = new EntityQueryExpression
                        {
                            Properties = new MetadataPropertiesExpression("ObjectTypeCode", "LogicalName"),
                            Criteria = new MetadataFilterExpression(LogicalOperator.Or)
                        };
                        foreach (var objectTypeCode in batch)
                            query.Criteria.Conditions.Add(new MetadataConditionExpression("ObjectTypeCode",
                                MetadataConditionOperator.Equals, objectTypeCode));
                        var response = context.Execute(new RetrieveMetadataChangesRequest { Query = query })
                            as RetrieveMetadataChangesResponse;
                        var metadata = response?.EntityMetadata ?? new EntityMetadataCollection();
                        cancellationToken.ThrowIfCancellationRequested();
                        if (metadata.Any(item => !item.ObjectTypeCode.HasValue ||
                            !batch.Contains(item.ObjectTypeCode.Value)))
                        {
                            SetEntityLogicalNameDiagnostics(batch,
                                "(metadata lookup returned conflicting ObjectTypeCode data)", results);
                            continue;
                        }

                        foreach (var objectTypeCode in batch)
                        {
                            var matches = metadata.Where(item => item.ObjectTypeCode == objectTypeCode).ToList();
                            if (matches.Count == 0)
                                results[objectTypeCode] = EntityLogicalNameDiagnostic.Unverified(
                                    "(no entity metadata match)");
                            else if (matches.Count > 1)
                                results[objectTypeCode] = EntityLogicalNameDiagnostic.Unverified(
                                    "(multiple entity metadata matches: " +
                                    string.Join(", ", matches.Select(item => item.LogicalName ?? "(blank)")) + ")");
                            else
                                results[objectTypeCode] = string.IsNullOrWhiteSpace(matches[0].LogicalName)
                                    ? EntityLogicalNameDiagnostic.Unverified(
                                        "(entity metadata LogicalName is blank)")
                                    : EntityLogicalNameDiagnostic.Verified(matches[0].LogicalName);
                        }
                    }
                    catch (OperationCanceledException) { throw; }
                    catch (FaultException ex)
                    {
                        cancellationToken.ThrowIfCancellationRequested();
                        SetEntityLogicalNameDiagnostics(batch,
                            "(entity metadata lookup failed: " + ex.Message + ")", results);
                    }
                }
                return results;
            }

            private IEnumerable<string> GetTeamTemplateDiagnosticEvidence(SolutionComponentRecord record)
            {
                if (record.ComponentType != TeamTemplateComponentType || !teamTemplateDiagnosticsLoaded)
                    return new string[0];
                if (!record.ObjectId.HasValue || record.ObjectId.Value == Guid.Empty)
                    return new[] { "TeamTemplate diagnostic lookup was not attempted because objectid is unavailable." };
                IReadOnlyList<string> evidence;
                return teamTemplateDiagnostics.TryGetValue(record.ObjectId.Value, out evidence)
                    ? evidence : new[] { "TeamTemplate diagnostic lookup produced no auditable result." };
            }

            private static void SetTeamTemplateFailures(IEnumerable<Guid> objectIds, string diagnostic,
                IDictionary<Guid, string> failures)
            {
                foreach (var objectId in objectIds) failures[objectId] = diagnostic;
            }

            private static void SetEntityLogicalNameDiagnostics(IEnumerable<int> objectTypeCodes, string diagnostic,
                IDictionary<int, EntityLogicalNameDiagnostic> results)
            {
                foreach (var objectTypeCode in objectTypeCodes)
                    results[objectTypeCode] = EntityLogicalNameDiagnostic.Unverified(diagnostic);
            }

            private static string DescribeTeamTemplate(Entity row, string entityLogicalName)
            {
                return (IsCompleteTeamTemplate(row) ? "TeamTemplate diagnostic lookup matched. " :
                    "TeamTemplate diagnostic lookup matched but returned incomplete data. ") +
                    "teamtemplateid=" + FormatTeamTemplateValue(row, "teamtemplateid") +
                    "; teamtemplatename=" + FormatTeamTemplateValue(row, "teamtemplatename") +
                    "; objecttypecode=" + FormatTeamTemplateValue(row, "objecttypecode") +
                    "; entitylogicalname=" + entityLogicalName +
                    "; defaultaccessrightsmask=" + FormatTeamTemplateValue(row, "defaultaccessrightsmask") +
                    "; componentidunique=" + FormatTeamTemplateValue(row, "componentidunique") +
                    "; componentstate=" + FormatTeamTemplateValue(row, "componentstate") +
                    "; ismanaged=" + FormatTeamTemplateValue(row, "ismanaged") +
                    ". Diagnostic evidence only; none of these values is used as a portable comparison identity.";
            }

            private static bool IsCompleteTeamTemplate(Entity row)
            {
                return HasGuid(row, "teamtemplateid") && HasText(row, "teamtemplatename") &&
                    ReadObjectTypeCode(row).HasValue && HasInteger(row, "defaultaccessrightsmask") &&
                    HasGuid(row, "componentidunique") && HasOption(row, "componentstate") &&
                    row.Attributes.ContainsKey("ismanaged") && row.Attributes["ismanaged"] is bool;
            }

            private static string FormatTeamTemplateValue(Entity row, string attributeName)
            {
                object value;
                if (!row.Attributes.TryGetValue(attributeName, out value)) return "(not supplied)";
                if (value == null) return "(null)";
                var option = value as OptionSetValue;
                if (option != null)
                    return AppendFormattedWorkflowEvidence(row, attributeName,
                        option.Value.ToString(System.Globalization.CultureInfo.InvariantCulture));
                if (value is Guid) return ((Guid)value).ToString("D");
                if (value is bool) return ((bool)value).ToString();
                if (value is int) return ((int)value).ToString(System.Globalization.CultureInfo.InvariantCulture);
                if (value is long) return ((long)value).ToString(System.Globalization.CultureInfo.InvariantCulture);
                var text = value as string;
                if (text != null) return "'" + EscapeDiagnosticText(text) + "'";
                return "(unexpected " + value.GetType().FullName + ")";
            }

            private static bool HasInteger(Entity row, string attributeName)
            {
                object value;
                return row.Attributes.TryGetValue(attributeName, out value) &&
                    (value is int || value is long);
            }

            private void ResolveAppModuleBatches(IReadOnlyList<LookupKey> keys,
                CancellationToken cancellationToken)
            {
                foreach (var batch in Batch(keys))
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    try
                    {
                        var query = new QueryExpression("appmodule")
                        {
                            ColumnSet = new ColumnSet("appmoduleid", "uniquename", "name",
                                "appmoduleidunique", "componentstate", "ismanaged")
                        };
                        query.Criteria.AddCondition(new ConditionExpression("appmoduleid", ConditionOperator.In,
                            batch.Select(item => (object)item.ObjectId).ToArray()));
                        var rows = context.Query(query);
                        cancellationToken.ThrowIfCancellationRequested();
                        if (rows.MoreRecords)
                        {
                            SetAppModuleResults(batch, IdentityResolutionStatus.Ambiguous,
                                "AppModule identity lookup returned an incomplete result set.");
                            continue;
                        }

                        var returned = new List<KeyValuePair<Guid, Entity>>();
                        var requested = new HashSet<Guid>(batch.Select(item => item.ObjectId));
                        bool invalidResponse = false;
                        foreach (var row in rows.Entities)
                        {
                            Guid appModuleId;
                            if (!string.Equals(row.LogicalName, "appmodule", StringComparison.OrdinalIgnoreCase) ||
                                !TryReadGuid(row, "appmoduleid", out appModuleId) ||
                                !requested.Contains(appModuleId) ||
                                row.Id != Guid.Empty && row.Id != appModuleId)
                            {
                                invalidResponse = true;
                                break;
                            }
                            returned.Add(new KeyValuePair<Guid, Entity>(appModuleId, row));
                        }
                        if (invalidResponse)
                        {
                            SetAppModuleResults(batch, IdentityResolutionStatus.Ambiguous,
                                "AppModule identity lookup returned conflicting or incomplete primary-key data.");
                            continue;
                        }

                        foreach (var key in batch)
                        {
                            var matches = returned.Where(item => item.Key == key.ObjectId)
                                .Select(item => item.Value).ToList();
                            if (matches.Count == 0)
                                identityCache[key] = ResolutionValue.Unresolved(key.Kind,
                                    "No appmodule row matched the component object ID.");
                            else if (matches.Count > 1)
                                identityCache[key] = ResolutionValue.Ambiguous(key.Kind,
                                    "An appmodule object ID lookup returned multiple records.",
                                    matches.Select(DescribeAppModule));
                            else
                            {
                                var row = matches[0];
                                var evidence = new[] { DescribeAppModule(row) };
                                var uniqueName = row.GetAttributeValue<string>("uniquename");
                                identityCache[key] = string.IsNullOrWhiteSpace(uniqueName)
                                    ? ResolutionValue.Unresolved(key.Kind,
                                        "The appmodule record has no nonblank uniquename.", evidence)
                                    : ResolutionValue.FromKey(key.Kind, uniqueName,
                                        diagnosticEvidence: evidence);
                            }
                        }
                    }
                    catch (OperationCanceledException) { throw; }
                    catch (FaultException ex)
                    {
                        cancellationToken.ThrowIfCancellationRequested();
                        SetAppModuleResults(batch, IdentityResolutionStatus.Unresolved,
                            "AppModule identity lookup failed: " + ex.Message);
                    }
                }
            }

            private void SetAppModuleResults(IEnumerable<LookupKey> keys, IdentityResolutionStatus status,
                string diagnostic)
            {
                foreach (var key in keys)
                    identityCache[key] = status == IdentityResolutionStatus.Ambiguous
                        ? ResolutionValue.Ambiguous(key.Kind, diagnostic)
                        : ResolutionValue.Unresolved(key.Kind, diagnostic);
            }

            private static string DescribeAppModule(Entity row)
            {
                bool complete = HasGuid(row, "appmoduleid") && HasText(row, "uniquename") &&
                    HasText(row, "name") && HasGuid(row, "appmoduleidunique") &&
                    HasOption(row, "componentstate") && row.Attributes.ContainsKey("ismanaged") &&
                    row.Attributes["ismanaged"] is bool;
                return (complete ? "AppModule identity lookup matched. " :
                    "AppModule identity lookup matched but returned incomplete diagnostic data. ") +
                    "appmoduleid=" + FormatAppModuleValue(row, "appmoduleid") +
                    "; uniquename=" + FormatAppModuleValue(row, "uniquename") +
                    "; name=" + FormatAppModuleValue(row, "name") +
                    "; appmoduleidunique=" + FormatAppModuleValue(row, "appmoduleidunique") +
                    "; componentstate=" + FormatAppModuleValue(row, "componentstate") +
                    "; ismanaged=" + FormatAppModuleValue(row, "ismanaged") +
                    ". Only nonblank uniquename is used as the portable comparison identity; all fields in this evidence are diagnostic.";
            }

            private static string FormatAppModuleValue(Entity row, string attributeName)
            {
                object value;
                if (!row.Attributes.TryGetValue(attributeName, out value)) return "(not supplied)";
                if (value == null) return "(null)";
                var option = value as OptionSetValue;
                if (option != null)
                    return AppendFormattedWorkflowEvidence(row, attributeName,
                        option.Value.ToString(System.Globalization.CultureInfo.InvariantCulture));
                if (value is Guid) return ((Guid)value).ToString("D");
                if (value is bool) return ((bool)value).ToString();
                var text = value as string;
                if (text != null) return "'" + EscapeDiagnosticText(text) + "'";
                return "(unexpected " + value.GetType().FullName + ")";
            }

            private static bool TryReadGuid(Entity row, string attributeName, out Guid value)
            {
                object raw;
                if (row.Attributes.TryGetValue(attributeName, out raw) && raw is Guid)
                {
                    value = (Guid)raw;
                    return value != Guid.Empty;
                }
                value = Guid.Empty;
                return false;
            }

            private static bool HasGuid(Entity row, string attributeName)
            {
                Guid value;
                return TryReadGuid(row, attributeName, out value);
            }

            private static bool HasText(Entity row, string attributeName)
            {
                object value;
                return row.Attributes.TryGetValue(attributeName, out value) &&
                    !string.IsNullOrWhiteSpace(value as string);
            }

            private static bool HasOption(Entity row, string attributeName)
            {
                object value;
                return row.Attributes.TryGetValue(attributeName, out value) && value is OptionSetValue;
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
                string diagnostic, string kind = null, IEnumerable<string> diagnosticEvidence = null) =>
                new ComponentIdentity(record, status, diagnostic: diagnostic, componentTypeKey: kind,
                    diagnosticEvidence: diagnosticEvidence);

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

            private sealed class EntityLogicalNameDiagnostic
            {
                private EntityLogicalNameDiagnostic(bool isVerified, string displayValue)
                {
                    IsVerified = isVerified;
                    DisplayValue = displayValue;
                }

                public bool IsVerified { get; }
                public string DisplayValue { get; }
                public static EntityLogicalNameDiagnostic Verified(string logicalName) =>
                    new EntityLogicalNameDiagnostic(true, logicalName);
                public static EntityLogicalNameDiagnostic Unverified(string diagnostic) =>
                    new EntityLogicalNameDiagnostic(false, diagnostic);
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
                private ResolutionValue(string kind, IdentityResolutionStatus status, string key, string diagnostic,
                    IEnumerable<string> diagnosticEvidence)
                {
                    Kind = kind;
                    Status = status;
                    Key = key;
                    Diagnostic = diagnostic;
                    DiagnosticEvidence = new List<string>(diagnosticEvidence ?? new string[0]).AsReadOnly();
                }
                public string Kind { get; }
                public IdentityResolutionStatus Status { get; }
                public string Key { get; }
                public string Diagnostic { get; }
                public IReadOnlyList<string> DiagnosticEvidence { get; }
                public ComponentIdentity ToIdentity(SolutionComponentRecord record) =>
                    new ComponentIdentity(record, Status, Key, Diagnostic, Kind,
                        diagnosticEvidence: DiagnosticEvidence);
                public static ResolutionValue FromKey(string kind, string key, string diagnostic = null,
                    IEnumerable<string> diagnosticEvidence = null) =>
                    string.IsNullOrWhiteSpace(key)
                    ? Unresolved(kind, "No strong portable identity was available; display names and local GUIDs are not used.")
                    : new ResolutionValue(kind, IdentityResolutionStatus.Resolved, key, diagnostic,
                        diagnosticEvidence);
                public static ResolutionValue Unresolved(string kind, string diagnostic,
                    IEnumerable<string> diagnosticEvidence = null) =>
                    new ResolutionValue(kind, IdentityResolutionStatus.Unresolved, null, diagnostic,
                        diagnosticEvidence);
                public static ResolutionValue Ambiguous(string kind, string diagnostic,
                    IEnumerable<string> diagnosticEvidence = null) =>
                    new ResolutionValue(kind, IdentityResolutionStatus.Ambiguous, null, diagnostic,
                        diagnosticEvidence);
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
