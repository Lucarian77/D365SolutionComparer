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

        // Published componenttype choices already assigned to non-Connection-Reference kinds.
        private static readonly HashSet<int> KnownNonConnectionReferenceTypes = new HashSet<int>
        {
            1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 16, 17, 18,
            20, 21, 22, 23, 24, 25, 26, 29, 31, 32, 33, 34, 35, 36, 37, 38, 39,
            44, 45, 46, 47, 48, 49, 50, 52, 53, 55, 59, 60, 61, 62, 63, 64, 65, 66, 68,
            70, 71, 90, 91, 92, 93, 95, 150, 151, 152, 153, 154, 155, 161, 162, 165, 166,
            201, 202, 203, 204, 205, 206, 207, 208, 210, 300, 371, 372, 380, 381,
            400, 401, 402, 430, 431, 432
        };

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
            private bool connectionMappingLoaded;
            private int? connectionTypeCode;
            private string connectionMappingDiagnostic;
            private IdentityResolutionStatus connectionMappingStatus = IdentityResolutionStatus.Unresolved;

            public ResolutionContext(DataverseReadContext context) { this.context = context; }

            public ComponentIdentity Resolve(SolutionComponentRecord record, CancellationToken cancellationToken)
            {
                cancellationToken.ThrowIfCancellationRequested();
                string kind;
                var immediate = Classify(record, cancellationToken, out kind);
                if (immediate != null) return immediate;
                var cacheKey = new LookupKey(kind, record.ObjectId.Value);
                ResolutionValue value;
                if (!identityCache.TryGetValue(cacheKey, out value))
                {
                    value = ResolveOne(cacheKey, cancellationToken);
                    identityCache.Add(cacheKey, value);
                }
                return value.ToIdentity(record);
            }

            public IReadOnlyList<ComponentIdentity> ResolveAll(IReadOnlyList<ComponentIdentity> components,
                CancellationToken cancellationToken)
            {
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
                    if (IsEntityBacked(group.Key)) ResolveEntityBatches(group.Key, keys, cancellationToken);
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
                        if (KnownNonConnectionReferenceTypes.Contains(record.ComponentType))
                            return Unknown(record, IdentityResolutionStatus.Unsupported,
                                "No identity resolver supports this known component type.");
                        LoadConnectionMapping();
                        cancellationToken.ThrowIfCancellationRequested();
                        if (connectionMappingDiagnostic != null)
                            return Unknown(record, connectionMappingStatus, connectionMappingDiagnostic);
                        if (!connectionTypeCode.HasValue || connectionTypeCode.Value != record.ComponentType)
                            return Unknown(record, IdentityResolutionStatus.Unsupported,
                                "No identity resolver supports this component type.");
                        kind = "connectionreference";
                        break;
                }
                if (!record.ObjectId.HasValue || record.ObjectId.Value == Guid.Empty)
                    return Unknown(record, IdentityResolutionStatus.Unresolved,
                        "The raw component has no usable object ID.", kind);
                return null;
            }

            private ResolutionValue ResolveOne(LookupKey key, CancellationToken cancellationToken)
            {
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
                            Criteria = new MetadataFilterExpression(LogicalOperator.And)
                        };
                        query.Criteria.Conditions.Add(new MetadataConditionExpression("MetadataId",
                            MetadataConditionOperator.In, batch.Select(item => item.ObjectId).Cast<object>().ToArray()));
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

            private static IEnumerable<IReadOnlyList<LookupKey>> Batch(IReadOnlyList<LookupKey> items)
            {
                for (int offset = 0; offset < items.Count; offset += BatchSize)
                    yield return items.Skip(offset).Take(Math.Min(BatchSize, items.Count - offset)).ToList();
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
                        else if (KnownNonConnectionReferenceTypes.Contains(connectionTypeCode.Value))
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
                public static ResolutionValue FromKey(string kind, string key) => string.IsNullOrWhiteSpace(key)
                    ? Unresolved(kind, "No strong portable identity was available; display names and local GUIDs are not used.")
                    : new ResolutionValue(kind, IdentityResolutionStatus.Resolved, key, null);
                public static ResolutionValue Unresolved(string kind, string diagnostic) =>
                    new ResolutionValue(kind, IdentityResolutionStatus.Unresolved, null, diagnostic);
                public static ResolutionValue Ambiguous(string kind, string diagnostic) =>
                    new ResolutionValue(kind, IdentityResolutionStatus.Ambiguous, null, diagnostic);
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
