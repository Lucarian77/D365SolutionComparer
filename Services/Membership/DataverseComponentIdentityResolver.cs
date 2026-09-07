using System;
using System.Collections.Generic;
using System.Threading;
using D365SolutionComparer.Models.Identity;
using D365SolutionComparer.Models.Membership;
using D365SolutionComparer.Services.Contracts;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Query;

namespace D365SolutionComparer.Services.Membership
{
    /// <summary>Published identity metadata only. No display-name or environment-local GUID fallback.</summary>
    public sealed class DataverseComponentIdentityResolver : IComponentIdentityResolver
    {
        // Published componenttype choices already assigned to non-Connection-Reference kinds.
        // https://learn.microsoft.com/en-us/power-apps/developer/data-platform/reference/entities/solutioncomponent#componenttype
        // Keep this set current when adding support for newly documented built-in component types.
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
            var context = new ResolutionContext(new DataverseReadContext(service, environment, cancellationToken));
            return context.Resolve(component, cancellationToken);
        }

        /// <summary>Mapping discovery is cached only for this snapshot operation, never across environments.</summary>
        public MembershipSnapshot ResolveSnapshot(IOrganizationService service, MembershipSnapshot snapshot, CancellationToken cancellationToken)
        {
            if (snapshot == null) throw new ArgumentNullException(nameof(snapshot));
            cancellationToken.ThrowIfCancellationRequested();
            if (snapshot.State != MembershipSnapshotState.Complete) return snapshot;
            var context = new ResolutionContext(new DataverseReadContext(service, snapshot.Environment, cancellationToken));
            var resolved = new List<ComponentIdentity>();
            foreach (var component in snapshot.Components)
                resolved.Add(context.Resolve(component.Record, cancellationToken));
            cancellationToken.ThrowIfCancellationRequested();
            return MembershipSnapshot.Complete(snapshot.Solution, resolved, snapshot.CapturedAt);
        }

        private sealed class ResolutionContext
        {
            private readonly DataverseReadContext context;
            private bool connectionMappingLoaded;
            private int? connectionTypeCode;
            private string connectionMappingDiagnostic;
            private IdentityResolutionStatus connectionMappingStatus = IdentityResolutionStatus.Unresolved;

            public ResolutionContext(DataverseReadContext context) { this.context = context; }

            public ComponentIdentity Resolve(SolutionComponentRecord record, CancellationToken cancellationToken)
            {
                cancellationToken.ThrowIfCancellationRequested();
                string kind;
                switch (record.ComponentType)
                {
                    case 1: kind = "table"; break;
                    case 2: kind = "column"; break;
                    case 10: kind = "relationship"; break;
                    case 61: kind = "webresource"; break;
                    case 29: kind = "process"; break;
                    case 20: kind = "securityrole"; break;
                    case 380: kind = "environmentvariabledefinition"; break;
                    case 3:  // Relationship
                    case 11: // Entity Relationship Role
                    case 12: // Entity Relationship Relationships
                        return Unknown(record, IdentityResolutionStatus.Unsupported, "This relationship component type is not supported in Phase 2A.");
                    default:
                        if (KnownNonConnectionReferenceTypes.Contains(record.ComponentType))
                            return Unknown(record, IdentityResolutionStatus.Unsupported, "No identity resolver supports this known component type.");
                        LoadConnectionMapping();
                        cancellationToken.ThrowIfCancellationRequested();
                        if (connectionMappingDiagnostic != null)
                            return Unknown(record, connectionMappingStatus, connectionMappingDiagnostic);
                        if (!connectionTypeCode.HasValue || connectionTypeCode.Value != record.ComponentType)
                            return Unknown(record, IdentityResolutionStatus.Unsupported, "No identity resolver supports this component type.");
                        kind = "connectionreference";
                        break;
                }
                if (!record.ObjectId.HasValue || record.ObjectId.Value == Guid.Empty)
                    return Unknown(record, IdentityResolutionStatus.Unresolved, "The raw component has no usable object ID.", kind);
                try
                {
                    var key = ResolveKey(record, kind);
                    cancellationToken.ThrowIfCancellationRequested();
                    return string.IsNullOrWhiteSpace(key)
                        ? Unknown(record, IdentityResolutionStatus.Unresolved, "No strong portable identity was available; display names and local GUIDs are not used.", kind)
                        : new ComponentIdentity(record, IdentityResolutionStatus.Resolved, key, componentTypeKey: kind);
                }
                catch (OperationCanceledException) { throw; }
                catch (AmbiguousIdentityException ex)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    return Unknown(record, IdentityResolutionStatus.Ambiguous, ex.Message, kind);
                }
                catch (System.ServiceModel.FaultException ex)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    return Unknown(record, IdentityResolutionStatus.Unresolved, "Identity read failed: " + ex.Message, kind);
                }
            }

            private string ResolveKey(SolutionComponentRecord record, string kind)
            {
                var id = record.ObjectId.Value;
                if (kind == "table")
                {
                    var response = context.Execute(new RetrieveEntityRequest { MetadataId = id, EntityFilters = Microsoft.Xrm.Sdk.Metadata.EntityFilters.Entity, RetrieveAsIfPublished = false }) as RetrieveEntityResponse;
                    return response?.EntityMetadata?.LogicalName;
                }
                if (kind == "column")
                {
                    var response = context.Execute(new RetrieveAttributeRequest { MetadataId = id, RetrieveAsIfPublished = false }) as RetrieveAttributeResponse;
                    var metadata = response?.AttributeMetadata;
                    return metadata == null || string.IsNullOrWhiteSpace(metadata.EntityLogicalName) || string.IsNullOrWhiteSpace(metadata.LogicalName)
                        ? null : metadata.EntityLogicalName + "." + metadata.LogicalName;
                }
                if (kind == "relationship")
                {
                    var response = context.Execute(new RetrieveRelationshipRequest { MetadataId = id, RetrieveAsIfPublished = false }) as RetrieveRelationshipResponse;
                    return response?.RelationshipMetadata?.SchemaName;
                }
                string table, primaryId, identityAttribute;
                switch (kind)
                {
                    case "webresource": table = "webresource"; primaryId = "webresourceid"; identityAttribute = "name"; break;
                    case "process": table = "workflow"; primaryId = "workflowid"; identityAttribute = "uniquename"; break;
                    case "securityrole": table = "role"; primaryId = "roleid"; identityAttribute = "roletemplateid"; break;
                    case "environmentvariabledefinition": table = "environmentvariabledefinition"; primaryId = "environmentvariabledefinitionid"; identityAttribute = "schemaname"; break;
                    default: table = "connectionreference"; primaryId = "connectionreferenceid"; identityAttribute = "connectionreferencelogicalname"; break;
                }
                var query = new QueryExpression(table) { ColumnSet = new ColumnSet(primaryId, identityAttribute), TopCount = 2 };
                query.Criteria.AddCondition(primaryId, ConditionOperator.Equal, id);
                var rows = context.Query(query);
                if (rows.MoreRecords || rows.Entities.Count > 1) throw new AmbiguousIdentityException("An object ID lookup returned multiple records.");
                if (rows.Entities.Count == 0) return null;
                if (rows.Entities[0].Id != id) throw new InvalidOperationException("An object ID lookup returned a different object.");
                if (kind == "securityrole")
                {
                    var template = rows.Entities[0].GetAttributeValue<EntityReference>(identityAttribute);
                    return template == null || template.Id == Guid.Empty ? null : template.Id.ToString("D");
                }
                return rows.Entities[0].GetAttributeValue<string>(identityAttribute);
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
                catch (System.ServiceModel.FaultException ex)
                {
                    connectionMappingDiagnostic = "Connection-reference component type mapping is unavailable: " + ex.Message;
                }
            }

            private static ComponentIdentity Unknown(SolutionComponentRecord record, IdentityResolutionStatus status, string diagnostic, string kind = null)
            {
                return new ComponentIdentity(record, status, diagnostic: diagnostic, componentTypeKey: kind);
            }

            private sealed class AmbiguousIdentityException : Exception
            {
                public AmbiguousIdentityException(string message) : base(message) { }
            }
        }
    }
}
