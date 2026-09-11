using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.Serialization;
using System.ServiceModel;
using System.Threading;
using D365SolutionComparer.Models.ComponentDetails;
using D365SolutionComparer.Models.Membership;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Metadata.Query;
using Microsoft.Xrm.Sdk.Query;

namespace D365SolutionComparer.Services.Membership
{
    /// <summary>
    /// Full published child inventories of the tables represented by authoritative membership.
    /// Filters belong exclusively to EntityMetadata. Child IDs are correlated locally, never
    /// sent as AttributeQuery/RelationshipQuery conditions. An operation commits its inventory
    /// only after every parent batch succeeds, so partial reads cannot establish a unique match.
    /// </summary>
    internal static class ParentEntityMetadataReader
    {
        internal const int BatchSize = 200;
        internal static readonly string[] Collections = { "Attributes", "OneToManyRelationships",
            "ManyToOneRelationships", "ManyToManyRelationships" };

        // SDK source properties, not flattened names from the definition equality contract.
        internal static readonly string[] AttributeProperties = { "MetadataId", "LogicalName",
            "EntityLogicalName", "SchemaName", "AttributeType", "AttributeTypeName", "RequiredLevel",
            "IsAuditEnabled", "IsSecured", "IsValidForAdvancedFind", "IsCustomizable", "MaxLength",
            "MinValue", "MaxValue", "Format", "Precision", "PrecisionSource", "DateTimeBehavior",
            "Targets", "OptionSet" };
        internal static readonly string[] RelationshipProperties = { "MetadataId", "SchemaName",
            "RelationshipType", "IsCustomizable", "ReferencedEntity", "ReferencedAttribute",
            "ReferencingEntity", "ReferencingAttribute", "CascadeConfiguration",
            "Entity1LogicalName", "Entity2LogicalName", "IntersectEntityName" };

        internal static void Ensure(DataverseReadContext context,
            IReadOnlyList<ComponentIdentity> components, CancellationToken cancellationToken)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (context.MetadataCache.ParentMetadata != null) return;
            var children = components.Where(item => item.Record.ComponentType == 2 ||
                item.Record.ComponentType == 10).ToList();
            if (children.Count == 0) return;
            var inventory = new ParentEntityMetadataInventory();
            var parentIds = components.Where(item => item.Record.ComponentType == 1 &&
                    item.Record.ObjectId.HasValue && item.Record.ObjectId.Value != Guid.Empty)
                .Select(item => item.Record.ObjectId.Value).Distinct().OrderBy(id => id).ToList();
            inventory.ParentIds.UnionWith(parentIds);
            var conditions = parentIds.Select(id => new MetadataConditionExpression("MetadataId",
                MetadataConditionOperator.Equals, id)).ToList();

            // Standalone definition callers can have a previously verified Column portable key.
            // This is an approved resolver result, not a name guessed from a raw GUID.
            if (conditions.Count == 0)
                conditions.AddRange(children.Where(item => item.Record.ComponentType == 2 &&
                        item.Status == IdentityResolutionStatus.Resolved &&
                        item.ComponentTypeKey == ComponentSemanticKinds.Column &&
                        item.ComparisonKey != null && item.ComparisonKey.IndexOf('.') > 0)
                    .Select(item => item.ComparisonKey.Substring(0, item.ComparisonKey.IndexOf('.')))
                    .Distinct(StringComparer.OrdinalIgnoreCase).OrderBy(name => name, StringComparer.OrdinalIgnoreCase)
                    .Select(name => new MetadataConditionExpression("LogicalName",
                        MetadataConditionOperator.Equals, name)));
            if (conditions.Count == 0)
            {
                inventory.Fail("No verified parent Table membership or previously resolved Column parent was available.");
                context.MetadataCache.ParentMetadata = inventory;
                return;
            }

            var returned = new List<EntityMetadata>();
            inventory.ScopeEvidence = "Parent metadata scope: tables=" + conditions.Count +
                "; entity batches=" + ((conditions.Count + BatchSize - 1) / BatchSize) +
                "; collections=" + string.Join(", ", Collections) + ".";
            for (int offset = 0; offset < conditions.Count; offset += BatchSize)
            {
                cancellationToken.ThrowIfCancellationRequested();
                var batch = conditions.Skip(offset).Take(BatchSize).ToList();
                string audit = "Column / Relationship parent metadata batch " + (offset / BatchSize + 1) +
                    "; parent entities=" + batch.Count + "; collections=" + string.Join(", ", Collections);
                try
                {
                    var query = new EntityQueryExpression
                    {
                        Properties = new MetadataPropertiesExpression(new[] { "MetadataId", "LogicalName" }
                            .Concat(ComponentDefinitionContractCatalog.For(ComponentSemanticKinds.Table)
                                .ComparableProperties).Concat(Collections).Distinct(StringComparer.Ordinal).ToArray()),
                        Criteria = new MetadataFilterExpression(LogicalOperator.Or),
                        AttributeQuery = new AttributeQueryExpression
                        {
                            Properties = new MetadataPropertiesExpression(AttributeProperties)
                        },
                        RelationshipQuery = new RelationshipQueryExpression
                        {
                            Properties = new MetadataPropertiesExpression(RelationshipProperties)
                        }
                    };
                    foreach (var condition in batch) query.Criteria.Conditions.Add(condition);
                    var response = context.Execute(new RetrieveMetadataChangesRequest { Query = query })
                        as RetrieveMetadataChangesResponse;
                    var entities = response?.EntityMetadata;
                    if (entities == null) throw new InvalidOperationException("No entity metadata collection was returned.");
                    if (entities.Count != batch.Count || entities.Any(entity => entity == null ||
                        !entity.MetadataId.HasValue || entity.MetadataId == Guid.Empty ||
                        string.IsNullOrWhiteSpace(entity.LogicalName)) ||
                        entities.GroupBy(entity => entity.MetadataId.Value).Any(group => group.Count() != 1) ||
                        entities.GroupBy(entity => entity.LogicalName, StringComparer.OrdinalIgnoreCase)
                            .Any(group => group.Count() != 1))
                        throw new InvalidOperationException("Parent metadata correlation is missing, duplicate, or incomplete.");
                    foreach (var condition in batch)
                        if (entities.Count(entity => condition.PropertyName == "MetadataId"
                            ? entity.MetadataId == (Guid)condition.Value
                            : StringComparer.OrdinalIgnoreCase.Equals(entity.LogicalName, (string)condition.Value)) != 1)
                            throw new InvalidOperationException("Returned parent metadata does not match the requested entities.");
                    if (entities.Any(entity => entity.Attributes == null ||
                        entity.OneToManyRelationships == null || entity.ManyToOneRelationships == null ||
                        entity.ManyToManyRelationships == null))
                        throw new InvalidOperationException("One or more requested child metadata collections were not supplied.");
                    returned.AddRange(entities);
                }
                catch (OperationCanceledException) { throw; }
                catch (Exception ex) when (ex is FaultException || ex is InvalidOperationException)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    inventory.Fail(audit + "; " + ex.Message);
                    context.MetadataCache.ParentMetadata = inventory;
                    return;
                }
            }
            cancellationToken.ThrowIfCancellationRequested();
            if (returned.GroupBy(item => item.MetadataId.Value).Any(group => group.Count() != 1) ||
                returned.GroupBy(item => item.LogicalName, StringComparer.OrdinalIgnoreCase)
                    .Any(group => group.Count() != 1))
                inventory.Fail("Column / Relationship parent metadata has conflicting parents across batches.");
            else inventory.Load(returned);
            cancellationToken.ThrowIfCancellationRequested();
            if (!inventory.Failed)
                foreach (var entity in returned) context.MetadataCache.Store(entity, true);
            context.MetadataCache.ParentMetadata = inventory;
        }
    }

    /// <summary>Separate Column and Relationship indices prevent cross-family GUID collisions.</summary>
    internal sealed class ParentEntityMetadataInventory
    {
        private readonly Dictionary<Guid, List<ChildMetadataCorrelation>> columns =
            new Dictionary<Guid, List<ChildMetadataCorrelation>>();
        private readonly Dictionary<Guid, List<ChildMetadataCorrelation>> relationships =
            new Dictionary<Guid, List<ChildMetadataCorrelation>>();
        internal string ScopeEvidence;
        internal readonly HashSet<Guid> ParentIds = new HashSet<Guid>();
        internal string FailureEvidence => failure;
        private string failure;
        internal bool Failed => failure != null;
        internal void Fail(string diagnostic) { failure = diagnostic; columns.Clear(); relationships.Clear(); }

        internal void Load(IEnumerable<EntityMetadata> entities)
        {
            foreach (var entity in entities)
            {
                foreach (var attribute in entity.Attributes)
                {
                    if (attribute?.MetadataId == null || attribute.MetadataId == Guid.Empty)
                    { Fail("Column parent metadata contains a child without a usable MetadataId."); return; }
                    Add(columns, attribute.MetadataId.Value, new ChildMetadataCorrelation
                    {
                        Attribute = attribute, ParentLogicalName = entity.LogicalName
                    });
                }
                foreach (var relationship in entity.OneToManyRelationships.Cast<RelationshipMetadataBase>()
                    .Concat(entity.ManyToOneRelationships).Concat(entity.ManyToManyRelationships))
                {
                    if (relationship?.MetadataId == null || relationship.MetadataId == Guid.Empty)
                    { Fail("Relationship parent metadata contains a child without a usable MetadataId."); return; }
                    Add(relationships, relationship.MetadataId.Value, new ChildMetadataCorrelation
                    {
                        Relationship = relationship, ParentLogicalName = entity.LogicalName
                    });
                }
            }
        }

        internal ChildMetadataCorrelation Correlate(string kind, Guid id)
        {
            var result = new ChildMetadataCorrelation();
            if (Failed)
            {
                result.Diagnostic = kind + " metadata inventory is incomplete; see diagnostic evidence.";
                result.Evidence = new[] { failure };
                return result;
            }
            var index = kind == ComponentSemanticKinds.Column ? columns : relationships;
            List<ChildMetadataCorrelation> matches;
            if (!index.TryGetValue(id, out matches))
            {
                result.Evidence = Evidence(id, kind);
                result.Diagnostic = "No " + kind + " metadata matched the component object ID within verified parent Tables.";
                return result;
            }
            // A relationship can appear from both endpoints. Collapse only byte-equivalent
            // SDK definitions, not merely matching GUIDs or schema names. Conflicts remain ambiguous.
            if (kind == ComponentSemanticKinds.Relationship)
                matches = matches.GroupBy(item => RelationshipSignature(item.Relationship), StringComparer.Ordinal)
                    .Select(group => group.First()).ToList();
            if (matches.Count != 1)
            {
                result.Evidence = Evidence(id, kind);
                result.Status = IdentityResolutionStatus.Ambiguous;
                result.Diagnostic = "Multiple conflicting " + kind + " metadata records matched the component object ID.";
                return result;
            }
            result = matches[0];
            result.Evidence = Evidence(id, kind);
            string key;
            if (result.Attribute != null)
            {
                var attribute = result.Attribute;
                if (string.IsNullOrWhiteSpace(attribute.LogicalName) ||
                    (!string.IsNullOrWhiteSpace(attribute.EntityLogicalName) &&
                     !StringComparer.OrdinalIgnoreCase.Equals(attribute.EntityLogicalName, result.ParentLogicalName)))
                {
                    result.Diagnostic = "Column metadata has a missing logical name or conflicting parent.";
                    return result;
                }
                key = result.ParentLogicalName + "." + attribute.LogicalName;
            }
            else key = result.Relationship.SchemaName;
            if (string.IsNullOrWhiteSpace(key))
            {
                result.Diagnostic = "The correlated " + kind + " metadata has no portable name.";
                return result;
            }
            result.PortableKey = key;
            result.Status = IdentityResolutionStatus.Resolved;
            result.Evidence = Evidence(id, kind).Concat(new[] {
                "Correlated parent=" + result.ParentLogicalName + "." });
            return result;
        }

        private IEnumerable<string> Evidence(Guid id, string kind) => new[] {
            kind + " MetadataId=" + id.ToString("D") + ". Local ID is audit evidence only.",
            ScopeEvidence ?? "Metadata supplied by the operation-scoped parent inventory." };

        private static void Add(Dictionary<Guid, List<ChildMetadataCorrelation>> index, Guid id,
            ChildMetadataCorrelation value)
        {
            List<ChildMetadataCorrelation> list;
            if (!index.TryGetValue(id, out list)) index[id] = list = new List<ChildMetadataCorrelation>();
            list.Add(value);
        }

        private static string RelationshipSignature(RelationshipMetadataBase metadata)
        {
            using (var stream = new MemoryStream())
            {
                new DataContractSerializer(metadata.GetType()).WriteObject(stream, metadata);
                return Convert.ToBase64String(stream.ToArray());
            }
        }
    }

    internal sealed class ChildMetadataCorrelation
    {
        internal IdentityResolutionStatus Status = IdentityResolutionStatus.Unresolved;
        internal string Diagnostic = string.Empty;
        internal string PortableKey;
        internal string ParentLogicalName;
        internal AttributeMetadata Attribute;
        internal RelationshipMetadataBase Relationship;
        internal IEnumerable<string> Evidence = new string[0];
    }
}
