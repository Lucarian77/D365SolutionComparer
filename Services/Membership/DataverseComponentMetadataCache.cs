using System;
using System.Collections.Generic;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Metadata;

namespace D365SolutionComparer.Services.Membership
{
    /// <summary>Operation-scoped published metadata shared by identity and definition reads.</summary>
    internal sealed class DataverseComponentMetadataCache
    {
        private readonly Dictionary<Guid, EntityMetadata> entities = new Dictionary<Guid, EntityMetadata>();
        private readonly Dictionary<Guid, AttributeMetadata> attributes = new Dictionary<Guid, AttributeMetadata>();
        private readonly Dictionary<Guid, RelationshipMetadataBase> relationships =
            new Dictionary<Guid, RelationshipMetadataBase>();
        private readonly HashSet<Guid> completeEntityDefinitions = new HashSet<Guid>();
        private readonly Dictionary<string, Entity> entityRows =
            new Dictionary<string, Entity>(StringComparer.OrdinalIgnoreCase);

        internal ParentEntityMetadataInventory ParentMetadata { get; set; }

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
