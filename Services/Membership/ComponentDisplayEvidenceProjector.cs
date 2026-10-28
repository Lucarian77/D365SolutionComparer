using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using D365SolutionComparer.Models.Membership;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Metadata;

namespace D365SolutionComparer.Services.Membership
{
    /// <summary>
    /// Read-only, one-way projection after authoritative comparison. No service dependency,
    /// diagnostic-string parsing, portable-key parsing, or retrieval pipeline. Future consumers
    /// can use Project directly; coordinated operations can supply their existing caches.
    /// Does not attach evidence to or replace the protected presentation/identity models.
    /// </summary>
    public static class ComponentDisplayEvidenceProjector
    {
        public static IReadOnlyList<ComponentDisplayEvidence> Project(
            MembershipComparisonPresentation presentation, CancellationToken cancellationToken = default(CancellationToken)) =>
            Project(presentation, null, null, cancellationToken);

        internal static IReadOnlyList<ComponentDisplayEvidence> Project(
            MembershipComparisonPresentation presentation, DataverseComponentMetadataCache sourceCache,
            DataverseComponentMetadataCache targetCache, CancellationToken cancellationToken = default(CancellationToken))
        {
            if (presentation == null) throw new ArgumentNullException(nameof(presentation));
            // Cancellation discards this optional projection, never cancels an authoritative comparison.
            // No partial display inventory is returned.
            if (cancellationToken.IsCancellationRequested) return new ComponentDisplayEvidence[0];
            var source = View(sourceCache, presentation.Source.Snapshot);
            var target = View(targetCache, presentation.Target.Snapshot);
            var results = new List<ComponentDisplayEvidence>(presentation.Rows.Count);
            foreach (var row in presentation.Rows)
            {
                if (cancellationToken.IsCancellationRequested) return new ComponentDisplayEvidence[0];
                results.Add(new ComponentDisplayEvidence(results.Count,
                    source.Project(row.SourceIdentity, DisplayEnvironmentSide.Source, presentation.Source.State),
                    target.Project(row.TargetIdentity, DisplayEnvironmentSide.Target, presentation.Target.State)));
            }
            return results.AsReadOnly();
        }

        private static CacheView View(DataverseComponentMetadataCache cache, MembershipSnapshot owner)
        {
            try { return new CacheView(cache, owner); }
            catch (Exception) { return new CacheView(null, owner, true); }
        }

        private sealed class CacheView
        {
            private readonly DataverseComponentMetadataCache cache;
            private readonly MembershipSnapshot owner;
            private readonly bool failed;
            private readonly Dictionary<Guid, List<OptionSetMetadataBase>> choices =
                new Dictionary<Guid, List<OptionSetMetadataBase>>();
            private readonly Dictionary<Guid, List<DisplayColumn>> columns =
                new Dictionary<Guid, List<DisplayColumn>>();
            private readonly Dictionary<Guid, List<RelationshipMetadataBase>> relationships =
                new Dictionary<Guid, List<RelationshipMetadataBase>>();
            private readonly Dictionary<Guid, List<Type92StepEvidence>> steps =
                new Dictionary<Guid, List<Type92StepEvidence>>();

            internal CacheView(DataverseComponentMetadataCache cache, MembershipSnapshot owner, bool failed = false)
            {
                this.cache = cache;
                this.owner = owner;
                this.failed = failed;
                foreach (var item in cache?.OptionSetCatalog ?? new OptionSetMetadataBase[0])
                {
                    if (item?.MetadataId == null) continue;
                    List<OptionSetMetadataBase> group;
                    if (!choices.TryGetValue(item.MetadataId.Value, out group))
                        choices[item.MetadataId.Value] = group = new List<OptionSetMetadataBase>();
                    group.Add(item);
                }
                // Index existing published child metadata once. Never call Correlate here:
                // that resolver method constructs keys and updates its internal correlation result.
                if (cache?.ParentMetadata != null && !cache.ParentMetadata.Failed)
                    foreach (var entity in cache.RetainedEntities)
                    {
                        foreach (var attribute in entity.Attributes ?? new AttributeMetadata[0])
                            if (attribute?.MetadataId != null) Add(columns, attribute.MetadataId.Value,
                                new DisplayColumn { Metadata = attribute, Parent = entity.LogicalName });
                        foreach (var relationship in (entity.OneToManyRelationships ?? new OneToManyRelationshipMetadata[0])
                            .Cast<RelationshipMetadataBase>().Concat(entity.ManyToOneRelationships ?? new OneToManyRelationshipMetadata[0])
                            .Concat(entity.ManyToManyRelationships ?? new ManyToManyRelationshipMetadata[0]))
                            if (relationship?.MetadataId != null) Add(relationships, relationship.MetadataId.Value, relationship);
                    }
                foreach (var step in cache?.Type92Evidence?.Steps ?? new Type92StepEvidence[0])
                    if (step != null) Add(steps, step.ObjectId, step);
            }

            internal ComponentSideDisplayEvidence Project(ComponentIdentity identity,
                DisplayEnvironmentSide side, MembershipSnapshotState state)
            {
                if (identity == null)
                {
                    // No local row is not a new absence inference. Unavailable snapshots remain unavailable.
                    var empty = Missing(state == MembershipSnapshotState.Unavailable
                        ? DisplayAvailability.Unavailable : DisplayAvailability.NotApplicable);
                    return new ComponentSideDisplayEvidence(side, null, null, null, empty, empty, empty,
                        DisplayBackingCorrelationStatus.NotAttempted, empty, ComponentObjectMetadataEvidence.NotRetrieved(side));
                }
                var category = Category(identity);
                try { return Read(identity, side, category); }
                catch (Exception)
                {
                    // Optional display failures cannot invalidate already computed membership or definitions.
                    // No SDK fault text/payload is retained here.
                    return Evidence(identity, side, category, Missing(DisplayAvailability.RetrievalFailed),
                        Missing(DisplayAvailability.RetrievalFailed), DisplayBackingCorrelationStatus.RetrievalFailed);
                }
            }

            private ComponentSideDisplayEvidence Read(ComponentIdentity identity,
                DisplayEnvironmentSide side, ComponentDisplayValue category)
            {
                var unknown = Missing(DisplayAvailability.NotRetrieved);
                if (failed)
                    return Evidence(identity, side, category, Missing(DisplayAvailability.RetrievalFailed),
                        Missing(DisplayAvailability.RetrievalFailed), DisplayBackingCorrelationStatus.RetrievalFailed);
                var id = identity.Record.ObjectId;
                if (!id.HasValue || id == Guid.Empty)
                    return Evidence(identity, side, category, unknown, unknown, DisplayBackingCorrelationStatus.NotAttempted);
                if (cache == null)
                    return Evidence(identity, side, category, unknown, unknown, DisplayBackingCorrelationStatus.Unavailable);

                if (identity.Record.ComponentType == 10139)
                {
                    MetadataForArchivalInformation info;
                    if (cache.MetadataForArchival != null && cache.MetadataForArchival.TryInformation(owner, identity, side, out info))
                        return new ComponentSideDisplayEvidence(side, identity.Record.SolutionComponentId, id, 10139, category,
                            info.Field("Name"), info.Field("Name"), info.Field("Name").Availability == DisplayAvailability.Available
                                ? DisplayBackingCorrelationStatus.Unique : DisplayBackingCorrelationStatus.Unavailable,
                            info.Field("Related Table"), archivalInformation: info);
                    return Evidence(identity, side, category, unknown, unknown, DisplayBackingCorrelationStatus.Unavailable);
                }

                MakerObjectDisplayEvidence maker;
                if (cache.TryGetMakerObjectDisplay(owner, identity, out maker))
                    return Evidence(identity, side, category, maker.FriendlyName, maker.InternalName, maker.Correlation, maker: maker);

                if (identity.Record.ComponentType == 300)
                {
                    DataverseComponentMetadataCache.NormalDisplayNameObservation canvas;
                    return cache.TryGetNormalDisplayName(owner, identity, out canvas)
                        ? Evidence(identity, side, category, canvas.FriendlyName, canvas.InternalName, canvas.Correlation)
                        : Evidence(identity, side, category, unknown, unknown, DisplayBackingCorrelationStatus.Unavailable);
                }

                if (identity.Record.ComponentType == 511 || identity.Record.ComponentType == 59)
                {
                    DataverseComponentMetadataCache.NormalDisplayNameObservation retained;
                    if (cache.TryGetNormalDisplayName(owner, identity, out retained))
                        return Evidence(identity, side, category, retained.FriendlyName, unknown, retained.Correlation, retained.ParentContext);
                }

                if (identity.Record.ComponentType == 31 || identity.Record.ComponentType == 36)
                {
                    // The coordinator verifies invocation/org/solution ownership before supplying
                    // the side's cache; the observation also requires its exact final snapshot/record.
                    DataverseComponentMetadataCache.NormalDisplayNameObservation observation;
                    return cache.TryGetNormalDisplayName(owner, identity, out observation)
                        ? Evidence(identity, side, category, observation.FriendlyName, unknown, observation.Correlation)
                        : Evidence(identity, side, category, unknown, unknown, DisplayBackingCorrelationStatus.Unavailable);
                }

                EntityMetadata entity;
                if (identity.Record.ComponentType == 1 && cache.TryGetEntity(id.Value, out entity))
                    return Evidence(identity, side, category, Label(entity.DisplayName, "EntityMetadata.DisplayName"),
                        Observed(entity.LogicalName, "EntityMetadata.LogicalName"), DisplayBackingCorrelationStatus.VerifiedMetadata);

                AttributeMetadata directColumn;
                if (identity.Record.ComponentType == 2 && cache.TryGetAttribute(id.Value, out directColumn))
                    return Column(identity, side, category, directColumn);
                List<DisplayColumn> columnRows;
                if (identity.Record.ComponentType == 2 && columns.TryGetValue(id.Value, out columnRows))
                {
                    if (columnRows.Count != 1) return Ambiguous(identity, side, category);
                    return Column(identity, side, category, columnRows[0].Metadata, columnRows[0].Parent);
                }
                RelationshipMetadataBase directRelationship;
                if (identity.Record.ComponentType == 10 && cache.TryGetRelationship(id.Value, out directRelationship))
                    return Evidence(identity, side, category, unknown,
                        Observed(directRelationship.SchemaName, "RelationshipMetadata.SchemaName"), DisplayBackingCorrelationStatus.VerifiedMetadata);
                List<RelationshipMetadataBase> relationshipRows;
                if (identity.Record.ComponentType == 10 && relationships.TryGetValue(id.Value, out relationshipRows))
                {
                    // Endpoint duplicates are already handled by membership. Display does not re-resolve them.
                    if (relationshipRows.Count != 1) return Ambiguous(identity, side, category);
                    return Evidence(identity, side, category, unknown,
                        Observed(relationshipRows[0].SchemaName, "RelationshipMetadata.SchemaName"), DisplayBackingCorrelationStatus.VerifiedMetadata);
                }
                List<Type92StepEvidence> stepRows;
                if (identity.Record.ComponentType == 92 && steps.TryGetValue(id.Value, out stepRows))
                {
                    if (stepRows.Count != 1) return Ambiguous(identity, side, category);
                    var step = stepRows[0];
                    if (step.Correlation == "Unique")
                        return Evidence(identity, side, category, Observed(step.Name, "CompletedType92Evidence.name"),
                            unknown, DisplayBackingCorrelationStatus.Unique);
                    if (step.Correlation == "Missing")
                        return Evidence(identity, side, category, Missing(DisplayAvailability.Unavailable), unknown, DisplayBackingCorrelationStatus.Missing);
                    if (step.Correlation == "DuplicateReturnedRow" || step.Correlation == "ConflictingPrimaryKey")
                        return Ambiguous(identity, side, category);
                    // IncompleteBatch is incomplete evidence, not proof that a request faulted.
                    return Evidence(identity, side, category, Missing(DisplayAvailability.Unavailable), unknown,
                        DisplayBackingCorrelationStatus.Unavailable);
                }

                List<OptionSetMetadataBase> options;
                if (identity.Record.ComponentType == 9 && choices.TryGetValue(id.Value, out options))
                {
                    if (options.Count != 1)
                        return Evidence(identity, side, category, Missing(DisplayAvailability.Ambiguous),
                            Missing(DisplayAvailability.Ambiguous), DisplayBackingCorrelationStatus.Ambiguous);
                    return Evidence(identity, side, category, Label(options[0].DisplayName, "OptionSetMetadata.DisplayName"),
                        Observed(options[0].Name, "OptionSetMetadata.Name"), DisplayBackingCorrelationStatus.VerifiedMetadata);
                }

                Entity workflow;
                if (identity.Record.ComponentType == 29 && cache.WorkflowDefinitions.TryGetValue(id.Value, out workflow))
                    return Evidence(identity, side, category, Field(workflow, "name", true), Field(workflow, "uniquename", true),
                        workflow.Id == id.Value ? DisplayBackingCorrelationStatus.Unique : DisplayBackingCorrelationStatus.VerifiedDefinition);

                string table, friendly, internalName;
                if (Fields(identity, out table, out friendly, out internalName))
                {
                    Entity row;
                    if (cache.TryGetEntityRow(table, id.Value, out row))
                    {
                        if (row.Id != id.Value || !StringComparer.OrdinalIgnoreCase.Equals(row.LogicalName, table))
                            return Evidence(identity, side, category, Missing(DisplayAvailability.Ambiguous),
                                Missing(DisplayAvailability.Ambiguous), DisplayBackingCorrelationStatus.Ambiguous);
                        return Evidence(identity, side, category, Field(row, friendly, friendly != null),
                            Field(row, internalName, internalName != null), DisplayBackingCorrelationStatus.Unique);
                    }
                }
                // Cache absence does not prove a missing backing row or an unattempted request.
                // Legacy string diagnostics are not a structured display contract.
                return Evidence(identity, side, category, unknown, unknown, DisplayBackingCorrelationStatus.Unavailable);
            }

            private static void Add<T>(Dictionary<Guid, List<T>> index, Guid id, T item)
            {
                List<T> group;
                if (!index.TryGetValue(id, out group)) index[id] = group = new List<T>();
                group.Add(item);
            }

            private sealed class DisplayColumn
            {
                internal AttributeMetadata Metadata;
                internal string Parent;
            }

            private static ComponentSideDisplayEvidence Column(ComponentIdentity identity, DisplayEnvironmentSide side,
                ComponentDisplayValue category, AttributeMetadata column, string parent = null)
            {
                var scope = parent ?? column.EntityLogicalName;
                var conflict = parent != null && !string.IsNullOrWhiteSpace(column.EntityLogicalName) &&
                    !StringComparer.OrdinalIgnoreCase.Equals(parent, column.EntityLogicalName);
                return Evidence(identity, side, category, Label(column.DisplayName, "AttributeMetadata.DisplayName"),
                    Observed(column.LogicalName, "AttributeMetadata.LogicalName"), DisplayBackingCorrelationStatus.VerifiedMetadata,
                    conflict ? Missing(DisplayAvailability.Unavailable) : string.IsNullOrWhiteSpace(scope)
                        ? Missing(DisplayAvailability.NotRetrieved) : Observed(scope, parent == null
                            ? "AttributeMetadata.EntityLogicalName" : "VerifiedParentMetadata.LogicalName"));
            }

            private static ComponentSideDisplayEvidence Ambiguous(ComponentIdentity identity, DisplayEnvironmentSide side,
                ComponentDisplayValue category) => Evidence(identity, side, category, Missing(DisplayAvailability.Ambiguous),
                    Missing(DisplayAvailability.Ambiguous), DisplayBackingCorrelationStatus.Ambiguous);
        }

        private static ComponentDisplayValue Field(Entity row, string field, bool requested)
        {
            if (!requested) return Missing(DisplayAvailability.NotRetrieved);
            object value;
            if (!row.Attributes.TryGetValue(field, out value))
                return new ComponentDisplayValue(DisplayAvailability.AttributeOmitted, displayEvidenceSource: row.LogicalName + "." + field);
            if (value == null || value is string) return Observed((string)value, row.LogicalName + "." + field);
            return Missing(DisplayAvailability.Unavailable);
        }

        private static ComponentDisplayValue Label(Label label, string provenance)
        {
            // Missing labels cannot reveal whether the metadata property was requested.
            var text = label?.UserLocalizedLabel?.Label;
            return text == null ? Missing(DisplayAvailability.NotRetrieved) : Observed(text, provenance);
        }

        private static ComponentDisplayValue Observed(string value, string provenance) =>
            new ComponentDisplayValue(DisplayAvailability.Available, value, provenance);
        private static ComponentDisplayValue Missing(DisplayAvailability state) => new ComponentDisplayValue(state);

        private static ComponentSideDisplayEvidence Evidence(ComponentIdentity identity, DisplayEnvironmentSide side,
            ComponentDisplayValue category, ComponentDisplayValue friendly, ComponentDisplayValue internalName,
            DisplayBackingCorrelationStatus correlation, ComponentDisplayValue parent = null, MakerObjectDisplayEvidence maker = null) =>
            new ComponentSideDisplayEvidence(side, identity.Record.SolutionComponentId, identity.Record.ObjectId,
                identity.Record.ComponentType, category, friendly, internalName, correlation,
                parent ?? Missing(DisplayAvailability.NotRetrieved), ComponentObjectMetadataEvidence.NotRetrieved(side), maker);

        private static bool Fields(ComponentIdentity identity, out string table, out string friendly, out string internalName)
        {
            table = friendly = internalName = null;
            switch (identity.SemanticKind)
            {
                case ComponentSemanticKinds.WebResource: table = "webresource"; friendly = "displayname"; internalName = "name"; break;
                case ComponentSemanticKinds.EnvironmentVariableDefinition: table = "environmentvariabledefinition"; friendly = "displayname"; internalName = "schemaname"; break;
                case ComponentSemanticKinds.ConnectionReference: table = "connectionreference"; friendly = "connectionreferencedisplayname"; internalName = "connectionreferencelogicalname"; break;
                case ComponentSemanticKinds.AppModule: table = "appmodule"; friendly = "name"; internalName = "uniquename"; break;
                case ComponentSemanticKinds.SystemForm: table = "systemform"; friendly = "name"; internalName = "uniquename"; break;
                case ComponentSemanticKinds.SiteMap: table = "sitemap"; friendly = "sitemapname"; internalName = "sitemapnameunique"; break;
                case ComponentSemanticKinds.SavedQuery: table = "savedquery"; friendly = "name"; break;
                case ComponentSemanticKinds.SavedQueryVisualization: table = "savedqueryvisualization"; friendly = "name"; break;
                case ComponentSemanticKinds.PluginAssembly: table = "pluginassembly"; internalName = "name"; break;
            }
            return table != null;
        }

        private static ComponentDisplayValue Category(ComponentIdentity identity)
        {
            string label;
            switch (identity.Record.ComponentType)
            {
                case 1: label = "Table"; break;
                case 2: label = "Column"; break;
                case 9: label = "Global Choice (subset)"; break;
                case 10: label = "Relationship"; break;
                case 26: label = "Saved Query"; break;
                case 29: label = "Process / Workflow"; break;
                case 31: label = "Report (subset)"; break;
                case 36: label = "Email Template"; break;
                case 59: label = "Chart / Saved Query Visualization"; break;
                case 60: label = "Form"; break;
                case 61: label = "Web Resource"; break;
                case 62: label = "Site Map"; break;
                case 80: label = "Model-Driven App"; break;
                case 10139: label = "MetadataForArchival"; break;
                case 91: label = "Plug-in Assembly"; break;
                case 92: label = "SDK Message Processing Step"; break;
                case 300: label = "Canvas App family"; break;
                case 380: label = "Environment Variable Definition"; break;
                default:
                    if (identity.SemanticKind == ComponentSemanticKinds.ConnectionReference)
                        return Observed("Connection Reference", "VerifiedResolver.SemanticKind");
                    if (identity.SemanticKind == ComponentSemanticKinds.TeamTemplate)
                        return Observed("Team Template", "VerifiedResolver.SemanticKind");
                    if (identity.RegisteredDefinition != null)
                        return Observed(identity.RegisteredDefinition.Name, "SolutionComponentDefinition.Name");
                    return Observed("Unknown", "ComponentCatalog.Unclassified");
            }
            return Observed(label, "PublishedComponentCatalog.ComponentType");
        }
    }
}
