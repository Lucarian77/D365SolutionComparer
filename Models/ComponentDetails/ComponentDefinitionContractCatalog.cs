using System;
using System.Collections.Generic;
using System.Linq;
using D365SolutionComparer.Models.Membership;

namespace D365SolutionComparer.Models.ComponentDetails
{
    /// <summary>
    /// Phase 2G.1 equality contract. Every listed property must be represented, including a
    /// represented null value, before a definition is complete enough to compare.
    /// Identity fields and audit-only local IDs are intentionally absent.
    /// </summary>
    internal static class ComponentDefinitionContractCatalog
    {
        private static readonly IDictionary<string, ComponentDefinitionContract> Contracts =
            new Dictionary<string, ComponentDefinitionContract>(StringComparer.OrdinalIgnoreCase)
            {
                [ComponentSemanticKinds.Table] = Contract(ComponentSemanticKinds.Table,
                    "SchemaName", "OwnershipType", "IsActivity", "IsIntersect", "HasActivities",
                    "HasNotes", "IsAuditEnabled", "IsValidForAdvancedFind",
                    "IsDuplicateDetectionEnabled", "IsCustomizable"),
                [ComponentSemanticKinds.Column] = Contract(ComponentSemanticKinds.Column,
                    "SchemaName", "AttributeType", "AttributeTypeName", "RequiredLevel",
                    "IsAuditEnabled", "IsSecured", "IsValidForAdvancedFind", "IsCustomizable",
                    "MaxLength", "MinValue", "MaxValue", "IntegerFormat", "Precision",
                    "PrecisionSource", "DateTimeFormat", "DateTimeBehavior", "Targets", "Options"),
                [ComponentSemanticKinds.Relationship] = Contract(ComponentSemanticKinds.Relationship,
                    "RelationshipType", "IsCustomizable", "ReferencedEntity", "ReferencedAttribute",
                    "ReferencingEntity", "ReferencingAttribute", "CascadeAssign", "CascadeDelete",
                    "CascadeMerge", "CascadeReparent", "CascadeShare", "CascadeUnshare",
                    "Entity1LogicalName", "Entity2LogicalName", "IntersectEntityName"),
                [ComponentSemanticKinds.WebResource] = Contract(ComponentSemanticKinds.WebResource,
                    "webresourcetype", "content", "displayname", "description"),
                [ComponentSemanticKinds.GlobalChoice] = Contract(ComponentSemanticKinds.GlobalChoice,
                    "OptionSetType", "IsCustomOptionSet", "Options"),
                [ComponentSemanticKinds.EnvironmentVariableDefinition] = Contract(
                    ComponentSemanticKinds.EnvironmentVariableDefinition, "type", "defaultvalue",
                    "valueschema", "isrequired", "secretstore", "displayname", "description"),
                [ComponentSemanticKinds.ConnectionReference] = Contract(
                    ComponentSemanticKinds.ConnectionReference, "connectionreferencedisplayname",
                    "connectorid", "description"),
                [ComponentSemanticKinds.AppModule] = Contract(ComponentSemanticKinds.AppModule,
                    "name", "description", "clienttype", "formfactor", "navigationtype")
            };

        public static ComponentDefinitionContract For(string semanticKind)
        {
            ComponentDefinitionContract contract;
            return semanticKind != null && Contracts.TryGetValue(semanticKind, out contract)
                ? contract : null;
        }

        private static ComponentDefinitionContract Contract(string kind, params string[] properties) =>
            new ComponentDefinitionContract(kind, properties, AuditOnlyProperties(kind),
                StringComparer.OrdinalIgnoreCase);

        private static IEnumerable<string> AuditOnlyProperties(string kind)
        {
            if (string.Equals(kind, ComponentSemanticKinds.Table, StringComparison.OrdinalIgnoreCase))
                return new[] { "MetadataId", "LogicalName", "IsManaged" };
            if (string.Equals(kind, ComponentSemanticKinds.Column, StringComparison.OrdinalIgnoreCase))
                return new[] { "MetadataId", "EntityLogicalName", "LogicalName", "IsManaged" };
            if (string.Equals(kind, ComponentSemanticKinds.Relationship, StringComparison.OrdinalIgnoreCase))
                return new[] { "MetadataId", "SchemaName", "IsManaged" };
            if (string.Equals(kind, ComponentSemanticKinds.WebResource, StringComparison.OrdinalIgnoreCase))
                return new[] { "webresourceid", "webresourceidunique", "name", "componentstate", "ismanaged" };
            if (string.Equals(kind, ComponentSemanticKinds.GlobalChoice, StringComparison.OrdinalIgnoreCase))
                return new[] { "MetadataId", "Name", "IsGlobal", "IsManaged" };
            if (string.Equals(kind, ComponentSemanticKinds.EnvironmentVariableDefinition,
                StringComparison.OrdinalIgnoreCase))
                return new[] { "environmentvariabledefinitionid", "schemaname", "componentstate", "ismanaged" };
            if (string.Equals(kind, ComponentSemanticKinds.ConnectionReference,
                StringComparison.OrdinalIgnoreCase))
                return new[] { "connectionreferenceid", "connectionreferencelogicalname", "componentstate", "ismanaged" };
            if (string.Equals(kind, ComponentSemanticKinds.AppModule, StringComparison.OrdinalIgnoreCase))
                return new[] { "appmoduleid", "appmoduleidunique", "uniquename", "componentstate", "ismanaged" };
            return Enumerable.Empty<string>();
        }
    }

    internal sealed class ComponentDefinitionContract
    {
        public ComponentDefinitionContract(string semanticKind, IEnumerable<string> comparableProperties,
            IEnumerable<string> auditOnlyProperties, StringComparer identityComparer)
        {
            SemanticKind = semanticKind;
            ComparableProperties = comparableProperties.ToList().AsReadOnly();
            AuditOnlyProperties = auditOnlyProperties.ToList().AsReadOnly();
            IdentityComparer = identityComparer;
        }

        public string SemanticKind { get; }
        public IReadOnlyList<string> ComparableProperties { get; }

        /// <summary>
        /// Correlation, identity, management, and local-state fields that must never participate
        /// in Phase 2G.1 definition equality.
        /// </summary>
        public IReadOnlyList<string> AuditOnlyProperties { get; }

        /// <summary>The existing membership comparer establishes these keys case-insensitively.</summary>
        public StringComparer IdentityComparer { get; }
    }
}
