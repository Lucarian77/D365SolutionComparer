using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.ServiceModel;
using System.Threading;
using D365SolutionComparer.Infrastructure;
using D365SolutionComparer.Models.ComponentDetails;
using D365SolutionComparer.Models.Membership;
using D365SolutionComparer.Services.Contracts;
using D365SolutionComparer.Services.Membership;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Metadata.Query;
using Microsoft.Xrm.Sdk.Query;

namespace D365SolutionComparer.Services.ComponentDetails
{
    /// <summary>
    /// Reads published component configuration only. It reuses established portable identities,
    /// batches entity-backed reads, and never issues a Dataverse write request.
    /// </summary>
    internal sealed class DataverseComponentDefinitionReader : IComponentDefinitionReader
    {
        private const int BatchSize = 200;
        private static readonly HashSet<string> SupportedKinds = new HashSet<string>(
            new[]
            {
                ComponentSemanticKinds.Table,
                ComponentSemanticKinds.Column,
                ComponentSemanticKinds.Relationship,
                ComponentSemanticKinds.WebResource,
                ComponentSemanticKinds.GlobalChoice,
                ComponentSemanticKinds.EnvironmentVariableDefinition,
                ComponentSemanticKinds.ConnectionReference,
                ComponentSemanticKinds.AppModule
            }, StringComparer.OrdinalIgnoreCase);

        public ComponentDefinitionSnapshot Read(IOrganizationService service,
            MembershipSnapshot membership, CancellationToken cancellationToken,
            DataverseRequestCounter requestCounter = null)
        {
            if (membership == null) throw new ArgumentNullException(nameof(membership));
            cancellationToken.ThrowIfCancellationRequested();
            if (membership.State != MembershipSnapshotState.Complete)
                return new ComponentDefinitionSnapshot(membership,
                    Enumerable.Empty<ComponentDefinition>(), membership.Diagnostic);
            var context = new DataverseReadContext(service, membership.Environment,
                cancellationToken, requestCounter);
            return Read(context, membership, cancellationToken);
        }

        internal ComponentDefinitionSnapshot Read(DataverseReadContext context,
            MembershipSnapshot membership, CancellationToken cancellationToken)
        {
            if (context == null) throw new ArgumentNullException(nameof(context));
            if (membership == null) throw new ArgumentNullException(nameof(membership));
            cancellationToken.ThrowIfCancellationRequested();
            if (membership.State != MembershipSnapshotState.Complete)
                return new ComponentDefinitionSnapshot(membership,
                    Enumerable.Empty<ComponentDefinition>(), membership.Diagnostic);

            var results = new Dictionary<Guid, ComponentDefinition>();
            foreach (var identity in membership.Components)
            {
                if (identity.Status != IdentityResolutionStatus.Resolved)
                {
                    Set(results, FromIdentityStatus(identity));
                    continue;
                }
                if (!SupportedKinds.Contains(identity.SemanticKind))
                {
                    Set(results, new ComponentDefinition(identity,
                        ComponentDefinitionReadStatus.Unsupported,
                        diagnostic: "Definition comparison does not support this component kind."));
                    continue;
                }
                if (!identity.Record.ObjectId.HasValue || identity.Record.ObjectId.Value == Guid.Empty)
                    Set(results, new ComponentDefinition(identity,
                        ComponentDefinitionReadStatus.Unresolved,
                        diagnostic: "The resolved membership record has no usable object ID."));
            }

            ReadTables(context, Pending(membership, results, ComponentSemanticKinds.Table),
                results, cancellationToken);
            ReadAttributes(context, Pending(membership, results, ComponentSemanticKinds.Column),
                results, cancellationToken);
            ReadRelationships(context, Pending(membership, results, ComponentSemanticKinds.Relationship),
                results, cancellationToken);
            ReadGlobalChoices(context, Pending(membership, results, ComponentSemanticKinds.GlobalChoice),
                results, cancellationToken);
            ReadEntityBacked(context, Pending(membership, results, ComponentSemanticKinds.WebResource),
                EntityDefinitionConfiguration.WebResource, results, cancellationToken);
            ReadEntityBacked(context, Pending(membership, results,
                ComponentSemanticKinds.EnvironmentVariableDefinition),
                EntityDefinitionConfiguration.EnvironmentVariableDefinition, results, cancellationToken);
            ReadEntityBacked(context, Pending(membership, results,
                ComponentSemanticKinds.ConnectionReference),
                EntityDefinitionConfiguration.ConnectionReference, results, cancellationToken);
            ReadEntityBacked(context, Pending(membership, results, ComponentSemanticKinds.AppModule),
                EntityDefinitionConfiguration.AppModule, results, cancellationToken);

            cancellationToken.ThrowIfCancellationRequested();
            var ordered = membership.Components.Select(identity =>
            {
                ComponentDefinition definition;
                return results.TryGetValue(identity.Record.SolutionComponentId, out definition)
                    ? definition
                    : new ComponentDefinition(identity, ComponentDefinitionReadStatus.Unresolved,
                        diagnostic: "Definition retrieval produced no result for this membership record.");
            }).ToList();
            return new ComponentDefinitionSnapshot(membership, ordered);
        }

        private static void ReadTables(DataverseReadContext context,
            IReadOnlyList<ComponentIdentity> identities,
            IDictionary<Guid, ComponentDefinition> results, CancellationToken cancellationToken)
        {
            var representatives = DistinctRepresentatives(identities);
            var remaining = new List<ComponentIdentity>();
            foreach (var identity in representatives)
            {
                EntityMetadata cached;
                if (context.MetadataCache.TryGetEntity(identity.Record.ObjectId.Value, out cached) &&
                    context.MetadataCache.HasCompleteEntityDefinition(identity.Record.ObjectId.Value))
                    Set(results, Available(identity, TableProperties(cached)));
                else remaining.Add(identity);
            }
            foreach (var batch in Batch(remaining))
            {
                cancellationToken.ThrowIfCancellationRequested();
                try
                {
                    var query = new EntityQueryExpression
                    {
                        Properties = new MetadataPropertiesExpression("MetadataId", "LogicalName",
                            "SchemaName", "OwnershipType", "IsActivity", "IsIntersect",
                            "HasActivities", "HasNotes", "IsAuditEnabled",
                            "IsValidForAdvancedFind", "IsDuplicateDetectionEnabled", "IsCustomizable"),
                        Criteria = new MetadataFilterExpression(LogicalOperator.Or)
                    };
                    foreach (var identity in batch)
                        query.Criteria.Conditions.Add(new MetadataConditionExpression("MetadataId",
                            MetadataConditionOperator.Equals, identity.Record.ObjectId.Value));
                    var response = context.Execute(new RetrieveMetadataChangesRequest { Query = query })
                        as RetrieveMetadataChangesResponse;
                    var metadata = response?.EntityMetadata;
                    if (metadata == null)
                    {
                        SetUnresolved(batch, results, "Table metadata retrieval returned no metadata collection.");
                        continue;
                    }
                    foreach (var item in metadata) context.MetadataCache.Store(item, true);
                    var requested = new HashSet<Guid>(batch.Select(item => item.Record.ObjectId.Value));
                    if (metadata.Any(item => !item.MetadataId.HasValue ||
                        !requested.Contains(item.MetadataId.Value)))
                    {
                        SetUnresolved(batch, results,
                            "Table metadata retrieval returned conflicting or incomplete identifiers.");
                        continue;
                    }
                    var indexed = metadata.GroupBy(item => item.MetadataId.Value)
                        .ToDictionary(group => group.Key, group => group.ToList());
                    foreach (var identity in batch)
                    {
                        List<EntityMetadata> matches;
                        if (!indexed.TryGetValue(identity.Record.ObjectId.Value, out matches))
                            Set(results, Unresolved(identity, "No table metadata matched the component object ID."));
                        else if (matches.Count != 1)
                            Set(results, Ambiguous(identity, "Multiple table metadata records matched the component object ID."));
                        else Set(results, Available(identity, TableProperties(matches[0])));
                    }
                }
                catch (OperationCanceledException) { throw; }
                catch (FaultException ex)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    SetUnresolved(batch, results, "Table metadata retrieval failed: " + ex.Message);
                }
            }
            CopyRepeatedObjectResults(identities, representatives, results);
        }

        private static void ReadAttributes(DataverseReadContext context,
            IReadOnlyList<ComponentIdentity> identities,
            IDictionary<Guid, ComponentDefinition> results, CancellationToken cancellationToken)
        {
            var representatives = DistinctRepresentatives(identities);
            var remaining = new List<ComponentIdentity>();
            foreach (var identity in representatives)
            {
                AttributeMetadata cached;
                if (context.MetadataCache.TryGetAttribute(identity.Record.ObjectId.Value, out cached))
                    Set(results, Available(identity, AttributeProperties(cached)));
                else remaining.Add(identity);
            }
            foreach (var batch in Batch(remaining))
            {
                cancellationToken.ThrowIfCancellationRequested();
                try
                {
                    var attributeQuery = new AttributeQueryExpression
                    {
                        Properties = MetadataProperties(ComponentSemanticKinds.Column,
                            "MetadataId", "LogicalName"),
                        Criteria = MetadataIdFilter(batch)
                    };
                    var query = new EntityQueryExpression
                    {
                        Properties = new MetadataPropertiesExpression("MetadataId", "LogicalName"),
                        AttributeQuery = attributeQuery
                    };
                    var response = context.Execute(new RetrieveMetadataChangesRequest { Query = query })
                        as RetrieveMetadataChangesResponse;
                    var entities = response?.EntityMetadata;
                    if (entities == null)
                    {
                        SetUnresolved(batch, results,
                            "Column metadata retrieval returned no metadata collection.");
                        continue;
                    }
                    var metadata = entities.Where(item => item?.Attributes != null)
                        .SelectMany(item => item.Attributes).Where(item => item != null).ToList();
                    foreach (var item in metadata) context.MetadataCache.Store(item);
                    CorrelateMetadata(batch, metadata, item => item.MetadataId,
                        item => AttributeProperties(item), "column", results);
                }
                catch (OperationCanceledException) { throw; }
                catch (FaultException ex)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    SetUnresolved(batch, results, "Column metadata retrieval failed: " + ex.Message);
                }
            }
            CopyRepeatedObjectResults(identities, representatives, results);
        }

        private static void ReadRelationships(DataverseReadContext context,
            IReadOnlyList<ComponentIdentity> identities,
            IDictionary<Guid, ComponentDefinition> results, CancellationToken cancellationToken)
        {
            var representatives = DistinctRepresentatives(identities);
            var remaining = new List<ComponentIdentity>();
            foreach (var identity in representatives)
            {
                RelationshipMetadataBase cached;
                if (context.MetadataCache.TryGetRelationship(identity.Record.ObjectId.Value, out cached))
                    Set(results, Available(identity, RelationshipProperties(cached)));
                else remaining.Add(identity);
            }
            foreach (var batch in Batch(remaining))
            {
                cancellationToken.ThrowIfCancellationRequested();
                try
                {
                    var relationshipQuery = new RelationshipQueryExpression
                    {
                        Properties = MetadataProperties(ComponentSemanticKinds.Relationship,
                            "MetadataId", "SchemaName"),
                        Criteria = MetadataIdFilter(batch)
                    };
                    var query = new EntityQueryExpression
                    {
                        Properties = new MetadataPropertiesExpression("MetadataId", "LogicalName"),
                        RelationshipQuery = relationshipQuery
                    };
                    var response = context.Execute(new RetrieveMetadataChangesRequest { Query = query })
                        as RetrieveMetadataChangesResponse;
                    var entities = response?.EntityMetadata;
                    if (entities == null)
                    {
                        SetUnresolved(batch, results,
                            "Relationship metadata retrieval returned no metadata collection.");
                        continue;
                    }
                    var metadata = entities.SelectMany(Relationships).Where(item => item != null).ToList();
                    foreach (var item in metadata) context.MetadataCache.Store(item);
                    CorrelateMetadata(batch, metadata, item => item.MetadataId,
                        item => RelationshipProperties(item), "relationship", results,
                        collapseEquivalentDuplicates: true);
                }
                catch (OperationCanceledException) { throw; }
                catch (FaultException ex)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    SetUnresolved(batch, results,
                        "Relationship metadata retrieval failed: " + ex.Message);
                }
            }
            CopyRepeatedObjectResults(identities, representatives, results);
        }

        private static void ReadGlobalChoices(DataverseReadContext context,
            IReadOnlyList<ComponentIdentity> identities,
            IDictionary<Guid, ComponentDefinition> results, CancellationToken cancellationToken)
        {
            if (identities.Count == 0) return;
            try
            {
                cancellationToken.ThrowIfCancellationRequested();
                OptionSetMetadataBase[] returned;
                if (context.MetadataCache.OptionSetCatalogAttempted)
                    returned = context.MetadataCache.OptionSetCatalog;
                else
                {
                    var response = context.Execute(new RetrieveAllOptionSetsRequest
                    {
                        RetrieveAsIfPublished = false
                    }) as RetrieveAllOptionSetsResponse;
                    returned = response != null && response.Results.Contains("OptionSetMetadata")
                        ? response.Results["OptionSetMetadata"] as OptionSetMetadataBase[] : null;
                    context.MetadataCache.StoreOptionSetCatalog(returned,
                        returned == null ? "RetrieveAllOptionSets returned no option-set catalog." : null);
                }
                if (returned == null)
                {
                    SetUnresolved(identities, results,
                        "Global Choice metadata retrieval returned no option-set catalog" +
                        (string.IsNullOrWhiteSpace(context.MetadataCache.OptionSetCatalogFailure) ? "." :
                            ": " + context.MetadataCache.OptionSetCatalogFailure));
                    return;
                }
                var indexed = returned.Where(item => item != null && item.MetadataId.HasValue)
                    .GroupBy(item => item.MetadataId.Value)
                    .ToDictionary(group => group.Key, group => group.ToList());
                foreach (var identity in identities)
                {
                    List<OptionSetMetadataBase> matches;
                    if (!indexed.TryGetValue(identity.Record.ObjectId.Value, out matches))
                        Set(results, Unresolved(identity,
                            "No Global Choice metadata matched the component object ID."));
                    else if (matches.Count != 1)
                        Set(results, Ambiguous(identity,
                            "Multiple Global Choice metadata records matched the component object ID."));
                    else if (matches[0].IsGlobal != true || string.IsNullOrWhiteSpace(matches[0].Name))
                        Set(results, Unresolved(identity,
                            "The correlated option-set metadata did not establish a complete Global Choice definition."));
                    else Set(results, Available(identity, OptionSetProperties(matches[0])));
                }
            }
            catch (OperationCanceledException) { throw; }
            catch (FaultException ex)
            {
                cancellationToken.ThrowIfCancellationRequested();
                context.MetadataCache.StoreOptionSetCatalog(null, ex.Message);
                SetUnresolved(identities, results, "Global Choice metadata retrieval failed: " + ex.Message);
            }
        }

        private static void ReadEntityBacked(DataverseReadContext context,
            IReadOnlyList<ComponentIdentity> identities, EntityDefinitionConfiguration configuration,
            IDictionary<Guid, ComponentDefinition> results, CancellationToken cancellationToken)
        {
            if (identities.Count == 0) return;
            var remaining = new List<ComponentIdentity>();
            foreach (var group in identities.GroupBy(item => item.Record.ObjectId.Value)
                .OrderBy(item => item.Key))
            {
                Entity cached;
                if (context.MetadataCache.TryGetEntityRow(configuration.EntityName,
                    group.Key, out cached))
                {
                    foreach (var identity in group)
                        Set(results, Available(identity,
                            EntityProperties(cached, configuration.ComparableColumns),
                            configuration.PrimaryIdAttribute + "=" + group.Key.ToString("D")));
                }
                else remaining.AddRange(group);
            }
            if (remaining.Count == 0) return;
            var records = remaining.Select(item => item.Record).ToList();
            var retrieval = new BatchedDiagnosticQueryReader(context).Read(records,
                configuration.EntityName, configuration.PrimaryIdAttribute,
                configuration.Columns, "Definition retrieval returned an incomplete result set.",
                "Definition retrieval returned conflicting or incomplete primary-key data.",
                ex => "Definition retrieval failed: " + ex.Message, cancellationToken);
            var byObjectId = remaining.GroupBy(item => item.Record.ObjectId.Value)
                .ToDictionary(group => group.Key, group => group.ToList());
            foreach (var pair in byObjectId)
            {
                var correlation = retrieval.GetCorrelation(pair.Key);
                foreach (var identity in pair.Value)
                {
                    if (correlation.Status == DiagnosticCorrelationStatus.Failed)
                        Set(results, Unresolved(identity, correlation.Failure));
                    else if (correlation.Status == DiagnosticCorrelationStatus.Missing)
                        Set(results, Unresolved(identity,
                            "No " + configuration.EntityName + " row matched the component object ID."));
                    else if (correlation.Status == DiagnosticCorrelationStatus.Duplicate)
                        Set(results, Ambiguous(identity,
                            "Multiple " + configuration.EntityName + " rows matched the component object ID."));
                    else
                    {
                        context.MetadataCache.StoreEntityRow(correlation.Rows[0], pair.Key);
                        Set(results, Available(identity,
                            EntityProperties(correlation.Rows[0], configuration.ComparableColumns),
                            configuration.PrimaryIdAttribute + "=" + pair.Key.ToString("D")));
                    }
                }
            }
        }

        private static IReadOnlyList<ComponentIdentity> Pending(MembershipSnapshot membership,
            IDictionary<Guid, ComponentDefinition> results, string semanticKind) =>
            membership.Components.Where(item => item.Status == IdentityResolutionStatus.Resolved &&
                string.Equals(item.SemanticKind, semanticKind, StringComparison.OrdinalIgnoreCase) &&
                item.Record.ObjectId.HasValue && item.Record.ObjectId.Value != Guid.Empty &&
                !results.ContainsKey(item.Record.SolutionComponentId)).ToList();

        private static ComponentDefinition FromIdentityStatus(ComponentIdentity identity)
        {
            var status = identity.Status == IdentityResolutionStatus.Unsupported
                ? ComponentDefinitionReadStatus.Unsupported
                : identity.Status == IdentityResolutionStatus.Ambiguous
                    ? ComponentDefinitionReadStatus.Ambiguous
                    : ComponentDefinitionReadStatus.Unresolved;
            return new ComponentDefinition(identity, status,
                diagnostic: "Portable identity resolution is not complete: " + identity.Diagnostic,
                diagnosticEvidence: identity.DiagnosticEvidence);
        }

        private static ComponentDefinition Available(ComponentIdentity identity,
            IEnumerable<KeyValuePair<string, string>> properties, string evidence = null) =>
            new ComponentDefinition(identity, ComponentDefinitionReadStatus.Available,
                CompleteProperties(identity.SemanticKind, properties),
                diagnosticEvidence: evidence == null ? null : new[] { evidence });

        private static IEnumerable<KeyValuePair<string, string>> CompleteProperties(string semanticKind,
            IEnumerable<KeyValuePair<string, string>> properties)
        {
            var contract = ComponentDefinitionContractCatalog.For(semanticKind) ??
                throw new InvalidOperationException("No component definition contract exists for " + semanticKind + ".");
            var supplied = properties.ToDictionary(item => item.Key, item => item.Value,
                StringComparer.OrdinalIgnoreCase);
            if (supplied.Keys.Any(item => !contract.ComparableProperties.Contains(item,
                StringComparer.OrdinalIgnoreCase)))
                throw new InvalidOperationException("A definition reader returned a property outside its equality contract.");
            return contract.ComparableProperties.Select(item => new KeyValuePair<string, string>(item,
                supplied.ContainsKey(item) ? supplied[item] : null));
        }

        private static ComponentDefinition Unresolved(ComponentIdentity identity, string diagnostic) =>
            new ComponentDefinition(identity, ComponentDefinitionReadStatus.Unresolved,
                diagnostic: diagnostic);

        private static ComponentDefinition Ambiguous(ComponentIdentity identity, string diagnostic) =>
            new ComponentDefinition(identity, ComponentDefinitionReadStatus.Ambiguous,
                diagnostic: diagnostic);

        private static void Set(IDictionary<Guid, ComponentDefinition> results,
            ComponentDefinition definition) =>
            results[definition.Identity.Record.SolutionComponentId] = definition;

        private static void SetUnresolved(IEnumerable<ComponentIdentity> identities,
            IDictionary<Guid, ComponentDefinition> results, string diagnostic)
        {
            foreach (var identity in identities) Set(results, Unresolved(identity, diagnostic));
        }

        private static IEnumerable<IReadOnlyList<ComponentIdentity>> Batch(
            IReadOnlyList<ComponentIdentity> identities)
        {
            for (int offset = 0; offset < identities.Count; offset += BatchSize)
                yield return identities.Skip(offset).Take(Math.Min(BatchSize,
                    identities.Count - offset)).ToList();
        }

        private static MetadataPropertiesExpression MetadataProperties(string semanticKind,
            params string[] correlationProperties)
        {
            var contract = ComponentDefinitionContractCatalog.For(semanticKind);
            return new MetadataPropertiesExpression(correlationProperties
                .Concat(contract.ComparableProperties)
                .Distinct(StringComparer.OrdinalIgnoreCase).ToArray());
        }

        private static MetadataFilterExpression MetadataIdFilter(
            IEnumerable<ComponentIdentity> identities)
        {
            var filter = new MetadataFilterExpression(LogicalOperator.Or);
            foreach (var identity in identities)
                filter.Conditions.Add(new MetadataConditionExpression("MetadataId",
                    MetadataConditionOperator.Equals, identity.Record.ObjectId.Value));
            return filter;
        }

        private static IEnumerable<RelationshipMetadataBase> Relationships(EntityMetadata entity)
        {
            if (entity == null) return Enumerable.Empty<RelationshipMetadataBase>();
            return (entity.OneToManyRelationships ?? new OneToManyRelationshipMetadata[0])
                .Cast<RelationshipMetadataBase>()
                .Concat((entity.ManyToOneRelationships ?? new OneToManyRelationshipMetadata[0])
                    .Cast<RelationshipMetadataBase>())
                .Concat((entity.ManyToManyRelationships ?? new ManyToManyRelationshipMetadata[0])
                    .Cast<RelationshipMetadataBase>());
        }

        private static void CorrelateMetadata<T>(IReadOnlyList<ComponentIdentity> identities,
            IEnumerable<T> returned, Func<T, Guid?> metadataId,
            Func<T, IEnumerable<KeyValuePair<string, string>>> properties, string family,
            IDictionary<Guid, ComponentDefinition> results,
            bool collapseEquivalentDuplicates = false)
        {
            var requested = new HashSet<Guid>(identities.Select(item => item.Record.ObjectId.Value));
            var rows = returned.ToList();
            if (rows.Any(item => !metadataId(item).HasValue ||
                !requested.Contains(metadataId(item).Value)))
            {
                SetUnresolved(identities, results, family +
                    " metadata retrieval returned conflicting or incomplete identifiers.");
                return;
            }
            var indexed = rows.GroupBy(item => metadataId(item).Value)
                .ToDictionary(group => group.Key, group => group.ToList());
            foreach (var identity in identities)
            {
                List<T> matches;
                if (!indexed.TryGetValue(identity.Record.ObjectId.Value, out matches))
                {
                    Set(results, Unresolved(identity, "No " + family +
                        " metadata matched the component object ID."));
                    continue;
                }
                if (collapseEquivalentDuplicates && matches.Count > 1)
                    matches = matches.GroupBy(item => PropertySignature(
                        CompleteProperties(identity.SemanticKind, properties(item))),
                        StringComparer.Ordinal).Select(group => group.First()).ToList();
                if (matches.Count != 1)
                    Set(results, Ambiguous(identity, "Multiple conflicting " + family +
                        " metadata records matched the component object ID."));
                else Set(results, Available(identity, properties(matches[0])));
            }
        }

        private static string PropertySignature(IEnumerable<KeyValuePair<string, string>> properties) =>
            string.Join("\u001f", properties.OrderBy(item => item.Key,
                StringComparer.OrdinalIgnoreCase).Select(item => item.Key.Length.ToString(
                    CultureInfo.InvariantCulture) + ":" + item.Key + "=" +
                    (item.Value == null ? "<null>" : item.Value.Length.ToString(
                        CultureInfo.InvariantCulture) + ":" + item.Value)));

        private static IReadOnlyList<ComponentIdentity> DistinctRepresentatives(
            IEnumerable<ComponentIdentity> identities) => identities
                .GroupBy(item => item.Record.ObjectId.Value)
                .OrderBy(group => group.Key)
                .Select(group => group.First()).ToList();

        private static void CopyRepeatedObjectResults(IEnumerable<ComponentIdentity> identities,
            IReadOnlyList<ComponentIdentity> representatives,
            IDictionary<Guid, ComponentDefinition> results)
        {
            var representativeByObjectId = representatives.ToDictionary(
                item => item.Record.ObjectId.Value);
            foreach (var identity in identities)
            {
                var representative = representativeByObjectId[identity.Record.ObjectId.Value];
                if (ReferenceEquals(identity, representative)) continue;
                var definition = results[representative.Record.SolutionComponentId];
                Set(results, new ComponentDefinition(identity, definition.Status,
                    definition.ComparableProperties, definition.Diagnostic,
                    definition.DiagnosticEvidence));
            }
        }

        private static IEnumerable<KeyValuePair<string, string>> TableProperties(EntityMetadata metadata)
        {
            yield return Pair("SchemaName", metadata.SchemaName);
            yield return Pair("OwnershipType", metadata.OwnershipType);
            yield return Pair("IsActivity", metadata.IsActivity);
            yield return Pair("IsIntersect", metadata.IsIntersect);
            yield return Pair("HasActivities", metadata.HasActivities);
            yield return Pair("HasNotes", metadata.HasNotes);
            yield return Pair("IsAuditEnabled", metadata.IsAuditEnabled?.Value);
            yield return Pair("IsValidForAdvancedFind", metadata.IsValidForAdvancedFind);
            yield return Pair("IsDuplicateDetectionEnabled", metadata.IsDuplicateDetectionEnabled?.Value);
            yield return Pair("IsCustomizable", metadata.IsCustomizable?.Value);
        }

        private static IEnumerable<KeyValuePair<string, string>> AttributeProperties(AttributeMetadata metadata)
        {
            yield return Pair("SchemaName", metadata.SchemaName);
            yield return Pair("AttributeType", metadata.AttributeType);
            yield return Pair("AttributeTypeName", metadata.AttributeTypeName?.Value);
            yield return Pair("RequiredLevel", metadata.RequiredLevel?.Value);
            yield return Pair("IsAuditEnabled", metadata.IsAuditEnabled?.Value);
            yield return Pair("IsSecured", metadata.IsSecured);
            yield return Pair("IsValidForAdvancedFind", metadata.IsValidForAdvancedFind?.Value);
            yield return Pair("IsCustomizable", metadata.IsCustomizable?.Value);

            var text = metadata as StringAttributeMetadata;
            if (text != null) yield return Pair("MaxLength", text.MaxLength);
            var memo = metadata as MemoAttributeMetadata;
            if (memo != null) yield return Pair("MaxLength", memo.MaxLength);
            var integer = metadata as IntegerAttributeMetadata;
            if (integer != null)
            {
                yield return Pair("MinValue", integer.MinValue);
                yield return Pair("MaxValue", integer.MaxValue);
                yield return Pair("IntegerFormat", integer.Format);
            }
            var decimalValue = metadata as DecimalAttributeMetadata;
            if (decimalValue != null)
            {
                yield return Pair("MinValue", decimalValue.MinValue);
                yield return Pair("MaxValue", decimalValue.MaxValue);
                yield return Pair("Precision", decimalValue.Precision);
            }
            var money = metadata as MoneyAttributeMetadata;
            if (money != null)
            {
                yield return Pair("MinValue", money.MinValue);
                yield return Pair("MaxValue", money.MaxValue);
                yield return Pair("Precision", money.Precision);
                yield return Pair("PrecisionSource", money.PrecisionSource);
            }
            var dateTime = metadata as DateTimeAttributeMetadata;
            if (dateTime != null)
            {
                yield return Pair("DateTimeFormat", dateTime.Format);
                yield return Pair("DateTimeBehavior", dateTime.DateTimeBehavior?.Value);
            }
            var lookup = metadata as LookupAttributeMetadata;
            if (lookup != null)
                yield return Pair("Targets", lookup.Targets == null ? null :
                    string.Join("|", lookup.Targets.OrderBy(item => item, StringComparer.OrdinalIgnoreCase)));
            var enumMetadata = metadata as EnumAttributeMetadata;
            if (enumMetadata?.OptionSet != null)
                yield return Pair("Options", SerializeOptions(enumMetadata.OptionSet.Options));
        }

        private static IEnumerable<KeyValuePair<string, string>> RelationshipProperties(
            RelationshipMetadataBase metadata)
        {
            yield return Pair("RelationshipType", metadata.RelationshipType);
            yield return Pair("IsCustomizable", metadata.IsCustomizable?.Value);
            var oneToMany = metadata as OneToManyRelationshipMetadata;
            if (oneToMany != null)
            {
                yield return Pair("ReferencedEntity", oneToMany.ReferencedEntity);
                yield return Pair("ReferencedAttribute", oneToMany.ReferencedAttribute);
                yield return Pair("ReferencingEntity", oneToMany.ReferencingEntity);
                yield return Pair("ReferencingAttribute", oneToMany.ReferencingAttribute);
                yield return Pair("CascadeAssign", oneToMany.CascadeConfiguration?.Assign);
                yield return Pair("CascadeDelete", oneToMany.CascadeConfiguration?.Delete);
                yield return Pair("CascadeMerge", oneToMany.CascadeConfiguration?.Merge);
                yield return Pair("CascadeReparent", oneToMany.CascadeConfiguration?.Reparent);
                yield return Pair("CascadeShare", oneToMany.CascadeConfiguration?.Share);
                yield return Pair("CascadeUnshare", oneToMany.CascadeConfiguration?.Unshare);
            }
            var manyToMany = metadata as ManyToManyRelationshipMetadata;
            if (manyToMany != null)
            {
                yield return Pair("Entity1LogicalName", manyToMany.Entity1LogicalName);
                yield return Pair("Entity2LogicalName", manyToMany.Entity2LogicalName);
                yield return Pair("IntersectEntityName", manyToMany.IntersectEntityName);
            }
        }

        private static IEnumerable<KeyValuePair<string, string>> OptionSetProperties(
            OptionSetMetadataBase metadata)
        {
            yield return Pair("OptionSetType", metadata.OptionSetType);
            yield return Pair("IsCustomOptionSet", metadata.IsCustomOptionSet);
            var optionSet = metadata as OptionSetMetadata;
            if (optionSet != null) yield return Pair("Options", SerializeOptions(optionSet.Options));
        }

        private static string SerializeOptions(IEnumerable<OptionMetadata> options)
        {
            if (options == null) return null;
            return string.Join("|", options.Where(item => item != null)
                .OrderBy(item => item.Value)
                .Select(item => Value(item.Value) + ":" + SerializeLabel(item.Label)));
        }

        private static string SerializeLabel(Label label)
        {
            if (label?.LocalizedLabels == null) return null;
            return string.Join(";", label.LocalizedLabels.Where(item => item != null)
                .OrderBy(item => item.LanguageCode)
                .Select(item => item.LanguageCode.ToString(CultureInfo.InvariantCulture) + "=" +
                    (item.Label ?? string.Empty)));
        }

        private static IEnumerable<KeyValuePair<string, string>> EntityProperties(Entity entity,
            IEnumerable<string> columns)
        {
            foreach (var column in columns)
                yield return Pair(column, entity.Attributes.ContainsKey(column)
                    ? EntityValue(entity.Attributes[column]) : null);
        }

        private static string EntityValue(object value)
        {
            if (value == null) return null;
            var option = value as OptionSetValue;
            if (option != null) return option.Value.ToString(CultureInfo.InvariantCulture);
            var managed = value as BooleanManagedProperty;
            if (managed != null) return managed.Value.ToString(CultureInfo.InvariantCulture);
            if (value is bool) return ((bool)value).ToString(CultureInfo.InvariantCulture);
            if (value is int) return ((int)value).ToString(CultureInfo.InvariantCulture);
            if (value is decimal) return ((decimal)value).ToString(CultureInfo.InvariantCulture);
            return value as string ?? Convert.ToString(value, CultureInfo.InvariantCulture);
        }

        private static KeyValuePair<string, string> Pair(string key, object value) =>
            new KeyValuePair<string, string>(key, Value(value));

        private static string Value(object value)
        {
            if (value == null) return null;
            if (value is bool) return ((bool)value).ToString(CultureInfo.InvariantCulture);
            var formattable = value as IFormattable;
            return formattable != null ? formattable.ToString(null, CultureInfo.InvariantCulture) : value.ToString();
        }

        private sealed class EntityDefinitionConfiguration
        {
            private EntityDefinitionConfiguration(string entityName, string primaryIdAttribute,
                string[] columns, string[] comparableColumns)
            {
                EntityName = entityName;
                PrimaryIdAttribute = primaryIdAttribute;
                Columns = columns;
                ComparableColumns = comparableColumns;
            }

            public string EntityName { get; }
            public string PrimaryIdAttribute { get; }
            public IReadOnlyList<string> Columns { get; }
            public IReadOnlyList<string> ComparableColumns { get; }

            public static readonly EntityDefinitionConfiguration WebResource = Create("webresource",
                "webresourceid", "name", "webresourcetype", "content", "displayname", "description");
            public static readonly EntityDefinitionConfiguration EnvironmentVariableDefinition = Create(
                "environmentvariabledefinition", "environmentvariabledefinitionid", "schemaname", "type",
                "defaultvalue", "valueschema", "isrequired", "secretstore", "displayname", "description");
            public static readonly EntityDefinitionConfiguration ConnectionReference = Create(
                "connectionreference", "connectionreferenceid", "connectionreferencelogicalname",
                "connectionreferencedisplayname", "connectorid", "description");
            public static readonly EntityDefinitionConfiguration AppModule = Create("appmodule",
                "appmoduleid", "uniquename", "name", "description", "clienttype", "formfactor",
                "navigationtype");

            private static EntityDefinitionConfiguration Create(string entityName, string primaryId,
                string identityColumn, params string[] comparableColumns)
            {
                return new EntityDefinitionConfiguration(entityName, primaryId,
                    new[] { primaryId, identityColumn }.Concat(comparableColumns).Distinct(
                        StringComparer.OrdinalIgnoreCase).ToArray(), comparableColumns);
            }
        }
    }
}
