using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.ServiceModel;
using System.Threading;
using D365SolutionComparer.Infrastructure;
using D365SolutionComparer.Models.Membership;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Query;

namespace D365SolutionComparer.Services.Membership
{
    /// <summary>
    /// Temporary, read-only evidence operation for investigating AppSetting solution components.
    /// It deliberately has no relationship with the normal identity resolver or comparer.
    /// </summary>
    internal sealed class DataverseAppSettingEvidenceOperation
    {
        internal const int BatchSize = 200;
        private static readonly string[] CandidateEntities = { "appsetting", "settingdefinition", "appmodule" };

        internal AppSettingEvidenceReport Capture(IOrganizationService service,
            MembershipSnapshot snapshot, CancellationToken cancellationToken,
            Action<string> progress = null, DataverseRequestCounter requestCounter = null)
        {
            if (snapshot == null) throw new ArgumentNullException(nameof(snapshot));
            if (snapshot.State != MembershipSnapshotState.Complete || snapshot.Solution == null)
                throw new InvalidOperationException("AppSetting evidence requires a complete solution membership snapshot.");
            cancellationToken.ThrowIfCancellationRequested();
            var counter = requestCounter ?? new DataverseRequestCounter();
            var diagnostics = new List<string>();
            var context = new DataverseReadContext(service, snapshot.Environment, cancellationToken, counter);
            progress?.Invoke("Reading solutioncomponent evidence...");
            var solutionComponents = ReadSolutionComponents(context, snapshot, cancellationToken);
            var summaries = BuildSummaries(solutionComponents);
            var candidateTypes = FindCandidateComponentTypes(
                snapshot.Components.Select(item => item.RegisteredDefinition), diagnostics);
            if (candidateTypes.Count == 0)
                diagnostics.Add("No unambiguous AppSetting/appsetting registered family was available in the completed membership snapshot.");
            else diagnostics.Add("Candidate types selected from completed membership registered definitions (Name=AppSetting; PrimaryEntityName=appsetting): " +
                string.Join(", ", candidateTypes.Select(item => item.ToString(CultureInfo.InvariantCulture))) + ".");

            progress?.Invoke("Inspecting candidate entity metadata...");
            var entityInfos = candidateTypes.Count == 0
                ? new List<EntityTypeInfo>()
                : CandidateEntities.Select(name => LoadMetadata(context, name,
                    cancellationToken, diagnostics)).ToList();
            var candidates = new List<AppSettingCandidateEvidence>();
            foreach (var type in candidateTypes)
            {
                cancellationToken.ThrowIfCancellationRequested();
                var rows = solutionComponents.Where(item => item.ComponentType == type).ToList();
                var ids = rows.Where(item => item.ObjectId.HasValue && item.ObjectId.Value != Guid.Empty)
                    .Select(item => item.ObjectId.Value).Distinct().OrderBy(item => item).ToList();
                progress?.Invoke("Correlating App Setting candidates for component type " +
                    type.ToString(CultureInfo.InvariantCulture) + "...");
                var correlation = new Dictionary<string, EntityCorrelationSet>(StringComparer.OrdinalIgnoreCase);
                foreach (var info in entityInfos)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    correlation[info.LogicalName] = QueryEntityByIds(context, info, ids, cancellationToken,
                        diagnostics);
                }
                var relatedSettingDefinitions = QueryRelated(context, entityInfos, correlation, "settingdefinition",
                    new[] { "settingdefinitionid", "settingdefinition", "_settingdefinitionid_value" },
                    cancellationToken, diagnostics);
                var relatedAppModules = QueryRelated(context, entityInfos, correlation, "appmodule",
                    new[] { "parentappmoduleid", "parentappmodule", "appmoduleid", "_parentappmoduleid_value" },
                    cancellationToken, diagnostics);

                foreach (var row in rows)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    var correlations = new List<AppSettingEntityCorrelation>();
                    foreach (var info in entityInfos)
                    {
                        EntityCorrelationSet set;
                        if (!correlation.TryGetValue(info.LogicalName, out set)) continue;
                        correlations.Add(set.For(row.ObjectId));
                    }
                    var appSet = correlation.ContainsKey("appsetting")
                        ? correlation["appsetting"].GetSingle(row.ObjectId) : null;
                    var settingId = GetRelatedGuid(appSet, "settingdefinitionid", "settingdefinition",
                        "_settingdefinitionid_value");
                    var appModuleId = GetRelatedGuid(appSet, "parentappmoduleid", "parentappmodule",
                        "appmoduleid", "_parentappmoduleid_value");
                    var setting = settingId.HasValue ? relatedSettingDefinitions.GetSingle(settingId) : null;
                    var appModule = appModuleId.HasValue ? relatedAppModules.GetSingle(appModuleId) : null;
                    var settingName = GetPortableName(setting);
                    var parentUniqueName = GetString(appModule, "uniquename");
                    var parentName = GetString(appModule, "name");
                    var composite = !string.IsNullOrWhiteSpace(settingName) &&
                        !string.IsNullOrWhiteSpace(parentUniqueName)
                        ? parentUniqueName + "+" + settingName : string.Empty;
                    var state = DetermineState(row, correlations, settingId, setting, appModuleId, appModule,
                        composite);
                    candidates.Add(new AppSettingCandidateEvidence(snapshot.Environment.DisplayName,
                        snapshot.SolutionUniqueName, snapshot.Solution.SolutionId, row.SolutionComponentId,
                        row.ComponentType, row.FormattedLabel, row.ObjectId, row.RootComponentBehavior,
                        row.RootSolutionComponentId, row.IsMetadata,
                        "Registered solutioncomponentdefinition: Name=AppSetting; PrimaryEntityName=appsetting (ordinal case-insensitive).", correlations, state,
                        settingId.HasValue ? settingId.Value.ToString("D") : string.Empty, settingName,
                        appModuleId.HasValue ? appModuleId.Value.ToString("D") : string.Empty,
                        parentUniqueName, parentName, composite,
                        BuildCandidateDiagnostic(row, state, settingId, setting, appModuleId, appModule)));
                }
            }
            cancellationToken.ThrowIfCancellationRequested();
            var requests = BuildRequestSummary(counter);
            return new AppSettingEvidenceReport(snapshot.Environment.DisplayName, snapshot.SolutionUniqueName,
                snapshot.Solution.SolutionId, summaries, entityInfos.Select(item => item.Evidence), candidates,
                requests, diagnostics);
        }

        private static List<SolutionComponentEvidence> ReadSolutionComponents(DataverseReadContext context,
            MembershipSnapshot snapshot, CancellationToken cancellationToken)
        {
            var query = CreateSolutionComponentEvidenceQuery(snapshot.Solution.SolutionId);
            var rows = new DataversePagedReader(context.Service).ReadAll(query, "solutioncomponentid", cancellationToken);
            return rows.Select(row => new SolutionComponentEvidence
            {
                SolutionComponentId = row.Id,
                ObjectId = row.GetAttributeValue<Guid?>("objectid"),
                ComponentType = row.GetAttributeValue<OptionSetValue>("componenttype")?.Value ?? -1,
                FormattedLabel = GetFormatted(row, "componenttype"),
                RootComponentBehavior = row.GetAttributeValue<OptionSetValue>("rootcomponentbehavior")?.Value,
                RootSolutionComponentId = row.GetAttributeValue<Guid?>("rootsolutioncomponentid"),
                IsMetadata = row.GetAttributeValue<bool?>("ismetadata")
            }).ToList();
        }

        internal static QueryExpression CreateSolutionComponentEvidenceQuery(Guid solutionId)
        {
            var query = new QueryExpression("solutioncomponent")
            {
                ColumnSet = new ColumnSet("solutioncomponentid", "solutionid", "componenttype", "objectid",
                    "rootcomponentbehavior", "rootsolutioncomponentid", "ismetadata")
            };
            query.Criteria.AddCondition("solutionid", ConditionOperator.Equal, solutionId);
            return query;
        }

        internal static IReadOnlyList<int> FindCandidateComponentTypes(
            IEnumerable<SolutionComponentDefinitionIdentity> registeredFamilies,
            IList<string> diagnostics = null)
        {
            // Reuse the structured registrations retained by the membership resolver and displayed
            // by Coverage Details. Labels and environment-local numeric constants are not evidence
            // of family identity. Repeated copies on raw membership rows are expected.
            var selected = new List<int>();
            foreach (var group in (registeredFamilies ?? Enumerable.Empty<SolutionComponentDefinitionIdentity>())
                .Where(item => item != null).GroupBy(item => item.ObjectTypeCode).OrderBy(group => group.Key))
            {
                var first = group.First();
                if (group.Any(item => !string.Equals(item.Name, first.Name, StringComparison.OrdinalIgnoreCase) ||
                    !string.Equals(item.PrimaryEntityName, first.PrimaryEntityName, StringComparison.OrdinalIgnoreCase)))
                {
                    diagnostics?.Add("Ambiguous registered definitions for raw component type " +
                        group.Key.ToString(CultureInfo.InvariantCulture) + "; excluded from AppSetting evidence candidates.");
                    continue;
                }
                if (string.Equals(first.Name, "AppSetting", StringComparison.OrdinalIgnoreCase) &&
                    string.Equals(first.PrimaryEntityName, "appsetting", StringComparison.OrdinalIgnoreCase))
                    selected.Add(group.Key);
            }
            return selected;
        }

        private static List<AppSettingComponentTypeSummary> BuildSummaries(IEnumerable<SolutionComponentEvidence> rows)
        {
            return rows.GroupBy(row => row.ComponentType)
                .OrderBy(group => group.Key)
                .Select(group => new AppSettingComponentTypeSummary(group.Key,
                    group.Select(row => row.FormattedLabel).FirstOrDefault(label =>
                        !string.IsNullOrWhiteSpace(label)) ?? string.Empty, group.Count(),
                    group.Select(row => row.ObjectId)
                        .Where(id => id.HasValue && id.Value != Guid.Empty).Distinct().Count(),
                    group.Count(row => !row.ObjectId.HasValue || row.ObjectId == Guid.Empty))).ToList();
        }

        private static EntityTypeInfo LoadMetadata(DataverseReadContext context, string logicalName,
            CancellationToken cancellationToken, IList<string> diagnostics)
        {
            try
            {
                var response = context.Execute(new RetrieveEntityRequest
                {
                    LogicalName = logicalName,
                    EntityFilters = EntityFilters.Entity | EntityFilters.Attributes | EntityFilters.Relationships,
                    RetrieveAsIfPublished = true
                }) as RetrieveEntityResponse;
                var metadata = response?.EntityMetadata;
                if (metadata == null || string.IsNullOrWhiteSpace(metadata.PrimaryIdAttribute))
                    throw new InvalidOperationException("No usable entity metadata was returned.");
                var attributes = (metadata.Attributes ?? new AttributeMetadata[0]).Where(item =>
                    item != null && !string.IsNullOrWhiteSpace(item.LogicalName)).Select(item => item.LogicalName)
                    .Distinct(StringComparer.OrdinalIgnoreCase).ToList();
                var lookups = (metadata.Attributes ?? new AttributeMetadata[0]).Where(item =>
                    item != null && !string.IsNullOrWhiteSpace(item.LogicalName) &&
                    item.AttributeType.HasValue && (item.AttributeType.Value == AttributeTypeCode.Lookup ||
                    item.AttributeType.Value == AttributeTypeCode.Customer || item.AttributeType.Value == AttributeTypeCode.Owner))
                    .Select(item => item.LogicalName).Distinct(StringComparer.OrdinalIgnoreCase).ToList();
                var relationships = (metadata.OneToManyRelationships ?? new OneToManyRelationshipMetadata[0])
                    .Cast<RelationshipMetadataBase>()
                    .Concat(metadata.ManyToOneRelationships == null
                        ? Enumerable.Empty<RelationshipMetadataBase>()
                        : metadata.ManyToOneRelationships.Cast<RelationshipMetadataBase>())
                    .Concat(metadata.ManyToManyRelationships == null
                        ? Enumerable.Empty<RelationshipMetadataBase>()
                        : metadata.ManyToManyRelationships.Cast<RelationshipMetadataBase>())
                    .Where(item => item != null && !string.IsNullOrWhiteSpace(item.SchemaName))
                    .Select(item => item.SchemaName)
                    .Distinct(StringComparer.OrdinalIgnoreCase).ToList();
                return new EntityTypeInfo(logicalName, metadata.PrimaryIdAttribute,
                    metadata.PrimaryNameAttribute, attributes, lookups, relationships,
                    new AppSettingEntityMetadataEvidence(logicalName, AppSettingEvidenceState.Confirmed,
                        metadata.PrimaryIdAttribute, metadata.PrimaryNameAttribute,
                        attributes, lookups, relationships, "Entity metadata retrieved successfully."));
            }
            catch (OperationCanceledException) { throw; }
            catch (FaultException ex)
            {
                diagnostics.Add(logicalName + " metadata unavailable: " + ex.Message);
                return EntityTypeInfo.Unavailable(logicalName, AppSettingEvidenceState.Unavailable,
                    ex.GetType().Name + ": " + ex.Message);
            }
            catch (Exception ex)
            {
                diagnostics.Add(logicalName + " metadata failed: " + ex.Message);
                return EntityTypeInfo.Unavailable(logicalName, AppSettingEvidenceState.Faulted,
                    ex.GetType().Name + ": " + ex.Message);
            }
        }

        private static EntityCorrelationSet QueryEntityByIds(DataverseReadContext context, EntityTypeInfo info,
            IReadOnlyList<Guid> ids, CancellationToken cancellationToken, IList<string> diagnostics)
        {
            if (info.State != AppSettingEvidenceState.Confirmed)
                return EntityCorrelationSet.Unavailable(info.LogicalName, info.PrimaryIdAttribute, info.State,
                    info.Evidence.Diagnostic);
            if (ids.Count == 0)
                return EntityCorrelationSet.Empty(info.LogicalName, info.PrimaryIdAttribute);
            var fields = SelectFields(info);
            var returned = new List<Entity>();
            try
            {
                for (int offset = 0; offset < ids.Count; offset += BatchSize)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    var batch = ids.Skip(offset).Take(BatchSize).ToArray();
                    var query = new QueryExpression(info.LogicalName)
                    {
                        ColumnSet = new ColumnSet(fields.ToArray())
                    };
                    query.Criteria.AddCondition(info.PrimaryIdAttribute, ConditionOperator.In,
                        batch.Cast<object>().ToArray());
                    returned.AddRange(new DataversePagedReader(context.Service).ReadAll(query,
                        info.PrimaryIdAttribute, cancellationToken));
                }
                var unexpected = returned.Where(entity => entity == null || entity.Id == Guid.Empty ||
                    !ids.Contains(entity.Id)).ToList();
                if (unexpected.Count > 0)
                    diagnostics.Add(info.LogicalName + " returned unexpected primary-key rows.");
                return new EntityCorrelationSet(info, returned);
            }
            catch (OperationCanceledException) { throw; }
            catch (FaultException ex)
            {
                diagnostics.Add(info.LogicalName + " query failed: " + ex.Message);
                return EntityCorrelationSet.Unavailable(info.LogicalName, info.PrimaryIdAttribute,
                    AppSettingEvidenceState.Faulted, ex.GetType().Name + ": " + ex.Message);
            }
            catch (Exception ex)
            {
                diagnostics.Add(info.LogicalName + " query failed: " + ex.Message);
                return EntityCorrelationSet.Unavailable(info.LogicalName, info.PrimaryIdAttribute,
                    AppSettingEvidenceState.Faulted, ex.GetType().Name + ": " + ex.Message);
            }
        }

        private static EntityCorrelationSet QueryRelated(DataverseReadContext context, IReadOnlyList<EntityTypeInfo> infos,
            IDictionary<string, EntityCorrelationSet> direct, string logicalName, string[] relationNames,
            CancellationToken cancellationToken, IList<string> diagnostics)
        {
            var info = infos.First(item => string.Equals(item.LogicalName, logicalName,
                StringComparison.OrdinalIgnoreCase));
            var ids = direct.Values.SelectMany(item => item.RecordsById.Values.SelectMany(list => list))
                .Select(record => GetRelatedGuid(record, relationNames)).Where(id => id.HasValue)
                .Select(id => id.Value).Distinct().OrderBy(id => id).ToList();
            EntityCorrelationSet existing;
            if (!direct.TryGetValue(logicalName, out existing))
                existing = EntityCorrelationSet.Empty(logicalName, info.PrimaryIdAttribute);
            if (existing.State != AppSettingEvidenceState.Confirmed) return existing;
            var missing = ids.Where(id => !existing.RecordsById.ContainsKey(id)).ToList();
            if (missing.Count == 0) return existing;
            return existing.Merge(QueryEntityByIds(context, info, missing, cancellationToken, diagnostics));
        }

        private static AppSettingEvidenceState DetermineState(SolutionComponentEvidence row,
            IEnumerable<AppSettingEntityCorrelation> correlations, Guid? settingId,
            AppSettingRecordEvidence setting, Guid? appModuleId, AppSettingRecordEvidence appModule,
            string composite)
        {
            if (!row.ObjectId.HasValue || row.ObjectId.Value == Guid.Empty) return AppSettingEvidenceState.Unresolved;
            var app = correlations.FirstOrDefault(item => string.Equals(item.EntityLogicalName, "appsetting",
                StringComparison.OrdinalIgnoreCase));
            if (app == null || app.State == AppSettingEvidenceState.Unavailable) return AppSettingEvidenceState.Unavailable;
            if (app.State != AppSettingEvidenceState.Confirmed) return app.State;
            if (app.MatchedRecordCount != 1) return app.MatchedRecordCount > 1
                ? AppSettingEvidenceState.Ambiguous : AppSettingEvidenceState.Unresolved;
            return !string.IsNullOrWhiteSpace(composite) ? AppSettingEvidenceState.Confirmed :
                AppSettingEvidenceState.Unresolved;
        }

        private static string BuildCandidateDiagnostic(SolutionComponentEvidence row,
            AppSettingEvidenceState state, Guid? settingId, AppSettingRecordEvidence setting,
            Guid? appModuleId, AppSettingRecordEvidence appModule)
        {
            var text = "Evidence state=" + state + ".";
            if (!row.ObjectId.HasValue) return text + " solutioncomponent.objectid is blank.";
            if (settingId.HasValue && setting == null) text += " Setting Definition correlation is missing.";
            if (appModuleId.HasValue && appModule == null) text += " Parent AppModule correlation is missing.";
            return text;
        }

        private static IEnumerable<string> SelectFields(EntityTypeInfo info)
        {
            var preferred = new[] { info.PrimaryIdAttribute, info.PrimaryNameAttribute, "uniquename",
                "name", "displayname", "title", "settingdefinitionid", "settingdefinition",
                "_settingdefinitionid_value", "parentappmoduleid", "parentappmodule", "appmoduleid",
                "_parentappmoduleid_value", "componentstate", "ismanaged", "isglobal", "isoverridable",
                "overridablelevel", "releaselevel", "datatype", "defaultvalue", "description", "informationurl",
                "value", "configuration", "languagecode", "objecttypecode" };
            return preferred.Where(name => !string.IsNullOrWhiteSpace(name) &&
                (string.Equals(name, info.PrimaryIdAttribute, StringComparison.OrdinalIgnoreCase) ||
                 string.Equals(name, info.PrimaryNameAttribute, StringComparison.OrdinalIgnoreCase) ||
                 info.Attributes.Contains(name, StringComparer.OrdinalIgnoreCase))).Distinct(StringComparer.OrdinalIgnoreCase);
        }

        private static Guid? GetRelatedGuid(AppSettingRecordEvidence record, params string[] names)
        {
            if (record == null) return null;
            foreach (var name in names)
            {
                string value;
                if (record.Fields.TryGetValue(name, out value) && Guid.TryParse(value, out var id) && id != Guid.Empty)
                    return id;
            }
            return null;
        }

        private static string GetPortableName(AppSettingRecordEvidence record)
        {
            if (record == null) return string.Empty;
            foreach (var key in new[] { "uniquename", "name", "schemaname", "logicalname" })
            {
                string value;
                if (record.Fields.TryGetValue(key, out value) && !string.IsNullOrWhiteSpace(value)) return value;
            }
            return string.Empty;
        }

        private static string GetString(AppSettingRecordEvidence record, string name)
        {
            if (record == null) return string.Empty;
            string value;
            return record.Fields.TryGetValue(name, out value) ? value : string.Empty;
        }

        private static string GetFormatted(Entity row, string attribute)
        {
            string value;
            return row != null && row.FormattedValues.TryGetValue(attribute, out value) ? value : string.Empty;
        }

        private static AppSettingRequestSummary BuildRequestSummary(DataverseRequestCounter counter)
        {
            var who = counter.GetExecuteCount("WhoAmI");
            var metadata = counter.GetExecuteCount("RetrieveEntity");
            var solution = counter.GetQueryCount("solutioncomponent");
            var setting = counter.GetQueryCount("settingdefinition");
            var appModule = counter.GetQueryCount("appmodule");
            var candidate = counter.GetQueryCount("appsetting");
            var known = who + metadata + solution + setting + appModule + candidate;
            return new AppSettingRequestSummary(who, solution, metadata, candidate, setting, appModule,
                Math.Max(0, counter.TotalRequests - known), counter.TotalRequests);
        }

        private sealed class EntityTypeInfo
        {
            internal EntityTypeInfo(string logicalName, string primaryId, string primaryName,
                IEnumerable<string> attributes, IEnumerable<string> lookups, IEnumerable<string> relationships,
                AppSettingEntityMetadataEvidence evidence, AppSettingEvidenceState state)
            {
                LogicalName = logicalName;
                PrimaryIdAttribute = primaryId ?? string.Empty;
                PrimaryNameAttribute = primaryName ?? string.Empty;
                Attributes = attributes.ToList();
                Lookups = lookups.ToList();
                Relationships = relationships.ToList();
                Evidence = evidence;
                State = state;
            }

            internal EntityTypeInfo(string logicalName, string primaryId, string primaryName,
                IEnumerable<string> attributes, IEnumerable<string> lookups, IEnumerable<string> relationships,
                AppSettingEntityMetadataEvidence evidence)
                : this(logicalName, primaryId, primaryName, attributes, lookups, relationships, evidence,
                    AppSettingEvidenceState.Confirmed) { }

            internal static EntityTypeInfo Unavailable(string logicalName, AppSettingEvidenceState state,
                string diagnostic) => new EntityTypeInfo(logicalName, string.Empty, string.Empty,
                    Enumerable.Empty<string>(), Enumerable.Empty<string>(),
                    Enumerable.Empty<string>(),
                    new AppSettingEntityMetadataEvidence(logicalName, state, string.Empty, string.Empty,
                        Enumerable.Empty<string>(), Enumerable.Empty<string>(), Enumerable.Empty<string>(), diagnostic), state);

            internal string LogicalName { get; }
            internal string PrimaryIdAttribute { get; }
            internal string PrimaryNameAttribute { get; }
            internal List<string> Attributes { get; }
            internal List<string> Lookups { get; }
            internal List<string> Relationships { get; }
            internal AppSettingEntityMetadataEvidence Evidence { get; }
            internal AppSettingEvidenceState State { get; }
        }

        private sealed class SolutionComponentEvidence
        {
            internal Guid SolutionComponentId;
            internal Guid? ObjectId;
            internal int ComponentType;
            internal string FormattedLabel;
            internal int? RootComponentBehavior;
            internal Guid? RootSolutionComponentId;
            internal bool? IsMetadata;
        }

        private sealed class EntityCorrelationSet
        {
            private readonly EntityTypeInfo info;
            internal readonly Dictionary<Guid, List<Entity>> RecordsById = new Dictionary<Guid, List<Entity>>();

            internal EntityCorrelationSet(EntityTypeInfo info, IEnumerable<Entity> records)
            {
                this.info = info;
                foreach (var record in records.Where(item => item != null))
                {
                    List<Entity> list;
                    if (!RecordsById.TryGetValue(record.Id, out list)) RecordsById[record.Id] = list = new List<Entity>();
                    list.Add(record);
                }
            }

            private EntityCorrelationSet(string logicalName, string primaryKey, AppSettingEvidenceState state,
                string diagnostic)
            {
                info = new EntityTypeInfo(logicalName, primaryKey, string.Empty, Enumerable.Empty<string>(),
                    Enumerable.Empty<string>(), Enumerable.Empty<string>(),
                    new AppSettingEntityMetadataEvidence(logicalName, state,
                        primaryKey, string.Empty, Enumerable.Empty<string>(), Enumerable.Empty<string>(),
                        Enumerable.Empty<string>(), diagnostic), state);
                State = state;
                Diagnostic = diagnostic;
            }

            internal AppSettingEvidenceState State { get; private set; } = AppSettingEvidenceState.Confirmed;
            internal string Diagnostic { get; private set; } = string.Empty;

            internal static EntityCorrelationSet Empty(string logicalName, string primaryKey) =>
                new EntityCorrelationSet(new EntityTypeInfo(logicalName, primaryKey, string.Empty,
                    Enumerable.Empty<string>(), Enumerable.Empty<string>(),
                    Enumerable.Empty<string>(),
                    new AppSettingEntityMetadataEvidence(logicalName, AppSettingEvidenceState.Confirmed,
                        primaryKey, string.Empty, Enumerable.Empty<string>(), Enumerable.Empty<string>(),
                        Enumerable.Empty<string>(),
                        "No IDs were requested.")), Enumerable.Empty<Entity>());

            internal static EntityCorrelationSet Unavailable(string logicalName, string primaryKey,
                AppSettingEvidenceState state, string diagnostic) => new EntityCorrelationSet(logicalName, primaryKey,
                    state, diagnostic);

            internal EntityCorrelationSet Merge(EntityCorrelationSet other)
            {
                if (other == null || other.State != AppSettingEvidenceState.Confirmed) return this;
                foreach (var pair in other.RecordsById)
                    RecordsById[pair.Key] = pair.Value;
                return this;
            }

            internal AppSettingEntityCorrelation For(Guid? objectId)
            {
                if (!objectId.HasValue || objectId.Value == Guid.Empty)
                    return new AppSettingEntityCorrelation(info.LogicalName, AppSettingEvidenceState.Unresolved,
                        info.PrimaryIdAttribute, 0, Enumerable.Empty<AppSettingRecordEvidence>(),
                        "The solutioncomponent.objectid is blank.");
                List<Entity> records;
                if (State != AppSettingEvidenceState.Confirmed)
                    return new AppSettingEntityCorrelation(info.LogicalName, State, info.PrimaryIdAttribute, 0,
                        Enumerable.Empty<AppSettingRecordEvidence>(), Diagnostic);
                if (!RecordsById.TryGetValue(objectId.Value, out records))
                    return new AppSettingEntityCorrelation(info.LogicalName, AppSettingEvidenceState.Unresolved,
                        info.PrimaryIdAttribute, 0, Enumerable.Empty<AppSettingRecordEvidence>(),
                        "No backing record matched the solutioncomponent.objectid.");
                return new AppSettingEntityCorrelation(info.LogicalName,
                    records.Count == 1 ? AppSettingEvidenceState.Confirmed : AppSettingEvidenceState.Ambiguous,
                    info.PrimaryIdAttribute, records.Count, records.Select(ToEvidence),
                    records.Count == 1 ? "Exactly one backing record matched." :
                    "Multiple backing records matched the same primary key.");
            }

            internal AppSettingRecordEvidence GetSingle(Guid? id)
            {
                if (!id.HasValue) return null;
                List<Entity> records;
                return RecordsById.TryGetValue(id.Value, out records) && records.Count == 1
                    ? ToEvidence(records[0]) : null;
            }

            private AppSettingRecordEvidence ToEvidence(Entity entity) =>
                new AppSettingRecordEvidence(entity.Id, info.Attributes.Select(name =>
                    new KeyValuePair<string, string>(name, FormatValue(name, entity.GetAttributeValue<object>(name)))));

            private static string FormatValue(string attribute, object value)
            {
                if (value == null) return "(null)";
                var sensitive = attribute.IndexOf("value", StringComparison.OrdinalIgnoreCase) >= 0 ||
                    attribute.IndexOf("secret", StringComparison.OrdinalIgnoreCase) >= 0 ||
                    attribute.IndexOf("token", StringComparison.OrdinalIgnoreCase) >= 0 ||
                    attribute.IndexOf("password", StringComparison.OrdinalIgnoreCase) >= 0 ||
                    attribute.IndexOf("connection", StringComparison.OrdinalIgnoreCase) >= 0;
                if (sensitive) return "(present; redacted)";
                var reference = value as EntityReference;
                if (reference != null) return reference.Id.ToString("D");
                var option = value as OptionSetValue;
                if (option != null) return option.Value.ToString(CultureInfo.InvariantCulture);
                return Convert.ToString(value, CultureInfo.InvariantCulture) ?? string.Empty;
            }
        }

        private static Guid? GetRelatedGuid(Entity entity, params string[] names)
        {
            if (entity == null) return null;
            foreach (var name in names)
            {
                object value;
                var key = entity.Attributes.Keys.FirstOrDefault(item =>
                    string.Equals(item, name, StringComparison.OrdinalIgnoreCase));
                if (key == null || !entity.Attributes.TryGetValue(key, out value)) continue;
                var reference = value as EntityReference;
                if (reference != null && reference.Id != Guid.Empty) return reference.Id;
                if (value is Guid && (Guid)value != Guid.Empty) return (Guid)value;
            }
            return null;
        }

        private static string GetPortableName(Entity entity)
        {
            if (entity == null) return string.Empty;
            foreach (var key in new[] { "uniquename", "name", "schemaname", "logicalname" })
            {
                var actual = entity.Attributes.Keys.FirstOrDefault(item =>
                    string.Equals(item, key, StringComparison.OrdinalIgnoreCase));
                if (actual == null) continue;
                var value = entity.GetAttributeValue<string>(actual);
                if (!string.IsNullOrWhiteSpace(value)) return value;
            }
            return string.Empty;
        }

        private static string GetString(Entity entity, string name)
        {
            if (entity == null) return string.Empty;
            var actual = entity.Attributes.Keys.FirstOrDefault(item =>
                string.Equals(item, name, StringComparison.OrdinalIgnoreCase));
            return actual == null ? string.Empty : entity.GetAttributeValue<string>(actual) ?? string.Empty;
        }
    }
}
