using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Threading;
using D365SolutionComparer.Models.ComponentDetails;
using D365SolutionComparer.Models.Membership;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;

namespace D365SolutionComparer.Services.Membership
{
    /// <summary>
    /// Registered AppSetting only. Values/defaults and arbitrary server exception text must
    /// never enter this operation's properties, diagnostics or shared read inventory.
    /// </summary>
    internal sealed class AppSettingResolutionOperation
    {
        // These allowlisted structural fields are queried ONLY after the live attribute
        // metadata confirms their exact logical names and primitive types. No label guessing.
        // Display/description/information URLs are excluded (localization/free-text risks).
        internal static readonly string[] StructuralFields =
            { "datatype", "isoverridable", "overridablelevel", "releaselevel" };
        private readonly DataverseReadContext context;
        private readonly Dictionary<Guid, AppSettingRowResult> settings = new Dictionary<Guid, AppSettingRowResult>();
        private readonly Dictionary<Guid, AppSettingRowResult> identityDefinitions =
            new Dictionary<Guid, AppSettingRowResult>();
        private readonly Dictionary<Guid, AppSettingRowResult> structuralDefinitions =
            new Dictionary<Guid, AppSettingRowResult>();
        private readonly Dictionary<Guid, AppSettingRowResult> parents = new Dictionary<Guid, AppSettingRowResult>();
        private readonly Dictionary<string, AttributeTypeCode> structuralTypes =
            new Dictionary<string, AttributeTypeCode>(StringComparer.Ordinal);
        private bool schemaAttempted;
        private bool structuralDefinitionsAttempted;
        private string structuralPreparationDiagnostic;

        internal AppSettingResolutionOperation(DataverseReadContext context) { this.context = context; }

        internal static string PortableKey(string parent, string definition) => "appsetting:v1:" +
            parent.Length.ToString(CultureInfo.InvariantCulture) + ":" + parent + ":" +
            definition.Length.ToString(CultureInfo.InvariantCulture) + ":" + definition;

        internal IReadOnlyList<ComponentIdentity> Resolve(IReadOnlyList<ComponentIdentity> candidates,
            Func<IReadOnlyList<Guid>, IReadOnlyDictionary<Guid, AppSettingRowResult>> readParents,
            CancellationToken token)
        {
            if (candidates.Count == 0) return candidates;
            ReadRows(candidates.Select(i => i.Record.ObjectId), "appsetting", "appsettingid",
                new[] { "appsettingid", "settingdefinitionid", "parentappmoduleid" }, settings, token);
            var rows = candidates.Where(i => i.Record.ObjectId.HasValue && settings.ContainsKey(i.Record.ObjectId.Value))
                .Select(i => settings[i.Record.ObjectId.Value].Row).Where(r => r != null).ToList();
            var definitionIds = rows.Select(r => Reference(r, "settingdefinitionid", "settingdefinition"));
            ReadRows(definitionIds, "settingdefinition", "settingdefinitionid",
                new[] { "settingdefinitionid", "uniquename" }, identityDefinitions, token,
                conflictsAreAmbiguous: true);
            var parentIds = rows.Select(r => Reference(r, "parentappmoduleid", "appmodule"))
                .Where(id => id.HasValue).Select(id => id.Value).Distinct().OrderBy(id => id)
                .Where(id => !parents.ContainsKey(id)).ToList();
            if (parentIds.Count > 0)
                foreach (var item in readParents(parentIds)) parents[item.Key] = item.Value;
            token.ThrowIfCancellationRequested();
            var resolved = candidates.Select(ResolveOne).ToList();
            var duplicates = new HashSet<string>(resolved.Where(i => i.Status == IdentityResolutionStatus.Resolved)
                .GroupBy(i => i.ComparisonKey, StringComparer.OrdinalIgnoreCase)
                .Where(g => g.Select(i => i.Record.SolutionComponentId).Distinct().Count() > 1)
                .Select(g => g.Key), StringComparer.OrdinalIgnoreCase);
            return resolved.Select(i => i.Status == IdentityResolutionStatus.Resolved && duplicates.Contains(i.ComparisonKey)
                ? new ComponentIdentity(i.Record, IdentityResolutionStatus.Ambiguous,
                    diagnostic: "Multiple AppSetting membership records share the same portable identity.",
                    componentTypeKey: ComponentSemanticKinds.AppSetting, registeredDefinition: i.RegisteredDefinition,
                    diagnosticEvidence: i.DiagnosticEvidence, blockerPortableIdentity: i.ComparisonKey)
                : i).ToList().AsReadOnly();
        }

        private ComponentIdentity ResolveOne(ComponentIdentity candidate)
        {
            var evidence = new List<string>
            {
                "Registered definition=" + candidate.RegisteredDefinition.Name +
                "; PrimaryEntityName=" + candidate.RegisteredDefinition.PrimaryEntityName,
                "solutioncomponentid=" + candidate.Record.SolutionComponentId.ToString("D") +
                "; objectid=" + candidate.Record.ObjectId
            };
            var setting = Find(settings, candidate.Record.ObjectId, "appsetting");
            if (setting.Row == null) return Identity(candidate, setting.Status, setting.Diagnostic, evidence);
            var definitionId = Reference(setting.Row, "settingdefinitionid", "settingdefinition");
            var parentId = Reference(setting.Row, "parentappmoduleid", "appmodule");
            evidence.Add("Exact appsetting primary-key correlation confirmed; settingdefinitionid=" +
                definitionId + "; parentappmoduleid=" + parentId);
            var definition = Find(identityDefinitions, definitionId, "settingdefinition");
            var parent = Find(parents, parentId, "appmodule");
            if (definition.Row == null) return Identity(candidate, definition.Status, definition.Diagnostic, evidence);
            if (parent.Row == null) return Identity(candidate, parent.Status, parent.Diagnostic, evidence);
            var definitionUniqueName = definition.Row.GetAttributeValue<object>("uniquename") as string;
            var parentUniqueName = parent.Row.GetAttributeValue<object>("uniquename") as string;
            if (string.IsNullOrWhiteSpace(definitionUniqueName) || string.IsNullOrWhiteSpace(parentUniqueName))
                return Identity(candidate, IdentityResolutionStatus.Unresolved,
                    "AppSetting requires a nonblank Setting Definition UniqueName and parent AppModule UniqueName.", evidence);
            evidence.Add("SettingDefinition.UniqueName=" + definitionUniqueName +
                "; ParentAppModule.UniqueName=" + parentUniqueName);
            return Identity(candidate, IdentityResolutionStatus.Resolved, "AppSetting portable identity resolved.",
                evidence, PortableKey(parentUniqueName, definitionUniqueName));
        }

        private static ComponentIdentity Identity(ComponentIdentity candidate, IdentityResolutionStatus status,
            string reason, IEnumerable<string> evidence, string key = null) =>
            new ComponentIdentity(candidate.Record, status, key, reason, ComponentSemanticKinds.AppSetting,
                registeredDefinition: candidate.RegisteredDefinition, diagnosticEvidence: evidence);

        /// <summary>
        /// Definition-only preparation. Identity resolution never depends on this metadata or
        /// structural-property query and remains valid if either read is unavailable.
        /// </summary>
        internal void PrepareDefinitions(IReadOnlyList<ComponentIdentity> identities,
            CancellationToken token)
        {
            if (structuralDefinitionsAttempted || identities.Count == 0) return;
            structuralDefinitionsAttempted = true;
            var definitionIds = identities.Select(identity =>
            {
                var setting = Find(settings, identity.Record.ObjectId, "appsetting");
                return Reference(setting.Row, "settingdefinitionid", "settingdefinition");
            }).Where(id => id.HasValue).ToList();
            if (definitionIds.Count == 0)
            {
                structuralPreparationDiagnostic =
                    "No Setting Definition references were available for structural comparison.";
                return;
            }
            EnsureSchema(token);
            if (StructuralFields.Any(field => !structuralTypes.ContainsKey(field)))
            {
                structuralPreparationDiagnostic =
                    "AppSetting structural metadata is missing, unverified or unavailable. Portable identity remains resolved.";
                return;
            }
            ReadRows(definitionIds, "settingdefinition", "settingdefinitionid",
                new[] { "settingdefinitionid" }.Concat(StructuralFields).ToArray(),
                structuralDefinitions, token);
        }

        internal ComponentDefinition Definition(ComponentIdentity identity)
        {
            var setting = Find(settings, identity.Record.ObjectId, "appsetting");
            var definition = Find(structuralDefinitions,
                Reference(setting.Row, "settingdefinitionid", "settingdefinition"),
                "settingdefinition");
            if (definition.Row == null)
                return new ComponentDefinition(identity,
                    definition.Status == IdentityResolutionStatus.Ambiguous
                        ? ComponentDefinitionReadStatus.Ambiguous
                        : ComponentDefinitionReadStatus.Unresolved,
                    diagnostic: structuralPreparationDiagnostic ?? definition.Diagnostic ??
                        "AppSetting structural comparison was not prepared. Portable identity remains resolved.",
                    diagnosticEvidence: identity.DiagnosticEvidence);
            var values = new Dictionary<string, string>(StringComparer.Ordinal);
            foreach (var field in StructuralFields)
            {
                AttributeTypeCode type;
                object value;
                if (!structuralTypes.TryGetValue(field, out type) ||
                    !definition.Row.Attributes.TryGetValue(field, out value) || value == null ||
                    !(type == AttributeTypeCode.Boolean ? value is bool :
                        type == AttributeTypeCode.Picklist ? value is OptionSetValue : value is int))
                    return new ComponentDefinition(identity, ComponentDefinitionReadStatus.Unresolved,
                        diagnostic: "AppSetting structural metadata is missing, unverified or incomplete. Values/defaults are not compared.",
                        diagnosticEvidence: identity.DiagnosticEvidence);
                values[field] = value is OptionSetValue ? ((OptionSetValue)value).Value.ToString(CultureInfo.InvariantCulture)
                    : Convert.ToString(value, CultureInfo.InvariantCulture);
            }
            return new ComponentDefinition(identity, ComponentDefinitionReadStatus.Available, values,
                "Setting Definition structural coverage only; AppSetting values and defaults are excluded.", identity.DiagnosticEvidence);
        }

        private void EnsureSchema(CancellationToken token)
        {
            if (schemaAttempted) return;
            schemaAttempted = true;
            try
            {
                var response = context.Execute(new RetrieveEntityRequest { LogicalName = "settingdefinition",
                    EntityFilters = EntityFilters.Entity | EntityFilters.Attributes, RetrieveAsIfPublished = false })
                    as RetrieveEntityResponse;
                var metadata = response?.EntityMetadata;
                if (metadata == null || !string.Equals(metadata.LogicalName, "settingdefinition", StringComparison.OrdinalIgnoreCase) ||
                    !string.Equals(metadata.PrimaryIdAttribute, "settingdefinitionid", StringComparison.OrdinalIgnoreCase) ||
                    metadata.Attributes == null) return;
                foreach (var field in StructuralFields)
                {
                    var matches = metadata.Attributes.Where(a => string.Equals(a.LogicalName, field, StringComparison.Ordinal)).ToList();
                    if (matches.Count != 1 || matches[0].IsValidForRead != true) continue;
                    var type = matches[0].AttributeType;
                    if (field == "isoverridable" ? type == AttributeTypeCode.Boolean :
                        type == AttributeTypeCode.Integer || type == AttributeTypeCode.Picklist)
                        structuralTypes[field] = type.Value;
                }
            }
            catch (OperationCanceledException) { throw; }
            catch (Exception) { token.ThrowIfCancellationRequested(); /* Never expose server fault text. */ }
        }

        private void ReadRows(IEnumerable<Guid?> ids, string table, string primaryId, string[] columns,
            IDictionary<Guid, AppSettingRowResult> cache, CancellationToken token,
            bool conflictsAreAmbiguous = false)
        {
            var pending = ids.Where(id => id.HasValue && id.Value != Guid.Empty).Select(id => id.Value)
                .Distinct().OrderBy(id => id).Where(id => !cache.ContainsKey(id)).ToList();
            if (pending.Count == 0) return;
            // Only requested, allowlisted fields enter this operation's cache, even if a
            // misbehaving service/fake returns value/defaultvalue or unrelated attributes.
            try
            {
                var records = pending.Select(id => new SolutionComponentRecord(id, 0, id)).ToList();
                var result = new BatchedDiagnosticQueryReader(context).Read(records, table, primaryId, columns,
                    table + " retrieval is incomplete.", table + " primary-key correlation is conflicting.",
                    fault => table + " retrieval failed; server details withheld.", token);
                foreach (var id in pending)
                {
                    var correlation = result.GetCorrelation(id);
                    if (correlation.Status == DiagnosticCorrelationStatus.Unique)
                    {
                        var row = correlation.Rows[0];
                        var safe = new Entity(table, id);
                        foreach (var field in columns) if (row.Contains(field)) safe[field] = row[field];
                        cache[id] = AppSettingRowResult.Unique(safe);
                    }
                    else cache[id] = new AppSettingRowResult(null,
                        correlation.Status == DiagnosticCorrelationStatus.Duplicate ||
                        conflictsAreAmbiguous && correlation.UnassociatedRows.Count > 0
                            ? IdentityResolutionStatus.Ambiguous : IdentityResolutionStatus.Unresolved,
                        correlation.Failure ?? (correlation.Status == DiagnosticCorrelationStatus.Duplicate
                            ? "Multiple " + table + " records matched the requested ID."
                            : "No " + table + " record matched the requested ID."));
                }
            }
            catch (OperationCanceledException) { throw; }
            catch (Exception)
            {
                token.ThrowIfCancellationRequested();
                foreach (var id in pending) cache[id] = new AppSettingRowResult(null, IdentityResolutionStatus.Unresolved,
                    table + " retrieval failed; server details withheld.");
            }
        }

        private static Guid? Reference(Entity row, string field, string table)
        {
            var reference = row?.GetAttributeValue<object>(field) as EntityReference;
            return reference != null && reference.Id != Guid.Empty &&
                string.Equals(reference.LogicalName, table, StringComparison.OrdinalIgnoreCase) ? reference.Id : (Guid?)null;
        }

        private static AppSettingRowResult Find(IDictionary<Guid, AppSettingRowResult> cache, Guid? id, string table)
        {
            AppSettingRowResult result;
            return id.HasValue && cache.TryGetValue(id.Value, out result) ? result :
                new AppSettingRowResult(null, IdentityResolutionStatus.Unresolved, "Missing or invalid " + table + " reference/object ID.");
        }
    }

    internal sealed class AppSettingRowResult
    {
        internal AppSettingRowResult(Entity row, IdentityResolutionStatus status, string diagnostic)
        { Row = row; Status = status; Diagnostic = diagnostic; }
        internal Entity Row { get; }
        internal IdentityResolutionStatus Status { get; }
        internal string Diagnostic { get; }
        internal static AppSettingRowResult Unique(Entity row) => new AppSettingRowResult(row, IdentityResolutionStatus.Resolved, null);
    }
}
