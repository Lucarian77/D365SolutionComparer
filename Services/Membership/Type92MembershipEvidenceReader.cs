using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Threading;
using D365SolutionComparer.Models.Membership;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace D365SolutionComparer.Services.Membership
{
    /// <summary>Batched read-only Type 92 evidence used by production membership resolution.</summary>
    internal sealed class Type92MembershipEvidenceReader
    {
        internal const int BatchSize = 200;
        internal const int ComponentType = 92;
        private static readonly string[] StepColumns = { "sdkmessageprocessingstepid", "sdkmessageprocessingstepidunique",
            "name", "eventhandler", "sdkmessageid", "sdkmessagefilterid", "stage", "mode",
            "supporteddeployment", "rank", "filteringattributes", "configuration", "statecode",
            "statuscode", "ismanaged", "componentstate" };
        private static readonly string[] TypeColumns = { "plugintypeid", "plugintypeidunique",
            "plugintypeexportkey", "typename", "name", "isworkflowactivity", "pluginassemblyid",
            "ismanaged", "componentstate" };
        private static readonly string[] AssemblyColumns = { "pluginassemblyid", "pluginassemblyidunique",
            "name", "publickeytoken", "culture", "version", "isolationmode", "sourcetype",
            "ismanaged", "componentstate" };
        private static readonly string[] MessageColumns = { "sdkmessageid", "name" };
        private static readonly string[] FilterColumns = { "sdkmessagefilterid", "sdkmessagefilteridunique",
            "name", "sdkmessageid", "primaryobjecttypecode", "secondaryobjecttypecode", "availability" };

        internal Type92EnvironmentEvidence Capture(IOrganizationService service, MembershipSnapshot snapshot,
            string solutionVersion, CancellationToken cancellationToken, Action<string> progress = null,
            DataverseComponentMetadataCache metadataCache = null)
        {
            if (service == null) throw new ArgumentNullException(nameof(service));
            if (snapshot == null || snapshot.State != MembershipSnapshotState.Complete || snapshot.Solution == null)
                throw new InvalidOperationException("Type 92 evidence requires a completed solution snapshot.");
            cancellationToken.ThrowIfCancellationRequested();
            var raw = snapshot.Components.Select(item => item.Record)
                .Where(item => item.ComponentType == ComponentType)
                .OrderBy(item => item.SolutionComponentId).ToList();
            return CaptureCore(service, snapshot, solutionVersion, raw, cancellationToken,
                progress, metadataCache);
        }

        internal Type92EnvironmentEvidence CaptureSingle(IOrganizationService service,
            SolutionComponentRecord record, CancellationToken cancellationToken,
            DataverseComponentMetadataCache metadataCache = null)
        {
            if (service == null) throw new ArgumentNullException(nameof(service));
            if (record == null) throw new ArgumentNullException(nameof(record));
            if (record.ComponentType != ComponentType)
                throw new ArgumentException("The component must be an SDK Message Processing Step.", nameof(record));
            cancellationToken.ThrowIfCancellationRequested();
            return CaptureCore(service, null, string.Empty, new[] { record }, cancellationToken,
                null, metadataCache);
        }

        private static Type92EnvironmentEvidence CaptureCore(IOrganizationService service,
            MembershipSnapshot snapshot, string solutionVersion, IReadOnlyList<SolutionComponentRecord> raw,
            CancellationToken cancellationToken, Action<string> progress,
            DataverseComponentMetadataCache metadataCache)
        {
            var ledger = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            progress?.Invoke("Reading SDK Message Processing Steps...");
            var steps = Read(service, "sdkmessageprocessingstep", "sdkmessageprocessingstepid",
                StepColumns, raw.Select(item => item.ObjectId), ledger, cancellationToken);
            var uniqueSteps = steps.UniqueRows.ToList();
            var handlers = uniqueSteps.Select(item => item.GetAttributeValue<EntityReference>("eventhandler"))
                .Where(item => item != null && item.Id != Guid.Empty &&
                    string.Equals(item.LogicalName, "plugintype", StringComparison.OrdinalIgnoreCase))
                .Select(item => (Guid?)item.Id);
            progress?.Invoke("Reading referenced Plug-in Types...");
            var types = Read(service, "plugintype", "plugintypeid", TypeColumns,
                handlers, ledger, cancellationToken);
            var typeRows = types.UniqueRows.ToList();
            progress?.Invoke("Reading parent Plug-in Assemblies...");
            var assemblyIds = typeRows.Select(item => Id(item, "pluginassemblyid"))
                .Where(item => item.HasValue && item.Value != Guid.Empty)
                .Select(item => item.Value).Distinct().ToList();
            var cachedAssemblies = new List<Entity>();
            if (metadataCache != null)
                foreach (var assemblyId in assemblyIds)
                {
                    Entity cached;
                    if (metadataCache.TryGetEntityRow("pluginassembly", assemblyId, out cached))
                        cachedAssemblies.Add(cached);
                }
            var assemblies = Read(service, "pluginassembly", "pluginassemblyid", AssemblyColumns,
                assemblyIds.Select(item => (Guid?)item), ledger, cancellationToken, cachedAssemblies);
            progress?.Invoke("Reading SDK Messages...");
            var messages = Read(service, "sdkmessage", "sdkmessageid", MessageColumns,
                uniqueSteps.Select(item => Id(item, "sdkmessageid")), ledger, cancellationToken);
            progress?.Invoke("Reading SDK Message Filters...");
            var filters = Read(service, "sdkmessagefilter", "sdkmessagefilterid", FilterColumns,
                uniqueSteps.Select(item => Id(item, "sdkmessagefilterid")), ledger, cancellationToken);

            var evidence = new List<Type92StepEvidence>();
            foreach (var rawGroup in raw.Where(item => item.ObjectId.HasValue && item.ObjectId.Value != Guid.Empty)
                .GroupBy(item => item.ObjectId.Value).OrderBy(group => group.Key))
            {
                cancellationToken.ThrowIfCancellationRequested();
                var step = steps.Get(rawGroup.Key);
                evidence.Add(BuildStep(step, rawGroup.Count(), types, assemblies, messages, filters));
            }
            foreach (var group in evidence.Where(item => item.Candidate != null)
                .GroupBy(item => item.Candidate))
            {
                if (group.Select(item => item.StepId).Distinct().Count() < 2) continue;
                foreach (var item in group) item.MarkAmbiguous("Different backing steps share this candidate tuple.");
            }
            cancellationToken.ThrowIfCancellationRequested();
            return new Type92EnvironmentEvidence(snapshot, solutionVersion, raw, evidence, ledger);
        }

        internal static string CanonicalFilteringAttributes(string value) => string.Join(",",
            (value ?? string.Empty).Split(',').Select(item => item.Trim().ToLowerInvariant())
                .Where(item => item.Length != 0).Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(item => item, StringComparer.OrdinalIgnoreCase));

        private static Type92StepEvidence BuildStep(Type92Correlation step, int membershipCount,
            Type92BatchResult types, Type92BatchResult assemblies, Type92BatchResult messages,
            Type92BatchResult filters)
        {
            var result = new Type92StepEvidence(step.Id, membershipCount, step.Status, step.Diagnostic);
            if (step.Status != "Unique") return result;
            var row = step.Row;
            result.StepId = row.Id;
            result.StepUniqueId = Id(row, "sdkmessageprocessingstepidunique");
            result.Name = row.GetAttributeValue<string>("name");
            result.Handler = row.GetAttributeValue<EntityReference>("eventhandler");
            result.MessageId = Id(row, "sdkmessageid");
            result.FilterId = Id(row, "sdkmessagefilterid");
            result.Stage = Number(row, "stage");
            result.Mode = Number(row, "mode");
            result.Deployment = Number(row, "supporteddeployment");
            result.Rank = Number(row, "rank");
            result.FilteringOriginal = row.GetAttributeValue<string>("filteringattributes") ?? string.Empty;
            result.FilteringCanonical = CanonicalFilteringAttributes(result.FilteringOriginal);
            var configuration = row.GetAttributeValue<string>("configuration");
            result.ConfigurationPresent = !string.IsNullOrWhiteSpace(configuration);
            result.ConfigurationLength = configuration?.Length ?? 0;
            result.State = Number(row, "statecode");
            result.Status = Number(row, "statuscode");
            result.IsManaged = row.GetAttributeValue<bool?>("ismanaged");
            result.ComponentState = Number(row, "componentstate");
            result.HandlerCategory = result.Handler == null || result.Handler.Id == Guid.Empty ? "missing" :
                string.IsNullOrWhiteSpace(result.Handler.LogicalName) ? "unknown" :
                string.Equals(result.Handler.LogicalName, "plugintype", StringComparison.OrdinalIgnoreCase) ? "plugintype" :
                string.Equals(result.Handler.LogicalName, "serviceendpoint", StringComparison.OrdinalIgnoreCase) ? "serviceendpoint" : "other";
            if (result.HandlerCategory != "plugintype")
            {
                result.CandidateStatus = "IncompleteHandler";
                return result;
            }
            var type = types.Get(result.Handler.Id);
            if (type.Status != "Unique") { result.CandidateStatus = "IncompleteHandler"; result.Diagnostic = type.Diagnostic; return result; }
            result.PluginTypeId = type.Row.Id;
            result.PluginTypeUniqueId = Id(type.Row, "plugintypeidunique");
            result.ExportKey = type.Row.GetAttributeValue<string>("plugintypeexportkey")?.Trim();
            result.TypeName = type.Row.GetAttributeValue<string>("typename")?.Trim();
            result.PluginTypeName = type.Row.GetAttributeValue<string>("name");
            result.IsWorkflowActivity = type.Row.GetAttributeValue<bool?>("isworkflowactivity");
            result.AssemblyId = Id(type.Row, "pluginassemblyid");
            if (result.AssemblyId.HasValue)
            {
                var assembly = assemblies.Get(result.AssemblyId.Value);
                if (assembly.Status == "Unique")
                {
                    var a = assembly.Row;
                    result.AssemblyVersion = a.GetAttributeValue<string>("version");
                    var name = a.GetAttributeValue<string>("name")?.Trim();
                    var token = a.GetAttributeValue<string>("publickeytoken")?.Trim();
                    var culture = a.GetAttributeValue<string>("culture")?.Trim();
                    if (!string.IsNullOrWhiteSpace(name) && !string.IsNullOrWhiteSpace(token) &&
                        !string.IsNullOrWhiteSpace(culture))
                        result.AssemblyIdentity = DataverseComponentIdentityResolver.PluginAssemblyPortableKey(
                            name, token, culture);
                }
            }
            if (string.IsNullOrWhiteSpace(result.TypeName) ||
                string.IsNullOrWhiteSpace(result.AssemblyIdentity))
            { result.CandidateStatus = "IncompleteHandler"; return result; }
            result.HandlerSemanticIdentity = result.AssemblyIdentity + " + " + result.TypeName;
            if (!result.MessageId.HasValue)
            { result.CandidateStatus = "MissingMessage"; return result; }
            var message = messages.Get(result.MessageId.Value);
            if (message.Status != "Unique")
            { result.CandidateStatus = "MissingMessage"; result.Diagnostic = message.Diagnostic; return result; }
            result.MessageName = message.Row.GetAttributeValue<string>("name")?.Trim();
            if (string.IsNullOrWhiteSpace(result.MessageName))
            { result.CandidateStatus = "BlankMessageName"; return result; }
            if (result.FilterId.HasValue)
            {
                var filter = filters.Get(result.FilterId.Value);
                if (filter.Status != "Unique")
                { result.CandidateStatus = "UnresolvedFilter"; result.Diagnostic = filter.Diagnostic; return result; }
                result.FilterName = filter.Row.GetAttributeValue<string>("name");
                result.PrimaryScope = filter.Row.GetAttributeValue<string>("primaryobjecttypecode")?.Trim();
                result.SecondaryScope = filter.Row.GetAttributeValue<string>("secondaryobjecttypecode")?.Trim();
                result.FilterMessageId = Id(filter.Row, "sdkmessageid");
                result.FilterUniqueId = Id(filter.Row, "sdkmessagefilteridunique");
                result.FilterAvailability = Number(filter.Row, "availability");
                if (!result.FilterMessageId.HasValue || result.FilterMessageId != result.MessageId)
                { result.CandidateStatus = "ConflictingFilter"; return result; }
            }
            else result.FilterStatus = "NoFilter";
            var primary = NormalizeEntityScope(result.PrimaryScope);
            var secondary = NormalizeEntityScope(result.SecondaryScope);
            if (result.FilterStatus == "NoFilter" || (primary.Length == 0 && secondary.Length == 0))
                result.SemanticScope = "Global/Unbound";
            else if (primary.Length == 0)
            { result.CandidateStatus = "IncompleteScope"; return result; }
            else result.SemanticScope = "EntityBound";
            if (!result.Stage.HasValue || !result.Mode.HasValue || !result.Deployment.HasValue)
            { result.CandidateStatus = "IncompleteOptions"; return result; }
            result.Candidate = new Type92Candidate(result.AssemblyIdentity, result.TypeName,
                result.MessageName, result.SemanticScope, primary, secondary,
                result.Stage.Value, result.Mode.Value, result.Deployment.Value);
            result.CandidateStatus = "Candidate";
            return result;
        }

        private static string NormalizeEntityScope(string value)
        {
            var normalized = value?.Trim() ?? string.Empty;
            return string.Equals(normalized, "none", StringComparison.OrdinalIgnoreCase) ?
                string.Empty : normalized;
        }

        private static Type92BatchResult Read(IOrganizationService service, string entityName, string primaryId,
            string[] columns, IEnumerable<Guid?> ids, IDictionary<string, int> ledger,
            CancellationToken cancellationToken, IEnumerable<Entity> cachedRows = null)
        {
            var requested = ids.Where(item => item.HasValue && item.Value != Guid.Empty)
                .Select(item => item.Value).Distinct().OrderBy(item => item).ToList();
            var rows = requested.ToDictionary(item => item, item => new List<Entity>());
            var failures = new Dictionary<Guid, string>();
            foreach (var cached in cachedRows ?? Enumerable.Empty<Entity>())
            {
                cancellationToken.ThrowIfCancellationRequested();
                var key = Id(cached, primaryId);
                if (!string.Equals(cached.LogicalName, entityName, StringComparison.OrdinalIgnoreCase) ||
                    !key.HasValue || !rows.ContainsKey(key.Value) ||
                    cached.Id != Guid.Empty && cached.Id != key.Value)
                    throw new InvalidOperationException("The operation-scoped parent cache contains conflicting primary-key data.");
                rows[key.Value].Add(cached);
            }
            var uncached = requested.Where(item => rows[item].Count == 0).ToList();
            for (int offset = 0; offset < uncached.Count; offset += BatchSize)
            {
                cancellationToken.ThrowIfCancellationRequested();
                var batch = uncached.Skip(offset).Take(BatchSize).ToList();
                var query = new QueryExpression(entityName) { ColumnSet = new ColumnSet(columns) };
                query.Criteria.AddCondition(new ConditionExpression(primaryId, ConditionOperator.In,
                    batch.Select(item => (object)item).ToArray()));
                query.AddOrder(primaryId, OrderType.Ascending);
                int previous;
                ledger.TryGetValue(entityName, out previous);
                ledger[entityName] = previous + 1;
                var response = service.RetrieveMultiple(query);
                cancellationToken.ThrowIfCancellationRequested();
                if (response == null) throw new InvalidOperationException(entityName + " returned no response.");
                if (response.MoreRecords)
                    foreach (var id in batch) failures[id] = "IncompleteBatch";
                foreach (var row in response.Entities)
                {
                    var key = Id(row, primaryId);
                    if (!string.Equals(row.LogicalName, entityName, StringComparison.OrdinalIgnoreCase) ||
                        !key.HasValue || !batch.Contains(key.Value))
                    {
                        foreach (var id in batch) failures[id] = "ConflictingPrimaryKey";
                        continue;
                    }
                    if (row.Id != Guid.Empty && row.Id != key.Value)
                    { failures[key.Value] = "ConflictingPrimaryKey"; continue; }
                    rows[key.Value].Add(row);
                }
            }
            return new Type92BatchResult(rows, failures);
        }

        private static Guid? Id(Entity row, string name)
        {
            if (row == null || !row.Attributes.ContainsKey(name)) return null;
            var value = row[name];
            if (value is EntityReference) return ((EntityReference)value).Id;
            if (value is Guid) return (Guid)value;
            return null;
        }

        private static int? Number(Entity row, string name)
        {
            if (row == null || !row.Attributes.ContainsKey(name)) return null;
            var value = row[name];
            if (value is OptionSetValue) return ((OptionSetValue)value).Value;
            if (value is int) return (int)value;
            return null;
        }
    }

    internal sealed class Type92BatchResult
    {
        private readonly IDictionary<Guid, List<Entity>> rows;
        private readonly IDictionary<Guid, string> failures;
        internal Type92BatchResult(IDictionary<Guid, List<Entity>> rows, IDictionary<Guid, string> failures)
        { this.rows = rows; this.failures = failures; }
        internal IEnumerable<Entity> UniqueRows => rows.Keys.Select(Get).Where(item => item.Status == "Unique")
            .Select(item => item.Row);
        internal Type92Correlation Get(Guid id)
        {
            List<Entity> found;
            if (!rows.TryGetValue(id, out found)) return new Type92Correlation(id, "Missing", null, "ID was not requested.");
            string failure;
            if (failures.TryGetValue(id, out failure)) return new Type92Correlation(id, failure, null, failure);
            return found.Count == 0 ? new Type92Correlation(id, "Missing", null, "No backing row.") :
                found.Count == 1 ? new Type92Correlation(id, "Unique", found[0], "") :
                new Type92Correlation(id, "DuplicateReturnedRow", null, "Multiple backing rows.");
        }
    }

    internal sealed class Type92Correlation
    {
        internal Type92Correlation(Guid id, string status, Entity row, string diagnostic)
        { Id = id; Status = status; Row = row; Diagnostic = diagnostic; }
        internal Guid Id { get; }
        internal string Status { get; }
        internal Entity Row { get; }
        internal string Diagnostic { get; }
    }

    internal sealed class Type92Candidate : IEquatable<Type92Candidate>
    {
        internal Type92Candidate(string assembly, string typeName, string message, string semanticScope,
            string primary, string secondary, int stage, int mode, int deployment)
        {
            Assembly = assembly.Trim(); TypeName = typeName.Trim(); Message = message.Trim();
            SemanticScope = semanticScope;
            Primary = primary?.Trim() ?? ""; Secondary = secondary?.Trim() ?? "";
            Stage = stage; Mode = mode; Deployment = deployment;
        }
        internal string Assembly { get; }
        internal string TypeName { get; }
        internal string Message { get; }
        internal string SemanticScope { get; }
        internal string Primary { get; }
        internal string Secondary { get; }
        internal int Stage { get; }
        internal int Mode { get; }
        internal int Deployment { get; }
        internal string PortableKey => "sdkmessageprocessingstep:v1:" +
            Frame(Assembly) + Frame(TypeName) + Frame(Message) + Frame(SemanticScope) +
            Frame(Primary) + Frame(Secondary) +
            Stage.ToString(CultureInfo.InvariantCulture) + ":" +
            Mode.ToString(CultureInfo.InvariantCulture) + ":" +
            Deployment.ToString(CultureInfo.InvariantCulture);
        private static string Frame(string value) => value.Length.ToString(CultureInfo.InvariantCulture) +
            ":" + value + ":";
        public bool Equals(Type92Candidate other) => other != null &&
            Same(Assembly, other.Assembly) && Same(TypeName, other.TypeName) &&
            Same(Message, other.Message) && Same(SemanticScope, other.SemanticScope) &&
            Same(Primary, other.Primary) &&
            Same(Secondary, other.Secondary) && Stage == other.Stage && Mode == other.Mode &&
            Deployment == other.Deployment;
        public override bool Equals(object obj) => Equals(obj as Type92Candidate);
        public override int GetHashCode()
        {
            unchecked
            {
                var hash = 17;
                foreach (var value in new[] { Assembly, TypeName, Message, SemanticScope, Primary, Secondary })
                    hash = hash * 31 + StringComparer.OrdinalIgnoreCase.GetHashCode(value);
                return (((hash * 31 + Stage) * 31 + Mode) * 31 + Deployment);
            }
        }
        private static bool Same(string a, string b) => string.Equals(a, b, StringComparison.OrdinalIgnoreCase);
        public override string ToString() => "handlerAssembly=" + Assembly + "; handlerType=" + TypeName +
            "; message=" + Message + "; semanticScope=" + SemanticScope + "/" + Primary +
            "/" + Secondary + "; stage=" + Stage +
            "; mode=" + Mode + "; deployment=" + Deployment;
    }

    internal sealed class Type92StepEvidence
    {
        internal Type92StepEvidence(Guid objectId, int membershipCount, string correlation, string diagnostic)
        { ObjectId = objectId; MembershipCount = membershipCount; Correlation = correlation;
            Diagnostic = diagnostic; CandidateStatus = correlation == "Unique" ? "Incomplete" : correlation; }
        internal Guid ObjectId { get; }
        internal int MembershipCount { get; }
        internal string Correlation { get; }
        internal string Diagnostic { get; set; }
        internal Guid StepId { get; set; }
        internal Guid? StepUniqueId { get; set; }
        internal string Name { get; set; }
        internal EntityReference Handler { get; set; }
        internal string HandlerCategory { get; set; }
        internal Guid PluginTypeId { get; set; }
        internal Guid? PluginTypeUniqueId { get; set; }
        internal string ExportKey { get; set; }
        internal string TypeName { get; set; }
        internal string PluginTypeName { get; set; }
        internal bool? IsWorkflowActivity { get; set; }
        internal Guid? AssemblyId { get; set; }
        internal string AssemblyIdentity { get; set; }
        internal string HandlerSemanticIdentity { get; set; }
        internal string AssemblyVersion { get; set; }
        internal Guid? MessageId { get; set; }
        internal string MessageName { get; set; }
        internal Guid? FilterId { get; set; }
        internal Guid? FilterUniqueId { get; set; }
        internal Guid? FilterMessageId { get; set; }
        internal string FilterName { get; set; }
        internal string FilterStatus { get; set; } = "Filter";
        internal string PrimaryScope { get; set; }
        internal string SecondaryScope { get; set; }
        internal string SemanticScope { get; set; }
        internal int? FilterAvailability { get; set; }
        internal int? Stage { get; set; }
        internal int? Mode { get; set; }
        internal int? Deployment { get; set; }
        internal int? Rank { get; set; }
        internal string FilteringOriginal { get; set; }
        internal string FilteringCanonical { get; set; }
        internal bool ConfigurationPresent { get; set; }
        internal int ConfigurationLength { get; set; }
        internal int? State { get; set; }
        internal int? Status { get; set; }
        internal bool? IsManaged { get; set; }
        internal int? ComponentState { get; set; }
        internal Type92Candidate Candidate { get; set; }
        internal string CandidateStatus { get; set; }
        internal void MarkAmbiguous(string reason) { CandidateStatus = "AmbiguousCandidate"; Diagnostic = reason; }
    }

    internal sealed class Type92EnvironmentEvidence
    {
        internal Type92EnvironmentEvidence(MembershipSnapshot snapshot, string solutionVersion,
            IReadOnlyList<SolutionComponentRecord> raw, IReadOnlyList<Type92StepEvidence> steps,
            IDictionary<string, int> ledger)
        {
            Snapshot = snapshot; SolutionVersion = solutionVersion ?? ""; Raw = raw; Steps = steps;
            Ledger = new Dictionary<string, int>(ledger, StringComparer.OrdinalIgnoreCase);
        }
        internal MembershipSnapshot Snapshot { get; }
        internal string SolutionVersion { get; }
        internal IReadOnlyList<SolutionComponentRecord> Raw { get; }
        internal IReadOnlyList<Type92StepEvidence> Steps { get; }
        internal IReadOnlyDictionary<string, int> Ledger { get; }
        internal int Count(string entity) { int value; return Ledger.TryGetValue(entity, out value) ? value : 0; }
    }

}
