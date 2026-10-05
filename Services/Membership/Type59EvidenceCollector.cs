#if DEBUG
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.ServiceModel;
using System.Text;
using System.Threading;
using System.Xml;
using D365SolutionComparer.Models.Identity;
using D365SolutionComparer.Models.Membership;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata.Query;
using Microsoft.Xrm.Sdk.Query;

namespace D365SolutionComparer.Services.Membership
{
    /// <summary>Temporary evidence only. Never supplies identities to the membership comparer.</summary>
    internal sealed class Type59EvidenceCollector
    {
        internal const string EntityName = "savedqueryvisualization";
        internal const string PrimaryId = "savedqueryvisualizationid";
        internal const int BatchSize = 200;
        internal static readonly string[] LightColumns = { PrimaryId, "savedqueryvisualizationidunique",
            "primaryentitytypecode", "name", "type", "charttype", "ismanaged" };
        internal static readonly string[] DetailColumns = LightColumns.Concat(new[] {
            "componentstate", "datadescription", "presentationdescription", "isdefault" }).ToArray();
        internal static readonly string[] MembershipColumns = { "solutioncomponentid", "componenttype",
            "objectid", "solutionid", "rootsolutioncomponentid", "rootcomponentbehavior", "ismetadata" };

        internal Type59EvidenceReport Capture(IOrganizationService sourceService, MembershipSnapshot source,
            string sourceVersion, IOrganizationService targetService, MembershipSnapshot target,
            string targetVersion, CancellationToken token, bool discover = false, Action<string> progress = null)
        {
            RequireSnapshot(source); RequireSnapshot(target);
            if (!StringComparer.OrdinalIgnoreCase.Equals(source.SolutionUniqueName, target.SolutionUniqueName))
                throw new ArgumentException("Source and Target must describe the same solution unique name.");
            var left = new Session(sourceService, source.Environment, token, progress);
            var right = new Session(targetService, target.Environment, token, progress);
            var report = new Type59EvidenceReport { Source = left.Evidence, Target = right.Evidence, Discovery = discover };
            if (discover)
            {
                left.Inventory(); right.Inventory();
                var shared = left.Evidence.Raw.Select(item => item.Solution).Intersect(
                    right.Evidence.Raw.Select(item => item.Solution), StringComparer.OrdinalIgnoreCase)
                    .OrderBy(item => item, StringComparer.OrdinalIgnoreCase).ToList();
                report.SharedSolutions.AddRange(shared);
                // One deduplicated detail inventory per side, including one-sided solutions.
                // Repeated solution references reuse the same backing row and XML hashes.
                left.Load(left.Evidence.Raw, true);
                right.Load(right.Evidence.Raw, true);
                report.Analyze();
                var selected = report.Pairs.Where(item => item.UniqueA && item.DifferentIds)
                    .OrderByDescending(item => item.UnmanagedToManaged)
                    .ThenBy(item => report.Population(item.Solution))
                    .ThenBy(item => item.Solution, StringComparer.OrdinalIgnoreCase)
                    .ThenBy(item => item.Source.CandidateA, StringComparer.OrdinalIgnoreCase)
                    .GroupBy(item => item.Source.ObjectId.ToString("D") + ":" + item.Target.ObjectId.ToString("D"))
                    .Select(item => item.First()).Take(2).ToList();
                report.SelectedPairs.AddRange(selected);
                left.LoadDetails(selected.Select(item => item.Source.ObjectId));
                right.LoadDetails(selected.Select(item => item.Target.ObjectId));
            }
            else
            {
                left.UseSnapshot(source, sourceVersion); right.UseSnapshot(target, targetVersion);
                report.SharedSolutions.Add(source.SolutionUniqueName);
                left.Load(left.Evidence.Raw, true); right.Load(right.Evidence.Raw, true);
                report.Analyze();
                report.SelectedPairs.AddRange(report.Pairs.Where(item => item.UniqueA));
            }
            token.ThrowIfCancellationRequested();
            return report;
        }

        private static void RequireSnapshot(MembershipSnapshot snapshot)
        {
            if (snapshot == null || snapshot.State != MembershipSnapshotState.Complete)
                throw new ArgumentException("Completed membership snapshots are required; incomplete evidence cannot prove absence.");
        }

        internal static string Text(Entity row, string column) => row != null &&
            row.Attributes.TryGetValue(column, out var value) && value is string ? ((string)value).Trim() : null;
        internal static int? Number(Entity row, string column)
        {
            if (row == null || !row.Attributes.TryGetValue(column, out var value)) return null;
            if (value is OptionSetValue) return ((OptionSetValue)value).Value;
            return value is int ? (int?)value : null;
        }
        internal static Guid? Id(Entity row, string column)
        {
            if (row == null || !row.Attributes.TryGetValue(column, out var value)) return null;
            return value is Guid && (Guid)value != Guid.Empty ? (Guid?)value : null;
        }
        private static Guid? SolutionId(Entity row)
        {
            if (row.Attributes.TryGetValue("solutionid", out var value) && value is EntityReference)
            {
                var reference = (EntityReference)value;
                return StringComparer.OrdinalIgnoreCase.Equals(reference.LogicalName, "solution") && reference.Id != Guid.Empty
                    ? (Guid?)reference.Id : null;
            }
            return Id(row, "solutionid");
        }
        internal static bool? Boolean(Entity row, string column) => row != null &&
            row.Attributes.TryGetValue(column, out var value) && value is bool ? (bool?)value : null;
        internal static string Frame(params string[] values) => string.Concat(values.Select(value =>
            value.Length.ToString(CultureInfo.InvariantCulture) + ":" + value + ":"));
        internal static string Safe(string value) => (value ?? "(not supplied)").Replace("\r", "\\r")
            .Replace("\n", "\\n").Replace("\t", "\\t");
        private static bool LogicalScope(string value) => !string.IsNullOrWhiteSpace(value) &&
            !StringComparer.OrdinalIgnoreCase.Equals(value, "none") &&
            (char.IsLetter(value[0]) || value[0] == '_') &&
            value.All(character => char.IsLetterOrDigit(character) || character == '_');
        private static IEnumerable<List<T>> Batches<T>(IEnumerable<T> values)
        {
            var list = values.ToList();
            for (int i = 0; i < list.Count; i += BatchSize) yield return list.Skip(i).Take(BatchSize).ToList();
        }

        private sealed class Session
        {
            private readonly IOrganizationService service;
            private readonly CancellationToken token;
            private readonly Action<string> progress;
            private readonly Dictionary<int, string> scopes = new Dictionary<int, string>();
            internal readonly Type59EnvironmentEvidence Evidence;
            internal Session(IOrganizationService service, EnvironmentIdentity environment,
                CancellationToken token, Action<string> progress)
            {
                this.service = service ?? throw new ArgumentNullException(nameof(service));
                this.token = token; this.progress = progress;
                Evidence = new Type59EnvironmentEvidence { Environment = environment, CapturedUtc = DateTimeOffset.UtcNow };
            }
            private EntityCollection Query(QueryExpression query)
            {
                token.ThrowIfCancellationRequested();
                Evidence.Requests[query.EntityName] = Evidence.Count(query.EntityName) + 1;
                progress?.Invoke(Evidence.Environment.DisplayName + ": reading " + query.EntityName);
                var rows = service.RetrieveMultiple(query);
                token.ThrowIfCancellationRequested();
                return rows ?? throw new InvalidOperationException("Dataverse returned no query response.");
            }
            internal void UseSnapshot(MembershipSnapshot snapshot, string version)
            {
                Evidence.SnapshotUtc = snapshot.CapturedAt;
                Evidence.Raw.AddRange(snapshot.Components.Where(item => item.Record.ComponentType == 59)
                    .Select(item => new Type59RawEvidence { Record = item.Record,
                        Solution = snapshot.SolutionUniqueName, Version = version }));
            }
            internal void Inventory()
            {
                var query = new QueryExpression("solutioncomponent") { ColumnSet = new ColumnSet(MembershipColumns),
                    PageInfo = new PagingInfo { Count = 5000, PageNumber = 1 } };
                query.Criteria.AddCondition("componenttype", ConditionOperator.Equal, 59);
                query.AddOrder("solutioncomponentid", OrderType.Ascending);
                var link = query.AddLink("solution", "solutionid", "solutionid", JoinOperator.Inner);
                link.EntityAlias = "evidenceSolution";
                link.Columns = new ColumnSet("uniquename", "version");
                var seen = new HashSet<Guid>(); var cookies = new HashSet<string>(StringComparer.Ordinal);
                var versions = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                while (true)
                {
                    var page = Query(query);
                    foreach (var row in page.Entities)
                    {
                        var id = Id(row, "solutioncomponentid");
                        var solution = Alias(row, "uniquename"); var version = Alias(row, "version");
                        if (!StringComparer.OrdinalIgnoreCase.Equals(row.LogicalName, "solutioncomponent") ||
                            !id.HasValue || (row.Id != Guid.Empty && row.Id != id) || !seen.Add(id.Value) ||
                            Number(row, "componenttype") != 59 || !SolutionId(row).HasValue ||
                            string.IsNullOrWhiteSpace(solution) || string.IsNullOrWhiteSpace(version))
                            throw new InvalidOperationException("Type 59 discovery returned incomplete/conflicting membership or solution evidence; no partial inventory is usable.");
                        if (versions.TryGetValue(solution, out var prior) && prior != version)
                            throw new InvalidOperationException("Type 59 discovery returned conflicting solution versions.");
                        versions[solution] = version;
                        Evidence.Raw.Add(new Type59RawEvidence { Solution = solution, Version = version,
                            Record = new SolutionComponentRecord(id.Value, 59, Id(row, "objectid"),
                                Number(row, "rootcomponentbehavior"), Id(row, "rootsolutioncomponentid"), Boolean(row, "ismetadata")) });
                    }
                    if (!page.MoreRecords) break;
                    if (page.Entities.Count == 0 || string.IsNullOrWhiteSpace(page.PagingCookie) || !cookies.Add(page.PagingCookie))
                        throw new InvalidOperationException("Type 59 membership discovery paging is incomplete; no partial inventory is usable.");
                    query.PageInfo.PageNumber++; query.PageInfo.PagingCookie = page.PagingCookie;
                }
            }
            private static string Alias(Entity row, string name) =>
                row.Attributes.TryGetValue("evidenceSolution." + name, out var value) && value is AliasedValue &&
                ((AliasedValue)value).Value is string ? ((string)((AliasedValue)value).Value).Trim() : null;

            internal void Load(IEnumerable<Type59RawEvidence> memberships, bool detail)
            {
                var ids = memberships.Select(item => item.Record.ObjectId).Where(item => item.HasValue && item != Guid.Empty)
                    .Select(item => item.Value).Distinct().OrderBy(item => item).ToList();
                foreach (var chart in Read(ids, detail)) Evidence.Charts[chart.ObjectId] = chart;
                ResolveScopes(Evidence.Charts.Values);
                foreach (var chart in Evidence.Charts.Values) chart.ConstructCandidates();
                MarkDuplicates(Evidence.Charts.Values, false); MarkDuplicates(Evidence.Charts.Values, true);
            }
            internal void LoadDetails(IEnumerable<Guid> ids)
            {
                var list = ids.Distinct().OrderBy(item => item).Where(item =>
                    !Evidence.Details.ContainsKey(item)).ToList();
                foreach (var chart in Read(list, true)) Evidence.Details[chart.ObjectId] = chart;
            }
            private IEnumerable<Type59ChartEvidence> Read(IEnumerable<Guid> ids, bool detail)
            {
                foreach (var batch in Batches(ids))
                {
                    var results = batch.ToDictionary(item => item, item => new List<Entity>());
                    string failure = null;
                    try
                    {
                        var query = new QueryExpression(EntityName) { ColumnSet = new ColumnSet(detail ? DetailColumns : LightColumns) };
                        query.Criteria.AddCondition(new ConditionExpression(PrimaryId, ConditionOperator.In,
                            batch.Select(item => (object)item).ToArray()));
                        var rows = Query(query);
                        if (rows.MoreRecords) failure = "IncompleteBatch: unexpected backing paging";
                        foreach (var row in rows.Entities)
                        {
                            var id = Id(row, PrimaryId);
                            if (!StringComparer.OrdinalIgnoreCase.Equals(row.LogicalName, EntityName) || !id.HasValue ||
                                !results.ContainsKey(id.Value) || (row.Id != Guid.Empty && row.Id != id))
                                failure = "ConflictingPrimaryKey or unexpected backing entity/ID";
                            else results[id.Value].Add(row);
                        }
                    }
                    catch (OperationCanceledException) { throw; }
                    catch (FaultException) { token.ThrowIfCancellationRequested(); failure = "Faulted backing retrieval; server details withheld"; }
                    catch (InvalidOperationException) { token.ThrowIfCancellationRequested(); failure = "Incomplete backing response"; }
                    foreach (var id in batch)
                    {
                        var chart = new Type59ChartEvidence { ObjectId = id, Rows = results[id], Detail = detail,
                            Correlation = failure != null ? "Incomplete" : results[id].Count == 0 ? "Missing" :
                                results[id].Count == 1 ? "Unique" : "Duplicate/Ambiguous", Diagnostic = failure };
                        yield return chart;
                        if (detail) Evidence.Details[id] = chart;
                    }
                }
            }
            private void ResolveScopes(IEnumerable<Type59ChartEvidence> charts)
            {
                var numeric = new List<int>();
                foreach (var chart in charts.Where(item => item.Correlation == "Unique"))
                {
                    var row = chart.Rows[0]; var scope = Text(row, "primaryentitytypecode");
                    int code;
                    if (!string.IsNullOrWhiteSpace(scope) && !int.TryParse(scope, NumberStyles.Integer,
                        CultureInfo.InvariantCulture, out code))
                    {
                        if (LogicalScope(scope)) chart.Entity = scope;
                        continue;
                    }
                    var value = Number(row, "primaryentitytypecode");
                    if (!value.HasValue && int.TryParse(scope, NumberStyles.Integer, CultureInfo.InvariantCulture, out code)) value = code;
                    if (value.HasValue) { chart.ScopeCode = value; numeric.Add(value.Value); }
                }
                foreach (var batch in Batches(numeric.Distinct().OrderBy(item => item).Where(item => !scopes.ContainsKey(item))))
                {
                    token.ThrowIfCancellationRequested();
                    var query = new EntityQueryExpression { Properties = new MetadataPropertiesExpression("ObjectTypeCode", "LogicalName"),
                        Criteria = new MetadataFilterExpression(LogicalOperator.Or) };
                    foreach (var code in batch) query.Criteria.Conditions.Add(new MetadataConditionExpression(
                        "ObjectTypeCode", MetadataConditionOperator.Equals, code));
                    foreach (var code in batch) scopes[code] = null;
                    try
                    {
                        Evidence.Requests["RetrieveMetadataChanges"] = Evidence.Count("RetrieveMetadataChanges") + 1;
                        var response = service.Execute(new RetrieveMetadataChangesRequest { Query = query }) as RetrieveMetadataChangesResponse;
                        token.ThrowIfCancellationRequested();
                        if (response?.EntityMetadata == null || response.EntityMetadata.Any(item =>
                            !item.ObjectTypeCode.HasValue || !batch.Contains(item.ObjectTypeCode.Value))) continue;
                        foreach (var code in batch)
                        {
                            var matches = response.EntityMetadata.Where(item => item.ObjectTypeCode == code).ToList();
                            if (matches.Count == 1 && LogicalScope(matches[0].LogicalName?.Trim())) scopes[code] = matches[0].LogicalName.Trim();
                        }
                    }
                    catch (OperationCanceledException) { throw; }
                    catch (FaultException) { token.ThrowIfCancellationRequested(); }
                }
                foreach (var chart in charts.Where(item => item.ScopeCode.HasValue)) chart.Entity = scopes[chart.ScopeCode.Value];
            }
            private static void MarkDuplicates(IEnumerable<Type59ChartEvidence> charts, bool b)
            {
                foreach (var group in charts.Where(item => (b ? item.CandidateB : item.CandidateA) != null)
                    .GroupBy(item => b ? item.CandidateB : item.CandidateA, StringComparer.OrdinalIgnoreCase))
                    if (group.Select(item => item.ObjectId).Distinct().Count() > 1)
                        foreach (var chart in group) { if (b) chart.StatusB = "Ambiguous"; else chart.StatusA = "Ambiguous"; }
            }
        }
    }

    internal sealed class Type59RawEvidence
    {
        internal string Solution, Version;
        internal SolutionComponentRecord Record;
    }
    internal sealed class Type59ChartEvidence
    {
        internal Guid ObjectId;
        internal List<Entity> Rows;
        internal string Correlation, Diagnostic, Entity, CandidateA, CandidateB;
        internal string StatusA = "Incomplete", StatusB = "Incomplete";
        internal int? ScopeCode;
        internal bool Detail;
        private readonly Dictionary<string, Type59XmlEvidence> xml = new Dictionary<string, Type59XmlEvidence>();
        internal Type59XmlEvidence Xml(string field)
        {
            if (!xml.TryGetValue(field, out var value)) xml[field] = value = Type59XmlEvidence.Create(Row, field);
            return value;
        }
        internal Entity Row => Correlation == "Unique" ? Rows[0] : null;
        internal void ConstructCandidates()
        {
            var name = Type59EvidenceCollector.Text(Row, "name");
            if (Row == null) return;
            if (string.IsNullOrWhiteSpace(Entity)) { Diagnostic = "Entity logical scope is unresolved; no entity-based candidate."; return; }
            if (string.IsNullOrWhiteSpace(name)) { Diagnostic = "Chart name is blank or unavailable; no candidate."; return; }
            CandidateA = "type59:evidence:A:" + Type59EvidenceCollector.Frame(Entity, name); StatusA = "Unique";
            var type = Type59EvidenceCollector.Number(Row, "type"); var charttype = Type59EvidenceCollector.Number(Row, "charttype");
            if (type.HasValue && charttype.HasValue)
            {
                CandidateB = "type59:evidence:B:" + Type59EvidenceCollector.Frame(Entity,
                    type.Value.ToString(CultureInfo.InvariantCulture), charttype.Value.ToString(CultureInfo.InvariantCulture), name);
                StatusB = "Unique";
            }
        }
    }
    internal sealed class Type59EnvironmentEvidence
    {
        internal EnvironmentIdentity Environment;
        internal DateTimeOffset CapturedUtc;
        internal DateTimeOffset? SnapshotUtc;
        internal readonly List<Type59RawEvidence> Raw = new List<Type59RawEvidence>();
        internal readonly Dictionary<Guid, Type59ChartEvidence> Charts = new Dictionary<Guid, Type59ChartEvidence>();
        internal readonly Dictionary<Guid, Type59ChartEvidence> Details = new Dictionary<Guid, Type59ChartEvidence>();
        internal readonly Dictionary<string, int> Requests = new Dictionary<string, int>();
        internal int Count(string table) => Requests.TryGetValue(table, out var count) ? count : 0;
        internal bool Complete(string solution) => Raw.Where(item => StringComparer.OrdinalIgnoreCase.Equals(item.Solution, solution))
            .All(item => item.Record.ObjectId.HasValue && item.Record.ObjectId != Guid.Empty &&
                Charts.TryGetValue(item.Record.ObjectId.Value, out var chart) && chart.Correlation == "Unique" && chart.CandidateA != null);
    }
    internal sealed class Type59PairEvidence
    {
        internal string Solution;
        internal Type59ChartEvidence Source, Target;
        internal bool UniqueInSolution;
        internal bool UniqueA => UniqueInSolution;
        internal bool DifferentIds => Source.ObjectId != Target.ObjectId;
        internal bool UnmanagedToManaged => Type59EvidenceCollector.Boolean(Source.Row, "ismanaged") == false &&
            Type59EvidenceCollector.Boolean(Target.Row, "ismanaged") == true;
    }
    internal sealed class Type59LifecycleEvidence
    {
        internal string Solution, CandidateA;
        internal Type59ChartEvidence Source, Target;
        internal readonly List<string> Outcomes = new List<string>();
    }
    internal sealed class Type59EvidenceReport
    {
        internal Type59EnvironmentEvidence Source, Target;
        internal bool Discovery;
        internal readonly List<string> SharedSolutions = new List<string>();
        internal readonly List<Type59PairEvidence> Pairs = new List<Type59PairEvidence>();
        internal readonly List<Type59PairEvidence> SelectedPairs = new List<Type59PairEvidence>();
        internal readonly List<Type59LifecycleEvidence> Lifecycle = new List<Type59LifecycleEvidence>();
        internal int Population(string solution) => Source.Raw.Concat(Target.Raw).Count(item =>
            StringComparer.OrdinalIgnoreCase.Equals(item.Solution, solution));
        internal void Analyze()
        {
            Pairs.Clear(); Lifecycle.Clear();
            foreach (var solution in SharedSolutions)
            {
                var left = InSolution(Source, solution); var right = InSolution(Target, solution);
                foreach (var chart in left.Where(item => item.CandidateA != null))
                    foreach (var match in right.Where(item => StringComparer.OrdinalIgnoreCase.Equals(item.CandidateA, chart.CandidateA)))
                        Pairs.Add(new Type59PairEvidence { Solution = solution, Source = chart, Target = match,
                            UniqueInSolution = left.Count(item => StringComparer.OrdinalIgnoreCase.Equals(item.CandidateA, chart.CandidateA)) == 1 &&
                                right.Count(item => StringComparer.OrdinalIgnoreCase.Equals(item.CandidateA, chart.CandidateA)) == 1 });
            }
            foreach (var solution in Source.Raw.Concat(Target.Raw).Select(item => item.Solution)
                .Distinct(StringComparer.OrdinalIgnoreCase).OrderBy(item => item, StringComparer.OrdinalIgnoreCase))
            {
                var left = InSolution(Source, solution); var right = InSolution(Target, solution);
                foreach (var key in left.Concat(right).Select(item => item.CandidateA).Where(item => item != null)
                    .Distinct(StringComparer.OrdinalIgnoreCase).OrderBy(item => item, StringComparer.OrdinalIgnoreCase))
                {
                    var a = left.Where(item => StringComparer.OrdinalIgnoreCase.Equals(item.CandidateA, key)).ToList();
                    var b = right.Where(item => StringComparer.OrdinalIgnoreCase.Equals(item.CandidateA, key)).ToList();
                    var evidence = new Type59LifecycleEvidence { Solution = solution, CandidateA = key,
                        Source = a.Count == 1 ? a[0] : null, Target = b.Count == 1 ? b[0] : null };
                    if (a.Count > 1 || b.Count > 1) evidence.Outcomes.Add("Ambiguous");
                    else if (a.Count == 0 || b.Count == 0)
                    {
                        evidence.Outcomes.Add("OneSidedEvidence");
                        if (!Source.Complete(solution) || !Target.Complete(solution)) evidence.Outcomes.Add("Incomplete");
                    }
                    else ClassifyPair(evidence);
                    Lifecycle.Add(evidence);
                }
                foreach (var raw in Source.Raw.Concat(Target.Raw).Where(item =>
                    StringComparer.OrdinalIgnoreCase.Equals(item.Solution, solution) &&
                    (!item.Record.ObjectId.HasValue || item.Record.ObjectId == Guid.Empty)))
                    Lifecycle.Add(Incomplete(solution));
                foreach (var chart in left.Concat(right).Where(item => item.CandidateA == null))
                {
                    var evidence = Incomplete(solution);
                    if (chart.Correlation == "Duplicate/Ambiguous") evidence.Outcomes.Add("Ambiguous");
                    Lifecycle.Add(evidence);
                }
            }
        }
        private static Type59LifecycleEvidence Incomplete(string solution)
        {
            var evidence = new Type59LifecycleEvidence { Solution = solution };
            evidence.Outcomes.Add("Incomplete"); return evidence;
        }
        private static void ClassifyPair(Type59LifecycleEvidence evidence)
        {
            var a = evidence.Source; var b = evidence.Target;
            evidence.Outcomes.Add(a.ObjectId == b.ObjectId ? "SamePrimaryId" : "DifferentPrimaryId");
            var uniqueA = Type59EvidenceCollector.Id(a.Row, "savedqueryvisualizationidunique");
            var uniqueB = Type59EvidenceCollector.Id(b.Row, "savedqueryvisualizationidunique");
            if (uniqueA.HasValue && uniqueB.HasValue)
                evidence.Outcomes.Add(uniqueA == uniqueB ? "SameUniqueId" : "DifferentUniqueId");
            else evidence.Outcomes.Add("Incomplete");
            bool completeDefinition = true, sameDefinition = true;
            foreach (var field in new[] { "datadescription", "presentationdescription" })
            {
                var x = a.Xml(field); var y = b.Xml(field);
                completeDefinition &= x.Hash != null && y.Hash != null;
                sameDefinition &= x.Hash == y.Hash;
            }
            foreach (var field in new[] { "type", "charttype" })
            {
                var x = Type59EvidenceCollector.Number(a.Row, field); var y = Type59EvidenceCollector.Number(b.Row, field);
                completeDefinition &= x.HasValue && y.HasValue; sameDefinition &= x == y;
            }
            var defaultA = Type59EvidenceCollector.Boolean(a.Row, "isdefault");
            var defaultB = Type59EvidenceCollector.Boolean(b.Row, "isdefault");
            completeDefinition &= defaultA.HasValue && defaultB.HasValue; sameDefinition &= defaultA == defaultB;
            if (completeDefinition) evidence.Outcomes.Add(sameDefinition ? "SameDefinition" : "DifferentDefinition");
            else if (!evidence.Outcomes.Contains("Incomplete")) evidence.Outcomes.Add("Incomplete");
            var managedA = Type59EvidenceCollector.Boolean(a.Row, "ismanaged");
            var managedB = Type59EvidenceCollector.Boolean(b.Row, "ismanaged");
            if (managedA.HasValue && managedB.HasValue)
            {
                if (managedA != managedB) evidence.Outcomes.Add("ManagedTransition");
                if (managedA == false && managedB == true) evidence.Outcomes.Add("UnmanagedToManaged");
            }
            else if (!evidence.Outcomes.Contains("Incomplete")) evidence.Outcomes.Add("Incomplete");
        }
        private static List<Type59ChartEvidence> InSolution(Type59EnvironmentEvidence side, string solution) =>
            side.Raw.Where(item => StringComparer.OrdinalIgnoreCase.Equals(item.Solution, solution) && item.Record.ObjectId.HasValue)
                .Select(item => item.Record.ObjectId.Value).Distinct().Where(side.Charts.ContainsKey).Select(item => side.Charts[item]).ToList();

        internal string Build()
        {
            var output = new StringBuilder("TYPE 59 EVIDENCE ONLY - NO PRODUCTION MEMBERSHIP DECISIONS OR DEFINITION CONTRACT\r\n");
            AppendSide(output, "Source", Source); AppendSide(output, "Target", Target);
            output.AppendLine("SHARED SOLUTION INVENTORY");
            foreach (var solution in SharedSolutions) output.AppendLine("  solution=" + Type59EvidenceCollector.Safe(solution) +
                "; Source raw=" + Source.Raw.Count(item => StringComparer.OrdinalIgnoreCase.Equals(item.Solution, solution)) +
                "; Target raw=" + Target.Raw.Count(item => StringComparer.OrdinalIgnoreCase.Equals(item.Solution, solution)));
            output.AppendLine("Source-only solution unique names=[" + string.Join(", ", Source.Raw.Select(item => item.Solution)
                .Distinct(StringComparer.OrdinalIgnoreCase).Except(Target.Raw.Select(item => item.Solution), StringComparer.OrdinalIgnoreCase)
                .OrderBy(item => item, StringComparer.OrdinalIgnoreCase).Select(Type59EvidenceCollector.Safe)) + "]");
            output.AppendLine("Target-only solution unique names=[" + string.Join(", ", Target.Raw.Select(item => item.Solution)
                .Distinct(StringComparer.OrdinalIgnoreCase).Except(Source.Raw.Select(item => item.Solution), StringComparer.OrdinalIgnoreCase)
                .OrderBy(item => item, StringComparer.OrdinalIgnoreCase).Select(Type59EvidenceCollector.Safe)) + "]");
            output.AppendLine("SOURCE / TARGET RECONCILIATION");
            foreach (var pair in Pairs.OrderBy(item => item.Solution, StringComparer.OrdinalIgnoreCase)
                .ThenBy(item => item.Source.CandidateA, StringComparer.OrdinalIgnoreCase).ThenBy(item => item.Source.ObjectId).ThenBy(item => item.Target.ObjectId))
                output.AppendLine("  solution=" + Type59EvidenceCollector.Safe(pair.Solution) + "; CandidateA=" + Type59EvidenceCollector.Safe(pair.Source.CandidateA) +
                    "; status=" + (pair.UniqueA ? "Unique candidate match" : "Ambiguous candidate group") +
                    "; Source id=" + pair.Source.ObjectId + "; Target id=" + pair.Target.ObjectId + "; differingIds=" + pair.DifferentIds +
                    "; unmanagedSourceToManagedTarget=" + pair.UnmanagedToManaged +
                    "; CandidateBEqual=" + (pair.Source.CandidateB != null && pair.Target.CandidateB != null &&
                        StringComparer.OrdinalIgnoreCase.Equals(pair.Source.CandidateB, pair.Target.CandidateB)) +
                    "; classification Source=" + Type59EvidenceCollector.Number(pair.Source.Row, "type") + "/" + Type59EvidenceCollector.Number(pair.Source.Row, "charttype") +
                    "; classification Target=" + Type59EvidenceCollector.Number(pair.Target.Row, "type") + "/" + Type59EvidenceCollector.Number(pair.Target.Row, "charttype"));
            foreach (var solution in SharedSolutions)
            {
                AppendOneSide(output, Source, Target, solution, "Source"); AppendOneSide(output, Target, Source, solution, "Target");
                var left = InSolution(Source, solution); var right = InSolution(Target, solution);
                foreach (var key in left.Concat(right).Select(item => item.CandidateB).Where(item => item != null)
                    .Distinct(StringComparer.OrdinalIgnoreCase).OrderBy(item => item, StringComparer.OrdinalIgnoreCase))
                {
                    var a = left.Where(item => StringComparer.OrdinalIgnoreCase.Equals(item.CandidateB, key)).ToList();
                    var b = right.Where(item => StringComparer.OrdinalIgnoreCase.Equals(item.CandidateB, key)).ToList();
                    output.AppendLine("  solution=" + Type59EvidenceCollector.Safe(solution) + "; CandidateB=" + Type59EvidenceCollector.Safe(key) +
                        "; status=" + (a.Count > 1 || b.Count > 1 ? "Ambiguous" : a.Count == 1 && b.Count == 1 ?
                            "Unique candidate match (B not approved)" : "One-sided Candidate B evidence only; no absence inference"));
                }
            }
            output.AppendLine("Unique matched solution/candidate references=" + Pairs.Count(item => item.UniqueA) +
                "; distinct differing-ID pairs=" + Pairs.Where(item => item.UniqueA && item.DifferentIds)
                    .Select(item => item.Source.ObjectId + ":" + item.Target.ObjectId).Distinct().Count());
            AppendLifecycle(output);
            output.AppendLine("DETAILED XML EVIDENCE: selected pairs=" + SelectedPairs.Count);
            foreach (var pair in SelectedPairs)
            {
                output.AppendLine("  solution=" + Type59EvidenceCollector.Safe(pair.Solution) + "; CandidateA=" + Type59EvidenceCollector.Safe(pair.Source.CandidateA));
                Source.Details.TryGetValue(pair.Source.ObjectId, out var left); Target.Details.TryGetValue(pair.Target.ObjectId, out var right);
                if (!ConsistentDetail(pair.Source, left) || !ConsistentDetail(pair.Target, right))
                {
                    output.AppendLine("  Incomplete or contradictory detailed correlation/identity evidence; no XML pair comparison is usable.");
                    continue;
                }
                foreach (var field in new[] { "datadescription", "presentationdescription" })
                {
                    var a = left.Xml(field); var b = right.Xml(field);
                    output.AppendLine("  " + field + ": Source " + a.Summary + "; Target " + b.Summary);
                    if (a.Canonical != null && b.Canonical != null && a.Canonical != b.Canonical)
                        output.AppendLine("    " + Type59XmlEvidence.Difference(a.Canonical, b.Canonical));
                }
                output.AppendLine("  isdefault Source=" + Type59EvidenceCollector.Boolean(left?.Row, "isdefault") +
                    "; Target=" + Type59EvidenceCollector.Boolean(right?.Row, "isdefault"));
                output.AppendLine("  detailed correlation Source=" + left?.Correlation + "; Target=" + right?.Correlation);
            }
            output.AppendLine("Names are localizable and rename-sensitive; portability is unproven. Candidate B is evidence only, not approved identity.");
            output.AppendLine("Duplicate backing candidates remain Ambiguous. Incomplete correlation/scope/name evidence cannot prove absence. One-sided labels are evidence only, never membership results.");
            output.AppendLine("Live Type 59 observations require review independently of production primary-ID membership resolution.");
            return output.ToString();
        }
        private void AppendLifecycle(StringBuilder output)
        {
            output.AppendLine("LIFECYCLE CORRELATION MATRIX (per solution; overlapping diagnostic categories, not membership results)");
            output.AppendLine("Definition evidence equality requires both canonical XML hashes, type, charttype and isdefault. Local IDs, componentstate and ismanaged are excluded.");
            foreach (var item in Lifecycle)
                output.AppendLine("  solution=" + Type59EvidenceCollector.Safe(item.Solution) + "; CandidateA=" +
                    Type59EvidenceCollector.Safe(item.CandidateA) + "; outcomes=" + string.Join(",", item.Outcomes) +
                    "; Source id=" + item.Source?.ObjectId + "; Target id=" + item.Target?.ObjectId +
                    "; Source uniqueId=" + Type59EvidenceCollector.Id(item.Source?.Row, "savedqueryvisualizationidunique") +
                    "; Target uniqueId=" + Type59EvidenceCollector.Id(item.Target?.Row, "savedqueryvisualizationidunique"));
            foreach (var category in new[] { "SamePrimaryId", "DifferentPrimaryId", "SameUniqueId", "DifferentUniqueId",
                "SameDefinition", "DifferentDefinition", "ManagedTransition", "UnmanagedToManaged", "OneSidedEvidence", "Ambiguous", "Incomplete" })
                output.AppendLine("  " + category + "=" + Lifecycle.Count(item => item.Outcomes.Contains(category)));
            var ambiguousB = Source.Raw.Concat(Target.Raw).Select(item => item.Solution)
                .Distinct(StringComparer.OrdinalIgnoreCase).Sum(solution => InSolution(Source, solution).Concat(InSolution(Target, solution))
                    .Where(item => item.CandidateB != null).GroupBy(item => item.CandidateB, StringComparer.OrdinalIgnoreCase)
                    .Count(group => InSolution(Source, solution).Count(item => StringComparer.OrdinalIgnoreCase.Equals(item.CandidateB, group.Key)) > 1 ||
                        InSolution(Target, solution).Count(item => StringComparer.OrdinalIgnoreCase.Equals(item.CandidateB, group.Key)) > 1));
            output.AppendLine("  CandidateB ambiguous solution/candidate groups=" + ambiguousB);
            output.AppendLine("DIFFERING PRIMARY-ID SEMANTIC PAIRS");
            var different = Pairs.Where(item => item.UniqueA && item.DifferentIds).ToList();
            if (different.Count == 0) output.AppendLine("  No uniquely matched Candidate A pair with differing savedqueryvisualizationid was observed in this capture.");
            foreach (var pair in different)
            {
                output.AppendLine("  solution=" + Type59EvidenceCollector.Safe(pair.Solution) + "; candidateUniqueness=UniqueWithinSolution; unmanagedSourceToManagedTarget=" + pair.UnmanagedToManaged);
                AppendChart(output, pair.Source, "Source"); AppendChart(output, pair.Target, "Target");
                foreach (var field in new[] { "datadescription", "presentationdescription" })
                    output.AppendLine("  " + field + " hashEquality=" + (pair.Source.Xml(field).Hash == null || pair.Target.Xml(field).Hash == null
                        ? "Incomplete" : pair.Source.Xml(field).Hash == pair.Target.Xml(field).Hash ? "EqualObserved" : "DifferentObserved"));
            }
            var uniquePairs = Lifecycle.Where(item => item.Outcomes.Contains("SamePrimaryId") || item.Outcomes.Contains("DifferentPrimaryId"))
                .GroupBy(item => item.Source.ObjectId.ToString("D") + ":" + item.Target.ObjectId.ToString("D"))
                .Select(item => item.First()).ToList();
            output.AppendLine("PRIMARY-ID PORTABILITY ASSESSMENT");
            output.AppendLine("  unique semantic backing pairs=" + uniquePairs.Count + "; matched solution/candidate references=" + Pairs.Count(item => item.UniqueA));
            foreach (var category in new[] { "SamePrimaryId", "DifferentPrimaryId", "DifferentUniqueId", "DifferentDefinition" })
                output.AppendLine("  distinct pair " + category + "=" + uniquePairs.Count(item => item.Outcomes.Contains(category)));
            output.AppendLine("  ambiguous semantic identity groups=" + Lifecycle.Count(item => item.CandidateA != null && item.Outcomes.Contains("Ambiguous")));
            output.AppendLine("  ambiguous Candidate B groups=" + ambiguousB);
            output.AppendLine("  Candidate B ambiguity is reported independently in SOURCE / TARGET RECONCILIATION; it never disambiguates Candidate A.");
            output.AppendLine("  Primary-ID identity: same-ID observations demonstrate retention only; differing-ID semantic pairs must be reviewed before concluding continuity or recreation.");
            output.AppendLine("  Semantic identity: unique name-based matches are observations, not proof of continued identity; names remain localizable/rename-sensitive and collision-prone.");
            output.AppendLine("  Sufficiency: neither identity is established for production by this capture alone. Lifecycle operation provenance and duplicate/incomplete cases require review; this report cannot override the independent production primary-ID resolver.");
        }
        private static bool ConsistentDetail(Type59ChartEvidence initial, Type59ChartEvidence detail) =>
            detail?.Row != null && initial.Row != null && new[] { "primaryentitytypecode", "name" }.All(field =>
                StringComparer.OrdinalIgnoreCase.Equals(Type59EvidenceCollector.Text(initial.Row, field), Type59EvidenceCollector.Text(detail.Row, field)) &&
                Type59EvidenceCollector.Number(initial.Row, field) == Type59EvidenceCollector.Number(detail.Row, field)) &&
            new[] { "type", "charttype" }.All(field => Type59EvidenceCollector.Number(initial.Row, field) == Type59EvidenceCollector.Number(detail.Row, field));
        private static void AppendOneSide(StringBuilder text, Type59EnvironmentEvidence side,
            Type59EnvironmentEvidence other, string solution, string label)
        {
            var opposite = InSolution(other, solution);
            foreach (var chart in InSolution(side, solution).Where(item => item.CandidateA != null &&
                !opposite.Any(row => StringComparer.OrdinalIgnoreCase.Equals(item.CandidateA, row.CandidateA))))
                text.AppendLine("  solution=" + Type59EvidenceCollector.Safe(solution) + "; " + label + " candidate=" +
                    Type59EvidenceCollector.Safe(chart.CandidateA) + "; status=" + (InSolution(side, solution).Count(item =>
                        StringComparer.OrdinalIgnoreCase.Equals(item.CandidateA, chart.CandidateA)) > 1 ? "Ambiguous" :
                        side.Complete(solution) && other.Complete(solution) ? label + "-only evidence" : "Indeterminate: incomplete inventory evidence"));
        }
        private static void AppendSide(StringBuilder output, string label, Type59EnvironmentEvidence side)
        {
            output.AppendLine(label + " environment=" + Type59EvidenceCollector.Safe(side.Environment.DisplayName) + "; capturedUtc=" +
                side.CapturedUtc.ToUniversalTime().ToString("o", CultureInfo.InvariantCulture) + "; snapshotUtc=" + side.SnapshotUtc?.ToUniversalTime().ToString("o") +
                "; raw=" + side.Raw.Count + "; distinct nonblank=" + side.Raw.Select(item => item.Record.ObjectId)
                    .Where(item => item.HasValue && item != Guid.Empty).Distinct().Count() + "; blank=" + side.Raw.Count(item => !item.Record.ObjectId.HasValue || item.Record.ObjectId == Guid.Empty));
            foreach (var group in side.Raw.GroupBy(item => item.Solution, StringComparer.OrdinalIgnoreCase).OrderBy(item => item.Key, StringComparer.OrdinalIgnoreCase))
                output.AppendLine("  solution=" + Type59EvidenceCollector.Safe(group.Key) + "; version=" + Type59EvidenceCollector.Safe(group.First().Version) + "; raw=" + group.Count());
            foreach (var raw in side.Raw.OrderBy(item => item.Solution, StringComparer.OrdinalIgnoreCase).ThenBy(item => item.Record.SolutionComponentId))
                output.AppendLine("  membership solution=" + Type59EvidenceCollector.Safe(raw.Solution) + "; solutioncomponentid=" + raw.Record.SolutionComponentId +
                    "; objectid=" + raw.Record.ObjectId + "; rootsolutioncomponentid=" + raw.Record.RootSolutionComponentId +
                    "; rootcomponentbehavior=" + raw.Record.RootComponentBehavior + "; ismetadata=" + raw.Record.IsMetadata);
            foreach (var group in side.Raw.Where(item => item.Record.ObjectId.HasValue && item.Record.ObjectId != Guid.Empty)
                .GroupBy(item => item.Record.ObjectId.Value).Where(item => item.Count() > 1).OrderBy(item => item.Key))
                output.AppendLine("  repeated raw membership objectid=" + group.Key + "; raw references=" + group.Count() + "; not duplicate backing records");
            foreach (var chart in side.Charts.Values.OrderBy(item => item.ObjectId)) AppendChart(output, chart, "backing");
            foreach (var chart in side.Details.Values.Where(item => !side.Charts.ContainsKey(item.ObjectId)).OrderBy(item => item.ObjectId)) AppendChart(output, chart, "detail audit");
            output.AppendLine(label + " SUMMARY: retrieved distinct IDs=" + side.Charts.Count +
                "; unique correlations=" + side.Charts.Values.Count(item => item.Correlation == "Unique") +
                "; missing correlations=" + side.Charts.Values.Count(item => item.Correlation == "Missing") +
                "; duplicate correlations=" + side.Charts.Values.Count(item => item.Correlation == "Duplicate/Ambiguous") +
                "; incomplete correlations=" + side.Charts.Values.Count(item => item.Correlation == "Incomplete") +
                "; Candidate A ambiguous backing records=" + side.Charts.Values.Count(item => item.StatusA == "Ambiguous") +
                "; Candidate B ambiguous backing records=" + side.Charts.Values.Count(item => item.StatusB == "Ambiguous") +
                "; Candidate A incomplete=" + side.Charts.Values.Count(item => item.StatusA == "Incomplete") +
                "; Candidate B incomplete=" + side.Charts.Values.Count(item => item.StatusB == "Incomplete"));
            output.AppendLine(label + " REQUEST LEDGER: solutioncomponent=" + side.Count("solutioncomponent") +
                "; savedqueryvisualization=" + side.Count(Type59EvidenceCollector.EntityName) + "; RetrieveMetadataChanges=" +
                side.Count("RetrieveMetadataChanges") + "; WhoAmI=0; writes=0");
        }
        private static void AppendChart(StringBuilder output, Type59ChartEvidence chart, string label)
        {
            output.AppendLine("  " + label + " objectid=" + chart.ObjectId + "; correlation=" + chart.Correlation +
                "; objectidEqualsPrimaryId=" + (chart.Row != null && Type59EvidenceCollector.Id(chart.Row, Type59EvidenceCollector.PrimaryId) == chart.ObjectId) +
                "; entity=" + Type59EvidenceCollector.Safe(chart.Entity) + "; CandidateA=" + Type59EvidenceCollector.Safe(chart.CandidateA) +
                "; statusA=" + chart.StatusA + "; CandidateB=" + Type59EvidenceCollector.Safe(chart.CandidateB) + "; statusB=" + chart.StatusB +
                "; diagnostic=" + Type59EvidenceCollector.Safe(chart.Diagnostic));
            foreach (var row in chart.Rows)
            {
                output.AppendLine("    " + string.Join("; ", Type59EvidenceCollector.LightColumns.Concat(new[] { "componentstate", "isdefault" })
                    .Select(column => column + "=" + Type59EvidenceCollector.Safe(row.Attributes.TryGetValue(column, out var value)
                        ? value is OptionSetValue ? ((OptionSetValue)value).Value.ToString(CultureInfo.InvariantCulture) : Convert.ToString(value, CultureInfo.InvariantCulture) : null))));
                if (chart.Detail)
                    foreach (var field in new[] { "datadescription", "presentationdescription" })
                        output.AppendLine("    " + field + " audit: " + (chart.Row == row ? chart.Xml(field) : Type59XmlEvidence.Create(row, field)).Summary);
            }
        }
    }
    internal sealed class Type59XmlEvidence
    {
        internal string Canonical, Status, Hash;
        internal string Summary => "status=" + Status + "; canonicalLength=" + (Canonical == null ? "(unavailable)" : Canonical.Length.ToString(CultureInfo.InvariantCulture)) +
            "; SHA256=" + (Hash ?? "(unavailable)");
        internal static Type59XmlEvidence Create(Entity row, string field)
        {
            var result = new Type59XmlEvidence();
            if (row == null || !row.Attributes.TryGetValue(field, out var value)) { result.Status = "NotSupplied"; return result; }
            if (value == null || value is string && string.IsNullOrWhiteSpace((string)value)) { result.Status = "Blank"; return result; }
            if (!(value is string)) { result.Status = "InvalidValueType"; return result; }
            try
            {
                var settings = new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null, IgnoreWhitespace = true };
                var document = new XmlDocument { XmlResolver = null, PreserveWhitespace = false };
                using (var reader = XmlReader.Create(new StringReader((string)value), settings)) document.Load(reader);
                result.Canonical = document.DocumentElement.OuterXml;
                using (var sha = SHA256.Create()) result.Hash = BitConverter.ToString(sha.ComputeHash(Encoding.UTF8.GetBytes(result.Canonical))).Replace("-", "").ToLowerInvariant();
                result.Status = "CanonicalXml";
            }
            catch (XmlException) { result.Status = "MalformedOrUnsafeXml"; }
            return result;
        }
        internal static string Difference(string source, string target)
        {
            int offset = 0; while (offset < Math.Min(source.Length, target.Length) && source[offset] == target[offset]) offset++;
            int start = Math.Max(0, offset - 40);
            return "First difference offset=" + offset + "; Source='" + Type59EvidenceCollector.Safe(source.Substring(start, Math.Min(120, source.Length - start))) +
                "'; Target='" + Type59EvidenceCollector.Safe(target.Substring(start, Math.Min(120, target.Length - start))) + "'";
        }
    }
}
#endif
