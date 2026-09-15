using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.Linq;
using System.Text;

namespace D365SolutionComparer.Models.Membership
{
    public enum AppSettingEvidenceState
    {
        Confirmed,
        Unresolved,
        Ambiguous,
        Unavailable,
        Faulted
    }

    public sealed class AppSettingComponentTypeSummary
    {
        public AppSettingComponentTypeSummary(int componentType, string formattedLabel, int count,
            int distinctObjectIdCount, int blankObjectIdCount)
        {
            ComponentType = componentType;
            FormattedLabel = formattedLabel ?? string.Empty;
            Count = count;
            DistinctObjectIdCount = distinctObjectIdCount;
            BlankObjectIdCount = blankObjectIdCount;
        }

        public int ComponentType { get; }
        public string FormattedLabel { get; }
        public int Count { get; }
        public int DistinctObjectIdCount { get; }
        public int BlankObjectIdCount { get; }
    }

    public sealed class AppSettingEntityMetadataEvidence
    {
        public AppSettingEntityMetadataEvidence(string logicalName, AppSettingEvidenceState state,
            string primaryIdAttribute, string primaryNameAttribute, IEnumerable<string> attributes,
            IEnumerable<string> lookups, IEnumerable<string> relationships, string diagnostic)
        {
            LogicalName = logicalName ?? string.Empty;
            State = state;
            PrimaryIdAttribute = primaryIdAttribute ?? string.Empty;
            PrimaryNameAttribute = primaryNameAttribute ?? string.Empty;
            Attributes = new ReadOnlyCollection<string>((attributes ?? Enumerable.Empty<string>()).ToList());
            LookupAttributes = new ReadOnlyCollection<string>((lookups ?? Enumerable.Empty<string>()).ToList());
            Relationships = new ReadOnlyCollection<string>((relationships ?? Enumerable.Empty<string>()).ToList());
            Diagnostic = diagnostic ?? string.Empty;
        }

        public string LogicalName { get; }
        public AppSettingEvidenceState State { get; }
        public string PrimaryIdAttribute { get; }
        public string PrimaryNameAttribute { get; }
        public IReadOnlyList<string> Attributes { get; }
        public IReadOnlyList<string> LookupAttributes { get; }
        public IReadOnlyList<string> Relationships { get; }
        public string Diagnostic { get; }
    }

    public sealed class AppSettingRecordEvidence
    {
        public AppSettingRecordEvidence(Guid recordId, IEnumerable<KeyValuePair<string, string>> fields)
        {
            RecordId = recordId;
            var copy = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (var item in fields ?? Enumerable.Empty<KeyValuePair<string, string>>())
                copy[item.Key] = item.Value;
            Fields = new ReadOnlyDictionary<string, string>(copy);
        }

        public Guid RecordId { get; }
        public IReadOnlyDictionary<string, string> Fields { get; }
    }

    public sealed class AppSettingEntityCorrelation
    {
        public AppSettingEntityCorrelation(string entityLogicalName, AppSettingEvidenceState state,
            string primaryIdAttribute, int matchedRecordCount, IEnumerable<AppSettingRecordEvidence> records,
            string diagnostic)
        {
            EntityLogicalName = entityLogicalName ?? string.Empty;
            State = state;
            PrimaryIdAttribute = primaryIdAttribute ?? string.Empty;
            MatchedRecordCount = matchedRecordCount;
            Records = new ReadOnlyCollection<AppSettingRecordEvidence>(
                (records ?? Enumerable.Empty<AppSettingRecordEvidence>()).ToList());
            Diagnostic = diagnostic ?? string.Empty;
        }

        public string EntityLogicalName { get; }
        public AppSettingEvidenceState State { get; }
        public string PrimaryIdAttribute { get; }
        public int MatchedRecordCount { get; }
        public IReadOnlyList<AppSettingRecordEvidence> Records { get; }
        public string Diagnostic { get; }
    }

    public sealed class AppSettingCandidateEvidence
    {
        public AppSettingCandidateEvidence(string environmentName, string solutionUniqueName,
            Guid solutionId, Guid solutionComponentId, int componentType, string formattedLabel,
            Guid? objectId, int? rootComponentBehavior, Guid? rootSolutionComponentId, bool? isMetadata,
            string candidateReason, IEnumerable<AppSettingEntityCorrelation> correlations,
            AppSettingEvidenceState state, string settingDefinitionId, string settingDefinitionName,
            string parentAppModuleId, string parentAppModuleUniqueName, string parentAppModuleName,
            string candidateCompositeIdentity, string diagnostic)
        {
            EnvironmentName = environmentName ?? string.Empty;
            SolutionUniqueName = solutionUniqueName ?? string.Empty;
            SolutionId = solutionId;
            SolutionComponentId = solutionComponentId;
            ComponentType = componentType;
            FormattedLabel = formattedLabel ?? string.Empty;
            ObjectId = objectId;
            RootComponentBehavior = rootComponentBehavior;
            RootSolutionComponentId = rootSolutionComponentId;
            IsMetadata = isMetadata;
            CandidateReason = candidateReason ?? string.Empty;
            Correlations = new ReadOnlyCollection<AppSettingEntityCorrelation>(
                (correlations ?? Enumerable.Empty<AppSettingEntityCorrelation>()).ToList());
            State = state;
            SettingDefinitionId = settingDefinitionId ?? string.Empty;
            SettingDefinitionName = settingDefinitionName ?? string.Empty;
            ParentAppModuleId = parentAppModuleId ?? string.Empty;
            ParentAppModuleUniqueName = parentAppModuleUniqueName ?? string.Empty;
            ParentAppModuleName = parentAppModuleName ?? string.Empty;
            CandidateCompositeIdentity = candidateCompositeIdentity ?? string.Empty;
            Diagnostic = diagnostic ?? string.Empty;
        }

        public string EnvironmentName { get; }
        public string SolutionUniqueName { get; }
        public Guid SolutionId { get; }
        public Guid SolutionComponentId { get; }
        public int ComponentType { get; }
        public string FormattedLabel { get; }
        public Guid? ObjectId { get; }
        public int? RootComponentBehavior { get; }
        public Guid? RootSolutionComponentId { get; }
        public bool? IsMetadata { get; }
        public string CandidateReason { get; }
        public IReadOnlyList<AppSettingEntityCorrelation> Correlations { get; }
        public AppSettingEvidenceState State { get; }
        public string SettingDefinitionId { get; }
        public string SettingDefinitionName { get; }
        public string ParentAppModuleId { get; }
        public string ParentAppModuleUniqueName { get; }
        public string ParentAppModuleName { get; }
        public string CandidateCompositeIdentity { get; }
        public string Diagnostic { get; }
    }

    public sealed class AppSettingRequestSummary
    {
        public AppSettingRequestSummary(int whoAmI, int solutionComponent, int metadataDiscovery,
            int candidateBackingEntity, int settingDefinition, int appModule, int other, int total)
        {
            WhoAmI = whoAmI;
            SolutionComponent = solutionComponent;
            MetadataDiscovery = metadataDiscovery;
            CandidateBackingEntity = candidateBackingEntity;
            SettingDefinition = settingDefinition;
            AppModule = appModule;
            Other = other;
            Total = total;
        }

        public int WhoAmI { get; }
        public int SolutionComponent { get; }
        public int MetadataDiscovery { get; }
        public int CandidateBackingEntity { get; }
        public int SettingDefinition { get; }
        public int AppModule { get; }
        public int Other { get; }
        public int Total { get; }
    }

    public sealed class AppSettingEvidenceReport
    {
        public AppSettingEvidenceReport(string environmentName, string solutionUniqueName, Guid solutionId,
            IEnumerable<AppSettingComponentTypeSummary> componentTypes,
            IEnumerable<AppSettingEntityMetadataEvidence> entityMetadata,
            IEnumerable<AppSettingCandidateEvidence> candidates,
            AppSettingRequestSummary requests, IEnumerable<string> diagnostics)
        {
            EnvironmentName = environmentName ?? string.Empty;
            SolutionUniqueName = solutionUniqueName ?? string.Empty;
            SolutionId = solutionId;
            ComponentTypes = new ReadOnlyCollection<AppSettingComponentTypeSummary>(
                (componentTypes ?? Enumerable.Empty<AppSettingComponentTypeSummary>()).ToList());
            EntityMetadata = new ReadOnlyCollection<AppSettingEntityMetadataEvidence>(
                (entityMetadata ?? Enumerable.Empty<AppSettingEntityMetadataEvidence>()).ToList());
            Candidates = new ReadOnlyCollection<AppSettingCandidateEvidence>(
                (candidates ?? Enumerable.Empty<AppSettingCandidateEvidence>()).ToList());
            CandidateStatistics = AppSettingCandidateStatistics.From(Candidates);
            Requests = requests ?? throw new ArgumentNullException(nameof(requests));
            Diagnostics = new ReadOnlyCollection<string>(
                (diagnostics ?? Enumerable.Empty<string>()).Where(item => item != null).ToList());
        }

        public string EnvironmentName { get; }
        public string SolutionUniqueName { get; }
        public Guid SolutionId { get; }
        public IReadOnlyList<AppSettingComponentTypeSummary> ComponentTypes { get; }
        public IReadOnlyList<AppSettingEntityMetadataEvidence> EntityMetadata { get; }
        public IReadOnlyList<AppSettingCandidateEvidence> Candidates { get; }
        public AppSettingCandidateStatistics CandidateStatistics { get; }
        public AppSettingRequestSummary Requests { get; }
        public IReadOnlyList<string> Diagnostics { get; }
    }

    public sealed class AppSettingCandidateStatistics
    {
        private AppSettingCandidateStatistics(int raw, int complete, int blankParent, int blankDefinition,
            int duplicate, int distinct)
        {
            RawCandidateCount = raw;
            CompleteCompositeCandidateCount = complete;
            BlankParentAppCount = blankParent;
            BlankSettingDefinitionNameCount = blankDefinition;
            DuplicateCompositeCandidateCount = duplicate;
            DistinctCompositeCandidateCount = distinct;
        }

        public int RawCandidateCount { get; }
        public int CompleteCompositeCandidateCount { get; }
        public int BlankParentAppCount { get; }
        public int BlankSettingDefinitionNameCount { get; }
        public int DuplicateCompositeCandidateCount { get; }
        public int DistinctCompositeCandidateCount { get; }

        internal static AppSettingCandidateStatistics From(IEnumerable<AppSettingCandidateEvidence> candidates)
        {
            var rows = candidates.ToList();
            var groups = rows.Where(item => !string.IsNullOrWhiteSpace(item.CandidateCompositeIdentity))
                .GroupBy(item => item.CandidateCompositeIdentity, StringComparer.OrdinalIgnoreCase).ToList();
            return new AppSettingCandidateStatistics(rows.Count, groups.Sum(item => item.Count()),
                rows.Count(item => string.IsNullOrWhiteSpace(item.ParentAppModuleUniqueName)),
                rows.Count(item => string.IsNullOrWhiteSpace(item.SettingDefinitionName)),
                groups.Count(item => item.Count() > 1), groups.Count);
        }
    }

    public sealed class AppSettingCandidateComparison
    {
        public AppSettingCandidateComparison(string identity, string outcome, int sourceCount, int targetCount)
        {
            Identity = identity ?? string.Empty;
            Outcome = outcome ?? string.Empty;
            SourceCount = sourceCount;
            TargetCount = targetCount;
        }

        public string Identity { get; }
        public string Outcome { get; }
        public int SourceCount { get; }
        public int TargetCount { get; }
    }

    public sealed class AppSettingEvidenceComparison
    {
        public AppSettingEvidenceComparison(AppSettingEvidenceReport source, AppSettingEvidenceReport target,
            IEnumerable<AppSettingCandidateComparison> candidates)
        {
            Source = source ?? throw new ArgumentNullException(nameof(source));
            Target = target ?? throw new ArgumentNullException(nameof(target));
            Candidates = new ReadOnlyCollection<AppSettingCandidateComparison>(
                (candidates ?? Enumerable.Empty<AppSettingCandidateComparison>()).ToList());
        }

        public AppSettingEvidenceReport Source { get; }
        public AppSettingEvidenceReport Target { get; }
        public IReadOnlyList<AppSettingCandidateComparison> Candidates { get; }

        public static AppSettingEvidenceComparison Create(AppSettingEvidenceReport source,
            AppSettingEvidenceReport target)
        {
            var sourceGroups = source.Candidates.Where(item => !string.IsNullOrWhiteSpace(item.CandidateCompositeIdentity))
                .GroupBy(item => item.CandidateCompositeIdentity, StringComparer.OrdinalIgnoreCase)
                .ToDictionary(group => group.Key, group => group.Count(), StringComparer.OrdinalIgnoreCase);
            var targetGroups = target.Candidates.Where(item => !string.IsNullOrWhiteSpace(item.CandidateCompositeIdentity))
                .GroupBy(item => item.CandidateCompositeIdentity, StringComparer.OrdinalIgnoreCase)
                .ToDictionary(group => group.Key, group => group.Count(), StringComparer.OrdinalIgnoreCase);
            var keys = sourceGroups.Keys.Concat(targetGroups.Keys)
                .Distinct(StringComparer.OrdinalIgnoreCase).OrderBy(item => item, StringComparer.OrdinalIgnoreCase);
            var rows = new List<AppSettingCandidateComparison>();
            foreach (var key in keys)
            {
                int sourceCount = sourceGroups.ContainsKey(key) ? sourceGroups[key] : 0;
                int targetCount = targetGroups.ContainsKey(key) ? targetGroups[key] : 0;
                string outcome = sourceCount > 1 || targetCount > 1 ? "AmbiguousCandidate" :
                    sourceCount == 1 && targetCount == 1 ? "CandidateStable" :
                    sourceCount == 1 ? "DEVOnlyCandidate" : "UATOnlyCandidate";
                rows.Add(new AppSettingCandidateComparison(key, outcome, sourceCount, targetCount));
            }
            return new AppSettingEvidenceComparison(source, target, rows);
        }
    }
}
