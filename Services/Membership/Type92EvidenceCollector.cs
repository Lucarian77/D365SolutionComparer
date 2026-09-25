#if DEBUG
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading;
using D365SolutionComparer.Models.Membership;
using Microsoft.Xrm.Sdk;

namespace D365SolutionComparer.Services.Membership
{
    /// <summary>Debug-only read-only evidence report; production resolution uses the shared reader.</summary>
    internal sealed class Type92EvidenceCollector
    {
        internal const int BatchSize = Type92MembershipEvidenceReader.BatchSize;

        internal Type92EnvironmentEvidence Capture(IOrganizationService service, MembershipSnapshot snapshot,
            string solutionVersion, CancellationToken cancellationToken, Action<string> progress = null,
            DataverseComponentMetadataCache metadataCache = null) =>
            new Type92MembershipEvidenceReader().Capture(service, snapshot, solutionVersion,
                cancellationToken, progress, metadataCache);

        internal static string CanonicalFilteringAttributes(string value) =>
            Type92MembershipEvidenceReader.CanonicalFilteringAttributes(value);
    }
    internal static class Type92EvidenceReport
    {
        internal static string Build(Type92EnvironmentEvidence source, Type92EnvironmentEvidence target)
        {
            if (source == null || target == null) throw new ArgumentNullException("A completed Source and Target evidence set is required.");
            var text = new StringBuilder("TYPE 92 EVIDENCE ONLY - NO MEMBERSHIP RESOLUTION\r\n");
            AppendEnvironment(text, source);
            AppendEnvironment(text, target);
            text.AppendLine("DEV/UAT RECONCILIATION");
            var sourceCandidates = source.Steps.Where(item => item.Candidate != null)
                .GroupBy(item => item.Candidate).ToDictionary(item => item.Key, item => item.ToList());
            var targetCandidates = target.Steps.Where(item => item.Candidate != null)
                .GroupBy(item => item.Candidate).ToDictionary(item => item.Key, item => item.ToList());
            int matched = 0, sourceOnly = 0, targetOnly = 0, ambiguousCount = 0;
            foreach (var key in sourceCandidates.Keys.Union(targetCandidates.Keys)
                .OrderBy(item => item.ToString(), StringComparer.OrdinalIgnoreCase))
            {
                List<Type92StepEvidence> leftRows, rightRows;
                var hasLeft = sourceCandidates.TryGetValue(key, out leftRows);
                var hasRight = targetCandidates.TryGetValue(key, out rightRows);
                var ambiguous = (hasLeft && leftRows.Any(item => item.CandidateStatus != "Candidate")) ||
                    (hasRight && rightRows.Any(item => item.CandidateStatus != "Candidate"));
                if (ambiguous) ambiguousCount++;
                else if (hasLeft && hasRight) matched++;
                else if (hasLeft) sourceOnly++;
                else targetOnly++;
                text.AppendLine((ambiguous ? "Ambiguous identity group" : hasLeft && hasRight ? "Unique match" :
                    hasLeft ? "DEV only" : "UAT only") + " | " + key);
                if (hasLeft && hasRight && !ambiguous)
                {
                    var left = leftRows[0];
                    var right = rightRows[0];
                    text.AppendLine("  definition evidence: rank=" + left.Rank + "/" + right.Rank +
                        "; filtering=" + left.FilteringCanonical + "/" + right.FilteringCanonical +
                        "; configuration present=" + left.ConfigurationPresent + "/" + right.ConfigurationPresent +
                        "; configuration length=" + left.ConfigurationLength + "/" + right.ConfigurationLength +
                        "; state=" + left.State + "/" + right.State +
                        "; status=" + left.Status + "/" + right.Status +
                        "; managed=" + left.IsManaged + "/" + right.IsManaged +
                        "; componentstate=" + left.ComponentState + "/" + right.ComponentState);
                }
            }
            var incomplete = source.Steps.Concat(target.Steps).Any(item =>
                    item.Candidate == null) ||
                source.Raw.Any(item => !item.ObjectId.HasValue || item.ObjectId == Guid.Empty) ||
                target.Raw.Any(item => !item.ObjectId.HasValue || item.ObjectId == Guid.Empty);
            text.AppendLine("Candidate totals: matched=" + matched + "; DEV-only=" + sourceOnly +
                "; UAT-only=" + targetOnly + "; ambiguous=" + ambiguousCount +
                "; incomplete evidence=" + incomplete);
            text.AppendLine("CONCLUSION");
            if (incomplete)
                text.AppendLine("Some Type 92 evidence is incomplete; review unresolved rows before interpreting one-sided candidates.");
            else text.AppendLine("Review unique, one-sided and ambiguous candidate groups before considering identity portability.");
            text.AppendLine("This report does not change Type 92 membership or definition comparison.");
            return text.ToString();
        }

        private static void AppendEnvironment(StringBuilder text, Type92EnvironmentEvidence evidence)
        {
            var snapshot = evidence.Snapshot;
            text.AppendLine("ENVIRONMENT " + snapshot.Environment.DisplayName + " | solution=" +
                snapshot.SolutionUniqueName + " | version=" + evidence.SolutionVersion +
                " | captured=" + snapshot.CapturedAt.ToUniversalTime().ToString("o", CultureInfo.InvariantCulture));
            text.AppendLine("RAW INVENTORY: count=" + evidence.Raw.Count + "; distinct nonblank objectids=" +
                evidence.Steps.Count + "; blank objectids=" + evidence.Raw.Count(item =>
                    !item.ObjectId.HasValue || item.ObjectId == Guid.Empty));
            text.AppendLine("REQUEST LEDGER (collector only): solutioncomponent=0 (completed snapshot reused); sdkmessageprocessingstep=" +
                evidence.Count("sdkmessageprocessingstep") + "; plugintype=" + evidence.Count("plugintype") +
                "; pluginassembly=" + evidence.Count("pluginassembly") + "; sdkmessage=" +
                evidence.Count("sdkmessage") + "; sdkmessagefilter=" + evidence.Count("sdkmessagefilter") +
                "; other-handler=0; WhoAmI=0; writes=0");
            text.AppendLine("RAW TYPE 92 MEMBERSHIP");
            foreach (var item in evidence.Raw)
                text.AppendLine("  solutioncomponentid=" + item.SolutionComponentId + "; objectid=" + item.ObjectId +
                    "; rootsolutioncomponentid=" + item.RootSolutionComponentId +
                    "; rootcomponentbehavior=" + item.RootComponentBehavior + "; ismetadata=" + item.IsMetadata);
            text.AppendLine("REPEATED MEMBERSHIP GROUPS");
            foreach (var group in evidence.Steps.Where(item => item.MembershipCount > 1))
                text.AppendLine("  objectid=" + group.ObjectId + "; raw rows=" + group.MembershipCount);
            text.AppendLine("CORRELATION AND HANDLER SUMMARY");
            foreach (var group in evidence.Steps.GroupBy(item => item.Correlation).OrderBy(item => item.Key))
                text.AppendLine("  correlation " + group.Key + "=" + group.Count());
            foreach (var group in evidence.Steps.GroupBy(item => item.HandlerCategory ?? "none").OrderBy(item => item.Key))
                text.AppendLine("  handler " + group.Key + "=" + group.Count());
            text.AppendLine("NORMALIZED STEP EVIDENCE");
            foreach (var item in evidence.Steps)
            {
                text.AppendLine("  objectid=" + item.ObjectId + "; correlation=" + item.Correlation +
                    "; stepid=" + item.StepId + "; stepidunique=" + item.StepUniqueId +
                    "; rawMembershipCount=" + item.MembershipCount + "; name=" + item.Name +
                    "; handler=" + item.Handler?.Id + "/" + item.Handler?.LogicalName +
                    "; handlerCategory=" + item.HandlerCategory + "; plugintypeid=" + item.PluginTypeId +
                    "; plugintypeidunique=" + item.PluginTypeUniqueId + "; exportkey=" + item.ExportKey +
                    "; typename=" + item.TypeName + "; pluginTypeName=" + item.PluginTypeName +
                    "; isworkflowactivity=" + item.IsWorkflowActivity +
                    "; assemblyid=" + item.AssemblyId + "; assemblyIdentity=" + item.AssemblyIdentity +
                    "; handlerSemanticIdentity=" + item.HandlerSemanticIdentity +
                    "; assemblyVersion=" + item.AssemblyVersion + "; sdkmessageid=" + item.MessageId +
                    "; message=" + item.MessageName + "; filterid=" + item.FilterId +
                    "; filteridunique=" + item.FilterUniqueId + "; filterName=" + item.FilterName +
                    "; filterMessageId=" + item.FilterMessageId + "; filterAvailability=" + item.FilterAvailability +
                    "; filterStatus=" + item.FilterStatus + "; primary=" + item.PrimaryScope +
                    "; secondary=" + item.SecondaryScope + "; semanticScope=" + item.SemanticScope +
                    "; stage=" + item.Stage + "; mode=" + item.Mode +
                    "; deployment=" + item.Deployment + "; rank=" + item.Rank +
                    "; filteringOriginal=" + item.FilteringOriginal +
                    "; filteringCanonical=" + item.FilteringCanonical +
                    "; configurationPresent=" + item.ConfigurationPresent +
                    "; configurationLength=" + item.ConfigurationLength +
                    "; state=" + item.State + "; status=" + item.Status +
                    "; ismanaged=" + item.IsManaged + "; componentstate=" + item.ComponentState +
                    "; candidateStatus=" + item.CandidateStatus + "; candidate=" + item.Candidate +
                    "; diagnostic=" + item.Diagnostic);
            }
            text.AppendLine("DUPLICATE/AMBIGUITY ANALYSIS");
            foreach (var item in evidence.Steps.Where(item => item.CandidateStatus != "Candidate"))
                text.AppendLine("  objectid=" + item.ObjectId + "; status=" + item.CandidateStatus +
                    "; diagnostic=" + item.Diagnostic);
        }
    }
}
#endif
