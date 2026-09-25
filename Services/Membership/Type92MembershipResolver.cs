using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using D365SolutionComparer.Models.Membership;

namespace D365SolutionComparer.Services.Membership
{
    /// <summary>Production SDK Message Processing Step identity resolution using batched read-only evidence.</summary>
    internal sealed class Type92MembershipResolver
    {
        internal ComponentIdentity Resolve(DataverseReadContext context, SolutionComponentRecord record,
            CancellationToken cancellationToken)
        {
            if (context == null) throw new ArgumentNullException(nameof(context));
            if (record == null) throw new ArgumentNullException(nameof(record));
            var evidence = new Type92MembershipEvidenceReader().CaptureSingle(context.Service, record,
                cancellationToken, context.MetadataCache);
            context.MetadataCache.Type92Evidence = evidence;
            return Resolve(record, evidence.Steps.ToDictionary(item => item.ObjectId));
        }

        internal MembershipSnapshot Apply(DataverseReadContext context, MembershipSnapshot snapshot,
            CancellationToken cancellationToken)
        {
            if (context == null) throw new ArgumentNullException(nameof(context));
            if (snapshot == null) throw new ArgumentNullException(nameof(snapshot));
            cancellationToken.ThrowIfCancellationRequested();
            if (snapshot.State != MembershipSnapshotState.Complete ||
                !snapshot.Components.Any(item => item.Record.ComponentType == 92)) return snapshot;

            // The instrumented operation service retains the shared request counter and WhoAmI context.
            var evidence = new Type92MembershipEvidenceReader().Capture(context.Service, snapshot,
                string.Empty, cancellationToken, metadataCache: context.MetadataCache);
            context.MetadataCache.Type92Evidence = evidence;
            var byObjectId = evidence.Steps.ToDictionary(item => item.ObjectId);
            var results = snapshot.Components.Select(item => item.Record.ComponentType == 92
                ? Resolve(item.Record, byObjectId) : item).ToList();
            cancellationToken.ThrowIfCancellationRequested();
            return MembershipSnapshot.Complete(snapshot.Solution, results, snapshot.CapturedAt);
        }

        private static ComponentIdentity Resolve(SolutionComponentRecord record,
            IDictionary<Guid, Type92StepEvidence> byObjectId)
        {
            Type92StepEvidence step;
            if (!record.ObjectId.HasValue || record.ObjectId.Value == Guid.Empty ||
                !byObjectId.TryGetValue(record.ObjectId.Value, out step))
                return new ComponentIdentity(record, IdentityResolutionStatus.Unresolved,
                    diagnostic: "Type 92 membership has no correlated backing step.",
                    componentTypeKey: ComponentSemanticKinds.SdkMessageProcessingStep,
                    semanticKind: ComponentSemanticKinds.SdkMessageProcessingStep);

            var audit = Evidence(step);
            if (step.Candidate != null)
            {
                var key = step.Candidate.PortableKey;
                if (step.CandidateStatus == "AmbiguousCandidate")
                    return new ComponentIdentity(record, IdentityResolutionStatus.Ambiguous,
                        diagnostic: "Different backing SDK steps share this portable candidate identity.",
                        componentTypeKey: ComponentSemanticKinds.SdkMessageProcessingStep,
                        semanticKind: ComponentSemanticKinds.SdkMessageProcessingStep,
                        diagnosticEvidence: audit, blockerPortableIdentity: key,
                        blockerScope: ResolutionBlockerScope.PortableIdentity);
                return new ComponentIdentity(record, IdentityResolutionStatus.Resolved, key,
                    diagnostic: "SDK Message Processing Step portable identity resolved; definition comparison is unavailable.",
                    componentTypeKey: ComponentSemanticKinds.SdkMessageProcessingStep,
                    semanticKind: ComponentSemanticKinds.SdkMessageProcessingStep,
                    diagnosticEvidence: audit);
            }

            var ambiguous = step.Correlation == "DuplicateReturnedRow" ||
                step.Correlation == "ConflictingPrimaryKey" ||
                step.CandidateStatus == "AmbiguousHandler";
            return new ComponentIdentity(record, ambiguous ? IdentityResolutionStatus.Ambiguous :
                    IdentityResolutionStatus.Unresolved,
                diagnostic: "Type 92 identity evidence is incomplete: " + step.CandidateStatus + ".",
                componentTypeKey: ComponentSemanticKinds.SdkMessageProcessingStep,
                semanticKind: ComponentSemanticKinds.SdkMessageProcessingStep,
                diagnosticEvidence: audit);
        }

        private static IEnumerable<string> Evidence(Type92StepEvidence step) => new[]
        {
            "objectid=" + step.ObjectId,
            "stepid=" + step.StepId,
            "rawMembershipCount=" + step.MembershipCount,
            "plugintypeid=" + step.PluginTypeId,
            "plugintypeexportkey=" + (step.ExportKey ?? "(blank)"),
            "typename=" + (step.TypeName ?? "(blank)"),
            "assemblyIdentity=" + (step.AssemblyIdentity ?? "(unavailable)"),
            "handlerSemanticIdentity=" + (step.HandlerSemanticIdentity ?? "(unavailable)"),
            "sdkmessage=" + (step.MessageName ?? "(unavailable)"),
            "filterStatus=" + step.FilterStatus,
            "primaryScope=" + (step.PrimaryScope ?? "(blank)"),
            "secondaryScope=" + (step.SecondaryScope ?? "(blank)"),
            "semanticScope=" + (step.SemanticScope ?? "(unavailable)"),
            "stage=" + step.Stage,
            "mode=" + step.Mode,
            "supporteddeployment=" + step.Deployment,
            "rank=" + step.Rank,
            "filteringAttributes=" + step.FilteringCanonical,
            "configurationPresent=" + step.ConfigurationPresent,
            "configurationLength=" + step.ConfigurationLength,
            "state=" + step.State,
            "status=" + step.Status,
            "ismanaged=" + step.IsManaged,
            "componentstate=" + step.ComponentState,
            "candidateStatus=" + step.CandidateStatus
        };
    }
}
