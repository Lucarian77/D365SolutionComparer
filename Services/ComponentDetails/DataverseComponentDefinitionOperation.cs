using System;
using System.Threading;
using D365SolutionComparer.Infrastructure;
using D365SolutionComparer.Models.ComponentDetails;
using D365SolutionComparer.Models.Identity;
using D365SolutionComparer.Services.Membership;
using Microsoft.Xrm.Sdk;

namespace D365SolutionComparer.Services.ComponentDetails
{
    /// <summary>
    /// Production bulk path. Environment verification, membership retrieval, identity
    /// resolution, and definition retrieval share one operation context and metadata cache.
    /// </summary>
    internal sealed class DataverseComponentDefinitionOperation
    {
        private readonly DataverseSolutionMembershipReader membershipReader =
            new DataverseSolutionMembershipReader();
        private readonly DataverseComponentIdentityResolver identityResolver =
            new DataverseComponentIdentityResolver();
        private readonly DataverseComponentDefinitionReader definitionReader =
            new DataverseComponentDefinitionReader();

        public ComponentDefinitionSnapshot ReadAndResolve(IOrganizationService service,
            SolutionIdentity solution, CancellationToken cancellationToken,
            Action<RetrievalProgress> progress = null, DataverseRequestCounter requestCounter = null)
        {
            if (solution == null) throw new ArgumentNullException(nameof(solution));
            var context = new DataverseReadContext(service, solution.Environment,
                cancellationToken, requestCounter);
            var membership = membershipReader.Read(context, solution, cancellationToken, progress);
            var resolved = identityResolver.ResolveSnapshot(context, membership, cancellationToken);
            return definitionReader.Read(context, resolved, cancellationToken);
        }

        /// <summary>
        /// Live UI path. Membership, identity resolution, and definition retrieval share one
        /// verified environment context, request counter, and operation-scoped metadata cache.
        /// </summary>
        public ComponentDefinitionSnapshot ReadAndResolve(IOrganizationService service,
            string environmentDisplayName, string solutionUniqueName,
            CancellationToken cancellationToken, Action<MembershipOperationProgress> progress,
            DataverseRequestCounter requestCounter = null)
        {
            if (string.IsNullOrWhiteSpace(solutionUniqueName))
                throw new ArgumentException("A solution Unique Name is required.",
                    nameof(solutionUniqueName));

            Report(progress, MembershipOperationStage.ValidatingEnvironment,
                "Validating the Dataverse environment...");
            var context = new DataverseReadContext(service, environmentDisplayName,
                cancellationToken, requestCounter);
            Report(progress, MembershipOperationStage.ReadingMembership,
                "Reading solution membership...");
            var membership = membershipReader.Read(context, context.Environment,
                solutionUniqueName, cancellationToken, item =>
                {
                    Report(progress, MembershipOperationStage.ReadingMembership,
                        "Reading solution membership: " + item.RecordsRetrieved +
                        " record(s) from " + item.PagesRetrieved + " page(s)...",
                        item.PagesRetrieved, item.RecordsRetrieved);
                });
            if (membership.State == Models.Membership.MembershipSnapshotState.Complete)
                Report(progress, MembershipOperationStage.ResolvingIdentities,
                    "Resolving portable component identities...");
            var resolved = identityResolver.ResolveSnapshot(context, membership,
                cancellationToken);
            if (resolved.State == Models.Membership.MembershipSnapshotState.Complete)
                Report(progress, MembershipOperationStage.ReadingDefinitions,
                    "Reading component definitions...");
            var definitions = definitionReader.Read(context, resolved, cancellationToken);
            Report(progress, MembershipOperationStage.Completed,
                "Membership and component definition comparison data completed.");
            return definitions;
        }

        private static void Report(Action<MembershipOperationProgress> progress,
            MembershipOperationStage stage, string message, int pagesRetrieved = 0,
            int recordsRetrieved = 0)
        {
            progress?.Invoke(new MembershipOperationProgress(stage, message,
                pagesRetrieved, recordsRetrieved));
        }
    }
}
