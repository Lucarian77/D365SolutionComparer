using System;
using System.Threading;
using D365SolutionComparer.Infrastructure;
using D365SolutionComparer.Models.Identity;
using D365SolutionComparer.Models.Membership;
using Microsoft.Xrm.Sdk;

namespace D365SolutionComparer.Services.Membership
{
    /// <summary>Normal bulk path: read and resolve with one verified environment context.</summary>
    public sealed class DataverseSolutionMembershipOperation
    {
        private readonly DataverseSolutionMembershipReader reader = new DataverseSolutionMembershipReader();
        private readonly DataverseComponentIdentityResolver resolver = new DataverseComponentIdentityResolver();

        public MembershipSnapshot ReadAndResolve(IOrganizationService service, SolutionIdentity solution,
            CancellationToken cancellationToken, Action<RetrievalProgress> progress = null,
            DataverseRequestCounter requestCounter = null)
        {
            if (solution == null) throw new ArgumentNullException(nameof(solution));
            var context = new DataverseReadContext(service, solution.Environment, cancellationToken, requestCounter);
            var snapshot = reader.Read(context, solution, cancellationToken, progress);
            return resolver.ResolveSnapshot(context, snapshot, cancellationToken);
        }

        public MembershipSnapshot ReadAndResolve(IOrganizationService service, EnvironmentIdentity environment,
            string solutionUniqueName, CancellationToken cancellationToken,
            Action<RetrievalProgress> progress = null, DataverseRequestCounter requestCounter = null)
        {
            if (environment == null) throw new ArgumentNullException(nameof(environment));
            var context = new DataverseReadContext(service, environment, cancellationToken, requestCounter);
            var snapshot = reader.Read(context, environment, solutionUniqueName, cancellationToken, progress);
            return resolver.ResolveSnapshot(context, snapshot, cancellationToken);
        }

        /// <summary>
        /// Live UI path. It captures and verifies the service organization once, then reuses that context
        /// for lookup, paging, and bulk identity resolution.
        /// </summary>
        public MembershipSnapshot ReadAndResolve(IOrganizationService service, string environmentDisplayName,
            string solutionUniqueName, CancellationToken cancellationToken,
            Action<MembershipOperationProgress> progress, DataverseRequestCounter requestCounter = null)
        {
            if (string.IsNullOrWhiteSpace(solutionUniqueName))
                throw new ArgumentException("A solution Unique Name is required.", nameof(solutionUniqueName));
            Report(progress, MembershipOperationStage.ValidatingEnvironment,
                "Validating the Dataverse environment...");
            var context = new DataverseReadContext(service, environmentDisplayName, cancellationToken, requestCounter);
            Report(progress, MembershipOperationStage.ReadingMembership,
                "Reading solution membership...");
            var snapshot = reader.Read(context, context.Environment, solutionUniqueName, cancellationToken, item =>
            {
                Report(progress, MembershipOperationStage.ReadingMembership,
                    "Reading solution membership: " + item.RecordsRetrieved + " record(s) from " +
                    item.PagesRetrieved + " page(s)...", item.PagesRetrieved, item.RecordsRetrieved);
            });
            if (snapshot.State == Models.Membership.MembershipSnapshotState.Complete)
                Report(progress, MembershipOperationStage.ResolvingIdentities,
                    "Resolving portable component identities...");
            var resolved = resolver.ResolveSnapshot(context, snapshot, cancellationToken);
            Report(progress, MembershipOperationStage.Completed,
                "Membership retrieval and identity resolution completed.");
            return resolved;
        }

        private static void Report(Action<MembershipOperationProgress> progress, MembershipOperationStage stage,
            string message, int pagesRetrieved = 0, int recordsRetrieved = 0)
        {
            progress?.Invoke(new MembershipOperationProgress(stage, message, pagesRetrieved, recordsRetrieved));
        }
    }
}
