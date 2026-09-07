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
    }
}
