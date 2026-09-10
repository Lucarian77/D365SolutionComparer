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
    /// Future production bulk path. Environment verification, membership retrieval, identity
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
    }
}
