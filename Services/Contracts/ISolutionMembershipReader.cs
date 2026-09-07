using System;
using System.Threading;
using D365SolutionComparer.Infrastructure;
using D365SolutionComparer.Models.Identity;
using D365SolutionComparer.Models.Membership;
using Microsoft.Xrm.Sdk;

namespace D365SolutionComparer.Services.Contracts
{
    public interface ISolutionMembershipReader
    {
        /// <summary>
        /// Read a known-present solution using the service for its captured environment.
        /// Revalidate the selected solution; return SolutionAbsent if its Unique Name no longer exists.
        /// Return all raw records wrapped as unresolved identities; propagate errors/cancellation.
        /// Callers represent a verified absent solution using MembershipSnapshot.Absent without reading it.
        /// </summary>
        MembershipSnapshot Read(IOrganizationService service, SolutionIdentity solution,
            CancellationToken cancellationToken, Action<RetrievalProgress> progress = null);
    }
}
