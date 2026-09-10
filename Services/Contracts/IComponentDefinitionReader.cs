using System.Threading;
using D365SolutionComparer.Infrastructure;
using D365SolutionComparer.Models.ComponentDetails;
using D365SolutionComparer.Models.Membership;
using Microsoft.Xrm.Sdk;

namespace D365SolutionComparer.Services.Contracts
{
    internal interface IComponentDefinitionReader
    {
        ComponentDefinitionSnapshot Read(IOrganizationService service, MembershipSnapshot membership,
            CancellationToken cancellationToken, DataverseRequestCounter requestCounter = null);
    }
}
