using System.Threading;
using D365SolutionComparer.Models.Identity;
using D365SolutionComparer.Models.Membership;
using Microsoft.Xrm.Sdk;

namespace D365SolutionComparer.Services.Contracts
{
    public interface IComponentIdentityResolver
    {
        /// <summary>
        /// Resolve within the captured environment. Preserve the record and unknown type codes.
        /// Return Unsupported/Unresolved/Ambiguous explicitly; cancellation must propagate.
        /// Matching keys must be scoped by component type and any required parent identity.
        /// </summary>
        ComponentIdentity Resolve(IOrganizationService service, EnvironmentIdentity environment,
            SolutionComponentRecord component, CancellationToken cancellationToken);
    }
}
