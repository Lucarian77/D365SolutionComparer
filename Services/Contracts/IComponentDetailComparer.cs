using System.Collections.Generic;
using D365SolutionComparer.Models.ComponentDetails;
using D365SolutionComparer.Models.Membership;

namespace D365SolutionComparer.Services.Contracts
{
    internal interface IComponentDetailComparer
    {
        IReadOnlyList<ComponentDetailCompareResult> Compare(
            IReadOnlyList<MembershipCompareResult> membershipResults,
            ComponentDefinitionSnapshot source, ComponentDefinitionSnapshot target);
    }
}
