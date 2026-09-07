using System.Collections.Generic;
using D365SolutionComparer.Models.Membership;

namespace D365SolutionComparer.Services.Contracts
{
    public interface ISolutionMembershipComparer
    {
        /// <summary>
        /// Compare the same Unique Name (OrdinalIgnoreCase) across captured snapshots, including absent sides.
        /// Match only resolved, type-scoped identities. Do not choose the first duplicate key.
        /// Unknown identities or unavailable opposite inventories cannot establish absence.
        /// CompleteResolvedInventory evidence requires complete retrieval AND unambiguous resolution
        /// of the relevant opposite inventory. Otherwise unmatched records remain Indeterminate.
        /// A verified absent solution supplies OppositeSolutionAbsent evidence for resolved records.
        /// Retain unsupported/unresolved/ambiguous records as Indeterminate even in a one-sided solution.
        /// Both solutions absent yield no component rows; callers retain the solution-level states.
        /// </summary>
        IReadOnlyList<MembershipCompareResult> Compare(MembershipSnapshot source, MembershipSnapshot target);
    }
}
