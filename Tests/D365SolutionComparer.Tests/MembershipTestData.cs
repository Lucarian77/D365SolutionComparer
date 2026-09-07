using System;
using D365SolutionComparer.Models.Identity;
using D365SolutionComparer.Models.Membership;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace D365SolutionComparer.Tests
{
    internal static class MembershipTestData
    {
        public static SolutionIdentity Solution(string uniqueName = "sample") =>
            new SolutionIdentity(new EnvironmentIdentity(Guid.NewGuid(), "Test environment"), Guid.NewGuid(), uniqueName);

        public static FakeOrganizationService Service(SolutionIdentity solution, Func<QueryExpression, EntityCollection> query = null)
        {
            return new FakeOrganizationService
            {
                RetrievePage = query,
                ExecuteRequest = request => request is WhoAmIRequest ? WhoAmI(solution.Environment.OrganizationId) : throw new NotSupportedException()
            };
        }

        public static WhoAmIResponse WhoAmI(Guid organizationId)
        {
            var response = new WhoAmIResponse();
            response.Results["OrganizationId"] = organizationId;
            return response;
        }

        public static Entity SolutionRow(SolutionIdentity solution) => new Entity("solution", solution.SolutionId)
        {
            ["uniquename"] = solution.UniqueName
        };

        public static Entity ComponentRow(SolutionIdentity solution, int type = 61) => new Entity("solutioncomponent", Guid.NewGuid())
        {
            ["solutionid"] = new EntityReference("solution", solution.SolutionId),
            ["componenttype"] = new OptionSetValue(type),
            ["objectid"] = Guid.NewGuid(),
            ["rootcomponentbehavior"] = new OptionSetValue(2),
            ["rootsolutioncomponentid"] = Guid.NewGuid(),
            ["ismetadata"] = false
        };

        public static EntityCollection Rows(params Entity[] rows)
        {
            var result = new EntityCollection();
            result.Entities.AddRange(rows);
            return result;
        }

        public static ComponentIdentity Identity(string key, int type = 61,
            IdentityResolutionStatus status = IdentityResolutionStatus.Resolved, string kind = null) =>
            new ComponentIdentity(new SolutionComponentRecord(Guid.NewGuid(), type, Guid.NewGuid()), status,
                status == IdentityResolutionStatus.Resolved ? key : null, componentTypeKey: kind);

        public static MembershipSnapshot Snapshot(params ComponentIdentity[] items) =>
            MembershipSnapshot.Complete(Solution(), items, DateTimeOffset.UtcNow);
    }
}
