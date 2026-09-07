using System;

namespace D365SolutionComparer.Models.Identity
{
    /// <summary>Organization identity captured with data; never stores credentials or a service client.</summary>
    public sealed class EnvironmentIdentity
    {
        public EnvironmentIdentity(Guid organizationId, string displayName)
        {
            if (organizationId == Guid.Empty) throw new ArgumentException("An organization ID is required.", nameof(organizationId));
            OrganizationId = organizationId;
            DisplayName = displayName ?? string.Empty;
        }

        public Guid OrganizationId { get; }
        public string DisplayName { get; }
    }
}
