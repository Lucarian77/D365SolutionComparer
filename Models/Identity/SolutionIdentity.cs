using System;

namespace D365SolutionComparer.Models.Identity
{
    /// <summary>The GUID is local to Environment. UniqueName remains the solution comparison key.</summary>
    public sealed class SolutionIdentity
    {
        public SolutionIdentity(EnvironmentIdentity environment, Guid solutionId, string uniqueName)
        {
            Environment = environment ?? throw new ArgumentNullException(nameof(environment));
            if (solutionId == Guid.Empty) throw new ArgumentException("A solution ID is required.", nameof(solutionId));
            if (string.IsNullOrWhiteSpace(uniqueName)) throw new ArgumentException("A unique name is required.", nameof(uniqueName));
            SolutionId = solutionId;
            UniqueName = uniqueName;
        }

        public EnvironmentIdentity Environment { get; }
        public Guid SolutionId { get; }
        public string UniqueName { get; }
    }
}
