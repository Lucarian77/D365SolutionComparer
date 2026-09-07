using System;

namespace D365SolutionComparer.Models.Membership
{
    /// <summary>Raw solutioncomponent data. Unknown type/behavior values and missing optional IDs are retained.</summary>
    public sealed class SolutionComponentRecord
    {
        public SolutionComponentRecord(Guid solutionComponentId, int componentType, Guid? objectId,
            int? rootComponentBehavior = null, Guid? rootSolutionComponentId = null, bool? isMetadata = null)
        {
            if (solutionComponentId == Guid.Empty) throw new ArgumentException("A membership record ID is required.", nameof(solutionComponentId));
            SolutionComponentId = solutionComponentId;
            ComponentType = componentType;
            ObjectId = objectId;
            RootComponentBehavior = rootComponentBehavior;
            RootSolutionComponentId = rootSolutionComponentId;
            IsMetadata = isMetadata;
        }

        public Guid SolutionComponentId { get; }
        public int ComponentType { get; }
        public Guid? ObjectId { get; }
        public int? RootComponentBehavior { get; }
        public Guid? RootSolutionComponentId { get; }
        public bool? IsMetadata { get; }
    }
}
