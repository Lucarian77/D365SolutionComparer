using System;

namespace D365SolutionComparer.Models.Membership
{
    /// <summary>
    /// Stable descriptive identity for a registered Solution Component Framework family.
    /// The environment-local object type code is retained for audit but is not part of the family key.
    /// </summary>
    public sealed class SolutionComponentDefinitionIdentity
    {
        public SolutionComponentDefinitionIdentity(int objectTypeCode, string name,
            string primaryEntityName = null)
        {
            if (string.IsNullOrWhiteSpace(name))
                throw new ArgumentException("A registered component definition name is required.", nameof(name));
            ObjectTypeCode = objectTypeCode;
            Name = name;
            PrimaryEntityName = primaryEntityName ?? string.Empty;
            SemanticKind = ComponentSemanticKinds.FromRegisteredDefinitionName(name);
        }

        public int ObjectTypeCode { get; }
        public string Name { get; }
        public string PrimaryEntityName { get; }
        public string SemanticKind { get; }
    }
}
