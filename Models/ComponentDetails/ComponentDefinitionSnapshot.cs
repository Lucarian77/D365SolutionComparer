using System;
using System.Collections.Generic;
using System.Linq;
using D365SolutionComparer.Models.Membership;

namespace D365SolutionComparer.Models.ComponentDetails
{
    internal sealed class ComponentDefinitionSnapshot
    {
        public ComponentDefinitionSnapshot(MembershipSnapshot membership,
            IEnumerable<ComponentDefinition> definitions, string diagnostic = null)
        {
            Membership = membership ?? throw new ArgumentNullException(nameof(membership));
            var copy = (definitions ?? throw new ArgumentNullException(nameof(definitions))).ToList();
            if (copy.Any(item => item == null))
                throw new ArgumentException("Null component definitions are not allowed.", nameof(definitions));
            Definitions = copy.AsReadOnly();
            Diagnostic = diagnostic ?? string.Empty;
        }

        public MembershipSnapshot Membership { get; }
        public IReadOnlyList<ComponentDefinition> Definitions { get; }
        public string Diagnostic { get; }
    }
}
