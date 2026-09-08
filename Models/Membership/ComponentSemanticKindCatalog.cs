using System.Collections.Generic;
using System.Globalization;

namespace D365SolutionComparer.Models.Membership
{
    /// <summary>Stable semantic coverage groups independent of portable-key resolution success.</summary>
    public static class ComponentSemanticKinds
    {
        public const string Table = "table";
        public const string Column = "column";
        public const string Relationship = "relationship";
        public const string WebResource = "webresource";
        public const string Process = "process";
        public const string SecurityRole = "securityrole";
        public const string EnvironmentVariableDefinition = "environmentvariabledefinition";
        public const string ConnectionReference = "connectionreference";
        public const string AppModule = "appmodule";
        internal const string RegisteredDefinitionPrefix = "registered:solutioncomponentdefinition:";

        // Published componenttype choices assigned to built-in kinds other than Connection Reference.
        private static readonly HashSet<int> KnownBuiltInTypes = new HashSet<int>
        {
            1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 16, 17, 18,
            20, 21, 22, 23, 24, 25, 26, 29, 31, 32, 33, 34, 35, 36, 37, 38, 39,
            44, 45, 46, 47, 48, 49, 50, 52, 53, 55, 59, 60, 61, 62, 63, 64, 65, 66, 68,
            70, 71, 80, 90, 91, 92, 93, 95, 150, 151, 152, 153, 154, 155, 161, 162, 165, 166,
            201, 202, 203, 204, 205, 206, 207, 208, 210, 300, 371, 372, 380, 381,
            400, 401, 402, 430, 431, 432
        };

        internal static bool IsKnownBuiltInType(int componentType) => KnownBuiltInTypes.Contains(componentType);

        internal static string FromRegisteredDefinitionName(string name) =>
            RegisteredDefinitionPrefix + name.Trim().ToLowerInvariant();

        internal static bool IsRegisteredDefinitionKind(string semanticKind) =>
            semanticKind != null && semanticKind.StartsWith(RegisteredDefinitionPrefix,
                System.StringComparison.OrdinalIgnoreCase);

        internal static string FromRawComponentType(int componentType)
        {
            switch (componentType)
            {
                case 1: return Table;
                case 2: return Column;
                case 3:
                case 8:
                case 10:
                case 11:
                case 12:
                    return Relationship;
                case 20: return SecurityRole;
                case 29: return Process;
                case 61: return WebResource;
                case 80: return AppModule;
                case 380: return EnvironmentVariableDefinition;
                default:
                    return IsKnownBuiltInType(componentType)
                        ? "unsupported:componenttype:" + componentType.ToString(CultureInfo.InvariantCulture)
                        : null;
            }
        }

        internal static string FromCanonicalTypeKey(string componentTypeKey)
        {
            if (string.IsNullOrWhiteSpace(componentTypeKey) ||
                componentTypeKey.StartsWith("componenttype:", System.StringComparison.OrdinalIgnoreCase))
                return null;
            return componentTypeKey;
        }
    }
}
