using System;
using System.Collections.Generic;
using System.Linq;
using D365SolutionComparer.Models.Identity;
using D365SolutionComparer.Models.Membership;

namespace D365SolutionComparer.Models.Presentation
{
    /// <summary>Six-column display; does not resolve, pair or aggregate comparison results.</summary>
    internal sealed class SolutionObjectsExplorerRow
    {
        internal SolutionObjectsExplorerRow(MembershipGridRow original, IEnumerable<int> source, IEnumerable<int> target)
        {
            Original = original;
            SourceOccurrences = source.ToList().AsReadOnly(); TargetOccurrences = target.ToList().AsReadOnly();
            SourceDisplay = Bound(original.SourceDisplayEvidence, original.AuthoritativeRow.SourceIdentity, DisplayEnvironmentSide.Source);
            TargetDisplay = Bound(original.TargetDisplayEvidence, original.AuthoritativeRow.TargetIdentity, DisplayEnvironmentSide.Target);
            DisplayName = Sides(Field(SourceDisplay?.FriendlyName), Field(TargetDisplay?.FriendlyName));
            Name = Sides(InternalName(SourceDisplay, original.AuthoritativeRow.SourceIdentity),
                InternalName(TargetDisplay, original.AuthoritativeRow.TargetIdentity));
            Type = (original.AuthoritativeRow.SourceIdentity == null || SourceDisplay?.ComponentCategory?.Availability == DisplayAvailability.Available) &&
                (original.AuthoritativeRow.TargetIdentity == null || TargetDisplay?.ComponentCategory?.Availability == DisplayAvailability.Available)
                ? original.TypeText : original.ComponentKind;
        }
        internal MembershipGridRow Original { get; }
        internal int OriginalRowIndex => Original.OriginalRowIndex;
        internal IReadOnlyList<int> SourceOccurrences { get; }
        internal IReadOnlyList<int> TargetOccurrences { get; }
        internal ComponentSideDisplayEvidence SourceDisplay { get; }
        internal ComponentSideDisplayEvidence TargetDisplay { get; }
        public string DisplayName { get; }
        public string Name { get; }
        public string Type { get; }
        public string Source => Original.SourcePresence;
        public string Target => Original.TargetPresence;
        public string ComparisonStatus => "Membership: " + Original.MembershipStatus + Environment.NewLine +
            "Definition: " + (string.IsNullOrEmpty(Original.DefinitionStatusText) ? "[Result not supplied]" : Original.DefinitionStatusText);

        private string Sides(string source, string target)
        {
            bool left = Original.AuthoritativeRow.SourceIdentity != null, right = Original.AuthoritativeRow.TargetIdentity != null;
            if (left && right && StringComparer.Ordinal.Equals(source, target)) return source;
            return (left ? "Source: " + source : "Source: [No local row]") + Environment.NewLine +
                (right ? "Target: " + target : "Target: [No local row]");
        }
        private static ComponentSideDisplayEvidence Bound(ComponentSideDisplayEvidence evidence,
            ComponentIdentity identity, DisplayEnvironmentSide side) => identity != null && evidence != null &&
            evidence.EnvironmentSide == side && evidence.SolutionComponentId == identity.Record.SolutionComponentId &&
            evidence.LocalObjectId == identity.Record.ObjectId && evidence.RawComponentType == identity.Record.ComponentType ? evidence : null;

        private static string InternalName(ComponentSideDisplayEvidence side, ComponentIdentity identity)
        {
            if (side?.InternalName?.Availability == DisplayAvailability.Available && !string.IsNullOrWhiteSpace(side.InternalName.Value))
                return Field(side.InternalName);
            // Formatting of existing resolved contracts only; never invent a friendly metadata value.
            var refined = MembershipGridRow.RefinedComponentName(null, identity);
            if (refined != null) return "Identity: " + MembershipGridRow.EscapeForGrid(refined.Value);
            return identity?.Status == IdentityResolutionStatus.Resolved
                ? "Identity: " + MembershipGridRow.EscapeForGrid(identity.ComparisonKey) : Field(side?.InternalName);
        }
        internal static string Field(ComponentDisplayValue value) => value == null ? "[Not retrieved]" :
            value.Availability == DisplayAvailability.Available ? value.Value == null ? "[Retrieved null]" :
                value.Value.Length == 0 ? "[Retrieved empty]" : MembershipGridRow.EscapeForGrid(value.Value) :
            "[" + value.Availability + "]";
        internal string Label(DisplayEnvironmentSide side)
        {
            var evidence = side == DisplayEnvironmentSide.Source ? SourceDisplay : TargetDisplay;
            var identity = side == DisplayEnvironmentSide.Source ? Original.AuthoritativeRow.SourceIdentity : Original.AuthoritativeRow.TargetIdentity;
            if (evidence?.FriendlyName?.Availability == DisplayAvailability.Available && !string.IsNullOrWhiteSpace(evidence.FriendlyName.Value))
                return Field(evidence.FriendlyName);
            return InternalName(evidence, identity);
        }
        internal string ObjectField(string column)
        {
            var left = SourceDisplay?.ObjectMetadata; var right = TargetDisplay?.ObjectMetadata;
            switch (column)
            {
                case "Managed": return Metadata(left?.Managed, right?.Managed);
                case "Customized": return Metadata(left?.Customized, right?.Customized);
                case "LastModified": return Metadata(left?.LastModified, right?.LastModified);
                case "Owner": return Metadata(left?.Owner, right?.Owner);
                default: return Metadata(left?.ObjectStatus, right?.ObjectStatus);
            }
        }
        private string Metadata<T>(ComponentObjectMetadataValue<T> left, ComponentObjectMetadataValue<T> right) =>
            Sides(ObjectValue(left), ObjectValue(right));
        private static string ObjectValue<T>(ComponentObjectMetadataValue<T> field) => field == null ? "[Not retrieved]" :
            field.Availability != DisplayAvailability.Available ? "[" + field.Availability + "]" :
            MembershipGridRow.EscapeForGrid(MembershipGridRow.FormatObjectValue(field)) ?? "[Retrieved null]";
    }
}
