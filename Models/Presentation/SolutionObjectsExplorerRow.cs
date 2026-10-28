using System;
using System.Collections.Generic;
using System.Linq;
using D365SolutionComparer.Models.Identity;
using D365SolutionComparer.Models.ComponentDetails;
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

        // Compact primary presentation only. The existing technical fields above stay unchanged.
        internal string PrimaryName(DisplayEnvironmentSide side, bool friendly)
        {
            var identity = side == DisplayEnvironmentSide.Source ? Original.AuthoritativeRow.SourceIdentity : Original.AuthoritativeRow.TargetIdentity;
            if (identity == null) return null;
            var evidence = side == DisplayEnvironmentSide.Source ? SourceDisplay : TargetDisplay;
            // Dashboard maker Name is the local systemform.name; technical uniquename/identity stay intact.
            if (!friendly && evidence?.MakerObject?.Kind == MakerObjectKind.Dashboard && evidence.MakerObject.IsFor(identity))
                return Readable(evidence.MakerObject.FriendlyName);
            if (!friendly && identity.Record.ComponentType == 31 && evidence?.BackingCorrelationStatus == DisplayBackingCorrelationStatus.Unique &&
                evidence.FriendlyName?.DisplayEvidenceSource == "NormalMembershipDiagnostic.report.name")
                return Readable(evidence.FriendlyName);
            if (!friendly && identity.Record.ComponentType == 511 && identity.SemanticKind == ComponentSemanticKinds.TeamTemplate &&
                evidence?.BackingCorrelationStatus == DisplayBackingCorrelationStatus.Unique &&
                evidence.FriendlyName?.DisplayEvidenceSource == "NormalMembershipDiagnostic.teamtemplate.teamtemplatename")
                return Readable(evidence.FriendlyName);
            if (evidence?.MakerObject?.Kind == MakerObjectKind.SecuredMaskingRule &&
                evidence.MakerObject.IsFor(identity) && evidence.MakerObject.Correlation == DisplayBackingCorrelationStatus.Unique)
                return Readable(evidence.MakerObject.FriendlyName);
            if (!friendly && evidence?.MakerObject?.Kind == MakerObjectKind.Process && evidence.MakerObject.IsFor(identity) &&
                (evidence.MakerObject.Correlation == DisplayBackingCorrelationStatus.Unique || evidence.MakerObject.Correlation == DisplayBackingCorrelationStatus.VerifiedDefinition))
                return Readable(evidence.MakerObject.FriendlyName);
            if (!friendly && identity.Record.ComponentType == 36 && evidence?.BackingCorrelationStatus == DisplayBackingCorrelationStatus.Unique &&
                evidence.FriendlyName?.DisplayEvidenceSource == "NormalMembershipDiagnostic.template.title")
                return Readable(evidence.FriendlyName);
            if (friendly)
            {
                string display = Readable(evidence?.FriendlyName);
                if (display != null) return display;
            }
            return Readable(evidence?.InternalName) ?? Readable(VerifiedIdentityName(identity));
        }
        internal string NameBasis(DisplayEnvironmentSide side)
        {
            var identity = side == DisplayEnvironmentSide.Source ? Original.AuthoritativeRow.SourceIdentity : Original.AuthoritativeRow.TargetIdentity;
            if (identity == null) return "No local observation; environment absence is not established";
            var evidence = side == DisplayEnvironmentSide.Source ? SourceDisplay : TargetDisplay;
            if (Readable(evidence?.FriendlyName) != null) return "Retained friendly/display name";
            if (Readable(evidence?.InternalName) != null) return "Internal/logical name fallback; friendly name: " + Field(evidence?.FriendlyName);
            if (Readable(VerifiedIdentityName(identity)) != null) return "Verified identity-name fallback; friendly name: " + Field(evidence?.FriendlyName);
            return "Name unavailable; friendly name: " + Field(evidence?.FriendlyName) + "; internal name: " + Field(evidence?.InternalName);
        }
        private static ComponentDisplayValue VerifiedIdentityName(ComponentIdentity identity)
        {
            string app, setting;
            if (identity?.SemanticKind == ComponentSemanticKinds.AppSetting && !MembershipGridRow.AppSettingIdentityNames(identity, out app, out setting)) return null;
            return MembershipGridRow.RefinedComponentName(null, identity);
        }
        private static string Readable(ComponentDisplayValue value)
        {
            Guid local;
            return value?.Availability == DisplayAvailability.Available && !string.IsNullOrWhiteSpace(value.Value) && !Guid.TryParse(value.Value, out local)
                ? MembershipGridRow.EscapeForGrid(value.Value) : null;
        }
        internal string PrimaryType
        {
            get
            {
                int raw = Original.AuthoritativeRow.TargetIdentity?.Record.ComponentType ?? Original.AuthoritativeRow.SourceIdentity?.Record.ComponentType ?? 0;
                if (raw == 29) return "Process/workflow (unclassified)";
                if (raw == 60) return "Form/dashboard (unclassified)";
                if (raw == 300) return "Canvas App/library (unclassified)";
                return Type;
            }
        }
        internal string MembershipPresentation => Original.AuthoritativeRow.Presence == MembershipPresence.PresentInBoth ? "Present in Both" :
            Original.AuthoritativeRow.Presence == MembershipPresence.OnlyInSource ? "Source only" :
            Original.AuthoritativeRow.Presence == MembershipPresence.OnlyInTarget ? "Target only" : "Needs review — " + Original.MembershipStatus;
        internal string DefinitionPresentation => Original.DefinitionStatus == "Match" ? "Match (compared properties)" :
            string.IsNullOrWhiteSpace(Original.DefinitionStatusText) ? "Not compared — result not supplied" : Original.DefinitionStatusText;
        internal IReadOnlyList<ComponentPropertyPresentation> RetainedComparedProperties => Original.AuthoritativeRow.DefinitionDetail?.Properties ?? new ComponentPropertyPresentation[0];
        internal IReadOnlyList<ComponentPropertyPresentation> OrderedComparedProperties => OrderProperties(RetainedComparedProperties);
        internal string ComparedPropertiesExplanation => PropertyExplanation(Original.DefinitionStatus,
            RetainedComparedProperties, Original.AuthoritativeRow.DefinitionDetail != null, Original.ChangedProperties) +
            (SourceDisplay?.ArchivalInformation == null ? "" : Environment.NewLine + SourceDisplay.ArchivalInformation.Text) +
            (TargetDisplay?.ArchivalInformation == null ? "" : Environment.NewLine + TargetDisplay.ArchivalInformation.Text);
        internal static IReadOnlyList<ComponentPropertyPresentation> OrderProperties(IEnumerable<ComponentPropertyPresentation> properties) =>
            (properties ?? Enumerable.Empty<ComponentPropertyPresentation>()).OrderByDescending(p => p.Changed && p.Comparison == "Different").ToList().AsReadOnly();
        internal static string PropertyComparison(ComponentPropertyPresentation property) =>
            property.Comparison == "Match" && (property.SourceValue == null || property.TargetValue == null) ? "Unavailable" :
            string.IsNullOrEmpty(property.Comparison) ? "Not compared" : property.Comparison;
        internal static string PropertyExplanation(string status, IEnumerable<ComponentPropertyPresentation> properties,
            bool retained, string reportedChanges = null)
        {
            var values = (properties ?? Enumerable.Empty<ComponentPropertyPresentation>()).ToList();
            var changed = values.Where(p => p.Changed && p.Comparison == "Different").Select(p => p.PropertyName).ToList();
            string summary = status == "Different" ? changed.Count == 0 ?
                "Definition: Different; difference-producing property details are unavailable. " :
                "Definition: Different. Changed properties shown first: " + string.Join(", ", changed) + ". " :
                "Definition: " + (string.IsNullOrEmpty(status) ? "Not supplied" : status) + ". ";
            if (status == "Different" && changed.Count == 0 && !string.IsNullOrWhiteSpace(reportedChanges))
                summary += "Reported changed properties: " + reportedChanges + "; their values are not retained here. ";
            summary += !retained ? "Compared properties not retained / not compared. " :
                values.Count == 0 ? "No retained compared properties. " :
                "Retained compared properties; full values remain in tooltips and Definition Details. ";
            if (values.Any(p => p.SourceValue == null || p.TargetValue == null))
                summary += "Some value evidence is unavailable; unavailable values do not establish equality or absence. ";
            return summary + "The authoritative definition result is unchanged; Match covers compared properties only.";
        }
        internal static string PropertyValue(string value) => value == null ? "[No value supplied; availability not established]" :
            value.Length == 0 ? "[Blank value supplied]" : value;


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
