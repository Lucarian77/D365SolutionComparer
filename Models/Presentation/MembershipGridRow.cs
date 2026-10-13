using System;
using System.Collections.Generic;
using System.Globalization;
using D365SolutionComparer.Models.Membership;

namespace D365SolutionComparer.Models.Presentation
{
    /// <summary>Display adapter only; the original comparison remains authoritative.</summary>
    internal sealed class MembershipGridRow
    {
        private const string BothObjectNotRetrieved = "Both: [Not retrieved]";
        private const string ObjectStatusMeaning = "Component object status, not membership or definition comparison.";
        private static readonly string ObjectNotRetrievedToolTip = "[Not retrieved]" + Environment.NewLine +
            "Source: NotRetrieved; requested: False; provenance: " + Environment.NewLine +
            "Target: NotRetrieved; requested: False; provenance: ";
        private static readonly string ObjectStatusNotRetrievedToolTip = ObjectStatusMeaning + Environment.NewLine + ObjectNotRetrievedToolTip;
        private readonly ComponentDisplayValue sourceComponentName;
        private readonly ComponentDisplayValue targetComponentName;
        private MembershipGridRow(int index, MembershipResultRow row, ComponentDisplayEvidence evidence,
            string unavailableReason, bool refinedPresentation = false)
        {
            OriginalRowIndex = index;
            AuthoritativeRow = row;
            SourceDisplayEvidence = evidence?.SourceDisplayEvidence;
            TargetDisplayEvidence = evidence?.TargetDisplayEvidence;
            DisplayContextDiagnostic = unavailableReason;
            ComponentCategoryText = RenderBothSides(SourceDisplayEvidence, TargetDisplayEvidence,
                SourceDisplayEvidence?.ComponentCategory, TargetDisplayEvidence?.ComponentCategory);
            FriendlyNameText = RenderBothSides(SourceDisplayEvidence, TargetDisplayEvidence,
                SourceDisplayEvidence?.FriendlyName, TargetDisplayEvidence?.FriendlyName);
            InternalNameText = RenderBothSides(SourceDisplayEvidence, TargetDisplayEvidence,
                SourceDisplayEvidence?.InternalName, TargetDisplayEvidence?.InternalName);
            ParentContextText = RenderBothSides(SourceDisplayEvidence, TargetDisplayEvidence,
                SourceDisplayEvidence?.ParentDisplayContext, TargetDisplayEvidence?.ParentDisplayContext);
            ComponentCategoryToolTip = ToolTip(ComponentCategoryText, SourceDisplayEvidence?.ComponentCategory, TargetDisplayEvidence?.ComponentCategory);
            FriendlyNameToolTip = ToolTip(FriendlyNameText, SourceDisplayEvidence?.FriendlyName, TargetDisplayEvidence?.FriendlyName);
            InternalNameToolTip = ToolTip(InternalNameText, SourceDisplayEvidence?.InternalName, TargetDisplayEvidence?.InternalName);
            ParentContextToolTip = ToolTip(ParentContextText, SourceDisplayEvidence?.ParentDisplayContext, TargetDisplayEvidence?.ParentDisplayContext);
            var sourceName = ComponentName(SourceDisplayEvidence, row.SourceIdentity);
            var targetName = ComponentName(TargetDisplayEvidence, row.TargetIdentity);
            if (refinedPresentation)
            {
                sourceName = RefinedComponentName(SourceDisplayEvidence, row.SourceIdentity) ?? sourceName;
                targetName = RefinedComponentName(TargetDisplayEvidence, row.TargetIdentity) ?? targetName;
                if (row.SourceIdentity?.SemanticKind == ComponentSemanticKinds.AppSetting &&
                    row.TargetIdentity?.SemanticKind == ComponentSemanticKinds.AppSetting &&
                    row.SourceIdentity.Status == IdentityResolutionStatus.Resolved && row.TargetIdentity.Status == IdentityResolutionStatus.Resolved &&
                    !string.Equals(row.SourceIdentity.ComparisonKey, row.TargetIdentity.ComparisonKey, StringComparison.OrdinalIgnoreCase))
                {
                    sourceName = QualifyName(sourceName, new ComponentDisplayValue(DisplayAvailability.Available,
                        row.SourceIdentity.ComparisonKey, "Original resolved identity"));
                    targetName = QualifyName(targetName, new ComponentDisplayValue(DisplayAvailability.Available,
                        row.TargetIdentity.ComparisonKey, "Original resolved identity"));
                }
            }
            if (HasName(SourceDisplayEvidence?.InternalName) && HasName(TargetDisplayEvidence?.InternalName) &&
                !string.Equals(SourceDisplayEvidence.InternalName.Value, TargetDisplayEvidence.InternalName.Value, StringComparison.Ordinal))
            {
                sourceName = QualifyName(sourceName, SourceDisplayEvidence.InternalName);
                targetName = QualifyName(targetName, TargetDisplayEvidence.InternalName);
            }
            ComponentText = WithoutBothPrefix(RenderBothSides(SourceDisplayEvidence, TargetDisplayEvidence, sourceName, targetName));
            sourceComponentName = sourceName;
            targetComponentName = targetName;
            var sourceObject = SourceDisplayEvidence?.ObjectMetadata;
            var targetObject = TargetDisplayEvidence?.ObjectMetadata;
            ManagedText = RenderObject(sourceObject?.Managed, targetObject?.Managed);
            CustomizedText = RenderObject(sourceObject?.Customized, targetObject?.Customized);
            LastModifiedText = RenderObject(sourceObject?.LastModified, targetObject?.LastModified);
            OwnerText = RenderObject(sourceObject?.Owner, targetObject?.Owner);
            ObjectStatusText = RenderObject(sourceObject?.ObjectStatus, targetObject?.ObjectStatus);
            ManagedToolTip = ObjectToolTip(ManagedText, sourceObject?.Managed, targetObject?.Managed);
            CustomizedToolTip = ObjectToolTip(CustomizedText, sourceObject?.Customized, targetObject?.Customized);
            LastModifiedToolTip = ObjectToolTip(LastModifiedText, sourceObject?.LastModified, targetObject?.LastModified);
            OwnerToolTip = ObjectToolTip(OwnerText, sourceObject?.Owner, targetObject?.Owner);
            var statusTip = ObjectToolTip(ObjectStatusText, sourceObject?.ObjectStatus, targetObject?.ObjectStatus);
            ObjectStatusToolTip = ReferenceEquals(statusTip, ObjectNotRetrievedToolTip) ? ObjectStatusNotRetrievedToolTip :
                ObjectStatusMeaning + Environment.NewLine + statusTip;
        }

        public int OriginalRowIndex { get; }
        public MembershipResultRow AuthoritativeRow { get; }
        public ComponentSideDisplayEvidence SourceDisplayEvidence { get; }
        public ComponentSideDisplayEvidence TargetDisplayEvidence { get; }
        public string DisplayContextDiagnostic { get; }
        public string ComponentCategoryText { get; }
        public string FriendlyNameText { get; }
        public string InternalNameText { get; }
        public string ParentContextText { get; }
        public string ComponentText { get; }
        public string TypeText => WithoutBothPrefix(ComponentCategoryText);
        public string DefinitionStatusText => DefinitionStatus == "Match" ? "Match (compared properties)" : DefinitionStatus;
        // Only visible/requested tooltip text is materialized, not the entire inventory.
        internal string ComponentToolTip => ToolTip(ComponentText, sourceComponentName, targetComponentName) + Environment.NewLine +
            "Verified display name, otherwise verified logical/internal name. AppSetting composite membership identity is used only when names are not available." +
            Environment.NewLine + FriendlyNameToolTip + Environment.NewLine + InternalNameToolTip +
            Environment.NewLine + "Parent context: " + ParentContextToolTip;
        internal const string DefinitionMeaning = "Match means equality of the implemented comparable properties only, not identity equivalence or full definition equality. Not Compared means definition comparison is unavailable. Double-click for comparable properties and diagnostic evidence.";
        public string ManagedText { get; }
        public string CustomizedText { get; }
        public string LastModifiedText { get; }
        public string OwnerText { get; }
        public string ObjectStatusText { get; }
        internal string ComponentCategoryToolTip { get; }
        internal string FriendlyNameToolTip { get; }
        internal string InternalNameToolTip { get; }
        internal string ParentContextToolTip { get; }
        internal string ManagedToolTip { get; }
        internal string CustomizedToolTip { get; }
        internal string LastModifiedToolTip { get; }
        internal string OwnerToolTip { get; }
        internal string ObjectStatusToolTip { get; }

        public string ComponentKind => AuthoritativeRow.ComponentKind;
        public string PortableKey => AuthoritativeRow.PortableKey;
        public string SourcePresence => AuthoritativeRow.SourcePresence;
        public string TargetPresence => AuthoritativeRow.TargetPresence;
        public string MembershipStatus => AuthoritativeRow.MembershipStatus;
        public string DefinitionStatus => AuthoritativeRow.DefinitionStatus;
        public string ChangedProperties => AuthoritativeRow.ChangedProperties;
        public string SourceResolutionStatus => AuthoritativeRow.SourceResolutionStatus;
        public string TargetResolutionStatus => AuthoritativeRow.TargetResolutionStatus;
        public string Diagnostic => AuthoritativeRow.Diagnostic;
        public int? SourceRawComponentType => AuthoritativeRow.SourceRawComponentType;
        public int? TargetRawComponentType => AuthoritativeRow.TargetRawComponentType;

        private static ComponentDisplayValue ComponentName(ComponentSideDisplayEvidence side, ComponentIdentity identity)
        {
            if (side == null || !side.HasLocalObservation) return side?.FriendlyName;
            if (HasName(side.FriendlyName)) return side.FriendlyName;
            if (HasName(side.InternalName)) return side.InternalName;
            // Use the existing verified composite verbatim. Do not parse it into display names,
            // inspect diagnostic strings, or invent parent/setting metadata.
            if (identity?.Status == IdentityResolutionStatus.Resolved && identity.SemanticKind == ComponentSemanticKinds.AppSetting)
                return new ComponentDisplayValue(DisplayAvailability.Available, identity.ComparisonKey,
                    "Resolved AppSetting composite membership identity (not retrieved display names)");
            return side.FriendlyName;
        }

        private static bool HasName(ComponentDisplayValue value) =>
            value?.Availability == DisplayAvailability.Available && !string.IsNullOrWhiteSpace(value.Value);

        // Presentation decoding of an already resolved contract only. Never resolves a component,
        // creates a parent link, or reads diagnostics. Original keys remain verbatim in evidence.
        internal static bool AppSettingIdentityNames(ComponentIdentity identity, out string parent, out string definition)
        {
            parent = definition = null;
            if (identity?.Status != IdentityResolutionStatus.Resolved || identity.SemanticKind != ComponentSemanticKinds.AppSetting ||
                identity.ComparisonKey == null || !identity.ComparisonKey.StartsWith("appsetting:v1:", StringComparison.Ordinal)) return false;
            var key = identity.ComparisonKey; int position = "appsetting:v1:".Length;
            string first, second;
            if (!ReadIdentitySegment(key, ref position, out first) || position >= key.Length || key[position++] != ':' ||
                !ReadIdentitySegment(key, ref position, out second) || position != key.Length ||
                string.IsNullOrWhiteSpace(first) || string.IsNullOrWhiteSpace(second)) return false;
            Guid guid;
            if (Guid.TryParse(first, out guid) || Guid.TryParse(second, out guid)) return false;
            parent = first; definition = second; return true;
        }

        private static bool ReadIdentitySegment(string key, ref int position, out string value)
        {
            value = null; int separator = key.IndexOf(':', position), length;
            if (separator < position || !int.TryParse(key.Substring(position, separator - position),
                NumberStyles.None, CultureInfo.InvariantCulture, out length) || length <= 0 ||
                key.Substring(position, separator - position) != length.ToString(CultureInfo.InvariantCulture) ||
                length > key.Length - separator - 1) return false;
            value = key.Substring(separator + 1, length); position = separator + 1 + length; return true;
        }

        internal static string QualifiedColumnIdentity(ComponentIdentity identity)
        {
            if (identity?.Status != IdentityResolutionStatus.Resolved || identity.SemanticKind != ComponentSemanticKinds.Column) return null;
            var key = identity.ComparisonKey; var parts = key?.Split('.');
            return parts?.Length == 2 && LogicalName(parts[0]) && LogicalName(parts[1]) ? key : null;
        }
        private static bool LogicalName(string value)
        {
            if (string.IsNullOrEmpty(value) || !(value[0] >= 'a' && value[0] <= 'z' || value[0] >= 'A' && value[0] <= 'Z' || value[0] == '_')) return false;
            foreach (var c in value)
                if (!(c >= 'a' && c <= 'z' || c >= 'A' && c <= 'Z' || c >= '0' && c <= '9' || c == '_')) return false;
            return true;
        }

        internal static ComponentDisplayValue RefinedComponentName(ComponentSideDisplayEvidence side, ComponentIdentity identity)
        {
            if (identity == null) return null;
            var column = QualifiedColumnIdentity(identity);
            if (column != null) return new ComponentDisplayValue(DisplayAvailability.Available, column, "Verified qualified column portable identity; display only");
            if (identity.Status == IdentityResolutionStatus.Resolved && identity.SemanticKind == ComponentSemanticKinds.AppSetting)
            {
                if (HasName(side?.FriendlyName) && side.HasLocalObservation &&
                    side.SolutionComponentId == identity.Record.SolutionComponentId && side.LocalObjectId == identity.Record.ObjectId &&
                    side.RawComponentType == identity.Record.ComponentType) return side.FriendlyName;
                string parent, setting;
                return new ComponentDisplayValue(DisplayAvailability.Available,
                    AppSettingIdentityNames(identity, out parent, out setting) ? setting : identity.ComparisonKey,
                    "Verified AppSetting membership identity formatting; not retrieved display names; not parent ownership proof");
            }
            return null;
        }

        private static ComponentDisplayValue QualifyName(ComponentDisplayValue name, ComponentDisplayValue internalName) =>
            HasName(name) && !string.Equals(name.Value, internalName.Value, StringComparison.Ordinal)
                ? new ComponentDisplayValue(DisplayAvailability.Available, name.Value + " (" + internalName.Value + ")",
                    name.DisplayEvidenceSource + "; " + internalName.DisplayEvidenceSource) : name;

        internal static string WithoutBothPrefix(string value) =>
            value != null && value.StartsWith("Both: ", StringComparison.Ordinal) ? value.Substring(6) : value;

        internal static IReadOnlyList<MembershipGridRow> CreateRows(
            MembershipComparisonPresentation presentation, CompletedMembershipComparison completed, bool refinedPresentation = false)
        {
            if (presentation == null) throw new ArgumentNullException(nameof(presentation));
            var reason = Validate(presentation, completed);
            var rows = new List<MembershipGridRow>(presentation.Rows.Count);
            for (var i = 0; i < presentation.Rows.Count; i++)
                rows.Add(new MembershipGridRow(i, presentation.Rows[i],
                    reason == null ? completed.DisplayEvidence[i] : null, reason, refinedPresentation));
            return rows.AsReadOnly();
        }

        private static string Validate(MembershipComparisonPresentation presentation,
            CompletedMembershipComparison completed)
        {
            if (completed == null) return "No completed display wrapper (presentation-only caller).";
            if (!ReferenceEquals(completed.Presentation, presentation)) return "Display wrapper belongs to another presentation.";
            if (completed.DisplayProjectionOutcome != DisplayProjectionOutcome.Completed) return "Optional display projection did not complete.";
            if (completed.SourceCacheAssociation == DisplayCacheAssociation.Rejected ||
                completed.TargetCacheAssociation == DisplayCacheAssociation.Rejected) return "Display cache association was rejected.";
            if (completed.DisplayEvidence.Count != presentation.Rows.Count) return "Display evidence row count differs from the authoritative rows.";
            for (var i = 0; i < completed.DisplayEvidence.Count; i++)
            {
                var evidence = completed.DisplayEvidence[i];
                if (evidence == null || evidence.RowIndex != i) return "Display evidence row index is invalid.";
                if (!ValidSide(evidence.SourceDisplayEvidence, DisplayEnvironmentSide.Source) ||
                    !ValidSide(evidence.TargetDisplayEvidence, DisplayEnvironmentSide.Target))
                    return "Display evidence environment side or field context is invalid.";
            }
            return null;
        }

        private static bool ValidSide(ComponentSideDisplayEvidence side, DisplayEnvironmentSide expected) =>
            side != null && side.EnvironmentSide == expected && side.ComponentCategory != null &&
            side.FriendlyName != null && side.InternalName != null && side.ParentDisplayContext != null &&
            side.ObjectMetadata != null && side.ObjectMetadata.EnvironmentSide == expected;

        private string RenderObject<T>(ComponentObjectMetadataValue<T> source, ComponentObjectMetadataValue<T> target)
        {
            // B1 observations remain unrequested. Share only identical immutable indicators,
            // never actual values, provenance, local associations or comparison results.
            if (SourceDisplayEvidence?.HasLocalObservation == true && TargetDisplayEvidence?.HasLocalObservation == true &&
                source?.Availability == DisplayAvailability.NotRetrieved && target?.Availability == DisplayAvailability.NotRetrieved)
                return BothObjectNotRetrieved;
            return RenderBothSides(SourceDisplayEvidence, TargetDisplayEvidence, DisplayObject(source), DisplayObject(target));
        }

        private static ComponentDisplayValue DisplayObject<T>(ComponentObjectMetadataValue<T> field) => field == null ? null :
            new ComponentDisplayValue(field.Availability, field.Availability == DisplayAvailability.Available
                ? FormatObjectValue(field) : null, field.Provenance);

        internal static string FormatObjectValue<T>(ComponentObjectMetadataValue<T> field)
        {
            if (field == null || field.Availability != DisplayAvailability.Available || (object)field.Value == null)
                return null;
            object value = field.Value;
            if (value is bool) return (bool)value ? "Yes" : "No";
            var timestamp = value as ComponentTimestampEvidence;
            if (timestamp != null)
                return timestamp.VerifiedInstant.HasValue
                    ? timestamp.VerifiedInstant.Value.ToString("o", CultureInfo.InvariantCulture)
                    : timestamp.RawTimestamp.ToString("yyyy-MM-dd HH:mm:ss.fffffff", CultureInfo.InvariantCulture) +
                        " [Kind=" + timestamp.Kind + "; timezone not interpreted]";
            var owner = value as ComponentOwnerEvidence;
            if (owner != null) return owner.DisplayNameVerified ? owner.DisplayName : "[Owner display name not verified]";
            var status = value as ComponentObjectStatusEvidence;
            if (status != null)
                return (status.LabelVerified ? status.DisplayLabel ?? "[Retrieved null label]" : "[Object status label not verified]") +
                    (status.RawNumericValue.HasValue ? " (code " + status.RawNumericValue.Value.ToString(CultureInfo.InvariantCulture) + ")" : "");
            return null;
        }

        private string ObjectToolTip<T>(string text, ComponentObjectMetadataValue<T> source, ComponentObjectMetadataValue<T> target)
        {
            if (DisplayContextDiagnostic == null && text == BothObjectNotRetrieved &&
                source?.Availability == DisplayAvailability.NotRetrieved && target?.Availability == DisplayAvailability.NotRetrieved &&
                source.Provenance == null && target.Provenance == null && !source.WasRequested && !target.WasRequested)
                return ObjectNotRetrievedToolTip;
            return WithoutBothPrefix(text) + Environment.NewLine + (DisplayContextDiagnostic ??
                ("Source: " + source?.Availability + "; requested: " + source?.WasRequested +
                    "; provenance: " + EscapeForGrid(source?.Provenance) + Environment.NewLine +
                 "Target: " + target?.Availability + "; requested: " + target?.WasRequested +
                    "; provenance: " + EscapeForGrid(target?.Provenance)));
        }

        internal static string RenderBothSides(ComponentSideDisplayEvidence source,
            ComponentSideDisplayEvidence target, ComponentDisplayValue sourceValue, ComponentDisplayValue targetValue)
        {
            if (source == null || target == null) return "[Display evidence unavailable]";
            if (source.HasLocalObservation && target.HasLocalObservation && sourceValue != null && targetValue != null)
            {
                if (sourceValue.Availability == DisplayAvailability.Available &&
                    targetValue.Availability == DisplayAvailability.Available &&
                    !string.IsNullOrWhiteSpace(sourceValue.Value) && !string.IsNullOrWhiteSpace(targetValue.Value) &&
                    string.Equals(sourceValue.Value, targetValue.Value, StringComparison.Ordinal))
                    return "Both: " + EscapeForGrid(sourceValue.Value);
                if (sourceValue.Availability != DisplayAvailability.Available &&
                    sourceValue.Availability == targetValue.Availability)
                    return "Both: " + RenderSideValue(source, sourceValue);
            }
            return "Source: " + RenderSideValue(source, sourceValue) + Environment.NewLine +
                "Target: " + RenderSideValue(target, targetValue);
        }

        internal static string RenderSideValue(ComponentSideDisplayEvidence side, ComponentDisplayValue value)
        {
            if (side == null) return "[Display evidence unavailable]";
            if (!side.HasLocalObservation) return "[No local observation]";
            if (value == null) return "[Display evidence unavailable]";
            switch (value.Availability)
            {
                case DisplayAvailability.Available:
                    if (value.Value == null) return "[Retrieved null]";
                    if (value.Value.Length == 0) return "[Retrieved empty]";
                    if (string.IsNullOrWhiteSpace(value.Value)) return "[Retrieved whitespace]";
                    return EscapeForGrid(value.Value);
                case DisplayAvailability.NotRetrieved: return "[Not retrieved]";
                case DisplayAvailability.AttributeOmitted: return "[Attribute omitted]";
                case DisplayAvailability.Unavailable: return "[Unavailable]";
                case DisplayAvailability.Ambiguous: return "[Ambiguous display evidence]";
                case DisplayAvailability.RetrievalFailed: return "[Display retrieval failed]";
                case DisplayAvailability.NotApplicable: return "[Not applicable]";
                default: return "[Unsupported display state: " + (int)value.Availability + "]";
            }
        }

        // Escaping is presentation only. The detached evidence retains the exact original string.
        internal static string EscapeForGrid(string value) => value?.Replace("\\", "\\\\")
            .Replace("\r", "\\r").Replace("\n", "\\n").Replace("\t", "\\t");

        private string ToolTip(string text, ComponentDisplayValue source, ComponentDisplayValue target) =>
            WithoutBothPrefix(text) + Environment.NewLine + (DisplayContextDiagnostic ??
                ("Source: " + source?.Availability + "; provenance: " + EscapeForGrid(source?.DisplayEvidenceSource) +
                 Environment.NewLine + "Target: " + target?.Availability + "; provenance: " + EscapeForGrid(target?.DisplayEvidenceSource)));
    }
}
