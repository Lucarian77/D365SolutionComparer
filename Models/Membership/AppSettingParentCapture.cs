using System;
using System.Collections.Generic;
using System.Linq;

namespace D365SolutionComparer.Models.Membership
{
    internal enum NormalReadDisposition
    { NotObserved, Unique, Missing, Duplicate, Conflicting, Incomplete, FailedReadOrIncomplete, Faulted }
    internal enum ParentCaptureDisposition { NotApplicable, Completed, Incomplete, Cancelled, Faulted }
    internal enum AppSettingReadKind { Child, Definition, Parent }

    // Source events contain only validated, allowlisted scalars. No SDK row/cache escapes.
    internal sealed class AppSettingReadObservation
    {
        internal AppSettingReadObservation(AppSettingReadKind kind, Guid requestedId,
            NormalReadDisposition correlation, int returnedRowCount, Guid? primaryId = null,
            string uniqueName = null, Guid? parentId = null, bool parentTableVerified = false,
            Guid? definitionId = null, bool definitionTableVerified = false)
        { Kind = kind; RequestedId = requestedId; Correlation = correlation;
            ReturnedRowCount = returnedRowCount; PrimaryId = primaryId; UniqueName = uniqueName;
            ParentId = parentId; ParentTableVerified = parentTableVerified;
            DefinitionId = definitionId; DefinitionTableVerified = definitionTableVerified; }
        internal AppSettingReadKind Kind { get; }
        internal Guid RequestedId { get; }
        internal NormalReadDisposition Correlation { get; }
        internal int ReturnedRowCount { get; }
        internal Guid? PrimaryId { get; }
        internal string UniqueName { get; }
        internal Guid? ParentId { get; }
        internal bool ParentTableVerified { get; }
        internal Guid? DefinitionId { get; }
        internal bool DefinitionTableVerified { get; }
    }

    internal sealed class AppSettingParentObservation
    {
        internal AppSettingParentObservation(int index, SolutionComponentRecord record,
            AppSettingReadObservation child, AppSettingReadObservation parent, AppSettingReadObservation definition)
        { RawOccurrenceIndex = index; OriginalRecord = record; Child = child; Parent = parent; Definition = definition; }
        internal int RawOccurrenceIndex { get; }
        internal SolutionComponentRecord OriginalRecord { get; }
        internal AppSettingReadObservation Child { get; }
        internal AppSettingReadObservation Parent { get; }
        internal AppSettingReadObservation Definition { get; }
        internal string ProvenanceCode => "NormalAppSettingRead:v1";
    }

    internal sealed class AppSettingParentCapture
    {
        internal AppSettingParentCapture(MembershipSnapshot owner, DisplayEnvironmentSide side,
            Guid organizationId, Guid solutionId, string solutionUniqueName, object invocationOwner,
            object readOwner, ParentCaptureDisposition disposition, IEnumerable<AppSettingParentObservation> observations)
        { Owner = owner; Side = side; OrganizationId = organizationId; SolutionId = solutionId;
            SolutionUniqueName = solutionUniqueName; ComparisonInvocationOwner = invocationOwner;
            ReadOwner = readOwner; Disposition = disposition; Observations = observations.ToList().AsReadOnly(); }
        internal MembershipSnapshot Owner { get; }
        internal DisplayEnvironmentSide Side { get; }
        internal Guid OrganizationId { get; }
        internal Guid SolutionId { get; }
        internal string SolutionUniqueName { get; }
        internal object ComparisonInvocationOwner { get; }
        internal object ReadOwner { get; }
        internal ParentCaptureDisposition Disposition { get; }
        internal IReadOnlyList<AppSettingParentObservation> Observations { get; }
        internal string DiagnosticCode => "AppSettingParentCapture:" + Disposition;
        internal bool BelongsTo(MembershipSnapshot snapshot, DisplayEnvironmentSide side,
            object invocationOwner, object readOwner) => snapshot != null && snapshot.Solution != null &&
            ReferenceEquals(Owner, snapshot) && Side == side && invocationOwner != null && readOwner != null &&
            ReferenceEquals(ComparisonInvocationOwner, invocationOwner) && ReferenceEquals(ReadOwner, readOwner) &&
            OrganizationId == snapshot.Environment.OrganizationId && SolutionId == snapshot.Solution.SolutionId &&
            string.Equals(SolutionUniqueName, snapshot.SolutionUniqueName, StringComparison.OrdinalIgnoreCase);
    }
}
