using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using D365SolutionComparer.Models.Membership;

namespace D365SolutionComparer.Services.Membership
{
    internal interface IAppSettingParentEvidenceSink
    { void Observe(AppSettingReadObservation observation); }

    /// <summary>One worker read only. Observes existing source processing; never reads a cache or service.</summary>
    internal sealed class AppSettingParentEvidenceEmission : IAppSettingParentEvidenceSink
    {
        private readonly MembershipSnapshot raw;
        private readonly DisplayEnvironmentSide side;
        private readonly object invocationOwner, readOwner;
        private readonly Dictionary<AppSettingReadKind, Dictionary<Guid, AppSettingReadObservation>> events =
            new Dictionary<AppSettingReadKind, Dictionary<Guid, AppSettingReadObservation>>();
        private bool failed, sealedCapture;
        internal AppSettingParentEvidenceEmission(MembershipSnapshot raw, DisplayEnvironmentSide side,
            object invocationOwner, object readOwner)
        {
            this.raw = raw ?? throw new ArgumentNullException(nameof(raw)); this.side = side;
            if (!Enum.IsDefined(typeof(DisplayEnvironmentSide), side)) throw new ArgumentOutOfRangeException(nameof(side));
            if (invocationOwner == null || readOwner == null) throw new ArgumentException("Read ownership is required.");
            this.invocationOwner = invocationOwner; this.readOwner = readOwner;
            foreach (AppSettingReadKind kind in Enum.GetValues(typeof(AppSettingReadKind)))
                events[kind] = new Dictionary<Guid, AppSettingReadObservation>();
        }
        // Isolation belongs at the source boundary, including projection of the SDK scalars.
        internal static void Emit(IAppSettingParentEvidenceSink sink, Func<AppSettingReadObservation> project)
        {
            if (sink == null) return;
            try { sink.Observe(project()); }
            catch (Exception) { (sink as AppSettingParentEvidenceEmission)?.Disable(); }
        }
        internal void Disable() { failed = true; }
        public void Observe(AppSettingReadObservation observation)
        {
            if (sealedCapture || observation == null || observation.RequestedId == Guid.Empty ||
                !Enum.IsDefined(typeof(AppSettingReadKind), observation.Kind) ||
                !Enum.IsDefined(typeof(NormalReadDisposition), observation.Correlation) || observation.ReturnedRowCount < 0 ||
                observation.Correlation == NormalReadDisposition.Unique &&
                    (observation.ReturnedRowCount != 1 || observation.PrimaryId != observation.RequestedId))
            { failed = true; return; }
            var local = events[observation.Kind];
            if (local.ContainsKey(observation.RequestedId)) { failed = true; return; }
            local.Add(observation.RequestedId, observation);
        }
        private AppSettingReadObservation Find(AppSettingReadKind kind, Guid? id) =>
            id.HasValue && events[kind].TryGetValue(id.Value, out var value) ? value : null;
        private static bool Failed(AppSettingReadObservation row) => row == null ||
            row.Correlation == NormalReadDisposition.NotObserved || row.Correlation == NormalReadDisposition.Conflicting ||
            row.Correlation == NormalReadDisposition.Incomplete || row.Correlation == NormalReadDisposition.Faulted ||
            row.Correlation == NormalReadDisposition.FailedReadOrIncomplete;

        internal AppSettingParentCapture SealAgainstFinalSnapshot(MembershipSnapshot finalOwner, CancellationToken token)
        {
            token.ThrowIfCancellationRequested();
            if (sealedCapture) return null;
            sealedCapture = true;
            var observations = new List<AppSettingParentObservation>();
            bool valid = !failed && raw.State == MembershipSnapshotState.Complete && raw.Solution != null &&
                finalOwner != null && finalOwner.State == MembershipSnapshotState.Complete && finalOwner.Solution != null &&
                raw.Environment.OrganizationId == finalOwner.Environment.OrganizationId &&
                raw.Solution.SolutionId == finalOwner.Solution.SolutionId &&
                string.Equals(raw.SolutionUniqueName, finalOwner.SolutionUniqueName, StringComparison.OrdinalIgnoreCase) &&
                raw.Components.Count == finalOwner.Components.Count;
            if (valid)
            {
                var selected = finalOwner.Components.Where(i => i.Record.ComponentType == 10075).ToList();
                valid = selected.Select(i => i.Record.SolutionComponentId).Distinct().Count() == selected.Count &&
                    selected.Where(i => i.Record.ObjectId.HasValue).Select(i => i.Record.ObjectId).Distinct().Count() ==
                    selected.Count(i => i.Record.ObjectId.HasValue);
                for (int index = 0; index < raw.Components.Count; index++)
                {
                    token.ThrowIfCancellationRequested();
                    var record = raw.Components[index].Record;
                    if (!ReferenceEquals(record, finalOwner.Components[index].Record)) { valid = false; continue; }
                    if (record.ComponentType != 10075) continue;
                    var child = Find(AppSettingReadKind.Child, record.ObjectId);
                    var parent = Find(AppSettingReadKind.Parent, child?.ParentId);
                    var definition = Find(AppSettingReadKind.Definition, child?.DefinitionId);
                    if (Failed(child) || child.Correlation == NormalReadDisposition.Unique &&
                        (child.ParentId.HasValue && Failed(parent) || child.DefinitionId.HasValue && Failed(definition))) valid = false;
                    observations.Add(new AppSettingParentObservation(index, record, child, parent, definition));
                }
            }
            var disposition = failed ? ParentCaptureDisposition.Faulted : !valid ? ParentCaptureDisposition.Incomplete :
                observations.Count == 0 ? ParentCaptureDisposition.NotApplicable : ParentCaptureDisposition.Completed;
            // No trusted partial data escapes failed/incomplete captures.
            return new AppSettingParentCapture(finalOwner, side, raw.Environment.OrganizationId,
                raw.Solution?.SolutionId ?? Guid.Empty, raw.SolutionUniqueName, invocationOwner, readOwner, disposition,
                valid ? observations : new List<AppSettingParentObservation>());
        }
    }
}
