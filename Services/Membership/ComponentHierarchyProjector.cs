using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using D365SolutionComparer.Models.Membership;
using D365SolutionComparer.Models.Presentation;

namespace D365SolutionComparer.Services.Membership
{
    /// <summary>Local H1 projection only. No service, resolver, cache mutation or UI dependency.</summary>
    public static class ComponentHierarchyProjector
    {
        public static ComponentHierarchyProjection Project(MembershipComparisonPresentation presentation,
            ComponentHierarchyParentEvidence sourceParents = null, ComponentHierarchyParentEvidence targetParents = null,
            CancellationToken cancellationToken = default(CancellationToken)) =>
            ProjectCore(presentation, null, sourceParents, targetParents, cancellationToken, null);

        public static ComponentHierarchyProjection Project(CompletedMembershipComparison completed,
            ComponentHierarchyParentEvidence sourceParents = null, ComponentHierarchyParentEvidence targetParents = null,
            CancellationToken cancellationToken = default(CancellationToken)) =>
            ProjectCore((completed ?? throw new ArgumentNullException(nameof(completed))).Presentation, completed,
                sourceParents, targetParents, cancellationToken, null);

        /// <summary>
        /// The caller must have retained the report's original snapshot provenance. Legacy reports
        /// do not contain an organization ID: do not bind imported reports by display name alone.
        /// Scope fields are checked; the resulting observations stay bound to this exact snapshot/role.
        /// </summary>
        public static ComponentHierarchyParentEvidence RetainAppSettingParents(MembershipSnapshot verifiedReportOwner,
            DisplayEnvironmentSide side, AppSettingEvidenceReport report)
        {
            if (verifiedReportOwner == null) throw new ArgumentNullException(nameof(verifiedReportOwner));
            if (!Enum.IsDefined(typeof(DisplayEnvironmentSide), side)) throw new ArgumentOutOfRangeException(nameof(side));
            var observations = new List<HierarchyStructuralObservation>();
            if (report == null || verifiedReportOwner.Solution == null ||
                report.SolutionId != verifiedReportOwner.Solution.SolutionId ||
                !Same(report.SolutionUniqueName, verifiedReportOwner.SolutionUniqueName) ||
                !Same(report.EnvironmentName, verifiedReportOwner.Environment.DisplayName))
                return new ComponentHierarchyParentEvidence(verifiedReportOwner, side, observations,
                    "AppSetting report scope is unavailable or rejected; no structural binding attached.");

            var byId = verifiedReportOwner.Components.GroupBy(i => i.Record.SolutionComponentId)
                .ToDictionary(g => g.Key, g => g.ToList());
            foreach (var candidate in report.Candidates)
            {
                List<ComponentIdentity> children;
                Guid parent;
                var correlation = candidate.Correlations.Where(c => Same(c.EntityLogicalName, "appsetting")).ToList();
                string primary = null, reference = null;
                if (correlation.Count == 1 && correlation[0].Records.Count == 1)
                {
                    correlation[0].Records[0].Fields.TryGetValue("appsettingid", out primary);
                    correlation[0].Records[0].Fields.TryGetValue("parentappmoduleid", out reference);
                }
                Guid primaryId, referenceId;
                bool verified = candidate.ComponentType == 10075 && candidate.ObjectId.HasValue && candidate.ObjectId != Guid.Empty &&
                    candidate.State == AppSettingEvidenceState.Confirmed &&
                    candidate.SolutionId == report.SolutionId && Same(candidate.SolutionUniqueName, report.SolutionUniqueName) &&
                    Same(candidate.EnvironmentName, report.EnvironmentName) &&
                    byId.TryGetValue(candidate.SolutionComponentId, out children) && children.Count == 1 &&
                    children[0].Record.ObjectId == candidate.ObjectId && children[0].Record.ComponentType == candidate.ComponentType &&
                    children[0].Record.RootSolutionComponentId == candidate.RootSolutionComponentId &&
                    children[0].Record.RootComponentBehavior == candidate.RootComponentBehavior && children[0].Record.IsMetadata == candidate.IsMetadata &&
                    correlation.Count == 1 && correlation[0].State == AppSettingEvidenceState.Confirmed &&
                    Same(correlation[0].PrimaryIdAttribute, "appsettingid") && correlation[0].MatchedRecordCount == 1 &&
                    correlation[0].Records.Count == 1 && correlation[0].Records[0].RecordId == candidate.ObjectId &&
                    Guid.TryParse(primary, out primaryId) && primaryId == candidate.ObjectId &&
                    Guid.TryParse(reference, out referenceId) && Guid.TryParse(candidate.ParentAppModuleId, out parent) &&
                    parent != Guid.Empty && referenceId == parent && !string.IsNullOrWhiteSpace(candidate.ParentAppModuleUniqueName);
                observations.Add(new HierarchyStructuralObservation(candidate.SolutionComponentId, candidate.ObjectId,
                    Guid.TryParse(candidate.ParentAppModuleId, out parent) ? (Guid?)parent : null,
                    ComponentSemanticKinds.AppModule, candidate.ParentAppModuleUniqueName,
                    "Bound completed AppSetting evidence: appsetting.parentappmoduleid + appmodule.uniquename", verified));
            }
            return new ComponentHierarchyParentEvidence(verifiedReportOwner, side, observations);
        }

        /// <summary>Strict normal-source adapter. Legacy report gates and grouping remain unchanged.</summary>
        internal static ComponentHierarchyParentEvidence RetainAppSettingParentsFromNormalRead(
            AppSettingParentCapture capture, MembershipSnapshot expectedOwner, DisplayEnvironmentSide expectedSide,
            object expectedComparisonOwner, object expectedReadOwner)
        {
            if (capture == null || capture.Disposition != ParentCaptureDisposition.Completed ||
                !capture.BelongsTo(expectedOwner, expectedSide, expectedComparisonOwner, expectedReadOwner)) return null;
            var children = expectedOwner.Components.GroupBy(i => i.Record.SolutionComponentId)
                .ToDictionary(g => g.Key, g => g.ToList());
            var objects = expectedOwner.Components.Where(i => i.Record.ComponentType == 10075 && i.Record.ObjectId.HasValue)
                .GroupBy(i => i.Record.ObjectId.Value).ToDictionary(g => g.Key, g => g.Count());
            var observations = new List<HierarchyStructuralObservation>();
            foreach (var item in capture.Observations)
            {
                var record = item.OriginalRecord;
                var child = item.Child; var parent = item.Parent; var definition = item.Definition;
                List<ComponentIdentity> matches;
                bool verified = record != null && record.ComponentType == 10075 && record.ObjectId.HasValue &&
                    record.ObjectId != Guid.Empty && item.RawOccurrenceIndex >= 0 && item.RawOccurrenceIndex < expectedOwner.Components.Count &&
                    ReferenceEquals(expectedOwner.Components[item.RawOccurrenceIndex].Record, record) &&
                    children.TryGetValue(record.SolutionComponentId, out matches) && matches.Count == 1 &&
                    matches[0].RegisteredDefinition?.IsAppSetting == true && objects[record.ObjectId.Value] == 1 &&
                    Unique(child, record.ObjectId) && child.ParentTableVerified && child.DefinitionTableVerified &&
                    Unique(parent, child.ParentId) && Unique(definition, child.DefinitionId) &&
                    !string.IsNullOrWhiteSpace(parent.UniqueName) && !string.IsNullOrWhiteSpace(definition.UniqueName);
                if (record != null)
                    observations.Add(new HierarchyStructuralObservation(record.SolutionComponentId, record.ObjectId,
                        child?.ParentId, ComponentSemanticKinds.AppModule, parent?.UniqueName,
                        "Bound normal AppSetting evidence: appsetting.parentappmoduleid + appmodule.uniquename", verified));
            }
            return new ComponentHierarchyParentEvidence(expectedOwner, expectedSide, observations);
        }

        private static bool Unique(AppSettingReadObservation observation, Guid? expectedId) =>
            observation != null && expectedId.HasValue && expectedId != Guid.Empty &&
            observation.Correlation == NormalReadDisposition.Unique && observation.ReturnedRowCount == 1 &&
            observation.RequestedId == expectedId && observation.PrimaryId == expectedId;

        // Stage callback is internal for deterministic cancellation/failure-isolation tests.
        internal static ComponentHierarchyProjection ProjectCore(MembershipComparisonPresentation presentation,
            CompletedMembershipComparison completed, ComponentHierarchyParentEvidence sourceParents,
            ComponentHierarchyParentEvidence targetParents, CancellationToken token, Action<string> stage)
        {
            if (presentation == null) throw new ArgumentNullException(nameof(presentation));
            var raw = Raw(presentation);
            var links = LinkRows(presentation, raw);
            var diagnostics = new List<string>();
            try
            {
                token.ThrowIfCancellationRequested(); stage?.Invoke("AccountingPrepared"); token.ThrowIfCancellationRequested();
                int unrepresented = raw.Count(r => r.Comparisons.Count == 0);
                if (unrepresented > 0) diagnostics.Add(unrepresented + " raw occurrence(s) have no represented comparison row; all retained.");
                int unmapped = links.Count(l => l.Row.SourceIdentity != null && l.SourceOccurrences.Count == 0 ||
                    l.Row.TargetIdentity != null && l.TargetOccurrences.Count == 0);
                if (unmapped > 0) diagnostics.Add(unmapped + " comparison row(s) have no corresponding retained raw occurrence; rows retained without guessing.");
                var associations = raw.Select(r => new List<ComponentHierarchyAssociation>()).ToArray();
                foreach (var side in new[] { DisplayEnvironmentSide.Source, DisplayEnvironmentSide.Target })
                {
                    var snapshot = side == DisplayEnvironmentSide.Source ? presentation.Source.Snapshot : presentation.Target.Snapshot;
                    if (snapshot == null || snapshot.State == MembershipSnapshotState.Unavailable)
                        diagnostics.Add(side + " raw inventory unavailable; zero retained records is not absence evidence.");
                    if (snapshot != null && !Same(snapshot.SolutionUniqueName, presentation.SolutionUniqueName))
                        throw new InvalidOperationException("Snapshot solution scope mismatch.");
                    var local = raw.Where(r => r.Side == side).ToList();
                    var byId = local.GroupBy(r => r.Identity.Record.SolutionComponentId).ToDictionary(g => g.Key, g => g.ToList());
                    Roots(local, byId, associations, token);
                    var observations = new List<HierarchyStructuralObservation>();
                    var supplied = side == DisplayEnvironmentSide.Source ? sourceParents : targetParents;
                    if (supplied != null)
                    {
                        if (!ReferenceEquals(supplied.Owner, snapshot) || supplied.Side != side)
                            diagnostics.Add(side + " structural evidence snapshot/role scope rejected.");
                        else
                        {
                            observations.AddRange(supplied.Observations);
                            if (supplied.Diagnostic != null) diagnostics.Add(supplied.Diagnostic);
                        }
                    }
                    observations.AddRange(Columns(completed, side, diagnostics));
                    Structural(local, byId, observations, associations, token);
                }
                AssociationCycles(raw, associations, token);
                stage?.Invoke("ParentsValidated"); token.ThrowIfCancellationRequested();
                var usedParents = new HashSet<int>(associations.SelectMany(a => a)
                    .Where(a => a.Status == HierarchyAssociationStatus.Verified && a.ParentOccurrenceIndex.HasValue)
                    .Select(a => a.ParentOccurrenceIndex.Value));
                var pairs = CorrelateParents(links, raw, usedParents);
                stage?.Invoke("Completed"); token.ThrowIfCancellationRequested();
                return Finish(presentation, raw, links, associations, pairs, HierarchyProjectionOutcome.Completed, diagnostics);
            }
            catch (OperationCanceledException)
            {
                diagnostics.Add("Hierarchy projection cancelled; all partial associations discarded. Authoritative data retained.");
                return Fallback(presentation, raw, links, HierarchyProjectionOutcome.Cancelled, diagnostics);
            }
            catch (Exception ex)
            {
                // Do not retain exception messages, service payloads, stack traces or partial relationships.
                diagnostics.Add("Hierarchy projection failed (" + ex.GetType().Name + "); authoritative data retained.");
                return Fallback(presentation, raw, links, HierarchyProjectionOutcome.Failed, diagnostics);
            }
        }

        private sealed class RawOccurrence
        {
            internal int Index, OriginalIndex;
            internal DisplayEnvironmentSide Side;
            internal MembershipSnapshot Snapshot;
            internal ComponentIdentity Identity;
            internal readonly List<int> Comparisons = new List<int>();
        }

        private static List<RawOccurrence> Raw(MembershipComparisonPresentation p)
        {
            var result = new List<RawOccurrence>();
            foreach (var side in new[] { DisplayEnvironmentSide.Source, DisplayEnvironmentSide.Target })
            {
                var snapshot = side == DisplayEnvironmentSide.Source ? p.Source.Snapshot : p.Target.Snapshot;
                if (snapshot == null) continue;
                var ordered = snapshot.Components.Select((identity, index) => new RawOccurrence
                    { Identity = identity, OriginalIndex = index, Snapshot = snapshot, Side = side })
                    .OrderBy(r => r.Identity.Record.SolutionComponentId).ThenBy(r => r.Identity.Record.ComponentType)
                    .ThenBy(r => r.Identity.Record.ObjectId).ThenBy(r => r.Identity.Record.RootSolutionComponentId)
                    .ThenBy(r => r.Identity.Record.RootComponentBehavior).ThenBy(r => r.Identity.Record.IsMetadata)
                    .ThenBy(r => r.Identity.ComponentTypeKey, StringComparer.Ordinal)
                    .ThenBy(r => r.Identity.ComparisonKey, StringComparer.Ordinal).ThenBy(r => r.Identity.Status)
                    .ThenBy(r => r.OriginalIndex);
                foreach (var item in ordered) { item.Index = result.Count; result.Add(item); }
            }
            return result;
        }

        private static IReadOnlyList<HierarchyComparisonRowLink> LinkRows(MembershipComparisonPresentation p, List<RawOccurrence> raw)
        {
            var maps = new[] { DisplayEnvironmentSide.Source, DisplayEnvironmentSide.Target }.ToDictionary(side => side,
                side => raw.Where(r => r.Side == side).GroupBy(r => r.Identity.Record).ToDictionary(g => g.Key, g => g.ToList()));
            var repeats = new[] { DisplayEnvironmentSide.Source, DisplayEnvironmentSide.Target }.ToDictionary(side => side,
                side => raw.Where(r => r.Side == side && Collapsible(r.Identity)).GroupBy(r => RepeatKey(r.Identity), StringComparer.OrdinalIgnoreCase)
                    .ToDictionary(g => g.Key, g => g.ToList(), StringComparer.OrdinalIgnoreCase));
            var links = new List<HierarchyComparisonRowLink>();
            for (int index = 0; index < p.Rows.Count; index++)
            {
                var row = p.Rows[index];
                Func<DisplayEnvironmentSide, ComponentIdentity, IEnumerable<int>> find = (side, identity) =>
                {
                    if (identity == null) return new int[0];
                    List<RawOccurrence> matches;
                    // Only the already approved collapsed-reference kinds use a backing/key mapping.
                    // This relates audit occurrences to an existing row; it does not collapse nodes or create matches.
                    if (!(Collapsible(identity) && repeats[side].TryGetValue(RepeatKey(identity), out matches)) &&
                        !maps[side].TryGetValue(identity.Record, out matches)) return new int[0];
                    foreach (var item in matches) item.Comparisons.Add(index);
                    return matches.Select(r => r.Index).ToList();
                };
                links.Add(new HierarchyComparisonRowLink(index, row, find(DisplayEnvironmentSide.Source, row.SourceIdentity),
                    find(DisplayEnvironmentSide.Target, row.TargetIdentity)));
            }
            return links.AsReadOnly();
        }

        private static bool Collapsible(ComponentIdentity i) => i.Status == IdentityResolutionStatus.Resolved &&
            i.Record.ObjectId.HasValue && i.Record.ObjectId != Guid.Empty &&
            (i.SemanticKind == ComponentSemanticKinds.SystemForm || i.SemanticKind == ComponentSemanticKinds.PluginAssembly ||
             i.SemanticKind == ComponentSemanticKinds.SdkMessageProcessingStep);
        private static string RepeatKey(ComponentIdentity i) => i.ComponentTypeKey.Length + ":" + i.ComponentTypeKey +
            i.ComparisonKey.Length + ":" + i.ComparisonKey + ":" + i.Record.ObjectId.Value.ToString("D");

        private static void Roots(List<RawOccurrence> local, Dictionary<Guid, List<RawOccurrence>> byId,
            List<ComponentHierarchyAssociation>[] associations, CancellationToken token)
        {
            var parent = new Dictionary<int, int>();
            var statuses = new Dictionary<int, HierarchyAssociationStatus>();
            foreach (var item in local)
            {
                token.ThrowIfCancellationRequested(); var root = item.Identity.Record.RootSolutionComponentId;
                List<RawOccurrence> matches;
                var status = !root.HasValue ? HierarchyAssociationStatus.NotSupplied : root == Guid.Empty ? HierarchyAssociationStatus.InvalidReference :
                    root == item.Identity.Record.SolutionComponentId ? HierarchyAssociationStatus.SelfReference :
                    !byId.TryGetValue(root.Value, out matches) ? HierarchyAssociationStatus.MissingParent :
                    matches.Count != 1 ? HierarchyAssociationStatus.DuplicateParent : HierarchyAssociationStatus.Verified;
                statuses[item.Index] = status;
                if (status == HierarchyAssociationStatus.Verified) parent[item.Index] = byId[root.Value][0].Index;
            }
            // Iterative, memoized root-chain validation avoids recursion and quadratic deep-tree walks.
            var valid = new Dictionary<int, bool>();
            foreach (var item in local)
            {
                var path = new List<int>(); var positions = new Dictionary<int, int>(); int current = item.Index;
                bool ok;
                while (!valid.TryGetValue(current, out ok))
                {
                    token.ThrowIfCancellationRequested(); int cycle;
                    if (positions.TryGetValue(current, out cycle))
                    {
                        foreach (var id in path.Skip(cycle)) statuses[id] = HierarchyAssociationStatus.CyclicReference;
                        ok = false; break;
                    }
                    positions[current] = path.Count; path.Add(current);
                    if (statuses[current] == HierarchyAssociationStatus.NotSupplied) { ok = true; break; }
                    if (statuses[current] != HierarchyAssociationStatus.Verified) { ok = false; break; }
                    current = parent[current];
                }
                foreach (var id in path)
                {
                    valid[id] = ok;
                    if (!ok && statuses[id] == HierarchyAssociationStatus.Verified) statuses[id] = HierarchyAssociationStatus.UnverifiedAncestor;
                }
            }
            foreach (var item in local)
                associations[item.Index].Add(new ComponentHierarchyAssociation(HierarchyRelationshipKind.InclusionRoot,
                    statuses[item.Index], parent.ContainsKey(item.Index) ? (int?)parent[item.Index] : null,
                    item.Identity.Record.RootSolutionComponentId, "Completed snapshot: solutioncomponent.rootsolutioncomponentid",
                    statuses[item.Index] == HierarchyAssociationStatus.Verified ? "Verified local inclusion-root link; not schema ownership or top-level proof." :
                        "Inclusion root: " + statuses[item.Index] + ". No top-level or absence inference."));
        }

        private static IEnumerable<HierarchyStructuralObservation> Columns(CompletedMembershipComparison completed,
            DisplayEnvironmentSide side, List<string> diagnostics)
        {
            var result = new List<HierarchyStructuralObservation>();
            if (completed == null || completed.DisplayProjectionOutcome != DisplayProjectionOutcome.Completed) return result;
            var scope = side == DisplayEnvironmentSide.Source ? completed.SourceCacheAssociation : completed.TargetCacheAssociation;
            if (scope != DisplayCacheAssociation.Verified) return result;
            foreach (var display in completed.DisplayEvidence)
            {
                var row = completed.Presentation.Rows[display.RowIndex];
                var identity = side == DisplayEnvironmentSide.Source ? row.SourceIdentity : row.TargetIdentity;
                var value = side == DisplayEnvironmentSide.Source ? display.SourceDisplayEvidence : display.TargetDisplayEvidence;
                if (identity == null || identity.Record.ComponentType != 2 || value == null) continue;
                if (value.EnvironmentSide != side || value.SolutionComponentId != identity.Record.SolutionComponentId ||
                    value.LocalObjectId != identity.Record.ObjectId || value.RawComponentType != identity.Record.ComponentType)
                { diagnostics.Add(side + " Column display context rejected."); continue; }
                var field = value.ParentDisplayContext;
                if (value.BackingCorrelationStatus != DisplayBackingCorrelationStatus.VerifiedMetadata ||
                    field?.Availability != DisplayAvailability.Available || string.IsNullOrWhiteSpace(field.Value) ||
                    !(field.DisplayEvidenceSource == "AttributeMetadata.EntityLogicalName" || field.DisplayEvidenceSource == "VerifiedParentMetadata.LogicalName")) continue;
                result.Add(new HierarchyStructuralObservation(identity.Record.SolutionComponentId, identity.Record.ObjectId, null,
                    ComponentSemanticKinds.Table, field.Value, field.DisplayEvidenceSource, true));
            }
            return result;
        }

        private static void Structural(List<RawOccurrence> local, Dictionary<Guid, List<RawOccurrence>> byId,
            List<HierarchyStructuralObservation> observations, List<ComponentHierarchyAssociation>[] associations, CancellationToken token)
        {
            var parents = local.Where(r => r.Identity.Status == IdentityResolutionStatus.Resolved)
                .GroupBy(r => ParentKey(r.Identity.SemanticKind, r.Identity.ComparisonKey), StringComparer.OrdinalIgnoreCase)
                .ToDictionary(g => g.Key, g => g.ToList(), StringComparer.OrdinalIgnoreCase);
            var evidence = observations.GroupBy(o => o.Child).ToDictionary(g => g.Key, g => g.ToList());
            var localObjects = local.Where(r => r.Identity.Record.ObjectId.HasValue)
                .GroupBy(r => r.Identity.Record.ObjectId.Value).ToDictionary(g => g.Key, g => g.ToList());
            var uncertainKinds = new HashSet<string>(local.Where(r => r.Identity.Status != IdentityResolutionStatus.Resolved)
                .Select(r => r.Identity.SemanticKind).Where(k => k != null), StringComparer.OrdinalIgnoreCase);
            foreach (var child in local)
            {
                token.ThrowIfCancellationRequested(); List<HierarchyStructuralObservation> entries;
                if (!evidence.TryGetValue(child.Identity.Record.SolutionComponentId, out entries)) continue;
                foreach (var entry in entries)
                {
                    List<RawOccurrence> matches = null;
                    bool verified = entry.Verified && entry.ChildObject == child.Identity.Record.ObjectId &&
                        byId[entry.Child].Count == 1 && !string.IsNullOrWhiteSpace(entry.ParentKey);
                    if (verified) parents.TryGetValue(ParentKey(entry.ParentKind, entry.ParentKey), out matches);
                    List<RawOccurrence> located = null;
                    if (entry.ParentObject.HasValue && localObjects.TryGetValue(entry.ParentObject.Value, out located))
                        located = located.Where(r => Same(r.Identity.SemanticKind, entry.ParentKind)).ToList();
                    bool identityUncertain = matches == null && (located?.Count == 1 &&
                        located[0].Identity.Status != IdentityResolutionStatus.Resolved ||
                        !entry.ParentObject.HasValue && uncertainKinds.Contains(entry.ParentKind));
                    var status = !verified ? HierarchyAssociationStatus.UnverifiedEvidence : matches == null ? HierarchyAssociationStatus.MissingParent :
                        matches.Count != 1 ? HierarchyAssociationStatus.DuplicateParent :
                        entry.ParentObject.HasValue && entry.ParentObject != matches[0].Identity.Record.ObjectId ? HierarchyAssociationStatus.UnverifiedEvidence :
                        matches[0].Index == child.Index ? HierarchyAssociationStatus.SelfReference : HierarchyAssociationStatus.Verified;
                    if (verified && identityUncertain) status = HierarchyAssociationStatus.ParentIdentityNotEstablished;
                    if (verified && located?.Count > 1) status = HierarchyAssociationStatus.DuplicateParent;
                    if (verified && located?.Count == 1 && located[0].Identity.Status == IdentityResolutionStatus.Resolved &&
                        !Same(located[0].Identity.ComparisonKey, entry.ParentKey)) status = HierarchyAssociationStatus.ConflictingEvidence;
                    associations[child.Index].Add(new ComponentHierarchyAssociation(HierarchyRelationshipKind.StructuralParent, status,
                        status == HierarchyAssociationStatus.DuplicateParent ? null : matches?.Count == 1 ? (int?)matches[0].Index :
                            identityUncertain && located?.Count == 1 ? (int?)located[0].Index : null, entry.ParentObject, entry.Provenance,
                        "Structural binding: " + status + ". This is not a raw inclusion-root reference."));
                }
                var verifiedParents = associations[child.Index].Where(a => a.Kind == HierarchyRelationshipKind.StructuralParent &&
                    a.Status == HierarchyAssociationStatus.Verified).Select(a => a.ParentOccurrenceIndex).Distinct().Count();
                if (verifiedParents > 1)
                    associations[child.Index].Add(new ComponentHierarchyAssociation(HierarchyRelationshipKind.StructuralParent,
                        HierarchyAssociationStatus.ConflictingEvidence, null, null, "Retained structural parent observations",
                        "Conflicting structural parent evidence; all observations retained without choosing a parent."));
            }
        }
        private static string ParentKey(string kind, string key) => (kind ?? "").Length + ":" + kind + key;
        private static bool Same(string left, string right) => StringComparer.OrdinalIgnoreCase.Equals(left, right);

        private static void AssociationCycles(List<RawOccurrence> raw, List<ComponentHierarchyAssociation>[] associations,
            CancellationToken token)
        {
            // Guard the prospective primary display path as well as raw root chains. This
            // does not declare structural scope equivalent to inclusion ownership.
            var parents = new Dictionary<int, ComponentHierarchyAssociation>();
            foreach (var item in raw)
            {
                var entries = associations[item.Index];
                if (entries.Any(a => a.Status == HierarchyAssociationStatus.ConflictingEvidence)) continue;
                var root = entries.First(a => a.Kind == HierarchyRelationshipKind.InclusionRoot);
                if (root.Status == HierarchyAssociationStatus.Verified) parents[item.Index] = root;
                else if (root.Status == HierarchyAssociationStatus.NotSupplied)
                {
                    var structure = entries.Where(a => a.Kind == HierarchyRelationshipKind.StructuralParent).ToList();
                    if (structure.Count > 0 && structure.All(a => a.Status == HierarchyAssociationStatus.Verified))
                        parents[item.Index] = structure[0];
                }
            }
            var visited = new HashSet<int>();
            foreach (var item in raw)
            {
                var path = new List<int>(); var positions = new Dictionary<int, int>(); int current = item.Index;
                while (!visited.Contains(current))
                {
                    token.ThrowIfCancellationRequested(); int cycle; ComponentHierarchyAssociation edge;
                    if (positions.TryGetValue(current, out cycle))
                    {
                        foreach (var id in path.Skip(cycle))
                            associations[id].Add(new ComponentHierarchyAssociation(parents[id].Kind,
                                HierarchyAssociationStatus.CyclicReference, parents[id].ParentOccurrenceIndex,
                                parents[id].ReferencedLocalId, "Local hierarchy association path validation",
                                "Cyclic association path; evidence retained without constructing a cyclic display tree."));
                        break;
                    }
                    positions[current] = path.Count; path.Add(current);
                    if (!parents.TryGetValue(current, out edge)) break;
                    current = edge.ParentOccurrenceIndex.Value;
                }
                visited.UnionWith(path);
            }
        }

        private static IEnumerable<HierarchyParentCorrelation> CorrelateParents(IReadOnlyList<HierarchyComparisonRowLink> links,
            List<RawOccurrence> raw, HashSet<int> usedParents)
        {
            foreach (var link in links)
            {
                var source = link.Row.SourceIdentity; var target = link.Row.TargetIdentity;
                if (link.Row.Presence != MembershipPresence.PresentInBoth || source?.Status != IdentityResolutionStatus.Resolved ||
                    target?.Status != IdentityResolutionStatus.Resolved || !Same(source.ComponentTypeKey, target.ComponentTypeKey) ||
                    !Same(source.ComparisonKey, target.ComparisonKey) || link.SourceOccurrences.Count != 1 || link.TargetOccurrences.Count != 1) continue;
                int s = link.SourceOccurrences[0], t = link.TargetOccurrences[0];
                if (!(usedParents.Contains(s) || usedParents.Contains(t)) || raw[s].Identity.Status != IdentityResolutionStatus.Resolved ||
                    raw[t].Identity.Status != IdentityResolutionStatus.Resolved ||
                    !Same(raw[s].Identity.ComparisonKey, source.ComparisonKey) || !Same(raw[t].Identity.ComparisonKey, target.ComparisonKey)) continue;
                yield return new HierarchyParentCorrelation(link.RowIndex, s, t, source.ComponentTypeKey, source.ComparisonKey);
            }
        }

        private static ComponentHierarchyProjection Fallback(MembershipComparisonPresentation p, List<RawOccurrence> raw,
            IReadOnlyList<HierarchyComparisonRowLink> links, HierarchyProjectionOutcome outcome, List<string> diagnostics) =>
            Finish(p, raw, links, raw.Select(r => new List<ComponentHierarchyAssociation> { new ComponentHierarchyAssociation(
                HierarchyRelationshipKind.InclusionRoot, HierarchyAssociationStatus.UnverifiedEvidence, null,
                r.Identity.Record.RootSolutionComponentId, "Completed snapshot retained; hierarchy unavailable", outcome.ToString()) }).ToArray(),
                new HierarchyParentCorrelation[0], outcome, diagnostics);

        private static ComponentHierarchyProjection Finish(MembershipComparisonPresentation p, List<RawOccurrence> raw,
            IReadOnlyList<HierarchyComparisonRowLink> links, List<ComponentHierarchyAssociation>[] associations,
            IEnumerable<HierarchyParentCorrelation> pairs, HierarchyProjectionOutcome outcome, List<string> diagnostics) =>
            new ComponentHierarchyProjection(p, outcome, raw.Select(r => new ComponentHierarchyOccurrence(r.Index,
                r.Side, r.OriginalIndex, r.Snapshot, r.Identity, r.Comparisons, associations[r.Index]
                    .OrderBy(a => a.Kind).ThenBy(a => a.Status).ThenBy(a => a.ParentOccurrenceIndex)
                    .ThenBy(a => a.ReferencedLocalId).ThenBy(a => a.Provenance, StringComparer.Ordinal))), links, pairs, diagnostics);
    }
}
