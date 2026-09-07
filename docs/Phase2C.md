# Phase 2C: minimal live membership validation UI

Starting checkpoint: `9d793f6ae25bda1bd31ff9a4cd11d8e03707066b` on `codex/1.2026.2.0`. These changes remain uncommitted pending review.

## Workflow and command gating

The existing solution comparison grid now has a **Compare Membership** action. It operates on exactly one selected `SolutionCompareResult` and is enabled only when both source and target connections still match the services from which their solution lists were loaded, both lists are loaded, no membership operation is running, and the selected Unique Name exists in at least one list. Existing solution comparison remains keyed by Unique Name.

The action queries both available environments independently. This positively distinguishes a present solution with zero components from a solution that is absent. A row that exists in only the source or target list therefore still produces a normal two-side result containing one `Complete` or `SolutionAbsent` snapshot per successful environment read.

The existing comparison grid, filters, row-details double-click action, settings, and export paths are unchanged.

## Background execution and cancellation

Live reads run in XrmToolBox `WorkAsync` with cancellation enabled. Source and target calls are sequential; no parallel Dataverse requests were added. A cancellation bridge observes the XrmToolBox `BackgroundWorker` and cancels the token consumed by `DataverseSolutionMembershipOperation`, the paged reader, and the resolver. Cancellation closes the work operation without opening a result window, so a partial inventory cannot appear complete.

The coordinated `ReadAndResolve` path now has a UI overload that captures the organization ID from its single `WhoAmIRequest`, builds the `EnvironmentIdentity`, and reuses the same `DataverseReadContext` for solution lookup, paged membership retrieval, and bulk resolution. It reports environment validation, membership paging, identity resolution, and completion stages. The existing overloads and standalone resolver APIs are unchanged.

A non-cancellation failure is contained to the affected side and represented as `Unavailable`, with no snapshot and nullable inventory/resolution counts. The other environment is still read for live diagnosis. Presentation logic treats every component opposite an unavailable side as indeterminate and never as missing.

## Results and diagnostics

The dedicated membership window shows:

- Component Kind
- Component Identity / Portable Key
- Source Presence
- Target Presence
- Membership Status
- Source Resolution Status
- Target Resolution Status
- Diagnostic / Reason
- Source Raw Component Type
- Target Raw Component Type

The header reports solution state for each environment and summary counts for Present in Both, Source Only, Target Only, Unsupported, Unresolved, and Ambiguous. Missing rows include the evidence used by `SolutionMembershipComparer`. Unsupported, unresolved, ambiguous, and retrieval-unavailable rows remain explicitly indeterminate. Duplicate canonical keys remain ambiguous even when the opposite environment is unavailable.

Per-side live diagnostics show total Dataverse request count, elapsed time, raw membership count, and resolved/unsupported/unresolved/ambiguous counts. Inventory and resolution counts display `n/a` when retrieval failed.

## Dataverse requests

No query shape from Phase 2B changed. Each successfully started side uses:

1. one `WhoAmIRequest` to capture and verify its organization identity;
2. the existing Unique Name solution query;
3. the existing cookie-paged `solutioncomponent` query ordered by `solutioncomponentid`;
4. the existing cached/batched identity resolver requests required by that snapshot.

The dynamic Connection Reference discovery query and all Phase 2B batching, collision, unresolved, and duplicate safeguards remain unchanged.

## Automated coverage and live assumptions

UI-independent tests cover command gating, source-only and target-only selection eligibility, portable-key matches, genuine missing evidence, all three indeterminate identity statuses, duplicate identities, unavailable environments, diagnostics counts, and the live overload's single-`WhoAmI` context reuse and progress stages. Existing Phase 2A/2B retrieval, paging, cancellation, request-count, batching, and safeguard tests remain in the suite.

Automated tests use fake `IOrganizationService` implementations. Live validation is still required for XrmToolBox work-dialog cancellation behavior, Dataverse permissions and service-protection behavior, environment-specific Connection Reference type discovery, metadata response shape, real latency, and the usability of the results window with large inventories. Synchronous SDK calls cannot be interrupted while a request is in flight; cancellation is enforced at the next request boundary.

No component-definition comparison, membership export, root-component-behavior comparison, or environment-local root GUID comparison is included.
