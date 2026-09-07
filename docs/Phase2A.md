# Phase 2A: functional membership services

Implementation checkpoint: `8bba3640a2f66b61681d8543f4650b6a3d5342c0` on `codex/1.2026.2.0`. The production regression baseline remains v1.2026.1.3, main commit `4ca0601`. This document describes uncommitted implementation for review; no commit or push is authorized yet.

## Read, resolve, compare

`DataverseSolutionMembershipReader` implements `ISolutionMembershipReader`. It verifies the service organization with WhoAmI, revalidates the solution by Unique Name, and reads all solutioncomponent pages with the existing `DataversePagedReader`. The supplied deterministic primary key is explicitly `solutioncomponentid`. The paging implementation is unchanged, including its case-insensitive duplicate-order check, query cloning, paging cookies, cancellation and progress behavior.

The interface overload accepts an existing SolutionIdentity and rejects a stale solution GUID. A concrete overload accepts EnvironmentIdentity and Unique Name, allowing independent lookup on a side that has no selected solution GUID. Neither overload fabricates identifiers.

Existing snapshot state names are retained:

| Solution state | Snapshot representation |
| --- | --- |
| Present, including zero members | Complete, after successful full retrieval |
| Absent | SolutionAbsent, after a successful zero-row solution lookup |
| Unavailable | Unavailable, with diagnostic and no inventory |

Strict `Read` methods propagate service failures, malformed responses, progress callback failures and cancellation. The explicit `ReadOrUnavailable` alternative converts non-cancellation failures to Unavailable. It never returns earlier pages as Complete, and it does not convert failures into SolutionAbsent. Cancellation still throws. Absence is established against the connected user's Dataverse view; these services cannot establish visibility beyond that user's permissions.

Every member retains its raw solutioncomponent ID, component type, nullable object ID, root component behavior, root solutioncomponent ID and nullable metadata flag. The snapshot supplies solution/environment provenance. Missing membership IDs or component types fail the inventory instead of producing invented values. Optional nulls and unknown numeric values are retained.

Retrieval initially wraps all raw records in Unresolved identities. `DataverseComponentIdentityResolver.ResolveSnapshot` creates a separate enriched snapshot without modifying the original. Unsupported types become explicitly Unsupported during this stage. `Resolve` also implements the existing per-record interface. Connection-type discovery is cached only within one resolution operation; no service or mapping is retained across environments.

`SolutionMembershipComparer` uses the existing comparison interface and accepts snapshots for the same Unique Name, compared with OrdinalIgnoreCase. It supports absent or unavailable sides. Both absent yields no component rows; callers retain the two snapshot states. Results describe recorded membership presence, not component definition equality or expanded effective membership.

## Resolver mapping strategy

Keys are compared case-insensitively within a canonical component kind. Raw local GUIDs are only lookup addresses. Display names are never fallback keys.

| Component type | Canonical kind | Portable identity used |
| --- | --- | --- |
| 1: table | table | EntityMetadata.LogicalName |
| 2: column | column | AttributeMetadata.EntityLogicalName + `.` + LogicalName; both required |
| 10: entity relationship | relationship | RelationshipMetadata.SchemaName |
| 61: web resource | webresource | webresource.name |
| 29: process/workflow | process | workflow.uniquename, when populated |
| 20: security role | securityrole | role.roletemplateid, when populated |
| 380: environment variable definition | environmentvariabledefinition | environmentvariabledefinition.schemaname |
| Discovered per environment: connection reference | connectionreference | connectionreference.connectionreferencelogicalname |

Metadata requests use the raw object ID as MetadataId and RetrieveAsIfPublished=false. Table retrieval requests EntityFilters.Entity. The column key requires its parent table to prevent equal column names on different tables from matching. Relationship support is limited to type 10; legacy relationship type 3 is not resolved in this phase.

Process Unique Name is optional; a process without it remains Unresolved. No workflow display-name or GUID fallback is used. See Microsoft's [workflow UniqueName reference](https://learn.microsoft.com/en-us/power-apps/developer/data-platform/reference/entities/workflow#uniquename).

Microsoft documents that template IDs, unlike local role IDs, are consistent across environments. Only template-backed roles are resolved here. Custom/template-less roles remain Unresolved, and multiple membership records with the same template key become Ambiguous in comparison. This does not claim equal privileges or business-unit scope. See [Security roles](https://learn.microsoft.com/en-us/power-apps/developer/data-platform/security-roles).

Connection-reference component type codes are discovered using solutioncomponentdefinition filtered by primaryentityname. They are not hardcoded. The Microsoft CoE team's [connection-reference mapping discussion](https://github.com/microsoft/coe-starter-kit/issues/1363) describes environment-specific codes and this lookup. The identity field is documented in the [connectionreference reference](https://learn.microsoft.com/en-us/power-apps/developer/data-platform/reference/entities/connectionreference).

The only model extension is optional `ComponentIdentity.ComponentTypeKey`, a canonical type scope. Existing callers default to the original numeric component-type scope. The result model now checks canonical type equality, allowing the same connection reference to match when numeric codes differ between environments. Raw numeric codes remain intact.

Unsupported types retain their original record and an Unsupported diagnostic. An unavailable/incomplete connection-type mapping leaves otherwise unknown types Unresolved, because their support cannot safely be determined; multiple mapping candidates produce Ambiguous. Missing object IDs, absent identity rows, incomplete names, and Dataverse identity-read faults produce Unresolved. Ambiguous lookups produce Ambiguous. Cancellation and unexpected non-service exceptions propagate rather than returning partial resolved snapshots.

## Comparison safeguards

- A match requires a resolved canonical kind and key on both sides.
- Duplicate resolved keys on either side become Ambiguous in result identities. All original records are retained; the engine does not choose the first candidate or mutate input snapshots.
- A one-sided classification requires a resolved member and either an explicitly absent opposite solution or a Complete opposite inventory with every identity resolved and unambiguous.
- Any unknown identity in an opposite present inventory conservatively blocks unproven missing classifications, across all component kinds. Independently established matches remain valid.
- Unsupported, Unresolved and Ambiguous records stay Indeterminate even against an absent solution. Unavailable snapshots cannot establish absence.
- Root behavior and metadata flags remain accessible on both matched raw records, but do not change presence or imply definition equality.

## Dataverse requests introduced

All requests are read-only. No joins, writes, ExecuteMultiple calls or all-column selections are introduced.

| Request | Columns / parameters | Predicate / behavior |
| --- | --- | --- |
| WhoAmIRequest | OrganizationId from response | Once per read or resolution operation; must match captured EnvironmentIdentity |
| QueryExpression: solution | solutionid, uniquename | uniquename = selected Unique Name; TopCount=2; reject multiple rows or MoreRecords |
| QueryExpression: solutioncomponent | solutioncomponentid, componenttype, objectid, rootcomponentbehavior, rootsolutioncomponentid, ismetadata, solutionid | solutionid = resolved solution GUID; solutioncomponentid ascending; cookie paging, default page size 5000 |
| RetrieveEntityRequest | MetadataId = objectid; EntityFilters.Entity; RetrieveAsIfPublished=false | Returns table logical name |
| RetrieveAttributeRequest | MetadataId = objectid; RetrieveAsIfPublished=false | Returns column and parent logical names |
| RetrieveRelationshipRequest | MetadataId = objectid; RetrieveAsIfPublished=false | Returns relationship schema name |
| QueryExpression: webresource | webresourceid, name | webresourceid = objectid; TopCount=2 |
| QueryExpression: workflow | workflowid, uniquename | workflowid = objectid; TopCount=2 |
| QueryExpression: role | roleid, roletemplateid | roleid = objectid; TopCount=2 |
| QueryExpression: environmentvariabledefinition | environmentvariabledefinitionid, schemaname | environmentvariabledefinitionid = objectid; TopCount=2 |
| QueryExpression: solutioncomponentdefinition | objecttypecode | primaryentityname = connectionreference; TopCount=2; once per operation when an unknown type needs mapping |
| QueryExpression: connectionreference | connectionreferenceid, connectionreferencelogicalname | connectionreferenceid = objectid; TopCount=2 |

TopCount lookups are bounded selection/identity queries, not inventory paging. Multiple rows or MoreRecords are treated as ambiguity. Only solutioncomponent uses DataversePagedReader. Identity resolution currently performs one metadata/record request per supported member; batching and additional caches are deferred.

## File-by-file changes

| File | Change |
| --- | --- |
| Services/Membership/DataverseReadContext.cs | Added operation-local organization verification and cancellation-aware read wrappers |
| Services/Membership/DataverseSolutionMembershipReader.cs | Added concrete retrieval, selection revalidation, one-sided lookup and explicit Unavailable alternative |
| Services/Membership/DataverseComponentIdentityResolver.cs | Added limited identity resolution and per-operation connection-type mapping |
| Services/Membership/SolutionMembershipComparer.cs | Added pure presence comparison, duplicate detection and conservative absence evidence |
| Models/Membership/ComponentIdentity.cs | Added optional canonical component kind while preserving raw type and prior defaults |
| Models/Membership/MembershipCompareResult.cs | Validate shared presence using canonical kind |
| Services/Contracts/ISolutionMembershipReader.cs | Clarified revalidation/absence documentation; signature unchanged |
| D365SolutionComparer.csproj | Include four new service files; dependencies unchanged |
| Tests/D365SolutionComparer.Tests/MembershipTestData.cs | Added fake response and snapshot fixtures |
| Tests/D365SolutionComparer.Tests/MembershipReaderTests.cs | Added 12 retrieval/state/failure/cancellation cases |
| Tests/D365SolutionComparer.Tests/ComponentIdentityResolverTests.cs | Added 18 resolver/mapping/failure/cancellation cases |
| Tests/D365SolutionComparer.Tests/MembershipComparisonTests.cs | Added 16 matching/one-sided/unknown/ambiguity cases |
| Tests/D365SolutionComparer.Tests/FakeOrganizationService.cs | Added Execute request callback and count for metadata and WhoAmI fixtures |
| Tests/D365SolutionComparer.Tests/D365SolutionComparer.Tests.csproj | Include four new test source files; test dependencies unchanged |
| docs/Phase2A.md | Implementation, queries, limitations, validation and review inventory |

## Validation

Executed 2026-09-07 with VS2019 Professional MSBuild and VSTest 16.11.0, targeting .NET Framework 4.8 / Any CPU.

| Configuration | Build | Warnings | Errors | Tests |
| --- | --- | --- | --- | --- |
| Debug | Passed | 0 | 0 | 113 passed, 0 failed |
| Release | Passed | 0 | 0 | 113 passed, 0 failed |

Each run contains all 67 checkpoint tests plus 46 new Phase 2A cases. Existing cases include 32 solution comparison/model, 15 membership model, and 20 paging regressions. New tests cover empty membership, three-page retrieval and raw-field preservation, stale/wrong-environment selection, absent solutions, both one-sided directions, all resolver families, unsupported/missing/ambiguous identities, duplicate keys, dynamic codes differing across environments, later-page faults, unavailable state, and cancellation.

Reproduce using the VS2019 commands in [Validation.md](Validation.md), replacing the results directory `obj/Phase1` with `obj/Phase2A`. Local ignored logs are `obj/Phase2A/Debug-build.log`, `Release-build.log`, `Debug-tests.trx`, and `Release-tests.trx`.

Preservation checks: HEAD remains the checkpoint; main and historical tag references remain unchanged. Production dependency targets/libraries match the pre-restore graph. No test frameworks/adapters/test assemblies or ClosedXML appear in production dependencies or output. Whitespace validation passes with `git -c core.whitespace=cr-at-eol diff --check` for the repository's existing line endings. The existing control, solution reader/comparer, exporters, settings, nuspec, assembly version, app.config and paged reader are unchanged.

These are fake-service and pure comparison tests, not live Dataverse or XrmToolBox certification. Live metadata/record visibility and connection-type discovery must be validated against supported environments before release. Cookie paging and later identity requests are not a transactional server snapshot; concurrent changes can still affect inventory consistency.

No membership UI or WorkAsync factory is added. Services execute synchronously on the calling thread, check cancellation between SDK calls, and cannot interrupt an in-flight SDK request. Future UI integration must call them through XrmToolBox WorkAsync, marshal progress, and reject stale completions after connection or selection changes. Effective membership expansion, definition comparison, exports and packaging changes remain outside this phase.
