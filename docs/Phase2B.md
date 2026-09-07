# Phase 2B: membership resolution performance and live-readiness

Starting checkpoint: `fc23a9e7647d7bf273ea17bd3d3a92262db878ce` on `codex/1.2026.2.0`. This phase is uncommitted pending review. It changes only membership retrieval/resolution infrastructure, tests, project source inclusions, and this document.

## Operation and cache boundaries

`DataverseSolutionMembershipOperation.ReadAndResolve` is the normal bulk path. It creates one `DataverseReadContext`, verifies the service organization once, reads the selected solution and its complete membership, then resolves that snapshot with the same verified context. The overload accepting EnvironmentIdentity and Unique Name retains one-sided solution support. An absent solution returns without issuing resolver requests.

Existing APIs remain available. `DataverseSolutionMembershipReader.Read`, `DataverseComponentIdentityResolver.Resolve`, and `ResolveSnapshot` still create and verify their own operation contexts. Calling reader and resolver separately therefore still executes two WhoAmI requests. Standalone `Resolve` retains its single-object request shape, including primary-key equality and TopCount=2 for entity-backed components. The new bulk path is the optimized route for future UI integration.

Identity cache entries are scoped to one `ResolveSnapshot` operation and keyed by canonical component kind plus local object ID. Entries include Resolved, Unresolved, and Ambiguous outcomes. Repeated raw membership records produce separate ComponentIdentity instances retaining their own SolutionComponentRecord, but reuse one cached resolution. This preserves later duplicate canonical-key detection: duplicate memberships are still marked Ambiguous by SolutionMembershipComparer and cannot create false matches.

Connection Reference type discovery remains one request at most per resolution operation. Zero, incomplete, ambiguous, failed, and valid discovery results are all cached for that operation. Published built-in component codes remain excluded from reinterpretation, and a discovered code colliding with one of them remains Ambiguous.

No service, identity, metadata, or mapping result is cached across operations or environments. This avoids stale cross-environment identities and bounds retained data to the current read.

## Batching strategy

All calls remain sequential. No Task, parallel loop, ExecuteMultiple, retry, or throttling mechanism was introduced.

The bulk resolver deduplicates usable object IDs before sending requests. It groups entity-backed identities by table and sends QueryExpression requests with one primary-key `In` condition in chunks of at most 200 IDs. Each query selects only the primary key and portable identity field.

| Component family | Grouped query | Portable field |
| --- | --- | --- |
| Web Resources | webresourceid In (...) | name |
| Workflows / Processes | workflowid In (...) | uniquename |
| Security Roles | roleid In (...) | roletemplateid |
| Environment Variable Definitions | environmentvariabledefinitionid In (...) | schemaname |
| Connection References | connectionreferenceid In (...) | connectionreferencelogicalname |

The 200-ID chunk is deliberately below Dataverse's 500-condition query ceiling and keeps response sizes bounded. Microsoft documents that ConditionOperator.In supports unique identifiers in [Filter rows using QueryExpression](https://learn.microsoft.com/en-us/power-apps/developer/data-platform/org-service/queryexpression/filter-rows).

Tables use one `RetrieveMetadataChangesRequest` per 200 distinct table IDs. Its EntityQueryExpression filters EntityMetadata.MetadataId with `In` and requests only MetadataId and LogicalName. This is the SDK's supported grouped table-metadata query mechanism; Microsoft describes RetrieveMetadataChanges as a way to retrieve only selected schema definitions in [Query schema definitions](https://learn.microsoft.com/en-us/power-apps/developer/data-platform/query-schema-definitions).

Columns and relationships remain one metadata request per distinct object ID. Their solutioncomponent records provide child/relationship metadata IDs but no parent table metadata ID. A broad RetrieveMetadataChanges query would have to inspect every table to locate those children, increasing payload and server work. `RetrieveAttributeRequest` and `RetrieveRelationshipRequest` therefore remain the safer scoped operations until parent context is available or live measurements justify another approach.

Grouped responses are validated before becoming cache entries. Unrequested or empty IDs and impossible `MoreRecords` responses fail the operation. Duplicate returned IDs become Ambiguous. Requested IDs with no row, missing portable fields, or a caught Dataverse FaultException become Unresolved. A batch fault conservatively marks every identity in that batch Unresolved; it is never evidence of absence. Cancellation and unexpected exceptions propagate, and a canceled bulk resolution never returns a completed partial snapshot.

## Request instrumentation

`DataverseRequestCounter` is an optional operation-scoped instrument. It counts total Execute and RetrieveMultiple calls and exposes counts by SDK request name or queried entity logical name. It is thread-safe for observation but does not cause parallel execution, retain the service, inspect credentials, or alter responses. `DataverseSolutionMembershipOperation` and the bulk ResolveSnapshot overload accept it for deterministic request-count tests and future diagnostics.

## Request counts

Let N be the number of supported components with usable object IDs, U the number of distinct canonical-kind/object-ID pairs, P the number of solutioncomponent pages, D be 1 when Connection Reference discovery is needed and 0 otherwise, and B be the sum of grouped table/entity query chunks. Column and relationship distinct-ID counts are C and R.

Phase 2A separate read plus snapshot resolution used `N + P + 3 + D` requests: two WhoAmI requests, one solution lookup, P membership pages, N identity lookups, and optional discovery.

Phase 2B's coordinated operation uses `P + 2 + B + C + R + D`: one WhoAmI, one solution lookup, P membership pages, grouped requests, distinct column/relationship requests, and optional discovery. Repeated IDs reduce U without removing raw records.

The test suite records these representative cases with one membership page:

| Snapshot | Phase 2A estimate | Phase 2B measured | Reduction |
| --- | ---: | ---: | ---: |
| 100 unique Web Resources | 104 | 4 | 100 requests (96.2%) |
| 500 unique Web Resources | 504 | 6 | 498 requests (98.8%) |
| 100 mixed across all eight supported families | 105 | 36 | 69 requests (65.7%) |
| 500 mixed across all eight supported families | 505 | 136 | 369 requests (73.1%) |

The mixed distribution cycles through table, column, relationship, web resource, workflow, security role, environment variable definition, and Connection Reference. Connection Reference discovery accounts for the extra request in both before/after estimates. At these sizes each grouped mixed family fits one 200-ID chunk; columns and relationships dominate the remaining calls.

## Dataverse request changes

Unchanged operation-level reads:

- WhoAmIRequest validates OrganizationId. The coordinated operation issues it once.
- solution is selected by Unique Name with TopCount=2.
- solutioncomponent remains cookie-paged, ordered by solutioncomponentid, and selects the same raw fields.
- solutioncomponentdefinition discovery remains filtered by primaryentityname = connectionreference with TopCount=2.

Changed bulk identity reads:

- The five entity-backed primary-key equality queries become primary-key `In` queries, up to 200 IDs, with no TopCount. Because primary keys are unique and at most 200 rows can match, MoreRecords is treated as an invalid response.
- Table `RetrieveEntityRequest` calls become grouped RetrieveMetadataChangesRequest calls filtered by up to 200 MetadataIds and limited to MetadataId and LogicalName.
- Column RetrieveAttributeRequest and relationship RetrieveRelationshipRequest remain unchanged except duplicate IDs are suppressed.
- Standalone Resolve retains the Phase 2A single-object requests.

No new runtime package or connection technology is introduced.

## File-by-file changes

| File | Change |
| --- | --- |
| Infrastructure/DataverseRequestCounter.cs | Added optional operation request counters by request and entity name |
| Services/Membership/DataverseReadContext.cs | Added optional instrumentation and exposes the verified operation service/environment to coordinated internal paths |
| Services/Membership/DataverseSolutionMembershipReader.cs | Added internal overloads that reuse a verified context; public behavior remains unchanged |
| Services/Membership/DataverseComponentIdentityResolver.cs | Added operation-local result cache, duplicate suppression, grouped entity/table resolution, response validation, and an instrumented bulk overload |
| Services/Membership/DataverseSolutionMembershipOperation.cs | Added coordinated read-and-resolve entry points with one WhoAmI |
| D365SolutionComparer.csproj | Includes the two new production source files; dependencies unchanged |
| Tests/D365SolutionComparer.Tests/ComponentIdentityResolverTests.cs | Updated dynamic Connection Reference expectations for one grouped lookup |
| Tests/D365SolutionComparer.Tests/MembershipPerformanceTests.cs | Added duplicate-cache, grouped metadata, 100/500 request-count, mixed-family, fault, cancellation, and context-reuse tests |
| Tests/D365SolutionComparer.Tests/D365SolutionComparer.Tests.csproj | Includes the new performance test file; dependencies unchanged |
| docs/Phase2B.md | Records design, request changes, measurements, validation, and remaining risks |

## Validation and remaining risks

Executed 2026-09-07 with Visual Studio 2019 Professional MSBuild and VSTest 16.11.0, targeting .NET Framework 4.8 / Any CPU.

| Configuration | Build | Warnings | Errors | Tests |
| --- | --- | --- | --- | --- |
| Debug | Passed | 0 | 0 | 138 passed, 0 failed |
| Release | Passed | 0 | 0 | 138 passed, 0 failed |

Both runs include all 129 Phase 2A checkpoint cases plus nine Phase 2B cases. Data-driven cases cover the 100/500 entity and mixed snapshots separately. Automated tests use a fake IOrganizationService and validate query shape, grouping, counts, failure semantics, cancellation, and comparison safeguards. They do not prove live server latency, permissions, metadata behavior, or throttling characteristics.

The primary unresolved performance cost is one request per distinct column and relationship. Large snapshots dominated by these types still scale linearly. Grouped entity queries reduce round trips but may make a transient fault affect up to 200 identities; those entries remain explicitly Unresolved. No retry or batch bisection is included. Solutioncomponent paging and subsequent resolution are not a transactional snapshot, so server changes during the operation can still affect consistency. Live testing should measure response sizes, service-protection behavior, and RetrieveMetadataChanges permissions before UI integration.

Existing solution comparison, UI, exports, settings, nuspec, assembly version, app.config, packages, and membership presence semantics are unchanged. No membership UI is added.
