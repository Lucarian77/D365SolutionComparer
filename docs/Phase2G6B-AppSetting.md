# Phase 2G.6B AppSetting identity and structural definition comparison

Pending live validation; no Dataverse writes, version change or deployment is part of this phase.

## Evidence and selection

The user-provided Phase 2G.6A EDU validation established 15 unique AppSettings per
environment, exact objectid -> appsetting primary-key correlation, related
Setting Definition UniqueName and parent AppModule UniqueName, and matching composites
despite differing local IDs. That evidence authorizes the production identity.

A registered definition must have Name `AppSetting` and PrimaryEntityName
`appsetting`, both OrdinalIgnoreCase, without conflicting definitions for the raw
type. Equivalent repeated registrations collapse. The raw numeric type and
formatted labels never select the resolver. Supplied completed-snapshot
AppSetting registrations are reused; otherwise existing registered-family
discovery supplies them. Static platform types and Connection Reference discovery
retain their existing priority and behavior.

The semantic kind is `appsetting`, displayed by the standard presenter as
**App Setting**. The key is:

`appsetting:v1:<parent.Length>:<parent>:<definition.Length>:<definition>`

Both strings must be nonblank. Preserve their original content; equality uses
OrdinalIgnoreCase. Framing prevents delimiter collisions. There is no GUID,
display name, setting value or numeric component type in the key.

Missing/invalid references, inaccessible rows and incomplete/conflicting reads
remain Unresolved; duplicate row correlations are Ambiguous. Different membership
rows producing one case-insensitive key become Ambiguous with a PortableIdentity
blocker. The unchanged generic membership engine permits absence for unrelated
keys, but never for the blocked key or when the opposite kind has incomplete
unscoped identity coverage. The comparer itself is unchanged.

## Definition contract and limits

Microsoft documents the conceptual setting structure (data type, override level
and release level) and immutable publisher-prefixed setting names in
[Use settings to provide customized app experiences](https://learn.microsoft.com/en-us/power-apps/maker/data-platform/create-edit-configure-settings).
That page does not establish the exact deployed table schema. This implementation
does not treat conceptual labels or guessed logical names as query authorization.

Definition comparison, after identity resolution, uses one published
`RetrieveEntityRequest` for `settingdefinition`, filters
`Entity | Attributes`, verifies the exact entity/primary ID and each candidate
attribute's logical name, unique metadata result, readable flag and SDK type.
Only verified attributes are added to a separate structural backing-row query:

| Comparable property | Required runtime metadata type | Equality |
| --- | --- | --- |
| `datatype` | Integer or Picklist | invariant integer value |
| `isoverridable` | Boolean | boolean value |
| `overridablelevel` | Integer or Picklist | invariant integer value |
| `releaselevel` | Integer or Picklist | invariant integer value |

All four must be schema-verified and present with the matching runtime value type
for an Available definition. Missing/unreadable/wrong-type/duplicate metadata,
metadata faults, missing/null values or wrong runtime values produce Unresolved
definitions, never Match. A missing field is not defaulted. Identity resolution
always reads only `settingdefinitionid` and `uniquename` and completes before structural
metadata is requested. Identity therefore remains resolved when structural metadata
or the structural-property query is unavailable. Live validation must establish
which of these structural fields actually exist in DEV/UAT; no assertion of live
field availability is made by fake tests.

Excluded from comparison and new setting queries: `value`, `defaultvalue`, local
primary/reference IDs, solutioncomponentid, componentidunique, componentstate,
ismanaged, owners, timestamps and solution IDs. Display/description and information
URL fields are excluded because no localization/free-text policy is approved.
Identity names and reference IDs are correlation/audit only. No secret/default
value is used to disambiguate duplicate identities.

The new AppSetting path copies only allowlisted queried attributes into its
inventory; unsolicited value/defaultvalue response fields are discarded. It never
formats arbitrary response rows or server exception messages into diagnostics.
Fault diagnostics are stable and redact server details. Existing Coverage CSV and
definition detail presentation consume these safe properties/evidence unchanged.

## Requests and reuse

No separate Phase 2G.6A evidence capture runs during normal Compare Membership.
The existing coordinated read context verifies the environment once and counts
every new request. New reads use the shared bounded diagnostic reader:

| Entity/API | Columns/request | Filter |
| --- | --- | --- |
| `appsetting` | `appsettingid`, `settingdefinitionid`, `parentappmoduleid` | `appsettingid IN` distinct nonempty raw object IDs |
| `settingdefinition` identity | `settingdefinitionid`, `uniquename` | `settingdefinitionid IN` distinct related IDs |
| `settingdefinition` structure | `settingdefinitionid`, plus all four schema-verified contract fields | separate `settingdefinitionid IN` distinct related IDs; definition phase only |
| `appmodule` | existing AppModule projection: `appmoduleid`, `uniquename`, `name`, `appmoduleidunique`, `componentstate`, `ismanaged`, `description`, `clienttype`, `formfactor`, `navigationtype` | `appmoduleid IN` related IDs not already attempted by the normal resolver |
| `RetrieveEntityRequest` | `LogicalName=settingdefinition`, `EntityFilters=Entity|Attributes`, `RetrieveAsIfPublished=false` | once when a related Setting Definition is required |

GUIDs are actual Guid-valued IN operands. New ID lists are sorted, deduplicated
and split into batches of at most 200. Missing, duplicate, unexpected-entity,
conflicting-PK and MoreRecords responses retain conservative outcomes. All
new stages propagate cancellation. No partial success escapes cancellation.

AppSetting/Setting Definition outcomes (including missing/faulted results) and
related AppModule outcomes are cached for the operation. Normal AppModule results
are reused before resolving AppSetting parents. Definition comparison reads the
same inventory without additional requests. A standalone definition read without
that coordinated inventory is explicitly Unresolved, rather than claiming Match.

For A distinct AppSetting IDs, D distinct referenced Setting Definitions and P
parent IDs not previously attempted: additional cost is
`ceil(A/200) + 2*ceil(D/200) + ceil(P/200) + one schema request`, when A/D are usable
and structural metadata is verified. If schema verification fails, the structural
row read is skipped and definitions remain Unresolved without changing identity.
For EDU's 15 rows: **+4 per environment** if all parent apps were already read;
**+5** if up to 200 additional parents are required. Existing registered-family
discovery is reused, not counted as a new query. No additional WhoAmI occurs.
No Type 9, workflow, metadata-inventory or solutioncomponent query is changed.

## Evidence UI and regression scope

AppSetting uses the standard result and definition detail view. Promoted records
appear in the supported semantic coverage bucket; their registrations stay on
the raw component evidence for CSV/Capture AppSetting. They are not counted again
under isolated registered families.

Phase 2G.6A remains available. Its display now explains that a successful related
Setting Definition/AppModule was reached through the AppSetting reference, not
through direct objectid correlation. The underlying direct-correlation states
and evidence are unchanged.

`AppSettingResolutionTests` explicitly covers registered selection and conflicts,
synthetic numeric types, standalone/coordinated APIs, exact correlation, every
reference failure, framing/case/local-ID independence, duplicate scoped blockers,
one-sided and unavailable coverage, structural differences, metadata validation,
secret exclusion in CSV/details, cancellation, stable diagnostic grouping,
coverage reconciliation and request counts. Existing Phase 2G.1–2G.6A tests remain
unchanged.

## Live read-only validation

1. Reload the rebuilt Release plugin in XrmToolBox. Connect Source
   CSC-ICMS-DEV and Target CSC-ICMS-UAT, load EDU and run the normal solution Compare.
2. Run Compare Membership. Record requests/timing and overall membership totals.
3. Open Coverage Details. Expect App Setting 15 total/15 resolved per environment,
   no unsupported/unresolved/ambiguous AppSettings, and 15 Present in Both if EDU
   is unchanged. Reconcile actual totals rather than forcing historical totals.
4. Double-click AppSetting results. Verify framed keys use parent UniqueName and
   Setting Definition UniqueName, and only the four approved structural fields appear.
   If schema/data are incomplete, expect Unresolved definitions; capture that
   diagnostic rather than interpreting it as Match. Differences must list only
   approved fields. No value/default may appear anywhere.
5. Run Capture AppSetting manually. Confirm all 15 raw rows per environment,
   dynamic definition provenance and exact object/reference correlations remain
   inspectable. Successful related rows should have the corrected wording.
6. Export Coverage Details CSV; verify safe evidence and reconciled counts, with
   no new Dataverse requests. Check that values/defaults and server fault details
   are absent. Cancel an operation and verify no completed partial inventory.
7. Recheck Columns (DEV438/UAT382), Relationships (DEV59/UAT55), three Site Maps
   per environment, existing workflow ambiguity and unsupported families against
   the supplied live baseline. Normal solution filters/exports/lifecycle tools
   must retain their existing behavior. Make no Dataverse changes for this test.

Do not commit until the user reviews this implementation and completes live EDU validation.
