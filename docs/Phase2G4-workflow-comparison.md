# Phase 2G.4 - Process / Workflow comparison

Status: implementation and automated validation only; live EDU validation is pending. No identity claim is based on matching environment-local GUIDs.

## Identity contract

- Family remains `process` / Process / Workflow, raw type 29.
- Existing nonblank `uniquename` remains the preferred key, using `StringComparer.OrdinalIgnoreCase`.
- A blank-name activation (`type=2`) follows its verified `parentworkflowid` to a definition (`type=1`). The parent supplies identity and definition configuration. The pointer itself is audit-only. No recursion through arbitrary process types.
- For a blank `uniquename` definition, the separately identified semantic candidate consists of **name + type=1 + category + nonblank primaryentity + applicable subtype**.
- Applicable subtype: category 4 uses businessprocesstype; 5 uses modernflowtype; 6 uses uiflowtype. Currently recognized subtype values are 0 and 1. Categories 0-7 are recognized, including the documented AI Flow category; newer/unknown categories remain unresolved for fallback and configuration coverage.
- Candidate encoding is `workflow-semantic:v1:` followed by length-framed exact field values. No whitespace, punctuation or meaningful characters are removed. Equality is OrdinalIgnoreCase. Uniquenames colliding with the reserved prefix remain unresolved.
- Name alone, mode, subprocess, state, locale/deployment status and local IDs cannot resolve duplicate candidates.

This composite is an explicitly identified **solution-scoped semantic correlation rule**, not a documented Dataverse alternate key or a guarantee of ALM component lineage. A rename or a scope/category change can change a fallback candidate. Independently authored definitions can share this semantic evidence. Do not interpret a configuration match as proof of identical execution logic or historic component identity.

## Ambiguity and absence

Candidate uniqueness is checked across distinct correlated definition rows in the environment operation, including named definitions. Repeated raw references to the same backing row do not fabricate a new semantic candidate; existing duplicate membership-key safeguards still apply.

Multiple complete equal candidates remain Ambiguous. Incomplete same-name candidates, missing rows, duplicate raw correlations, missing object IDs or faulted batches prevent definitive fallback resolution when uniqueness cannot be established. Conflicting entity/primary IDs or unexpected paging fail the operation rather than returning partial success. Cancellation propagates.

Phase 2G.5 classifies a duplicate Process / Workflow fallback key as a `PortableIdentity` blocker carrying that key. It blocks absence only for the same key. A workflow with no trustworthy candidate carries a `SemanticKind` blocker and continues to block absence for the whole Process / Workflow kind. Other component families retain their previous kind-wide duplicate safeguards. This scope is evaluated by one authoritative `CanEstablishAbsence` path using ordinal case-insensitive keys.

Uniquename and semantic-fallback strategies are never silently paired. If opposite-side semantic evidence makes a differently keyed record a plausible counterpart, the Process-only comparison guard returns indeterminate results, not false one-sided results. Snapshot coverage describes per-environment resolution; this additional cross-strategy uncertainty is shown in the comparison row's reason.

The existing semantic-kind coverage rule still governs absence. Incomplete Process identity coverage blocks Source Only/Target Only for that family; SolutionAbsent proves absence, Unavailable does not. Other component families retain their existing behavior.

## Definition contract

Compared fields: `type`, `category`, `primaryentity`, `mode`, `subprocess`, `businessprocesstype`, `modernflowtype`, `uiflowtype`.

A verified type-1 definition, recognized category, nonblank entity scope, mode 0/1, Boolean subprocess, and any applicable recognized subtype are required. Inapplicable subtype properties are represented as null. Entity logical names are case-normalized; numeric options and Booleans use deterministic invariant representations. Available definitions contain every contract property. Missing required configuration remains Unresolved.

Names and identity keys do not automatically become equality fields. Workflow IDs, parent references, owner IDs, workflowidunique, componentstate, ismanaged, statecode and statuscode remain audit-only.

**Configuration-only coverage** is shown in definition diagnostics. `xaml`, `clientdata`, and `definition` are not requested or compared. Microsoft exposes payload fields, but no safe normalization of their environment-specific references has been established here. No GUID stripping, XML/JSON normalization or content hashing is applied, so workflow logic is not inadvertently erased from a purported full-logic comparison.

## Retrieval and request cost

Existing `workflow` QueryExpression: `workflowid IN (distinct Guid values)`, deterministic GUID batches of up to 200. Main projection is unchanged:

`workflowid, uniquename, name, type, category, primaryentity, mode, parentworkflowid, workflowidunique, statecode, statuscode, componentstate, ismanaged, subprocess, businessprocesstype, modernflowtype, uiflowtype`.

Parent lookup projection expands from workflowid/uniquename/type to those same 17 fields. Already read parents are reused. No new discovery or payload requests are introduced. Shared definition retrieval consumes only cached rows and makes no second workflow pass.

Workflow requests per environment = `ceil(D / 200) + ceil(P / 200)`, where D is the distinct usable raw workflow ID count and P is the distinct required parent IDs not already read during the operation. The coordinated operation retains one WhoAmI per environment. DEV 19 and UAT 25 each fit in one direct batch; parent costs depend on the live activation records. Zero incremental direct workflow requests; parent payload increases, and caching can remove redundant parent reads. Zero Dataverse writes.

## Microsoft evidence versus policy

[Microsoft Workflow table reference](https://learn.microsoft.com/en-us/power-apps/developer/data-platform/reference/entities/workflow) documents the fields, record types and categories. [Workflow Web API reference](https://learn.microsoft.com/en-us/power-apps/developer/data-platform/webapi/reference/workflow?view=dataverse-latest) provides their API representation. Neither is treated as proof that the composite above is a globally unique portable key.

## Read-only live validation

1. Load the rebuilt Release plugin; use CSC-ICMS-DEV as Source and CSC-ICMS-UAT as Target, solution EDU.
2. Run normal Compare and Compare Membership. Record solution results, raw counts, request counts and elapsed times.
3. Confirm Columns 438/382 and Relationships 59/55 remain fully resolved, or reconcile changed solution contents.
4. Confirm all three Site Maps remain Present in Both: ava_CaseManagementSystem Different, ava_ICMSSystemHelpdesk Match, msdyn_FSMobile Different, with sitemapxml as the known changed property.
5. Inspect every Process fallback key and subtype, original diagnostic fields, duplicate and missing-field reasons. Do not force the initial 18/20 unresolved populations to zero.
6. Inspect the two UAT Modern Flow names `[BPF] - Inspection - Approval from Inspection to Investigation` and `[BPF] - Investigation - Approval from Investigation to Prosecution`. Search all DEV Process evidence for counterparts, including differing identity strategies. Target Only is justified only by unique usable identity and complete opposite coverage; current supplied evidence cannot establish their absence.
7. Double-click Process rows. Verify comparable source/target values, mode/subprocess differences, audit-only GUID/state differences, and the configuration-only coverage notice.
8. Check Coverage Details, CSV/lifecycle functions, filters, cancellation and request counts. This phase performs no environment modifications or ALM lifecycle actions.

Do not mark complete, commit or push before review and live validation.
