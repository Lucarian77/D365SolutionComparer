# Validated production baseline

Recorded 2026-09-07 for the 1.2026.2.0 Phase 0/1 development work.

## Source identity

- User-designated validated production baseline: **v1.2026.1.3**.
- Baseline commit: `4ca0601f923574d7a0fd9997b5bd386cb148f074`.
- Commit subject: `Synchronize stable v1.2026.1.3 production baseline`.
- Development branch: `codex/1.2026.2.0`, starting at that commit.
- Historical tag `1.2026.1.3`: `ff524ce09d57f4c258c0d3dd3bc3cfb7ef1937c6`.
- The historical tag is not the validated source snapshot. Do not move it or use it as the regression checkout. Neither main nor historical tags are changed by this work.

The validated-production designation comes from the project owner. This phase does not independently certify the deployed production package. Before the development build, the existing local `bin/Release/D365SolutionComparer.dll` declared file version `1.2026.1.3` and had SHA-256 `e002b631492f6544f79ae24afef0efd043c863eb10d7901b240674e8dfddbb90`. This identifies the inspected local artifact, not a verified published NuGet package. Development builds overwrite local build outputs; the commit is the durable source baseline.

## Required platform and dependencies

Keep the traditional Visual Studio 2019 C# project, .NET Framework 4.8, WinForms, and `MultipleConnectionsPluginControlBase`.

| Direct PackageReference | Version |
| --- | --- |
| DocumentFormat.OpenXml | 2.13.1 |
| MscrmTools.Xrm.Connection | 1.2025.9.64 |
| XrmToolBoxPackage | 1.2025.10.74 |

The assembly and package versions remain `1.2026.1.3` during this foundation phase. No ClosedXML dependency is permitted. See [the binding audit](Binding-Redirect-Audit.md).

## Observable behavior to preserve

- Source solutions load from the primary XrmToolBox connection; target solutions use the additional connection.
- Existing solution retrieval selects visible solutions and is unchanged in this phase. The new paged reader is not wired into it.
- Solutions match using Unique Name with `OrdinalIgnoreCase`; keys are not trimmed. Display Name and local GUIDs are not comparison keys.
- Null input lists behave as empty inventories in the pure comparison service. The existing UI still requires nonempty loaded lists.
- Duplicate keys keep the first record on each side. Null and empty keys group together. Preserve these baseline quirks; do not reuse them for future ambiguous component identities.
- Compared display name, version, publisher and derived package type ignore case and surrounding whitespace. Versions are strings, not semantic versions.
- A single difference produces its field status; multiple differences produce `Multiple Differences`. Package Type Status remains independent.
- Managed state is nullable: unknown displays blank, unknown versus known currently produces `Managed/Unmanaged Mismatch`, and two unknown values match.
- Results retain original source/target text, using empty strings for null values, and sort by Unique Name with `OrdinalIgnoreCase`.
- Status multi-selection, Changed only, managed/unmanaged-only, Reset Filters, summary counts, row details, and saved filter preferences retain their existing behavior.
- Package Type Differences is an additional intersection with selected statuses, not another union status.
- Existing source connection/reload behavior is not changed. Future membership integration must independently prevent stale connection results.
- XLSX remains the default export using raw Open XML SDK; Excel XML and CSV remain available. Export uses currently visible comparison rows and the existing 11 columns, metadata, formatting and text normalization.

Pure regression coverage and the remaining manual checks are recorded in [Validation.md](Validation.md). UI, export, connection and settings behavior has been preserved by leaving the implementation files unchanged; it has not been re-exercised in a live host during this phase.
