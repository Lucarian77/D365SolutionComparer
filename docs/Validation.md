# Phase 0/1 validation

Baseline: `4ca0601`; development branch: `codex/1.2026.2.0`. No commits or pushes are made as part of this implementation review.

## Reproduce with Visual Studio 2019

Run from the repository root in PowerShell. Adjust only the Visual Studio installation path if needed.

```powershell
$vs2019 = 'C:\Program Files (x86)\Microsoft Visual Studio\2019\Professional'
$msbuild = "$vs2019\MSBuild\Current\Bin\MSBuild.exe"
$vstest = "$vs2019\Common7\IDE\CommonExtensions\Microsoft\TestWindow\vstest.console.exe"
New-Item -ItemType Directory -Path obj\Phase1 -Force | Out-Null
& $msbuild D365SolutionComparer.sln /restore /t:Build /p:Configuration=Debug '/p:Platform=Any CPU'
& $vstest Tests\D365SolutionComparer.Tests\bin\Debug\D365SolutionComparer.Tests.dll '/TestAdapterPath:Tests\D365SolutionComparer.Tests\bin\Debug' '/Logger:trx;LogFileName=Debug-tests.trx' /ResultsDirectory:obj\Phase1 /Framework:.NETFramework,Version=v4.8
& $msbuild D365SolutionComparer.sln /restore /t:Build /p:Configuration=Release '/p:Platform=Any CPU'
& $vstest Tests\D365SolutionComparer.Tests\bin\Release\D365SolutionComparer.Tests.dll '/TestAdapterPath:Tests\D365SolutionComparer.Tests\bin\Release' '/Logger:trx;LogFileName=Release-tests.trx' /ResultsDirectory:obj\Phase1 /Framework:.NETFramework,Version=v4.8
```

The VS2019 Test Explorer can also discover the test project after restore/build. Build output and test results are under ignored bin/obj directories. The baseline Debug post-build step copies the plugin and Open XML DLL into the local output Plugins directory; it does not install or publish the plugin.

## Automated coverage

| File | Coverage |
| --- | --- |
| SolutionComparisonTests.cs | All 16 difference combinations; Unique Name case/whitespace; display-name independence; field normalization and retained text; version strings; duplicate/blank keys; null/empty and one-sided inventories; sorting; nullable managed state; model alias semantics |
| MembershipModelTests.cs | Local identities; required IDs; unknown raw type/behavior values; unresolved/unsupported/ambiguous safeguards; explicit absence evidence; both one-sided directions; complete-empty/absent/unavailable distinction; immutable collection; invalid states |
| DataversePagedReaderTests.cs | Query preservation; deterministic order; cookies; cumulative progress; complete-empty retrieval; cancellation; SDK/callback faults; no partial success; missing/repeated cookies; null/nonadvancing response; input restrictions |

## Execution record

Executed 2026-09-07 using Visual Studio Professional 2019 16.11.59, its MSBuild, and VSTest 16.11.0.

| Configuration | Build | Warnings | Errors | Tests |
| --- | --- | --- | --- | --- |
| Debug / Any CPU | Passed | 0 | 0 | 61 passed, 0 failed |
| Release / Any CPU | Passed | 0 | 0 | 61 passed, 0 failed |

Each configuration ran 32 baseline comparison/model cases, 15 membership-model cases, and 14 paging cases. Build logs and TRX results are in `obj/Phase1/Debug-build.log`, `Release-build.log`, `Debug-tests.trx`, and `Release-tests.trx` (local ignored artifacts).

Preservation checks passed:

- Main still points to `4ca0601f923574d7a0fd9997b5bd386cb148f074`; historical tag `1.2026.1.3` still points to `ff524ce09d57f4c258c0d3dd3bc3cfb7ef1937c6`.
- Existing source files, settings, exporters, app.config, nuspec and assembly version are unchanged. The only modified pre-existing files are the project and solution manifests, adding source inclusions and the test project.
- Production `project.assets.json` dependency targets and libraries are identical to the captured pre-restore graph. No test packages or ClosedXML appear in that graph.
- No test framework, adapter, test-platform or test-assembly DLLs appear in the production Debug/Release output or Debug Plugins directory. Production package file entries remain unchanged.
- `git -c core.whitespace=cr-at-eol diff --check` passed. This treats the existing project/solution CRLF line endings as line endings while checking whitespace; their original newline conventions are preserved.

## Manual validation still required before a release

These checks have not been performed in this phase and are not implied by passing unit tests:

- Live XrmToolBox discovery, primary/additional connection transitions, failed reloads and actual host assembly bindings.
- Existing filter combinations, Reset Filters, summaries, row details, button enablement and resizing.
- Existing settings restart, corrupt-file and unwritable-file behavior.
- Opening filtered XLSX and Excel XML exports in Excel, including styles, frozen rows and filters; CSV encoding, quoting and visible rows.
- Actual multi-page Dataverse retrieval and permission failures. Tests use a fake service; they are not a live integration certification.
- Published package installation/update and production package identity.

The phase deliberately leaves the existing UI/export/settings/retrieval/comparison implementations and production packaging/version/configuration untouched. The full v1.2026.1.3 behavior matrix is documented in [Baseline-1.2026.1.3.md](Baseline-1.2026.1.3.md).
