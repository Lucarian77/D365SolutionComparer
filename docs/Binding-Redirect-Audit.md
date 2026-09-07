# Binding redirect audit

Audit date: 2026-09-07. Source baseline: `4ca0601`. Decision: **leave app.config unchanged**.

## Evidence and method

1. Read the production project's direct PackageReferences and the `obj/project.assets.json` .NET Framework 4.8 target to identify selected runtime assets.
2. Inspect each selected package DLL's assembly manifest with the installed .NET Framework 4.8 `ildasm.exe /text /nobar`. Use `AssemblyVersion`, not NuGet version or Windows file version, for redirect comparison.
3. Compute SHA-256 of each package DLL and compare it with the pre-build local Release copy. All DLLs listed below matched exactly.
4. Inspect source `app.config` and the existing local `D365SolutionComparer.dll.config`; their redirects matched.

## Redirects against final resolved assets

All existing oldVersion ranges begin at `0.0.0.0` and end at the current target version shown below.

| Assembly | NuGet package version | Actual AssemblyVersion | Redirect target | Assessment |
| --- | --- | --- | --- | --- |
| System.Runtime.CompilerServices.Unsafe | 6.0.0 | 6.0.0.0 | 6.0.0.0 | Matches |
| System.Text.Json | 8.0.6 | 8.0.0.6 | 8.0.0.6 | Matches |
| Newtonsoft.Json | 13.0.1 | 13.0.0.0 | 13.0.0.0 | Matches |
| Microsoft.IdentityModel.Clients.ActiveDirectory | 5.3.0 | 5.3.0.0 | 5.3.0.0 | Matches |
| McTools.Xrm.Connection | MscrmTools.Xrm.Connection 1.2025.9.64 | 1.2025.9.64 | 1.2025.7.63 | Target differs from selected asset |
| McTools.Xrm.Connection.WinForms | MscrmTools.Xrm.Connection 1.2025.9.64 | 1.2025.9.64 | 1.2025.7.63 | Target differs from selected asset |

Other verified assets: Open XML package `2.13.1` contains assembly `2.13.1.0`; XrmToolBoxPackage `1.2025.10.74` contains Extensibility and ToolLibrary assemblies `1.2025.10.74`.

The XrmToolBox package declares a connection-package dependency minimum of `1.2025.7.63`, while the direct PackageReference selects `1.2025.9.64`. The inspected Extensibility assembly itself references `McTools.Xrm.Connection` assembly `1.2025.9.64`.

The two older redirect ranges do **not** include requests for `1.2025.9.64`. This is a documented discrepancy, not evidence that the validated plugin is currently failing. Whether the host applies plugin `.dll.config` redirects also requires runtime evidence. The unchanged `.nuspec` does not package the plugin configuration file.

Microsoft notes that plugin hosts may or may not honor `.dll.config`, and binding ultimately depends on the host/AppDomain configuration. See [Redirect assembly versions](https://learn.microsoft.com/en-us/dotnet/framework/configure-apps/redirect-assembly-versions#redirect-versions-for-tests-plugins-or-libraries-used-by-another-component).

## Package DLL SHA-256 values

| DLL | SHA-256 |
| --- | --- |
| DocumentFormat.OpenXml.dll | `1d2253ef392406a366d865fee651354a40123e46de0910bd6b5817884af3ce03` |
| Microsoft.IdentityModel.Clients.ActiveDirectory.dll | `99b4ca3049fbb0fff1456c3d89bc01fe04bbd8ddeb221d13b8f749195f01fd81` |
| McTools.Xrm.Connection.WinForms.dll | `c4ad36a9bed0d0d5d946fe2faf6a4ffe739206983d34fe6d651a2a29416c8e91` |
| McTools.Xrm.Connection.dll | `dcbdd6e0ffd5bdae26e4491cdad1b209589e90da4f2e6c0950f407d009ff14cc` |
| Newtonsoft.Json.dll | `b624949df8b0e3a6153fdfb730a7c6f4990b6592ee0d922e1788433d276610f3` |
| System.Runtime.CompilerServices.Unsafe.dll | `37768488e8ef45729bc7d9a2677633c6450042975bb96516e186da6cb9cd0dcf` |
| System.Text.Json.dll | `0cf5a9763c98f09e94dfcdaa4437e9508eabe2e50b408ba421d5faf734c5556a` |
| XrmToolBox.Extensibility.dll | `c81e47028e306fc07b6ff9d092ec369c957e906692e5a605659a1d5b7de9dee9` |
| XrmToolBox.ToolLibrary.dll | `84286503fad0a1b12f1732b35753a684c26e0807bb96e5451cb7866eab4b6c9d` |

## Before considering any redirect change

Reproduce an actual binding problem in a supported XrmToolBox installation. Capture the requested assembly identity, loaded assembly identities/paths, effective host configuration and relevant binding logs. Distinguish the deployed package from Debug output. Test a targeted change in that host before altering working redirects. Do not edit host configuration or add runtime assembly-resolution handlers as part of this phase.

The separate test project generates its own binding redirects using `AutoGenerateBindingRedirects` and `GenerateBindingRedirectsOutputType`. This isolates test-runner requirements from the production configuration.
