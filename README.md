# D365 Solution Comparer

D365 Solution Comparer is a custom XrmToolBox plugin for Microsoft Dynamics 365 and Dataverse. It helps compare solutions between a source environment and a target environment so teams can quickly spot differences before deployments, validation, or troubleshooting.

## Overview

This tool was built to simplify solution comparison across environments by providing a clear side by side view of key solution metadata.

It compares:
- Solution Unique Name
- Display Name
- Version
- Publisher
- Package Type based on managed or unmanaged state

The plugin is designed for internal ALM checks, release validation, support reviews, and general environment comparison work.

## Features

- Load source environment solutions
- Connect to a target environment
- Load target environment solutions
- Compare source and target solutions
- Highlight differences by status
- Multi-select result filtering
- Package Type Differences filter
- Managed/Unmanaged-only checkbox filter
- Summary counts for visible results
- Excel export for visible filtered rows

## Comparison Statuses

The tool can identify and display these statuses:

- Match
- Version Mismatch
- Publisher Mismatch
- Display Name Mismatch
- Managed/Unmanaged Mismatch
- Multiple Differences
- Missing in Source
- Missing in Target

## Requirements

- Windows
- XrmToolBox
- Microsoft Dynamics 365 / Dataverse access
- .NET Framework 4.8
- Appropriate permissions to read solution metadata in the source and target environments

## How to Use
1. Open XrmToolBox.
2. Open **D365 Solution Comparer**.
3. Connect to your source environment.
4. Click **Load Source**.
5. Click **Connect Target** and connect to the target environment.
6. Click **Load Target**.
7. Click **Compare**.
8. Review the results grid, summary counts, and status coloring.
9. Use filters as needed:
   - status filter
   - package type differences
   - managed/unmanaged differences only
10. Export visible results to Excel if needed.

## Excel Export
The tool exports the currently visible filtered rows to Excel.

Export includes:
- Solution Unique Name
- Source Display Name
- Target Display Name
- Source Version
- Target Version
- Source Publisher
- Target Publisher
- Source Package Type
- Target Package Type
- Package Type Status
- Overall Status

## Repository
GitHub project:
`https://github.com/Lucarian77/D365SolutionComparer`

NuGet package:
`https://www.nuget.org/packages/D365SolutionComparer`

## Current Version
`1.0.0.1`

## Notes
- Package Type in the tool is derived from the solution managed state.
- If metadata changes do not appear after a rebuild, confirm that the latest DLL is being loaded and remove stale plugin manifest files if needed.
- NuGet package version and assembly version should remain aligned for clean update behavior.

## Author
Adrian Lucaci
