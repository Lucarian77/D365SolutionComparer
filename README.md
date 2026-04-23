# D365 Solution Comparer

D365 Solution Comparer is a custom XrmToolBox plugin for Microsoft Dynamics 365 and Dataverse. It helps compare solutions between a source environment and a target environment so teams can quickly spot differences before deployments, validation, troubleshooting, and general ALM review.

## Overview

This tool provides a side by side comparison of solution metadata across two environments and highlights the differences that matter most during release validation and support work.

It compares:
- Solution Unique Name
- Display Name
- Version
- Publisher
- Package Type based on managed or unmanaged state

The plugin is designed for deployment validation, release readiness checks, ALM drift review, support investigations, and general environment comparison work.

## Features

- Load source environment solutions from the current XrmToolBox connection
- Connect to a separate target environment
- Load target environment solutions
- Compare source and target solutions side by side
- Highlight differences by status with color-coded results
- Multi-select status filtering
- Changed only quick filter
- Managed/unmanaged differences only filter
- Reset Filters button
- Summary counts for visible results
- Double-click row details popup
- About dialog with version and dependency information
- Export visible filtered rows
  - Excel export when the required runtime dependencies are available
  - CSV export fallback when Excel dependencies are not available in the current XrmToolBox host

## Comparison Statuses

The tool can identify and display these statuses:

- Match
- Version Mismatch
- Publisher Mismatch
- Display Name Mismatch
- Package Type Mismatch
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
3. Connect XrmToolBox to your source environment.
4. Click **Load Source**.
5. Click **Connect Target** and connect to the target environment.
6. Click **Load Target**.
7. Click **Compare**.
8. Review the results grid, summary counts, and status coloring.
9. Use filters as needed:
   - status filter
   - Changed only
   - Managed/unmanaged differences only
   - Reset Filters
10. Double-click any comparison row to open detailed row information.
11. Export the visible results when needed.
   - On hosts where Excel export is available, the tool exports an Excel workbook.
   - On hosts where Excel dependencies cannot be validated, the tool uses CSV export instead.

## Export Output

The export includes the currently visible rows and includes:

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

When Excel export is available, the tool generates a formatted workbook with report metadata and filters applied.  
When CSV export is used, the file is plain text and final column sizing or styling depends on the application used to open it.

## Notes

- Package Type in the tool is derived from the solution managed state.
- Filter state is saved between sessions.
- On some machines, Excel export may not be available because of dependency resolution in the shared XrmToolBox host. In those cases the tool remains fully usable and exports CSV instead.
- NuGet package version and assembly version should remain aligned for clean update behavior.

## Repository

GitHub project:  
`https://github.com/Lucarian77/D365SolutionComparer`

NuGet package:  
`https://www.nuget.org/packages/D365SolutionComparer`

## Current Version

`1.1.0.0`

## Author

Adrian Lucaci

## License

MIT
