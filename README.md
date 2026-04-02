# Purpose
D365 Solution Comparer is a custom XrmToolBox plugin used to compare solutions across a source and a target Microsoft Dynamics 365 / Dataverse environment.

## Current Key Features
- Load source environment solutions
- Connect to target environment
- Load target environment solutions
- Compare source and target solutions
- Compare by:
  - Solution Unique Name
  - Display Name
  - Publisher
  - Version
  - Package Type (Managed / Unmanaged)
- Multi-select result filtering
- Package Type Differences filter
- Managed/Unmanaged-only checkbox filter
- Color-coded comparison results
- Summary counts
- Excel export of visible filtered rows

## Main Comparison Statuses
- Match
- Version Mismatch
- Publisher Mismatch
- Display Name Mismatch
- Managed/Unmanaged Mismatch
- Multiple Differences
- Missing in Source
- Missing in Target

## Typical Use
1. Open the plugin in XrmToolBox.
2. Load the source solutions.
3. Connect to the target environment.
4. Load the target solutions.
5. Run Compare.
6. Use the filters as needed.
7. Export the visible results to Excel if required.
