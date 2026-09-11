# ReportExtract

ReportExtract is a portable Windows desktop application for filtering, validating, processing, and exporting structured Excel worksheets.

It is designed to reduce repetitive worksheet-processing work by allowing users to select a worksheet, apply reusable filters, calculate MAVL status, and export only the relevant records into a new Excel workbook.

## Features

- Import `.xlsx` workbooks
- Select a worksheet for processing
- Optionally include hidden rows and columns
- Validate required worksheet structure before processing
- Filter by:
  - `P+F part number`
  - `Category`
  - `Manufacturer`
- Save and reuse filter presets
- Search available filter values
- Preserve unavailable preset selections and warn before export
- Automatically calculate MAVL status
- Export only rows matching the active filters
- Preserve useful worksheet characteristics, including:
  - source values and value types
  - basic cell formatting and number formats
  - column widths
  - row heights
  - header formatting
  - freeze panes
- Show export progress
- Cancel an export in progress
- Store application settings and temporary data inside the portable application folder
- Run as a self-contained Windows x64 application without requiring a separate .NET runtime

## Required Worksheet Columns

The selected worksheet must contain the following headers:

```text
P+F part number
Category
Manufacturer
```

Header matching ignores leading/trailing whitespace and letter case, but otherwise requires the expected header names.

A missing or duplicate required header causes worksheet validation to fail.

### Empty Cells

Blank cells are permitted in general worksheet data.

The following fields are required for every eligible data row:

```text
P+F part number
Manufacturer
```

A blank, empty, or whitespace-only value in either field causes worksheet validation to fail.

`Category` may be blank. Blank categories are exposed as `(Blank)` in the filter interface.

Other worksheet columns may contain empty cells.

Fatal worksheet-validation errors are shown in a modal error dialog and also retained in the application status area.

## MAVL

ReportExtract generates a `MAVL` column immediately to the right of `P+F part number`.

MAVL is calculated from the complete eligible worksheet before user filters are applied.

For each `P+F part number`:

```text
More than one distinct Manufacturer -> Yes
One distinct Manufacturer           -> No
```

Repeated rows from the same manufacturer do not cause MAVL to become `Yes`.

`Category` does not affect MAVL calculation.

## Filtering

Filtering uses the following rules:

- Multiple selections within the same field use **OR**
- Different filter fields use **AND**
- No selections for a field means that field is not filtered

Example:

```text
Category:
Electrical OR Mechanical

Manufacturer:
Manufacturer A OR Manufacturer B
```

is evaluated as:

```text
(Electrical OR Mechanical)
AND
(Manufacturer A OR Manufacturer B)
```

Leading and trailing whitespace is ignored for analytical comparison. Internal whitespace and case are otherwise preserved.

## Filter Presets

Filter selections can be saved as presets and reused later.

Presets store selections for:

```text
P+F part number
Category
Manufacturer
```

ReportExtract also persists the last-used filter selections and basic application preferences.

If a preset contains a value that is not available in the currently loaded worksheet, ReportExtract preserves the selection and warns the user rather than silently discarding it.

## Export

ReportExtract creates a new `.xlsx` workbook containing only rows that match the active filters.

The source workbook is never modified.

The exported worksheet reconstructs the filtered result while preserving source values and their underlying Excel value types where practical, along with useful layout information such as formatting, column widths, row heights, and freeze panes.

Complex workbook structures such as conditional-formatting rules, named ranges, data validation, tables, and other advanced workbook metadata are not the primary focus of the export.

If no rows match the selected filters, ReportExtract displays a notification and does not create an output workbook.

## Download and Run

ReportExtract is distributed as a self-contained Windows x64 package.

No installation of Visual Studio, the .NET SDK, Python, or other development tools is required to run the published application.

1. Download `ReportExtract-win-x64.zip` from the GitHub Releases page.
2. Extract the ZIP completely.
3. Open the extracted `ReportExtract-win-x64` folder.
4. Run:

```text
ReportExtract.exe
```

Do not run the application directly from inside the ZIP archive.

The top-level `ReportExtract.exe` is a lightweight launcher that starts the actual application from the `Backend` folder. Keep the complete extracted folder together.

## Portable Folder Layout

The published package uses the following layout:

```text
ReportExtract-win-x64/
├── ReportExtract.exe
├── Backend/
│   ├── ReportExtract.exe
│   ├── .NET runtime files
│   ├── DuckDB dependencies
│   └── other application dependencies
└── Data/
    ├── settings.json
    ├── Temp/
    └── Logs/
```

The `Data` directory contains writable application state:

- `settings.json` — presets and application preferences
- `Temp/` — temporary DuckDB session data
- `Logs/` — bounded performance/runtime logs

`settings.json` is created when settings are first saved.

ReportExtract does not require LocalAppData for its active settings storage.

Because the application writes to its own `Data` directory, extract it to a location where the current user has write permission, such as `Documents` or `Desktop`.

Avoid placing the portable application inside protected locations such as `Program Files`.

Deleting the extracted ReportExtract folder removes the application's portable settings, temporary data, logs, and runtime files together.

## Technology

ReportExtract is built with:

- C#
- .NET 8
- WPF
- ClosedXML
- DuckDB

## Building from Source

### Requirements

For normal development:

- Windows
- .NET 8 SDK
- Visual Studio or another suitable .NET development environment

Build the solution with:

```powershell
dotnet build ReportExtract.sln
```

## Creating the Portable Windows Package

The repository includes:

```text
publish-win-x64.ps1
Launcher/ReportExtractLauncher.c
```

These files work together to create the distributable package.

`publish-win-x64.ps1`:

- restores the `win-x64` runtime packages when required
- publishes ReportExtract as a self-contained .NET 8 Windows application
- places the application runtime in `Backend`
- compiles `ReportExtractLauncher.c` into the top-level `ReportExtract.exe`
- creates the portable `Data` directories
- validates required DuckDB/runtime files
- creates `dist/ReportExtract-win-x64.zip`

Building the portable launcher requires Visual Studio Build Tools with the **Desktop development with C++** workload.

From the repository root:

```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
.\publish-win-x64.ps1
```

Generated output:

```text
dist/
├── ReportExtract-win-x64/
└── ReportExtract-win-x64.zip
```

The `dist` directory is generated output and should not be committed to source control.

## Repository Structure

```text
ReportExtract/
├── Launcher/
│   └── ReportExtractLauncher.c
├── ReportExtract/
│   ├── Models/
│   ├── Services/
│   ├── Utilities/
│   ├── ViewModels/
│   ├── Views/
│   └── ReportExtract.csproj
├── README.md
├── ReportExtract.sln
└── publish-win-x64.ps1
```

## Supported Platform

Current release target:

```text
Windows x64
```

The published release is self-contained and does not require a separately installed .NET runtime.

## Usage and Support

ReportExtract is provided as a general-purpose utility for processing structured Excel worksheets.

Support, maintenance, future updates, and compatibility with future operating-system or dependency changes are not guaranteed.

Always retain appropriate backups of important source workbooks and exported data.
