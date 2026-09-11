# ReportExtract

ReportExtract is a portable Windows desktop application for filtering, validating, processing, and exporting structured Excel worksheets.

It is designed to simplify repetitive worksheet-processing tasks by allowing users to select a worksheet, apply reusable filters, calculate MAVL status, and export only the relevant records into a new Excel workbook.

## Features

- Import `.xlsx` workbooks
- Select a worksheet for processing
- Optional inclusion of hidden rows and columns
- Validate required worksheet fields before processing
- Filter by:
  - `P+F part number`
  - `Category`
  - `Manufacturer`
- Save and reuse filter presets
- Search available filter values
- Automatically calculate MAVL status
- Export only rows matching the active filters
- Preserve useful worksheet characteristics including:
  - cell values and value types
  - number formats and basic cell formatting
  - column widths
  - row heights
  - header formatting
  - freeze panes
- Export progress indication
- Manual export cancellation
- Portable application storage with no installer required

## Required Worksheet Columns

The selected worksheet must contain the following headers:

```text
P+F part number
Category
Manufacturer
```

Header matching ignores leading/trailing whitespace and letter case, but otherwise requires the expected header names.

### Empty Cells

Blank cells are permitted in general worksheet data.

The following fields are required for every eligible data row:

```text
P+F part number
Manufacturer
```

A blank, empty, or whitespace-only value in either field causes worksheet validation to fail.

`Category` may be blank. Blank categories are available as `(Blank)` in the filter interface.

Other worksheet columns may contain empty cells.

## MAVL

ReportExtract generates a `MAVL` column immediately to the right of `P+F part number`.

MAVL is calculated using the complete eligible worksheet before user filters are applied.

For each `P+F part number`:

```text
More than one distinct Manufacturer -> Yes
One distinct Manufacturer           -> No
```

Repeated rows from the same manufacturer do not cause MAVL to become `Yes`.

Category does not affect MAVL calculation.

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

Leading and trailing whitespace is ignored for comparison purposes.

## Filter Presets

Filter selections can be saved as presets and reused later.

Presets store selections for:

```text
P+F part number
Category
Manufacturer
```

If a preset contains a value that is not available in the currently loaded worksheet, ReportExtract warns the user rather than silently removing the selection.

## Export

ReportExtract creates a new `.xlsx` workbook containing only rows that match the active filters.

The source workbook is never modified.

The exported worksheet preserves source values and their underlying Excel value types where practical while retaining useful layout information such as column widths, row heights, formatting, and freeze panes.

Advanced workbook structures such as formulas as formulas, conditional formatting rules, named ranges, and complex table metadata are not the primary focus of the export.

If no rows match the selected filters, ReportExtract displays a notification and does not create an output workbook.

## Portable Distribution

ReportExtract is distributed as a self-contained Windows x64 application.

No installation of the following is required:

- Visual Studio
- .NET SDK
- Python
- development tools

### Running ReportExtract

1. Download the latest `ReportExtract-win-x64.zip` from the GitHub Releases page.
2. Extract the ZIP completely.
3. Open the extracted `ReportExtract-win-x64` folder.
4. Run:

```text
ReportExtract.exe
```

Do not run the application directly from inside the ZIP archive.

The top-level `ReportExtract.exe` launcher starts the application contained in the `Backend` directory.

The `Backend` folder should not be moved, renamed, or separated from the launcher.

## Portable Data Storage

ReportExtract stores writable application data inside the portable application directory:

```text
ReportExtract-win-x64/
├── ReportExtract.exe
├── Backend/
└── Data/
    ├── settings.json
    ├── Temp/
    └── Logs/
```

The `Data` directory contains:

- saved filter presets and application preferences
- temporary DuckDB session data
- performance logs

This allows ReportExtract to be removed cleanly by deleting the extracted application folder.

The application should be extracted to a location where the current user has write permission, such as `Documents` or `Desktop`.

Avoid placing the portable application inside protected locations such as `Program Files`.

## Previous Settings

When possible, ReportExtract can migrate settings from earlier versions stored under:

```text
%LOCALAPPDATA%\ReportExtract
```

Existing legacy settings are used only for migration. New application state is stored in the portable `Data` directory.

## Technology

ReportExtract is built with:

- C#
- .NET 8
- WPF
- ClosedXML
- DuckDB

## Building from Source

### Requirements

For development:

- Windows
- .NET 8 SDK
- Visual Studio or another suitable .NET development environment

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
└── ReportExtract.sln
```

## Supported Platform

Current release target:

```text
Windows x64
```

The distributed release is self-contained and does not require a separately installed .NET runtime.

## Usage and Support

ReportExtract is provided as a general-purpose utility for processing structured Excel worksheets.

Support, maintenance, updates, and compatibility with future operating-system or dependency changes are not guaranteed.

Use of the software is at your own risk. Always retain appropriate backups of important source workbooks and exported data.
