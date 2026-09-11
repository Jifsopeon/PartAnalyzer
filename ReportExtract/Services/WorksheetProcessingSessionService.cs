using System.IO;
using ClosedXML.Excel;
using ReportExtract.Models;

namespace ReportExtract.Services;

public sealed class WorksheetProcessingSessionService
{
    private static readonly IReadOnlyDictionary<RequiredWorksheetColumn, string> RequiredHeaders =
        new Dictionary<RequiredWorksheetColumn, string>
        {
            [RequiredWorksheetColumn.PartNumber] = "P+F part number",
            [RequiredWorksheetColumn.Category] = "Category",
            [RequiredWorksheetColumn.Manufacturer] = "Manufacturer"
        };

    public Task<WorksheetProcessingSession> CreateAsync(
        string workbookPath,
        WorksheetInfo worksheetInfo,
        bool includeHiddenRowsAndColumns,
        CancellationToken cancellationToken = default)
    {
        return Task.Run(
            () => Create(workbookPath, worksheetInfo, includeHiddenRowsAndColumns, cancellationToken),
            cancellationToken);
    }

    private static WorksheetProcessingSession Create(
        string workbookPath,
        WorksheetInfo worksheetInfo,
        bool includeHiddenRowsAndColumns,
        CancellationToken cancellationToken)
    {
        using var measurement = PerformanceLogger.Measure("Worksheet validation", worksheetInfo.Name);
        if (string.IsNullOrWhiteSpace(workbookPath) || !File.Exists(workbookPath))
        {
            throw new WorksheetValidationException("The source workbook is no longer available.");
        }

        if (!string.Equals(Path.GetExtension(workbookPath), ".xlsx", StringComparison.OrdinalIgnoreCase))
        {
            throw new WorksheetValidationException("Choose an .xlsx workbook.");
        }

        try
        {
            using var stream = File.Open(workbookPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            using var workbook = new XLWorkbook(stream);
            var worksheet = workbook.Worksheets.FirstOrDefault(sheet => sheet.Name == worksheetInfo.Name)
                ?? throw new WorksheetValidationException("The selected worksheet could not be found in the workbook.");
            var range = worksheet.RangeUsed()
                ?? throw new WorksheetValidationException("The selected worksheet does not contain usable data.");
            var headerRow = range.FirstRowUsed()
                ?? throw new WorksheetValidationException("The selected worksheet does not contain a header row.");

            cancellationToken.ThrowIfCancellationRequested();
            if (worksheet.Row(headerRow.RowNumber()).IsHidden)
            {
                throw new WorksheetValidationException($"Header row {headerRow.RowNumber()} is hidden. Make it visible and try again.");
            }

            var firstColumnNumber = range.FirstColumn().ColumnNumber();
            var lastColumnNumber = range.LastColumn().ColumnNumber();
            var requiredColumns = ResolveRequiredColumns(worksheet, headerRow.RowNumber(), firstColumnNumber, lastColumnNumber);

            using var layoutMeasurement = PerformanceLogger.Measure("Effective worksheet layout", worksheetInfo.Name);
            var effectiveColumns = Enumerable.Range(firstColumnNumber, lastColumnNumber - firstColumnNumber + 1)
                .Where(columnNumber => includeHiddenRowsAndColumns || !worksheet.Column(columnNumber).IsHidden)
                .Select(columnNumber => new WorksheetSessionColumn
                {
                    ExcelColumnNumber = columnNumber,
                    OriginalHeaderText = GetCellDisplayValue(worksheet.Cell(headerRow.RowNumber(), columnNumber))
                })
                .ToList();

            var effectiveColumnNumbers = effectiveColumns.Select(column => column.ExcelColumnNumber).ToHashSet();
            var unavailableRequiredColumns = requiredColumns.Where(column => !effectiveColumnNumbers.Contains(column.ExcelColumnNumber)).ToList();
            if (unavailableRequiredColumns.Count > 0)
            {
                throw new WorksheetValidationException($"Required column(s) excluded by the hidden-content policy: {string.Join(", ", unavailableRequiredColumns.Select(column => RequiredHeaders[column.LogicalColumn]))}.");
            }

            var rows = new List<WorksheetSessionRow>();
            var partColumn = requiredColumns.Single(column => column.LogicalColumn == RequiredWorksheetColumn.PartNumber);
            var categoryColumn = requiredColumns.Single(column => column.LogicalColumn == RequiredWorksheetColumn.Category);
            var manufacturerColumn = requiredColumns.Single(column => column.LogicalColumn == RequiredWorksheetColumn.Manufacturer);
            var lastRowNumber = range.LastRow().RowNumber();

            for (var rowNumber = headerRow.RowNumber() + 1; rowNumber <= lastRowNumber; rowNumber++)
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (!includeHiddenRowsAndColumns && worksheet.Row(rowNumber).IsHidden)
                {
                    continue;
                }

                if (effectiveColumns.All(column => string.IsNullOrEmpty(GetCellDisplayValue(worksheet.Cell(rowNumber, column.ExcelColumnNumber)))))
                {
                    continue;
                }

                var partNumber = GetCellDisplayValue(worksheet.Cell(rowNumber, partColumn.ExcelColumnNumber));
                var category = GetCellDisplayValue(worksheet.Cell(rowNumber, categoryColumn.ExcelColumnNumber));
                var manufacturer = GetCellDisplayValue(worksheet.Cell(rowNumber, manufacturerColumn.ExcelColumnNumber));
                rows.Add(new WorksheetSessionRow
                {
                    ExcelRowNumber = rowNumber,
                    PartNumberRawValue = partNumber,
                    PartNumberComparisonValue = AnalyticalValueIdentity.Normalize(partNumber),
                    CategoryRawValue = category,
                    CategoryComparisonValue = AnalyticalValueIdentity.Normalize(category),
                    ManufacturerRawValue = manufacturer,
                    ManufacturerComparisonValue = AnalyticalValueIdentity.Normalize(manufacturer)
                });
            }

            ValidateRequiredValues(rows);
            return new WorksheetProcessingSession
            {
                WorkbookPath = workbookPath,
                WorksheetName = worksheet.Name,
                HeaderRowNumber = headerRow.RowNumber(),
                IncludeHiddenRowsAndColumns = includeHiddenRowsAndColumns,
                RequiredColumns = requiredColumns,
                EffectiveColumns = effectiveColumns,
                EligibleRows = rows
            };
        }
        catch (WorksheetValidationException)
        {
            throw;
        }
        catch (OperationCanceledException)
        {
            throw;
        }
        catch (IOException ex)
        {
            throw new WorksheetValidationException("The workbook could not be opened. Close it in Excel and try again.", ex);
        }
        catch (UnauthorizedAccessException ex)
        {
            throw new WorksheetValidationException("The workbook could not be opened because access was denied.", ex);
        }
        catch (Exception ex)
        {
            throw new WorksheetValidationException("The selected worksheet could not be validated.", ex);
        }
    }

    private static IReadOnlyList<WorksheetColumnMapping> ResolveRequiredColumns(
        IXLWorksheet worksheet,
        int headerRowNumber,
        int firstColumnNumber,
        int lastColumnNumber)
    {
        var headerCells = Enumerable.Range(firstColumnNumber, lastColumnNumber - firstColumnNumber + 1)
            .Select(columnNumber => new
            {
                ColumnNumber = columnNumber,
                OriginalHeaderText = GetCellDisplayValue(worksheet.Cell(headerRowNumber, columnNumber)),
                NormalizedHeader = NormalizeHeader(GetCellDisplayValue(worksheet.Cell(headerRowNumber, columnNumber)))
            })
            .ToList();

        var mappings = new List<WorksheetColumnMapping>();
        foreach (var requiredHeader in RequiredHeaders)
        {
            var matches = headerCells
                .Where(header => string.Equals(header.NormalizedHeader, requiredHeader.Value, StringComparison.OrdinalIgnoreCase))
                .ToList();
            if (matches.Count == 0)
            {
                throw new WorksheetValidationException($"Missing required column: {requiredHeader.Value}.");
            }

            if (matches.Count > 1)
            {
                var columns = string.Join(", ", matches.Select(match => XLHelper.GetColumnLetterFromNumber(match.ColumnNumber)));
                throw new WorksheetValidationException($"Duplicate required column '{requiredHeader.Value}' found in worksheet columns {columns}.");
            }

            var match = matches[0];
            mappings.Add(new WorksheetColumnMapping
            {
                LogicalColumn = requiredHeader.Key,
                OriginalHeaderText = match.OriginalHeaderText,
                ExcelColumnNumber = match.ColumnNumber
            });
        }

        return mappings;
    }

    private static void ValidateRequiredValues(IReadOnlyList<WorksheetSessionRow> rows)
    {
        using var measurement = PerformanceLogger.Measure("Required-cell blank validation", $"rows={rows.Count}");
        var issues = new List<WorksheetValidationIssue>();
        AddIssue(RequiredWorksheetColumn.PartNumber, rows.Where(row => string.IsNullOrWhiteSpace(row.PartNumberRawValue)).Select(row => row.ExcelRowNumber), issues);
        AddIssue(RequiredWorksheetColumn.Manufacturer, rows.Where(row => string.IsNullOrWhiteSpace(row.ManufacturerRawValue)).Select(row => row.ExcelRowNumber), issues);

        if (issues.Count > 0)
        {
            throw new WorksheetValidationException(issues);
        }
    }

    private static void AddIssue(RequiredWorksheetColumn column, IEnumerable<int> rowNumbers, ICollection<WorksheetValidationIssue> issues)
    {
        var numbers = rowNumbers.ToList();
        if (numbers.Count == 0)
        {
            return;
        }

        issues.Add(new WorksheetValidationIssue
        {
            LogicalColumn = column,
            AffectedRowCount = numbers.Count,
            FirstAffectedExcelRowNumbers = numbers.Take(5).ToList()
        });
    }

    private static string NormalizeHeader(string value) => value.Trim();

    private static string GetCellDisplayValue(IXLCell cell) => cell.GetFormattedString();
}
