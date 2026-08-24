using System.IO;
using ClosedXML.Excel;
using PartAnalyzer.Models;

namespace PartAnalyzer.Services;

public sealed class ExcelWorkbookService
{
    private static readonly HashSet<string> SupportedExtensions = new(StringComparer.OrdinalIgnoreCase)
    {
        ".xlsx",
        ".xlsm"
    };

    public Task<WorkbookInfo> InspectWorkbookAsync(
        string filePath,
        bool ignoreHiddenRows,
        CancellationToken cancellationToken = default)
    {
        return Task.Run(() => InspectWorkbook(filePath, ignoreHiddenRows, cancellationToken), cancellationToken);
    }

    private static WorkbookInfo InspectWorkbook(
        string filePath,
        bool ignoreHiddenRows,
        CancellationToken cancellationToken)
    {
        if (string.IsNullOrWhiteSpace(filePath))
        {
            throw new WorkbookInspectionException("Choose an Excel workbook to inspect.");
        }

        if (!File.Exists(filePath))
        {
            throw new WorkbookInspectionException("The selected workbook no longer exists.");
        }

        var extension = Path.GetExtension(filePath);
        if (!SupportedExtensions.Contains(extension))
        {
            throw new WorkbookInspectionException("Choose an .xlsx or .xlsm workbook.");
        }

        try
        {
            using var stream = File.Open(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            using var workbook = new XLWorkbook(stream);
            var result = new WorkbookInfo
            {
                FileName = Path.GetFileName(filePath),
                FullPath = filePath
            };

            foreach (var worksheet in workbook.Worksheets)
            {
                cancellationToken.ThrowIfCancellationRequested();
                result.Worksheets.Add(InspectWorksheet(worksheet, ignoreHiddenRows));
            }

            if (result.Worksheets.Count == 0)
            {
                throw new WorkbookInspectionException("The workbook does not contain any worksheets.");
            }

            if (!result.Worksheets.Any(sheet => !sheet.IsHidden && sheet.UsedRowCount > 0 && sheet.UsedColumnCount > 0))
            {
                throw new WorkbookInspectionException("The workbook does not contain a visible worksheet with usable data.");
            }

            return result;
        }
        catch (WorkbookInspectionException)
        {
            throw;
        }
        catch (OperationCanceledException)
        {
            throw;
        }
        catch (IOException ex)
        {
            throw new WorkbookInspectionException("The workbook could not be opened. Close it in Excel and try again.", ex);
        }
        catch (UnauthorizedAccessException ex)
        {
            throw new WorkbookInspectionException("The workbook could not be opened because access was denied.", ex);
        }
        catch (Exception ex)
        {
            throw new WorkbookInspectionException("The selected file is not a valid workbook or could not be inspected.", ex);
        }
    }

    private static WorksheetInfo InspectWorksheet(IXLWorksheet worksheet, bool ignoreHiddenRows)
    {
        var range = worksheet.RangeUsed();
        var visibility = worksheet.Visibility.ToString();
        var isHidden = worksheet.Visibility != XLWorksheetVisibility.Visible;

        if (range is null)
        {
            return new WorksheetInfo
            {
                Name = worksheet.Name,
                Visibility = visibility,
                IsHidden = isHidden
            };
        }

        var firstRowNumber = range.FirstRow().RowNumber();
        var lastRowNumber = range.LastRow().RowNumber();
        var firstColumnNumber = range.FirstColumn().ColumnNumber();
        var lastColumnNumber = range.LastColumn().ColumnNumber();

        var headerRow = range.FirstRowUsed();
        var headerRowNumber = headerRow?.RowNumber() ?? firstRowNumber;
        var hiddenRowCount = 0;
        var hiddenDataRowCount = 0;
        var totalDataRowCount = 0;
        var eligibleDataRowCount = 0;
        for (var rowNumber = firstRowNumber; rowNumber <= lastRowNumber; rowNumber++)
        {
            var isRowHidden = worksheet.Row(rowNumber).IsHidden;
            if (isRowHidden)
            {
                hiddenRowCount++;
            }

            if (rowNumber <= headerRowNumber)
            {
                continue;
            }

            totalDataRowCount++;
            if (isRowHidden)
            {
                hiddenDataRowCount++;
            }

            if (!ignoreHiddenRows || !isRowHidden)
            {
                eligibleDataRowCount++;
            }
        }

        var hiddenColumnCount = 0;
        for (var columnNumber = firstColumnNumber; columnNumber <= lastColumnNumber; columnNumber++)
        {
            if (worksheet.Column(columnNumber).IsHidden)
            {
                hiddenColumnCount++;
            }
        }

        var headers = new List<HeaderInfo>();
        if (headerRow is not null)
        {
            for (var columnNumber = firstColumnNumber; columnNumber <= lastColumnNumber; columnNumber++)
            {
                var cell = worksheet.Cell(headerRow.RowNumber(), columnNumber);
                var displayOrder = columnNumber - firstColumnNumber + 1;
                headers.Add(new HeaderInfo
                {
                    ColumnIndex = columnNumber,
                    ColumnLetter = XLHelper.GetColumnLetterFromNumber(columnNumber),
                    InternalColumnName = $"col_{displayOrder:000}",
                    Name = cell.GetFormattedString().Trim()
                });
            }
        }

        return new WorksheetInfo
        {
            Name = worksheet.Name,
            Visibility = visibility,
            IsHidden = isHidden,
            UsedRowCount = lastRowNumber - firstRowNumber + 1,
            HeaderRowNumber = headerRowNumber,
            TotalDataRowCount = totalDataRowCount,
            EligibleDataRowCount = eligibleDataRowCount,
            HiddenDataRowCount = hiddenDataRowCount,
            UsedColumnCount = lastColumnNumber - firstColumnNumber + 1,
            HiddenRowCount = hiddenRowCount,
            HiddenColumnCount = hiddenColumnCount,
            Headers = new(headers)
        };
    }
}
