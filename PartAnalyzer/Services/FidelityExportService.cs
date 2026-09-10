using System.IO;
using ClosedXML.Excel;
using PartAnalyzer.Models;

namespace PartAnalyzer.Services;

public sealed class FidelityExportService
{
    public Task<FidelityExportResult> ExportAsync(WorksheetProcessingSession session, MavlResult mavl, IReadOnlyCollection<int> matchingRows, string destinationPath, CancellationToken cancellationToken = default)
    {
        return Task.Run(() => Export(session, mavl, matchingRows, destinationPath, cancellationToken), cancellationToken);
    }

    private static FidelityExportResult Export(WorksheetProcessingSession session, MavlResult mavl, IReadOnlyCollection<int> matchingRows, string destinationPath, CancellationToken cancellationToken)
    {
        using var measurement = PerformanceLogger.Measure("High-fidelity worksheet export", session.WorksheetName);
        if (matchingRows.Count == 0) throw new DuckDbDataException("No matching rows.");
        if (string.Equals(Path.GetFullPath(session.WorkbookPath), Path.GetFullPath(destinationPath), StringComparison.OrdinalIgnoreCase)) throw new DuckDbDataException("The source workbook cannot be overwritten.");
        string? temporaryPath = null;
        try
        {
            using var source = new XLWorkbook(session.WorkbookPath);
            var sourceSheet = source.Worksheet(session.WorksheetName);
            using var output = new XLWorkbook();
            var sheet = sourceSheet.CopyTo(output, session.WorksheetName);
            var partHeader = session.GetRequiredColumn(RequiredWorksheetColumn.PartNumber).OriginalHeaderText.Trim();
            var partColumn = session.GetRequiredColumn(RequiredWorksheetColumn.PartNumber).ExcelColumnNumber;
            var mavlColumns = sheet.Row(session.HeaderRowNumber).CellsUsed().Where(cell => string.Equals(cell.GetFormattedString().Trim(), "MAVL", StringComparison.OrdinalIgnoreCase)).Select(cell => cell.Address.ColumnNumber).OrderByDescending(value => value).ToList();
            var adjacent = partColumn + 1;
            foreach (var column in mavlColumns.Where(column => column != adjacent)) sheet.Column(column).Delete();
            partColumn = sheet.Row(session.HeaderRowNumber).CellsUsed().Single(cell => string.Equals(cell.GetFormattedString().Trim(), partHeader, StringComparison.OrdinalIgnoreCase)).Address.ColumnNumber;
            adjacent = partColumn + 1;
            if (!mavlColumns.Contains(adjacent)) sheet.Column(partColumn).InsertColumnsAfter(1);
            sheet.Cell(session.HeaderRowNumber, adjacent).SetValue("MAVL");
            sheet.Column(adjacent).Style = sheet.Column(partColumn).Style;
            sheet.Column(adjacent).Width = sheet.Column(partColumn).Width;
            foreach (var pair in mavl.ByExcelRowNumber) sheet.Cell(pair.Key, adjacent).SetValue(pair.Value == MavlClassification.Yes ? "Yes" : "No");
            var eligible = session.EligibleRows.Select(row => row.ExcelRowNumber).ToHashSet();
            var keep = matchingRows.ToHashSet();
            var lastRow = sourceSheet.RangeUsed()?.LastRow().RowNumber() ?? session.HeaderRowNumber;
            for (var row = lastRow; row > session.HeaderRowNumber; row--) if (!eligible.Contains(row) || !keep.Contains(row)) sheet.Row(row).Delete();
            if (!session.IncludeHiddenRowsAndColumns)
                for (var column = (sourceSheet.RangeUsed()?.LastColumn().ColumnNumber() ?? 0); column >= 1; column--) if (sourceSheet.Column(column).IsHidden) sheet.Column(column).Delete();
            temporaryPath = Path.Combine(Path.GetDirectoryName(destinationPath)!, $"{Path.GetFileNameWithoutExtension(destinationPath)}.{Guid.NewGuid():N}.tmp.xlsx");
            output.SaveAs(temporaryPath);
            File.Move(temporaryPath, destinationPath, true);
            return new FidelityExportResult { DestinationPath = destinationPath, ExportedDataRowCount = matchingRows.Count, MavlColumnNumber = adjacent };
        }
        catch (Exception ex) when (ex is not DuckDbDataException) { throw new DuckDbDataException("The filtered workbook could not be exported.", ex); }
        finally { if (temporaryPath is not null && File.Exists(temporaryPath)) File.Delete(temporaryPath); }
    }
}
