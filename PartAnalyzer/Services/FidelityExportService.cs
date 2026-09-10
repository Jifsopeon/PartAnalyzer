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
            if (!session.IncludeHiddenRowsAndColumns)
                for (var column = sheet.RangeUsed()?.LastColumn().ColumnNumber() ?? 0; column >= 1; column--)
                    if (sheet.Column(column).IsHidden) sheet.Column(column).Delete();
            var partHeader = session.GetRequiredColumn(RequiredWorksheetColumn.PartNumber).OriginalHeaderText.Trim();
            foreach (var column in FindHeaderColumns(sheet, session.HeaderRowNumber, "MAVL").OrderByDescending(column => column)) sheet.Column(column).Delete();
            var partColumn = FindHeaderColumns(sheet, session.HeaderRowNumber, partHeader).Single();
            sheet.Column(partColumn).InsertColumnsAfter(1);
            var adjacent = partColumn + 1;
            sheet.Cell(session.HeaderRowNumber, adjacent).SetValue("MAVL");
            sheet.Column(adjacent).Style = sheet.Column(partColumn).Style;
            sheet.Column(adjacent).Width = sheet.Column(partColumn).Width;
            foreach (var pair in mavl.ByExcelRowNumber) sheet.Cell(pair.Key, adjacent).SetValue(pair.Value == MavlClassification.Yes ? "Yes" : "No");
            var eligible = session.EligibleRows.Select(row => row.ExcelRowNumber).ToHashSet();
            var keep = matchingRows.ToHashSet();
            var lastRow = sourceSheet.RangeUsed()?.LastRow().RowNumber() ?? session.HeaderRowNumber;
            for (var row = lastRow; row > session.HeaderRowNumber; row--)
                if ((eligible.Contains(row) && !keep.Contains(row)) || (!session.IncludeHiddenRowsAndColumns && sourceSheet.Row(row).IsHidden)) sheet.Row(row).Delete();
            var finalMavlColumns = FindHeaderColumns(sheet, session.HeaderRowNumber, "MAVL");
            var finalPartColumn = FindHeaderColumns(sheet, session.HeaderRowNumber, partHeader).Single();
            if (finalMavlColumns.Count != 1 || finalMavlColumns[0] != finalPartColumn + 1) throw new DuckDbDataException("MAVL column placement could not be verified.");
            temporaryPath = Path.Combine(Path.GetDirectoryName(destinationPath)!, $"{Path.GetFileNameWithoutExtension(destinationPath)}.{Guid.NewGuid():N}.tmp.xlsx");
            output.SaveAs(temporaryPath);
            File.Move(temporaryPath, destinationPath, true);
            return new FidelityExportResult { DestinationPath = destinationPath, ExportedDataRowCount = matchingRows.Count, MavlColumnNumber = adjacent };
        }
        catch (Exception ex) when (ex is not DuckDbDataException) { throw new DuckDbDataException("The filtered workbook could not be exported.", ex); }
        finally { if (temporaryPath is not null && File.Exists(temporaryPath)) File.Delete(temporaryPath); }
    }

    private static List<int> FindHeaderColumns(IXLWorksheet sheet, int headerRowNumber, string header)
    {
        return sheet.Row(headerRowNumber).CellsUsed()
            .Where(cell => string.Equals(cell.GetFormattedString().Trim(), header, StringComparison.OrdinalIgnoreCase))
            .Select(cell => cell.Address.ColumnNumber).ToList();
    }
}
