using System.Diagnostics;
using System.IO;
using ClosedXML.Excel;
using PartAnalyzer.Models;

namespace PartAnalyzer.Services;

public sealed class FidelityExportService
{
    private const int CancellationCheckInterval = 256;

    public Task<FidelityExportResult> ExportAsync(WorksheetProcessingSession session, MavlResult mavl, IReadOnlyCollection<int> matchingRows, string destinationPath, CancellationToken cancellationToken = default, IProgress<ExportProgress>? progress = null)
    {
        return Task.Run(() => Export(session, mavl, matchingRows, destinationPath, cancellationToken, progress), cancellationToken);
    }

    private static FidelityExportResult Export(WorksheetProcessingSession session, MavlResult mavl, IReadOnlyCollection<int> matchingRows, string destinationPath, CancellationToken cancellationToken, IProgress<ExportProgress>? progress)
    {
        if (matchingRows.Count == 0)
        {
            throw new DuckDbDataException("No matching rows.");
        }

        var fullSourcePath = Path.GetFullPath(session.WorkbookPath);
        var fullDestinationPath = Path.GetFullPath(destinationPath);
        if (string.Equals(fullSourcePath, fullDestinationPath, StringComparison.OrdinalIgnoreCase))
        {
            throw new DuckDbDataException("The source workbook cannot be overwritten.");
        }

        var totalTimer = Stopwatch.StartNew();
        string? temporaryPath = null;
        try
        {
            cancellationToken.ThrowIfCancellationRequested();
            Report(progress, 0, "Preparing export...", 0, matchingRows.Count);
            var openTimer = Stopwatch.StartNew();
            using var source = new XLWorkbook(fullSourcePath);
            var sourceSheet = source.Worksheet(session.WorksheetName);
            LogStage("open source workbook", openTimer, session.WorksheetName);
            Report(progress, 5, "Opening workbook...", 0, matchingRows.Count);

            cancellationToken.ThrowIfCancellationRequested();
            var prepareTimer = Stopwatch.StartNew();
            var sourcePartColumn = session.GetRequiredColumn(RequiredWorksheetColumn.PartNumber).ExcelColumnNumber;
            var sourceMavlColumns = session.EffectiveColumns
                .Where(column => string.Equals(column.OriginalHeaderText.Trim(), "MAVL", StringComparison.OrdinalIgnoreCase))
                .Select(column => column.ExcelColumnNumber)
                .ToHashSet();
            var eligibleRowNumbers = session.EligibleRows.Select(row => row.ExcelRowNumber).ToHashSet();
            var exportedRows = matchingRows
                .Where(eligibleRowNumbers.Contains)
                .Distinct()
                .OrderBy(rowNumber => rowNumber)
                .ToList();
            if (exportedRows.Count == 0)
            {
                throw new DuckDbDataException("No matching rows.");
            }

            var columns = CreateExportColumns(session.EffectiveColumns, sourceMavlColumns, sourcePartColumn);
            var mavlColumnNumber = columns.Single(column => column.IsGeneratedMavl).OutputColumnNumber;
            var sourceUsedRows = sourceSheet.RangeUsed()?.RowCount() ?? 0;
            PerformanceLogger.Write($"Filtered workbook export context worksheet=\"{session.WorksheetName}\" source_used_rows={sourceUsedRows} matching_rows={exportedRows.Count} exported_source_columns={columns.Count(column => !column.IsGeneratedMavl)} generated_columns=1");
            LogStage("determine export columns", prepareTimer, $"source_columns={columns.Count(column => !column.IsGeneratedMavl)}");
            Report(progress, 10, "Preparing columns and layout...", 0, exportedRows.Count);

            cancellationToken.ThrowIfCancellationRequested();
            var createTimer = Stopwatch.StartNew();
            using var output = new XLWorkbook();
            var outputSheet = output.Worksheets.Add(session.WorksheetName);
            LogStage("create output worksheet", createTimer, session.WorksheetName);

            cancellationToken.ThrowIfCancellationRequested();
            var layoutTimer = Stopwatch.StartNew();
            CopyHeaderAndColumnLayout(sourceSheet, outputSheet, session.HeaderRowNumber, columns, sourcePartColumn);
            LogStage("copy header and layout", layoutTimer, $"columns={columns.Count}");
            Report(progress, 20, "Copying matching rows...", 0, exportedRows.Count);

            cancellationToken.ThrowIfCancellationRequested();
            var copyTimer = Stopwatch.StartNew();
            CopyDataRows(sourceSheet, outputSheet, exportedRows, columns, cancellationToken, progress);
            LogStage("copy data rows", copyTimer, $"rows={exportedRows.Count}");

            cancellationToken.ThrowIfCancellationRequested();
            var mavlTimer = Stopwatch.StartNew();
            WriteMavl(outputSheet, exportedRows, mavl, mavlColumnNumber, columns.Single(column => column.SourceColumnNumber == sourcePartColumn).OutputColumnNumber, cancellationToken, progress);
            LogStage("write MAVL", mavlTimer, $"rows={exportedRows.Count}");

            cancellationToken.ThrowIfCancellationRequested();
            var freezeTimer = Stopwatch.StartNew();
            ApplyFreezePanes(sourceSheet, outputSheet, columns, sourcePartColumn);
            LogStage("apply freeze panes", freezeTimer);
            Report(progress, 90, "Applying worksheet layout...", exportedRows.Count, exportedRows.Count);

            cancellationToken.ThrowIfCancellationRequested();
            var destinationDirectory = Path.GetDirectoryName(fullDestinationPath) ?? throw new DuckDbDataException("The export destination is invalid.");
            temporaryPath = Path.Combine(destinationDirectory, $"{Path.GetFileNameWithoutExtension(fullDestinationPath)}.{Guid.NewGuid():N}.tmp.xlsx");
            var saveTimer = Stopwatch.StartNew();
            Report(progress, 90, "Saving filtered workbook...", exportedRows.Count, exportedRows.Count, isIndeterminate: true);
            output.SaveAs(temporaryPath);
            cancellationToken.ThrowIfCancellationRequested();
            File.Move(temporaryPath, fullDestinationPath, true);
            temporaryPath = null;
            LogStage("save workbook", saveTimer, fullDestinationPath);
            totalTimer.Stop();
            PerformanceLogger.Write($"Filtered workbook export total elapsed_ms={totalTimer.ElapsedMilliseconds} rows={exportedRows.Count} columns={columns.Count}");
            Report(progress, 100, "Export completed.", exportedRows.Count, exportedRows.Count);

            return new FidelityExportResult
            {
                DestinationPath = fullDestinationPath,
                ExportedDataRowCount = exportedRows.Count,
                MavlColumnNumber = mavlColumnNumber
            };
        }
        catch (Exception ex) when (ex is not DuckDbDataException and not OperationCanceledException)
        {
            throw new DuckDbDataException("The filtered workbook could not be exported.", ex);
        }
        finally
        {
            if (temporaryPath is not null && File.Exists(temporaryPath))
            {
                File.Delete(temporaryPath);
            }
        }
    }

    private static List<ExportColumn> CreateExportColumns(IReadOnlyList<WorksheetSessionColumn> effectiveColumns, IReadOnlySet<int> sourceMavlColumns, int sourcePartColumn)
    {
        var columns = new List<ExportColumn>();
        var outputColumn = 1;
        foreach (var sourceColumn in effectiveColumns)
        {
            if (sourceMavlColumns.Contains(sourceColumn.ExcelColumnNumber))
            {
                continue;
            }

            columns.Add(new ExportColumn(sourceColumn.ExcelColumnNumber, outputColumn++, false));
            if (sourceColumn.ExcelColumnNumber == sourcePartColumn)
            {
                columns.Add(new ExportColumn(null, outputColumn++, true));
            }
        }

        if (!columns.Any(column => column.IsGeneratedMavl))
        {
            throw new DuckDbDataException("The P+F part number column could not be mapped for MAVL export.");
        }

        return columns;
    }

    private static void CopyHeaderAndColumnLayout(IXLWorksheet sourceSheet, IXLWorksheet outputSheet, int sourceHeaderRow, IReadOnlyList<ExportColumn> columns, int sourcePartColumn)
    {
        outputSheet.Row(1).Height = sourceSheet.Row(sourceHeaderRow).Height;
        foreach (var column in columns)
        {
            var targetColumn = outputSheet.Column(column.OutputColumnNumber);
            if (column.IsGeneratedMavl)
            {
                var sourcePartHeader = sourceSheet.Cell(sourceHeaderRow, sourcePartColumn);
                CopyCellValueAndStyle(sourcePartHeader, outputSheet.Cell(1, column.OutputColumnNumber));
                outputSheet.Cell(1, column.OutputColumnNumber).SetValue("MAVL");
                targetColumn.Width = sourceSheet.Column(sourcePartColumn).Width;
                continue;
            }

            var sourceColumn = column.SourceColumnNumber!.Value;
            CopyCellValueAndStyle(sourceSheet.Cell(sourceHeaderRow, sourceColumn), outputSheet.Cell(1, column.OutputColumnNumber));
            targetColumn.Width = sourceSheet.Column(sourceColumn).Width;
        }
    }

    private static void CopyDataRows(IXLWorksheet sourceSheet, IXLWorksheet outputSheet, IReadOnlyList<int> exportedRows, IReadOnlyList<ExportColumn> columns, CancellationToken cancellationToken, IProgress<ExportProgress>? progress)
    {
        for (var sourceRowIndex = 0; sourceRowIndex < exportedRows.Count; sourceRowIndex++)
        {
            if (sourceRowIndex % CancellationCheckInterval == 0)
            {
                cancellationToken.ThrowIfCancellationRequested();
            }

            var sourceRowNumber = exportedRows[sourceRowIndex];
            var outputRowNumber = sourceRowIndex + 2;
            outputSheet.Row(outputRowNumber).Height = sourceSheet.Row(sourceRowNumber).Height;
            foreach (var column in columns.Where(column => !column.IsGeneratedMavl))
            {
                CopyCellValueAndStyle(
                    sourceSheet.Cell(sourceRowNumber, column.SourceColumnNumber!.Value),
                    outputSheet.Cell(outputRowNumber, column.OutputColumnNumber));
            }

            if (sourceRowIndex % CancellationCheckInterval == 0 || sourceRowIndex == exportedRows.Count - 1)
            {
                Report(progress, 20 + (sourceRowIndex + 1) * 55 / exportedRows.Count, "Copying matching rows...", sourceRowIndex + 1, exportedRows.Count);
            }
        }
    }

    private static void WriteMavl(IXLWorksheet outputSheet, IReadOnlyList<int> exportedRows, MavlResult mavl, int mavlColumnNumber, int partOutputColumnNumber, CancellationToken cancellationToken, IProgress<ExportProgress>? progress)
    {
        for (var sourceRowIndex = 0; sourceRowIndex < exportedRows.Count; sourceRowIndex++)
        {
            if (sourceRowIndex % CancellationCheckInterval == 0)
            {
                cancellationToken.ThrowIfCancellationRequested();
            }

            var outputRowNumber = sourceRowIndex + 2;
            var mavlCell = outputSheet.Cell(outputRowNumber, mavlColumnNumber);
            mavlCell.Style = outputSheet.Cell(outputRowNumber, partOutputColumnNumber).Style;
            mavlCell.SetValue(mavl.GetForExcelRow(exportedRows[sourceRowIndex]) == MavlClassification.Yes ? "Yes" : "No");
            if (sourceRowIndex % CancellationCheckInterval == 0 || sourceRowIndex == exportedRows.Count - 1)
            {
                Report(progress, 75 + (sourceRowIndex + 1) * 10 / exportedRows.Count, "Writing MAVL...", sourceRowIndex + 1, exportedRows.Count);
            }
        }
    }

    private static void CopyCellValueAndStyle(IXLCell source, IXLCell target)
    {
        if (source.HasFormula && source.NeedsRecalculation)
        {
            throw new DuckDbDataException($"Formula cell {source.Address} has no reliable calculated value for export.");
        }

        target.SetValue(source.HasFormula ? source.CachedValue : source.Value);
        target.Style = source.Style;
    }

    private static void ApplyFreezePanes(IXLWorksheet sourceSheet, IXLWorksheet outputSheet, IReadOnlyList<ExportColumn> columns, int sourcePartColumn)
    {
        var frozenSourceRows = sourceSheet.SheetView.SplitRow;
        var frozenSourceColumns = sourceSheet.SheetView.SplitColumn;
        var frozenRows = Math.Min(frozenSourceRows, 1);
        var frozenColumns = columns.Count(column => !column.IsGeneratedMavl && column.SourceColumnNumber <= frozenSourceColumns);
        if (sourcePartColumn <= frozenSourceColumns)
        {
            frozenColumns++;
        }

        if (frozenRows > 0 || frozenColumns > 0)
        {
            outputSheet.SheetView.Freeze(frozenRows, frozenColumns);
        }
    }

    private static void LogStage(string stage, Stopwatch stopwatch, string? detail = null)
    {
        stopwatch.Stop();
        PerformanceLogger.Write($"Filtered workbook export {stage} elapsed_ms={stopwatch.ElapsedMilliseconds}{(string.IsNullOrWhiteSpace(detail) ? string.Empty : $" detail=\"{detail}\"")}");
    }

    private static void Report(IProgress<ExportProgress>? progress, int percent, string stage, int processedRows, int totalRows, bool isIndeterminate = false)
    {
        progress?.Report(new ExportProgress(percent, stage, processedRows, totalRows, isIndeterminate));
    }

    private sealed record ExportColumn(int? SourceColumnNumber, int OutputColumnNumber, bool IsGeneratedMavl);
}
