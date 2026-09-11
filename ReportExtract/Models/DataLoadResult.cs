namespace ReportExtract.Models;

public sealed class DataLoadResult
{
    public string WorkbookPath { get; init; } = string.Empty;

    public string WorksheetName { get; init; } = string.Empty;

    public int HeaderRowNumber { get; init; }

    public long ImportedRowCount { get; init; }

    public int ImportedColumnCount { get; init; }

    public int SkippedBlankRowCount { get; init; }

    public IReadOnlyList<ImportedColumnInfo> Columns { get; init; } = Array.Empty<ImportedColumnInfo>();

    public WorksheetProcessingSession? ProcessingSession { get; init; }
}
