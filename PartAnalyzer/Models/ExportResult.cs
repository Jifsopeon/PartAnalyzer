namespace PartAnalyzer.Models;

public sealed class ExportResult
{
    public string DestinationPath { get; init; } = string.Empty;

    public long SourceRowCount { get; init; }

    public long GroupedRowCount { get; init; }

    public long TotalRowCount => SourceRowCount + GroupedRowCount;
}
