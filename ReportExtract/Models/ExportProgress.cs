namespace ReportExtract.Models;

public sealed record ExportProgress(
    int Percent,
    string Stage,
    int ProcessedRows,
    int TotalRows,
    bool IsIndeterminate = false);
