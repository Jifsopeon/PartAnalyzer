namespace ReportExtract.Models;

public sealed class FidelityExportResult
{
    public string DestinationPath { get; init; } = string.Empty;
    public int ExportedDataRowCount { get; init; }
    public int MavlColumnNumber { get; init; }
}
