namespace PartAnalyzer.Models;

public sealed class WorksheetColumnMapping
{
    public RequiredWorksheetColumn LogicalColumn { get; init; }

    public string OriginalHeaderText { get; init; } = string.Empty;

    public int ExcelColumnNumber { get; init; }
}
