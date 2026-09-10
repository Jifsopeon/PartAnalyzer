namespace PartAnalyzer.Models;

public sealed class WorksheetSessionColumn
{
    public int ExcelColumnNumber { get; init; }

    public string OriginalHeaderText { get; init; } = string.Empty;
}
