namespace PartAnalyzer.Models;

public sealed class ColumnMapping
{
    public string InternalColumnName { get; set; } = string.Empty;

    public int ExcelColumnNumber { get; set; }

    public string ExcelColumnLetter { get; set; } = string.Empty;

    public string DisplayHeader { get; set; } = string.Empty;
}
