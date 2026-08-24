namespace PartAnalyzer.Models;

public sealed class ImportedColumnInfo
{
    public int DisplayOrder { get; init; }

    public int ExcelColumnNumber { get; init; }

    public string ExcelColumnLetter { get; init; } = string.Empty;

    public string OriginalHeader { get; init; } = string.Empty;

    public string DisplayHeader { get; init; } = string.Empty;

    public string GridHeader { get; init; } = string.Empty;

    public string InternalColumnName { get; init; } = string.Empty;
}
