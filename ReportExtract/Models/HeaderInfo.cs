namespace ReportExtract.Models;

public sealed class HeaderInfo
{
    public int ColumnIndex { get; init; }

    public string ColumnLetter { get; init; } = string.Empty;

    public string? Name { get; init; }

    public string InternalColumnName { get; init; } = string.Empty;

    public string DisplayName => string.IsNullOrWhiteSpace(Name)
        ? $"Column {ColumnIndex}"
        : Name;

    public string GridHeader => string.IsNullOrWhiteSpace(Name)
        ? $"Column {ColumnLetter}"
        : Name;
}
