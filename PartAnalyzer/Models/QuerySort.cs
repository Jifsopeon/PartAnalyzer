namespace PartAnalyzer.Models;

public sealed class QuerySort
{
    public string ColumnKey { get; init; } = string.Empty;

    public SortDirection Direction { get; init; } = SortDirection.Ascending;
}
