using System.Data;

namespace PartAnalyzer.Models;

public sealed class DataPageResult
{
    public DataTable Rows { get; init; } = new();

    public int PageNumber { get; init; }

    public int PageSize { get; init; }

    public long TotalRows { get; init; }

    public int TotalPages { get; init; }

    public long FirstDisplayRow => TotalRows == 0 ? 0 : ((long)(PageNumber - 1) * PageSize) + 1;

    public long LastDisplayRow => Math.Min((long)PageNumber * PageSize, TotalRows);
}
