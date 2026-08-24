namespace PartAnalyzer.Models;

public sealed class FilterDefinition
{
    public string Id { get; init; } = string.Empty;

    public string DisplayName { get; init; } = string.Empty;

    public string InternalColumnName { get; init; } = string.Empty;

    public int DisplayOrder { get; init; }

    public FilterKind Kind { get; init; }

    public FilterTarget Target { get; init; }

    public long DistinctNonBlankCount { get; init; }

    public long BlankCount { get; init; }

    public bool HasBlankValues => BlankCount > 0;
}
