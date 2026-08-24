namespace PartAnalyzer.Models;

public sealed class FilterValueOption
{
    public string DisplayValue { get; init; } = string.Empty;

    public string? Value { get; init; }

    public long Count { get; init; }

    public bool IsBlank => Value is null;

    public string DisplayWithCount => $"{DisplayValue} ({Count})";
}
