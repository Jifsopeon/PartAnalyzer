namespace PartAnalyzer.Models;

public sealed class FilterCriteria
{
    public FilterDefinition Definition { get; init; } = new();

    public IReadOnlyList<string?> SelectedValues { get; init; } = Array.Empty<string?>();

    public TextFilterMode TextMode { get; init; } = TextFilterMode.Contains;

    public string? Text { get; init; }

    public double? Minimum { get; init; }

    public double? Maximum { get; init; }

    public bool IsActive { get; init; }
}
