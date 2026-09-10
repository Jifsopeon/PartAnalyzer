namespace PartAnalyzer.Models;

public sealed class ExportRequest
{
    public required ExportMode Mode { get; init; }

    public required string DestinationPath { get; init; }

    public IReadOnlyList<FilterCriteria> Filters { get; init; } = Array.Empty<FilterCriteria>();
}
