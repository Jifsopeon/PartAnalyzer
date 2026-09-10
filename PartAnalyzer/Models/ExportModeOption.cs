namespace PartAnalyzer.Models;

public sealed class ExportModeOption
{
    public required ExportMode Mode { get; init; }

    public required string DisplayName { get; init; }

    public string Description { get; init; } = string.Empty;
}
