namespace PartAnalyzer.Models;

public sealed class PartAnalysisMapping
{
    public ImportedColumnInfo PrimaryPartIdentifier { get; init; } = new();

    public ImportedColumnInfo Manufacturer { get; init; } = new();

    public ImportedColumnInfo? ManufacturerPartNumber { get; init; }
}
