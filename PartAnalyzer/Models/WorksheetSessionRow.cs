namespace PartAnalyzer.Models;

public sealed class WorksheetSessionRow
{
    public int ExcelRowNumber { get; init; }

    public string PartNumberRawValue { get; init; } = string.Empty;

    public string CategoryRawValue { get; init; } = string.Empty;

    public string ManufacturerRawValue { get; init; } = string.Empty;

    // This is intentionally the only normalization applied to Manufacturer identity.
    public string ManufacturerComparisonValue { get; init; } = string.Empty;
}
