namespace PartAnalyzer.Models;

public sealed class WorksheetSessionRow
{
    public int ExcelRowNumber { get; init; }

    public string PartNumberRawValue { get; init; } = string.Empty;

    public string PartNumberComparisonValue { get; init; } = string.Empty;

    public string CategoryRawValue { get; init; } = string.Empty;

    public string CategoryComparisonValue { get; init; } = string.Empty;

    public string ManufacturerRawValue { get; init; } = string.Empty;

    // Every core analytical identity uses leading/trailing trim only.
    public string ManufacturerComparisonValue { get; init; } = string.Empty;
}
