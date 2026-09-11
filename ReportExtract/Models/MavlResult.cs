namespace ReportExtract.Models;

public sealed class MavlResult
{
    public IReadOnlyDictionary<int, MavlClassification> ByExcelRowNumber { get; init; } = new Dictionary<int, MavlClassification>();

    public MavlClassification GetForExcelRow(int excelRowNumber) => ByExcelRowNumber[excelRowNumber];
}
