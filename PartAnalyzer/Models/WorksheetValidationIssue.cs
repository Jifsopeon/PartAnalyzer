namespace PartAnalyzer.Models;

public sealed class WorksheetValidationIssue
{
    public RequiredWorksheetColumn LogicalColumn { get; init; }

    public long AffectedRowCount { get; init; }

    public IReadOnlyList<int> FirstAffectedExcelRowNumbers { get; init; } = Array.Empty<int>();

    public bool HasAdditionalRows => AffectedRowCount > FirstAffectedExcelRowNumbers.Count;
}
