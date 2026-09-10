namespace PartAnalyzer.Models;

public sealed class WorksheetProcessingSession
{
    public string WorkbookPath { get; init; } = string.Empty;

    public string WorksheetName { get; init; } = string.Empty;

    public int HeaderRowNumber { get; init; }

    public bool IncludeHiddenRowsAndColumns { get; init; }

    public IReadOnlyList<WorksheetColumnMapping> RequiredColumns { get; init; } = Array.Empty<WorksheetColumnMapping>();

    public IReadOnlyList<WorksheetSessionColumn> EffectiveColumns { get; init; } = Array.Empty<WorksheetSessionColumn>();

    public IReadOnlyList<WorksheetSessionRow> EligibleRows { get; init; } = Array.Empty<WorksheetSessionRow>();

    public long EligibleRowCount => EligibleRows.Count;

    public WorksheetColumnMapping GetRequiredColumn(RequiredWorksheetColumn logicalColumn)
    {
        return RequiredColumns.Single(column => column.LogicalColumn == logicalColumn);
    }
}
