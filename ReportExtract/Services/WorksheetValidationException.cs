using ReportExtract.Models;

namespace ReportExtract.Services;

public sealed class WorksheetValidationException : Exception
{
    public WorksheetValidationException(string message, Exception? innerException = null)
        : base(message, innerException)
    {
    }

    public WorksheetValidationException(IReadOnlyList<WorksheetValidationIssue> issues)
        : base(CreateMessage(issues))
    {
        Issues = issues;
    }

    public IReadOnlyList<WorksheetValidationIssue> Issues { get; } = Array.Empty<WorksheetValidationIssue>();

    private static string CreateMessage(IReadOnlyList<WorksheetValidationIssue> issues)
    {
        var summaries = issues.Select(issue =>
        {
            var rows = string.Join(", ", issue.FirstAffectedExcelRowNumbers);
            var additional = issue.HasAdditionalRows ? " Showing first 5 rows." : string.Empty;
            return $"Blank {ToDisplayName(issue.LogicalColumn)}: rows {rows}. {issue.AffectedRowCount} affected row(s).{additional}";
        });

        return $"Correct the source worksheet and try again. {string.Join(" ", summaries)}";
    }

    private static string ToDisplayName(RequiredWorksheetColumn column)
    {
        return column switch
        {
            RequiredWorksheetColumn.PartNumber => "P+F part number",
            RequiredWorksheetColumn.Category => "Category",
            RequiredWorksheetColumn.Manufacturer => "Manufacturer",
            _ => column.ToString()
        };
    }
}
