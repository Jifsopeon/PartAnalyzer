using System.Collections.ObjectModel;

namespace PartAnalyzer.Models;

public sealed class WorkbookInfo
{
    public string FileName { get; init; } = string.Empty;

    public string FullPath { get; init; } = string.Empty;

    public ObservableCollection<WorksheetInfo> Worksheets { get; init; } = new();
}
