using System.Collections.ObjectModel;

namespace ReportExtract.Models;

public sealed class WorksheetInfo
{
    public string Name { get; init; } = string.Empty;

    public string Visibility { get; init; } = string.Empty;

    public bool IsHidden { get; init; }

    public int UsedRowCount { get; init; }

    public int HeaderRowNumber { get; init; }

    public int TotalDataRowCount { get; init; }

    public int EligibleDataRowCount { get; init; }

    public int HiddenDataRowCount { get; init; }

    public int UsedColumnCount { get; init; }

    public int HiddenRowCount { get; init; }

    public int HiddenColumnCount { get; init; }

    public ObservableCollection<HeaderInfo> Headers { get; init; } = new();

    public string DisplayName => IsHidden ? $"{Name} ({Visibility})" : Name;
}
