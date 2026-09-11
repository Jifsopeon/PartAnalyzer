namespace ReportExtract.Models;

public sealed class FilterPreset
{
    public string Name { get; set; } = string.Empty;
    public FilterSelectionSnapshot Selections { get; set; } = new();
}
