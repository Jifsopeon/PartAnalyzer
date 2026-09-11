namespace PartAnalyzer.Models;

public sealed class AppSettings
{
    public bool IncludeHiddenRowsAndColumns { get; set; }

    public string? LastImportDirectory { get; set; }

    public string? PreferredWorksheet { get; set; }

    public List<FilterPreset> FilterPresets { get; set; } = new();

    public FilterSelectionSnapshot LastUsedFilterSelections { get; set; } = new();
}
