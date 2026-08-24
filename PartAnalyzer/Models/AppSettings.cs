namespace PartAnalyzer.Models;

public sealed class AppSettings
{
    public bool IgnoreHiddenRows { get; set; } = true;

    public string? LastImportDirectory { get; set; }

    public string? PreferredWorksheet { get; set; }

    public ColumnMapping? PrimaryPartIdentifierMapping { get; set; }

    public ColumnMapping? ManufacturerMapping { get; set; }

    public ColumnMapping? ManufacturerPartNumberMapping { get; set; }
}
