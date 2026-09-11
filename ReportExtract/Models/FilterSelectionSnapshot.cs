namespace ReportExtract.Models;

public sealed class FilterSelectionSnapshot
{
    public List<string> PartNumbers { get; set; } = new();
    public List<string> Categories { get; set; } = new();
    public List<string> Manufacturers { get; set; } = new();

    public IReadOnlyList<string> ValuesFor(SessionFilterColumn column) => column switch
    {
        SessionFilterColumn.PartNumber => PartNumbers,
        SessionFilterColumn.Category => Categories,
        SessionFilterColumn.Manufacturer => Manufacturers,
        _ => Array.Empty<string>()
    };
}
