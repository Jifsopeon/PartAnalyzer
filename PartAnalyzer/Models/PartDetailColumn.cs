namespace PartAnalyzer.Models;

public sealed class PartDetailColumn
{
    public string Header { get; init; } = string.Empty;

    public string BindingPath { get; init; } = string.Empty;

    public int DisplayOrder { get; init; }
}
