namespace ReportExtract.Models;

public sealed class UnavailableFilterSelection
{
    public SessionFilterColumn Column { get; init; }
    public IReadOnlyList<string> Values { get; init; } = Array.Empty<string>();
}
