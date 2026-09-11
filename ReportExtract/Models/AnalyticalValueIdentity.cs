namespace ReportExtract.Models;

public static class AnalyticalValueIdentity
{
    public static string Normalize(string? rawValue) => rawValue?.Trim() ?? string.Empty;
}
