namespace PartAnalyzer.Models;

public sealed class PartAnalysisResult
{
    public long UniquePartCount { get; init; }

    public long DuplicatePartCount { get; init; }

    public long MultipleManufacturerPartCount { get; init; }

    public long BlankPrimaryIdentifierRowCount { get; init; }
}
