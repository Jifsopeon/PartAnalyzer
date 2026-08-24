namespace PartAnalyzer.Services;

public sealed class DuckDbDataException : Exception
{
    public DuckDbDataException(string message)
        : base(message)
    {
    }

    public DuckDbDataException(string message, Exception innerException)
        : base(message, innerException)
    {
    }
}
