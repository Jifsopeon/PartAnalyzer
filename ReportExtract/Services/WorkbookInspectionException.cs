namespace ReportExtract.Services;

public sealed class WorkbookInspectionException : Exception
{
    public WorkbookInspectionException(string message)
        : base(message)
    {
    }

    public WorkbookInspectionException(string message, Exception innerException)
        : base(message, innerException)
    {
    }
}
