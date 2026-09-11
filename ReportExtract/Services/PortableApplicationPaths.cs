using System.IO;

namespace ReportExtract.Services;

public sealed class PortableApplicationPaths
{
    private static readonly Lazy<PortableApplicationPaths> CurrentPaths = new(() => new PortableApplicationPaths());

    private PortableApplicationPaths()
    {
        var executableDirectory = Path.GetFullPath(AppContext.BaseDirectory).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
        var directoryName = Path.GetFileName(executableDirectory);
        ApplicationRoot = string.Equals(directoryName, "Backend", StringComparison.OrdinalIgnoreCase)
            ? new DirectoryInfo(executableDirectory).Parent?.FullName
                ?? throw new InvalidOperationException("The portable application root could not be resolved.")
            : executableDirectory;
        DataDirectory = Path.Combine(ApplicationRoot, "Data");
        TempDirectory = Path.Combine(DataDirectory, "Temp");
        LogsDirectory = Path.Combine(DataDirectory, "Logs");
    }

    public static PortableApplicationPaths Current => CurrentPaths.Value;

    public string ApplicationRoot { get; }

    public string DataDirectory { get; }

    public string TempDirectory { get; }

    public string LogsDirectory { get; }

    public void EnsureWritable()
    {
        Directory.CreateDirectory(DataDirectory);
        Directory.CreateDirectory(TempDirectory);
        Directory.CreateDirectory(LogsDirectory);
        var probePath = Path.Combine(DataDirectory, $".write-probe-{Guid.NewGuid():N}");
        File.WriteAllText(probePath, string.Empty);
        File.Delete(probePath);
    }
}
