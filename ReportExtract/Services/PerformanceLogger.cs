using System.Diagnostics;
using System.IO;

namespace ReportExtract.Services;

public static class PerformanceLogger
{
    private const long MaxLogBytes = 1_000_000;
    private static readonly object SyncRoot = new();
    private static readonly string LogDirectory = PortableApplicationPaths.Current.LogsDirectory;

    private static readonly string LogPath = Path.Combine(LogDirectory, "performance.log");

    public static IDisposable Measure(string operation, string? detail = null)
    {
        return new Measurement(operation, detail);
    }

    public static void Write(string message)
    {
        WriteLine(message);
    }

    public static long ManagedMemoryBytes()
    {
        return GC.GetTotalMemory(forceFullCollection: false);
    }

    private static void WriteLine(string message)
    {
        try
        {
            lock (SyncRoot)
            {
                Directory.CreateDirectory(LogDirectory);
                RollLogIfNeeded();
                File.AppendAllText(LogPath, $"{DateTimeOffset.Now:O} {message}{Environment.NewLine}");
            }
        }
        catch
        {
            // Diagnostics must never prevent the local workflow from continuing.
        }
    }

    private static void RollLogIfNeeded()
    {
        var info = new FileInfo(LogPath);
        if (!info.Exists || info.Length < MaxLogBytes)
        {
            return;
        }

        var oldPath = Path.Combine(LogDirectory, "performance.1.log");
        if (File.Exists(oldPath))
        {
            File.Delete(oldPath);
        }

        File.Move(LogPath, oldPath);
    }

    private sealed class Measurement : IDisposable
    {
        private readonly string _operation;
        private readonly string? _detail;
        private readonly long _startMemory;
        private readonly Stopwatch _stopwatch;

        public Measurement(string operation, string? detail)
        {
            _operation = operation;
            _detail = detail;
            _startMemory = ManagedMemoryBytes();
            _stopwatch = Stopwatch.StartNew();
            WriteLine($"START operation=\"{_operation}\" detail=\"{_detail}\" managed_mb={BytesToMegabytes(_startMemory):F1}");
        }

        public void Dispose()
        {
            _stopwatch.Stop();
            var endMemory = ManagedMemoryBytes();
            WriteLine(
                $"END operation=\"{_operation}\" detail=\"{_detail}\" elapsed_ms={_stopwatch.ElapsedMilliseconds} managed_mb={BytesToMegabytes(endMemory):F1} delta_mb={BytesToMegabytes(endMemory - _startMemory):F1}");
        }

        private static double BytesToMegabytes(long bytes)
        {
            return bytes / 1024d / 1024d;
        }
    }
}
