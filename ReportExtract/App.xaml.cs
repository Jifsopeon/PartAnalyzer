using System.IO;
using System.Windows;
using System.Windows.Threading;
using ReportExtract.Services;

namespace ReportExtract;

public partial class App : Application
{
    protected override void OnStartup(StartupEventArgs e)
    {
        try
        {
            PortableApplicationPaths.Current.EnsureWritable();
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            MessageBox.Show(
                "ReportExtract cannot write to its application folder. Move the extracted ReportExtract folder to a writable location such as Documents or Desktop.",
                "ReportExtract",
                MessageBoxButton.OK,
                MessageBoxImage.Error);
            Shutdown(-1);
            return;
        }

        base.OnStartup(e);
        DispatcherUnhandledException += AppDispatcherUnhandledException;
        AppDomain.CurrentDomain.UnhandledException += CurrentDomainUnhandledException;
        PerformanceLogger.Write("STARTUP");
    }

    protected override void OnExit(ExitEventArgs e)
    {
        PerformanceLogger.Write($"EXIT code={e.ApplicationExitCode}");
        base.OnExit(e);
    }

    private static void AppDispatcherUnhandledException(object sender, DispatcherUnhandledExceptionEventArgs e)
    {
        PerformanceLogger.Write($"UNHANDLED_DISPATCHER_EXCEPTION type=\"{e.Exception.GetType().FullName}\" message=\"{e.Exception.Message}\"");
    }

    private static void CurrentDomainUnhandledException(object sender, UnhandledExceptionEventArgs e)
    {
        if (e.ExceptionObject is Exception ex)
        {
            PerformanceLogger.Write($"UNHANDLED_DOMAIN_EXCEPTION type=\"{ex.GetType().FullName}\" message=\"{ex.Message}\" terminating={e.IsTerminating}");
        }
    }
}
