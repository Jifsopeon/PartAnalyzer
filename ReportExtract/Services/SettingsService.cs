using System.IO;
using System.Text.Json;
using ReportExtract.Models;

namespace ReportExtract.Services;

public sealed class SettingsService
{
    private const string PreviousApplicationDirectoryName = "ReportExtract";
    private const string LegacyApplicationDirectoryName = "PartAnalyzer";
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        WriteIndented = true
    };

    public string SettingsDirectory { get; } = PortableApplicationPaths.Current.DataDirectory;

    public string SettingsPath => Path.Combine(SettingsDirectory, "settings.json");

    private string PreviousSettingsPath => Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        PreviousApplicationDirectoryName,
        "settings.json");

    private string LegacySettingsPath => Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        LegacyApplicationDirectoryName,
        "settings.json");

    public AppSettings Load()
    {
        try
        {
            if (File.Exists(SettingsPath))
            {
                return LoadSettingsFile(SettingsPath);
            }

            var migrationSourcePath = File.Exists(PreviousSettingsPath)
                ? PreviousSettingsPath
                : File.Exists(LegacySettingsPath)
                    ? LegacySettingsPath
                    : null;
            if (migrationSourcePath is null)
            {
                return new AppSettings();
            }

            try
            {
                Directory.CreateDirectory(SettingsDirectory);
                File.Copy(migrationSourcePath, SettingsPath, overwrite: false);
                return LoadSettingsFile(SettingsPath);
            }
            catch (IOException ex)
            {
                PerformanceLogger.Write($"SETTINGS_MIGRATION_FAILED source=\"{migrationSourcePath}\" message=\"{ex.Message}\"");
                return LoadSettingsFile(migrationSourcePath);
            }
            catch (UnauthorizedAccessException ex)
            {
                PerformanceLogger.Write($"SETTINGS_MIGRATION_FAILED source=\"{migrationSourcePath}\" message=\"{ex.Message}\"");
                return LoadSettingsFile(migrationSourcePath);
            }
        }
        catch (JsonException)
        {
            return new AppSettings();
        }
        catch (IOException)
        {
            return new AppSettings();
        }
        catch (UnauthorizedAccessException)
        {
            return new AppSettings();
        }
    }

    public void Save(AppSettings settings)
    {
        Directory.CreateDirectory(SettingsDirectory);
        var json = JsonSerializer.Serialize(settings, JsonOptions);
        File.WriteAllText(SettingsPath, json);
    }

    private static AppSettings LoadSettingsFile(string path)
    {
        var json = File.ReadAllText(path);
        return JsonSerializer.Deserialize<AppSettings>(json, JsonOptions) ?? new AppSettings();
    }
}
