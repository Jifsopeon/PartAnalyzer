using System.IO;
using System.Text.Json;
using ReportExtract.Models;

namespace ReportExtract.Services;

public sealed class SettingsService
{
    private const string LegacyApplicationDirectoryName = "PartAnalyzer";
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        WriteIndented = true
    };

    public string SettingsDirectory { get; } = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "ReportExtract");

    public string SettingsPath => Path.Combine(SettingsDirectory, "settings.json");

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

            if (!File.Exists(LegacySettingsPath))
            {
                return new AppSettings();
            }

            try
            {
                Directory.CreateDirectory(SettingsDirectory);
                File.Copy(LegacySettingsPath, SettingsPath, overwrite: false);
                return LoadSettingsFile(SettingsPath);
            }
            catch (IOException)
            {
                return LoadSettingsFile(LegacySettingsPath);
            }
            catch (UnauthorizedAccessException)
            {
                return LoadSettingsFile(LegacySettingsPath);
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
