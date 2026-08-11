using System.IO;
using System.Text.Json;
using System.Text.Json.Serialization;
using XlsxDocxGenerator.Models;

namespace XlsxDocxGenerator.Services.Settings;

public sealed class SettingsService
{
    public const int CurrentSettingsVersion = 1;

    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        PropertyNameCaseInsensitive = true,
        WriteIndented = true,
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull
    };

    public SettingsService(string? settingsPath = null)
    {
        SettingsPath = settingsPath ?? Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "ReceiptXcel",
            "settings.json");
    }

    public string SettingsPath { get; }

    public ApplicationSettings Load()
    {
        try
        {
            if (!File.Exists(SettingsPath))
            {
                return new ApplicationSettings();
            }

            var settings = JsonSerializer.Deserialize<ApplicationSettings>(
                File.ReadAllText(SettingsPath),
                JsonOptions);
            if (settings is null || settings.SettingsVersion > CurrentSettingsVersion)
            {
                return new ApplicationSettings();
            }

            return settings with { SettingsVersion = CurrentSettingsVersion };
        }
        catch
        {
            return new ApplicationSettings();
        }
    }

    public void Save(ApplicationSettings settings)
    {
        ArgumentNullException.ThrowIfNull(settings);
        var temporaryPath = SettingsPath + $".{Guid.NewGuid():N}.tmp";

        try
        {
            var directory = Path.GetDirectoryName(SettingsPath);
            if (!string.IsNullOrWhiteSpace(directory))
            {
                Directory.CreateDirectory(directory);
            }

            var normalized = settings with { SettingsVersion = CurrentSettingsVersion };
            File.WriteAllText(temporaryPath, JsonSerializer.Serialize(normalized, JsonOptions));
            File.Move(temporaryPath, SettingsPath, overwrite: true);
        }
        catch
        {
            TryDelete(temporaryPath);
        }
    }

    private static void TryDelete(string path)
    {
        try
        {
            if (File.Exists(path))
            {
                File.Delete(path);
            }
        }
        catch
        {
            // Settings persistence is best effort and must never crash the app.
        }
    }
}
