namespace XlsxDocxGenerator.Models;

public sealed record ApplicationSettings
{
    public int SettingsVersion { get; init; } = 1;

    public string? LastTemplateConfigPath { get; init; }

    public string? LastOutputDirectory { get; init; }

    public string? LastSelectionMode { get; init; }

    public string? LastOutputFilenamePattern { get; init; }

    public double? WindowWidth { get; init; }

    public double? WindowHeight { get; init; }
}
