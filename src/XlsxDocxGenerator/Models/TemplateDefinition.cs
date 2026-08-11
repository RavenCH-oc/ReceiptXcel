namespace XlsxDocxGenerator.Models;

using System.Text.Json.Serialization;

/// <summary>
/// Template configuration only. It does not contain the imported Excel records.
/// </summary>
public sealed record TemplateDefinition
{
    public int SchemaVersion { get; init; } = 1;

    public required string TemplateName { get; init; }

    public required string WordTemplatePath { get; init; }

    public string PreferredWorksheetName { get; init; } = string.Empty;

    public required int HeaderRowNumber { get; init; }

    public IReadOnlyList<FieldMapping> FieldMappings { get; init; } = [];

    public string? OutputFileNamePattern { get; init; }

    /// <summary>Only the row used while creating the template; never used as a runtime exclusion.</summary>
    public int? CreationMappingMarkerRowNumber { get; init; }

    // Compatibility aliases for the Phase 3 API. They are deliberately not persisted.
    [JsonIgnore]
    public string WorksheetName
    {
        get => PreferredWorksheetName;
        init => PreferredWorksheetName = value;
    }

    [JsonIgnore]
    public int MarkerRowNumber
    {
        get => CreationMappingMarkerRowNumber ?? 0;
        init => CreationMappingMarkerRowNumber = value;
    }
}
