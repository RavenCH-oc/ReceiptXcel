using XlsxDocxGenerator.Models;
using XlsxDocxGenerator.Services.Excel;
using XlsxDocxGenerator.Services.Errors;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Text.RegularExpressions;

namespace XlsxDocxGenerator.Services.Templates;

using System.IO;

/// <summary>
/// Application service for creating a template definition from an Excel mapping row.
/// </summary>
public sealed class TemplateService(
    IExcelReader excelReader,
    MarkerParser markerParser)
{
    public const int CurrentSchemaVersion = 1;

    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        PropertyNameCaseInsensitive = true,
        WriteIndented = true,
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull
    };

    public async Task<TemplateDefinition> CreateDefinitionAsync(
        string excelPath,
        string templateName,
        string wordTemplatePath,
        string worksheetName,
        int headerRowNumber = 1,
        int? markerRowNumber = null,
        string? outputFileNamePattern = null,
        CancellationToken cancellationToken = default)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(templateName);
        ArgumentException.ThrowIfNullOrWhiteSpace(worksheetName);

        var worksheet = await excelReader.ReadAsync(
            excelPath,
            worksheetName,
            headerRowNumber,
            cancellationToken);

        var markerRow = markerRowNumber.HasValue
            ? FindExplicitMarkerRow(worksheet, markerRowNumber.Value, headerRowNumber)
            : FindAutomaticMarkerRow(worksheet);

        if (markerRow.RowNumber <= headerRowNumber)
        {
            throw new TemplateDefinitionException(
                TemplateDefinitionErrorCode.MarkerRowHeaderRelationshipInvalid,
                $"Marker row {markerRow.RowNumber} 必須位於 header row {headerRowNumber} 下方。");
        }

        var mappings = markerParser.ParseMappingRow(markerRow, worksheet.Headers);

        return new TemplateDefinition
        {
            TemplateName = templateName,
            WordTemplatePath = wordTemplatePath ?? string.Empty,
            PreferredWorksheetName = worksheet.WorksheetName,
            HeaderRowNumber = headerRowNumber,
            CreationMappingMarkerRowNumber = markerRow.RowNumber,
            FieldMappings = mappings,
            OutputFileNamePattern = outputFileNamePattern
        };
    }

    public void SaveTemplate(string path, TemplateDefinition definition)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(path);
        ArgumentNullException.ThrowIfNull(definition);
        ValidateDefinition(definition);

        try
        {
            var fullPath = Path.GetFullPath(path);
            var directory = Path.GetDirectoryName(fullPath);
            if (!string.IsNullOrWhiteSpace(directory))
            {
                Directory.CreateDirectory(directory);
            }

            File.WriteAllText(fullPath, JsonSerializer.Serialize(definition, JsonOptions));
        }
        catch (TemplatePersistenceException)
        {
            throw;
        }
        catch (Exception exception) when (exception is IOException or UnauthorizedAccessException or JsonException)
        {
            throw new TemplatePersistenceException(
                TemplatePersistenceErrorCode.InvalidTemplateFile,
                $"無法儲存模板設定檔：{Path.GetFileName(path)}。請確認路徑與權限。",
                exception);
        }
    }

    public TemplateDefinition LoadTemplate(string path)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(path);
        var fullPath = Path.GetFullPath(path);
        if (!File.Exists(fullPath))
        {
            throw new TemplatePersistenceException(
                TemplatePersistenceErrorCode.TemplateFileNotFound,
                $"找不到模板設定檔：{fullPath}");
        }

        try
        {
            var json = File.ReadAllText(fullPath);
            using var document = JsonDocument.Parse(json);
            if (!document.RootElement.TryGetProperty("schemaVersion", out var schemaVersionElement)
                || !schemaVersionElement.TryGetInt32(out var schemaVersion))
            {
                throw new TemplatePersistenceException(
                    TemplatePersistenceErrorCode.MissingRequiredField,
                    "模板設定檔缺少 schemaVersion。");
            }

            if (schemaVersion > CurrentSchemaVersion)
            {
                throw new TemplatePersistenceException(
                    TemplatePersistenceErrorCode.UnsupportedTemplateVersion,
                    $"模板設定檔版本 {schemaVersion} 高於目前支援版本 {CurrentSchemaVersion}。");
            }

            var definition = JsonSerializer.Deserialize<TemplateDefinition>(json, JsonOptions);
            if (definition is null)
            {
                throw new TemplatePersistenceException(
                    TemplatePersistenceErrorCode.InvalidTemplateFile,
                    "模板設定檔內容為空。請選擇有效的 .docxcel.json 檔案。");
            }

            ValidateDefinition(definition);
            return definition;
        }
        catch (TemplatePersistenceException)
        {
            throw;
        }
        catch (JsonException exception)
        {
            throw new TemplatePersistenceException(
                TemplatePersistenceErrorCode.InvalidTemplateFile,
                "模板設定檔不是有效的 JSON，請重新建立模板。",
                exception);
        }
        catch (Exception exception) when (exception is IOException or UnauthorizedAccessException)
        {
            throw new TemplatePersistenceException(
                TemplatePersistenceErrorCode.InvalidTemplateFile,
                "無法讀取模板設定檔，請確認檔案與權限。",
                exception);
        }
    }

    private static void ValidateDefinition(TemplateDefinition definition)
    {
        if (definition.SchemaVersion > CurrentSchemaVersion)
        {
            throw new TemplatePersistenceException(
                TemplatePersistenceErrorCode.UnsupportedTemplateVersion,
                $"模板設定檔版本 {definition.SchemaVersion} 高於目前支援版本 {CurrentSchemaVersion}。");
        }

        if (definition.SchemaVersion <= 0
            || string.IsNullOrWhiteSpace(definition.TemplateName)
            || string.IsNullOrWhiteSpace(definition.WordTemplatePath)
            || string.IsNullOrWhiteSpace(definition.PreferredWorksheetName)
            || definition.HeaderRowNumber <= 0)
        {
            throw new TemplatePersistenceException(
                TemplatePersistenceErrorCode.MissingRequiredField,
                "模板設定檔缺少必要欄位（模板名稱、工作表、標題列或版本）。");
        }

        var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var columns = new HashSet<int>();
        foreach (var mapping in definition.FieldMappings)
        {
            if (string.IsNullOrWhiteSpace(mapping.PlaceholderName)
                || mapping.ColumnIndex <= 0
                || !Regex.IsMatch(mapping.PlaceholderName.Trim(), "^[A-Za-z_][A-Za-z0-9_]*$", RegexOptions.CultureInvariant)
                || !names.Add(mapping.PlaceholderName.Trim())
                || !columns.Add(mapping.ColumnIndex))
            {
                throw new TemplatePersistenceException(
                    TemplatePersistenceErrorCode.InvalidFieldMapping,
                    "模板設定檔包含重複、空白或無效的欄位對應。");
            }
        }
    }

    private ExcelRowData FindExplicitMarkerRow(
        ExcelWorksheetData worksheet,
        int markerRowNumber,
        int headerRowNumber)
    {
        if (markerRowNumber <= 0)
        {
            throw new TemplateDefinitionException(
                TemplateDefinitionErrorCode.InvalidMarkerRow,
                "Marker row 必須是大於 0 的 Excel 列號。");
        }

        if (markerRowNumber <= headerRowNumber)
        {
            throw new TemplateDefinitionException(
                TemplateDefinitionErrorCode.MarkerRowHeaderRelationshipInvalid,
                $"Marker row {markerRowNumber} 必須位於 header row {headerRowNumber} 下方。");
        }

        var markerRow = worksheet.Rows.FirstOrDefault(row => row.RowNumber == markerRowNumber);
        if (markerRow is null)
        {
            throw new TemplateDefinitionException(
                TemplateDefinitionErrorCode.InvalidMarkerRow,
                $"找不到指定的 marker row：Excel 第 {markerRowNumber} 列。");
        }

        return markerRow;
    }

    private ExcelRowData FindAutomaticMarkerRow(ExcelWorksheetData worksheet)
    {
        var candidates = worksheet.Rows
            .Where(markerParser.ContainsValidMarker)
            .OrderByDescending(row => row.RowNumber)
            .ToArray();

        if (candidates.Length > 1)
        {
            var rowNumbers = string.Join(", ", candidates.Select(row => row.RowNumber));
            throw new TemplateDefinitionException(
                TemplateDefinitionErrorCode.AmbiguousMarkerRows,
                $"找到多個可能的 marker row（{rowNumbers}），請明確指定 marker row。");
        }

        if (candidates.Length == 1)
        {
            return candidates[0];
        }

        var invalidMarkerRow = worksheet.Rows.FirstOrDefault(row =>
            row.Cells.Values.Any(markerParser.LooksLikePlaceholder));
        if (invalidMarkerRow is not null)
        {
            throw new TemplateDefinitionException(
                TemplateDefinitionErrorCode.InvalidMarkerSyntax,
                $"Excel 第 {invalidMarkerRow.RowNumber} 列包含格式無效的 placeholder marker。");
        }

        throw new TemplateDefinitionException(
            TemplateDefinitionErrorCode.NoMarkerRow,
            "找不到包含合法 placeholder 的 marker row。");
    }
}
