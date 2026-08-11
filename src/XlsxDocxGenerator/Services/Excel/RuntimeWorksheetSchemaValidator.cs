using XlsxDocxGenerator.Models;
using XlsxDocxGenerator.Services.Errors;

namespace XlsxDocxGenerator.Services.Excel;

public sealed class RuntimeWorksheetSchemaValidator
{
    public RuntimeWorksheetContext Validate(
        ExcelWorksheetData worksheet,
        TemplateDefinition template) =>
        Validate(worksheet, template, null);

    public RuntimeWorksheetContext Validate(
        ExcelWorksheetData worksheet,
        TemplateDefinition template,
        int? explicitRuntimeMappingMarkerRowNumber)
    {
        ArgumentNullException.ThrowIfNull(worksheet);
        ArgumentNullException.ThrowIfNull(template);

        if (template.FieldMappings.GroupBy(mapping => mapping.ColumnIndex).Any(group => group.Count() > 1))
        {
            throw new TemplateDefinitionException(
                TemplateDefinitionErrorCode.InvalidFieldMapping,
                "模板欄位對應包含重複的 Excel 欄位，無法進行 runtime 驗證。");
        }

        var mismatches = template.FieldMappings
            .Where(mapping => !string.IsNullOrWhiteSpace(mapping.HeaderName))
            .Select(mapping =>
            {
                var actual = mapping.ColumnIndex <= worksheet.Headers.Count
                    ? worksheet.Headers[mapping.ColumnIndex - 1]
                    : null;
                return new { Mapping = mapping, Actual = actual };
            })
            .Where(item => !string.Equals(
                item.Mapping.HeaderName?.Trim(),
                item.Actual?.Trim(),
                StringComparison.OrdinalIgnoreCase))
            .ToArray();

        if (mismatches.Length > 0)
        {
            var details = string.Join(
                "; ",
                mismatches.Select(item =>
                    $"{item.Mapping.ColumnLetter}: 預期「{item.Mapping.HeaderName}」，實際「{item.Actual ?? "(空白)"}」"));
            throw new TemplateDefinitionException(
                TemplateDefinitionErrorCode.ExcelSchemaMismatch,
                $"Excel 欄位結構與模板不一致：{details}。程式不會自動重新對應欄位。");
        }

        if (explicitRuntimeMappingMarkerRowNumber.HasValue)
        {
            if (!RuntimeMappingMarkerDetector.IsValidExplicitRow(
                worksheet,
                template,
                explicitRuntimeMappingMarkerRowNumber.Value))
            {
                throw new TemplateDefinitionException(
                    TemplateDefinitionErrorCode.InvalidRuntimeMappingMarker,
                    $"目前 Excel 第 {explicitRuntimeMappingMarkerRowNumber.Value} 列不符合模板的欄位對應列格式，因此不會排除該列。");
            }

            return new RuntimeWorksheetContext(
                worksheet,
                explicitRuntimeMappingMarkerRowNumber.Value);
        }

        var markerRows = RuntimeMappingMarkerDetector.FindCandidates(worksheet, template);
        if (markerRows.Count > 1)
        {
            throw new TemplateDefinitionException(
                TemplateDefinitionErrorCode.AmbiguousRuntimeMarkerRows,
                $"目前 Excel 找到多個完整欄位對應列（{string.Join(", ", markerRows)}），請只保留一列。");
        }

        return new RuntimeWorksheetContext(
            worksheet,
            markerRows.Count == 0 ? null : markerRows[0]);
    }
}

public sealed record RuntimeWorksheetContext(
    ExcelWorksheetData Worksheet,
    int? RuntimeMappingMarkerRowNumber);

public static class RuntimeMappingMarkerDetector
{
    public static bool IsValidExplicitRow(
        ExcelWorksheetData worksheet,
        TemplateDefinition template,
        int rowNumber)
    {
        if (rowNumber <= worksheet.HeaderRowNumber)
        {
            return false;
        }

        var row = worksheet.Rows.FirstOrDefault(item => item.RowNumber == rowNumber);
        if (row is null)
        {
            return false;
        }

        var mappings = template.FieldMappings
            .GroupBy(mapping => mapping.ColumnIndex)
            .Select(group => group.Single())
            .ToArray();
        if (mappings.Length == 0)
        {
            return false;
        }

        var expectedColumns = mappings.Select(mapping => mapping.ColumnIndex).ToHashSet();
        var actualNames = mappings
            .Select(mapping => NormalizePlaceholder(row.GetCellText(mapping.ColumnIndex)))
            .ToArray();
        var expectedNames = mappings
            .Select(mapping => mapping.PlaceholderName.ToUpperInvariant())
            .ToHashSet(StringComparer.OrdinalIgnoreCase);

        return actualNames.All(name => !string.IsNullOrWhiteSpace(name))
            && actualNames.Length == actualNames.Distinct(StringComparer.OrdinalIgnoreCase).Count()
            && actualNames.ToHashSet(StringComparer.OrdinalIgnoreCase).SetEquals(expectedNames)
            && row.Cells.Keys.All(expectedColumns.Contains);
    }

    public static IReadOnlyList<int> FindCandidates(
        ExcelWorksheetData worksheet,
        TemplateDefinition template)
    {
        var mappings = template.FieldMappings
            .GroupBy(mapping => mapping.ColumnIndex)
            .Select(group => group.Single())
            .ToArray();

        // A one-field row containing {{FIELD}} is indistinguishable from ordinary
        // data. Stay conservative rather than silently dropping that data row.
        if (mappings.Length < 2)
        {
            return [];
        }

        return worksheet.Rows
            .Where(row => row.RowNumber > worksheet.HeaderRowNumber)
            .Where(row => IsValidExplicitRow(worksheet, template, row.RowNumber))
            .Select(row => row.RowNumber)
            .OrderBy(rowNumber => rowNumber)
            .ToArray();
    }

    private static string? NormalizePlaceholder(string? value)
    {
        var trimmed = value?.Trim();
        if (string.IsNullOrWhiteSpace(trimmed)
            || !trimmed.StartsWith("{{", StringComparison.Ordinal)
            || !trimmed.EndsWith("}}", StringComparison.Ordinal))
        {
            return null;
        }

        var name = trimmed[2..^2];
        return name.Length > 0 && name.All(character => char.IsLetterOrDigit(character) || character == '_')
            ? name.ToUpperInvariant()
            : null;
    }
}
