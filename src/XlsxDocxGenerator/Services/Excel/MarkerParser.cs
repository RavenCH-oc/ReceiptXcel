using System.Text.RegularExpressions;
using XlsxDocxGenerator.Models;
using XlsxDocxGenerator.Services.Errors;

namespace XlsxDocxGenerator.Services.Excel;

public sealed class MarkerParser
{
    private static readonly Regex PlaceholderPattern = new(
        "^\\{\\{(?<name>[A-Za-z_][A-Za-z0-9_]*)\\}\\}$",
        RegexOptions.Compiled | RegexOptions.CultureInvariant);

    /// <summary>
    /// Parses a cell as a placeholder and returns the canonical uppercase key.
    /// </summary>
    public bool TryParse(string? cellText, out string placeholderName)
    {
        var match = PlaceholderPattern.Match(cellText?.Trim() ?? string.Empty);
        if (!match.Success)
        {
            placeholderName = string.Empty;
            return false;
        }

        placeholderName = match.Groups["name"].Value.ToUpperInvariant();
        return true;
    }

    public bool LooksLikePlaceholder(string? cellText)
    {
        var value = cellText?.Trim() ?? string.Empty;
        return value.Contains('{', StringComparison.Ordinal)
            || value.Contains('}', StringComparison.Ordinal);
    }

    public bool ContainsValidMarker(ExcelRowData row)
    {
        ArgumentNullException.ThrowIfNull(row);
        return row.Cells.Values.Any(value => TryParse(value, out _));
    }

    public IReadOnlyList<FieldMapping> ParseMappingRow(
        ExcelRowData row,
        IReadOnlyList<string?> headers)
    {
        ArgumentNullException.ThrowIfNull(row);
        ArgumentNullException.ThrowIfNull(headers);

        var mappings = new List<FieldMapping>();
        var seenNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        foreach (var cell in row.Cells.OrderBy(item => item.Key))
        {
            if (string.IsNullOrWhiteSpace(cell.Value))
            {
                continue;
            }

            if (!TryParse(cell.Value, out var placeholderName))
            {
                if (LooksLikePlaceholder(cell.Value))
                {
                    throw new TemplateDefinitionException(
                        TemplateDefinitionErrorCode.InvalidMarkerSyntax,
                        $"Marker row 第 {row.RowNumber} 列的 placeholder 格式無效：{cell.Value}");
                }

                continue;
            }

            if (!seenNames.Add(placeholderName))
            {
                throw new TemplateDefinitionException(
                    TemplateDefinitionErrorCode.DuplicatePlaceholder,
                    $"Marker row 重複使用 placeholder「{placeholderName}」。");
            }

            mappings.Add(new FieldMapping
            {
                PlaceholderName = placeholderName,
                ColumnIndex = cell.Key,
                HeaderName = cell.Key <= headers.Count ? headers[cell.Key - 1] : null
            });
        }

        if (mappings.Count == 0)
        {
            throw new TemplateDefinitionException(
                TemplateDefinitionErrorCode.InvalidMarkerRow,
                $"Excel 第 {row.RowNumber} 列沒有任何合法 placeholder marker。");
        }

        return mappings;
    }
}
