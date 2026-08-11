using XlsxDocxGenerator.Models;

namespace XlsxDocxGenerator.Services.Excel;

public sealed record ExcelWorksheetData
{
    public required string WorksheetName { get; init; }

    public required int HeaderRowNumber { get; init; }

    public required int FirstUsedRowNumber { get; init; }

    public required int LastUsedRowNumber { get; init; }

    /// <summary>
    /// One-based Excel columns represented by zero-based list positions.
    /// Blank headers are represented by null.
    /// </summary>
    public required IReadOnlyList<string?> Headers { get; init; }

    /// <summary>
    /// Non-empty rows below the header row. Blank rows are intentionally omitted.
    /// </summary>
    public required IReadOnlyList<ExcelRowData> Rows { get; init; }

    /// <summary>Detected from the current workbook, never copied from a saved template.</summary>
    public int? RuntimeMappingMarkerRowNumber { get; init; }

    public IReadOnlyList<ExcelRecord> Records => Rows
        .Select(row => new ExcelRecord
        {
            RowNumber = row.RowNumber,
            ColumnValues = row.Cells,
            Values = Headers
                .Select((header, index) => new { Header = header, ColumnIndex = index + 1 })
                .Where(item => !string.IsNullOrWhiteSpace(item.Header))
                .GroupBy(item => item.Header!, StringComparer.OrdinalIgnoreCase)
                .ToDictionary(
                    group => group.Key,
                    group => row.GetCellText(group.First().ColumnIndex),
                    StringComparer.OrdinalIgnoreCase)
        })
        .ToArray();
}

public sealed record ExcelRowData
{
    public required int RowNumber { get; init; }

    /// <summary>
    /// One-based Excel column number to displayed/text cell value.
    /// </summary>
    public required IReadOnlyDictionary<int, string?> Cells { get; init; }

    public string? GetCellText(int columnIndex) =>
        Cells.TryGetValue(columnIndex, out var value) ? value : null;
}
