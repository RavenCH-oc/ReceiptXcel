namespace XlsxDocxGenerator.Models;

/// <summary>
/// A generic Excel row. The generator intentionally does not know business-specific columns.
/// </summary>
public sealed record ExcelRecord
{
    public required int RowNumber { get; init; }

    public required IReadOnlyDictionary<string, string?> Values { get; init; }

    public IReadOnlyDictionary<int, string?> ColumnValues { get; init; } =
        new Dictionary<int, string?>();

    public string? this[string header] => Values.TryGetValue(header, out var value) ? value : null;

    public string? GetValueByColumnIndex(int columnIndex) =>
        ColumnValues.TryGetValue(columnIndex, out var value) ? value : null;
}
