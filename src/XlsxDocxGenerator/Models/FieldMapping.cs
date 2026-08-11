namespace XlsxDocxGenerator.Models;

/// <summary>
/// Maps one canonical Word field name to one Excel column.
/// ColumnIndex is one-based, matching Excel's column numbering.
/// </summary>
public sealed record FieldMapping
{
    public required string PlaceholderName { get; init; }

    public required int ColumnIndex { get; init; }

    public string ColumnLetter => ExcelColumnName.FromIndex(ColumnIndex);

    public string? HeaderName { get; init; }
}

public static class ExcelColumnName
{
    public static string FromIndex(int columnIndex)
    {
        if (columnIndex <= 0)
        {
            throw new ArgumentOutOfRangeException(nameof(columnIndex), "Excel 欄號必須大於 0。");
        }

        var result = string.Empty;
        var current = columnIndex;
        while (current > 0)
        {
            current--;
            result = (char)('A' + current % 26) + result;
            current /= 26;
        }

        return result;
    }
}
