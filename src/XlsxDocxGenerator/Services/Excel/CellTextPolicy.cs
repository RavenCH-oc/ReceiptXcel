using ClosedXML.Excel;

namespace XlsxDocxGenerator.Services.Excel;

public interface ICellTextPolicy
{
    string? GetText(IXLCell cell);
}

/// <summary>
/// Uses ClosedXML formatted display text so number formats, leading zeroes,
/// dates and percentages follow what the user sees in Excel.
/// </summary>
public sealed class DisplayCellTextPolicy : ICellTextPolicy
{
    public string? GetText(IXLCell cell)
    {
        ArgumentNullException.ThrowIfNull(cell);

        var text = cell.GetFormattedString().Trim();
        return text.Length == 0 ? null : text;
    }
}
