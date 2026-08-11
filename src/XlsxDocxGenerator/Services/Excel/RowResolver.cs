using XlsxDocxGenerator.Models;
using XlsxDocxGenerator.Services.Errors;

namespace XlsxDocxGenerator.Services.Excel;

public sealed class RowResolver
{
    public RowSelectionResult Resolve(
        ExcelWorksheetData worksheet,
        TemplateDefinition template,
        SelectionRule rule) =>
        Resolve(worksheet, template, rule, template.CreationMappingMarkerRowNumber);

    public RowSelectionResult Resolve(
        ExcelWorksheetData worksheet,
        TemplateDefinition template,
        SelectionRule rule,
        int? runtimeMarkerRowNumber)
    {
        ArgumentNullException.ThrowIfNull(worksheet);
        ArgumentNullException.ThrowIfNull(template);
        ArgumentNullException.ThrowIfNull(rule);

        var availableRecords = worksheet.Records
            .Where(record => record.RowNumber > template.HeaderRowNumber
                && record.RowNumber != runtimeMarkerRowNumber)
            .OrderBy(record => record.RowNumber)
            .ToArray();

        if (rule is SelectionRule.Marker markerRule
            && !worksheet.Headers.Any(header =>
                string.Equals(header, markerRule.SelectionMarkerColumnHeader, StringComparison.OrdinalIgnoreCase)))
        {
            throw new SelectionValidationException(
                SelectionErrorCode.MissingMarkerColumn,
                $"工作表中找不到標記欄位「{markerRule.SelectionMarkerColumnHeader}」。");
        }

        return rule switch
        {
            SelectionRule.Marker marker => ResolveByMarker(availableRecords, marker),
            SelectionRule.ExcelRows rows => ResolveByRows(availableRecords, rows, template, runtimeMarkerRowNumber),
            SelectionRule.Latest latest => ResolveLatest(availableRecords, latest),
            _ => throw new SelectionValidationException(
                SelectionErrorCode.InvalidSyntax,
                "不支援的資料選擇規則。")
        };
    }

    private static RowSelectionResult ResolveByMarker(
        IReadOnlyList<ExcelRecord> records,
        SelectionRule.Marker rule)
    {
        var requestedTokens = rule.Tokens.ToHashSet(StringComparer.OrdinalIgnoreCase);
        var selected = records
            .Where(record => SplitTokens(record[rule.SelectionMarkerColumnHeader])
                .Any(requestedTokens.Contains))
            .Select(record => RowSelectionItem.Success(record))
            .ToArray();

        return new RowSelectionResult(selected, records.Count, null);
    }

    private static RowSelectionResult ResolveByRows(
        IReadOnlyList<ExcelRecord> records,
        SelectionRule.ExcelRows rule,
        TemplateDefinition template,
        int? runtimeMarkerRowNumber)
    {
        var recordsByRow = records.ToDictionary(record => record.RowNumber);
        var selections = new List<RowSelectionItem>();

        foreach (var rowNumber in rule.RowNumbers)
        {
            if (rowNumber <= template.HeaderRowNumber)
            {
                selections.Add(RowSelectionItem.Failure(
                    rowNumber,
                    GenerationErrorCode.InvalidExcelDataRow,
                    $"Excel 第 {rowNumber} 列是 header row 或位於 header row 之前。"));
            }
            else if (rowNumber == runtimeMarkerRowNumber)
            {
                selections.Add(RowSelectionItem.Failure(
                    rowNumber,
                    GenerationErrorCode.InvalidExcelDataRow,
                    $"Excel 第 {rowNumber} 列是 mapping marker row，不是正式資料列。"));
            }
            else if (recordsByRow.TryGetValue(rowNumber, out var record))
            {
                selections.Add(RowSelectionItem.Success(record));
            }
            else
            {
                selections.Add(RowSelectionItem.Failure(
                    rowNumber,
                    GenerationErrorCode.InvalidExcelDataRow,
                    $"Excel 第 {rowNumber} 列不存在或是空白資料列。"));
            }
        }

        return new RowSelectionResult(selections, records.Count, null);
    }

    private static RowSelectionResult ResolveLatest(
        IReadOnlyList<ExcelRecord> records,
        SelectionRule.Latest rule)
    {
        var selected = records
            .TakeLast(rule.Count)
            .Select(record => RowSelectionItem.Success(record))
            .ToArray();

        var notice = rule.Count > records.Count
            ? $"要求 {rule.Count} 筆，實際找到 {records.Count} 筆。"
            : null;

        return new RowSelectionResult(selected, records.Count, notice);
    }

    private static IEnumerable<string> SplitTokens(string? value) =>
        (value ?? string.Empty)
            .Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
}

public sealed record RowSelectionResult(
    IReadOnlyList<RowSelectionItem> Items,
    int AvailableRecordCount,
    string? Notice)
{
    public int TotalSelected => Items.Count;

    public IReadOnlyList<ExcelRecord> Records => Items
        .Where(item => item.Record is not null)
        .Select(item => item.Record!)
        .ToArray();
}

public sealed record RowSelectionItem(
    int RowNumber,
    ExcelRecord? Record,
    GenerationErrorCode? ErrorCode,
    string? ErrorMessage)
{
    public bool IsValid => Record is not null;

    public static RowSelectionItem Success(ExcelRecord record) =>
        new(record.RowNumber, record, null, null);

    public static RowSelectionItem Failure(
        int rowNumber,
        GenerationErrorCode errorCode,
        string message) =>
        new(rowNumber, null, errorCode, message);
}
