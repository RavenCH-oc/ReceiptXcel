using XlsxDocxGenerator.Services.Excel;
using XlsxDocxGenerator.Models;

namespace XlsxDocxGenerator.Services.Receipts;

/// <summary>
/// The production receipt register schema. It deliberately does not infer or
/// remap columns: the fixed form only accepts the exact A:H contract.
/// </summary>
public sealed class ReceiptWorksheetSchema
{
    public const string WorksheetName = "工作表1";

    public const int HeaderRowNumber = 2;

    public static IReadOnlyList<string> ExpectedHeaders { get; } =
    [
        "年",
        "月",
        "日",
        "編號",
        "繳款人",
        "數字金額",
        "事由",
        "承辦人"
    ];

    public void Validate(ExcelWorksheetData worksheet)
    {
        ArgumentNullException.ThrowIfNull(worksheet);

        if (!string.Equals(worksheet.WorksheetName, WorksheetName, StringComparison.Ordinal))
        {
            throw new ReceiptValidationException(
                ReceiptValidationErrorCode.ExcelSchemaMismatch,
                $"Excel 格式不符：工作表預期為「{WorksheetName}」，實際為「{worksheet.WorksheetName}」。");
        }

        if (worksheet.HeaderRowNumber != HeaderRowNumber)
        {
            throw new ReceiptValidationException(
                ReceiptValidationErrorCode.ExcelSchemaMismatch,
                $"Excel 格式不符：固定 header 必須位於第 {HeaderRowNumber} 列。");
        }

        for (var index = 0; index < ExpectedHeaders.Count; index++)
        {
            var actual = index < worksheet.Headers.Count
                ? worksheet.Headers[index]?.Trim()
                : null;

            if (string.Equals(actual, ExpectedHeaders[index], StringComparison.Ordinal))
            {
                continue;
            }

            var column = ExcelColumnName.FromIndex(index + 1);
            throw new ReceiptValidationException(
                ReceiptValidationErrorCode.ExcelSchemaMismatch,
                $"Excel 格式不符：{column} 欄預期為「{ExpectedHeaders[index]}」，實際為「{actual ?? "空白"}」。");
        }

        if (worksheet.Headers.Count > ExpectedHeaders.Count)
        {
            var column = ExcelColumnName.FromIndex(ExpectedHeaders.Count + 1);
            throw new ReceiptValidationException(
                ReceiptValidationErrorCode.ExcelSchemaMismatch,
                $"Excel 格式不符：{column} 欄不屬於固定收據格式，程式不會自動重新對應欄位。");
        }
    }
}
