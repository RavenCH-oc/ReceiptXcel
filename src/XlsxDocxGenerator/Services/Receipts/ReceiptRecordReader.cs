using System.Globalization;
using System.Text.RegularExpressions;
using XlsxDocxGenerator.Models;
using XlsxDocxGenerator.Services.Excel;

namespace XlsxDocxGenerator.Services.Receipts;

/// <summary>
/// Reads the fixed receipt worksheet and converts rows into closed domain
/// records. It never exposes ClosedXML cells to callers.
/// </summary>
public sealed class ReceiptRecordReader
{
    private static readonly Regex PositiveIntegerPattern = new(
        @"^\d+$",
        RegexOptions.Compiled | RegexOptions.CultureInvariant);

    private readonly IExcelReader _excelReader;
    private readonly ReceiptWorksheetSchema _schema;
    private readonly ReceiptAmountParser _amountParser;

    public ReceiptRecordReader(
        IExcelReader? excelReader = null,
        ReceiptWorksheetSchema? schema = null,
        ReceiptAmountParser? amountParser = null)
    {
        _excelReader = excelReader ?? new ExcelReader();
        _schema = schema ?? new ReceiptWorksheetSchema();
        _amountParser = amountParser ?? new ReceiptAmountParser();
    }

    public async Task<IReadOnlyList<ReceiptRecord>> ReadAsync(
        string workbookPath,
        CancellationToken cancellationToken = default)
    {
        var worksheet = await ReadWorksheetAsync(
            workbookPath,
            ReceiptWorksheetSchema.WorksheetName,
            cancellationToken);
        return worksheet.Rows.Select(ParseRecord).ToArray();
    }

    public async Task<ReceiptRecord> ReadRecordAsync(
        string workbookPath,
        int rowNumber,
        CancellationToken cancellationToken = default)
    {
        return await ReadRecordAsync(
            workbookPath,
            ReceiptWorksheetSchema.WorksheetName,
            rowNumber,
            cancellationToken);
    }

    public async Task<ReceiptRecord> ReadRecordAsync(
        string workbookPath,
        string worksheetName,
        int rowNumber,
        CancellationToken cancellationToken = default)
    {
        if (rowNumber <= ReceiptWorksheetSchema.HeaderRowNumber)
        {
            throw new ReceiptValidationException(
                ReceiptValidationErrorCode.RecordNotFound,
                $"Excel 第 {rowNumber} 列不是可產生收據的資料列。");
        }

        var worksheet = await ReadWorksheetAsync(workbookPath, worksheetName, cancellationToken);
        var row = worksheet.Rows.FirstOrDefault(item => item.RowNumber == rowNumber);
        if (row is null)
        {
            throw new ReceiptValidationException(
                ReceiptValidationErrorCode.RecordNotFound,
                $"找不到 Excel 第 {rowNumber} 列，或該列沒有有效資料。");
        }

        return ParseRecord(row);
    }

    private async Task<ExcelWorksheetData> ReadWorksheetAsync(
        string workbookPath,
        string worksheetName,
        CancellationToken cancellationToken)
    {
        var worksheet = await _excelReader.ReadAsync(
            workbookPath,
            worksheetName,
            ReceiptWorksheetSchema.HeaderRowNumber,
            cancellationToken);
        _schema.Validate(worksheet);
        return worksheet;
    }

    private ReceiptRecord ParseRecord(ExcelRowData row)
    {
        var rocYear = ParsePositiveInteger(row, 1, "年");
        var month = ParsePositiveInteger(row, 2, "月");
        var day = ParsePositiveInteger(row, 3, "日");
        ValidateDate(row.RowNumber, rocYear, month, day);

        var serial = Required(row, 4, "編號");
        if (!PositiveIntegerPattern.IsMatch(serial))
        {
            throw new ReceiptValidationException(
                ReceiptValidationErrorCode.InvalidReceiptSerial,
                $"第 {row.RowNumber} 列編號格式錯誤，請使用數字且保留 Excel 顯示的前導零。");
        }

        return new ReceiptRecord
        {
            ExcelRowNumber = row.RowNumber,
            ROCYear = rocYear,
            Month = month,
            Day = day,
            ReceiptSerial = serial,
            Payer = Required(row, 5, "繳款人"),
            Amount = _amountParser.Parse(Required(row, 6, "數字金額"), row.RowNumber),
            Reason = Required(row, 7, "事由"),
            Handler = Optional(row, 8)
        };
    }

    private static int ParsePositiveInteger(ExcelRowData row, int column, string name)
    {
        var value = Required(row, column, name);
        if (!PositiveIntegerPattern.IsMatch(value)
            || !int.TryParse(value, NumberStyles.None, CultureInfo.InvariantCulture, out var result)
            || result <= 0)
        {
            throw new ReceiptValidationException(
                ReceiptValidationErrorCode.InvalidRecord,
                $"第 {row.RowNumber} 列{name}格式錯誤，請輸入正整數。");
        }

        return result;
    }

    private static void ValidateDate(int rowNumber, int rocYear, int month, int day)
    {
        try
        {
            _ = new DateOnly(rocYear + 1911, month, day);
        }
        catch (ArgumentOutOfRangeException exception)
        {
            throw new ReceiptValidationException(
                ReceiptValidationErrorCode.InvalidDate,
                $"第 {rowNumber} 列日期無效：民國 {rocYear} 年 {month} 月 {day} 日。",
                exception);
        }
    }

    private static string Required(ExcelRowData row, int column, string name)
    {
        var value = row.GetCellText(column)?.Trim();
        if (!string.IsNullOrWhiteSpace(value))
        {
            return value;
        }

        throw new ReceiptValidationException(
            ReceiptValidationErrorCode.InvalidRecord,
            $"第 {row.RowNumber} 列缺少{name}資料。");
    }

    private static string Optional(ExcelRowData row, int column) =>
        row.GetCellText(column)?.Trim() ?? string.Empty;
}
