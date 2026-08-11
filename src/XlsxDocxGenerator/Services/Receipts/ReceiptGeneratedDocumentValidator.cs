using System.IO;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using XlsxDocxGenerator.Models;

namespace XlsxDocxGenerator.Services.Receipts;

/// <summary>
/// Validates the generated artifact, including the fixed text contract and
/// each of the three receipt tables.
/// </summary>
public sealed class ReceiptGeneratedDocumentValidator
{
    private static readonly Regex AnyPlaceholder = new(
        @"\{\{[^{}]*\}\}",
        RegexOptions.Compiled | RegexOptions.CultureInvariant);

    public void Validate(
        string documentPath,
        ReceiptRecord record,
        FormattedReceiptAmount amount)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(documentPath);
        ArgumentNullException.ThrowIfNull(record);
        ArgumentNullException.ThrowIfNull(amount);

        try
        {
            using var document = WordprocessingDocument.Open(documentPath, false);
            var body = document.MainDocumentPart?.Document.Body;
            if (body is null)
            {
                throw InvalidOutput("DOCX 缺少主要文件內容。");
            }

            var allText = string.Concat(body.Descendants<Text>().Select(text => text.Text ?? string.Empty));
            if (AnyPlaceholder.IsMatch(allText))
            {
                throw InvalidOutput("產生後仍存在未替換 placeholder。");
            }

            RequireCount(allText, record.Payer, 3, "繳款人");
            RequireCount(allText, record.Reason, 3, "事由");
            RequireCount(allText, record.ReceiptNumber, 3, "收據編號");
            RequireCount(allText, amount.ChineseUppercase, 3, "中文財務大寫");
            RequireCount(
                allText,
                $"中華民國{record.ROCYear}年{record.Month}月{record.Day}日",
                3,
                "日期");

            foreach (var fixedText in new[]
            {
                "南投縣水里鄉郡坑國民小學 自行收納款項統一收據",
                "【第三聯】",
                "【第二聯】",
                "【第一聯】",
                "郡坑收字第",
                "號",
                "21813799"
            })
            {
                if (!allText.Contains(fixedText, StringComparison.Ordinal))
                {
                    throw InvalidOutput($"固定文字不存在：{fixedText}");
                }
            }

            var tables = body.Descendants<Table>().ToArray();
            if (tables.Length != 3)
            {
                throw InvalidOutput($"收據表格數量為 {tables.Length}，預期為 3。");
            }

            var expectedCells = new[]
            {
                amount.MillionCell,
                amount.HundredThousandCell,
                amount.TenThousandCell,
                amount.ThousandCell,
                amount.HundredCell,
                amount.TenCell,
                amount.OneCell
            };

            foreach (var table in tables)
            {
                var rows = table.Elements<TableRow>().ToArray();
                var grid = table.GetFirstChild<TableGrid>();
                if (rows.Length != 5 || grid?.Elements<GridColumn>().Count() != 17)
                {
                    throw InvalidOutput("收據表格的列數或 grid 結構已改變。");
                }

                var dataCells = rows[2].Elements<TableCell>().ToArray();
                if (dataCells.Length < 9)
                {
                    throw InvalidOutput("收據金額資料列缺少固定欄位。");
                }

                for (var index = 0; index < expectedCells.Length; index++)
                {
                    var actual = GetCellText(dataCells[index + 2]);
                    if (!string.Equals(actual, expectedCells[index], StringComparison.Ordinal))
                    {
                        throw InvalidOutput(
                            $"金額第 {index + 1} 格不符：實際「{actual}」，預期「{expectedCells[index]}」。");
                    }
                }

                var handlerCells = rows[4].Elements<TableCell>().ToArray();
                if (handlerCells.Length < 2)
                {
                    throw InvalidOutput("找不到經手人填寫格。");
                }

                var handlerText = GetCellText(handlerCells[1]);
                if (!string.Equals(handlerText, record.Handler, StringComparison.Ordinal))
                {
                    throw InvalidOutput("經手人填寫格與承辦人資料不一致。");
                }
            }
        }
        catch (ReceiptGenerationException)
        {
            throw;
        }
        catch (Exception exception) when (exception is IOException
            or InvalidDataException
            or OpenXmlPackageException
            or UnauthorizedAccessException)
        {
            throw InvalidOutput("無法重新開啟或檢查產生的 DOCX。", exception);
        }
    }

    private static string GetCellText(TableCell cell) =>
        string.Concat(cell.Descendants<Text>().Select(text => text.Text ?? string.Empty));

    private static void RequireCount(string text, string value, int expected, string fieldName)
    {
        if (string.IsNullOrEmpty(value))
        {
            return;
        }

        var actual = CountOccurrences(text, value);
        if (actual != expected)
        {
            throw InvalidOutput($"{fieldName}出現 {actual} 次，預期 {expected} 次。");
        }
    }

    private static int CountOccurrences(string text, string value)
    {
        var count = 0;
        var index = 0;
        while ((index = text.IndexOf(value, index, StringComparison.Ordinal)) >= 0)
        {
            count++;
            index += value.Length;
        }

        return count;
    }

    private static ReceiptGenerationException InvalidOutput(
        string detail,
        Exception? innerException = null) =>
        new(
            ReceiptGenerationErrorCode.GeneratedDocumentInvalid,
            $"產生的收據 DOCX 驗證失敗：{detail}",
            innerException);
}
