using XlsxDocxGenerator.Models;

namespace XlsxDocxGenerator.Services.Receipts;

public sealed record ReceiptWordFieldMapping(
    string Placeholder,
    string TargetDescription);

/// <summary>
/// The only place where fixed receipt fields are mapped to internal Word
/// placeholders. In particular, Excel 承辦人 maps to each 聯's 經手人 cell.
/// </summary>
public static class ReceiptTemplateMapping
{
    public const string RelativeTemplatePath =
        "src/XlsxDocxGenerator/Assets/Templates/receipt-template.docx";

    public static IReadOnlyList<ReceiptWordFieldMapping> Fields { get; } =
    [
        new("ROC_YEAR", "三聯日期的民國年"),
        new("MONTH", "三聯日期的月份"),
        new("DAY", "三聯日期的日期"),
        new("RECEIPT_NUMBER", "三聯「郡坑收字第」與「號」之間"),
        new("PAYER", "三聯繳款人"),
        new("REASON", "三聯事由"),
        new("HANDLER", "三聯「經手人」填寫位置"),
        new("AMOUNT_1000000", "百萬格"),
        new("AMOUNT_100000", "十萬格"),
        new("AMOUNT_10000", "萬格"),
        new("AMOUNT_1000", "千格"),
        new("AMOUNT_100", "百格"),
        new("AMOUNT_10", "十格"),
        new("AMOUNT_1", "元格"),
        new("AMOUNT_UPPER", "三聯金額大寫")
    ];

    public static IReadOnlyDictionary<string, string> CreateValues(
        ReceiptRecord record,
        FormattedReceiptAmount amount)
    {
        ArgumentNullException.ThrowIfNull(record);
        ArgumentNullException.ThrowIfNull(amount);

        return new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            ["ROC_YEAR"] = record.ROCYear.ToString(),
            ["MONTH"] = record.Month.ToString(),
            ["DAY"] = record.Day.ToString(),
            ["RECEIPT_NUMBER"] = record.ReceiptNumber,
            ["PAYER"] = record.Payer,
            ["REASON"] = record.Reason,
            ["HANDLER"] = record.Handler,
            ["AMOUNT_1000000"] = amount.MillionCell,
            ["AMOUNT_100000"] = amount.HundredThousandCell,
            ["AMOUNT_10000"] = amount.TenThousandCell,
            ["AMOUNT_1000"] = amount.ThousandCell,
            ["AMOUNT_100"] = amount.HundredCell,
            ["AMOUNT_10"] = amount.TenCell,
            ["AMOUNT_1"] = amount.OneCell,
            ["AMOUNT_UPPER"] = amount.ChineseUppercase
        };
    }
}
