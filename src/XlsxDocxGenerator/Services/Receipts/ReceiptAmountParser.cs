using System.Globalization;
using System.Text.RegularExpressions;

namespace XlsxDocxGenerator.Services.Receipts;

public sealed class ReceiptAmountParser
{
    public const int MaxAmount = 9_999_999;

    private static readonly Regex IntegerPattern = new(
        @"^(?:\d+|\d{1,3}(?:,\d{3})+)$",
        RegexOptions.Compiled | RegexOptions.CultureInvariant);

    public int Parse(string? rawValue, int? excelRowNumber = null)
    {
        var value = rawValue?.Trim();
        if (string.IsNullOrWhiteSpace(value) || !IntegerPattern.IsMatch(value))
        {
            throw InvalidAmount(excelRowNumber);
        }

        var normalized = value.Replace(",", string.Empty, StringComparison.Ordinal);
        if (!long.TryParse(normalized, NumberStyles.None, CultureInfo.InvariantCulture, out var amount)
            || amount <= 0)
        {
            throw InvalidAmount(excelRowNumber);
        }

        return EnsureWithinFormLimit(amount, excelRowNumber);
    }

    public int Parse(long rawValue, int? excelRowNumber = null)
    {
        if (rawValue <= 0)
        {
            throw InvalidAmount(excelRowNumber);
        }

        return EnsureWithinFormLimit(rawValue, excelRowNumber);
    }

    private static int EnsureWithinFormLimit(long amount, int? excelRowNumber)
    {
        if (amount > MaxAmount)
        {
            var formatted = amount.ToString("#,##0", CultureInfo.InvariantCulture);
            var prefix = excelRowNumber.HasValue
                ? $"第 {excelRowNumber.Value} 列"
                : "金額";
            throw new ReceiptValidationException(
                ReceiptValidationErrorCode.AmountExceedsLimit,
                $"{prefix}金額 {formatted} 超過表單上限 {MaxAmount:#,##0}，請確認 Excel 資料。");
        }

        return (int)amount;
    }

    private static ReceiptValidationException InvalidAmount(int? excelRowNumber) =>
        new(
            ReceiptValidationErrorCode.InvalidAmount,
            excelRowNumber.HasValue
                ? $"第 {excelRowNumber.Value} 列「數字金額」格式不正確，請輸入不含角分的正整數。"
                : "「數字金額」格式不正確，請輸入不含角分的正整數。\n可接受：1000、1,000；不接受貨幣符號、負數或小數。");
}
