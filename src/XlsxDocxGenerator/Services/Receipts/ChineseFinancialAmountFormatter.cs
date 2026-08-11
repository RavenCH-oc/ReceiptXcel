namespace XlsxDocxGenerator.Services.Receipts;

public sealed class ChineseFinancialAmountFormatter
{
    private static readonly string[] Digits =
    [
        "零", "壹", "貳", "參", "肆", "伍", "陸", "柒", "捌", "玖"
    ];

    private static readonly string[] SmallUnits = ["", "拾", "佰", "仟"];

    public string Format(int amount)
    {
        if (amount is <= 0 or > ReceiptAmountParser.MaxAmount)
        {
            throw new ArgumentOutOfRangeException(
                nameof(amount),
                $"金額必須介於 1 與 {ReceiptAmountParser.MaxAmount:#,##0} 之間。");
        }

        var highGroup = amount / 10_000;
        var lowGroup = amount % 10_000;
        var result = string.Empty;

        if (highGroup > 0)
        {
            result = FormatUnderTenThousand(highGroup) + "萬";
            if (lowGroup > 0)
            {
                if (lowGroup < 1_000)
                {
                    result += "零";
                }

                result += FormatUnderTenThousand(lowGroup);
            }
        }
        else
        {
            result = FormatUnderTenThousand(lowGroup);
        }

        return result + "元整";
    }

    private static string FormatUnderTenThousand(int value)
    {
        var result = string.Empty;
        var zeroPending = false;

        for (var position = 3; position >= 0; position--)
        {
            var divisor = (int)Math.Pow(10, position);
            var digit = value / divisor % 10;
            if (digit == 0)
            {
                if (result.Length > 0)
                {
                    zeroPending = true;
                }

                continue;
            }

            if (zeroPending)
            {
                result += Digits[0];
                zeroPending = false;
            }

            result += Digits[digit] + SmallUnits[position];
        }

        return result;
    }
}
