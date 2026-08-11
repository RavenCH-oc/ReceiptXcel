namespace XlsxDocxGenerator.Services.Receipts;

public sealed record AmountDigitCells(
    string Million,
    string HundredThousand,
    string TenThousand,
    string Thousand,
    string Hundred,
    string Ten,
    string One);

/// <summary>
/// Formats a positive receipt amount into the seven fixed cells. The dollar
/// sign occupies the first non-zero cell only when a higher cell exists.
/// </summary>
public sealed class AmountDigitFormatter
{
    public AmountDigitCells Format(int amount)
    {
        if (amount is <= 0 or > ReceiptAmountParser.MaxAmount)
        {
            throw new ArgumentOutOfRangeException(
                nameof(amount),
                $"金額必須介於 1 與 {ReceiptAmountParser.MaxAmount:#,##0} 之間。");
        }

        var digits = amount.ToString("D7", System.Globalization.CultureInfo.InvariantCulture)
            .Select(character => character.ToString())
            .ToArray();
        var firstNonZero = Array.FindIndex(digits, digit => digit != "0");

        for (var index = 0; index < firstNonZero; index++)
        {
            digits[index] = string.Empty;
        }

        if (firstNonZero > 0)
        {
            digits[firstNonZero - 1] = "$";
        }

        return new AmountDigitCells(
            digits[0],
            digits[1],
            digits[2],
            digits[3],
            digits[4],
            digits[5],
            digits[6]);
    }
}
