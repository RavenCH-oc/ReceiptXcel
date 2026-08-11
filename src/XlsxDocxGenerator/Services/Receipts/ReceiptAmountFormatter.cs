namespace XlsxDocxGenerator.Services.Receipts;

/// <summary>
/// Creates all deterministic amount-derived values once, before Word mapping.
/// </summary>
public sealed class ReceiptAmountFormatter
{
    private readonly AmountDigitFormatter _digitFormatter;
    private readonly ChineseFinancialAmountFormatter _chineseFormatter;

    public ReceiptAmountFormatter(
        AmountDigitFormatter? digitFormatter = null,
        ChineseFinancialAmountFormatter? chineseFormatter = null)
    {
        _digitFormatter = digitFormatter ?? new AmountDigitFormatter();
        _chineseFormatter = chineseFormatter ?? new ChineseFinancialAmountFormatter();
    }

    public FormattedReceiptAmount Format(int amount)
    {
        var cells = _digitFormatter.Format(amount);
        return new FormattedReceiptAmount
        {
            OriginalAmount = amount,
            MillionCell = cells.Million,
            HundredThousandCell = cells.HundredThousand,
            TenThousandCell = cells.TenThousand,
            ThousandCell = cells.Thousand,
            HundredCell = cells.Hundred,
            TenCell = cells.Ten,
            OneCell = cells.One,
            ChineseUppercase = _chineseFormatter.Format(amount)
        };
    }
}
