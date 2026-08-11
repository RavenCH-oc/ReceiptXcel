namespace XlsxDocxGenerator.Services.Receipts;

public sealed record FormattedReceiptAmount
{
    public required int OriginalAmount { get; init; }

    public required string MillionCell { get; init; }

    public required string HundredThousandCell { get; init; }

    public required string TenThousandCell { get; init; }

    public required string ThousandCell { get; init; }

    public required string HundredCell { get; init; }

    public required string TenCell { get; init; }

    public required string OneCell { get; init; }

    public required string ChineseUppercase { get; init; }
}
