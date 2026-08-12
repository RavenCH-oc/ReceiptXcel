namespace XlsxDocxGenerator.Models;

/// <summary>
/// Fixed receipt data read from one Excel row. ClosedXML objects never cross
/// the receipt domain boundary.
/// </summary>
public sealed record ReceiptRecord
{
    public required int ExcelRowNumber { get; init; }

    public required int ROCYear { get; init; }

    public required int Month { get; init; }

    public required int Day { get; init; }

    /// <summary>Displayed Excel text, including any leading zeroes.</summary>
    public required string ReceiptSerial { get; init; }

    public required string Payer { get; init; }

    public required int Amount { get; init; }

    public required string Reason { get; init; }

    /// <summary>
    /// Raw Excel column H value retained only for compatibility with the
    /// fixed A:H register schema. ReceiptXcel v0.x never maps this source-only
    /// value to Word; the three "經手人" cells are intentionally left blank.
    /// </summary>
    public required string Handler { get; init; }

    public string ReceiptNumber => $"{ROCYear}{ReceiptSerial}";

    public DateOnly GregorianDate => new(ROCYear + 1911, Month, Day);
}
