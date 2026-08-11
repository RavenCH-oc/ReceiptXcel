namespace XlsxDocxGenerator.Services.Receipts;

public enum ReceiptValidationErrorCode
{
    ExcelSchemaMismatch,
    WorksheetNotFound,
    RecordNotFound,
    InvalidRecord,
    InvalidDate,
    InvalidReceiptSerial,
    InvalidAmount,
    AmountExceedsLimit
}

/// <summary>
/// Friendly, fail-closed validation error for the fixed receipt format.
/// </summary>
public sealed class ReceiptValidationException : Exception
{
    public ReceiptValidationException(
        ReceiptValidationErrorCode code,
        string userMessage,
        Exception? innerException = null)
        : base(userMessage, innerException)
    {
        Code = code;
        UserMessage = userMessage;
    }

    public ReceiptValidationErrorCode Code { get; }

    public string UserMessage { get; }
}
