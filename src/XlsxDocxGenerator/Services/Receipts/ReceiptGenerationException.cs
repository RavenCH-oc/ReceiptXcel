namespace XlsxDocxGenerator.Services.Receipts;

public enum ReceiptGenerationErrorCode
{
    ReceiptTemplateInvalid,
    OutputAlreadyExists,
    InvalidOutputPath,
    OutputWriteFailed,
    GeneratedDocumentInvalid
}

/// <summary>
/// Friendly boundary error for the fixed receipt generation workflow.
/// </summary>
public sealed class ReceiptGenerationException : Exception
{
    public ReceiptGenerationException(
        ReceiptGenerationErrorCode code,
        string userMessage,
        Exception? innerException = null)
        : base(userMessage, innerException)
    {
        Code = code;
        UserMessage = userMessage;
    }

    public ReceiptGenerationErrorCode Code { get; }

    public string UserMessage { get; }
}
