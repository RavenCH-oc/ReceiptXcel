namespace XlsxDocxGenerator.Services.Errors;

public enum GenerationErrorCode
{
    WordTemplateNotFound,
    InvalidWordTemplate,
    MissingMapping,
    InvalidExcelDataRow,
    OutputAlreadyExists,
    OutputWriteFailed,
    InvalidOutputFileName,
    PlaceholderReplacementFailed
}

/// <summary>
/// User-facing error for Word validation and single-document generation.
/// </summary>
public sealed class GenerationException : Exception
{
    public GenerationException(
        GenerationErrorCode code,
        string userMessage,
        Exception? innerException = null)
        : base(userMessage, innerException)
    {
        Code = code;
        UserMessage = userMessage;
    }

    public GenerationErrorCode Code { get; }

    public string UserMessage { get; }
}
