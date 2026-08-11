namespace XlsxDocxGenerator.Services.Errors;

public enum BatchFatalErrorCode
{
    ExcelReadFailed,
    WordTemplateInvalid,
    TemplateDefinitionInvalid,
    OutputDirectoryUnavailable
}

public sealed class BatchGenerationException : Exception
{
    public BatchGenerationException(
        BatchFatalErrorCode code,
        string userMessage,
        Exception? innerException = null)
        : base(userMessage, innerException)
    {
        Code = code;
        UserMessage = userMessage;
    }

    public BatchFatalErrorCode Code { get; }

    public string UserMessage { get; }
}
