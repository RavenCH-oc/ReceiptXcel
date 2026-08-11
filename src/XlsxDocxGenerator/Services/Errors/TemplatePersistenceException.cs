namespace XlsxDocxGenerator.Services.Errors;

public enum TemplatePersistenceErrorCode
{
    TemplateFileNotFound,
    InvalidTemplateFile,
    UnsupportedTemplateVersion,
    MissingRequiredField,
    InvalidFieldMapping
}

public sealed class TemplatePersistenceException : Exception
{
    public TemplatePersistenceException(
        TemplatePersistenceErrorCode code,
        string userMessage,
        Exception? innerException = null)
        : base(userMessage, innerException)
    {
        Code = code;
        UserMessage = userMessage;
    }

    public TemplatePersistenceErrorCode Code { get; }

    public string UserMessage { get; }
}
