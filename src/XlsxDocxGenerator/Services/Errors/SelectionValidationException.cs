namespace XlsxDocxGenerator.Services.Errors;

public enum SelectionErrorCode
{
    InvalidSyntax,
    MissingMarkerColumn,
    MissingMarkerToken,
    InvalidLatestCount
}

public sealed class SelectionValidationException : Exception
{
    public SelectionValidationException(
        SelectionErrorCode code,
        string userMessage)
        : base(userMessage)
    {
        Code = code;
        UserMessage = userMessage;
    }

    public SelectionErrorCode Code { get; }

    public string UserMessage { get; }
}
