namespace XlsxDocxGenerator.Services.Errors;

public enum TemplateDefinitionErrorCode
{
    ExcelFileNotFound,
    InvalidExcelFile,
    WorksheetNotFound,
    NoMarkerRow,
    AmbiguousMarkerRows,
    InvalidMarkerSyntax,
    DuplicatePlaceholder,
    InvalidHeaderRow,
    InvalidMarkerRow,
    MarkerRowHeaderRelationshipInvalid,
    ExcelSchemaMismatch,
    AmbiguousRuntimeMarkerRows,
    InvalidRuntimeMappingMarker,
    UnsupportedTemplateVersion,
    TemplateFileNotFound,
    InvalidTemplateFile,
    MissingRequiredField,
    InvalidFieldMapping
}

/// <summary>
/// User-facing error for Excel mapping and template definition creation.
/// The exception message is intentionally understandable without exposing library details.
/// </summary>
public sealed class TemplateDefinitionException : Exception
{
    public TemplateDefinitionException(
        TemplateDefinitionErrorCode code,
        string userMessage,
        Exception? innerException = null)
        : base(userMessage, innerException)
    {
        Code = code;
        UserMessage = userMessage;
    }

    public TemplateDefinitionErrorCode Code { get; }

    public string UserMessage { get; }
}
