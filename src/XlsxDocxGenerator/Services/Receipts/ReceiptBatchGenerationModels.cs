using XlsxDocxGenerator.Models;

namespace XlsxDocxGenerator.Services.Receipts;

public enum ReceiptSelectionMode
{
    Latest,
    ExcelRows
}

public sealed record ReceiptSelectionRequest(
    ReceiptSelectionMode Mode,
    string Parameter);

public sealed record ReceiptRowPreview(
    int ExcelRowNumber,
    ReceiptRecord? Record,
    string? ErrorCode,
    string? UserMessage)
{
    public bool IsValid => Record is not null;

    public string? ReceiptNumber => Record?.ReceiptNumber;

    public static ReceiptRowPreview Success(ReceiptRecord record) =>
        new(record.ExcelRowNumber, record, null, null);

    public static ReceiptRowPreview Failure(
        int rowNumber,
        string errorCode,
        string message) =>
        new(rowNumber, null, errorCode, message);
}

public sealed record ReceiptBatchPreview(
    string ExcelPath,
    string WorksheetName,
    bool SchemaValid,
    int AvailableRowCount,
    IReadOnlyList<ReceiptRowPreview> SelectedRows,
    string? Notice,
    string? BlockingErrorCode,
    string? BlockingUserMessage)
{
    public int TotalSelected => SelectedRows.Count;

    public bool CanGenerate =>
        SchemaValid
        && TotalSelected > 0
        && string.IsNullOrWhiteSpace(BlockingErrorCode);
}

public enum ReceiptBatchStatus
{
    Completed,
    Cancelled
}

public sealed record ReceiptBatchProgress(
    int Completed,
    int Total,
    int ExcelRowNumber);

public sealed record ReceiptRowGenerationResult(
    int ExcelRowNumber,
    string? ReceiptNumber,
    bool Success,
    string? OutputPath,
    string? ErrorCode,
    string? UserMessage)
{
    public static ReceiptRowGenerationResult Succeeded(
        ReceiptRecord record,
        string outputPath) =>
        new(record.ExcelRowNumber, record.ReceiptNumber, true, outputPath, null, null);

    public static ReceiptRowGenerationResult Failed(
        int rowNumber,
        string? receiptNumber,
        string errorCode,
        string message) =>
        new(rowNumber, receiptNumber, false, null, errorCode, message);
}

public sealed record ReceiptBatchGenerationResult(
    ReceiptBatchStatus Status,
    int TotalSelected,
    IReadOnlyList<ReceiptRowGenerationResult> Results,
    string? Notice = null)
{
    public bool IsCancelled => Status == ReceiptBatchStatus.Cancelled;

    public int SuccessCount => Results.Count(result => result.Success);

    public int FailureCount => Results.Count(result => !result.Success);

    public int UnprocessedCount => Math.Max(0, TotalSelected - Results.Count);
}

public enum ReceiptBatchFatalErrorCode
{
    ExcelReadFailed,
    WorksheetMissing,
    ExcelSchemaMismatch,
    InvalidSelection,
    DuplicateReceiptNumber,
    OutputDirectoryUnavailable,
    InternalTemplateInvalid
}

public sealed class ReceiptBatchGenerationException : Exception
{
    public ReceiptBatchGenerationException(
        ReceiptBatchFatalErrorCode code,
        string userMessage,
        Exception? innerException = null)
        : base(userMessage, innerException)
    {
        Code = code;
        UserMessage = userMessage;
    }

    public ReceiptBatchFatalErrorCode Code { get; }

    public string UserMessage { get; }
}
