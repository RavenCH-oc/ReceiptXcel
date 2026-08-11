using XlsxDocxGenerator.Services.Errors;

namespace XlsxDocxGenerator.Services.Generation;

public sealed record BatchProgress(int Completed, int Total, int RowNumber);

public enum BatchGenerationStatus
{
    Completed,
    Cancelled
}

public sealed record RowGenerationResult(
    int ExcelRowNumber,
    bool Success,
    string? OutputPath,
    GenerationErrorCode? ErrorCode,
    string? Message)
{
    public static RowGenerationResult Succeeded(int rowNumber, string outputPath) =>
        new(rowNumber, true, outputPath, null, null);

    public static RowGenerationResult Failed(
        int rowNumber,
        GenerationErrorCode errorCode,
        string message) =>
        new(rowNumber, false, null, errorCode, message);
}

public sealed record BatchGenerationResult(
    IReadOnlyList<RowGenerationResult> Results,
    string? Notice,
    BatchGenerationStatus Status = BatchGenerationStatus.Completed,
    int? RequestedCount = null)
{
    public bool IsCancelled => Status == BatchGenerationStatus.Cancelled;

    public int TotalSelected => RequestedCount ?? Results.Count;

    public int SuccessCount => Results.Count(result => result.Success);

    public int FailureCount => Results.Count(result => !result.Success);

    public int UnprocessedCount => Math.Max(0, TotalSelected - Results.Count);
}
