using System.IO;
using XlsxDocxGenerator.Services.Errors;
using XlsxDocxGenerator.Services.Excel;

namespace XlsxDocxGenerator.Services.Receipts;

/// <summary>
/// Specialized application workflow for fixed receipts. It owns selection,
/// preflight, duplicate detection, progress, cancellation, and per-row
/// continuation. The UI calls this service once per batch.
/// </summary>
public sealed class ReceiptBatchGenerationService
{
    private readonly ReceiptRecordReader _recordReader;
    private readonly ReceiptGenerationService _receiptGenerationService;
    private readonly ReceiptFilenamePolicy _filenamePolicy;
    private readonly SelectionRuleParser _selectionRuleParser;

    public ReceiptBatchGenerationService(
        ReceiptRecordReader? recordReader = null,
        ReceiptGenerationService? receiptGenerationService = null,
        ReceiptFilenamePolicy? filenamePolicy = null,
        SelectionRuleParser? selectionRuleParser = null)
    {
        _recordReader = recordReader ?? new ReceiptRecordReader();
        _receiptGenerationService = receiptGenerationService ?? new ReceiptGenerationService();
        _filenamePolicy = filenamePolicy ?? new ReceiptFilenamePolicy();
        _selectionRuleParser = selectionRuleParser ?? new SelectionRuleParser();
    }

    public string InternalTemplatePath => _receiptGenerationService.InternalTemplatePath;

    public bool IsInternalTemplateAvailable(out string userMessage)
    {
        try
        {
            _receiptGenerationService.ValidateInternalTemplate();
            userMessage = "內建收據模板可用。";
            return true;
        }
        catch (ReceiptGenerationException)
        {
            userMessage = "內建收據模板無法使用，請重新安裝 ReceiptXcel。";
            return false;
        }
    }

    public async Task<ReceiptBatchPreview> PreviewAsync(
        string excelPath,
        string worksheetName,
        ReceiptSelectionRequest request,
        CancellationToken cancellationToken = default)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(excelPath);
        ArgumentException.ThrowIfNullOrWhiteSpace(worksheetName);
        ArgumentNullException.ThrowIfNull(request);

        ExcelWorksheetData worksheet;
        try
        {
            worksheet = await _recordReader.ReadWorksheetAsync(
                excelPath,
                worksheetName,
                cancellationToken);
        }
        catch (OperationCanceledException)
        {
            throw;
        }
        catch (ReceiptValidationException exception)
        {
            return InvalidPreview(
                excelPath,
                worksheetName,
                "ExcelSchemaMismatch",
                exception.UserMessage);
        }
        catch (TemplateDefinitionException exception)
        {
            var code = exception.Code == TemplateDefinitionErrorCode.WorksheetNotFound
                ? ReceiptBatchFatalErrorCode.WorksheetMissing.ToString()
                : ReceiptBatchFatalErrorCode.ExcelReadFailed.ToString();
            return InvalidPreview(excelPath, worksheetName, code, exception.UserMessage);
        }
        catch (Exception)
        {
            return InvalidPreview(
                excelPath,
                worksheetName,
                ReceiptBatchFatalErrorCode.ExcelReadFailed.ToString(),
                "無法讀取 Excel 檔案，請確認檔案未損壞或被其他程式鎖定。");
        }

        IReadOnlyList<int> selectedRowNumbers;
        string? notice = null;
        try
        {
            selectedRowNumbers = ResolveSelectedRows(worksheet, request, out notice);
        }
        catch (SelectionValidationException exception)
        {
            return new ReceiptBatchPreview(
                excelPath,
                worksheetName,
                true,
                AvailableRowCount(worksheet),
                [],
                null,
                "InvalidSelection",
                exception.UserMessage);
        }

        var rowByNumber = worksheet.Rows.ToDictionary(row => row.RowNumber);
        var selectedRows = selectedRowNumbers
            .Select(rowNumber => ParsePreviewRow(rowNumber, rowByNumber))
            .ToArray();
        var duplicateMessage = FindDuplicateReceiptNumbers(selectedRows);

        return new ReceiptBatchPreview(
            excelPath,
            worksheetName,
            true,
            AvailableRowCount(worksheet),
            selectedRows,
            notice,
            duplicateMessage is null ? null : ReceiptBatchFatalErrorCode.DuplicateReceiptNumber.ToString(),
            duplicateMessage);
    }

    public async Task<ReceiptBatchGenerationResult> GenerateBatchAsync(
        string excelPath,
        string worksheetName,
        ReceiptSelectionRequest request,
        string outputDirectory,
        IProgress<ReceiptBatchProgress>? progress = null,
        CancellationToken cancellationToken = default)
    {
        if (cancellationToken.IsCancellationRequested)
        {
            return new ReceiptBatchGenerationResult(
                ReceiptBatchStatus.Cancelled,
                0,
                [],
                "已取消");
        }

        var preview = await PreviewAsync(excelPath, worksheetName, request, cancellationToken);
        if (!preview.SchemaValid || !string.IsNullOrWhiteSpace(preview.BlockingErrorCode))
        {
            throw new ReceiptBatchGenerationException(
                Enum.TryParse<ReceiptBatchFatalErrorCode>(preview.BlockingErrorCode, out var code)
                    ? code
                    : ReceiptBatchFatalErrorCode.ExcelSchemaMismatch,
                preview.BlockingUserMessage ?? "Excel 資料無法通過固定收據格式驗證。");
        }

        if (preview.TotalSelected == 0)
        {
            return new ReceiptBatchGenerationResult(
                ReceiptBatchStatus.Completed,
                0,
                [],
                preview.Notice ?? "沒有符合條件的資料列。");
        }

        try
        {
            _receiptGenerationService.ValidateInternalTemplate();
        }
        catch (ReceiptGenerationException exception)
        {
            throw new ReceiptBatchGenerationException(
                ReceiptBatchFatalErrorCode.InternalTemplateInvalid,
                "內建收據模板無法使用，請重新安裝 ReceiptXcel。",
                exception);
        }

        var fullOutputDirectory = EnsureOutputDirectory(outputDirectory);
        var results = new List<ReceiptRowGenerationResult>(preview.TotalSelected);
        progress?.Report(new ReceiptBatchProgress(0, preview.TotalSelected, 0));

        for (var index = 0; index < preview.SelectedRows.Count; index++)
        {
            var row = preview.SelectedRows[index];
            if (cancellationToken.IsCancellationRequested)
            {
                return Cancelled(results, preview.TotalSelected, preview.Notice);
            }

            if (!row.IsValid)
            {
                results.Add(ReceiptRowGenerationResult.Failed(
                    row.ExcelRowNumber,
                    row.ReceiptNumber,
                    row.ErrorCode ?? ReceiptValidationErrorCode.InvalidRecord.ToString(),
                    row.UserMessage ?? "Excel 資料列無效。"));
                progress?.Report(new ReceiptBatchProgress(index + 1, preview.TotalSelected, row.ExcelRowNumber));
                continue;
            }

            var record = row.Record!;
            var outputPath = _filenamePolicy.GetOutputPath(fullOutputDirectory, record);
            try
            {
                var generatedPath = await _receiptGenerationService.GenerateReceiptAsync(
                    record,
                    outputPath,
                    cancellationToken);
                results.Add(ReceiptRowGenerationResult.Succeeded(record, generatedPath));
            }
            catch (OperationCanceledException)
            {
                return Cancelled(results, preview.TotalSelected, preview.Notice);
            }
            catch (ReceiptGenerationException exception)
            {
                var message = exception.Code == ReceiptGenerationErrorCode.OutputAlreadyExists
                    ? $"Row {record.ExcelRowNumber}：{Path.GetFileName(outputPath)} 已存在，未覆寫。"
                    : exception.UserMessage;
                results.Add(ReceiptRowGenerationResult.Failed(
                    record.ExcelRowNumber,
                    record.ReceiptNumber,
                    exception.Code.ToString(),
                    message));
            }
            catch (Exception)
            {
                results.Add(ReceiptRowGenerationResult.Failed(
                    record.ExcelRowNumber,
                    record.ReceiptNumber,
                    ReceiptGenerationErrorCode.OutputWriteFailed.ToString(),
                    "產生收據時發生錯誤。"));
            }

            progress?.Report(new ReceiptBatchProgress(index + 1, preview.TotalSelected, record.ExcelRowNumber));
        }

        return new ReceiptBatchGenerationResult(
            ReceiptBatchStatus.Completed,
            preview.TotalSelected,
            results,
            preview.Notice);
    }

    private IReadOnlyList<int> ResolveSelectedRows(
        ExcelWorksheetData worksheet,
        ReceiptSelectionRequest request,
        out string? notice)
    {
        notice = null;
        var availableRows = worksheet.Rows
            .Where(row => row.RowNumber > ReceiptWorksheetSchema.HeaderRowNumber)
            .OrderBy(row => row.RowNumber)
            .Select(row => row.RowNumber)
            .ToArray();

        return request.Mode switch
        {
            ReceiptSelectionMode.Latest => ResolveLatest(availableRows, request.Parameter, out notice),
            ReceiptSelectionMode.ExcelRows => ResolveExcelRows(request.Parameter),
            _ => throw new SelectionValidationException(
                SelectionErrorCode.InvalidSyntax,
                "不支援的資料選擇方式。")
        };
    }

    private IReadOnlyList<int> ResolveLatest(
        IReadOnlyList<int> availableRows,
        string parameter,
        out string? notice)
    {
        var rule = _selectionRuleParser.ParseLatest(parameter);
        notice = rule.Count > availableRows.Count
            ? $"要求 {rule.Count} 筆，實際找到 {availableRows.Count} 筆。"
            : null;
        return availableRows.TakeLast(rule.Count).ToArray();
    }

    private IReadOnlyList<int> ResolveExcelRows(string parameter) =>
        _selectionRuleParser.ParseExcelRows(parameter).RowNumbers;

    private ReceiptRowPreview ParsePreviewRow(
        int rowNumber,
        IReadOnlyDictionary<int, ExcelRowData> rows)
    {
        if (rowNumber <= ReceiptWorksheetSchema.HeaderRowNumber)
        {
            var detail = rowNumber == ReceiptWorksheetSchema.HeaderRowNumber
                ? $"第 {rowNumber} 列是欄位標題，不是收據資料。"
                : $"第 {rowNumber} 列是標題內容，不是收據資料。";
            return ReceiptRowPreview.Failure(
                rowNumber,
                ReceiptValidationErrorCode.RecordNotFound.ToString(),
                detail);
        }

        if (!rows.TryGetValue(rowNumber, out var row))
        {
            return ReceiptRowPreview.Failure(
                rowNumber,
                ReceiptValidationErrorCode.RecordNotFound.ToString(),
                $"第 {rowNumber} 列不存在或是空白資料列。" );
        }

        try
        {
            return ReceiptRowPreview.Success(_recordReader.ParseRecord(row));
        }
        catch (ReceiptValidationException exception)
        {
            return ReceiptRowPreview.Failure(
                rowNumber,
                exception.Code.ToString(),
                exception.UserMessage);
        }
    }

    private static string? FindDuplicateReceiptNumbers(
        IReadOnlyList<ReceiptRowPreview> rows)
    {
        var duplicateGroups = rows
            .Where(row => row.Record is not null)
            .GroupBy(row => row.Record!.ReceiptNumber, StringComparer.OrdinalIgnoreCase)
            .Where(group => group.Count() > 1)
            .OrderBy(group => group.Key, StringComparer.OrdinalIgnoreCase)
            .ToArray();

        if (duplicateGroups.Length == 0)
        {
            return null;
        }

        var details = duplicateGroups.Select(group =>
            $"發現重複收據編號 {group.Key}：{string.Join("、", group.Select(item => $"Row {item.ExcelRowNumber}"))}");
        return string.Join("\n", details) + "\n請確認 Excel 資料。";
    }

    private static int AvailableRowCount(ExcelWorksheetData worksheet) =>
        worksheet.Rows.Count(row => row.RowNumber > ReceiptWorksheetSchema.HeaderRowNumber);

    private static ReceiptBatchPreview InvalidPreview(
        string excelPath,
        string worksheetName,
        string errorCode,
        string message) =>
        new(excelPath, worksheetName, false, 0, [], null, errorCode, message);

    private static string EnsureOutputDirectory(string outputDirectory)
    {
        try
        {
            ArgumentException.ThrowIfNullOrWhiteSpace(outputDirectory);
            var fullPath = Path.GetFullPath(outputDirectory);
            Directory.CreateDirectory(fullPath);
            return fullPath;
        }
        catch (Exception exception) when (exception is ArgumentException
            or IOException
            or UnauthorizedAccessException
            or NotSupportedException
            or PathTooLongException)
        {
            throw new ReceiptBatchGenerationException(
                ReceiptBatchFatalErrorCode.OutputDirectoryUnavailable,
                "無法使用輸出資料夾，請確認資料夾存在且具有寫入權限。",
                exception);
        }
    }

    private static ReceiptBatchGenerationResult Cancelled(
        IReadOnlyList<ReceiptRowGenerationResult> results,
        int total,
        string? notice) =>
        new(ReceiptBatchStatus.Cancelled, total, results.ToArray(), notice ?? "已取消");
}
