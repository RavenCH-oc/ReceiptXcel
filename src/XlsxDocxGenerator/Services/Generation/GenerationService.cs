using System.IO;
using XlsxDocxGenerator.Models;
using XlsxDocxGenerator.Services.Errors;
using XlsxDocxGenerator.Services.Excel;
using XlsxDocxGenerator.Services.Word;

namespace XlsxDocxGenerator.Services.Generation;

public sealed class GenerationService
{
    private readonly IExcelReader _excelReader;
    private readonly IDocxGenerator _docxGenerator;
    private readonly RowResolver _rowResolver;
    private readonly FilenamePolicy _filenamePolicy;
    private readonly RuntimeWorksheetSchemaValidator _schemaValidator;

    public GenerationService(
        IExcelReader excelReader,
        IDocxGenerator docxGenerator,
        RowResolver? rowResolver = null,
        FilenamePolicy? filenamePolicy = null,
        RuntimeWorksheetSchemaValidator? schemaValidator = null)
    {
        _excelReader = excelReader;
        _docxGenerator = docxGenerator;
        _rowResolver = rowResolver ?? new RowResolver();
        _filenamePolicy = filenamePolicy ?? new FilenamePolicy();
        _schemaValidator = schemaValidator ?? new RuntimeWorksheetSchemaValidator();
    }

    public async Task GenerateSingleAsync(
        string workbookPath,
        TemplateDefinition template,
        int rowNumber,
        string outputPath,
        CancellationToken cancellationToken = default)
    {
        var record = await _excelReader.ReadRecordAsync(
            workbookPath,
            template,
            rowNumber,
            cancellationToken);

        await _docxGenerator.GenerateAsync(template, record, outputPath, cancellationToken);
    }

    public async Task<RowSelectionResult> PreviewSelectionAsync(
        string workbookPath,
        TemplateDefinition template,
        SelectionRule selectionRule,
        CancellationToken cancellationToken = default,
        int? explicitRuntimeMappingMarkerRowNumber = null)
    {
        var worksheet = await _excelReader.ReadAsync(
            workbookPath,
            template.PreferredWorksheetName,
            template.HeaderRowNumber,
            cancellationToken);

        var context = _schemaValidator.Validate(
            worksheet,
            template,
            explicitRuntimeMappingMarkerRowNumber);
        return _rowResolver.Resolve(
            context.Worksheet,
            template,
            selectionRule,
            context.RuntimeMappingMarkerRowNumber);
    }

    public async Task<BatchGenerationResult> GenerateBatchAsync(
        string workbookPath,
        TemplateDefinition template,
        SelectionRule selectionRule,
        string outputDirectory,
        IProgress<BatchProgress>? progress = null,
        CancellationToken cancellationToken = default,
        int? explicitRuntimeMappingMarkerRowNumber = null)
    {
        ArgumentNullException.ThrowIfNull(template);
        ArgumentNullException.ThrowIfNull(selectionRule);

        if (cancellationToken.IsCancellationRequested)
        {
            return CancelledResult(GetKnownRequestedCount(selectionRule));
        }

        try
        {
            EnsureOutputDirectory(outputDirectory);
            await ValidateTemplateForBatchAsync(template, cancellationToken);
        }
        catch (OperationCanceledException)
        {
            return CancelledResult(0);
        }

        ExcelWorksheetData worksheet;
        try
        {
            worksheet = await _excelReader.ReadAsync(
                workbookPath,
                template.PreferredWorksheetName,
                template.HeaderRowNumber,
                cancellationToken);
        }
        catch (OperationCanceledException)
        {
            return CancelledResult(0);
        }
        catch (BatchGenerationException)
        {
            throw;
        }
        catch (Exception exception)
        {
            throw new BatchGenerationException(
                BatchFatalErrorCode.ExcelReadFailed,
                "無法讀取 Excel 資料，請確認檔案未損壞或被其他程式鎖定。",
                exception);
        }

        RuntimeWorksheetContext context;
        try
        {
            context = _schemaValidator.Validate(
                worksheet,
                template,
                explicitRuntimeMappingMarkerRowNumber);
        }
        catch (TemplateDefinitionException exception)
        {
            throw new BatchGenerationException(
                BatchFatalErrorCode.TemplateDefinitionInvalid,
                exception.UserMessage,
                exception);
        }

        RowSelectionResult selection;
        try
        {
            selection = _rowResolver.Resolve(
                context.Worksheet,
                template,
                selectionRule,
                context.RuntimeMappingMarkerRowNumber);
        }
        catch (SelectionValidationException exception)
        {
            throw new BatchGenerationException(
                BatchFatalErrorCode.TemplateDefinitionInvalid,
                exception.UserMessage,
                exception);
        }

        if (selection.Items.Count == 0)
        {
            return new BatchGenerationResult(
                [],
                selection.Notice ?? "沒有符合條件的資料列。",
                BatchGenerationStatus.Completed,
                0);
        }

        var results = new List<RowGenerationResult>(selection.Items.Count);
        var allocator = new FilenameAllocator();
        var total = selection.Items.Count;
        progress?.Report(new BatchProgress(0, total, 0));

        for (var index = 0; index < selection.Items.Count; index++)
        {
            if (cancellationToken.IsCancellationRequested)
            {
                return CancelledResult(results, total);
            }

            var item = selection.Items[index];
            if (!item.IsValid)
            {
                results.Add(RowGenerationResult.Failed(
                    item.RowNumber,
                    item.ErrorCode ?? GenerationErrorCode.InvalidExcelDataRow,
                    item.ErrorMessage ?? "Excel 資料列無效。"));
                progress?.Report(new BatchProgress(index + 1, total, item.RowNumber));
                continue;
            }

            try
            {
                var baseName = _filenamePolicy.ResolveBaseName(
                    template.OutputFileNamePattern,
                    template,
                    item.Record!);
                var outputPath = allocator.Allocate(outputDirectory, baseName);

                await _docxGenerator.GenerateAsync(
                    template,
                    item.Record!,
                    outputPath,
                    cancellationToken);

                results.Add(RowGenerationResult.Succeeded(item.RowNumber, outputPath));
            }
            catch (OperationCanceledException)
            {
                return CancelledResult(results, total);
            }
            catch (GenerationException exception)
            {
                results.Add(RowGenerationResult.Failed(
                    item.RowNumber,
                    exception.Code,
                    exception.UserMessage));
            }
            catch (Exception)
            {
                results.Add(RowGenerationResult.Failed(
                    item.RowNumber,
                    GenerationErrorCode.OutputWriteFailed,
                    "產生 Word 文件時發生錯誤。"));
            }

            progress?.Report(new BatchProgress(index + 1, total, item.RowNumber));
        }

        return new BatchGenerationResult(
            results,
            selection.Notice,
            BatchGenerationStatus.Completed,
            total);
    }

    private static BatchGenerationResult CancelledResult(int requestedCount) =>
        CancelledResult([], requestedCount);

    private static int GetKnownRequestedCount(SelectionRule selectionRule) => selectionRule switch
    {
        SelectionRule.ExcelRows excelRows => excelRows.RowNumbers.Count,
        _ => 0
    };

    private static BatchGenerationResult CancelledResult(
        IReadOnlyList<RowGenerationResult> results,
        int requestedCount) =>
        new(
            results,
            "已取消",
            BatchGenerationStatus.Cancelled,
            requestedCount);

    private async Task ValidateTemplateForBatchAsync(
        TemplateDefinition template,
        CancellationToken cancellationToken)
    {
        try
        {
            await _docxGenerator.ValidateTemplateAsync(template, cancellationToken);
        }
        catch (GenerationException exception)
        {
            var code = exception.Code is GenerationErrorCode.WordTemplateNotFound
                or GenerationErrorCode.InvalidWordTemplate
                ? BatchFatalErrorCode.WordTemplateInvalid
                : BatchFatalErrorCode.TemplateDefinitionInvalid;
            throw new BatchGenerationException(code, exception.UserMessage, exception);
        }
    }

    private static void EnsureOutputDirectory(string outputDirectory)
    {
        try
        {
            ArgumentException.ThrowIfNullOrWhiteSpace(outputDirectory);
            Directory.CreateDirectory(outputDirectory);
        }
        catch (Exception exception)
        {
            throw new BatchGenerationException(
                BatchFatalErrorCode.OutputDirectoryUnavailable,
                "無法使用輸出資料夾，請確認資料夾存在且具有寫入權限。",
                exception);
        }
    }
}

public static class OutputPathPolicy
{
    public static string CreateSingleDocumentPath(string outputDirectory, int rowNumber)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(outputDirectory);
        if (rowNumber <= 0)
        {
            throw new ArgumentOutOfRangeException(nameof(rowNumber));
        }

        return Path.Combine(outputDirectory, $"output-row-{rowNumber}.docx");
    }
}
