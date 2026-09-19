using System.IO;
using DocumentFormat.OpenXml.Packaging;
using XlsxDocxGenerator.Models;
using XlsxDocxGenerator.Services.Word;

namespace XlsxDocxGenerator.Services.Receipts;

/// <summary>
/// Fixed single-row Excel-to-receipt generation boundary. The service owns
/// validation, derived values, placeholder replacement, and safe file output;
/// callers never need to assemble production mappings themselves.
/// </summary>
public sealed class ReceiptGenerationService
{
    private readonly ReceiptRecordReader _recordReader;
    private readonly ReceiptAmountFormatter _amountFormatter;
    private readonly ReceiptPlaceholderValuesBuilder _valuesBuilder;
    private readonly ReceiptTemplateContractValidator _templateValidator;
    private readonly ReceiptGeneratedDocumentValidator _generatedValidator;
    private readonly ReceiptFilenamePolicy _filenamePolicy;
    private readonly IPlaceholderResolver _placeholderResolver;
    private readonly IReceiptTemplateProvider _templateProvider;

    public ReceiptGenerationService(
        ReceiptRecordReader? recordReader = null,
        ReceiptAmountFormatter? amountFormatter = null,
        ReceiptPlaceholderValuesBuilder? valuesBuilder = null,
        ReceiptTemplateContractValidator? templateValidator = null,
        ReceiptGeneratedDocumentValidator? generatedValidator = null,
        ReceiptFilenamePolicy? filenamePolicy = null,
        IPlaceholderResolver? placeholderResolver = null,
        string? internalTemplatePath = null,
        IReceiptTemplateProvider? templateProvider = null)
    {
        _recordReader = recordReader ?? new ReceiptRecordReader();
        _amountFormatter = amountFormatter ?? new ReceiptAmountFormatter();
        _valuesBuilder = valuesBuilder ?? new ReceiptPlaceholderValuesBuilder();
        _placeholderResolver = placeholderResolver ?? new PlaceholderResolver();
        _templateValidator = templateValidator
            ?? new ReceiptTemplateContractValidator(_placeholderResolver);
        _generatedValidator = generatedValidator ?? new ReceiptGeneratedDocumentValidator();
        _filenamePolicy = filenamePolicy ?? new ReceiptFilenamePolicy();
        if (internalTemplatePath is not null && templateProvider is not null)
        {
            throw new ArgumentException("Specify either a template provider or an explicit template path.");
        }
        _templateProvider = templateProvider ?? (internalTemplatePath is null
            ? new EmbeddedReceiptTemplateProvider()
            : new FileReceiptTemplateProvider(internalTemplatePath));
    }

    /// <summary>
    /// Generates exactly one DOCX at the requested path. Existing files are
    /// never overwritten.
    /// </summary>
    public async Task<string> GenerateReceiptAsync(
        string excelPath,
        string worksheetName,
        int rowNumber,
        string outputPath,
        CancellationToken cancellationToken = default)
    {
        var record = await _recordReader.ReadRecordAsync(
            excelPath,
            worksheetName,
            rowNumber,
            cancellationToken);
        var amount = _amountFormatter.Format(record.Amount);
        return await GenerateRecordAsync(record, amount, outputPath, cancellationToken);
    }

    /// <summary>
    /// Convenience API for the fixed Phase 1 filename policy:
    /// 收據_{ReceiptNumber}.docx.
    /// </summary>
    public async Task<string> GenerateReceiptToDirectoryAsync(
        string excelPath,
        string worksheetName,
        int rowNumber,
        string outputDirectory,
        CancellationToken cancellationToken = default)
    {
        var record = await _recordReader.ReadRecordAsync(
            excelPath,
            worksheetName,
            rowNumber,
            cancellationToken);
        var amount = _amountFormatter.Format(record.Amount);
        var outputPath = _filenamePolicy.GetOutputPath(outputDirectory, record);
        return await GenerateRecordAsync(record, amount, outputPath, cancellationToken);
    }

    /// <summary>
    /// Generates a validated receipt from an already-parsed record. Batch
    /// orchestration uses this overload so the Excel workbook is read once
    /// and row-level validation can continue independently.
    /// </summary>
    public Task<string> GenerateReceiptAsync(
        ReceiptRecord record,
        string outputPath,
        CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(record);
        var amount = _amountFormatter.Format(record.Amount);
        return GenerateRecordAsync(record, amount, outputPath, cancellationToken);
    }

    public Task<string> GenerateReceiptToDirectoryAsync(
        ReceiptRecord record,
        string outputDirectory,
        CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(record);
        var outputPath = _filenamePolicy.GetOutputPath(outputDirectory, record);
        return GenerateReceiptAsync(record, outputPath, cancellationToken);
    }

    public ReceiptTemplateContract ValidateInternalTemplate()
    {
        using var template = _templateProvider.Acquire();
        return _templateValidator.Validate(template.TemplatePath);
    }

    internal ReceiptTemplateLease AcquireValidatedTemplate(CancellationToken cancellationToken)
    {
        var template = _templateProvider.Acquire(cancellationToken);
        try
        {
            _templateValidator.Validate(template.TemplatePath);
            return template;
        }
        catch
        {
            template.Dispose();
            throw;
        }
    }

    internal Task<string> GenerateReceiptAsync(
        ReceiptRecord record,
        string outputPath,
        ReceiptTemplateLease template,
        CancellationToken cancellationToken) =>
        GenerateRecordAsync(record, _amountFormatter.Format(record.Amount),
            outputPath, cancellationToken, template.TemplatePath);

    private async Task<string> GenerateRecordAsync(
        ReceiptRecord record,
        FormattedReceiptAmount amount,
        string outputPath,
        CancellationToken cancellationToken)
    {
        using var template = AcquireValidatedTemplate(cancellationToken);
        return await GenerateRecordAsync(record, amount, outputPath, cancellationToken, template.TemplatePath);
    }

    private Task<string> GenerateRecordAsync(
        ReceiptRecord record,
        FormattedReceiptAmount amount,
        string outputPath,
        CancellationToken cancellationToken,
        string templatePath)
    {
        cancellationToken.ThrowIfCancellationRequested();

        var finalPath = GetFullPath(outputPath);
        if (string.Equals(finalPath, templatePath, StringComparison.OrdinalIgnoreCase))
        {
            throw new ReceiptGenerationException(
                ReceiptGenerationErrorCode.OutputWriteFailed,
                "輸出路徑不可與內建收據模板相同。");
        }

        if (File.Exists(finalPath))
        {
            throw new ReceiptGenerationException(
                ReceiptGenerationErrorCode.OutputAlreadyExists,
                $"輸出檔案已存在，為避免覆蓋既有文件而停止：{finalPath}");
        }

        var outputDirectory = Path.GetDirectoryName(finalPath);
        if (string.IsNullOrWhiteSpace(outputDirectory))
        {
            throw new ReceiptGenerationException(
                ReceiptGenerationErrorCode.InvalidOutputPath,
                "無法判斷收據輸出資料夾。");
        }

        var temporaryPath = Path.Combine(
            outputDirectory,
            $".{Path.GetFileName(finalPath)}.{Guid.NewGuid():N}.tmp");

        try
        {
            Directory.CreateDirectory(outputDirectory);
            cancellationToken.ThrowIfCancellationRequested();
            File.Copy(templatePath, temporaryPath, overwrite: false);

            using (var document = WordprocessingDocument.Open(temporaryPath, true))
            {
                var mainDocumentPart = document.MainDocumentPart;
                var body = mainDocumentPart?.Document.Body;
                if (mainDocumentPart is null || body is null)
                {
                    throw new ReceiptGenerationException(
                        ReceiptGenerationErrorCode.ReceiptTemplateInvalid,
                        "內建收據模板缺少主要文件內容。");
                }

                var values = _valuesBuilder.Build(record, amount);
                _placeholderResolver.ReplaceOccurrences(body, values);
                mainDocumentPart.Document.Save();

                foreach (var part in ReceiptOpenXmlParts.EnumerateHeaders(mainDocumentPart))
                {
                    if (part.Header is not null)
                    {
                        _placeholderResolver.ReplaceOccurrences(part.Header, values);
                        part.Header.Save();
                    }
                }

                foreach (var part in ReceiptOpenXmlParts.EnumerateFooters(mainDocumentPart))
                {
                    if (part.Footer is not null)
                    {
                        _placeholderResolver.ReplaceOccurrences(part.Footer, values);
                        part.Footer.Save();
                    }
                }
            }

            _generatedValidator.Validate(temporaryPath, record, amount);
            cancellationToken.ThrowIfCancellationRequested();
            File.Move(temporaryPath, finalPath, overwrite: false);
            return Task.FromResult(finalPath);
        }
        catch (OperationCanceledException)
        {
            DeleteTemporaryFile(temporaryPath);
            throw;
        }
        catch (ReceiptGenerationException)
        {
            DeleteTemporaryFile(temporaryPath);
            throw;
        }
        catch (Exception exception)
        {
            DeleteTemporaryFile(temporaryPath);
            throw new ReceiptGenerationException(
                ReceiptGenerationErrorCode.OutputWriteFailed,
                $"產生收據 DOCX 失敗：{Path.GetFileName(finalPath)}",
                exception);
        }
    }

    private static string GetFullPath(string outputPath)
    {
        try
        {
            ArgumentException.ThrowIfNullOrWhiteSpace(outputPath);
            var fullPath = Path.GetFullPath(outputPath);
            if (!string.Equals(
                    Path.GetExtension(fullPath),
                    ".docx",
                    StringComparison.OrdinalIgnoreCase))
            {
                throw new ReceiptGenerationException(
                    ReceiptGenerationErrorCode.InvalidOutputPath,
                    "收據輸出檔案必須使用 .docx 副檔名。");
            }

            return fullPath;
        }
        catch (ReceiptGenerationException)
        {
            throw;
        }
        catch (Exception exception) when (exception is ArgumentException
            or NotSupportedException
            or PathTooLongException)
        {
            throw new ReceiptGenerationException(
                ReceiptGenerationErrorCode.InvalidOutputPath,
                "收據輸出路徑無效。",
                exception);
        }
    }

    private static void DeleteTemporaryFile(string temporaryPath)
    {
        try
        {
            if (File.Exists(temporaryPath))
            {
                File.Delete(temporaryPath);
            }
        }
        catch
        {
            // Preserve the original generation failure.
        }
    }
}
