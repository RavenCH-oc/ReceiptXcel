using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using XlsxDocxGenerator.Models;
using XlsxDocxGenerator.Services.Errors;

namespace XlsxDocxGenerator.Services.Word;

public interface IDocxGenerator
{
    Task ValidateTemplateAsync(
        TemplateDefinition template,
        CancellationToken cancellationToken = default);

    Task GenerateAsync(
        TemplateDefinition template,
        ExcelRecord record,
        string outputPath,
        CancellationToken cancellationToken = default);
}

public sealed class DocxGenerator : IDocxGenerator
{
    private readonly IDocxTemplateReader _templateReader;
    private readonly IPlaceholderResolver _placeholderResolver;
    private readonly TemplateDefinitionValidator _validator;

    public DocxGenerator(
        IDocxTemplateReader? templateReader = null,
        IPlaceholderResolver? placeholderResolver = null,
        TemplateDefinitionValidator? validator = null)
    {
        _placeholderResolver = placeholderResolver ?? new PlaceholderResolver();
        _templateReader = templateReader ?? new DocxTemplateReader(_placeholderResolver);
        _validator = validator ?? new TemplateDefinitionValidator();
    }

    public async Task ValidateTemplateAsync(
        TemplateDefinition template,
        CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(template);
        var templatePath = GetFullPath(template.WordTemplatePath, GenerationErrorCode.WordTemplateNotFound);

        if (!File.Exists(templatePath))
        {
            throw new GenerationException(
                GenerationErrorCode.WordTemplateNotFound,
                $"找不到 Word 模板：{templatePath}");
        }

        try
        {
            var placeholders = await _templateReader.ScanPlaceholdersAsync(templatePath, cancellationToken);
            _validator.ValidatePlaceholders(template, placeholders);
        }
        catch (GenerationException)
        {
            throw;
        }
        catch (OperationCanceledException)
        {
            throw;
        }
        catch (Exception exception)
        {
            throw new GenerationException(
                GenerationErrorCode.InvalidWordTemplate,
                "無法驗證 Word 模板內容。",
                exception);
        }
    }

    public async Task GenerateAsync(
        TemplateDefinition template,
        ExcelRecord record,
        string outputPath,
        CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(template);
        ArgumentNullException.ThrowIfNull(record);

        var templatePath = GetFullPath(template.WordTemplatePath, GenerationErrorCode.WordTemplateNotFound);
        var finalPath = GetFullPath(outputPath, GenerationErrorCode.OutputWriteFailed);

        if (!File.Exists(templatePath))
        {
            throw new GenerationException(
                GenerationErrorCode.WordTemplateNotFound,
                $"找不到 Word 模板：{templatePath}");
        }

        if (string.Equals(templatePath, finalPath, StringComparison.OrdinalIgnoreCase))
        {
            throw new GenerationException(
                GenerationErrorCode.OutputWriteFailed,
                "輸出路徑不可與原始 Word 模板相同。");
        }

        if (File.Exists(finalPath))
        {
            throw new GenerationException(
                GenerationErrorCode.OutputAlreadyExists,
                $"輸出檔案已存在，為避免覆蓋既有文件而停止：{finalPath}");
        }

        await ValidateTemplateAsync(template, cancellationToken);

        var outputDirectory = Path.GetDirectoryName(finalPath);
        if (string.IsNullOrWhiteSpace(outputDirectory))
        {
            throw new GenerationException(
                GenerationErrorCode.OutputWriteFailed,
                "無法判斷輸出資料夾。");
        }

        var temporaryPath = Path.Combine(
            outputDirectory,
            $".{Path.GetFileName(finalPath)}.{Guid.NewGuid():N}.tmp");

        try
        {
            cancellationToken.ThrowIfCancellationRequested();
            Directory.CreateDirectory(outputDirectory);
            File.Copy(templatePath, temporaryPath, overwrite: false);

            using (var document = WordprocessingDocument.Open(temporaryPath, true))
            {
                var mainDocumentPart = document.MainDocumentPart;
                var body = mainDocumentPart?.Document.Body;
                if (mainDocumentPart is null || body is null)
                {
                    throw new GenerationException(
                        GenerationErrorCode.InvalidWordTemplate,
                        "Word 模板缺少主要文件內容。");
                }

                var values = template.FieldMappings
                    .ToDictionary(
                        mapping => mapping.PlaceholderName.ToUpperInvariant(),
                        mapping => record.GetValueByColumnIndex(mapping.ColumnIndex),
                        StringComparer.OrdinalIgnoreCase);

                _placeholderResolver.ReplaceOccurrences(body, values);
                mainDocumentPart.Document.Save();

                foreach (var part in EnumerateHeaderParts(mainDocumentPart))
                {
                    if (part.Header is not null)
                    {
                        _placeholderResolver.ReplaceOccurrences(part.Header, values);
                        part.Header.Save();
                    }
                }

                foreach (var part in EnumerateFooterParts(mainDocumentPart))
                {
                    if (part.Footer is not null)
                    {
                        _placeholderResolver.ReplaceOccurrences(part.Footer, values);
                        part.Footer.Save();
                    }
                }
            }

            cancellationToken.ThrowIfCancellationRequested();
            File.Move(temporaryPath, finalPath, overwrite: false);
        }
        catch (OperationCanceledException)
        {
            DeleteTemporaryFile(temporaryPath);
            throw;
        }
        catch (GenerationException)
        {
            DeleteTemporaryFile(temporaryPath);
            throw;
        }
        catch (Exception exception)
        {
            DeleteTemporaryFile(temporaryPath);
            throw new GenerationException(
                GenerationErrorCode.OutputWriteFailed,
                $"產生 Word 文件失敗：{Path.GetFileName(finalPath)}",
                exception);
        }
    }

    private static string GetFullPath(string path, GenerationErrorCode errorCode)
    {
        try
        {
            ArgumentException.ThrowIfNullOrWhiteSpace(path);
            return Path.GetFullPath(path);
        }
        catch (Exception exception) when (exception is ArgumentException or NotSupportedException or PathTooLongException)
        {
            throw new GenerationException(
                errorCode,
                "檔案路徑無效。",
                exception);
        }
    }

    private static IEnumerable<HeaderPart> EnumerateHeaderParts(MainDocumentPart mainDocumentPart) =>
        mainDocumentPart.HeaderParts
            .GroupBy(part => part.Uri.ToString(), StringComparer.OrdinalIgnoreCase)
            .Select(group => group.First())
            .OrderBy(part => part.Uri.ToString(), StringComparer.OrdinalIgnoreCase);

    private static IEnumerable<FooterPart> EnumerateFooterParts(MainDocumentPart mainDocumentPart) =>
        mainDocumentPart.FooterParts
            .GroupBy(part => part.Uri.ToString(), StringComparer.OrdinalIgnoreCase)
            .Select(group => group.First())
            .OrderBy(part => part.Uri.ToString(), StringComparer.OrdinalIgnoreCase);

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
            // The original generation error is more useful than cleanup failure.
        }
    }
}
