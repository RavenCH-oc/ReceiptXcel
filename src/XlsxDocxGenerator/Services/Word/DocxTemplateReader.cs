using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using XlsxDocxGenerator.Services.Errors;

namespace XlsxDocxGenerator.Services.Word;

public sealed class DocxTemplateReader : IDocxTemplateReader
{
    private readonly IPlaceholderResolver _placeholderResolver;

    public DocxTemplateReader(IPlaceholderResolver? placeholderResolver = null)
    {
        _placeholderResolver = placeholderResolver ?? new PlaceholderResolver();
    }

    public Task<DocxTemplateInfo> ReadAsync(
        string templatePath,
        CancellationToken cancellationToken = default)
    {
        var fullPath = ValidatePath(templatePath);
        cancellationToken.ThrowIfCancellationRequested();

        try
        {
            using var document = WordprocessingDocument.Open(fullPath, false);
            if (document.MainDocumentPart?.Document.Body is null)
            {
                throw new InvalidOperationException("Word 文件缺少主要文件內容。");
            }
        }
        catch (Exception exception) when (exception is IOException
            or InvalidDataException
            or OpenXmlPackageException
            or InvalidOperationException)
        {
            throw new GenerationException(
                GenerationErrorCode.InvalidWordTemplate,
                $"無法讀取 Word 模板，檔案可能損壞或格式不受支援：{Path.GetFileName(fullPath)}",
                exception);
        }

        return Task.FromResult(new DocxTemplateInfo(
            Path.GetFileNameWithoutExtension(fullPath),
            fullPath));
    }

    public Task<IReadOnlyList<string>> ScanPlaceholdersAsync(
        string templatePath,
        CancellationToken cancellationToken = default)
    {
        var fullPath = ValidatePath(templatePath);
        cancellationToken.ThrowIfCancellationRequested();

        try
        {
            using var document = WordprocessingDocument.Open(fullPath, false);
            var mainDocumentPart = document.MainDocumentPart;
            var body = mainDocumentPart?.Document.Body;
            if (body is null)
            {
                throw new InvalidOperationException("Word 文件缺少主要文件內容。");
            }

            var names = EnumerateRoots(mainDocumentPart!)
                .SelectMany(root => _placeholderResolver.FindOccurrences(root))
                .Select(occurrence => occurrence.Name)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToArray();

            return Task.FromResult<IReadOnlyList<string>>(names);
        }
        catch (GenerationException)
        {
            throw;
        }
        catch (Exception exception) when (exception is IOException
            or InvalidDataException
            or OpenXmlPackageException
            or InvalidOperationException)
        {
            throw new GenerationException(
                GenerationErrorCode.InvalidWordTemplate,
                $"無法掃描 Word 模板，檔案可能損壞或格式不受支援：{Path.GetFileName(fullPath)}",
                exception);
        }
    }

    private static IEnumerable<OpenXmlElement> EnumerateRoots(MainDocumentPart mainDocumentPart)
    {
        var body = mainDocumentPart.Document.Body;
        if (body is not null)
        {
            yield return body;
        }

        foreach (var part in mainDocumentPart.HeaderParts
            .GroupBy(part => part.Uri.ToString(), StringComparer.OrdinalIgnoreCase)
            .Select(group => group.First())
            .OrderBy(part => part.Uri.ToString(), StringComparer.OrdinalIgnoreCase))
        {
            if (part.Header is not null)
            {
                yield return part.Header;
            }
        }

        foreach (var part in mainDocumentPart.FooterParts
            .GroupBy(part => part.Uri.ToString(), StringComparer.OrdinalIgnoreCase)
            .Select(group => group.First())
            .OrderBy(part => part.Uri.ToString(), StringComparer.OrdinalIgnoreCase))
        {
            if (part.Footer is not null)
            {
                yield return part.Footer;
            }
        }
    }

    private static string ValidatePath(string templatePath)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(templatePath);
        var fullPath = Path.GetFullPath(templatePath);

        if (!File.Exists(fullPath))
        {
            throw new GenerationException(
                GenerationErrorCode.WordTemplateNotFound,
                $"找不到 Word 模板：{fullPath}");
        }

        return fullPath;
    }
}
