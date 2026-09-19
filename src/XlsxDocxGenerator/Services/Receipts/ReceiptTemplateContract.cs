using System.IO;
using DocumentFormat.OpenXml.Packaging;
using XlsxDocxGenerator.Services.Word;

namespace XlsxDocxGenerator.Services.Receipts;

public sealed record ReceiptTemplateContract(
    string TemplatePath,
    IReadOnlyDictionary<string, int> OccurrenceCounts);

/// <summary>
/// Validates the embedded production template before generation. A malformed
/// template fails closed instead of producing a partial receipt.
/// </summary>
public sealed class ReceiptTemplateContractValidator
{
    public const int ExpectedOccurrenceCountPerReceipt = 3;

    private readonly IPlaceholderResolver _placeholderResolver;

    public ReceiptTemplateContractValidator(
        IPlaceholderResolver? placeholderResolver = null)
    {
        _placeholderResolver = placeholderResolver ?? new PlaceholderResolver();
    }

    public ReceiptTemplateContract Validate(string templatePath)
    {
        var fullPath = GetFullPath(templatePath);
        var expectedNames = ReceiptTemplateMapping.Fields
            .Select(field => field.Placeholder.ToUpperInvariant())
            .ToArray();

        if (expectedNames.Length != expectedNames.Distinct(StringComparer.OrdinalIgnoreCase).Count())
        {
            throw InvalidTemplate("程式內建 placeholder contract 包含重複名稱。");
        }

        try
        {
            // Own the stream even when package construction fails, so a bad
            // extracted template cannot leave a handle blocking lease cleanup.
            using var stream = File.OpenRead(fullPath);
            using var document = WordprocessingDocument.Open(stream, false);
            var mainDocumentPart = document.MainDocumentPart;
            if (mainDocumentPart?.Document.Body is null)
            {
                throw new InvalidDataException("Word 文件缺少主要文件內容。");
            }

            var counts = ReceiptOpenXmlParts.EnumerateRoots(mainDocumentPart)
                .SelectMany(root => _placeholderResolver.FindOccurrences(root))
                .GroupBy(occurrence => occurrence.Name.ToUpperInvariant(), StringComparer.OrdinalIgnoreCase)
                .ToDictionary(
                    group => group.Key,
                    group => group.Count(),
                    StringComparer.OrdinalIgnoreCase);

            var expected = expectedNames.ToHashSet(StringComparer.OrdinalIgnoreCase);
            var unknown = counts.Keys
                .Where(name => !expected.Contains(name))
                .OrderBy(name => name, StringComparer.OrdinalIgnoreCase)
                .ToArray();
            var mismatches = expectedNames
                .Where(name => !counts.TryGetValue(name, out var count)
                    || count != ExpectedOccurrenceCountPerReceipt)
                .Select(name =>
                {
                    counts.TryGetValue(name, out var count);
                    return $"{name} 實際 {count} 次，預期 {ExpectedOccurrenceCountPerReceipt} 次";
                })
                .ToArray();

            if (unknown.Length > 0 || mismatches.Length > 0)
            {
                var details = new List<string>();
                if (mismatches.Length > 0)
                {
                    details.Add($"occurrence 不符：{string.Join("、", mismatches)}");
                }

                if (unknown.Length > 0)
                {
                    details.Add($"未知 placeholder：{string.Join("、", unknown)}");
                }

                throw InvalidTemplate(string.Join("；", details));
            }

            return new ReceiptTemplateContract(fullPath, counts);
        }
        catch (ReceiptGenerationException)
        {
            throw;
        }
        catch (Exception exception) when (exception is IOException
            or FileFormatException
            or InvalidDataException
            or InvalidOperationException
            or OpenXmlPackageException
            or UnauthorizedAccessException)
        {
            throw InvalidTemplate(
                $"無法以 Open XML SDK 開啟：{Path.GetFileName(fullPath)}",
                exception);
        }
    }

    private static string GetFullPath(string templatePath)
    {
        try
        {
            ArgumentException.ThrowIfNullOrWhiteSpace(templatePath);
            var fullPath = Path.GetFullPath(templatePath);
            if (!File.Exists(fullPath))
            {
                throw InvalidTemplate($"找不到內建收據模板：{fullPath}");
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
            throw InvalidTemplate("內建收據模板路徑無效。", exception);
        }
    }

    private static ReceiptGenerationException InvalidTemplate(
        string detail,
        Exception? innerException = null) =>
        new(
            ReceiptGenerationErrorCode.ReceiptTemplateInvalid,
            $"ReceiptTemplateInvalid：內建收據模板不符合 production contract。{detail}",
            innerException);
}
