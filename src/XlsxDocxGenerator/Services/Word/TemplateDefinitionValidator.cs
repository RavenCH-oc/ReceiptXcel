using XlsxDocxGenerator.Models;
using XlsxDocxGenerator.Services.Errors;

namespace XlsxDocxGenerator.Services.Word;

public sealed class TemplateDefinitionValidator
{
    public void ValidatePlaceholders(
        TemplateDefinition template,
        IReadOnlyCollection<string> placeholders)
    {
        ArgumentNullException.ThrowIfNull(template);
        ArgumentNullException.ThrowIfNull(placeholders);

        var mappedNames = template.FieldMappings
            .Select(mapping => mapping.PlaceholderName.ToUpperInvariant())
            .ToHashSet(StringComparer.OrdinalIgnoreCase);

        var missingNames = placeholders
            .Select(name => name.ToUpperInvariant())
            .Where(name => !mappedNames.Contains(name))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToArray();

        if (missingNames.Length == 0)
        {
            return;
        }

        throw new GenerationException(
            GenerationErrorCode.MissingMapping,
            $"Word 模板包含尚未建立 Excel mapping 的欄位：{string.Join(", ", missingNames)}。");
    }
}
