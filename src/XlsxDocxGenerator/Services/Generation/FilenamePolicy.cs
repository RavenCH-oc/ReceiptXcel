using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using XlsxDocxGenerator.Models;
using XlsxDocxGenerator.Services.Errors;

namespace XlsxDocxGenerator.Services.Generation;

public sealed class FilenamePolicy
{
    private static readonly Regex ValidPlaceholder = new(
        "^\\{\\{(?<name>[A-Za-z_][A-Za-z0-9_]*)\\}\\}$",
        RegexOptions.Compiled | RegexOptions.CultureInvariant);

    private static readonly Regex PlaceholderLike = new(
        "\\{\\{[^{}]*\\}\\}",
        RegexOptions.Compiled | RegexOptions.CultureInvariant);

    private static readonly Regex ReservedDeviceName = new(
        "^(CON|PRN|AUX|NUL|COM[1-9]|LPT[1-9])$",
        RegexOptions.Compiled | RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);

    private static readonly HashSet<char> InvalidWindowsCharacters =
        new("<>:\"/\\|?*".ToCharArray());

    public string ResolveBaseName(
        string? pattern,
        TemplateDefinition template,
        ExcelRecord record)
    {
        ArgumentNullException.ThrowIfNull(template);
        ArgumentNullException.ThrowIfNull(record);

        var effectivePattern = string.IsNullOrWhiteSpace(pattern)
            ? "output-row-{{ROW}}"
            : pattern.Trim();

        var extension = Path.GetExtension(effectivePattern);
        if (!string.IsNullOrEmpty(extension))
        {
            if (!string.Equals(extension, ".docx", StringComparison.OrdinalIgnoreCase))
            {
                throw InvalidFilename("輸出檔名格式只能使用 .docx 副檔名。");
            }

            effectivePattern = effectivePattern[..^extension.Length];
        }

        var mappingByName = template.FieldMappings
            .ToDictionary(
                mapping => mapping.PlaceholderName.ToUpperInvariant(),
                mapping => mapping,
                StringComparer.OrdinalIgnoreCase);

        var replaced = PlaceholderLike.Replace(effectivePattern, match =>
        {
            var marker = ValidPlaceholder.Match(match.Value);
            if (!marker.Success)
            {
                throw InvalidFilename($"輸出檔名格式包含無效 placeholder：{match.Value}");
            }

            var name = marker.Groups["name"].Value.ToUpperInvariant();
            if (name == "ROW")
            {
                return record.RowNumber.ToString();
            }

            if (!mappingByName.TryGetValue(name, out var mapping))
            {
                throw InvalidFilename($"輸出檔名 placeholder「{name}」沒有對應 mapping。");
            }

            return record.GetValueByColumnIndex(mapping.ColumnIndex) ?? string.Empty;
        });

        if (replaced.Contains('{', StringComparison.Ordinal)
            || replaced.Contains('}', StringComparison.Ordinal))
        {
            throw InvalidFilename("輸出檔名格式包含未完成的 placeholder。");
        }

        var sanitized = Sanitize(replaced);
        if (sanitized.Length == 0 || sanitized is "." or "..")
        {
            throw InvalidFilename("輸出檔名不可為空白或特殊目錄名稱。");
        }

        return sanitized;
    }

    public void ValidatePatternSyntax(string? pattern, TemplateDefinition template)
    {
        ArgumentNullException.ThrowIfNull(template);

        var effectivePattern = string.IsNullOrWhiteSpace(pattern)
            ? "output-row-{{ROW}}"
            : pattern.Trim();
        var extension = Path.GetExtension(effectivePattern);
        if (!string.IsNullOrEmpty(extension)
            && !string.Equals(extension, ".docx", StringComparison.OrdinalIgnoreCase))
        {
            throw InvalidFilename("輸出檔名只能使用 .docx 副檔名。");
        }

        if (string.Equals(extension, ".docx", StringComparison.OrdinalIgnoreCase))
        {
            effectivePattern = effectivePattern[..^extension.Length];
        }

        var mappingNames = template.FieldMappings
            .Select(mapping => mapping.PlaceholderName.ToUpperInvariant())
            .ToHashSet(StringComparer.OrdinalIgnoreCase);
        foreach (Match match in PlaceholderLike.Matches(effectivePattern))
        {
            var marker = ValidPlaceholder.Match(match.Value);
            if (!marker.Success)
            {
                throw InvalidFilename($"輸出檔名包含無效 placeholder：{match.Value}");
            }

            var name = marker.Groups["name"].Value.ToUpperInvariant();
            if (name != "ROW" && !mappingNames.Contains(name))
            {
                throw InvalidFilename($"輸出檔名使用了未定義欄位：{name}");
            }
        }

        if (effectivePattern.Contains('{', StringComparison.Ordinal)
            || effectivePattern.Contains('}', StringComparison.Ordinal))
        {
            throw InvalidFilename("輸出檔名包含未完成的 placeholder。");
        }
    }

    public string ResolveOutputPath(
        string outputDirectory,
        string? pattern,
        TemplateDefinition template,
        ExcelRecord record)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(outputDirectory);
        var baseName = ResolveBaseName(pattern, template, record);
        return Path.Combine(outputDirectory, $"{baseName}.docx");
    }

    private static string Sanitize(string value)
    {
        var builder = new StringBuilder(value.Length);
        foreach (var character in value)
        {
            builder.Append(InvalidWindowsCharacters.Contains(character) ? '_' : character);
        }

        var sanitized = builder.ToString().TrimEnd(' ', '.');
        if (ReservedDeviceName.IsMatch(sanitized))
        {
            sanitized = $"_{sanitized}";
        }

        return sanitized;
    }

    private static GenerationException InvalidFilename(string message) =>
        new(GenerationErrorCode.InvalidOutputFileName, message);
}

public sealed class FilenameAllocator
{
    private readonly HashSet<string> _reservedPaths = new(StringComparer.OrdinalIgnoreCase);

    public string Allocate(string outputDirectory, string baseName)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(outputDirectory);
        ArgumentException.ThrowIfNullOrWhiteSpace(baseName);

        var suffix = 1;
        while (true)
        {
            var candidateName = suffix == 1
                ? $"{baseName}.docx"
                : $"{baseName} ({suffix}).docx";
            var candidatePath = Path.Combine(outputDirectory, candidateName);

            if (!_reservedPaths.Contains(candidatePath) && !File.Exists(candidatePath))
            {
                _reservedPaths.Add(candidatePath);
                return candidatePath!;
            }

            suffix++;
        }
    }
}
