using System.IO;
using System.Text;
using XlsxDocxGenerator.Models;

namespace XlsxDocxGenerator.Services.Receipts;

/// <summary>
/// Phase 1's fixed, deterministic output filename policy. User-configurable
/// patterns remain intentionally outside this phase.
/// </summary>
public sealed class ReceiptFilenamePolicy
{
    private static readonly HashSet<char> InvalidWindowsCharacters =
        new("<>:\"/\\|?*".ToCharArray());

    public string GetFileName(ReceiptRecord record)
    {
        ArgumentNullException.ThrowIfNull(record);
        return Sanitize($"收據_{record.ReceiptNumber}.docx");
    }

    public string GetOutputPath(string outputDirectory, ReceiptRecord record)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(outputDirectory);
        ArgumentNullException.ThrowIfNull(record);
        return Path.Combine(outputDirectory, GetFileName(record));
    }

    private static string Sanitize(string value)
    {
        var builder = new StringBuilder(value.Length);
        foreach (var character in value)
        {
            builder.Append(InvalidWindowsCharacters.Contains(character) ? '_' : character);
        }

        return builder.ToString().TrimEnd(' ', '.');
    }
}
