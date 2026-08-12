using XlsxDocxGenerator.Models;

namespace XlsxDocxGenerator.Services.Receipts;

/// <summary>
/// Centralized production mapping from derived receipt data to internal Word
/// placeholders. Services and future UI code must consume this builder rather
/// than reconstructing placeholder names independently. The fixed G/H
/// semantics are enforced by the mapping: G supplies REASON, while source-only
/// H never supplies HANDLER.
/// </summary>
public sealed class ReceiptPlaceholderValuesBuilder
{
    public IReadOnlyDictionary<string, string?> Build(
        ReceiptRecord record,
        FormattedReceiptAmount amount)
    {
        ArgumentNullException.ThrowIfNull(record);
        ArgumentNullException.ThrowIfNull(amount);

        return ReceiptTemplateMapping.CreateValues(record, amount);
    }
}
