using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;

namespace XlsxDocxGenerator.Services.Receipts;

internal static class ReceiptOpenXmlParts
{
    public static IEnumerable<OpenXmlElement> EnumerateRoots(MainDocumentPart mainDocumentPart)
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

    public static IEnumerable<HeaderPart> EnumerateHeaders(MainDocumentPart mainDocumentPart) =>
        mainDocumentPart.HeaderParts
            .GroupBy(part => part.Uri.ToString(), StringComparer.OrdinalIgnoreCase)
            .Select(group => group.First())
            .OrderBy(part => part.Uri.ToString(), StringComparer.OrdinalIgnoreCase);

    public static IEnumerable<FooterPart> EnumerateFooters(MainDocumentPart mainDocumentPart) =>
        mainDocumentPart.FooterParts
            .GroupBy(part => part.Uri.ToString(), StringComparer.OrdinalIgnoreCase)
            .Select(group => group.First())
            .OrderBy(part => part.Uri.ToString(), StringComparer.OrdinalIgnoreCase);
}
