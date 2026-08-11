namespace XlsxDocxGenerator.Services.Word;

public interface IDocxTemplateReader
{
    Task<DocxTemplateInfo> ReadAsync(
        string templatePath,
        CancellationToken cancellationToken = default);

    Task<IReadOnlyList<string>> ScanPlaceholdersAsync(
        string templatePath,
        CancellationToken cancellationToken = default);
}

public sealed record DocxTemplateInfo(string TemplateName, string WordTemplatePath);
