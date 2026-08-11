using XlsxDocxGenerator.Models;

namespace XlsxDocxGenerator.Services.Excel;

public interface IExcelReader
{
    Task<IReadOnlyList<string>> GetWorksheetNamesAsync(
        string workbookPath,
        CancellationToken cancellationToken = default);

    Task<ExcelWorksheetData> ReadAsync(
        string workbookPath,
        string worksheetName,
        int headerRowNumber = 1,
        CancellationToken cancellationToken = default);

    Task<ExcelRecord> ReadRecordAsync(
        string workbookPath,
        TemplateDefinition template,
        int rowNumber,
        CancellationToken cancellationToken = default);
}
