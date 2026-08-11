using System.IO;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Packaging;
using XlsxDocxGenerator.Models;
using XlsxDocxGenerator.Services.Errors;

namespace XlsxDocxGenerator.Services.Excel;

/// <summary>
/// Excel infrastructure boundary. ClosedXML types never leave this service.
/// </summary>
public sealed class ExcelReader : IExcelReader
{
    private readonly ICellTextPolicy _cellTextPolicy;

    public ExcelReader(ICellTextPolicy? cellTextPolicy = null)
    {
        _cellTextPolicy = cellTextPolicy ?? new DisplayCellTextPolicy();
    }

    public Task<IReadOnlyList<string>> GetWorksheetNamesAsync(
        string workbookPath,
        CancellationToken cancellationToken = default)
    {
        var workbook = OpenWorkbook(workbookPath);
        using (workbook)
        {
            cancellationToken.ThrowIfCancellationRequested();
            return Task.FromResult<IReadOnlyList<string>>(
                workbook.Worksheets.Select(sheet => sheet.Name).ToArray());
        }
    }

    public Task<ExcelWorksheetData> ReadAsync(
        string workbookPath,
        string worksheetName,
        int headerRowNumber = 1,
        CancellationToken cancellationToken = default)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(worksheetName);

        if (headerRowNumber <= 0)
        {
            throw new TemplateDefinitionException(
                TemplateDefinitionErrorCode.InvalidHeaderRow,
                "Header row 必須是大於 0 的 Excel 列號。");
        }

        var workbook = OpenWorkbook(workbookPath);
        using (workbook)
        {
            cancellationToken.ThrowIfCancellationRequested();

            var worksheet = workbook.Worksheets.FirstOrDefault(sheet =>
                string.Equals(sheet.Name, worksheetName, StringComparison.OrdinalIgnoreCase));

            if (worksheet is null)
            {
                throw new TemplateDefinitionException(
                    TemplateDefinitionErrorCode.WorksheetNotFound,
                    $"找不到工作表「{worksheetName}」。");
            }

            return Task.FromResult(ReadWorksheet(worksheet, worksheetName, headerRowNumber, cancellationToken));
        }
    }

    public async Task<ExcelRecord> ReadRecordAsync(
        string workbookPath,
        TemplateDefinition template,
        int rowNumber,
        CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(template);

        if (rowNumber <= template.HeaderRowNumber)
        {
            throw new GenerationException(
                GenerationErrorCode.InvalidExcelDataRow,
                $"Excel 第 {rowNumber} 列不是可產生文件的資料列。");
        }

        var worksheet = await ReadAsync(
            workbookPath,
            template.PreferredWorksheetName,
            template.HeaderRowNumber,
            cancellationToken);
        var context = new RuntimeWorksheetSchemaValidator().Validate(worksheet, template);
        if (rowNumber == context.RuntimeMappingMarkerRowNumber
            || IsSingleMappingMarkerRow(worksheet, template, rowNumber))
        {
            throw new GenerationException(
                GenerationErrorCode.InvalidExcelDataRow,
                $"Excel 列 {rowNumber} 是目前工作簿的欄位對應列，不能作為資料列。");
        }
        var record = worksheet.Records.FirstOrDefault(item => item.RowNumber == rowNumber);

        if (record is null)
        {
            throw new GenerationException(
                GenerationErrorCode.InvalidExcelDataRow,
                $"找不到 Excel 第 {rowNumber} 列，或該列沒有有效資料。");
        }

        return record;
    }

    private static bool IsSingleMappingMarkerRow(
        ExcelWorksheetData worksheet,
        TemplateDefinition template,
        int rowNumber)
    {
        if (template.FieldMappings.Count != 1)
        {
            return false;
        }

        var mapping = template.FieldMappings[0];
        var row = worksheet.Rows.FirstOrDefault(item => item.RowNumber == rowNumber);
        var expected = $"{{{{{mapping.PlaceholderName}}}}}";
        return row is not null
            && row.Cells.Keys.All(column => column == mapping.ColumnIndex)
            && string.Equals(row.GetCellText(mapping.ColumnIndex)?.Trim(), expected, StringComparison.OrdinalIgnoreCase);
    }

    private static XLWorkbook OpenWorkbook(string workbookPath)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(workbookPath);

        if (!File.Exists(workbookPath))
        {
            throw new TemplateDefinitionException(
                TemplateDefinitionErrorCode.ExcelFileNotFound,
                $"找不到 Excel 檔案：{workbookPath}");
        }

        try
        {
            return new XLWorkbook(workbookPath);
        }
        catch (Exception exception) when (exception is IOException
            or UnauthorizedAccessException
            or InvalidDataException
            or ArgumentException
            or OpenXmlPackageException)
        {
            throw new TemplateDefinitionException(
                TemplateDefinitionErrorCode.InvalidExcelFile,
                $"無法讀取 Excel 檔案，檔案可能損壞或格式不受支援：{Path.GetFileName(workbookPath)}",
                exception);
        }
        catch (Exception exception)
        {
            throw new TemplateDefinitionException(
                TemplateDefinitionErrorCode.InvalidExcelFile,
                $"無法讀取 Excel 檔案，檔案可能損壞或格式不受支援：{Path.GetFileName(workbookPath)}",
                exception);
        }
    }

    private ExcelWorksheetData ReadWorksheet(
        IXLWorksheet worksheet,
        string worksheetName,
        int headerRowNumber,
        CancellationToken cancellationToken)
    {
        // Contents excludes cells that only contain formatting, preventing an inflated UsedRange.
        var usedCells = worksheet.CellsUsed(XLCellsUsedOptions.Contents).ToArray();
        if (usedCells.Length == 0)
        {
            throw new TemplateDefinitionException(
                TemplateDefinitionErrorCode.InvalidHeaderRow,
                $"工作表「{worksheetName}」沒有可用內容，無法建立 header row。");
        }

        var firstUsedRow = usedCells.Min(cell => cell.Address.RowNumber);
        var lastUsedRow = usedCells.Max(cell => cell.Address.RowNumber);
        var lastUsedColumn = usedCells.Max(cell => cell.Address.ColumnNumber);

        if (headerRowNumber < firstUsedRow || headerRowNumber > lastUsedRow)
        {
            throw new TemplateDefinitionException(
                TemplateDefinitionErrorCode.InvalidHeaderRow,
                $"Header row {headerRowNumber} 不在工作表的有效內容範圍內（{firstUsedRow} 至 {lastUsedRow}）。");
        }

        var headers = Enumerable.Range(1, lastUsedColumn)
            .Select(column => _cellTextPolicy.GetText(worksheet.Cell(headerRowNumber, column)))
            .ToArray();

        var rows = new List<ExcelRowData>();
        foreach (var rowNumber in Enumerable.Range(headerRowNumber + 1, lastUsedRow - headerRowNumber))
        {
            cancellationToken.ThrowIfCancellationRequested();

            var cells = Enumerable.Range(1, lastUsedColumn)
                .Select(column => new
                {
                    Column = column,
                    Value = _cellTextPolicy.GetText(worksheet.Cell(rowNumber, column))
                })
                .Where(cell => !string.IsNullOrEmpty(cell.Value))
                .ToDictionary(cell => cell.Column, cell => cell.Value);

            if (cells.Count == 0)
            {
                continue;
            }

            rows.Add(new ExcelRowData
            {
                RowNumber = rowNumber,
                Cells = cells
            });
        }

        return new ExcelWorksheetData
        {
            WorksheetName = worksheetName,
            HeaderRowNumber = headerRowNumber,
            FirstUsedRowNumber = firstUsedRow,
            LastUsedRowNumber = lastUsedRow,
            Headers = headers,
            Rows = rows
        };
    }
}
