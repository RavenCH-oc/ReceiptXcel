using System.Security.Cryptography;
using ClosedXML.Excel;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using XlsxDocxGenerator.Models;
using XlsxDocxGenerator.Services.Receipts;

namespace XlsxDocxGenerator.Tests;

public sealed class ReceiptGenerationTests
{
    private const string WorksheetName = "工作表1";
    private const string QualityAssuranceOutput = "artifacts/qa/receipt-row-3.docx";

    [Fact]
    public void InternalTemplateSatisfiesProductionContract()
    {
        var contract = new ReceiptTemplateContractValidator().Validate(TemplatePath());

        Assert.Equal(ReceiptTemplateMapping.Fields.Count, contract.OccurrenceCounts.Count);
        Assert.All(
            ReceiptTemplateMapping.Fields,
            field => Assert.Equal(
                ReceiptTemplateContractValidator.ExpectedOccurrenceCountPerReceipt,
                contract.OccurrenceCounts[field.Placeholder]));
    }

    [Fact]
    public void UnknownTemplatePlaceholderFailsClosed()
    {
        var directory = CreateTemporaryDirectory();
        var malformed = Path.Combine(directory, "unknown-placeholder.docx");

        try
        {
            File.Copy(TemplatePath(), malformed);
            ReplaceFirstPlaceholder(malformed, "{{PAYER}}", "{{UNKNOWN}}");

            var exception = Assert.Throws<ReceiptGenerationException>(
                () => new ReceiptTemplateContractValidator().Validate(malformed));

            Assert.Equal(ReceiptGenerationErrorCode.ReceiptTemplateInvalid, exception.Code);
        }
        finally
        {
            DeleteTemporaryDirectory(directory);
        }
    }

    [Fact]
    public void TemplateOccurrenceMismatchFailsClosed()
    {
        var directory = CreateTemporaryDirectory();
        var malformed = Path.Combine(directory, "missing-occurrence.docx");

        try
        {
            File.Copy(TemplatePath(), malformed);
            ReplaceFirstPlaceholder(malformed, "{{PAYER}}", string.Empty);

            var exception = Assert.Throws<ReceiptGenerationException>(
                () => new ReceiptTemplateContractValidator().Validate(malformed));

            Assert.Equal(ReceiptGenerationErrorCode.ReceiptTemplateInvalid, exception.Code);
        }
        finally
        {
            DeleteTemporaryDirectory(directory);
        }
    }

    [Fact]
    public async Task ReferenceRowThreeGeneratesValidatedQaArtifactWithoutChangingInputs()
    {
        var excelPath = ReferenceExcelPath();
        var wordPath = ReferenceWordPath();
        var excelHashBefore = Hash(excelPath);
        var wordHashBefore = Hash(wordPath);
        var outputPath = Path.Combine(RepositoryRoot(), QualityAssuranceOutput);
        Directory.CreateDirectory(Path.GetDirectoryName(outputPath)!);
        if (File.Exists(outputPath))
        {
            File.Delete(outputPath);
        }

        var service = new ReceiptGenerationService(internalTemplatePath: TemplatePath());
        var generatedPath = await service.GenerateReceiptAsync(
            excelPath,
            WorksheetName,
            3,
            outputPath);

        var record = await new ReceiptRecordReader().ReadRecordAsync(excelPath, WorksheetName, 3);
        var amount = new ReceiptAmountFormatter().Format(record.Amount);
        new ReceiptGeneratedDocumentValidator().Validate(generatedPath, record, amount);

        Assert.Equal(Path.GetFullPath(outputPath), generatedPath);
        Assert.True(File.Exists(generatedPath));
        Assert.Equal(excelHashBefore, Hash(excelPath));
        Assert.Equal(wordHashBefore, Hash(wordPath));
        Assert.Empty(Directory.GetFiles(Path.GetDirectoryName(generatedPath)!, ".*.tmp"));

        using var document = WordprocessingDocument.Open(generatedPath, false);
        var body = document.MainDocumentPart!.Document.Body!;
        var text = AllText(body);
        Assert.DoesNotContain("{{", text, StringComparison.Ordinal);
        Assert.Contains("21813799", text, StringComparison.Ordinal);
        Assert.Equal(3, CountOccurrences(text, record.Payer));
        Assert.Equal(3, CountOccurrences(text, record.Reason));
        Assert.Equal(3, CountOccurrences(text, record.ReceiptNumber));
        Assert.Equal(3, CountOccurrences(text, amount.ChineseUppercase));

        var tables = body.Descendants<Table>().ToArray();
        Assert.Equal(3, tables.Length);
        Assert.All(tables, table =>
        {
            Assert.Equal(5, table.Elements<TableRow>().Count());
            Assert.Equal(17, table.GetFirstChild<TableGrid>()!.Elements<GridColumn>().Count());
            Assert.NotNull(table.Descendants<TableBorders>().SingleOrDefault());

            var dataCells = table.Elements<TableRow>().ElementAt(2).Elements<TableCell>().ToArray();
            var amountCells = dataCells.Skip(2).Take(7).Select(CellText).ToArray();
            Assert.Equal(
                [
                    amount.MillionCell,
                    amount.HundredThousandCell,
                    amount.TenThousandCell,
                    amount.ThousandCell,
                    amount.HundredCell,
                    amount.TenCell,
                    amount.OneCell
                ],
                amountCells);

            var handlerCell = table.Elements<TableRow>().ElementAt(4).Elements<TableCell>().ElementAt(1);
            Assert.Equal(string.Empty, CellText(handlerCell));
        });

        Assert.Equal(CaptureLayout(TemplatePath()), CaptureLayout(generatedPath));
    }

    [Theory]
    [InlineData(300)]
    [InlineData(12_000)]
    [InlineData(60_000)]
    [InlineData(888_888)]
    [InlineData(1_000_000)]
    public async Task AmountBoundaryCasesGenerateThroughTheProductionService(int amountValue)
    {
        var directory = CreateTemporaryDirectory();
        var excelPath = Path.Combine(directory, "input.xlsx");

        try
        {
            CreateWorkbook(excelPath, amountValue);
            var service = new ReceiptGenerationService(internalTemplatePath: TemplatePath());
            var generatedPath = await service.GenerateReceiptToDirectoryAsync(
                excelPath,
                WorksheetName,
                3,
                directory);

            var record = await new ReceiptRecordReader().ReadRecordAsync(excelPath, WorksheetName, 3);
            var formatted = new ReceiptAmountFormatter().Format(amountValue);
            new ReceiptGeneratedDocumentValidator().Validate(generatedPath, record, formatted);

            Assert.Equal(new ReceiptFilenamePolicy().GetFileName(record), Path.GetFileName(generatedPath));
            Assert.True(File.Exists(generatedPath));
        }
        finally
        {
            DeleteTemporaryDirectory(directory);
        }
    }

    [Fact]
    public void PlaceholderBuilderContainsAllCentralizedValues()
    {
        var record = new ReceiptRecord
        {
            ExcelRowNumber = 3,
            ROCYear = 112,
            Month = 12,
            Day = 30,
            ReceiptSerial = "0051",
            Payer = "A實業股份有限公司",
            Amount = 1_000,
            Reason = "運動會禮金",
            Handler = string.Empty
        };
        var amount = new ReceiptAmountFormatter().Format(record.Amount);

        var values = new ReceiptPlaceholderValuesBuilder().Build(record, amount);

        Assert.Equal(ReceiptTemplateMapping.Fields.Count, values.Count);
        Assert.Equal("112", values["ROC_YEAR"]);
        Assert.Equal("0051", record.ReceiptSerial);
        Assert.Equal(record.ReceiptNumber, values["RECEIPT_NUMBER"]);
        Assert.Equal(record.Payer, values["PAYER"]);
        Assert.Equal(record.Reason, values["REASON"]);
        Assert.Equal(string.Empty, values["HANDLER"]);
        Assert.Equal(amount.ChineseUppercase, values["AMOUNT_UPPER"]);
        Assert.Equal(amount.ThousandCell, values["AMOUNT_1000"]);
    }

    [Fact]
    public async Task ExistingOutputIsNotOverwritten()
    {
        var directory = CreateTemporaryDirectory();
        var outputPath = Path.Combine(directory, "existing.docx");
        const string sentinel = "do not overwrite";
        File.WriteAllText(outputPath, sentinel);

        try
        {
            var service = new ReceiptGenerationService(internalTemplatePath: TemplatePath());

            var exception = await Assert.ThrowsAsync<ReceiptGenerationException>(() =>
                service.GenerateReceiptAsync(
                    ReferenceExcelPath(),
                    WorksheetName,
                    3,
                    outputPath));

            Assert.Equal(ReceiptGenerationErrorCode.OutputAlreadyExists, exception.Code);
            Assert.Equal(sentinel, File.ReadAllText(outputPath));
            Assert.Empty(Directory.GetFiles(directory, ".*.tmp"));
        }
        finally
        {
            DeleteTemporaryDirectory(directory);
        }
    }

    [Fact]
    public async Task PreCancelledGenerationLeavesNoOutputOrTemporaryFile()
    {
        var directory = CreateTemporaryDirectory();
        var outputPath = Path.Combine(directory, "cancelled.docx");

        try
        {
            using var cancellation = new CancellationTokenSource();
            cancellation.Cancel();
            var service = new ReceiptGenerationService(internalTemplatePath: TemplatePath());

            await Assert.ThrowsAnyAsync<OperationCanceledException>(() =>
                service.GenerateReceiptAsync(
                    ReferenceExcelPath(),
                    WorksheetName,
                    3,
                    outputPath,
                    cancellation.Token));

            Assert.False(File.Exists(outputPath));
            Assert.Empty(Directory.GetFiles(directory, ".*.tmp"));
        }
        finally
        {
            DeleteTemporaryDirectory(directory);
        }
    }

    private static string RepositoryRoot()
    {
        var directory = new DirectoryInfo(AppContext.BaseDirectory);
        while (directory is not null)
        {
            if (File.Exists(Path.Combine(directory.FullName, "ReceiptXcel.sln")))
            {
                return directory.FullName;
            }

            directory = directory.Parent;
        }

        throw new DirectoryNotFoundException("ReceiptXcel solution root was not found.");
    }

    private static string TemplatePath() =>
        Path.Combine(RepositoryRoot(), ReceiptTemplateMapping.RelativeTemplatePath);

    private static string ReferenceExcelPath() =>
        Path.Combine(RepositoryRoot(), "reference", "excel.xlsx");

    private static string ReferenceWordPath() =>
        Path.Combine(RepositoryRoot(), "reference", "word.docx");

    private static string CreateTemporaryDirectory()
    {
        var directory = Path.Combine(
            Path.GetTempPath(),
            "ReceiptXcelTests",
            Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        return directory;
    }

    private static void DeleteTemporaryDirectory(string directory)
    {
        if (Directory.Exists(directory))
        {
            Directory.Delete(directory, recursive: true);
        }
    }

    private static void CreateWorkbook(string path, int amount)
    {
        using var workbook = new XLWorkbook();
        var worksheet = workbook.Worksheets.Add(WorksheetName);
        var headers = ReceiptWorksheetSchema.ExpectedHeaders;
        for (var column = 0; column < headers.Count; column++)
        {
            worksheet.Cell(2, column + 1).Value = headers[column];
        }

        worksheet.Cell(3, 1).Value = 112;
        worksheet.Cell(3, 2).Value = 12;
        worksheet.Cell(3, 3).Value = 30;
        worksheet.Cell(3, 4).Value = "0051";
        worksheet.Cell(3, 4).Style.NumberFormat.Format = "@";
        worksheet.Cell(3, 5).Value = "A實業股份有限公司";
        worksheet.Cell(3, 6).Value = amount;
        worksheet.Cell(3, 6).Style.NumberFormat.Format = "#,##0";
        worksheet.Cell(3, 7).Value = "運動會禮金";
        worksheet.Cell(3, 8).Value = string.Empty;
        workbook.SaveAs(path);
    }

    private static void ReplaceFirstPlaceholder(string path, string source, string replacement)
    {
        using var document = WordprocessingDocument.Open(path, true);
        var text = document.MainDocumentPart!
            .Document
            .Descendants<Text>()
            .FirstOrDefault(node => (node.Text ?? string.Empty).Contains(source, StringComparison.Ordinal));
        Assert.NotNull(text);
        text!.Text = text.Text!.Replace(source, replacement, StringComparison.Ordinal);
        document.MainDocumentPart.Document.Save();
    }

    private static string Hash(string path) =>
        Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(path)));

    private static string AllText(OpenXmlElement root) =>
        string.Concat(root.Descendants<Text>().Select(text => text.Text ?? string.Empty));

    private static string CellText(TableCell cell) => AllText(cell);

    private static int CountOccurrences(string text, string value)
    {
        var count = 0;
        var index = 0;
        while ((index = text.IndexOf(value, index, StringComparison.Ordinal)) >= 0)
        {
            count++;
            index += value.Length;
        }

        return count;
    }

    private static LayoutSnapshot CaptureLayout(string path)
    {
        using var document = WordprocessingDocument.Open(path, false);
        var body = document.MainDocumentPart!.Document.Body!;
        var section = body.Elements<SectionProperties>().Single();
        var tablesFingerprint = string.Join(
            "\n",
            body.Descendants<Table>().Select(table => string.Join(
                "\n",
                new[]
                {
                    table.GetFirstChild<TableProperties>()?.OuterXml ?? string.Empty,
                    table.GetFirstChild<TableGrid>()?.OuterXml ?? string.Empty
                }.Concat(table.Elements<TableRow>().Select(row => string.Join(
                    "\n",
                    new[]
                    {
                        row.GetFirstChild<TableRowProperties>()?.OuterXml ?? string.Empty
                    }.Concat(row.Elements<TableCell>().Select(cell =>
                        cell.GetFirstChild<TableCellProperties>()?.OuterXml ?? string.Empty))))))));

        return new LayoutSnapshot(
            section.GetFirstChild<PageSize>()?.OuterXml ?? string.Empty,
            section.GetFirstChild<PageMargin>()?.OuterXml ?? string.Empty,
            tablesFingerprint);
    }

    private sealed record LayoutSnapshot(
        string PageSize,
        string PageMargin,
        string TablesFingerprint);
}
