using System.Security.Cryptography;
using ClosedXML.Excel;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using XlsxDocxGenerator.Models;
using XlsxDocxGenerator.Services.Receipts;

namespace XlsxDocxGenerator.Tests;

public sealed class ReceiptGhFieldSemanticsTests
{
    private const string ReasonSentinel = "運動會禮金";
    private const string HColumnSentinel = "6/6匯款";
    private const string SecondReasonSentinel = "REASON_SENTINEL";
    private const string SecondHColumnSentinel = "H_COLUMN_SENTINEL";

    [Fact]
    public async Task ColumnGMapsToReasonAndColumnHRemainsSourceOnly()
    {
        using var workspace = new TemporaryWorkspace();
        var workbook = CreateWorkbook(
            workspace.Path,
            "direct-column-mapping.xlsx",
            reason: SecondReasonSentinel,
            sourceColumnH: SecondHColumnSentinel);

        var record = await new ReceiptRecordReader().ReadRecordAsync(
            workbook,
            ReceiptWorksheetSchema.WorksheetName,
            3);

        Assert.Equal("PAYER_SENTINEL", record.Payer);
        Assert.Equal(1_000, record.Amount);
        Assert.Equal(SecondReasonSentinel, record.Reason);
        Assert.Equal(SecondHColumnSentinel, record.Handler);
    }

    [Theory]
    [InlineData("")]
    [InlineData("6/6匯款")]
    [InlineData("ABC")]
    [InlineData("王小明")]
    public void PlaceholderBuilderAlwaysLeavesHandlerEmptyForSourceColumnH(string sourceColumnH)
    {
        var record = CreateRecord(ReasonSentinel, sourceColumnH);
        var values = new ReceiptPlaceholderValuesBuilder().Build(
            record,
            new ReceiptAmountFormatter().Format(record.Amount));

        Assert.Equal(ReasonSentinel, values["REASON"]);
        Assert.Equal(string.Empty, values["HANDLER"]);
    }

    [Fact]
    public void InternalTemplateKeepsReasonAndHandlerPlaceholdersInTheirFixedCells()
    {
        using var document = WordprocessingDocument.Open(TemplatePath(), false);
        var body = document.MainDocumentPart!.Document.Body!;
        var tables = body.Descendants<Table>().ToArray();

        Assert.Equal(3, tables.Length);
        foreach (var table in tables)
        {
            var rows = table.Elements<TableRow>().ToArray();
            Assert.Equal(5, rows.Length);

            var reasonCells = rows[2].Elements<TableCell>().ToArray();
            var handlerCells = rows[4].Elements<TableCell>().ToArray();
            Assert.Equal("{{REASON}}", CellText(reasonCells[9]));
            Assert.Equal("經手人", CellText(handlerCells[0]));
            Assert.Equal("{{HANDLER}}", CellText(handlerCells[1]));
        }

        Assert.Equal(3, body.Descendants<Text>().Count(text => text.Text == "{{REASON}}"));
        Assert.Equal(3, body.Descendants<Text>().Count(text => text.Text == "{{HANDLER}}"));
    }

    [Fact]
    public async Task HColumnNeverAppearsInDocxAndAllThreeHandlerCellsStayBlank()
    {
        using var workspace = new TemporaryWorkspace();
        var excelHashBefore = Hash(ReferenceExcelPath());
        var wordHashBefore = Hash(ReferenceWordPath());
        var workbook = CreateWorkbook(
            workspace.Path,
            "gh-primary-sentinel.xlsx",
            reason: ReasonSentinel,
            sourceColumnH: HColumnSentinel);
        var output = Path.Combine(workspace.Path, "gh-primary-sentinel.docx");

        var generated = await new ReceiptGenerationService(internalTemplatePath: TemplatePath())
            .GenerateReceiptAsync(
                workbook,
                ReceiptWorksheetSchema.WorksheetName,
                3,
                output);

        var record = await new ReceiptRecordReader().ReadRecordAsync(
            workbook,
            ReceiptWorksheetSchema.WorksheetName,
            3);
        Assert.Equal(ReasonSentinel, record.Reason);
        Assert.Equal(HColumnSentinel, record.Handler);
        new ReceiptGeneratedDocumentValidator().Validate(
            generated,
            record,
            new ReceiptAmountFormatter().Format(record.Amount));

        AssertGeneratedSemantics(generated, ReasonSentinel, HColumnSentinel);
        Assert.Equal(CaptureLayout(TemplatePath()), CaptureLayout(generated));
        Assert.Equal(excelHashBefore, Hash(ReferenceExcelPath()));
        Assert.Equal(wordHashBefore, Hash(ReferenceWordPath()));
    }

    [Fact]
    public async Task SecondGHSentinelUsesGInEveryReasonCellAndExcludesHCompletely()
    {
        using var workspace = new TemporaryWorkspace();
        var workbook = CreateWorkbook(
            workspace.Path,
            "gh-secondary-sentinel.xlsx",
            reason: SecondReasonSentinel,
            sourceColumnH: SecondHColumnSentinel);
        var output = Path.Combine(workspace.Path, "gh-secondary-sentinel.docx");

        var generated = await new ReceiptGenerationService(internalTemplatePath: TemplatePath())
            .GenerateReceiptAsync(
                workbook,
                ReceiptWorksheetSchema.WorksheetName,
                3,
                output);

        AssertGeneratedSemantics(generated, SecondReasonSentinel, SecondHColumnSentinel);
    }

    [Fact]
    public async Task ChangingOnlyColumnHDoesNotChangeGeneratedReceiptContentOrFilename()
    {
        using var workspace = new TemporaryWorkspace();
        var emptyHWorkbook = CreateWorkbook(
            workspace.Path,
            "empty-h.xlsx",
            reason: ReasonSentinel,
            sourceColumnH: string.Empty);
        var populatedHWorkbook = CreateWorkbook(
            workspace.Path,
            "populated-h.xlsx",
            reason: ReasonSentinel,
            sourceColumnH: SecondHColumnSentinel);
        var emptyHOutput = Path.Combine(workspace.Path, "empty-h.docx");
        var populatedHOutput = Path.Combine(workspace.Path, "populated-h.docx");
        var service = new ReceiptGenerationService(internalTemplatePath: TemplatePath());

        var emptyHRecord = await new ReceiptRecordReader().ReadRecordAsync(
            emptyHWorkbook,
            ReceiptWorksheetSchema.WorksheetName,
            3);
        var populatedHRecord = await new ReceiptRecordReader().ReadRecordAsync(
            populatedHWorkbook,
            ReceiptWorksheetSchema.WorksheetName,
            3);
        var emptyHGenerated = await service.GenerateReceiptAsync(emptyHRecord, emptyHOutput);
        var populatedHGenerated = await service.GenerateReceiptAsync(populatedHRecord, populatedHOutput);

        Assert.Equal(string.Empty, emptyHRecord.Handler);
        Assert.Equal(SecondHColumnSentinel, populatedHRecord.Handler);
        Assert.Equal(emptyHRecord.ROCYear, populatedHRecord.ROCYear);
        Assert.Equal(emptyHRecord.Month, populatedHRecord.Month);
        Assert.Equal(emptyHRecord.Day, populatedHRecord.Day);
        Assert.Equal(emptyHRecord.ReceiptNumber, populatedHRecord.ReceiptNumber);
        Assert.Equal(emptyHRecord.Payer, populatedHRecord.Payer);
        Assert.Equal(emptyHRecord.Amount, populatedHRecord.Amount);
        Assert.Equal(emptyHRecord.Reason, populatedHRecord.Reason);
        Assert.Equal(
            new ReceiptFilenamePolicy().GetFileName(emptyHRecord),
            new ReceiptFilenamePolicy().GetFileName(populatedHRecord));
        Assert.Equal(AllDocumentText(emptyHGenerated), AllDocumentText(populatedHGenerated));
        AssertGeneratedSemantics(populatedHGenerated, ReasonSentinel, SecondHColumnSentinel);
    }

    [Fact]
    public async Task QaArtifactUsesGReasonAndExcludesHColumnValue()
    {
        using var workspace = new TemporaryWorkspace();
        var qaDirectory = Path.Combine(RepositoryRoot(), "artifacts", "qa", "phase2");
        Directory.CreateDirectory(qaDirectory);
        var output = Path.Combine(qaDirectory, "gh-mapping-hotfix.docx");
        if (File.Exists(output))
        {
            File.Delete(output);
        }

        var generated = await new ReceiptGenerationService(internalTemplatePath: TemplatePath())
            .GenerateReceiptAsync(
                CreateWorkbook(
                    workspace.Path,
                    "gh-mapping-hotfix.xlsx",
                    reason: ReasonSentinel,
                    sourceColumnH: HColumnSentinel),
                ReceiptWorksheetSchema.WorksheetName,
                3,
                output);

        Assert.Equal(Path.GetFullPath(output), generated);
        AssertGeneratedSemantics(generated, ReasonSentinel, HColumnSentinel);
    }

    private static ReceiptRecord CreateRecord(string reason, string sourceColumnH) => new()
    {
        ExcelRowNumber = 3,
        ROCYear = 112,
        Month = 12,
        Day = 30,
        ReceiptSerial = "0051",
        Payer = "PAYER_SENTINEL",
        Amount = 1_000,
        Reason = reason,
        Handler = sourceColumnH
    };

    private static string CreateWorkbook(
        string directory,
        string fileName,
        string reason,
        string sourceColumnH)
    {
        var path = Path.Combine(directory, fileName);
        using var workbook = new XLWorkbook();
        var worksheet = workbook.Worksheets.Add(ReceiptWorksheetSchema.WorksheetName);
        foreach (var (header, index) in ReceiptWorksheetSchema.ExpectedHeaders.Select((value, index) => (value, index)))
        {
            worksheet.Cell(2, index + 1).Value = header;
        }

        worksheet.Cell(3, 1).Value = 112;
        worksheet.Cell(3, 2).Value = 12;
        worksheet.Cell(3, 3).Value = 30;
        worksheet.Cell(3, 4).Value = "0051";
        worksheet.Cell(3, 4).Style.NumberFormat.Format = "@";
        worksheet.Cell(3, 5).Value = "PAYER_SENTINEL";
        worksheet.Cell(3, 6).Value = "1000";
        worksheet.Cell(3, 7).Value = reason;
        worksheet.Cell(3, 8).Value = sourceColumnH;
        workbook.SaveAs(path);
        return path;
    }

    private static void AssertGeneratedSemantics(
        string documentPath,
        string expectedReason,
        string excludedSourceColumnH)
    {
        using var document = WordprocessingDocument.Open(documentPath, false);
        var body = document.MainDocumentPart!.Document.Body!;
        var visibleText = AllText(body);
        var tables = body.Descendants<Table>().ToArray();

        Assert.Equal(3, CountOccurrences(visibleText, expectedReason));
        Assert.DoesNotContain(excludedSourceColumnH, visibleText, StringComparison.Ordinal);
        Assert.Equal(3, tables.Length);
        foreach (var table in tables)
        {
            var rows = table.Elements<TableRow>().ToArray();
            Assert.Equal(5, rows.Length);
            Assert.Equal(17, table.GetFirstChild<TableGrid>()!.Elements<GridColumn>().Count());
            Assert.Equal(expectedReason, CellText(rows[2].Elements<TableCell>().ElementAt(9)));
            Assert.Equal(string.Empty, CellText(rows[4].Elements<TableCell>().ElementAt(1)));
        }
    }

    private static string AllDocumentText(string documentPath)
    {
        using var document = WordprocessingDocument.Open(documentPath, false);
        return AllText(document.MainDocumentPart!.Document.Body!);
    }

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

    private static string Hash(string path) =>
        Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(path)));

    private static string TemplatePath() =>
        Path.Combine(RepositoryRoot(), ReceiptTemplateMapping.RelativeTemplatePath);

    private static string ReferenceExcelPath() =>
        Path.Combine(RepositoryRoot(), "reference", "excel.xlsx");

    private static string ReferenceWordPath() =>
        Path.Combine(RepositoryRoot(), "reference", "word.docx");

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

    private sealed class TemporaryWorkspace : IDisposable
    {
        public TemporaryWorkspace()
        {
            Path = System.IO.Path.Combine(
                System.IO.Path.GetTempPath(),
                $"receiptxcel-gh-semantics-{Guid.NewGuid():N}");
            Directory.CreateDirectory(Path);
        }

        public string Path { get; }

        public void Dispose()
        {
            if (Directory.Exists(Path))
            {
                Directory.Delete(Path, recursive: true);
            }
        }
    }
}
