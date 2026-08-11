using ClosedXML.Excel;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using XlsxDocxGenerator.Models;
using XlsxDocxGenerator.Services.Errors;
using XlsxDocxGenerator.Services.Excel;
using XlsxDocxGenerator.Services.Generation;
using XlsxDocxGenerator.Services.Word;

namespace XlsxDocxGenerator.Tests;

public sealed class WordGenerationTests
{
    [Fact]
    public async Task ReplacesSinglePlaceholder()
    {
        using var workspace = new TemporaryWorkspace();
        var templatePath = CreateDocument(workspace.Path, new Paragraph(new Run(new Text("Hello {{NAME}}"))));

        var outputPath = await GenerateAsync(workspace, templatePath, [Mapping("NAME", 1, "Alice")]);

        Assert.Equal("Hello Alice", ReadAllText(outputPath));
    }

    [Fact]
    public async Task ReplacesSamePlaceholderTwice()
    {
        using var workspace = new TemporaryWorkspace();
        var templatePath = CreateDocument(workspace.Path, new Paragraph(new Run(new Text("{{NAME}} / {{NAME}}"))));

        var outputPath = await GenerateAsync(workspace, templatePath, [Mapping("NAME", 1, "Alice")]);

        Assert.Equal("Alice / Alice", ReadAllText(outputPath));
    }

    [Fact]
    public async Task ReplacesMultiplePlaceholdersInOneParagraph()
    {
        using var workspace = new TemporaryWorkspace();
        var templatePath = CreateDocument(workspace.Path, new Paragraph(new Run(new Text("{{FIRST}} {{LAST}}"))));

        var outputPath = await GenerateAsync(workspace, templatePath,
            [Mapping("FIRST", 1, "Ada"), Mapping("LAST", 2, "Lovelace")]);

        Assert.Equal("Ada Lovelace", ReadAllText(outputPath));
    }

    [Fact]
    public async Task ReplacesPlaceholderAcrossTwoRuns()
    {
        using var workspace = new TemporaryWorkspace();
        var templatePath = CreateDocument(workspace.Path,
            new Paragraph(new Run(new Text("{{NA")), new Run(new Text("ME}}"))));

        var outputPath = await GenerateAsync(workspace, templatePath, [Mapping("NAME", 1, "Alice")]);

        Assert.Equal("Alice", ReadAllText(outputPath));
    }

    [Fact]
    public async Task ReplacesPlaceholderAcrossFourRunsAndKeepsTextAroundIt()
    {
        using var workspace = new TemporaryWorkspace();
        var templatePath = CreateDocument(workspace.Path,
            new Paragraph(
                new Run(new Text("Before ")),
                new Run(new Text("{{")),
                new Run(new Text("NA")),
                new Run(new Text("ME")),
                new Run(new Text("}}")),
                new Run(new Text(" After"))));

        var outputPath = await GenerateAsync(workspace, templatePath, [Mapping("NAME", 1, "Alice")]);

        Assert.Equal("Before Alice After", ReadAllText(outputPath));
    }

    [Fact]
    public async Task ReplacesPlaceholderInsideTableCell()
    {
        using var workspace = new TemporaryWorkspace();
        var table = new Table(
            new TableRow(
                new TableCell(new Paragraph(new Run(new Text("Address: {{ADDRESS}}"))))));
        var templatePath = CreateDocument(workspace.Path, table);

        var outputPath = await GenerateAsync(workspace, templatePath, [Mapping("ADDRESS", 1, "台北市")]);

        Assert.Equal("Address: 台北市", ReadAllText(outputPath));
    }

    [Fact]
    public async Task SupportsChineseReplacement()
    {
        using var workspace = new TemporaryWorkspace();
        var templatePath = CreateDocument(workspace.Path, new Paragraph(new Run(new Text("{{NAME}}"))));

        var outputPath = await GenerateAsync(workspace, templatePath, [Mapping("NAME", 1, "王小明")]);

        Assert.Equal("王小明", ReadAllText(outputPath));
    }

    [Fact]
    public async Task SupportsEmptyReplacement()
    {
        using var workspace = new TemporaryWorkspace();
        var templatePath = CreateDocument(workspace.Path, new Paragraph(new Run(new Text("A{{NAME}}B"))));

        var outputPath = await GenerateAsync(workspace, templatePath, [Mapping("NAME", 1, null)]);

        Assert.Equal("AB", ReadAllText(outputPath));
    }

    [Fact]
    public async Task SupportsXmlSpecialCharacters()
    {
        using var workspace = new TemporaryWorkspace();
        var templatePath = CreateDocument(workspace.Path, new Paragraph(new Run(new Text("{{VALUE}}"))));

        var outputPath = await GenerateAsync(workspace, templatePath, [Mapping("VALUE", 1, "A & B < C > D")]);

        Assert.Equal("A & B < C > D", ReadAllText(outputPath));
    }

    [Fact]
    public async Task PreservesTextFormattingAtPlaceholderStart()
    {
        using var workspace = new TemporaryWorkspace();
        var boldRun = new Run(new RunProperties(new Bold()), new Text("{{NAME}}"));
        var templatePath = CreateDocument(workspace.Path, new Paragraph(boldRun));

        var outputPath = await GenerateAsync(workspace, templatePath, [Mapping("NAME", 1, "Alice")]);

        using var document = WordprocessingDocument.Open(outputPath, false);
        var replacementRun = document.MainDocumentPart!.Document.Body!
            .Descendants<Run>()
            .First(run => run.InnerText == "Alice");
        Assert.NotNull(replacementRun.RunProperties?.Bold);
    }

    [Fact]
    public async Task PreservesWhitespaceInReplacement()
    {
        using var workspace = new TemporaryWorkspace();
        var templatePath = CreateDocument(workspace.Path, new Paragraph(new Run(new Text("{{NAME}}"))));

        var outputPath = await GenerateAsync(workspace, templatePath, [Mapping("NAME", 1, " 王小明 ")]);

        using var document = WordprocessingDocument.Open(outputPath, false);
        var text = document.MainDocumentPart!.Document.Body!.Descendants<Text>().Single();
        Assert.Equal(" 王小明 ", text.Text);
        Assert.Equal(SpaceProcessingModeValues.Preserve, text.Space?.Value);
    }

    [Fact]
    public async Task ScansOnlyValidPlaceholders()
    {
        using var workspace = new TemporaryWorkspace();
        var templatePath = CreateDocument(workspace.Path,
            new Paragraph(new Run(new Text("{{NAME}} {{BAD FIELD}} {{123}} {NAME}"))));

        var names = await new DocxTemplateReader().ScanPlaceholdersAsync(templatePath);

        Assert.Equal(["NAME"], names);
    }

    [Fact]
    public async Task ScannerDeduplicatesPlaceholders()
    {
        using var workspace = new TemporaryWorkspace();
        var templatePath = CreateDocument(workspace.Path,
            new Paragraph(new Run(new Text("{{NAME}} {{NAME}} {{ADDRESS}}"))));

        var names = await new DocxTemplateReader().ScanPlaceholdersAsync(templatePath);

        Assert.Equal(["NAME", "ADDRESS"], names);
    }

    [Fact]
    public async Task ScannerCanonicalizesCase()
    {
        using var workspace = new TemporaryWorkspace();
        var templatePath = CreateDocument(workspace.Path,
            new Paragraph(new Run(new Text("{{name}} {{Name}}"))));

        var names = await new DocxTemplateReader().ScanPlaceholdersAsync(templatePath);

        Assert.Equal(["NAME"], names);
    }

    [Fact]
    public async Task MissingMappingIsFatal()
    {
        using var workspace = new TemporaryWorkspace();
        var templatePath = CreateDocument(workspace.Path,
            new Paragraph(new Run(new Text("{{NAME}} {{PHONE}}"))));

        var exception = await Assert.ThrowsAsync<GenerationException>(() =>
            GenerateAsync(workspace, templatePath, [Mapping("NAME", 1, "Alice")]));

        Assert.Equal(GenerationErrorCode.MissingMapping, exception.Code);
    }

    [Fact]
    public async Task UnusedMappingDoesNotBlockGeneration()
    {
        using var workspace = new TemporaryWorkspace();
        var templatePath = CreateDocument(workspace.Path, new Paragraph(new Run(new Text("{{NAME}}"))));

        var outputPath = await GenerateAsync(workspace, templatePath,
            [Mapping("NAME", 1, "Alice"), Mapping("UNUSED", 2, "ignored")]);

        Assert.Equal("Alice", ReadAllText(outputPath));
    }

    [Fact]
    public async Task ReopensGeneratedDocumentSuccessfully()
    {
        using var workspace = new TemporaryWorkspace();
        var templatePath = CreateDocument(workspace.Path, new Paragraph(new Run(new Text("{{NAME}}"))));

        var outputPath = await GenerateAsync(workspace, templatePath, [Mapping("NAME", 1, "Alice")]);

        using var document = WordprocessingDocument.Open(outputPath, false);
        Assert.NotNull(document.MainDocumentPart?.Document.Body);
    }

    [Fact]
    public async Task ExistingOutputFailsWithoutOverwriting()
    {
        using var workspace = new TemporaryWorkspace();
        var templatePath = CreateDocument(workspace.Path, new Paragraph(new Run(new Text("{{NAME}}"))));
        var outputPath = Path.Combine(workspace.Path, "output.docx");
        File.WriteAllText(outputPath, "existing");

        var exception = await Assert.ThrowsAsync<GenerationException>(() =>
            GenerateAsync(workspace, templatePath, [Mapping("NAME", 1, "Alice")], outputPath));

        Assert.Equal(GenerationErrorCode.OutputAlreadyExists, exception.Code);
        Assert.Equal("existing", File.ReadAllText(outputPath));
    }

    [Fact]
    public async Task OriginalTemplateIsNotModified()
    {
        using var workspace = new TemporaryWorkspace();
        var templatePath = CreateDocument(workspace.Path, new Paragraph(new Run(new Text("{{NAME}}"))));

        await GenerateAsync(workspace, templatePath, [Mapping("NAME", 1, "Alice")]);

        Assert.Equal("{{NAME}}", ReadAllText(templatePath));
    }

    [Fact]
    public async Task FailedGenerationLeavesNoTemporaryDocx()
    {
        using var workspace = new TemporaryWorkspace();
        var templatePath = CreateDocument(workspace.Path, new Paragraph(new Run(new Text("{{MISSING}}"))));
        var outputPath = Path.Combine(workspace.Path, "output.docx");

        await Assert.ThrowsAsync<GenerationException>(() =>
            new DocxGenerator().GenerateAsync(
                CreateDefinition(templatePath, [Mapping("NAME", 1, "Alice")]),
                CreateRecord((1, "Alice")),
                outputPath));

        Assert.Empty(Directory.GetFiles(workspace.Path, "*.tmp", SearchOption.TopDirectoryOnly));
        Assert.False(File.Exists(outputPath));
    }

    [Fact]
    public async Task GenerationServiceCreatesSingleDocumentFromExcelRow()
    {
        using var workspace = new TemporaryWorkspace();
        var workbookPath = CreateWorkbook(workspace.Path, (sheet, _) =>
        {
            SetRow(sheet, 1, "姓名");
            SetRow(sheet, 2, "王小明");
            SetRow(sheet, 3, "{{NAME}}");
        });
        var templatePath = CreateDocument(workspace.Path, new Paragraph(new Run(new Text("姓名：{{NAME}}"))));
        var outputPath = Path.Combine(workspace.Path, "single.docx");
        var definition = CreateDefinition(templatePath, [Mapping("NAME", 1, null)], "資料", 1, 3);

        await new GenerationService(new ExcelReader(), new DocxGenerator())
            .GenerateSingleAsync(workbookPath, definition, 2, outputPath);

        Assert.Equal("姓名：王小明", ReadAllText(outputPath));
    }

    [Fact]
    public async Task MarkerRowCannotBeUsedAsDataRow()
    {
        using var workspace = new TemporaryWorkspace();
        var workbookPath = CreateWorkbook(workspace.Path, (sheet, _) =>
        {
            SetRow(sheet, 1, "姓名");
            SetRow(sheet, 2, "王小明");
            SetRow(sheet, 3, "{{NAME}}");
        });
        var definition = CreateDefinition("template.docx", [Mapping("NAME", 1, null)], "資料", 1, 3);

        var exception = await Assert.ThrowsAsync<GenerationException>(() =>
            new ExcelReader().ReadRecordAsync(workbookPath, definition, 3));

        Assert.Equal(GenerationErrorCode.InvalidExcelDataRow, exception.Code);
    }

    [Fact]
    public async Task BlankRowCannotBeGenerated()
    {
        using var workspace = new TemporaryWorkspace();
        var workbookPath = CreateWorkbook(workspace.Path, (sheet, _) =>
        {
            SetRow(sheet, 1, "姓名");
            SetRow(sheet, 3, "{{NAME}}");
        });
        var definition = CreateDefinition("template.docx", [Mapping("NAME", 1, null)], "資料", 1, 3);

        var exception = await Assert.ThrowsAsync<GenerationException>(() =>
            new ExcelReader().ReadRecordAsync(workbookPath, definition, 2));

        Assert.Equal(GenerationErrorCode.InvalidExcelDataRow, exception.Code);
    }

    [Fact]
    public async Task LeadingZeroFormattingIsReadAsDisplayedText()
    {
        using var workspace = new TemporaryWorkspace();
        var workbookPath = CreateWorkbook(workspace.Path, (sheet, _) =>
        {
            SetRow(sheet, 1, "編號");
            sheet.Cell(2, 1).Value = 123;
            sheet.Cell(2, 1).Style.NumberFormat.Format = "00000";
            SetRow(sheet, 3, "{{ID}}");
        });
        var definition = CreateDefinition("template.docx", [Mapping("ID", 1, null)], "資料", 1, 3);

        var record = await new ExcelReader().ReadRecordAsync(workbookPath, definition, 2);

        Assert.Equal("00123", record.GetValueByColumnIndex(1));
    }

    [Fact]
    public async Task DateFormattingIsReadAsDisplayedText()
    {
        using var workspace = new TemporaryWorkspace();
        var workbookPath = CreateWorkbook(workspace.Path, (sheet, _) =>
        {
            SetRow(sheet, 1, "日期");
            sheet.Cell(2, 1).Value = new DateTime(2026, 8, 9);
            sheet.Cell(2, 1).Style.NumberFormat.Format = "yyyy/mm/dd";
            SetRow(sheet, 3, "{{DATE}}");
        });
        var definition = CreateDefinition("template.docx", [Mapping("DATE", 1, null)], "資料", 1, 3);

        var record = await new ExcelReader().ReadRecordAsync(workbookPath, definition, 2);

        Assert.Equal("2026/08/09", record.GetValueByColumnIndex(1));
    }

    private static async Task<string> GenerateAsync(
        TemporaryWorkspace workspace,
        string templatePath,
        IReadOnlyList<TestMapping> mappings,
        string? outputPath = null)
    {
        outputPath ??= Path.Combine(workspace.Path, $"output-{Guid.NewGuid():N}.docx");
        var definition = CreateDefinition(templatePath, mappings);
        var record = CreateRecord(mappings.Select(mapping => (mapping.ColumnIndex, mapping.Value)).ToArray());
        await new DocxGenerator().GenerateAsync(definition, record, outputPath);
        return outputPath;
    }

    private static TemplateDefinition CreateDefinition(
        string templatePath,
        IReadOnlyList<TestMapping> mappings,
        string worksheetName = "資料",
        int headerRowNumber = 1,
        int markerRowNumber = 99) => new()
    {
        TemplateName = "測試模板",
        WordTemplatePath = templatePath,
        WorksheetName = worksheetName,
        HeaderRowNumber = headerRowNumber,
        MarkerRowNumber = markerRowNumber,
        FieldMappings = mappings.Select(mapping => new FieldMapping
        {
            PlaceholderName = mapping.Name,
            ColumnIndex = mapping.ColumnIndex,
            HeaderName = mapping.HeaderName
        }).ToArray()
    };

    private static ExcelRecord CreateRecord(params (int ColumnIndex, string? Value)[] values) => new()
    {
        RowNumber = 2,
        Values = values.ToDictionary(
            value => $"COLUMN_{value.ColumnIndex}",
            value => value.Value),
        ColumnValues = values.ToDictionary(value => value.ColumnIndex, value => value.Value)
    };

    private static TestMapping Mapping(string name, int columnIndex, string? value) =>
        new(name, columnIndex, value, null);

    private static string CreateDocument(string directory, params OpenXmlElement[] bodyElements)
    {
        var path = Path.Combine(directory, $"template-{Guid.NewGuid():N}.docx");
        using var document = WordprocessingDocument.Create(
            path,
            WordprocessingDocumentType.Document);
        var mainPart = document.AddMainDocumentPart();
        mainPart.Document = new Document(new Body(bodyElements));
        mainPart.Document.Save();
        return path;
    }

    private static string ReadAllText(string path)
    {
        using var document = WordprocessingDocument.Open(path, false);
        return string.Concat(document.MainDocumentPart!.Document.Body!.Descendants<Text>()
            .Select(text => text.Text));
    }

    private static string CreateWorkbook(string directory, Action<IXLWorksheet, XLWorkbook> configure)
    {
        var path = Path.Combine(directory, $"workbook-{Guid.NewGuid():N}.xlsx");
        using var workbook = new XLWorkbook();
        var worksheet = workbook.AddWorksheet("資料");
        configure(worksheet, workbook);
        workbook.SaveAs(path);
        return path;
    }

    private static void SetRow(IXLWorksheet worksheet, int rowNumber, params string[] values)
    {
        for (var index = 0; index < values.Length; index++)
        {
            worksheet.Cell(rowNumber, index + 1).Value = values[index];
        }
    }

    private sealed record TestMapping(string Name, int ColumnIndex, string? Value, string? HeaderName);

    private sealed class TemporaryWorkspace : IDisposable
    {
        public TemporaryWorkspace()
        {
            Path = System.IO.Path.Combine(
                System.IO.Path.GetTempPath(),
                $"xlsx-docx-phase2-{Guid.NewGuid():N}");
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
