using ClosedXML.Excel;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using XlsxDocxGenerator.Models;
using XlsxDocxGenerator.Services.Errors;
using XlsxDocxGenerator.Services.Excel;
using XlsxDocxGenerator.Services.Generation;
using XlsxDocxGenerator.Services.Diagnostics;
using XlsxDocxGenerator.Services.Settings;
using XlsxDocxGenerator.Services.Templates;
using XlsxDocxGenerator.ViewModels;

namespace XlsxDocxGenerator.Tests;

public sealed class Phase4BRegressionTests
{
    [Fact]
    public void CreationMarkerAtRow100DoesNotExcludeNormalRow100()
    {
        var worksheet = Worksheet(["姓名", "地址"], (100, "Alice", "Taipei"));
        var template = Template(100);
        var context = new RuntimeWorksheetSchemaValidator().Validate(worksheet, template);

        var result = Resolve(worksheet, template, new SelectionRule.Latest(1), context);

        Assert.Null(context.RuntimeMappingMarkerRowNumber);
        Assert.Equal([100], result.Records.Select(record => record.RowNumber));
    }

    [Fact]
    public void ExistingRuntimeMarkerIsExcludedEvenWhenCreationRowDiffers()
    {
        var worksheet = Worksheet(["姓名", "地址"],
            (2, "Alice", "Taipei"),
            (7, "{{NAME}}", "{{ADDRESS}}"));
        var template = Template(100);
        var context = new RuntimeWorksheetSchemaValidator().Validate(worksheet, template);

        var result = Resolve(worksheet, template, new SelectionRule.ExcelRows([7]), context);

        Assert.Equal(7, context.RuntimeMappingMarkerRowNumber);
        Assert.False(result.Items.Single().IsValid);
    }

    [Fact]
    public void WorkbookWithoutRuntimeMarkerKeepsAllNormalRows()
    {
        var worksheet = Worksheet(["姓名", "地址"],
            (2, "Alice", "Taipei"),
            (3, "Bob", "Kaohsiung"));
        var template = Template(100);
        var context = new RuntimeWorksheetSchemaValidator().Validate(worksheet, template);

        var result = Resolve(worksheet, template, new SelectionRule.Latest(5), context);

        Assert.Null(context.RuntimeMappingMarkerRowNumber);
        Assert.Equal([2, 3], result.Records.Select(record => record.RowNumber));
    }

    [Theory]
    [InlineData("{{NAME}}", "normal text")]
    [InlineData("{{NAME}}", "")]
    [InlineData("{{UNKNOWN}}", "{{ADDRESS}}")]
    [InlineData("{{NAME}}", "ordinary text")]
    public void StrictRuntimeMarkerDetectionDoesNotSkipPartialOrAccidentalRows(
        string nameCell,
        string addressCell)
    {
        var worksheet = Worksheet(["姓名", "地址"], (2, nameCell, addressCell));
        var template = Template(100);

        var context = new RuntimeWorksheetSchemaValidator().Validate(worksheet, template);

        Assert.Null(context.RuntimeMappingMarkerRowNumber);
    }

    [Fact]
    public void OneFieldPlaceholderIsNotAutoDetectedAsRuntimeMarker()
    {
        var worksheet = new ExcelWorksheetData
        {
            WorksheetName = "資料",
            HeaderRowNumber = 1,
            FirstUsedRowNumber = 1,
            LastUsedRowNumber = 2,
            Headers = ["姓名"],
            Rows =
            [
                new ExcelRowData
                {
                    RowNumber = 2,
                    Cells = new Dictionary<int, string?> { [1] = "{{NAME}}" }
                }
            ]
        };
        var template = Template(null) with
        {
            FieldMappings = [Mapping("NAME", 1, "姓名")]
        };

        var context = new RuntimeWorksheetSchemaValidator().Validate(worksheet, template);

        Assert.Null(context.RuntimeMappingMarkerRowNumber);
    }

    [Fact]
    public void LatestAndMarkerSelectionDoNotIncludeRuntimeMarker()
    {
        var worksheet = Worksheet(["姓名", "地址", "OUTPUT"],
            (2, "Alice", "Taipei", "A"),
            (3, "Bob", "Kaohsiung", "B"),
            (8, "{{NAME}}", "{{ADDRESS}}", ""));
        var template = Template(100);
        var context = new RuntimeWorksheetSchemaValidator().Validate(worksheet, template);

        var latest = Resolve(worksheet, template, new SelectionRule.Latest(5), context);
        var marker = Resolve(worksheet, template, new SelectionRule.Marker("OUTPUT", ["A", "B"]), context);

        Assert.Equal([2, 3], latest.Records.Select(record => record.RowNumber));
        Assert.Equal([2, 3], marker.Records.Select(record => record.RowNumber));
    }

    [Fact]
    public void TemplatePersistenceCoversValidationPolicies()
    {
        var service = TemplateService();
        var path = TempFile(".docxcel.json");
        try
        {
            var source = Template(null) with
            {
                TemplateName = "模板 測試",
                WordTemplatePath = "missing-template.docx",
                FieldMappings =
                [
                    Mapping("NAME", 1, "姓名"),
                    Mapping("ADDRESS", 2, "地址")
                ],
                OutputFileNamePattern = "報表-{{ROW}}"
            };
            service.SaveTemplate(path, source);
            var json = File.ReadAllText(path);
            var loaded = service.LoadTemplate(path);

            Assert.Contains("\"schemaVersion\": 1", json);
            Assert.Equal("模板 測試", loaded.TemplateName);
            Assert.Equal("姓名", loaded.FieldMappings[0].HeaderName);
            Assert.Equal("報表-{{ROW}}", loaded.OutputFileNamePattern);
            Assert.False(File.Exists(loaded.WordTemplatePath));
        }
        finally
        {
            Delete(path);
        }
    }

    [Fact]
    public void TemplatePersistenceRejectsMalformedAndMissingFiles()
    {
        var service = TemplateService();
        var malformed = TempFile(".docxcel.json");
        try
        {
            File.WriteAllText(malformed, "{ not json }");
            var malformedException = Assert.Throws<TemplatePersistenceException>(() => service.LoadTemplate(malformed));
            Assert.Equal(TemplatePersistenceErrorCode.InvalidTemplateFile, malformedException.Code);

            var missingException = Assert.Throws<TemplatePersistenceException>(() => service.LoadTemplate(malformed + ".missing"));
            Assert.Equal(TemplatePersistenceErrorCode.TemplateFileNotFound, missingException.Code);
        }
        finally
        {
            Delete(malformed);
        }
    }

    [Fact]
    public void TemplatePersistenceRejectsMissingSchemaAndInvalidMappings()
    {
        var service = TemplateService();
        var missingSchema = TempFile(".docxcel.json");
        try
        {
            File.WriteAllText(missingSchema, "{\"templateName\":\"T\"}");
            var missingException = Assert.Throws<TemplatePersistenceException>(() => service.LoadTemplate(missingSchema));
            Assert.Equal(TemplatePersistenceErrorCode.MissingRequiredField, missingException.Code);

            foreach (var mapping in new[]
            {
                Mapping("NAME", 0, "姓名"),
                Mapping("NAME", -1, "姓名"),
                Mapping("BAD NAME", 1, "姓名"),
                Mapping("NAME", 1, "姓名")
            })
            {
                var definition = Template(null) with { FieldMappings = [mapping, mapping] };
                var exception = Assert.Throws<TemplatePersistenceException>(() => service.SaveTemplate(missingSchema, definition));
                Assert.Equal(TemplatePersistenceErrorCode.InvalidFieldMapping, exception.Code);
            }
        }
        finally
        {
            Delete(missingSchema);
        }
    }

    [Fact]
    public void EmptyMappingCollectionIsAllowedByCurrentPolicy()
    {
        var service = TemplateService();
        var path = TempFile(".docxcel.json");
        try
        {
            service.SaveTemplate(path, Template(null) with { FieldMappings = [] });
            Assert.Empty(service.LoadTemplate(path).FieldMappings);
        }
        finally
        {
            Delete(path);
        }
    }

    [Fact]
    public void SchemaValidationAllowsExactWhitespaceCaseAndNullHeaderMappings()
    {
        var worksheet = Worksheet([" 姓名 ", "ADDRESS"], (2, "Alice", "Taipei"));
        var template = Template(null) with
        {
            FieldMappings =
            [
                Mapping("NAME", 1, "姓名"),
                Mapping("ADDRESS", 2, "address"),
                Mapping("OPTIONAL", 3, null)
            ]
        };

        var context = new RuntimeWorksheetSchemaValidator().Validate(worksheet, template);

        Assert.Same(worksheet, context.Worksheet);
    }

    [Fact]
    public void SchemaMismatchReportsColumnExpectedAndActualHeader()
    {
        var worksheet = Worksheet(["地址", "姓名"], (2, "Taipei", "Alice"));
        var template = Template(null);

        var exception = Assert.Throws<TemplateDefinitionException>(() =>
            new RuntimeWorksheetSchemaValidator().Validate(worksheet, template));

        Assert.Equal(TemplateDefinitionErrorCode.ExcelSchemaMismatch, exception.Code);
        Assert.Contains("A", exception.UserMessage);
        Assert.Contains("姓名", exception.UserMessage);
        Assert.Contains("地址", exception.UserMessage);
    }

    [Fact]
    public void SchemaMismatchWithOneWrongColumnReportsOnlyThatColumn()
    {
        var worksheet = Worksheet(["姓名", "城市"], (2, "Alice", "Taipei"));
        var template = Template(null);

        var exception = Assert.Throws<TemplateDefinitionException>(() =>
            new RuntimeWorksheetSchemaValidator().Validate(worksheet, template));

        Assert.Contains("B", exception.UserMessage);
        Assert.Contains("地址", exception.UserMessage);
        Assert.Contains("城市", exception.UserMessage);
    }

    [Fact]
    public async Task PreferredWorksheetCanFallbackAndThenValidateSelectedSheet()
    {
        var path = CreateWorkbook("資料", ["姓名", "地址"], (2, "Alice", "Taipei"));
        try
        {
            var reader = new ExcelReader();
            var names = await reader.GetWorksheetNamesAsync(path);
            var template = Template(null) with { PreferredWorksheetName = "Sheet1" };

            Assert.DoesNotContain("Sheet1", names);
            Assert.Contains("資料", names);

            var selectedWorksheet = await reader.ReadAsync(path, "資料", 1);
            var validated = new RuntimeWorksheetSchemaValidator().Validate(
                selectedWorksheet,
                template with { PreferredWorksheetName = "資料" });
            Assert.Equal("資料", validated.Worksheet.WorksheetName);
        }
        finally
        {
            Delete(path);
        }
    }

    [Fact]
    public async Task CompatibleWorkbookCompletesTemplateLifecycleEndToEnd()
    {
        var directory = CreateDirectory();
        try
        {
            var wordPath = CreateDocument(directory, "姓名={{NAME}} 地址={{ADDRESS}}");
            var workbookA = CreateWorkbook(directory, "資料", ["姓名", "地址"],
                (2, "建立者", "台北"), (3, "{{NAME}}", "{{ADDRESS}}"));
            var workbookB = CreateWorkbook(directory, "資料", ["姓名", "地址"],
                (2, "Alice", "Kaohsiung"), (7, "{{NAME}}", "{{ADDRESS}}"));
            var configPath = Path.Combine(directory, "template.docxcel.json");
            var outputDirectory = Path.Combine(directory, "output");
            Directory.CreateDirectory(outputDirectory);

            var templateService = TemplateService();
            var created = await templateService.CreateDefinitionAsync(
                workbookA, "整合測試模板", wordPath, "資料");
            templateService.SaveTemplate(configPath, created);
            var loaded = templateService.LoadTemplate(configPath);
            var result = await new GenerationService(new ExcelReader(), new XlsxDocxGenerator.Services.Word.DocxGenerator())
                .GenerateBatchAsync(
                    workbookB,
                    loaded,
                    new SelectionRule.ExcelRows([2]),
                    outputDirectory);

            var output = Assert.Single(result.Results);
            Assert.True(output.Success);
            var generatedText = ReadBodyText(output.OutputPath!);
            Assert.Contains("Alice", generatedText);
            Assert.Contains("Kaohsiung", generatedText);
            using var reopened = WordprocessingDocument.Open(output.OutputPath!, false);
            Assert.NotNull(reopened.MainDocumentPart);
        }
        finally
        {
            Directory.Delete(directory, true);
        }
    }

    [Fact]
    public async Task IncompatibleReorderedWorkbookIsBlockedBeforeOutput()
    {
        var directory = CreateDirectory();
        try
        {
            var wordPath = CreateDocument(directory, "{{NAME}} {{ADDRESS}}");
            var workbookA = CreateWorkbook(directory, "資料", ["姓名", "地址"],
                (2, "建立者", "台北"), (3, "{{NAME}}", "{{ADDRESS}}"));
            var workbookB = CreateWorkbook(directory, "資料", ["地址", "姓名"],
                (2, "Kaohsiung", "Alice"));
            var template = await TemplateService().CreateDefinitionAsync(
                workbookA, "模板", wordPath, "資料");
            var outputDirectory = Path.Combine(directory, "output");

            var exception = await Assert.ThrowsAsync<BatchGenerationException>(() =>
                new GenerationService(new ExcelReader(), new XlsxDocxGenerator.Services.Word.DocxGenerator())
                    .GenerateBatchAsync(
                        workbookB,
                        template,
                        new SelectionRule.ExcelRows([2]),
                        outputDirectory));

            Assert.Equal(BatchFatalErrorCode.TemplateDefinitionInvalid, exception.Code);
            Assert.Empty(Directory.Exists(outputDirectory)
                ? Directory.GetFiles(outputDirectory, "*.docx")
                : []);
        }
        finally
        {
            Directory.Delete(directory, true);
        }
    }

    [Fact]
    public async Task MissingWordPathLoadsWithoutCrashAndCanBeReassigned()
    {
        var directory = CreateDirectory();
        try
        {
            var configPath = Path.Combine(directory, "template.docxcel.json");
            var validWordPath = CreateDocument(directory, "{{NAME}} {{ADDRESS}}");
            TemplateService().SaveTemplate(
                configPath,
                Template(null) with { WordTemplatePath = Path.Combine(directory, "missing.docx") });

            var viewModel = new MainWindowViewModel(
                new SettingsService(Path.Combine(directory, "settings.json")),
                new DiagnosticsLogger(Path.Combine(directory, "Logs")));
            await viewModel.LoadTemplateAsync(configPath);

            Assert.Contains("Word", viewModel.StatusText);
            Assert.False(viewModel.CanGenerateBatch);

            viewModel.WordTemplatePath = validWordPath;
            Assert.Equal(validWordPath, viewModel.CurrentTemplateDefinition!.WordTemplatePath);
            Assert.True(File.Exists(viewModel.WordTemplatePath));
        }
        finally
        {
            Directory.Delete(directory, true);
        }
    }

    private static RowSelectionResult Resolve(
        ExcelWorksheetData worksheet,
        TemplateDefinition template,
        SelectionRule rule,
        RuntimeWorksheetContext context) =>
        new RowResolver().Resolve(worksheet, template, rule, context.RuntimeMappingMarkerRowNumber);

    private static TemplateDefinition Template(int? creationMarkerRow) => new()
    {
        TemplateName = "測試模板",
        WordTemplatePath = "missing.docx",
        PreferredWorksheetName = "資料",
        HeaderRowNumber = 1,
        CreationMappingMarkerRowNumber = creationMarkerRow,
        FieldMappings =
        [
            Mapping("NAME", 1, "姓名"),
            Mapping("ADDRESS", 2, "地址")
        ]
    };

    private static FieldMapping Mapping(string name, int columnIndex, string? headerName) => new()
    {
        PlaceholderName = name,
        ColumnIndex = columnIndex,
        HeaderName = headerName
    };

    private static ExcelWorksheetData Worksheet(
        IReadOnlyList<string?> headers,
        params (int Row, string? Name, string? Address, string? Output)[] rows)
    {
        return new ExcelWorksheetData
        {
            WorksheetName = "資料",
            HeaderRowNumber = 1,
            FirstUsedRowNumber = 1,
            LastUsedRowNumber = rows.Max(row => row.Row),
            Headers = headers,
            Rows = rows.Select(row => new ExcelRowData
            {
                RowNumber = row.Row,
                Cells = new Dictionary<int, string?>
                {
                    [1] = row.Name,
                    [2] = row.Address,
                    [3] = row.Output
                }.Where(pair => !string.IsNullOrEmpty(pair.Value))
                 .ToDictionary(pair => pair.Key, pair => pair.Value)
            }).ToArray()
        };
    }

    private static ExcelWorksheetData Worksheet(
        IReadOnlyList<string?> headers,
        params (int Row, string? Name, string? Address)[] rows) =>
        Worksheet(
            headers,
            rows.Select(row => (row.Row, row.Name, row.Address, (string?)null)).ToArray());

    private static TemplateService TemplateService() => new(new ExcelReader(), new MarkerParser());

    private static string CreateWorkbook(
        string directory,
        string sheetName,
        IReadOnlyList<string> headers,
        params (int Row, string Name, string Address)[] rows)
    {
        var path = Path.Combine(directory, $"workbook-{Guid.NewGuid():N}.xlsx");
        using var workbook = new XLWorkbook();
        var sheet = workbook.AddWorksheet(sheetName);
        for (var index = 0; index < headers.Count; index++)
        {
            sheet.Cell(1, index + 1).Value = headers[index];
        }

        foreach (var row in rows)
        {
            sheet.Cell(row.Row, 1).Value = row.Name;
            sheet.Cell(row.Row, 2).Value = row.Address;
        }

        workbook.SaveAs(path);
        return path;
    }

    private static string CreateWorkbook(
        string sheetName,
        IReadOnlyList<string> headers,
        params (int Row, string Name, string Address)[] rows)
    {
        var path = Path.Combine(Path.GetTempPath(), $"docxcel-{Guid.NewGuid():N}.xlsx");
        return CreateWorkbook(Path.GetDirectoryName(path)!, sheetName, headers, rows);
    }

    private static string CreateDocument(string directory, string text)
    {
        var path = Path.Combine(directory, $"template-{Guid.NewGuid():N}.docx");
        using var document = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document);
        var main = document.AddMainDocumentPart();
        main.Document = new Document(new Body(new Paragraph(new Run(new Text(text)))));
        main.Document.Save();
        return path;
    }

    private static string ReadBodyText(string path)
    {
        using var document = WordprocessingDocument.Open(path, false);
        return document.MainDocumentPart!.Document.Body!.InnerText;
    }

    private static string TempFile(string extension) =>
        Path.Combine(Path.GetTempPath(), $"docxcel-{Guid.NewGuid():N}{extension}");

    private static string CreateDirectory()
    {
        var path = Path.Combine(Path.GetTempPath(), $"docxcel-{Guid.NewGuid():N}");
        Directory.CreateDirectory(path);
        return path;
    }

    private static void Delete(string path)
    {
        if (File.Exists(path)) File.Delete(path);
    }
}
