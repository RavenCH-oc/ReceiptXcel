using XlsxDocxGenerator.Models;
using XlsxDocxGenerator.Services.Errors;
using XlsxDocxGenerator.Services.Excel;
using XlsxDocxGenerator.Services.Generation;
using XlsxDocxGenerator.Services.Diagnostics;
using XlsxDocxGenerator.Services.Settings;
using XlsxDocxGenerator.Services.Templates;
using XlsxDocxGenerator.Services.Word;
using XlsxDocxGenerator.ViewModels;

namespace XlsxDocxGenerator.Tests;

public sealed class Phase4CRegressionTests
{
    [Fact]
    public void SingleMappingNormalPlaceholderDataIsNotAutomaticallyExcluded()
    {
        var worksheet = SingleWorksheet((20, "{{NAME}}"));
        var template = SingleTemplate(150);

        var context = new RuntimeWorksheetSchemaValidator().Validate(worksheet, template);
        var result = Resolve(worksheet, template, new SelectionRule.Latest(1), context);

        Assert.Null(context.RuntimeMappingMarkerRowNumber);
        Assert.Equal([20], result.Records.Select(record => record.RowNumber));
    }

    [Fact]
    public void SingleMappingRuntimePlaceholderDoesNotTriggerAutomaticDetection()
    {
        var worksheet = SingleWorksheet((20, "Alice"), (150, "{{NAME}}"));
        var template = SingleTemplate(20);

        var context = new RuntimeWorksheetSchemaValidator().Validate(worksheet, template);
        var result = Resolve(worksheet, template, new SelectionRule.Latest(5), context);

        Assert.Null(context.RuntimeMappingMarkerRowNumber);
        Assert.Equal([20, 150], result.Records.Select(record => record.RowNumber));
    }

    [Fact]
    public void SingleMappingExplicitRuntimeMarkerIsValidatedAndExcluded()
    {
        var worksheet = SingleWorksheet((2, "Alice"), (150, "{{NAME}}"));
        var template = SingleTemplate(20);
        var context = new RuntimeWorksheetSchemaValidator().Validate(worksheet, template, 150);

        var result = Resolve(worksheet, template, new SelectionRule.ExcelRows([150]), context);

        Assert.Equal(150, context.RuntimeMappingMarkerRowNumber);
        Assert.False(result.Items.Single().IsValid);
    }

    [Fact]
    public void SingleMappingExplicitRuntimeMarkerRejectsNonMatchingRow()
    {
        var worksheet = SingleWorksheet((150, "王小明"));

        var exception = Assert.Throws<TemplateDefinitionException>(() =>
            new RuntimeWorksheetSchemaValidator().Validate(worksheet, SingleTemplate(20), 150));

        Assert.Equal(TemplateDefinitionErrorCode.InvalidRuntimeMappingMarker, exception.Code);
    }

    [Fact]
    public void SingleMappingExplicitMarkerIsExcludedFromLatestSelection()
    {
        var worksheet = SingleWorksheet((2, "Alice"), (150, "{{NAME}}"));
        var template = SingleTemplate(null);
        var context = new RuntimeWorksheetSchemaValidator().Validate(worksheet, template, 150);

        var result = Resolve(worksheet, template, new SelectionRule.Latest(5), context);

        Assert.Equal([2], result.Records.Select(record => record.RowNumber));
    }

    [Fact]
    public void SingleMappingExplicitMarkerIsExcludedFromMarkerSelection()
    {
        var worksheet = SingleWorksheetWithOutput(
            (2, "Alice", "A"),
            (150, "{{NAME}}", ""));
        var template = SingleTemplate(null);
        var context = new RuntimeWorksheetSchemaValidator().Validate(worksheet, template, 150);

        var result = Resolve(
            worksheet,
            template,
            new SelectionRule.Marker("OUTPUT", ["A"]),
            context);

        Assert.Equal([2], result.Records.Select(record => record.RowNumber));
    }

    [Fact]
    public void SingleMappingExplicitMarkerCannotBeSelectedAsExplicitRow()
    {
        var worksheet = SingleWorksheet((150, "{{NAME}}"));
        var template = SingleTemplate(null);
        var context = new RuntimeWorksheetSchemaValidator().Validate(worksheet, template, 150);

        var item = Resolve(
            worksheet,
            template,
            new SelectionRule.ExcelRows([150]),
            context).Items.Single();

        Assert.False(item.IsValid);
        Assert.Equal(GenerationErrorCode.InvalidExcelDataRow, item.ErrorCode);
    }

    [Fact]
    public void MultiMappingAutomaticDetectionStillExcludesCompleteMarker()
    {
        var worksheet = MultiWorksheet(
            (2, "Alice", "Taipei"),
            (150, "{{NAME}}", "{{ADDRESS}}"));
        var template = MultiTemplate(20);

        var context = new RuntimeWorksheetSchemaValidator().Validate(worksheet, template);
        var result = Resolve(worksheet, template, new SelectionRule.Latest(5), context);

        Assert.Equal(150, context.RuntimeMappingMarkerRowNumber);
        Assert.Equal([2], result.Records.Select(record => record.RowNumber));
    }

    [Fact]
    public void MultiMappingExplicitMarkerOverridesAutomaticAmbiguity()
    {
        var worksheet = MultiWorksheet(
            (2, "Alice", "Taipei"),
            (7, "{{NAME}}", "{{ADDRESS}}"),
            (8, "{{NAME}}", "{{ADDRESS}}"),
            (150, "{{NAME}}", "{{ADDRESS}}"));
        var template = MultiTemplate(20);

        var context = new RuntimeWorksheetSchemaValidator().Validate(worksheet, template, 150);

        Assert.Equal(150, context.RuntimeMappingMarkerRowNumber);
        var result = Resolve(worksheet, template, new SelectionRule.ExcelRows([7, 8, 150]), context);
        Assert.True(result.Items[0].IsValid);
        Assert.True(result.Items[1].IsValid);
        Assert.False(result.Items[2].IsValid);
    }

    [Fact]
    public void CreationMarkerMetadataNeverBecomesRuntimeExplicitMarker()
    {
        var worksheet = SingleWorksheet((150, "{{NAME}}"));
        var template = SingleTemplate(150);

        var context = new RuntimeWorksheetSchemaValidator().Validate(worksheet, template);

        Assert.Null(context.RuntimeMappingMarkerRowNumber);
    }

    [Fact]
    public async Task GenerationServicePassesExplicitRuntimeMarkerToSelection()
    {
        var worksheet = SingleWorksheet((2, "Alice"), (150, "{{NAME}}"));
        var service = new GenerationService(
            new StubExcelReader(worksheet),
            new NoOpDocxGenerator());

        var result = await service.PreviewSelectionAsync(
            "ignored.xlsx",
            SingleTemplate(null),
            new SelectionRule.Latest(5),
            explicitRuntimeMappingMarkerRowNumber: 150);

        Assert.Equal([2], result.Records.Select(record => record.RowNumber));
    }

    [Fact]
    public async Task ViewModelShowsSingleMappingSafetyInstructionWithoutPersistingRuntimeState()
    {
        var directory = Path.Combine(Path.GetTempPath(), $"docxcel-phase4c-{Guid.NewGuid():N}");
        Directory.CreateDirectory(directory);
        var path = Path.Combine(directory, "template.docxcel.json");
        try
        {
            var definition = SingleTemplate(null) with { WordTemplatePath = "template.docx" };
            var service = new TemplateService(new ExcelReader(), new MarkerParser());
            service.SaveTemplate(path, definition);

            var viewModel = new MainWindowViewModel(
                new SettingsService(Path.Combine(directory, "settings.json")),
                new DiagnosticsLogger(Path.Combine(directory, "Logs")));
            await viewModel.LoadTemplateAsync(path);

            Assert.True(viewModel.HasSingleMappingTemplate);
            Assert.Contains("無法安全自動辨識", viewModel.RuntimeMarkerInstructionText);
            Assert.Equal(string.Empty, viewModel.RuntimeMarkerRowText);
            Assert.DoesNotContain("runtimeMappingMarker", File.ReadAllText(path));
        }
        finally
        {
            if (Directory.Exists(directory)) Directory.Delete(directory, true);
        }
    }

    private static RowSelectionResult Resolve(
        ExcelWorksheetData worksheet,
        TemplateDefinition template,
        SelectionRule rule,
        RuntimeWorksheetContext context) =>
        new RowResolver().Resolve(worksheet, template, rule, context.RuntimeMappingMarkerRowNumber);

    private static TemplateDefinition SingleTemplate(int? creationMarkerRow) => new()
    {
        TemplateName = "Single mapping",
        WordTemplatePath = "template.docx",
        PreferredWorksheetName = "資料",
        HeaderRowNumber = 1,
        CreationMappingMarkerRowNumber = creationMarkerRow,
        FieldMappings = [new FieldMapping
        {
            PlaceholderName = "NAME",
            ColumnIndex = 1,
            HeaderName = "姓名"
        }]
    };

    private static TemplateDefinition MultiTemplate(int? creationMarkerRow) => new()
    {
        TemplateName = "Multi mapping",
        WordTemplatePath = "template.docx",
        PreferredWorksheetName = "資料",
        HeaderRowNumber = 1,
        CreationMappingMarkerRowNumber = creationMarkerRow,
        FieldMappings =
        [
            new FieldMapping { PlaceholderName = "NAME", ColumnIndex = 1, HeaderName = "姓名" },
            new FieldMapping { PlaceholderName = "ADDRESS", ColumnIndex = 2, HeaderName = "地址" }
        ]
    };

    private static ExcelWorksheetData SingleWorksheet(params (int Row, string Value)[] rows) => new()
    {
        WorksheetName = "資料",
        HeaderRowNumber = 1,
        FirstUsedRowNumber = 1,
        LastUsedRowNumber = rows.Max(row => row.Row),
        Headers = ["姓名"],
        Rows = rows.Select(row => new ExcelRowData
        {
            RowNumber = row.Row,
            Cells = new Dictionary<int, string?> { [1] = row.Value }
        }).ToArray()
    };

    private static ExcelWorksheetData SingleWorksheetWithOutput(
        params (int Row, string Name, string Output)[] rows) => new()
    {
        WorksheetName = "資料",
        HeaderRowNumber = 1,
        FirstUsedRowNumber = 1,
        LastUsedRowNumber = rows.Max(row => row.Row),
        Headers = ["姓名", "OUTPUT"],
        Rows = rows.Select(row => new ExcelRowData
        {
            RowNumber = row.Row,
            Cells = new Dictionary<int, string?>
            {
                [1] = row.Name,
                [2] = row.Output
            }.Where(item => !string.IsNullOrEmpty(item.Value))
             .ToDictionary(item => item.Key, item => item.Value)
        }).ToArray()
    };

    private static ExcelWorksheetData MultiWorksheet(
        params (int Row, string Name, string Address)[] rows) => new()
    {
        WorksheetName = "資料",
        HeaderRowNumber = 1,
        FirstUsedRowNumber = 1,
        LastUsedRowNumber = rows.Max(row => row.Row),
        Headers = ["姓名", "地址"],
        Rows = rows.Select(row => new ExcelRowData
        {
            RowNumber = row.Row,
            Cells = new Dictionary<int, string?>
            {
                [1] = row.Name,
                [2] = row.Address
            }.Where(item => !string.IsNullOrEmpty(item.Value))
             .ToDictionary(item => item.Key, item => item.Value)
        }).ToArray()
    };

    private sealed class StubExcelReader(ExcelWorksheetData worksheet) : IExcelReader
    {
        public Task<IReadOnlyList<string>> GetWorksheetNamesAsync(
            string workbookPath,
            CancellationToken cancellationToken = default) =>
            Task.FromResult<IReadOnlyList<string>>([worksheet.WorksheetName]);

        public Task<ExcelWorksheetData> ReadAsync(
            string workbookPath,
            string worksheetName,
            int headerRowNumber = 1,
            CancellationToken cancellationToken = default) =>
            Task.FromResult(worksheet);

        public Task<ExcelRecord> ReadRecordAsync(
            string workbookPath,
            TemplateDefinition template,
            int rowNumber,
            CancellationToken cancellationToken = default) =>
            throw new NotSupportedException();
    }

    private sealed class NoOpDocxGenerator : IDocxGenerator
    {
        public Task ValidateTemplateAsync(
            TemplateDefinition template,
            CancellationToken cancellationToken = default) =>
            Task.CompletedTask;

        public Task GenerateAsync(
            TemplateDefinition template,
            ExcelRecord record,
            string outputPath,
            CancellationToken cancellationToken = default) =>
            Task.CompletedTask;
    }
}
