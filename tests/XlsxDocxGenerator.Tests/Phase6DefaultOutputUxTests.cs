using ClosedXML.Excel;
using XlsxDocxGenerator.Models;
using XlsxDocxGenerator.Services.Diagnostics;
using XlsxDocxGenerator.Services.Excel;
using XlsxDocxGenerator.Services.Generation;
using XlsxDocxGenerator.Services.Settings;
using XlsxDocxGenerator.Services.Templates;
using XlsxDocxGenerator.Services.Word;
using XlsxDocxGenerator.ViewModels;

namespace XlsxDocxGenerator.Tests;

public sealed class Phase6DefaultOutputUxTests
{
    [Fact]
    public async Task ExcelSelectionPrefillsOutputDirectoryNextToWorkbook()
    {
        using var workspace = new TemporaryWorkspace();
        var workbookPath = CreateWorkbook(workspace.Path, "測試資料.xlsx");
        var viewModel = CreateViewModel(workspace.Path);

        await viewModel.LoadExcelAsync(workbookPath);

        Assert.Equal(Path.Combine(workspace.Path, "ReceiptXcel輸出"), viewModel.OutputDirectory);
    }

    [Fact]
    public async Task AutomaticOutputDirectoryIsNotCreatedDuringExcelSelection()
    {
        using var workspace = new TemporaryWorkspace();
        var workbookPath = CreateWorkbook(workspace.Path, "資料.xlsx");
        var outputDirectory = Path.Combine(workspace.Path, "ReceiptXcel輸出");
        var viewModel = CreateViewModel(workspace.Path);

        await viewModel.LoadExcelAsync(workbookPath);

        Assert.False(Directory.Exists(outputDirectory));
    }

    [Fact]
    public async Task GenerationBoundaryCreatesMissingOutputDirectory()
    {
        using var workspace = new TemporaryWorkspace();
        var outputDirectory = Path.Combine(workspace.Path, "ReceiptXcel輸出");
        var service = new GenerationService(new StubExcelReader(), new WritingDocxGenerator());
        var result = await service.GenerateBatchAsync(
            "ignored.xlsx",
            CreateTemplate("row-{{ROW}}"),
            new SelectionRule.ExcelRows([2]),
            outputDirectory);

        Assert.True(result.SuccessCount == 1);
        Assert.True(Directory.Exists(outputDirectory));
        Assert.Single(Directory.GetFiles(outputDirectory, "*.docx"));
    }

    [Fact]
    public async Task UserDefinedOutputDirectorySurvivesChangingExcel()
    {
        using var workspace = new TemporaryWorkspace();
        var firstWorkbook = CreateWorkbook(workspace.Path, "first.xlsx");
        var secondWorkbook = CreateWorkbook(workspace.Path, "second.xlsx");
        var customDirectory = Path.Combine(workspace.Path, "CustomOutput");
        var viewModel = CreateViewModel(workspace.Path);
        viewModel.OutputDirectory = customDirectory;

        await viewModel.LoadExcelAsync(firstWorkbook);
        await viewModel.LoadExcelAsync(secondWorkbook);

        Assert.Equal(customDirectory, viewModel.OutputDirectory);
    }

    [Fact]
    public async Task UserDefinedOutputDirectorySurvivesSelectionAndPreviewChanges()
    {
        using var workspace = new TemporaryWorkspace();
        var workbookPath = CreateWorkbook(workspace.Path, "資料.xlsx");
        var customDirectory = Path.Combine(workspace.Path, "CustomOutput");
        var viewModel = CreateViewModel(workspace.Path);
        viewModel.OutputDirectory = customDirectory;

        await viewModel.LoadExcelAsync(workbookPath);
        viewModel.SelectedMode = viewModel.SelectionModes[1];
        viewModel.WordTemplatePath = Path.Combine(workspace.Path, "通知書.docx");
        await viewModel.RefreshSelectionPreviewAsync();

        Assert.Equal(customDirectory, viewModel.OutputDirectory);
    }

    [Fact]
    public void WordTemplateStemCreatesRowPatternWithoutExtension()
    {
        using var workspace = new TemporaryWorkspace();
        var viewModel = CreateViewModel(workspace.Path);

        viewModel.WordTemplatePath = Path.Combine(workspace.Path, "員工資料表.docx");

        Assert.Equal("員工資料表_{{ROW}}", viewModel.OutputPattern);
        Assert.DoesNotContain(".docx", viewModel.OutputPattern, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void UserDefinedPatternSurvivesChangingWordTemplate()
    {
        using var workspace = new TemporaryWorkspace();
        var viewModel = CreateViewModel(workspace.Path);
        viewModel.WordTemplatePath = Path.Combine(workspace.Path, "員工資料表.docx");
        viewModel.OutputPattern = "{{ID}}_{{NAME}}";

        viewModel.WordTemplatePath = Path.Combine(workspace.Path, "通知書.docx");

        Assert.Equal("{{ID}}_{{NAME}}", viewModel.OutputPattern);
    }

    [Fact]
    public async Task PersistedTemplatePatternWinsOverCurrentAutomaticOrUserPattern()
    {
        using var workspace = new TemporaryWorkspace();
        var configPath = Path.Combine(workspace.Path, "template.docxcel.json");
        var definition = CreateTemplate("saved-{{ROW}}") with
        {
            WordTemplatePath = Path.Combine(workspace.Path, "通知書.docx")
        };
        new TemplateService(new ExcelReader(), new MarkerParser()).SaveTemplate(configPath, definition);
        var viewModel = CreateViewModel(workspace.Path);
        viewModel.OutputPattern = "{{ID}}_{{NAME}}";

        await viewModel.LoadTemplateAsync(configPath);

        Assert.Equal("saved-{{ROW}}", viewModel.OutputPattern);
        Assert.Equal("saved-{{ROW}}", viewModel.CurrentTemplateDefinition!.OutputFileNamePattern);
    }

    [Fact]
    public async Task UserPatternSurvivesTemplateWithoutPersistedPattern()
    {
        using var workspace = new TemporaryWorkspace();
        var configPath = Path.Combine(workspace.Path, "template.docxcel.json");
        var definition = CreateTemplate(null) with
        {
            WordTemplatePath = Path.Combine(workspace.Path, "通知書.docx")
        };
        new TemplateService(new ExcelReader(), new MarkerParser()).SaveTemplate(configPath, definition);
        var viewModel = CreateViewModel(workspace.Path);
        viewModel.OutputPattern = "{{ID}}_{{NAME}}";

        await viewModel.LoadTemplateAsync(configPath);

        Assert.Equal("{{ID}}_{{NAME}}", viewModel.OutputPattern);
    }

    [Theory]
    [InlineData("2026通知書.docx", "2026通知書_{{ROW}}")]
    [InlineData("Example Form.docx", "Example Form_{{ROW}}")]
    public void UnicodeAndSpaceWordStemsCreateExpectedPatterns(string filename, string expected)
    {
        using var workspace = new TemporaryWorkspace();
        var viewModel = CreateViewModel(workspace.Path);

        viewModel.WordTemplatePath = Path.Combine(workspace.Path, filename);

        Assert.Equal(expected, viewModel.OutputPattern);
    }

    [Fact]
    public async Task SettingsRestoreTakesPriorityOverAutomaticDefaults()
    {
        using var workspace = new TemporaryWorkspace();
        var settingsPath = Path.Combine(workspace.Path, "settings.json");
        var customDirectory = Path.Combine(workspace.Path, "RememberedOutput");
        var settings = new SettingsService(settingsPath);
        settings.Save(new ApplicationSettings
        {
            LastOutputDirectory = customDirectory,
            LastOutputFilenamePattern = "remembered-{{ROW}}"
        });
        var workbookPath = CreateWorkbook(workspace.Path, "資料.xlsx");
        var viewModel = CreateViewModel(workspace.Path);

        await viewModel.LoadExcelAsync(workbookPath);
        viewModel.WordTemplatePath = Path.Combine(workspace.Path, "通知書.docx");

        Assert.Equal(customDirectory, viewModel.OutputDirectory);
        Assert.Equal("remembered-{{ROW}}", viewModel.OutputPattern);
    }

    [Fact]
    public async Task AutomaticDefaultsAreNotPersistedAsUserPreferences()
    {
        using var workspace = new TemporaryWorkspace();
        var settingsPath = Path.Combine(workspace.Path, "settings.json");
        var settings = new SettingsService(settingsPath);
        var viewModel = CreateViewModel(workspace.Path);
        var workbookPath = CreateWorkbook(workspace.Path, "資料.xlsx");

        await viewModel.LoadExcelAsync(workbookPath);
        viewModel.WordTemplatePath = Path.Combine(workspace.Path, "通知書.docx");

        var loaded = settings.Load();
        Assert.Null(loaded.LastOutputDirectory);
        Assert.Null(loaded.LastOutputFilenamePattern);
    }

    private static MainWindowViewModel CreateViewModel(string directory) =>
        new(
            new SettingsService(Path.Combine(directory, "settings.json")),
            new DiagnosticsLogger(Path.Combine(directory, "Logs")));

    private static string CreateWorkbook(string directory, string filename)
    {
        var path = Path.Combine(directory, filename);
        using var workbook = new XLWorkbook();
        var worksheet = workbook.AddWorksheet("Sheet1");
        worksheet.Cell(1, 1).Value = "NAME";
        worksheet.Cell(2, 1).Value = "Alice";
        workbook.SaveAs(path);
        return path;
    }

    private static TemplateDefinition CreateTemplate(string? outputPattern) => new()
    {
        TemplateName = "UX template",
        WordTemplatePath = "template.docx",
        PreferredWorksheetName = "Sheet1",
        HeaderRowNumber = 1,
        OutputFileNamePattern = outputPattern,
        FieldMappings = [new FieldMapping { PlaceholderName = "NAME", ColumnIndex = 1 }]
    };

    private sealed class StubExcelReader : IExcelReader
    {
        private static readonly ExcelWorksheetData Worksheet = new()
        {
            WorksheetName = "Sheet1",
            HeaderRowNumber = 1,
            FirstUsedRowNumber = 1,
            LastUsedRowNumber = 2,
            Headers = ["NAME"],
            Rows = [new ExcelRowData
            {
                RowNumber = 2,
                Cells = new Dictionary<int, string?> { [1] = "Alice" }
            }]
        };

        public Task<IReadOnlyList<string>> GetWorksheetNamesAsync(
            string workbookPath,
            CancellationToken cancellationToken = default) =>
            Task.FromResult<IReadOnlyList<string>>(["Sheet1"]);

        public Task<ExcelWorksheetData> ReadAsync(
            string workbookPath,
            string worksheetName,
            int headerRowNumber = 1,
            CancellationToken cancellationToken = default) =>
            Task.FromResult(Worksheet);

        public Task<ExcelRecord> ReadRecordAsync(
            string workbookPath,
            TemplateDefinition template,
            int rowNumber,
            CancellationToken cancellationToken = default) =>
            throw new NotSupportedException();
    }

    private sealed class WritingDocxGenerator : IDocxGenerator
    {
        public Task ValidateTemplateAsync(
            TemplateDefinition template,
            CancellationToken cancellationToken = default) =>
            Task.CompletedTask;

        public Task GenerateAsync(
            TemplateDefinition template,
            ExcelRecord record,
            string outputPath,
            CancellationToken cancellationToken = default)
        {
            File.WriteAllText(outputPath, record.GetValueByColumnIndex(1) ?? string.Empty);
            return Task.CompletedTask;
        }
    }

    private sealed class TemporaryWorkspace : IDisposable
    {
        public TemporaryWorkspace()
        {
            Path = System.IO.Path.Combine(
                System.IO.Path.GetTempPath(),
                $"docxcel-phase6-ux-{Guid.NewGuid():N}");
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
