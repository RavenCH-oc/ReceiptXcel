using ClosedXML.Excel;
using XlsxDocxGenerator.Models;
using XlsxDocxGenerator.Services.Diagnostics;
using XlsxDocxGenerator.Services.Excel;
using XlsxDocxGenerator.Services.Settings;
using XlsxDocxGenerator.Services.Templates;
using XlsxDocxGenerator.ViewModels;

namespace XlsxDocxGenerator.Tests;

public sealed class Phase6DGenerateWorkflowTests
{
    [Fact]
    public async Task RestoredTemplateAndExcelAutomaticallyValidateSchema()
    {
        using var workspace = new TemporaryWorkspace();
        var viewModel = CreateViewModelWithRecentTemplate(workspace.Path);
        var workbookPath = CreateWorkbook(workspace.Path, "valid.xlsx");
        viewModel.SelectedMode = viewModel.SelectionModes.Single(option => option.Mode == SelectionMode.Latest);
        viewModel.LatestCountText = "1";

        await viewModel.LoadExcelAsync(workbookPath);
        await WaitForAsync(() => viewModel.ExcelFormatStatusText == "Excel 格式符合目前模板");

        Assert.Equal("Excel 格式符合目前模板", viewModel.ExcelFormatStatusText);
    }

    [Fact]
    public async Task LatestOneAutomaticallyProducesOneSelectionPreview()
    {
        using var workspace = new TemporaryWorkspace();
        var viewModel = CreateViewModelWithRecentTemplate(workspace.Path);
        await viewModel.LoadExcelAsync(CreateWorkbook(workspace.Path, "valid.xlsx"));
        viewModel.SelectedMode = viewModel.SelectionModes.Single(option => option.Mode == SelectionMode.Latest);
        viewModel.LatestCountText = "1";

        await WaitForAsync(() => viewModel.SelectionSummaryText.Contains("符合條件：1 筆"));

        Assert.Contains("預計產生：1 個 Word", viewModel.SelectionSummaryText);
    }

    [Fact]
    public async Task ValidPreviewEnablesGenerationWhenOtherFieldsAreValid()
    {
        using var workspace = new TemporaryWorkspace();
        var outputDirectory = Path.Combine(workspace.Path, "Output");
        Directory.CreateDirectory(outputDirectory);
        var viewModel = CreateViewModelWithRecentTemplate(workspace.Path);
        viewModel.OutputDirectory = outputDirectory;
        await viewModel.LoadExcelAsync(CreateWorkbook(workspace.Path, "valid.xlsx"));
        viewModel.SelectedMode = viewModel.SelectionModes.Single(option => option.Mode == SelectionMode.Latest);
        viewModel.LatestCountText = "1";

        await WaitForAsync(() => viewModel.CanGenerateBatch);

        Assert.True(viewModel.CanGenerateBatch);
    }

    [Fact]
    public async Task ChangingExcelInvalidatesPreviousPreviewImmediately()
    {
        using var workspace = new TemporaryWorkspace();
        var viewModel = CreateViewModelWithRecentTemplate(workspace.Path);
        await PrepareLatestOneAsync(viewModel, CreateWorkbook(workspace.Path, "first.xlsx"));

        var secondLoad = viewModel.LoadExcelAsync(CreateWorkbook(workspace.Path, "second.xlsx"));

        Assert.False(viewModel.CanGenerateBatch);
        await secondLoad;
        await WaitForAsync(() => viewModel.ExcelFormatStatusText == "Excel 格式符合目前模板");
    }

    [Fact]
    public async Task ChangingWorksheetInvalidatesPreviousPreview()
    {
        using var workspace = new TemporaryWorkspace();
        var viewModel = CreateViewModelWithRecentTemplate(workspace.Path);
        await PrepareLatestOneAsync(viewModel, CreateWorkbook(workspace.Path, "workbook.xlsx", includeSecondSheet: true));

        viewModel.SelectedWorksheet = "Second";

        Assert.False(viewModel.CanGenerateBatch);
        await WaitForAsync(() => viewModel.SelectionSummaryText.Contains("符合條件：1 筆"));
    }

    [Fact]
    public async Task ChangingTemplateInvalidatesPreviousPreview()
    {
        using var workspace = new TemporaryWorkspace();
        var viewModel = CreateViewModelWithRecentTemplate(workspace.Path);
        await PrepareLatestOneAsync(viewModel, CreateWorkbook(workspace.Path, "workbook.xlsx"));
        var replacementPath = CreateAlternativeTemplate(workspace.Path);

        var loadTask = viewModel.LoadTemplateAsync(replacementPath);

        Assert.False(viewModel.CanGenerateBatch);
        await loadTask;
        await WaitForAsync(() => viewModel.SelectionSummaryText.Contains("符合條件：1 筆"));
    }

    [Fact]
    public async Task ChangingSelectionModeRebuildsPreview()
    {
        using var workspace = new TemporaryWorkspace();
        var viewModel = CreateViewModelWithRecentTemplate(workspace.Path);
        await PrepareLatestOneAsync(viewModel, CreateWorkbook(workspace.Path, "workbook.xlsx"));

        viewModel.SelectedMode = viewModel.SelectionModes.Single(option => option.Mode == SelectionMode.ExcelRows);
        viewModel.RowExpressionText = "2";

        Assert.False(viewModel.CanGenerateBatch);
        await WaitForAsync(() => viewModel.SelectionSummaryText.Contains("符合條件：1 筆"));
    }

    [Fact]
    public async Task ChangingLatestCountRebuildsPreview()
    {
        using var workspace = new TemporaryWorkspace();
        var viewModel = CreateViewModelWithRecentTemplate(workspace.Path);
        await PrepareLatestOneAsync(viewModel, CreateWorkbook(workspace.Path, "workbook.xlsx"));

        viewModel.LatestCountText = "2";

        Assert.False(viewModel.CanGenerateBatch);
        await WaitForAsync(() => viewModel.SelectionSummaryText.Contains("符合條件：2 筆"));
    }

    [Fact]
    public async Task ChangingExcelRowExpressionRebuildsPreview()
    {
        using var workspace = new TemporaryWorkspace();
        var viewModel = CreateViewModelWithRecentTemplate(workspace.Path);
        await PrepareLatestOneAsync(viewModel, CreateWorkbook(workspace.Path, "workbook.xlsx"));
        viewModel.SelectedMode = viewModel.SelectionModes.Single(option => option.Mode == SelectionMode.ExcelRows);
        viewModel.RowExpressionText = "2";
        await WaitForAsync(() => viewModel.SelectionSummaryText.Contains("符合條件：1 筆"));

        viewModel.RowExpressionText = "3";

        Assert.False(viewModel.CanGenerateBatch);
        await WaitForAsync(() => viewModel.SelectionSummaryText.Contains("符合條件：1 筆"));
    }

    [Fact]
    public async Task ChangingMarkerTokenRebuildsPreview()
    {
        using var workspace = new TemporaryWorkspace();
        var viewModel = CreateViewModelWithRecentTemplate(workspace.Path);
        await viewModel.LoadExcelAsync(CreateWorkbook(workspace.Path, "workbook.xlsx"));
        viewModel.SelectedMode = viewModel.SelectionModes.Single(option => option.Mode == SelectionMode.Marker);
        viewModel.SelectedMarkerColumn = "OUTPUT";
        viewModel.MarkerTokensText = "A";
        await WaitForAsync(() => viewModel.SelectionSummaryText.Contains("符合條件：1 筆"));

        viewModel.MarkerTokensText = "B";

        Assert.False(viewModel.CanGenerateBatch);
        await WaitForAsync(() => viewModel.SelectionSummaryText.Contains("符合條件：1 筆"));
    }

    [Fact]
    public async Task SchemaMismatchDisablesGeneration()
    {
        using var workspace = new TemporaryWorkspace();
        var viewModel = CreateViewModelWithRecentTemplate(workspace.Path);
        viewModel.SelectedMode = viewModel.SelectionModes.Single(option => option.Mode == SelectionMode.Latest);
        viewModel.LatestCountText = "1";

        await viewModel.LoadExcelAsync(CreateWorkbook(workspace.Path, "mismatch.xlsx", schemaMismatch: true));
        await WaitForAsync(() => viewModel.ExcelFormatStatusText.Contains("不符"));

        Assert.False(viewModel.CanGenerateBatch);
        Assert.Contains("Excel 欄位結構與模板不一致", viewModel.StatusText);
    }

    [Fact]
    public async Task ZeroMatchesDisablesGenerationWithFriendlyStatus()
    {
        using var workspace = new TemporaryWorkspace();
        var viewModel = CreateViewModelWithRecentTemplate(workspace.Path);
        await viewModel.LoadExcelAsync(CreateWorkbook(workspace.Path, "workbook.xlsx"));
        viewModel.SelectedMode = viewModel.SelectionModes.Single(option => option.Mode == SelectionMode.Marker);
        viewModel.SelectedMarkerColumn = "OUTPUT";
        viewModel.MarkerTokensText = "NO_MATCH";

        await WaitForAsync(() => viewModel.SelectionSummaryText.Contains("符合條件：0 筆"));

        Assert.False(viewModel.CanGenerateBatch);
        Assert.Contains("沒有符合條件的資料", viewModel.StatusText);
    }

    [Fact]
    public async Task StalePreviewCannotEnableGenerationAfterExcelChanges()
    {
        using var workspace = new TemporaryWorkspace();
        var viewModel = CreateViewModelWithRecentTemplate(workspace.Path);
        await PrepareLatestOneAsync(viewModel, CreateWorkbook(workspace.Path, "first.xlsx"));

        var loadTask = viewModel.LoadExcelAsync(CreateWorkbook(workspace.Path, "mismatch.xlsx", schemaMismatch: true));
        Assert.False(viewModel.CanGenerateBatch);
        await loadTask;
        await WaitForAsync(() => viewModel.ExcelFormatStatusText.Contains("不符"));

        Assert.False(viewModel.CanGenerateBatch);
    }

    [Fact]
    public async Task RestoredTemplateDoesNotRequestCreationMapping()
    {
        using var workspace = new TemporaryWorkspace();
        var viewModel = CreateViewModelWithRecentTemplate(workspace.Path);
        await PrepareLatestOneAsync(viewModel, CreateWorkbook(workspace.Path, "workbook.xlsx"));

        Assert.DoesNotContain("建立欄位對應", viewModel.StatusText);
        Assert.NotEmpty(viewModel.MappingPreview);
    }

    [Fact]
    public async Task InvalidIntermediateSelectionInputDoesNotThrow()
    {
        using var workspace = new TemporaryWorkspace();
        var viewModel = CreateViewModelWithRecentTemplate(workspace.Path);
        await viewModel.LoadExcelAsync(CreateWorkbook(workspace.Path, "workbook.xlsx"));
        viewModel.SelectedMode = viewModel.SelectionModes.Single(option => option.Mode == SelectionMode.ExcelRows);

        var exception = Record.Exception(() => viewModel.RowExpressionText = "10-");
        await WaitForAsync(() => viewModel.StatusText.Contains("無法解析"));

        Assert.Null(exception);
        Assert.False(viewModel.CanGenerateBatch);
    }

    [Fact]
    public void MainWindowStartupBindingsRemainReadonly()
    {
        var xaml = File.ReadAllText(Path.Combine(FindSolutionDirectory(), "src", "XlsxDocxGenerator", "MainWindow.xaml"));

        Assert.Contains("Maximum=\"{Binding ProgressMaximum, Mode=OneWay}\"", xaml);
        Assert.Contains("Value=\"{Binding ProgressValue, Mode=OneWay}\"", xaml);
        Assert.DoesNotContain("Value=\"{Binding ProgressValue}\"", xaml);
    }

    [Fact]
    public void OutputFilenameHelperIsGroupedWithOutputFilenameField()
    {
        var xaml = File.ReadAllText(Path.Combine(FindSolutionDirectory(), "src", "XlsxDocxGenerator", "MainWindow.xaml"));
        var labelIndex = xaml.IndexOf("Text=\"輸出檔名\"", StringComparison.Ordinal);
        var helperIndex = xaml.IndexOf("Text=\"可使用 {{ID}}、{{NAME}} 等欄位代號\"", StringComparison.Ordinal);

        Assert.True(labelIndex >= 0);
        Assert.True(helperIndex > labelIndex);
        Assert.Contains("FontSize=\"13\"", xaml[helperIndex..(helperIndex + 180)]);
    }

    private static async Task PrepareLatestOneAsync(MainWindowViewModel viewModel, string workbookPath)
    {
        await viewModel.LoadExcelAsync(workbookPath);
        viewModel.SelectedMode = viewModel.SelectionModes.Single(option => option.Mode == SelectionMode.Latest);
        viewModel.LatestCountText = "1";
        await WaitForAsync(() => viewModel.SelectionSummaryText.Contains("符合條件：1 筆"));
    }

    private static async Task WaitForAsync(Func<bool> condition)
    {
        var deadline = DateTime.UtcNow.AddSeconds(5);
        while (!condition() && DateTime.UtcNow < deadline)
        {
            await Task.Delay(30);
        }

        Assert.True(condition(), "Preview condition timed out.");
    }

    private static MainWindowViewModel CreateViewModelWithRecentTemplate(string directory)
    {
        var wordPath = Path.Combine(directory, "通知書.docx");
        File.WriteAllText(wordPath, "placeholder");
        var templatePath = Path.Combine(directory, "通知書.docxcel.json");
        var definition = new TemplateDefinition
        {
            TemplateName = "通知書",
            WordTemplatePath = wordPath,
            PreferredWorksheetName = "Data",
            HeaderRowNumber = 1,
            FieldMappings =
            [
                new FieldMapping { PlaceholderName = "ID", ColumnIndex = 1, HeaderName = "ID" },
                new FieldMapping { PlaceholderName = "NAME", ColumnIndex = 2, HeaderName = "NAME" },
                new FieldMapping { PlaceholderName = "OUTPUT", ColumnIndex = 3, HeaderName = "OUTPUT" }
            ],
            OutputFileNamePattern = "test-{{ROW}}"
        };
        new TemplateService(new ExcelReader(), new MarkerParser()).SaveTemplate(templatePath, definition);
        new SettingsService(Path.Combine(directory, "settings.json")).Save(new ApplicationSettings
        {
            LastTemplateConfigPath = templatePath
        });
        return new MainWindowViewModel(
            new SettingsService(Path.Combine(directory, "settings.json")),
            new DiagnosticsLogger(Path.Combine(directory, "Logs")));
    }

    private static string CreateWorkbook(
        string directory,
        string filename,
        bool schemaMismatch = false,
        bool includeSecondSheet = false)
    {
        var path = Path.Combine(directory, filename);
        using var workbook = new XLWorkbook();
        AddWorksheet(workbook, "Data", schemaMismatch);
        if (includeSecondSheet)
        {
            AddWorksheet(workbook, "Second", schemaMismatch: false);
        }

        workbook.SaveAs(path);
        return path;
    }

    private static string CreateAlternativeTemplate(string directory)
    {
        var wordPath = Path.Combine(directory, "替代通知.docx");
        File.WriteAllText(wordPath, "placeholder");
        var templatePath = Path.Combine(directory, "替代通知.docxcel.json");
        var definition = new TemplateDefinition
        {
            TemplateName = "替代通知",
            WordTemplatePath = wordPath,
            PreferredWorksheetName = "Data",
            HeaderRowNumber = 1,
            FieldMappings =
            [
                new FieldMapping { PlaceholderName = "ID", ColumnIndex = 1, HeaderName = "ID" },
                new FieldMapping { PlaceholderName = "NAME", ColumnIndex = 2, HeaderName = "NAME" },
                new FieldMapping { PlaceholderName = "OUTPUT", ColumnIndex = 3, HeaderName = "OUTPUT" }
            ],
            OutputFileNamePattern = "alternative-{{ROW}}"
        };
        new TemplateService(new ExcelReader(), new MarkerParser()).SaveTemplate(templatePath, definition);
        return templatePath;
    }

    private static void AddWorksheet(XLWorkbook workbook, string name, bool schemaMismatch)
    {
        var worksheet = workbook.AddWorksheet(name);
        worksheet.Cell(1, 1).Value = "ID";
        worksheet.Cell(1, 2).Value = schemaMismatch ? "ADDRESS" : "NAME";
        worksheet.Cell(1, 3).Value = "OUTPUT";
        worksheet.Cell(2, 1).Value = "1";
        worksheet.Cell(2, 2).Value = "Alice";
        worksheet.Cell(2, 3).Value = "A";
        worksheet.Cell(3, 1).Value = "2";
        worksheet.Cell(3, 2).Value = "Bob";
        worksheet.Cell(3, 3).Value = "B";
    }

    private static string FindSolutionDirectory()
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

        throw new DirectoryNotFoundException("Could not locate the solution directory.");
    }

    private sealed class TemporaryWorkspace : IDisposable
    {
        public TemporaryWorkspace()
        {
            Path = System.IO.Path.Combine(
                System.IO.Path.GetTempPath(),
                $"docxcel-phase6d-{Guid.NewGuid():N}");
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
