using ClosedXML.Excel;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using XlsxDocxGenerator.Models;
using XlsxDocxGenerator.Services.Diagnostics;
using XlsxDocxGenerator.Services.Excel;
using XlsxDocxGenerator.Services.Settings;
using XlsxDocxGenerator.Services.Templates;
using XlsxDocxGenerator.ViewModels;

namespace XlsxDocxGenerator.Tests;

public sealed class Phase6CAutoRestoreTests
{
    [Fact]
    public void EmptyLastTemplatePathLeavesTemplateUnloaded()
    {
        using var workspace = new TemporaryWorkspace();
        var viewModel = CreateViewModel(workspace.Path);

        Assert.Null(viewModel.CurrentTemplateDefinition);
        Assert.Equal(string.Empty, viewModel.TemplateName);
        Assert.Equal(MainWindowMode.GenerateDocuments, viewModel.ActiveMode);
        Assert.DoesNotContain("最近", viewModel.StatusText);
    }

    [Fact]
    public void MissingRecentTemplateDoesNotThrowDuringStartup()
    {
        using var workspace = new TemporaryWorkspace();
        var missingPath = Path.Combine(workspace.Path, "missing.docxcel.json");
        SaveSettings(workspace.Path, new ApplicationSettings { LastTemplateConfigPath = missingPath });

        var exception = Record.Exception(() => CreateViewModel(workspace.Path));

        Assert.Null(exception);
    }

    [Fact]
    public void MissingRecentTemplateLeavesCurrentTemplateNull()
    {
        using var workspace = new TemporaryWorkspace();
        var missingPath = Path.Combine(workspace.Path, "missing.docxcel.json");
        SaveSettings(workspace.Path, new ApplicationSettings { LastTemplateConfigPath = missingPath });

        var viewModel = CreateViewModel(workspace.Path);

        Assert.Null(viewModel.CurrentTemplateDefinition);
        Assert.Contains("不存在", viewModel.StatusText);
    }

    [Fact]
    public void ValidRecentTemplateIsAutomaticallyRestored()
    {
        using var workspace = new TemporaryWorkspace();
        var definition = PrepareRecentTemplate(workspace.Path, "通知書", "通知書_{{ROW}}", wordExists: true);

        var viewModel = CreateViewModel(workspace.Path);

        Assert.NotNull(viewModel.CurrentTemplateDefinition);
        Assert.Equal(definition.TemplateName, viewModel.CurrentTemplateDefinition!.TemplateName);
        Assert.Contains("自動載入", viewModel.StatusText);
    }

    [Fact]
    public void RestoredTemplateNameIsCorrect()
    {
        using var workspace = new TemporaryWorkspace();
        PrepareRecentTemplate(workspace.Path, "員工通知", "員工通知_{{ROW}}", wordExists: true);

        var viewModel = CreateViewModel(workspace.Path);

        Assert.Equal("員工通知", viewModel.TemplateName);
        Assert.Equal("員工通知", viewModel.TemplateDisplayName);
    }

    [Fact]
    public void RestoredFieldMappingsAreCorrect()
    {
        using var workspace = new TemporaryWorkspace();
        PrepareRecentTemplate(workspace.Path, "資料模板", "資料_{{ROW}}", wordExists: true);

        var viewModel = CreateViewModel(workspace.Path);

        Assert.Equal(["ID", "NAME"], viewModel.CurrentTemplateDefinition!.FieldMappings.Select(x => x.PlaceholderName));
        Assert.Equal([1, 2], viewModel.CurrentTemplateDefinition.FieldMappings.Select(x => x.ColumnIndex));
    }

    [Fact]
    public void RestoredOutputFilenamePatternIsCorrect()
    {
        using var workspace = new TemporaryWorkspace();
        PrepareRecentTemplate(workspace.Path, "通知書", "saved-{{ID}}-{{ROW}}", wordExists: true);

        var viewModel = CreateViewModel(workspace.Path);

        Assert.Equal("saved-{{ID}}-{{ROW}}", viewModel.OutputPattern);
        Assert.Equal("saved-{{ID}}-{{ROW}}", viewModel.CurrentTemplateDefinition!.OutputFileNamePattern);
    }

    [Fact]
    public void AutoRestoreDoesNotLoadExcel()
    {
        using var workspace = new TemporaryWorkspace();
        PrepareRecentTemplate(workspace.Path, "通知書", "通知書_{{ROW}}", wordExists: true);

        var viewModel = CreateViewModel(workspace.Path);

        Assert.Equal(string.Empty, viewModel.ExcelPath);
        Assert.Empty(viewModel.WorksheetNames);
        Assert.Empty(viewModel.BatchResults);
    }

    [Fact]
    public void CanGenerateBatchRemainsFalseWithoutExcel()
    {
        using var workspace = new TemporaryWorkspace();
        PrepareRecentTemplate(workspace.Path, "通知書", "通知書_{{ROW}}", wordExists: true);

        var viewModel = CreateViewModel(workspace.Path);

        Assert.False(viewModel.CanGenerateBatch);
    }

    [Fact]
    public void MalformedRecentTemplateDoesNotCrashStartup()
    {
        using var workspace = new TemporaryWorkspace();
        var path = Path.Combine(workspace.Path, "malformed.docxcel.json");
        File.WriteAllText(path, "{ not valid json }");
        SaveSettings(workspace.Path, new ApplicationSettings { LastTemplateConfigPath = path });

        var exception = Record.Exception(() => CreateViewModel(workspace.Path));
        var viewModel = CreateViewModel(workspace.Path);

        Assert.Null(exception);
        Assert.Null(viewModel.CurrentTemplateDefinition);
        Assert.Contains("無法載入", viewModel.StatusText);
    }

    [Fact]
    public void UnsupportedRecentTemplateVersionDoesNotCrashStartup()
    {
        using var workspace = new TemporaryWorkspace();
        var path = Path.Combine(workspace.Path, "future.docxcel.json");
        File.WriteAllText(path, "{\"schemaVersion\":999,\"templateName\":\"Future\"}");
        SaveSettings(workspace.Path, new ApplicationSettings { LastTemplateConfigPath = path });

        var viewModel = CreateViewModel(workspace.Path);

        Assert.Null(viewModel.CurrentTemplateDefinition);
        Assert.Contains("無法載入", viewModel.StatusText);
    }

    [Fact]
    public void MissingWordTemplateKeepsRestoredDefinitionAndDoesNotCrash()
    {
        using var workspace = new TemporaryWorkspace();
        PrepareRecentTemplate(workspace.Path, "通知書", "通知書_{{ROW}}", wordExists: false);

        var viewModel = CreateViewModel(workspace.Path);

        Assert.NotNull(viewModel.CurrentTemplateDefinition);
        Assert.Equal("通知書", viewModel.TemplateName);
        Assert.Contains("找不到 Word 模板", viewModel.StatusText);
        Assert.False(viewModel.CanGenerateBatch);
    }

    [Fact]
    public async Task ManualLoadOfTemplateBUpdatesLastTemplatePath()
    {
        using var workspace = new TemporaryWorkspace();
        var pathB = WriteTemplate(workspace.Path, "B", "b-{{ROW}}", wordExists: true);
        var viewModel = CreateViewModel(workspace.Path);

        await viewModel.LoadTemplateAsync(pathB);

        Assert.Equal(pathB, new SettingsService(SettingsPath(workspace.Path)).Load().LastTemplateConfigPath);
    }

    [Fact]
    public async Task NewCreatedTemplateUpdatesLastTemplatePath()
    {
        using var workspace = new TemporaryWorkspace();
        var excelPath = CreateWorkbook(workspace.Path);
        var wordPath = CreateWordTemplate(workspace.Path);
        var savePath = Path.Combine(workspace.Path, "created.docxcel.json");
        var viewModel = CreateViewModel(workspace.Path);
        await viewModel.LoadExcelAsync(excelPath);
        viewModel.WordTemplatePath = wordPath;
        await viewModel.ReadMappingAsync();

        await viewModel.CreateTemplateAsync(savePath);

        Assert.Equal(savePath, new SettingsService(SettingsPath(workspace.Path)).Load().LastTemplateConfigPath);
    }

    [Fact]
    public async Task NextStartupRestoresLatestTemplateBInsteadOfOlderTemplateA()
    {
        using var workspace = new TemporaryWorkspace();
        var pathA = WriteTemplate(workspace.Path, "A", "a-{{ROW}}", wordExists: true);
        var pathB = WriteTemplate(workspace.Path, "B", "b-{{ROW}}", wordExists: true);
        SaveSettings(workspace.Path, new ApplicationSettings { LastTemplateConfigPath = pathA });
        var viewModel = CreateViewModel(workspace.Path);

        await viewModel.LoadTemplateAsync(pathB);
        var nextStartup = CreateViewModel(workspace.Path);

        Assert.Equal("B", nextStartup.TemplateName);
        Assert.Equal(pathB, new SettingsService(SettingsPath(workspace.Path)).Load().LastTemplateConfigPath);
    }

    [Fact]
    public void AutoRestoreDoesNotOverwriteLastOutputDirectory()
    {
        using var workspace = new TemporaryWorkspace();
        var outputDirectory = Path.Combine(workspace.Path, "RememberedOutput");
        PrepareRecentTemplate(workspace.Path, "通知書", "通知書_{{ROW}}", wordExists: true, outputDirectory: outputDirectory);

        var viewModel = CreateViewModel(workspace.Path);

        Assert.Equal(outputDirectory, viewModel.OutputDirectory);
        Assert.Equal(outputDirectory, new SettingsService(SettingsPath(workspace.Path)).Load().LastOutputDirectory);
    }

    [Fact]
    public void PersistedTemplatePatternTakesPrecedenceOverSettingsPattern()
    {
        using var workspace = new TemporaryWorkspace();
        PrepareRecentTemplate(
            workspace.Path,
            "通知書",
            "template-{{ROW}}",
            wordExists: true,
            outputDirectory: null,
            settingsPattern: "settings-{{ROW}}");

        var viewModel = CreateViewModel(workspace.Path);

        Assert.Equal("template-{{ROW}}", viewModel.OutputPattern);
    }

    [Fact]
    public void AutoRestoreKeepsDefaultGenerateDocumentsMode()
    {
        using var workspace = new TemporaryWorkspace();
        PrepareRecentTemplate(workspace.Path, "通知書", "通知書_{{ROW}}", wordExists: true);

        var viewModel = CreateViewModel(workspace.Path);

        Assert.True(viewModel.IsGenerateDocumentsMode);
        Assert.False(viewModel.IsCreateTemplateMode);
    }

    [Fact]
    public void AutoRestoreUsesIsolatedSettingsPath()
    {
        using var workspace = new TemporaryWorkspace();
        var settingsPath = SettingsPath(workspace.Path);
        var viewModel = CreateViewModel(workspace.Path);

        Assert.Equal(settingsPath, new SettingsService(settingsPath).SettingsPath);
        Assert.Null(viewModel.CurrentTemplateDefinition);
    }

    [Fact]
    public void StartupXamlStillUsesReadonlyProgressBindings()
    {
        var xamlPath = Path.Combine(FindSolutionDirectory(), "src", "XlsxDocxGenerator", "MainWindow.xaml");
        var xaml = File.ReadAllText(xamlPath);

        Assert.Contains("Maximum=\"{Binding ProgressMaximum, Mode=OneWay}\"", xaml);
        Assert.Contains("Value=\"{Binding ProgressValue, Mode=OneWay}\"", xaml);
        Assert.DoesNotContain("Value=\"{Binding ProgressValue}\"", xaml);
    }

    private static MainWindowViewModel CreateViewModel(string directory) =>
        new(
            new SettingsService(SettingsPath(directory)),
            new DiagnosticsLogger(Path.Combine(directory, "Logs")));

    private static string SettingsPath(string directory) => Path.Combine(directory, "settings.json");

    private static void SaveSettings(string directory, ApplicationSettings settings) =>
        new SettingsService(SettingsPath(directory)).Save(settings);

    private static TemplateDefinition PrepareRecentTemplate(
        string directory,
        string name,
        string pattern,
        bool wordExists,
        string? outputDirectory = null,
        string? settingsPattern = null)
    {
        var path = WriteTemplate(directory, name, pattern, wordExists);
        SaveSettings(directory, new ApplicationSettings
        {
            LastTemplateConfigPath = path,
            LastOutputDirectory = outputDirectory,
            LastOutputFilenamePattern = settingsPattern
        });
        return new TemplateService(new ExcelReader(), new MarkerParser()).LoadTemplate(path);
    }

    private static string WriteTemplate(string directory, string name, string pattern, bool wordExists)
    {
        var wordPath = Path.Combine(directory, $"{name}.docx");
        if (wordExists)
        {
            File.WriteAllText(wordPath, "placeholder");
        }

        var path = Path.Combine(directory, $"{name}.docxcel.json");
        var definition = new TemplateDefinition
        {
            TemplateName = name,
            WordTemplatePath = wordPath,
            PreferredWorksheetName = "資料",
            HeaderRowNumber = 1,
            FieldMappings =
            [
                new FieldMapping { PlaceholderName = "ID", ColumnIndex = 1, HeaderName = "編號" },
                new FieldMapping { PlaceholderName = "NAME", ColumnIndex = 2, HeaderName = "姓名" }
            ],
            OutputFileNamePattern = pattern
        };
        new TemplateService(new ExcelReader(), new MarkerParser()).SaveTemplate(path, definition);
        return path;
    }

    private static string CreateWorkbook(string directory)
    {
        var path = Path.Combine(directory, "資料.xlsx");
        using var workbook = new XLWorkbook();
        var worksheet = workbook.AddWorksheet("資料");
        worksheet.Cell(1, 1).Value = "NAME";
        worksheet.Cell(2, 1).Value = "{{NAME}}";
        worksheet.Cell(3, 1).Value = "Alice";
        workbook.SaveAs(path);
        return path;
    }

    private static string CreateWordTemplate(string directory)
    {
        var path = Path.Combine(directory, "通知書.docx");
        using var document = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document);
        var mainPart = document.AddMainDocumentPart();
        mainPart.Document = new Document(
            new Body(
                new Paragraph(
                    new Run(new Text("{{NAME}}")))));
        mainPart.Document.Save();
        return path;
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
                $"docxcel-phase6c-{Guid.NewGuid():N}");
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
