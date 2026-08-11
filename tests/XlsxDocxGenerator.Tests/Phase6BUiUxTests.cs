using XlsxDocxGenerator.Models;
using XlsxDocxGenerator.Services.Diagnostics;
using XlsxDocxGenerator.Services.Excel;
using XlsxDocxGenerator.Services.Settings;
using XlsxDocxGenerator.Services.Templates;
using XlsxDocxGenerator.ViewModels;

namespace XlsxDocxGenerator.Tests;

public sealed class Phase6BUiUxTests
{
    [Fact]
    public void DefaultModeIsGenerateDocuments()
    {
        using var workspace = new TemporaryWorkspace();
        var viewModel = CreateViewModel(workspace.Path);

        Assert.Equal(MainWindowMode.GenerateDocuments, viewModel.ActiveMode);
        Assert.True(viewModel.IsGenerateDocumentsMode);
        Assert.False(viewModel.IsCreateTemplateMode);
    }

    [Fact]
    public void SwitchingToCreateTemplateChangesOnlyPresentationMode()
    {
        using var workspace = new TemporaryWorkspace();
        var viewModel = CreateViewModel(workspace.Path);
        viewModel.ActiveMode = MainWindowMode.CreateTemplate;

        Assert.True(viewModel.IsCreateTemplateMode);
        Assert.False(viewModel.IsGenerateDocumentsMode);
    }

    [Fact]
    public void SwitchingBackToGenerateDocumentsRestoresGenerateMode()
    {
        using var workspace = new TemporaryWorkspace();
        var viewModel = CreateViewModel(workspace.Path);
        viewModel.ActiveMode = MainWindowMode.CreateTemplate;

        viewModel.ActiveMode = MainWindowMode.GenerateDocuments;

        Assert.True(viewModel.IsGenerateDocumentsMode);
        Assert.False(viewModel.IsCreateTemplateMode);
    }

    [Fact]
    public void GenerateModeDoesNotRequireTemplateCreationOnlyState()
    {
        using var workspace = new TemporaryWorkspace();
        var viewModel = CreateViewModel(workspace.Path);

        Assert.True(viewModel.IsGenerateDocumentsMode);
        Assert.False(viewModel.IsCreateTemplateMode);
        Assert.False(viewModel.CanGenerateBatch);
    }

    [Fact]
    public async Task SwitchingModePreservesLoadedTemplateSemantics()
    {
        using var workspace = new TemporaryWorkspace();
        var configPath = Path.Combine(workspace.Path, "template.docxcel.json");
        var definition = CreateDefinition("saved-{{ROW}}", "通知模板");
        new TemplateService(new ExcelReader(), new MarkerParser())
            .SaveTemplate(configPath, definition);
        var viewModel = CreateViewModel(workspace.Path);

        await viewModel.LoadTemplateAsync(configPath);
        viewModel.ActiveMode = MainWindowMode.CreateTemplate;
        viewModel.ActiveMode = MainWindowMode.GenerateDocuments;

        Assert.Equal("通知模板", viewModel.CurrentTemplateDefinition!.TemplateName);
        Assert.Equal("saved-{{ROW}}", viewModel.CurrentTemplateDefinition.OutputFileNamePattern);
        Assert.Single(viewModel.CurrentTemplateDefinition.FieldMappings);
    }

    [Fact]
    public void SwitchingModePreservesExcelSelection()
    {
        using var workspace = new TemporaryWorkspace();
        var viewModel = CreateViewModel(workspace.Path);
        viewModel.ExcelPath = Path.Combine(workspace.Path, "資料.xlsx");
        viewModel.SelectedWorksheet = "Sheet1";

        viewModel.ActiveMode = MainWindowMode.CreateTemplate;
        viewModel.ActiveMode = MainWindowMode.GenerateDocuments;

        Assert.Equal(Path.Combine(workspace.Path, "資料.xlsx"), viewModel.ExcelPath);
        Assert.Equal("Sheet1", viewModel.SelectedWorksheet);
    }

    [Fact]
    public void SwitchingModePreservesOutputPreferences()
    {
        using var workspace = new TemporaryWorkspace();
        var viewModel = CreateViewModel(workspace.Path);
        viewModel.OutputDirectory = Path.Combine(workspace.Path, "Output");
        viewModel.OutputPattern = "{{ID}}_{{NAME}}";

        viewModel.ActiveMode = MainWindowMode.CreateTemplate;
        viewModel.ActiveMode = MainWindowMode.GenerateDocuments;

        Assert.Equal(Path.Combine(workspace.Path, "Output"), viewModel.OutputDirectory);
        Assert.Equal("{{ID}}_{{NAME}}", viewModel.OutputPattern);
    }

    [Fact]
    public async Task TemplateNameEditPreservesExistingMappingsAndPattern()
    {
        using var workspace = new TemporaryWorkspace();
        var configPath = Path.Combine(workspace.Path, "template.docxcel.json");
        var definition = CreateDefinition("saved-{{ROW}}", "原始名稱");
        new TemplateService(new ExcelReader(), new MarkerParser())
            .SaveTemplate(configPath, definition);
        var viewModel = CreateViewModel(workspace.Path);

        await viewModel.LoadTemplateAsync(configPath);
        viewModel.TemplateName = "新的名稱";

        Assert.Equal("新的名稱", viewModel.CurrentTemplateDefinition!.TemplateName);
        Assert.Equal("saved-{{ROW}}", viewModel.CurrentTemplateDefinition.OutputFileNamePattern);
        Assert.Equal("NAME", viewModel.CurrentTemplateDefinition.FieldMappings[0].PlaceholderName);
    }

    [Fact]
    public void RuntimeMarkerInstructionUsesPlainUserLanguage()
    {
        using var workspace = new TemporaryWorkspace();
        var viewModel = CreateViewModel(workspace.Path);

        Assert.DoesNotContain("strict", viewModel.RuntimeMarkerInstructionText, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("runtime marker algorithm", viewModel.RuntimeMarkerInstructionText, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void MainWindowXamlUsesCentralStylesAndReadonlyProgressBindings()
    {
        var xamlPath = Path.Combine(FindSolutionDirectory(), "src", "XlsxDocxGenerator", "MainWindow.xaml");
        var xaml = File.ReadAllText(xamlPath);

        Assert.Contains("PrimaryButtonStyle", xaml);
        Assert.Contains("SegmentButtonStyle", xaml);
        Assert.Contains("Maximum=\"{Binding ProgressMaximum, Mode=OneWay}\"", xaml);
        Assert.Contains("Value=\"{Binding ProgressValue, Mode=OneWay}\"", xaml);
        Assert.DoesNotContain("Value=\"{Binding ProgressValue}\"", xaml);
    }

    [Fact]
    public void MainWindowXamlContainsBothUserFacingWorkModes()
    {
        var xamlPath = Path.Combine(FindSolutionDirectory(), "src", "XlsxDocxGenerator", "MainWindow.xaml");
        var xaml = File.ReadAllText(xamlPath);

        Assert.Contains("Content=\"產生文件\"", xaml);
        Assert.Contains("Content=\"建立模板\"", xaml);
        Assert.Contains("Header=\"進階選項\"", xaml);
        Assert.DoesNotContain("strict automatic detection", xaml, StringComparison.OrdinalIgnoreCase);
    }

    private static MainWindowViewModel CreateViewModel(string directory) =>
        new(
            new SettingsService(Path.Combine(directory, "settings.json")),
            new DiagnosticsLogger(Path.Combine(directory, "Logs")));

    private static TemplateDefinition CreateDefinition(string pattern, string name) => new()
    {
        TemplateName = name,
        WordTemplatePath = "template.docx",
        PreferredWorksheetName = "Sheet1",
        HeaderRowNumber = 1,
        FieldMappings = [new FieldMapping { PlaceholderName = "NAME", ColumnIndex = 1 }],
        OutputFileNamePattern = pattern
    };

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
                $"docxcel-phase6b-{Guid.NewGuid():N}");
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
