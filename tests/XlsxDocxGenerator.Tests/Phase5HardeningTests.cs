using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using XlsxDocxGenerator.Models;
using XlsxDocxGenerator.Services.Diagnostics;
using XlsxDocxGenerator.Services.Excel;
using XlsxDocxGenerator.Services.Generation;
using XlsxDocxGenerator.Services.Settings;
using XlsxDocxGenerator.Services.Word;

namespace XlsxDocxGenerator.Tests;

public sealed class Phase5HardeningTests
{
    [Fact]
    public void ViewModelStartsInSafeIdleState()
    {
        using var workspace = new TemporaryWorkspace();
        var settings = new SettingsService(Path.Combine(workspace.Path, "settings.json"));
        var logger = new DiagnosticsLogger(Path.Combine(workspace.Path, "Logs"));
        var viewModel = new XlsxDocxGenerator.ViewModels.MainWindowViewModel(settings, logger);

        Assert.False(viewModel.IsGenerating);
        Assert.True(viewModel.IsNotGenerating);
        Assert.False(viewModel.CanCancelBatch);
        Assert.False(viewModel.CanGenerateBatch);
    }

    [Fact]
    public async Task CancellationBeforeFirstRowReturnsCancelledWithoutFailures()
    {
        using var workspace = new TemporaryWorkspace();
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        var progress = new RecordingProgress<BatchProgress>();

        var result = await CreateGenerationService(5).GenerateBatchAsync(
            "ignored.xlsx",
            Template(),
            new SelectionRule.ExcelRows([2, 3, 4, 5, 6]),
            workspace.Path,
            progress,
            cancellation.Token);

        Assert.True(result.IsCancelled);
        Assert.Empty(result.Results);
        Assert.Equal(0, result.FailureCount);
        Assert.Equal(5, result.TotalSelected);
        Assert.Equal(5, result.UnprocessedCount);
        Assert.Empty(progress.Values);
        Assert.Empty(Directory.GetFiles(workspace.Path, "*.docx"));
    }

    [Fact]
    public async Task MidBatchCancellationKeepsCompletedOutputsAndStopsProgress()
    {
        using var workspace = new TemporaryWorkspace();
        using var cancellation = new CancellationTokenSource();
        var progress = new RecordingProgress<BatchProgress>(value =>
        {
            if (value.Completed == 2)
            {
                cancellation.Cancel();
            }
        });

        var result = await CreateGenerationService(5).GenerateBatchAsync(
            "ignored.xlsx",
            Template(),
            new SelectionRule.ExcelRows([2, 3, 4, 5, 6]),
            workspace.Path,
            progress,
            cancellation.Token);

        Assert.True(result.IsCancelled);
        Assert.Equal(2, result.SuccessCount);
        Assert.Equal(0, result.FailureCount);
        Assert.Equal(3, result.UnprocessedCount);
        Assert.Equal([0, 1, 2], progress.Values.Select(value => value.Completed));
        Assert.Equal(2, Directory.GetFiles(workspace.Path, "*.docx").Length);
        Assert.Empty(Directory.GetFiles(workspace.Path, "*.tmp"));
        Assert.DoesNotContain(result.Results, result => result.ExcelRowNumber > 3);
    }

    [Fact]
    public async Task CancellationIsNotConvertedIntoRowFailureAndBatchCanRunAgain()
    {
        using var workspace = new TemporaryWorkspace();
        using var cancellation = new CancellationTokenSource();
        var progress = new RecordingProgress<BatchProgress>(value =>
        {
            if (value.Completed == 1)
            {
                cancellation.Cancel();
            }
        });

        var service = CreateGenerationService(3);
        var cancelled = await service.GenerateBatchAsync(
            "ignored.xlsx", Template(), new SelectionRule.ExcelRows([2, 3, 4]),
            workspace.Path, progress, cancellation.Token);

        Assert.True(cancelled.IsCancelled);
        Assert.Equal(0, cancelled.FailureCount);

        var completed = await CreateGenerationService(1).GenerateBatchAsync(
            "ignored.xlsx", Template(), new SelectionRule.ExcelRows([2]), workspace.Path);

        Assert.False(completed.IsCancelled);
        Assert.Equal(1, completed.SuccessCount);
    }

    [Fact]
    public async Task CancellationDuringDocxGenerationCleansTemporaryOutput()
    {
        using var workspace = new TemporaryWorkspace();
        var templatePath = CreateDocument(workspace.Path);
        var outputPath = Path.Combine(workspace.Path, "output.docx");
        using var cancellation = new CancellationTokenSource();
        var definition = new TemplateDefinition
        {
            TemplateName = "Cancellation document test",
            WordTemplatePath = templatePath,
            PreferredWorksheetName = "Sheet1",
            HeaderRowNumber = 1,
            FieldMappings = [new FieldMapping { PlaceholderName = "NAME", ColumnIndex = 1 }]
        };
        var record = new ExcelRecord
        {
            RowNumber = 2,
            Values = new Dictionary<string, string?> { ["NAME"] = "Alice" },
            ColumnValues = new Dictionary<int, string?> { [1] = "Alice" }
        };

        var generator = new DocxGenerator(
            new FixedTemplateReader(),
            new CancellingPlaceholderResolver(cancellation));

        await Assert.ThrowsAsync<OperationCanceledException>(() => generator.GenerateAsync(
            definition,
            record,
            outputPath,
            cancellation.Token));

        Assert.False(File.Exists(outputPath));
        Assert.Empty(Directory.GetFiles(workspace.Path, "*.tmp"));
        Assert.Equal("{{NAME}}", ReadDocumentText(templatePath));
    }

    [Fact]
    public void MissingSettingsUseDefaults()
    {
        var settings = new SettingsService(Path.Combine(Path.GetTempPath(), $"docxcel-{Guid.NewGuid():N}", "settings.json"));

        var loaded = settings.Load();

        Assert.Equal(1, loaded.SettingsVersion);
        Assert.Null(loaded.LastTemplateConfigPath);
        Assert.Null(loaded.LastOutputDirectory);
    }

    [Fact]
    public void SettingsRoundTripAndAtomicReplacementKeepOnlyFinalJson()
    {
        using var workspace = new TemporaryWorkspace();
        var path = Path.Combine(workspace.Path, "settings.json");
        var service = new SettingsService(path);
        service.Save(new ApplicationSettings
        {
            LastTemplateConfigPath = "C:\\文件\\通知.docxcel.json",
            LastOutputDirectory = "D:\\輸出資料",
            LastSelectionMode = "Latest",
            LastOutputFilenamePattern = "notice-{{ROW}}",
            WindowWidth = 1200,
            WindowHeight = 900
        });
        var firstLoad = service.Load();
        Assert.Equal("C:\\文件\\通知.docxcel.json", firstLoad.LastTemplateConfigPath);
        Assert.Equal("D:\\輸出資料", firstLoad.LastOutputDirectory);
        Assert.Equal(1200, firstLoad.WindowWidth);
        Assert.Equal(900, firstLoad.WindowHeight);

        service.Save(new ApplicationSettings { LastOutputDirectory = "D:\\Output" });

        var loaded = service.Load();

        Assert.Equal("D:\\Output", loaded.LastOutputDirectory);
        Assert.Null(loaded.LastTemplateConfigPath);
        Assert.Null(loaded.WindowWidth);
        Assert.Null(loaded.WindowHeight);
        Assert.Single(Directory.GetFiles(workspace.Path, "settings.json"));
        Assert.Empty(Directory.GetFiles(workspace.Path, "*.tmp"));
    }

    [Fact]
    public void MalformedAndUnsupportedSettingsFallBackToDefaults()
    {
        using var workspace = new TemporaryWorkspace();
        var path = Path.Combine(workspace.Path, "settings.json");
        var service = new SettingsService(path);

        File.WriteAllText(path, "{ not json }");
        Assert.Null(service.Load().LastOutputDirectory);

        File.WriteAllText(path, "{\"settingsVersion\":99,\"lastOutputDirectory\":\"secret\"}");
        var unsupported = service.Load();
        Assert.Equal(1, unsupported.SettingsVersion);
        Assert.Null(unsupported.LastOutputDirectory);
    }

    [Fact]
    public void SettingsWriteFailureDoesNotThrow()
    {
        using var workspace = new TemporaryWorkspace();
        var blocker = Path.Combine(workspace.Path, "blocker");
        File.WriteAllText(blocker, "not a directory");
        var service = new SettingsService(Path.Combine(blocker, "settings.json"));

        var exception = Record.Exception(() => service.Save(new ApplicationSettings
        {
            LastOutputDirectory = "C:\\Output"
        }));

        Assert.Null(exception);
    }

    [Fact]
    public void DiagnosticsLogContainsMetadataButNotDocumentData()
    {
        using var workspace = new TemporaryWorkspace();
        var logger = new DiagnosticsLogger(workspace.Path);

        logger.LogException(
            "TestOperation",
            "TestCode",
            new InvalidOperationException("failure without document contents"));

        var log = File.ReadAllText(Directory.GetFiles(workspace.Path, "*.log").Single());
        Assert.Contains("TestOperation", log);
        Assert.Contains("TestCode", log);
        Assert.Contains(nameof(InvalidOperationException), log);
        Assert.DoesNotContain("Excel cell secret value", log);
    }

    [Fact]
    public void DiagnosticsWriteFailureDoesNotThrow()
    {
        using var workspace = new TemporaryWorkspace();
        var blocker = Path.Combine(workspace.Path, "blocker");
        File.WriteAllText(blocker, "not a directory");
        var logger = new DiagnosticsLogger(Path.Combine(blocker, "Logs"));

        var exception = Record.Exception(() => logger.LogException(
            "TestOperation",
            "TestCode",
            new InvalidOperationException("failure")));

        Assert.Null(exception);
    }

    private static GenerationService CreateGenerationService(int rowCount) =>
        new(
            new StubExcelReader(CreateWorksheet(rowCount)),
            new WritingDocxGenerator());

    private static TemplateDefinition Template() => new()
    {
        TemplateName = "Cancellation test",
        WordTemplatePath = "template.docx",
        PreferredWorksheetName = "資料",
        HeaderRowNumber = 1,
        FieldMappings = [new FieldMapping { PlaceholderName = "NAME", ColumnIndex = 1 }]
    };

    private static ExcelWorksheetData CreateWorksheet(int rowCount) => new()
    {
        WorksheetName = "資料",
        HeaderRowNumber = 1,
        FirstUsedRowNumber = 1,
        LastUsedRowNumber = rowCount + 1,
        Headers = ["NAME"],
        Rows = Enumerable.Range(2, rowCount)
            .Select(row => new ExcelRowData
            {
                RowNumber = row,
                Cells = new Dictionary<int, string?> { [1] = $"Name {row}" }
            })
            .ToArray()
    };

    private static string CreateDocument(string directory)
    {
        var path = Path.Combine(directory, "template.docx");
        using var document = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document);
        var mainPart = document.AddMainDocumentPart();
        mainPart.Document = new Document(
            new Body(new Paragraph(new Run(new Text("{{NAME}}")))));
        mainPart.Document.Save();
        return path;
    }

    private static string ReadDocumentText(string path)
    {
        using var document = WordprocessingDocument.Open(path, false);
        return string.Concat(document.MainDocumentPart!.Document!.Body!.Descendants<Text>().Select(text => text.Text));
    }

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
            cancellationToken.ThrowIfCancellationRequested();
            File.WriteAllText(outputPath, record.GetValueByColumnIndex(1) ?? string.Empty);
            return Task.CompletedTask;
        }
    }

    private sealed class FixedTemplateReader : IDocxTemplateReader
    {
        public Task<DocxTemplateInfo> ReadAsync(
            string templatePath,
            CancellationToken cancellationToken = default) =>
            Task.FromResult(new DocxTemplateInfo("Cancellation document test", templatePath));

        public Task<IReadOnlyList<string>> ScanPlaceholdersAsync(
            string templatePath,
            CancellationToken cancellationToken = default) =>
            Task.FromResult<IReadOnlyList<string>>(["NAME"]);
    }

    private sealed class CancellingPlaceholderResolver(CancellationTokenSource cancellation)
        : IPlaceholderResolver
    {
        public IReadOnlyList<PlaceholderOccurrence> FindOccurrences(OpenXmlElement root) => [];

        public void ReplaceOccurrences(
            OpenXmlElement root,
            IReadOnlyDictionary<string, string?> values) =>
            cancellation.Cancel();
    }

    private sealed class RecordingProgress<T>(Action<T>? action = null) : IProgress<T>
    {
        public List<T> Values { get; } = [];

        public void Report(T value)
        {
            Values.Add(value);
            action?.Invoke(value);
        }
    }

    private sealed class TemporaryWorkspace : IDisposable
    {
        public TemporaryWorkspace()
        {
            Path = System.IO.Path.Combine(
                System.IO.Path.GetTempPath(),
                $"xlsx-docx-phase5-{Guid.NewGuid():N}");
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
