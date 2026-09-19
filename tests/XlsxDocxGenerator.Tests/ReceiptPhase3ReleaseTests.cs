using System.Security.Cryptography;
using System.Diagnostics;
using System.Reflection;
using System.Text.Json;
using ClosedXML.Excel;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using XlsxDocxGenerator.Models;
using XlsxDocxGenerator.Services.Diagnostics;
using XlsxDocxGenerator.Services.Receipts;
using XlsxDocxGenerator.Services.Settings;
using XlsxDocxGenerator.ViewModels;

namespace XlsxDocxGenerator.Tests;

public sealed class ReceiptPhase3ReleaseTests
{
    private const string PayerSentinel = "PAYER_SENTINEL";
    private const string ReasonSentinel = "運動會禮金";
    private const string HColumnSentinel = "6/6匯款";

    [Fact]
    public void ReleaseIdentityAndPortableProfileAreExplicit()
    {
        var project = File.ReadAllText(Path.Combine(
            RepositoryRoot(),
            "src",
            "XlsxDocxGenerator",
            "XlsxDocxGenerator.csproj"));
        var profile = File.ReadAllText(Path.Combine(
            RepositoryRoot(),
            "src",
            "XlsxDocxGenerator",
            "Properties",
            "PublishProfiles",
            "WinX64Portable.pubxml"));

        Assert.Contains("<AssemblyName>ReceiptXcel 收據產生工具</AssemblyName>", project);
        Assert.Contains("<Product>ReceiptXcel</Product>", project);
        Assert.Contains("<Title>ReceiptXcel｜自行收納款項收據產生工具</Title>", project);
        Assert.Contains("<Version>0.1.1</Version>", project);
        Assert.Contains("<AssemblyVersion>0.1.1.0</AssemblyVersion>", project);
        Assert.Contains("<FileVersion>0.1.1.0</FileVersion>", project);
        Assert.Contains("<InformationalVersion>0.1.1</InformationalVersion>", project);
        Assert.Contains("<CopyToOutputDirectory>PreserveNewest</CopyToOutputDirectory>", project);
        Assert.Contains("<CopyToPublishDirectory>PreserveNewest</CopyToPublishDirectory>", project);
        Assert.Contains("<RuntimeIdentifier>win-x64</RuntimeIdentifier>", profile);
        Assert.Contains("<SelfContained>true</SelfContained>", profile);
        Assert.Contains("<PublishTrimmed>false</PublishTrimmed>", profile);
        Assert.Contains("<PublishSingleFile>false</PublishSingleFile>", profile);
        Assert.Contains("win-x64-phase3-rc", profile);

        var assembly = typeof(XlsxDocxGenerator.App).Assembly;
        Assert.Equal("ReceiptXcel 收據產生工具", assembly.GetName().Name);
        Assert.Equal(new Version(0, 1, 1, 0), assembly.GetName().Version);
        Assert.Equal("ReceiptXcel", assembly.GetCustomAttribute<AssemblyProductAttribute>()!.Product);
        Assert.Equal("ReceiptXcel 收據產生工具.dll", Path.GetFileName(assembly.Location));
        var metadata = FileVersionInfo.GetVersionInfo(assembly.Location);
        Assert.Equal("0.1.1.0", metadata.FileVersion);
        Assert.StartsWith("0.1.1", metadata.ProductVersion);
    }

    [Fact]
    public void DefaultReceiptServiceUsesRuntimeInternalTemplateNotReferenceFiles()
    {
        var service = new ReceiptGenerationService();

        Assert.Equal(
            Path.Combine(AppContext.BaseDirectory, "Assets", "Templates", "receipt-template.docx"),
            service.InternalTemplatePath);
        Assert.True(File.Exists(service.InternalTemplatePath));
        Assert.DoesNotContain("reference", service.InternalTemplatePath, StringComparison.OrdinalIgnoreCase);
        Assert.True(new ReceiptBatchGenerationService().IsInternalTemplateAvailable(out _));
    }

    [Fact]
    public void DiagnosticsDoNotPersistBusinessValuesFromExceptions()
    {
        using var workspace = new TemporaryWorkspace();
        var logger = new DiagnosticsLogger(Path.Combine(workspace.Path, "Logs"));

        try
        {
            throw new InvalidOperationException(
                $"{PayerSentinel}|{ReasonSentinel}|1000|{HColumnSentinel}|{{{{REASON}}}}");
        }
        catch (InvalidOperationException exception)
        {
            logger.LogException("ReleasePrivacyAudit", "SentinelFailure", exception);
        }

        var log = File.ReadAllText(Directory.GetFiles(logger.LogDirectory, "*.log").Single());
        Assert.Contains("Operation: ReleasePrivacyAudit", log);
        Assert.Contains("InternalErrorCode: SentinelFailure", log);
        Assert.Contains("ExceptionType: System.InvalidOperationException", log);
        Assert.DoesNotContain(PayerSentinel, log, StringComparison.Ordinal);
        Assert.DoesNotContain(ReasonSentinel, log, StringComparison.Ordinal);
        Assert.DoesNotContain("1000", log, StringComparison.Ordinal);
        Assert.DoesNotContain(HColumnSentinel, log, StringComparison.Ordinal);
        Assert.DoesNotContain("{{REASON}}", log, StringComparison.Ordinal);
    }

    [Fact]
    public void SpecializedSettingsPersistOnlyReceiptXcelUiPreferences()
    {
        using var workspace = new TemporaryWorkspace();
        var settingsPath = Path.Combine(workspace.Path, "settings.json");
        var viewModel = new ReceiptMainWindowViewModel(
            new SettingsService(settingsPath),
            new DiagnosticsLogger(Path.Combine(workspace.Path, "Logs")),
            CreateBatchService());

        Assert.False(viewModel.CanGenerate);
        viewModel.OutputDirectory = Path.Combine(workspace.Path, "Output");
        viewModel.SaveWindowSize(980, 850);

        using var document = JsonDocument.Parse(File.ReadAllText(settingsPath));
        var names = document.RootElement.EnumerateObject()
            .Select(property => property.Name)
            .OrderBy(name => name)
            .ToArray();

        Assert.Equal(
            [
                "lastOutputDirectory",
                "lastSelectionMode",
                "settingsVersion",
                "windowHeight",
                "windowWidth"
            ],
            names);
        var json = document.RootElement.GetRawText();
        Assert.DoesNotContain("ExcelPath", json, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("Payer", json, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("Reason", json, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("Amount", json, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("RowExpression", json, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("ReceiptNumber", json, StringComparison.OrdinalIgnoreCase);
    }

    [Theory]
    [InlineData(300, "", "", "", "$", "3", "0", "0", "參佰元整")]
    [InlineData(1_000, "", "", "$", "1", "0", "0", "0", "壹仟元整")]
    [InlineData(12_000, "", "$", "1", "2", "0", "0", "0", "壹萬貳仟元整")]
    [InlineData(60_000, "", "$", "6", "0", "0", "0", "0", "陸萬元整")]
    [InlineData(888_888, "$", "8", "8", "8", "8", "8", "8", "捌拾捌萬捌仟捌佰捌拾捌元整")]
    [InlineData(1_000_000, "1", "0", "0", "0", "0", "0", "0", "壹佰萬元整")]
    [InlineData(9_999_999, "9", "9", "9", "9", "9", "9", "9", "玖佰玖拾玖萬玖仟玖佰玖拾玖元整")]
    public void ReleaseAmountFormattingKeepsAllSevenCellsAndChineseUppercase(
        int amount,
        string million,
        string hundredThousand,
        string tenThousand,
        string thousand,
        string hundred,
        string ten,
        string one,
        string uppercase)
    {
        var formatted = new ReceiptAmountFormatter().Format(amount);

        Assert.Equal(
            [million, hundredThousand, tenThousand, thousand, hundred, ten, one],
            [
                formatted.MillionCell,
                formatted.HundredThousandCell,
                formatted.TenThousandCell,
                formatted.ThousandCell,
                formatted.HundredCell,
                formatted.TenCell,
                formatted.OneCell
            ]);
        Assert.Equal(uppercase, formatted.ChineseUppercase);
    }

    [Fact]
    public async Task Phase3QaArtifactsCoverReleaseBlockingScenarios()
    {
        var qaDirectory = Path.Combine(RepositoryRoot(), "artifacts", "qa", "phase3");
        Directory.CreateDirectory(qaDirectory);
        var workbook = Path.Combine(qaDirectory, "phase3-qa.xlsx");
        var referenceExcelHash = Hash(ReferenceExcelPath());
        var referenceWordHash = Hash(ReferenceWordPath());
        CreateQaWorkbook(workbook);

        var service = CreateBatchService();
        var normalDirectory = PrepareDirectory(qaDirectory, "normal-gh");
        var batchDirectory = PrepareDirectory(qaDirectory, "multi-row-batch");
        var existingDirectory = PrepareDirectory(qaDirectory, "existing-output");
        var overLimitDirectory = PrepareDirectory(qaDirectory, "amount-over-limit");
        var duplicateDirectory = Path.Combine(qaDirectory, "duplicate-blocked");
        if (Directory.Exists(duplicateDirectory))
        {
            Directory.Delete(duplicateDirectory, recursive: true);
        }

        var cancellationDirectory = PrepareDirectory(qaDirectory, "cancellation");
        var normal = await service.GenerateBatchAsync(
            workbook,
            ReceiptWorksheetSchema.WorksheetName,
            new ReceiptSelectionRequest(ReceiptSelectionMode.ExcelRows, "3"),
            normalDirectory);
        var normalOutput = Assert.Single(normal.Results).OutputPath;
        Assert.NotNull(normalOutput);
        Assert.Equal(1, normal.SuccessCount);
        AssertDocx(normalOutput!, ReasonSentinel, HColumnSentinel);

        var batch = await service.GenerateBatchAsync(
            workbook,
            ReceiptWorksheetSchema.WorksheetName,
            new ReceiptSelectionRequest(ReceiptSelectionMode.ExcelRows, "3,4"),
            batchDirectory);
        Assert.Equal(2, batch.SuccessCount);
        Assert.Equal(2, Directory.GetFiles(batchDirectory, "*.docx").Length);

        var existingPath = Path.Combine(existingDirectory, "收據_1120051.docx");
        File.WriteAllText(existingPath, "preserve-existing-output");
        var existing = await service.GenerateBatchAsync(
            workbook,
            ReceiptWorksheetSchema.WorksheetName,
            new ReceiptSelectionRequest(ReceiptSelectionMode.ExcelRows, "3,4"),
            existingDirectory);
        Assert.Equal(1, existing.SuccessCount);
        Assert.Equal(1, existing.FailureCount);
        Assert.Equal("preserve-existing-output", File.ReadAllText(existingPath));
        Assert.DoesNotContain(Directory.GetFiles(existingDirectory), path => Path.GetFileName(path).Contains("(2)", StringComparison.Ordinal));

        var overLimit = await service.GenerateBatchAsync(
            workbook,
            ReceiptWorksheetSchema.WorksheetName,
            new ReceiptSelectionRequest(ReceiptSelectionMode.ExcelRows, "3,5"),
            overLimitDirectory);
        Assert.Equal(1, overLimit.SuccessCount);
        Assert.Equal(1, overLimit.FailureCount);
        Assert.Contains(overLimit.Results, result => result.ErrorCode == nameof(ReceiptValidationErrorCode.AmountExceedsLimit));

        var duplicate = await Assert.ThrowsAsync<ReceiptBatchGenerationException>(() =>
            service.GenerateBatchAsync(
                workbook,
                ReceiptWorksheetSchema.WorksheetName,
                new ReceiptSelectionRequest(ReceiptSelectionMode.ExcelRows, "3,7"),
                duplicateDirectory));
        Assert.Equal(ReceiptBatchFatalErrorCode.DuplicateReceiptNumber, duplicate.Code);
        Assert.False(Directory.Exists(duplicateDirectory));

        using var cancellation = new CancellationTokenSource();
        var cancelled = await service.GenerateBatchAsync(
            workbook,
            ReceiptWorksheetSchema.WorksheetName,
            new ReceiptSelectionRequest(ReceiptSelectionMode.ExcelRows, "3,4,6"),
            cancellationDirectory,
            new CancelAfterFirstProgress(cancellation),
            cancellation.Token);
        Assert.True(cancelled.IsCancelled);
        Assert.Equal(1, cancelled.SuccessCount);
        Assert.Equal(2, cancelled.UnprocessedCount);
        Assert.Single(Directory.GetFiles(cancellationDirectory, "*.docx"));
        Assert.Empty(Directory.GetFiles(cancellationDirectory, ".*.tmp"));

        Assert.Equal(referenceExcelHash, Hash(ReferenceExcelPath()));
        Assert.Equal(referenceWordHash, Hash(ReferenceWordPath()));
    }

    private static ReceiptBatchGenerationService CreateBatchService() =>
        new(receiptGenerationService: new ReceiptGenerationService(internalTemplatePath: TemplatePath()));

    private static void CreateQaWorkbook(string path)
    {
        using var workbook = new XLWorkbook();
        var worksheet = workbook.Worksheets.Add(ReceiptWorksheetSchema.WorksheetName);
        foreach (var (header, index) in ReceiptWorksheetSchema.ExpectedHeaders.Select((value, index) => (value, index)))
        {
            worksheet.Cell(2, index + 1).Value = header;
        }

        WriteRow(worksheet, 3, "0051", PayerSentinel, 1_000, ReasonSentinel, HColumnSentinel);
        WriteRow(worksheet, 4, "0052", "BATCH_PAYER", 12_000, "第二筆事由", "INTERNAL_NOTE_B");
        WriteRow(worksheet, 5, "0053", "OVER_LIMIT_PAYER", 10_000_000, "超額金額", "INTERNAL_NOTE_C");
        WriteRow(worksheet, 6, "0054", "CANCEL_PAYER", 300, "取消測試", "INTERNAL_NOTE_D");
        WriteRow(worksheet, 7, "0051", "DUPLICATE_PAYER", 60_000, "重複收據編號", "INTERNAL_NOTE_E");
        workbook.SaveAs(path);
    }

    private static void WriteRow(
        IXLWorksheet worksheet,
        int row,
        string serial,
        string payer,
        int amount,
        string reason,
        string sourceColumnH)
    {
        worksheet.Cell(row, 1).Value = 112;
        worksheet.Cell(row, 2).Value = 12;
        worksheet.Cell(row, 3).Value = 30;
        worksheet.Cell(row, 4).Value = serial;
        worksheet.Cell(row, 4).Style.NumberFormat.Format = "@";
        worksheet.Cell(row, 5).Value = payer;
        worksheet.Cell(row, 6).Value = amount.ToString();
        worksheet.Cell(row, 7).Value = reason;
        worksheet.Cell(row, 8).Value = sourceColumnH;
    }

    private static string PrepareDirectory(string root, string name)
    {
        var path = Path.Combine(root, name);
        if (Directory.Exists(path))
        {
            Directory.Delete(path, recursive: true);
        }

        Directory.CreateDirectory(path);
        return path;
    }

    private static void AssertDocx(string path, string expectedReason, string excludedHValue)
    {
        using var document = WordprocessingDocument.Open(path, false);
        var body = document.MainDocumentPart!.Document.Body!;
        var text = AllText(body);
        var tables = body.Descendants<Table>().ToArray();

        Assert.DoesNotContain("{{", text, StringComparison.Ordinal);
        Assert.Contains("21813799", text, StringComparison.Ordinal);
        Assert.Equal(3, CountOccurrences(text, expectedReason));
        Assert.DoesNotContain(excludedHValue, text, StringComparison.Ordinal);
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

    private sealed class CancelAfterFirstProgress(CancellationTokenSource source)
        : IProgress<ReceiptBatchProgress>
    {
        public void Report(ReceiptBatchProgress value)
        {
            if (value.Completed > 0)
            {
                source.Cancel();
            }
        }
    }

    private sealed class TemporaryWorkspace : IDisposable
    {
        public TemporaryWorkspace()
        {
            Path = System.IO.Path.Combine(
                System.IO.Path.GetTempPath(),
                $"receiptxcel-phase3-release-{Guid.NewGuid():N}");
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
