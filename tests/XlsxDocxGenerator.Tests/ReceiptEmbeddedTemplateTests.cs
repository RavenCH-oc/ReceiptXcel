using System.Security.Cryptography;
using System.Xml.Linq;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using XlsxDocxGenerator.Models;
using XlsxDocxGenerator.Services.Receipts;

namespace XlsxDocxGenerator.Tests;

public sealed class ReceiptEmbeddedTemplateTests
{
    private const string TemplateHash = "C0C31D76A42CEE0D9D95F0AAF09736DB37F2840323919A9B5C9FFFDCD2BB690E";
    private static readonly XNamespace Wp = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
    private static readonly XNamespace A = "http://schemas.openxmlformats.org/drawingml/2006/main";

    [Fact]
    public void ManifestContainsNonemptyProductionTemplateBytes()
    {
        var assembly = typeof(EmbeddedReceiptTemplateProvider).Assembly;
        Assert.Contains(EmbeddedReceiptTemplateProvider.ResourceName, assembly.GetManifestResourceNames());
        using var stream = assembly.GetManifestResourceStream(EmbeddedReceiptTemplateProvider.ResourceName)!;
        Assert.NotNull(stream);
        Assert.True(stream.Length > 0);
        Assert.Equal(TemplateHash, Convert.ToHexString(SHA256.HashData(stream)));
    }

    [Fact]
    public void ExtractionOpensWithOpenXmlAndPreservesContract()
    {
        using var workspace = new Workspace();
        using var lease = new EmbeddedReceiptTemplateProvider(workspace.Temp).Acquire();
        using var document = WordprocessingDocument.Open(lease.TemplatePath, false);
        Assert.NotNull(document.MainDocumentPart!.Document.Body);
        var contract = new ReceiptTemplateContractValidator().Validate(lease.TemplatePath);
        Assert.Equal(15, contract.OccurrenceCounts.Count);
        Assert.All(contract.OccurrenceCounts.Values, count => Assert.Equal(3, count));
        Assert.Equal(TemplateHash, Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(lease.TemplatePath))));
    }

    [Fact]
    public void CandidateBEmbeddedLinesRemainUnchanged()
    {
        using var workspace = new Workspace();
        using var lease = new EmbeddedReceiptTemplateProvider(workspace.Temp).Acquire();
        AssertLines(lease.TemplatePath);
    }

    [Fact]
    public void IndependentLeasesNeverShareOrDeleteEachOthersTemplates()
    {
        using var workspace = new Workspace();
        var provider = new EmbeddedReceiptTemplateProvider(workspace.Temp);
        using var first = provider.Acquire();
        using var second = provider.Acquire();
        Assert.NotEqual(first.TemplatePath, second.TemplatePath);
        first.Dispose();
        first.Dispose();
        Assert.False(Directory.Exists(Path.GetDirectoryName(first.TemplatePath)));
        Assert.True(File.Exists(second.TemplatePath));
        second.Dispose();
        Assert.Empty(Directory.EnumerateFileSystemEntries(workspace.Temp));
    }

    [Fact]
    public void ExplicitTemplateOverrideRemainsCallerOwned()
    {
        using var workspace = new Workspace();
        using var lease = new EmbeddedReceiptTemplateProvider(workspace.Temp).Acquire();
        new ReceiptGenerationService(internalTemplatePath: lease.TemplatePath).ValidateInternalTemplate();
        Assert.True(File.Exists(lease.TemplatePath));
    }

    [Fact]
    public async Task GenerationFromEmbeddingKeepsBusinessAndDrawingContracts()
    {
        using var workspace = new Workspace();
        var provider = new TrackingProvider(workspace.Temp);
        var path = Path.Combine(workspace.Path, "receipt.docx");
        await new ReceiptGenerationService(templateProvider: provider).GenerateReceiptAsync(Record(), path);
        new ReceiptGeneratedDocumentValidator().Validate(path, Record(), new ReceiptAmountFormatter().Format(1000));
        AssertReceipt(path);
        AssertLines(path);
        AssertClean(provider, workspace);
    }

    [Fact]
    public async Task BatchAcquiresExactlyOnceAndCleansAfterSuccess()
    {
        using var workspace = new Workspace();
        var provider = new TrackingProvider(workspace.Temp);
        var result = await Batch(provider).GenerateBatchAsync(
            Workbook(workspace), "工作表1", Selection(), Path.Combine(workspace.Path, "out"),
            new CallbackProgress(_ => Assert.True(File.Exists(Assert.Single(provider.Paths)))));
        Assert.Equal(2, result.SuccessCount);
        Assert.All(result.Results, row => AssertReceipt(row.OutputPath!));
        AssertClean(provider, workspace);
    }

    [Fact]
    public async Task BatchRowFailureStillCleansAndContinues()
    {
        using var workspace = new Workspace();
        var provider = new TrackingProvider(workspace.Temp);
        var output = Directory.CreateDirectory(Path.Combine(workspace.Path, "out")).FullName;
        var existing = Path.Combine(output, "收據_1120051.docx");
        File.WriteAllText(existing, "keep-existing");
        var result = await Batch(provider).GenerateBatchAsync(Workbook(workspace), "工作表1", Selection(), output);
        Assert.Equal(1, result.FailureCount);
        Assert.Equal(1, result.SuccessCount);
        Assert.Equal("keep-existing", File.ReadAllText(existing));
        AssertClean(provider, workspace);
    }

    [Fact]
    public async Task FatalOutputPreflightFailureCleansTemplate()
    {
        using var workspace = new Workspace();
        var provider = new TrackingProvider(workspace.Temp);
        var blocked = Path.Combine(workspace.Path, "not-directory");
        File.WriteAllText(blocked, "file");
        var error = await Assert.ThrowsAsync<ReceiptBatchGenerationException>(() =>
            Batch(provider).GenerateBatchAsync(Workbook(workspace), "工作表1", Selection(), blocked));
        Assert.Equal(ReceiptBatchFatalErrorCode.OutputDirectoryUnavailable, error.Code);
        AssertClean(provider, workspace);
    }

    [Fact]
    public async Task CancellationAfterFirstRowCleansBatchTemplate()
    {
        using var workspace = new Workspace();
        var provider = new TrackingProvider(workspace.Temp);
        using var cancellation = new CancellationTokenSource();
        var result = await Batch(provider).GenerateBatchAsync(
            Workbook(workspace), "工作表1", Selection(), Path.Combine(workspace.Path, "out"),
            new CallbackProgress(value => { if (value.Completed == 1) cancellation.Cancel(); }),
            cancellation.Token);
        Assert.True(result.IsCancelled);
        Assert.Equal(1, result.SuccessCount);
        Assert.Equal(1, result.UnprocessedCount);
        AssertClean(provider, workspace);
    }

    [Fact]
    public async Task PreCancelledBatchDoesNotExtract()
    {
        using var workspace = new Workspace();
        var provider = new TrackingProvider(workspace.Temp);
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        var result = await Batch(provider).GenerateBatchAsync(
            "unused.xlsx", "工作表1", Selection(), workspace.Path, cancellationToken: cancellation.Token);
        Assert.True(result.IsCancelled);
        Assert.Empty(provider.Paths);
        Assert.False(Directory.Exists(workspace.Temp));
        Assert.Throws<OperationCanceledException>(() =>
            new EmbeddedReceiptTemplateProvider(workspace.Temp).Acquire(cancellation.Token));
    }

    [Fact]
    public async Task InvalidExtractedTemplateIsCleanedBeforeBatchFails()
    {
        using var workspace = new Workspace();
        var provider = new TrackingProvider(workspace.Temp, corrupt: true);
        var error = await Assert.ThrowsAsync<ReceiptBatchGenerationException>(() =>
            Batch(provider).GenerateBatchAsync(Workbook(workspace), "工作表1", Selection(), workspace.Path));
        Assert.Equal(ReceiptBatchFatalErrorCode.InternalTemplateInvalid, error.Code);
        AssertClean(provider, workspace);
    }

    [Fact]
    public async Task UnexpectedProgressFailureStillCleansTemplate()
    {
        using var workspace = new Workspace();
        var provider = new TrackingProvider(workspace.Temp);
        await Assert.ThrowsAsync<InvalidOperationException>(() =>
            Batch(provider).GenerateBatchAsync(Workbook(workspace), "工作表1", Selection(), workspace.Path,
                new CallbackProgress(_ => throw new InvalidOperationException("progress-failure"))));
        AssertClean(provider, workspace);
    }

    [Fact]
    public async Task SingleReceiptOutputFailureCleansTemplate()
    {
        using var workspace = new Workspace();
        var provider = new TrackingProvider(workspace.Temp);
        var error = await Assert.ThrowsAsync<ReceiptGenerationException>(() =>
            new ReceiptGenerationService(templateProvider: provider).GenerateReceiptAsync(
                Record(), Path.Combine(workspace.Path, "invalid.txt")));
        Assert.Equal(ReceiptGenerationErrorCode.InvalidOutputPath, error.Code);
        AssertClean(provider, workspace);
    }

    [Fact]
    public void CleanupFailureNeverEscapesDispose()
    {
        using var workspace = new Workspace();
        var lease = new EmbeddedReceiptTemplateProvider(workspace.Temp).Acquire();
        using (var locked = new FileStream(lease.TemplatePath, FileMode.Open, FileAccess.Read, FileShare.None))
        {
            Assert.Null(Xunit.Record.Exception(lease.Dispose));
            Assert.True(File.Exists(lease.TemplatePath));
        }
        // Windows releases the deliberate test lock here; Workspace owns cleanup.
    }

    [Fact]
    public void ExtractionIoFailureIsAFriendlyTemplateFailure()
    {
        using var workspace = new Workspace();
        File.WriteAllText(workspace.Temp, "not-directory");
        var error = Assert.Throws<ReceiptGenerationException>(() =>
            new EmbeddedReceiptTemplateProvider(workspace.Temp).Acquire());
        Assert.Equal(ReceiptGenerationErrorCode.ReceiptTemplateInvalid, error.Code);
        Assert.Equal("not-directory", File.ReadAllText(workspace.Temp));
    }

    [Fact]
    public void BuildEmbedsTemplateWithoutExternalRuntimeContent()
    {
        var project = XDocument.Load(Path.Combine(Root(), "src/XlsxDocxGenerator/XlsxDocxGenerator.csproj"));
        var item = Assert.Single(project.Descendants("EmbeddedResource"));
        Assert.Equal("ReceiptXcel.Templates.Receipt.docx", (string?)item.Attribute("LogicalName"));
        Assert.Empty(project.Descendants("CopyToOutputDirectory"));
        Assert.Empty(project.Descendants("CopyToPublishDirectory"));
        using var lease = new EmbeddedReceiptTemplateProvider().Acquire();
        Assert.False(lease.TemplatePath.StartsWith(AppContext.BaseDirectory, StringComparison.OrdinalIgnoreCase));
        Assert.DoesNotContain("Assets", lease.TemplatePath, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void SingleFileProfileUsesUntrimmedSelfContainedNativeExtraction()
    {
        var profile = XDocument.Load(Path.Combine(Root(),
            "src/XlsxDocxGenerator/Properties/PublishProfiles/WinX64SingleFile.pubxml"));
        foreach (var (key, value) in new[]
        {
            ("RuntimeIdentifier", "win-x64"), ("SelfContained", "true"), ("PublishSingleFile", "true"),
            ("PublishTrimmed", "false"), ("IncludeNativeLibrariesForSelfExtract", "true"),
            ("EnableCompressionInSingleFile", "false"), ("DebugType", "none")
        })
            Assert.Equal(value, profile.Descendants(key).Single().Value);
        Assert.Empty(profile.Descendants("PublishAot"));
        Assert.Contains("win-x64-single-file-rc", profile.Descendants("PublishDir").Single().Value);
        Assert.True(File.Exists(Path.Combine(Root(),
            "src/XlsxDocxGenerator/Properties/PublishProfiles/WinX64Portable.pubxml")));
    }

    private static void AssertReceipt(string path)
    {
        using var document = WordprocessingDocument.Open(path, false);
        var body = document.MainDocumentPart!.Document.Body!;
        var text = body.InnerText;
        Assert.DoesNotContain("{{", text);
        Assert.DoesNotContain("6/6匯款", text);
        Assert.Contains("21813799", text);
        foreach (var table in body.Elements<Table>())
        {
            var rows = table.Elements<TableRow>().ToArray();
            Assert.Equal("運動會禮金", rows[2].Elements<TableCell>().ElementAt(9).InnerText);
            Assert.Equal("", rows[4].Elements<TableCell>().ElementAt(1).InnerText);
        }
        Assert.Equal(3, text.Split("壹仟元整").Length - 1);
    }

    private static void AssertLines(string path)
    {
        using var document = WordprocessingDocument.Open(path, false);
        var xml = XDocument.Parse(document.MainDocumentPart!.Document.OuterXml);
        var anchors = xml.Descendants(Wp + "anchor").ToArray();
        Assert.Equal(2, anchors.Length);
        Assert.Equal(["119380", "274955"], anchors.Select(anchor => anchor.Element(Wp + "positionV")!
            .Element(Wp + "posOffset")!.Value).ToArray());
        foreach (var anchor in anchors)
        {
            Assert.NotNull(anchor.Element(Wp + "wrapNone"));
            Assert.Equal("6840220", (string?)anchor.Element(Wp + "extent")!.Attribute("cx"));
            var line = anchor.Descendants(A + "ln").Single();
            Assert.Equal("9525", (string?)line.Attribute("w"));
            Assert.Equal("dash", (string?)line.Element(A + "prstDash")!.Attribute("val"));
            Assert.Equal("808080", (string?)line.Descendants(A + "srgbClr").Single().Attribute("val"));
            Assert.Equal("straightConnector1", (string?)anchor.Descendants(A + "prstGeom").Single().Attribute("prst"));
        }
    }

    private static void AssertClean(TrackingProvider provider, Workspace workspace)
    {
        var path = Assert.Single(provider.Paths);
        Assert.False(File.Exists(path));
        Assert.False(Directory.Exists(Path.GetDirectoryName(path)));
        Assert.Empty(Directory.EnumerateFileSystemEntries(workspace.Temp));
    }

    private static ReceiptRecord Record() => new()
    {
        ExcelRowNumber = 3, ROCYear = 112, Month = 12, Day = 30, ReceiptSerial = "0051",
        Payer = "SINGLE_FILE_QA", Amount = 1000, Reason = "運動會禮金", Handler = "6/6匯款"
    };

    private static ReceiptSelectionRequest Selection() => new(ReceiptSelectionMode.ExcelRows, "3,4");
    private static ReceiptBatchGenerationService Batch(IReceiptTemplateProvider provider) =>
        new(receiptGenerationService: new ReceiptGenerationService(templateProvider: provider));

    private static string Workbook(Workspace workspace)
    {
        var path = System.IO.Path.Combine(workspace.Path, "input.xlsx");
        using var workbook = new XLWorkbook();
        var sheet = workbook.Worksheets.Add("工作表1");
        for (var i = 0; i < ReceiptWorksheetSchema.ExpectedHeaders.Count; i++)
            sheet.Cell(2, i + 1).Value = ReceiptWorksheetSchema.ExpectedHeaders[i];
        for (var row = 3; row <= 4; row++)
        {
            sheet.Cell(row, 1).Value = 112;
            sheet.Cell(row, 2).Value = 12;
            sheet.Cell(row, 3).Value = 30;
            sheet.Cell(row, 4).Value = row == 3 ? "0051" : "0052";
            sheet.Cell(row, 5).Value = "SINGLE_FILE_QA";
            sheet.Cell(row, 6).Value = 1000;
            sheet.Cell(row, 7).Value = "運動會禮金";
            sheet.Cell(row, 8).Value = "6/6匯款";
        }
        workbook.SaveAs(path);
        return path;
    }

    private static string Root()
    {
        for (var directory = new DirectoryInfo(AppContext.BaseDirectory); directory is not null; directory = directory.Parent)
            if (File.Exists(Path.Combine(directory.FullName, "ReceiptXcel.sln"))) return directory.FullName;
        throw new DirectoryNotFoundException();
    }

    private sealed class TrackingProvider(string root, bool corrupt = false) : IReceiptTemplateProvider
    {
        public List<string> Paths { get; } = [];
        public ReceiptTemplateLease Acquire(CancellationToken cancellationToken = default)
        {
            var lease = new EmbeddedReceiptTemplateProvider(root).Acquire(cancellationToken);
            Paths.Add(lease.TemplatePath);
            if (corrupt) File.WriteAllText(lease.TemplatePath, "invalid-docx");
            return lease;
        }
    }

    private sealed class CallbackProgress(Action<ReceiptBatchProgress> callback) : IProgress<ReceiptBatchProgress>
    {
        public void Report(ReceiptBatchProgress value) => callback(value);
    }

    private sealed class Workspace : IDisposable
    {
        public string Path { get; } = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "receiptxcel-embedded-tests-" + Guid.NewGuid().ToString("N"));
        public string Temp => System.IO.Path.Combine(Path, "templates");
        public Workspace() => Directory.CreateDirectory(Path);
        public void Dispose() => Directory.Delete(Path, recursive: true);
    }
}
