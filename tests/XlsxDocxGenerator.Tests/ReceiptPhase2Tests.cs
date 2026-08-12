using ClosedXML.Excel;
using XlsxDocxGenerator.Models;
using XlsxDocxGenerator.Services.Diagnostics;
using XlsxDocxGenerator.Services.Excel;
using XlsxDocxGenerator.Services.Receipts;
using XlsxDocxGenerator.Services.Settings;
using XlsxDocxGenerator.ViewModels;

namespace XlsxDocxGenerator.Tests;

public sealed class ReceiptPhase2Tests
{
    [Fact]
    public void StartupWithoutExcelCannotGenerate()
    {
        using var workspace = new TemporaryWorkspace();
        var viewModel = CreateViewModel(workspace.Path);

        Assert.False(viewModel.CanGenerate);
        Assert.Null(viewModel.ExcelSchemaValid);
    }

    [Fact]
    public async Task ValidExcelAutomaticallyPassesFixedSchemaValidation()
    {
        using var workspace = new TemporaryWorkspace();
        var viewModel = CreateViewModel(workspace.Path);

        await viewModel.LoadExcelAsync(CreateWorkbook(workspace.Path, ValidRows(3)));
        await WaitForAsync(() => viewModel.ExcelFormatStatusText == "✓ Excel 格式正確");

        Assert.True(viewModel.ExcelSchemaValid);
        Assert.Contains("可以開始產生收據", viewModel.StatusText);
    }

    [Fact]
    public async Task WorksheetOneIsSelectedAutomaticallyWhenAvailable()
    {
        using var workspace = new TemporaryWorkspace();
        var workbook = CreateWorkbook(workspace.Path, ValidRows(3), includeOtherSheet: true);
        var viewModel = CreateViewModel(workspace.Path);

        await viewModel.LoadExcelAsync(workbook);

        Assert.Equal(ReceiptWorksheetSchema.WorksheetName, viewModel.SelectedWorksheet);
        Assert.Contains(ReceiptWorksheetSchema.WorksheetName, viewModel.WorksheetNames);
    }

    [Fact]
    public async Task PreviewCountsRowsAfterHeaderOnly()
    {
        using var workspace = new TemporaryWorkspace();
        var viewModel = CreateViewModel(workspace.Path);

        await PrepareLatestAsync(viewModel, CreateWorkbook(workspace.Path, ValidRows(3, 4, 6)), 3);

        Assert.Contains("找到 3 筆收據資料", viewModel.SelectionSummaryText);
        Assert.Contains("符合條件：3 筆", viewModel.SelectionSummaryText);
    }

    [Fact]
    public void DefaultSelectionModeIsLatest()
    {
        using var workspace = new TemporaryWorkspace();
        var viewModel = CreateViewModel(workspace.Path);

        Assert.Equal(ReceiptSelectionMode.Latest, viewModel.SelectedMode.Mode);
        Assert.True(viewModel.IsLatestMode);
        Assert.False(viewModel.IsExcelRowsMode);
    }

    [Fact]
    public async Task LatestOneSelectsTheLastAvailableRow()
    {
        using var workspace = new TemporaryWorkspace();
        var viewModel = CreateViewModel(workspace.Path);

        await PrepareLatestAsync(viewModel, CreateWorkbook(workspace.Path, ValidRows(3, 4, 6)), 1);

        Assert.Equal([6], viewModel.PreviewRowNumbers);
    }

    [Fact]
    public async Task LatestNSelectsRowsInAscendingExcelOrder()
    {
        using var workspace = new TemporaryWorkspace();
        var viewModel = CreateViewModel(workspace.Path);

        await PrepareLatestAsync(viewModel, CreateWorkbook(workspace.Path, ValidRows(3, 4, 6)), 2);

        Assert.Equal([4, 6], viewModel.PreviewRowNumbers);
    }

    [Fact]
    public async Task LatestMoreThanAvailableUsesAllRowsAndShowsNotice()
    {
        using var workspace = new TemporaryWorkspace();
        var viewModel = CreateViewModel(workspace.Path);

        await PrepareLatestAsync(viewModel, CreateWorkbook(workspace.Path, ValidRows(3, 4)), 20, 2);

        Assert.Equal([3, 4], viewModel.PreviewRowNumbers);
        Assert.Contains("要求 20 筆，實際找到 2 筆", viewModel.SelectionSummaryText);
    }

    [Fact]
    public async Task ExcelRowsSingleExpressionSelectsRequestedRow()
    {
        using var workspace = new TemporaryWorkspace();
        var viewModel = CreateViewModel(workspace.Path);
        await viewModel.LoadExcelAsync(CreateWorkbook(workspace.Path, ValidRows(3, 4, 6)));
        viewModel.SelectedMode = viewModel.SelectionModes.Single(item => item.Mode == ReceiptSelectionMode.ExcelRows);
        viewModel.RowExpressionText = "6";

        await WaitForAsync(() => viewModel.PreviewRowNumbers.SequenceEqual([6]));
        Assert.Equal([6], viewModel.PreviewRowNumbers);
    }

    [Fact]
    public async Task ExcelRowsListAndRangeAreDistinctAndAscending()
    {
        using var workspace = new TemporaryWorkspace();
        var viewModel = CreateViewModel(workspace.Path);
        await viewModel.LoadExcelAsync(CreateWorkbook(workspace.Path, ValidRows(3, 4, 5, 6, 8, 9, 10)));
        viewModel.SelectedMode = viewModel.SelectionModes.Single(item => item.Mode == ReceiptSelectionMode.ExcelRows);
        viewModel.RowExpressionText = "10, 3, 5, 8-10";

        await WaitForAsync(() => viewModel.PreviewRowNumbers.SequenceEqual([3, 5, 8, 9, 10]));
        Assert.Equal([3, 5, 8, 9, 10], viewModel.PreviewRowNumbers);
    }

    [Fact]
    public async Task ExcelRowOneIsRejectedWithFriendlyMessage()
    {
        var preview = await PreviewRowsAsync("1");

        Assert.False(preview.SelectedRows.Single().IsValid);
        Assert.Contains("第 1 列", preview.SelectedRows.Single().UserMessage);
        Assert.Contains("不是收據資料", preview.SelectedRows.Single().UserMessage);
    }

    [Fact]
    public async Task ExcelRowTwoIsRejectedAsHeader()
    {
        var preview = await PreviewRowsAsync("2");

        Assert.False(preview.SelectedRows.Single().IsValid);
        Assert.Equal("第 2 列是欄位標題，不是收據資料。", preview.SelectedRows.Single().UserMessage);
    }

    [Fact]
    public async Task MalformedExcelRowsExpressionDoesNotEnableGeneration()
    {
        using var workspace = new TemporaryWorkspace();
        var viewModel = CreateViewModel(workspace.Path);
        await viewModel.LoadExcelAsync(CreateWorkbook(workspace.Path, ValidRows(3)));
        viewModel.SelectedMode = viewModel.SelectionModes.Single(item => item.Mode == ReceiptSelectionMode.ExcelRows);
        viewModel.RowExpressionText = "3-";

        await WaitForAsync(() => viewModel.StatusText.Contains("無法解析", StringComparison.Ordinal));

        Assert.False(viewModel.CanGenerate);
    }

    [Fact]
    public async Task ChangingExcelImmediatelyInvalidatesPreviousPreview()
    {
        using var workspace = new TemporaryWorkspace();
        var viewModel = CreateViewModel(workspace.Path);
        await PrepareLatestAsync(viewModel, CreateWorkbook(workspace.Path, ValidRows(3)), 1);
        var first = viewModel.PreviewRowNumbers;

        var loadTask = viewModel.LoadExcelAsync(CreateWorkbook(workspace.Path, ValidRows(3, 4)));

        Assert.False(viewModel.CanGenerate);
        await loadTask;
        Assert.Equal(first, new[] { 3 });
    }

    [Fact]
    public async Task ChangingWorksheetImmediatelyInvalidatesPreview()
    {
        using var workspace = new TemporaryWorkspace();
        var workbook = CreateWorkbook(workspace.Path, ValidRows(3), includeOtherSheet: true);
        var viewModel = CreateViewModel(workspace.Path);
        await PrepareLatestAsync(viewModel, workbook, 1);

        viewModel.SelectedWorksheet = "Other";

        Assert.False(viewModel.CanGenerate);
        await WaitForAsync(() => viewModel.ExcelFormatStatusText.Contains("不符", StringComparison.Ordinal));
    }

    [Fact]
    public async Task SelectingExcelPrefillsOutputDirectoryBesideWorkbook()
    {
        using var workspace = new TemporaryWorkspace();
        var workbook = CreateWorkbook(workspace.Path, ValidRows(3));
        var viewModel = CreateViewModel(workspace.Path);

        await viewModel.LoadExcelAsync(workbook);

        Assert.Equal(Path.Combine(workspace.Path, "ReceiptXcel輸出"), viewModel.OutputDirectory);
    }

    [Fact]
    public async Task AutomaticOutputDirectoryIsNotCreatedDuringPreview()
    {
        using var workspace = new TemporaryWorkspace();
        var workbook = CreateWorkbook(workspace.Path, ValidRows(3));
        var output = Path.Combine(workspace.Path, "ReceiptXcel輸出");
        var viewModel = CreateViewModel(workspace.Path);

        await PrepareLatestAsync(viewModel, workbook, 1);

        Assert.False(Directory.Exists(output));
    }

    [Fact]
    public async Task UserOutputOverrideSurvivesExcelAndSelectionChanges()
    {
        using var workspace = new TemporaryWorkspace();
        var custom = Path.Combine(workspace.Path, "CustomOutput");
        var viewModel = CreateViewModel(workspace.Path);
        viewModel.OutputDirectory = custom;

        await viewModel.LoadExcelAsync(CreateWorkbook(workspace.Path, ValidRows(3, 4)));
        viewModel.LatestCountText = "1";
        await WaitForAsync(() => viewModel.PreviewRowNumbers.SequenceEqual([4]));
        await viewModel.LoadExcelAsync(CreateWorkbook(workspace.Path, ValidRows(3, 4, 6)));

        Assert.Equal(custom, viewModel.OutputDirectory);
    }

    [Fact]
    public async Task BatchUsesFixedReceiptFilenamePolicy()
    {
        using var workspace = new TemporaryWorkspace();
        var result = await CreateBatchService().GenerateBatchAsync(
            CreateWorkbook(workspace.Path, ValidRows(3)),
            ReceiptWorksheetSchema.WorksheetName,
            new ReceiptSelectionRequest(ReceiptSelectionMode.ExcelRows, "3"),
            Path.Combine(workspace.Path, "Output"));

        Assert.True(result.Results.Single().Success);
        Assert.Equal("收據_1120051.docx", Path.GetFileName(result.Results.Single().OutputPath));
    }

    [Fact]
    public async Task DuplicateReceiptNumbersBlockBatchBeforeAnyOutput()
    {
        using var workspace = new TemporaryWorkspace();
        var output = Path.Combine(workspace.Path, "Output");
        var exception = await Assert.ThrowsAsync<ReceiptBatchGenerationException>(() =>
            CreateBatchService().GenerateBatchAsync(
                CreateWorkbook(workspace.Path, ValidRows(3)[0], ValidRows(8)[0] with { Serial = "0051" }),
                ReceiptWorksheetSchema.WorksheetName,
                new ReceiptSelectionRequest(ReceiptSelectionMode.ExcelRows, "3,8"),
                output));

        Assert.Equal(ReceiptBatchFatalErrorCode.DuplicateReceiptNumber, exception.Code);
        Assert.Contains("1120051", exception.UserMessage);
        Assert.False(Directory.Exists(output));
    }

    [Fact]
    public async Task ExistingOutputFailsOnlyThatRow()
    {
        using var workspace = new TemporaryWorkspace();
        var output = Path.Combine(workspace.Path, "Output");
        Directory.CreateDirectory(output);
        File.WriteAllText(Path.Combine(output, "收據_1120051.docx"), "keep");

        var result = await CreateBatchService().GenerateBatchAsync(
            CreateWorkbook(workspace.Path, ValidRows(3, 4)),
            ReceiptWorksheetSchema.WorksheetName,
            new ReceiptSelectionRequest(ReceiptSelectionMode.ExcelRows, "3,4"),
            output);

        var failed = result.Results.Single(item => item.ExcelRowNumber == 3);
        var succeeded = result.Results.Single(item => item.ExcelRowNumber == 4);
        Assert.False(failed.Success);
        Assert.Equal(nameof(ReceiptGenerationErrorCode.OutputAlreadyExists), failed.ErrorCode);
        Assert.Contains("未覆寫", failed.UserMessage);
        Assert.True(succeeded.Success);
        Assert.Equal("keep", File.ReadAllText(Path.Combine(output, "收據_1120051.docx")));
    }

    [Fact]
    public async Task InvalidAmountRowFailsAndLaterValidRowContinues()
    {
        using var workspace = new TemporaryWorkspace();
        var result = await CreateBatchService().GenerateBatchAsync(
            CreateWorkbook(workspace.Path, CombineRows(ValidRows(3), [InvalidAmountRow(4)], ValidRows(5))),
            ReceiptWorksheetSchema.WorksheetName,
            new ReceiptSelectionRequest(ReceiptSelectionMode.ExcelRows, "3,4,5"),
            Path.Combine(workspace.Path, "Output"));

        Assert.Equal(2, result.SuccessCount);
        Assert.Equal(1, result.FailureCount);
        Assert.True(result.Results.Single(item => item.ExcelRowNumber == 3).Success);
        Assert.False(result.Results.Single(item => item.ExcelRowNumber == 4).Success);
        Assert.True(result.Results.Single(item => item.ExcelRowNumber == 5).Success);
    }

    [Fact]
    public async Task OverLimitRowUsesFriendlyFailureMessage()
    {
        using var workspace = new TemporaryWorkspace();
        var result = await CreateBatchService().GenerateBatchAsync(
            CreateWorkbook(workspace.Path, CombineRows(ValidRows(3), [OverLimitRow(4)], ValidRows(5))),
            ReceiptWorksheetSchema.WorksheetName,
            new ReceiptSelectionRequest(ReceiptSelectionMode.ExcelRows, "3,4,5"),
            Path.Combine(workspace.Path, "Output"));

        var failed = result.Results.Single(item => item.ExcelRowNumber == 4);
        Assert.False(failed.Success);
        Assert.Equal(nameof(ReceiptValidationErrorCode.AmountExceedsLimit), failed.ErrorCode);
        Assert.Contains("超過表單上限 9,999,999", failed.UserMessage);
    }

    [Fact]
    public async Task BatchResultReportsSuccessFailureAndTotalCounts()
    {
        using var workspace = new TemporaryWorkspace();
        var result = await CreateBatchService().GenerateBatchAsync(
            CreateWorkbook(workspace.Path, CombineRows(ValidRows(3), [InvalidAmountRow(4)], ValidRows(5))),
            ReceiptWorksheetSchema.WorksheetName,
            new ReceiptSelectionRequest(ReceiptSelectionMode.ExcelRows, "3,4,5"),
            Path.Combine(workspace.Path, "Output"));

        Assert.Equal(ReceiptBatchStatus.Completed, result.Status);
        Assert.Equal(3, result.TotalSelected);
        Assert.Equal(2, result.SuccessCount);
        Assert.Equal(1, result.FailureCount);
        Assert.Equal(0, result.UnprocessedCount);
    }

    [Fact]
    public async Task CancellationBeforeFirstDoesNotCreateOutput()
    {
        using var workspace = new TemporaryWorkspace();
        var output = Path.Combine(workspace.Path, "Output");
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();

        var result = await CreateBatchService().GenerateBatchAsync(
            CreateWorkbook(workspace.Path, ValidRows(3)),
            ReceiptWorksheetSchema.WorksheetName,
            new ReceiptSelectionRequest(ReceiptSelectionMode.ExcelRows, "3"),
            output,
            cancellationToken: cancellation.Token);

        Assert.True(result.IsCancelled);
        Assert.Equal(0, result.SuccessCount);
        Assert.False(Directory.Exists(output));
    }

    [Fact]
    public async Task MidBatchCancellationPreservesCompletedRowsAndReportsUnprocessed()
    {
        using var workspace = new TemporaryWorkspace();
        using var cancellation = new CancellationTokenSource();
        var progress = new CancelAfterFirstProgress(cancellation);
        var result = await CreateBatchService().GenerateBatchAsync(
            CreateWorkbook(workspace.Path, ValidRows(3, 4, 5)),
            ReceiptWorksheetSchema.WorksheetName,
            new ReceiptSelectionRequest(ReceiptSelectionMode.ExcelRows, "3,4,5"),
            Path.Combine(workspace.Path, "Output"),
            progress,
            cancellation.Token);

        Assert.True(result.IsCancelled);
        Assert.Equal(1, result.SuccessCount);
        Assert.Equal(2, result.UnprocessedCount);
        Assert.True(File.Exists(result.Results.Single().OutputPath));
    }

    [Fact]
    public async Task CancellationCleansTemporaryFiles()
    {
        using var workspace = new TemporaryWorkspace();
        using var cancellation = new CancellationTokenSource();
        var output = Path.Combine(workspace.Path, "Output");
        var result = await CreateBatchService().GenerateBatchAsync(
            CreateWorkbook(workspace.Path, ValidRows(3, 4)),
            ReceiptWorksheetSchema.WorksheetName,
            new ReceiptSelectionRequest(ReceiptSelectionMode.ExcelRows, "3,4"),
            output,
            new CancelAfterFirstProgress(cancellation),
            cancellation.Token);

        Assert.True(result.IsCancelled);
        Assert.Empty(Directory.GetFiles(output, ".*.tmp"));
    }

    [Fact]
    public void SettingsUseReceiptXcelLocalAppDataBoundary()
    {
        var path = new SettingsService().SettingsPath;

        Assert.Contains(Path.Combine("ReceiptXcel", "settings.json"), path, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("DocXcel", path, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void DiagnosticsUseReceiptXcelLocalAppDataBoundary()
    {
        var path = new DiagnosticsLogger().LogDirectory;

        Assert.Contains(Path.Combine("ReceiptXcel", "Logs"), path, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("DocXcel", path, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void InternalTemplateIsAvailableThroughSpecializedService()
    {
        var service = CreateBatchService();

        Assert.True(service.IsInternalTemplateAvailable(out var message));
        Assert.Contains("可用", message);
    }

    [Fact]
    public void MainWindowXamlUsesReadonlyProgressBindings()
    {
        var xaml = File.ReadAllText(Path.Combine(RepositoryRoot(), "src", "XlsxDocxGenerator", "MainWindow.xaml"));

        Assert.Contains("Maximum=\"{Binding ProgressMaximum, Mode=OneWay}\"", xaml);
        Assert.Contains("Value=\"{Binding ProgressValue, Mode=OneWay}\"", xaml);
        Assert.DoesNotContain("Value=\"{Binding ProgressValue}\"", xaml);
    }

    [Fact]
    public void MainWindowUsesReceiptProductTitle()
    {
        var xaml = File.ReadAllText(Path.Combine(RepositoryRoot(), "src", "XlsxDocxGenerator", "MainWindow.xaml"));

        Assert.Contains("ReceiptXcel｜自行收納款項收據產生工具", xaml);
        Assert.Contains("產生收據", xaml);
    }

    [Fact]
    public void UserFacingUiContainsNoUniversalTemplateWorkflow()
    {
        var xaml = File.ReadAllText(Path.Combine(RepositoryRoot(), "src", "XlsxDocxGenerator", "MainWindow.xaml"));

        Assert.DoesNotContain("建立模板", xaml);
        Assert.DoesNotContain("載入模板", xaml);
        Assert.DoesNotContain("欄位對應", xaml);
        Assert.DoesNotContain("marker", xaml, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain(".docxcel.json", xaml, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void UserFacingUiContainsNoFilenamePatternEditor()
    {
        var xaml = File.ReadAllText(Path.Combine(RepositoryRoot(), "src", "XlsxDocxGenerator", "MainWindow.xaml"));

        Assert.DoesNotContain("輸出檔名", xaml);
        Assert.DoesNotContain("{{ReceiptNumber}}", xaml, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("固定檔名", xaml);
    }

    [Fact]
    public async Task SpecializedViewModelCanGenerateAfterValidLatestPreview()
    {
        using var workspace = new TemporaryWorkspace();
        var viewModel = CreateViewModel(workspace.Path);
        viewModel.OutputDirectory = Path.Combine(workspace.Path, "Output");

        await PrepareLatestAsync(viewModel, CreateWorkbook(workspace.Path, ValidRows(3)), 1);

        Assert.True(viewModel.CanGenerate);
    }

    [Fact]
    public async Task BlankPayerIsRowFailureWithoutHandlerWarning()
    {
        using var workspace = new TemporaryWorkspace();
        var result = await CreateBatchService().GenerateBatchAsync(
            CreateWorkbook(workspace.Path, ValidRows(3)[0] with { Payer = string.Empty }),
            ReceiptWorksheetSchema.WorksheetName,
            new ReceiptSelectionRequest(ReceiptSelectionMode.ExcelRows, "3"),
            Path.Combine(workspace.Path, "Output"));

        Assert.False(result.Results.Single().Success);
        Assert.Contains("繳款人", result.Results.Single().UserMessage);
        Assert.DoesNotContain("承辦人", result.Results.Single().UserMessage);
    }

    [Fact]
    public async Task BlankReasonIsRowFailureAndBlankHandlerIsAllowed()
    {
        using var workspace = new TemporaryWorkspace();
        var row = ValidRows(3)[0] with { Reason = string.Empty };
        var result = await CreateBatchService().GenerateBatchAsync(
            CreateWorkbook(workspace.Path, row),
            ReceiptWorksheetSchema.WorksheetName,
            new ReceiptSelectionRequest(ReceiptSelectionMode.ExcelRows, "3"),
            Path.Combine(workspace.Path, "Output"));

        Assert.False(result.Results.Single().Success);
        Assert.Contains("事由", result.Results.Single().UserMessage);
    }

    [Fact]
    public async Task InvalidSchemaBlocksBatchBeforeOutputDirectoryCreation()
    {
        using var workspace = new TemporaryWorkspace();
        var workbook = CreateWorkbook(workspace.Path, ValidRows(3), wrongHeader: true);
        var output = Path.Combine(workspace.Path, "Output");

        var exception = await Assert.ThrowsAsync<ReceiptBatchGenerationException>(() =>
            CreateBatchService().GenerateBatchAsync(
                workbook,
                ReceiptWorksheetSchema.WorksheetName,
                new ReceiptSelectionRequest(ReceiptSelectionMode.Latest, "1"),
                output));

        Assert.Equal(ReceiptBatchFatalErrorCode.ExcelSchemaMismatch, exception.Code);
        Assert.False(Directory.Exists(output));
    }

    [Fact]
    public async Task MissingInternalTemplateBlocksBatchBeforeOutput()
    {
        using var workspace = new TemporaryWorkspace();
        var missingTemplate = Path.Combine(workspace.Path, "missing.docx");
        var service = new ReceiptBatchGenerationService(
            receiptGenerationService: new ReceiptGenerationService(internalTemplatePath: missingTemplate));
        var output = Path.Combine(workspace.Path, "Output");

        var exception = await Assert.ThrowsAsync<ReceiptBatchGenerationException>(() =>
            service.GenerateBatchAsync(
                CreateWorkbook(workspace.Path, ValidRows(3)),
                ReceiptWorksheetSchema.WorksheetName,
                new ReceiptSelectionRequest(ReceiptSelectionMode.Latest, "1"),
                output));

        Assert.Equal(ReceiptBatchFatalErrorCode.InternalTemplateInvalid, exception.Code);
        Assert.Contains("重新安裝 ReceiptXcel", exception.UserMessage);
        Assert.False(Directory.Exists(output));
    }

    [Fact]
    public async Task BatchCanGenerateMultipleThreePartDocuments()
    {
        using var workspace = new TemporaryWorkspace();
        var output = Path.Combine(workspace.Path, "Output");
        var result = await CreateBatchService().GenerateBatchAsync(
            CreateWorkbook(workspace.Path, ValidRows(3, 4, 5)),
            ReceiptWorksheetSchema.WorksheetName,
            new ReceiptSelectionRequest(ReceiptSelectionMode.Latest, "3"),
            output);

        Assert.Equal(3, result.SuccessCount);
        Assert.Equal(
            ["收據_1120051.docx", "收據_1120052.docx", "收據_1120053.docx"],
            Directory.GetFiles(output, "*.docx").Select(Path.GetFileName).OrderBy(name => name));
    }

    [Fact]
    public async Task Phase2QaAssetsAreGeneratedForManualReview()
    {
        var qaDirectory = Path.Combine(RepositoryRoot(), "artifacts", "qa", "phase2");
        Directory.CreateDirectory(qaDirectory);
        var workbookPath = Path.Combine(qaDirectory, "receipt-phase2-qa.xlsx");
        if (File.Exists(workbookPath))
        {
            File.Delete(workbookPath);
        }

        CreateWorkbookFile(workbookPath, ValidRows(3, 4, 5, 6));
        var service = CreateBatchService();
        var latestDirectory = Path.Combine(qaDirectory, "latest-1");
        var rowsDirectory = Path.Combine(qaDirectory, "rows-3");
        var batchDirectory = Path.Combine(qaDirectory, "batch");
        var cancellationDirectory = Path.Combine(qaDirectory, "cancellation");
        foreach (var directory in new[] { latestDirectory, rowsDirectory, batchDirectory, cancellationDirectory })
        {
            Directory.CreateDirectory(directory);
            foreach (var file in Directory.GetFiles(directory, "*.docx"))
            {
                File.Delete(file);
            }
        }

        var latest = await service.GenerateBatchAsync(
            workbookPath,
            ReceiptWorksheetSchema.WorksheetName,
            new ReceiptSelectionRequest(ReceiptSelectionMode.Latest, "1"),
            latestDirectory);
        var row = await service.GenerateBatchAsync(
            workbookPath,
            ReceiptWorksheetSchema.WorksheetName,
            new ReceiptSelectionRequest(ReceiptSelectionMode.ExcelRows, "3"),
            rowsDirectory);
        var batch = await service.GenerateBatchAsync(
            workbookPath,
            ReceiptWorksheetSchema.WorksheetName,
            new ReceiptSelectionRequest(ReceiptSelectionMode.ExcelRows, "3,4,5"),
            batchDirectory);
        using var cancellation = new CancellationTokenSource();
        var cancelled = await service.GenerateBatchAsync(
            workbookPath,
            ReceiptWorksheetSchema.WorksheetName,
            new ReceiptSelectionRequest(ReceiptSelectionMode.ExcelRows, "3,4,5,6"),
            cancellationDirectory,
            new CancelAfterFirstProgress(cancellation),
            cancellation.Token);

        Assert.Equal(1, latest.SuccessCount);
        Assert.Equal(1, row.SuccessCount);
        Assert.Equal(3, batch.SuccessCount);
        Assert.True(cancelled.IsCancelled);
        Assert.Equal(1, cancelled.SuccessCount);
        Assert.Single(Directory.GetFiles(latestDirectory, "*.docx"));
        Assert.Single(Directory.GetFiles(rowsDirectory, "*.docx"));
        Assert.Equal(3, Directory.GetFiles(batchDirectory, "*.docx").Length);
        Assert.Single(Directory.GetFiles(cancellationDirectory, "*.docx"));
    }

    private static ReceiptMainWindowViewModel CreateViewModel(string directory) =>
        new(
            new SettingsService(Path.Combine(directory, "settings.json")),
            new DiagnosticsLogger(Path.Combine(directory, "Logs")),
            CreateBatchService());

    private static ReceiptBatchGenerationService CreateBatchService() =>
        new(receiptGenerationService: new ReceiptGenerationService(internalTemplatePath: TemplatePath()));

    private static async Task PrepareLatestAsync(
        ReceiptMainWindowViewModel viewModel,
        string workbook,
        int count,
        int? expectedSelectedCount = null)
    {
        await viewModel.LoadExcelAsync(workbook);
        viewModel.LatestCountText = count.ToString();
        await WaitForAsync(() =>
            !viewModel.IsPreviewRefreshing
            && viewModel.ExcelSchemaValid == true
            && viewModel.PreviewRowNumbers.Count == (expectedSelectedCount ?? count));
    }

    private static async Task<ReceiptBatchPreview> PreviewRowsAsync(string expression)
    {
        using var workspace = new TemporaryWorkspace();
        var service = CreateBatchService();
        return await service.PreviewAsync(
            CreateWorkbook(workspace.Path, ValidRows(3)),
            ReceiptWorksheetSchema.WorksheetName,
            new ReceiptSelectionRequest(ReceiptSelectionMode.ExcelRows, expression));
    }

    private static async Task WaitForAsync(Func<bool> condition)
    {
        var deadline = DateTime.UtcNow.AddSeconds(8);
        while (!condition() && DateTime.UtcNow < deadline)
        {
            await Task.Delay(30);
        }

        Assert.True(condition(), "ReceiptXcel preview condition timed out.");
    }

    private static string CreateWorkbook(
        string directory,
        params RowSpec[] rows) =>
        CreateWorkbook(directory, rows, includeOtherSheet: false, wrongHeader: false);

    private static string CreateWorkbook(
        string directory,
        IReadOnlyList<RowSpec> rows,
        bool includeOtherSheet = false,
        bool wrongHeader = false)
    {
        var path = Path.Combine(directory, $"receipt-{Guid.NewGuid():N}.xlsx");
        CreateWorkbookFile(path, rows, includeOtherSheet, wrongHeader);
        return path;
    }

    private static void CreateWorkbookFile(
        string path,
        IReadOnlyList<RowSpec> rows,
        bool includeOtherSheet = false,
        bool wrongHeader = false)
    {
        using var workbook = new XLWorkbook();
        var worksheet = workbook.Worksheets.Add(ReceiptWorksheetSchema.WorksheetName);
        WriteWorksheet(worksheet, rows, wrongHeader);
        if (includeOtherSheet)
        {
            var other = workbook.Worksheets.Add("Other");
            WriteWorksheet(other, rows, false);
        }

        workbook.SaveAs(path);
    }

    private static void WriteWorksheet(IXLWorksheet worksheet, IReadOnlyList<RowSpec> rows, bool wrongHeader)
    {
        var headers = ReceiptWorksheetSchema.ExpectedHeaders.ToArray();
        if (wrongHeader)
        {
            headers[5] = "金額";
        }

        for (var index = 0; index < headers.Length; index++)
        {
            worksheet.Cell(2, index + 1).Value = headers[index];
        }

        foreach (var row in rows)
        {
            worksheet.Cell(row.RowNumber, 1).Value = row.ROCYear;
            worksheet.Cell(row.RowNumber, 2).Value = row.Month;
            worksheet.Cell(row.RowNumber, 3).Value = row.Day;
            worksheet.Cell(row.RowNumber, 4).Value = row.Serial;
            worksheet.Cell(row.RowNumber, 4).Style.NumberFormat.Format = "@";
            worksheet.Cell(row.RowNumber, 5).Value = row.Payer;
            worksheet.Cell(row.RowNumber, 6).Value = row.Amount.ToString() ?? string.Empty;
            worksheet.Cell(row.RowNumber, 7).Value = row.Reason;
            if (!string.IsNullOrEmpty(row.Handler))
            {
                worksheet.Cell(row.RowNumber, 8).Value = row.Handler;
            }
        }
    }

    private static RowSpec[] ValidRows(params int[] rowNumbers) =>
        rowNumbers
            .Select(row => new RowSpec(row, (row + 48).ToString("D4"), "A實業股份有限公司", 1000, "運動會禮金"))
            .ToArray();

    private static RowSpec[] CombineRows(params IEnumerable<RowSpec>[] groups) =>
        groups.SelectMany(group => group).ToArray();

    private static RowSpec InvalidAmountRow(int row) =>
        new(row, $"00{row - 2:00}", "A實業股份有限公司", "bad", "運動會禮金");

    private static RowSpec OverLimitRow(int row) =>
        new(row, $"00{row - 2:00}", "A實業股份有限公司", 10_000_000, "運動會禮金");

    private static string TemplatePath() =>
        Path.Combine(RepositoryRoot(), ReceiptTemplateMapping.RelativeTemplatePath);

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

    private sealed record RowSpec(
        int RowNumber,
        string Serial,
        string Payer,
        object Amount,
        string Reason,
        int ROCYear = 112,
        int Month = 12,
        int Day = 30,
        string Handler = "");

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
                $"receiptxcel-phase2-{Guid.NewGuid():N}");
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
