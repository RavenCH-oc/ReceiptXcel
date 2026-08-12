using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Runtime.CompilerServices;
using XlsxDocxGenerator.Services.Diagnostics;
using XlsxDocxGenerator.Services.Errors;
using XlsxDocxGenerator.Services.Excel;
using XlsxDocxGenerator.Services.Receipts;
using XlsxDocxGenerator.Services.Settings;

namespace XlsxDocxGenerator.ViewModels;

/// <summary>
/// ReceiptXcel-only presentation state. The inherited MainWindowViewModel is
/// intentionally left available for regression coverage, but the product UI
/// binds exclusively to this specialized ViewModel.
/// </summary>
public sealed class ReceiptMainWindowViewModel : INotifyPropertyChanged
{
    private const string FixedWorksheetName = ReceiptWorksheetSchema.WorksheetName;

    private readonly IExcelReader _excelReader;
    private readonly ReceiptBatchGenerationService _batchService;
    private readonly SettingsService _settingsService;
    private readonly DiagnosticsLogger _diagnosticsLogger;
    private CancellationTokenSource? _previewCancellationSource;
    private CancellationTokenSource? _batchCancellationSource;
    private ReceiptBatchPreview? _preview;
    private int _previewRequestVersion;
    private bool _isPreviewRefreshing;
    private bool? _excelSchemaValid;
    private bool _isGenerating;
    private bool _isUpdatingAutomaticOutput;
    private bool _isOutputDirectoryUserDefined;
    private string _excelPath = string.Empty;
    private string _outputDirectory = string.Empty;
    private string? _selectedWorksheet;
    private ReceiptSelectionModeOption _selectedMode;
    private string _latestCountText = "1";
    private string _rowExpressionText = string.Empty;
    private string _excelFormatStatusText = "尚未載入 Excel";
    private string _selectionSummaryText = "選擇 Excel 後顯示資料筆數。";
    private string _statusText = "請選擇 Excel 收據登記表。";
    private string _progressText = string.Empty;
    private double _progressValue;
    private double _progressMaximum = 1;
    private bool _internalTemplateAvailable;
    private string _internalTemplateStatusText;
    private double? _windowWidth;
    private double? _windowHeight;

    public ReceiptMainWindowViewModel(
        SettingsService? settingsService = null,
        DiagnosticsLogger? diagnosticsLogger = null,
        ReceiptBatchGenerationService? batchService = null,
        IExcelReader? excelReader = null)
    {
        _settingsService = settingsService ?? new SettingsService();
        _diagnosticsLogger = diagnosticsLogger ?? new DiagnosticsLogger();
        _batchService = batchService ?? new ReceiptBatchGenerationService();
        _excelReader = excelReader ?? new ExcelReader();

        SelectionModes =
        [
            new(ReceiptSelectionMode.Latest, "最新資料"),
            new(ReceiptSelectionMode.ExcelRows, "Excel 列號")
        ];
        _selectedMode = SelectionModes[0];

        var settings = _settingsService.Load();
        _outputDirectory = settings.LastOutputDirectory ?? string.Empty;
        _isOutputDirectoryUserDefined = !string.IsNullOrWhiteSpace(settings.LastOutputDirectory);
        _windowWidth = settings.WindowWidth;
        _windowHeight = settings.WindowHeight;

        if (Enum.TryParse<ReceiptSelectionMode>(settings.LastSelectionMode, true, out var savedMode))
        {
            _selectedMode = SelectionModes.FirstOrDefault(item => item.Mode == savedMode)
                ?? SelectionModes[0];
        }

        _internalTemplateAvailable = _batchService.IsInternalTemplateAvailable(
            out var internalTemplateStatus);
        _internalTemplateStatusText = _internalTemplateAvailable
            ? "✓ 固定收據格式已就緒"
            : internalTemplateStatus;
    }

    public event PropertyChangedEventHandler? PropertyChanged;

    public ObservableCollection<ReceiptSelectionModeOption> SelectionModes { get; }

    public ObservableCollection<string> WorksheetNames { get; } = [];

    public ObservableCollection<ReceiptResultItem> BatchResults { get; } = [];

    public string ProductTitle => "ReceiptXcel｜自行收納款項收據產生工具";

    public string ExcelPath
    {
        get => _excelPath;
        private set => SetField(ref _excelPath, value);
    }

    public string ExcelFileName => string.IsNullOrWhiteSpace(ExcelPath)
        ? string.Empty
        : Path.GetFileName(ExcelPath);

    public string? SelectedWorksheet
    {
        get => _selectedWorksheet;
        set
        {
            if (!SetField(ref _selectedWorksheet, value))
            {
                return;
            }

            RequestPreviewRefresh();
        }
    }

    public ReceiptSelectionModeOption SelectedMode
    {
        get => _selectedMode;
        set
        {
            ArgumentNullException.ThrowIfNull(value);
            if (!SetField(ref _selectedMode, value))
            {
                return;
            }

            OnPropertyChanged(nameof(IsLatestMode));
            OnPropertyChanged(nameof(IsExcelRowsMode));
            SaveSettings();
            RequestPreviewRefresh();
        }
    }

    public bool IsLatestMode => SelectedMode.Mode == ReceiptSelectionMode.Latest;

    public bool IsExcelRowsMode => SelectedMode.Mode == ReceiptSelectionMode.ExcelRows;

    public string LatestCountText
    {
        get => _latestCountText;
        set
        {
            if (SetField(ref _latestCountText, value))
            {
                RequestPreviewRefresh();
            }
        }
    }

    public string RowExpressionText
    {
        get => _rowExpressionText;
        set
        {
            if (SetField(ref _rowExpressionText, value))
            {
                RequestPreviewRefresh();
            }
        }
    }

    public string OutputDirectory
    {
        get => _outputDirectory;
        set
        {
            if (!SetField(ref _outputDirectory, value))
            {
                return;
            }

            if (!_isUpdatingAutomaticOutput)
            {
                _isOutputDirectoryUserDefined = !string.IsNullOrWhiteSpace(value);
                SaveSettings();
            }

            OnPropertyChanged(nameof(CanGenerate));
        }
    }

    public bool IsGenerating
    {
        get => _isGenerating;
        private set
        {
            if (!SetField(ref _isGenerating, value))
            {
                return;
            }

            OnPropertyChanged(nameof(CanGenerate));
            OnPropertyChanged(nameof(CanCancel));
            OnPropertyChanged(nameof(CanEditInputs));
        }
    }

    public bool CanEditInputs => !IsGenerating;

    public bool CanCancel => IsGenerating && _batchCancellationSource is not null;

    public bool CanGenerate =>
        !IsGenerating
        && _internalTemplateAvailable
        && _preview?.CanGenerate == true
        && !string.IsNullOrWhiteSpace(OutputDirectory);

    public bool IsPreviewRefreshing => _isPreviewRefreshing;

    public bool? ExcelSchemaValid => _excelSchemaValid;

    public string ExcelFormatStatusText
    {
        get => _excelFormatStatusText;
        private set => SetField(ref _excelFormatStatusText, value);
    }

    public string SelectionSummaryText
    {
        get => _selectionSummaryText;
        private set => SetField(ref _selectionSummaryText, value);
    }

    public string StatusText
    {
        get => _statusText;
        set => SetField(ref _statusText, value);
    }

    public string ProgressText
    {
        get => _progressText;
        private set => SetField(ref _progressText, value);
    }

    public double ProgressValue
    {
        get => _progressValue;
        private set => SetField(ref _progressValue, value);
    }

    public double ProgressMaximum
    {
        get => _progressMaximum;
        private set => SetField(ref _progressMaximum, value);
    }

    public string InternalTemplateStatusText
    {
        get => _internalTemplateStatusText;
        private set => SetField(ref _internalTemplateStatusText, value);
    }

    public double? SavedWindowWidth => _windowWidth;

    public double? SavedWindowHeight => _windowHeight;

    public IReadOnlyList<int> PreviewRowNumbers =>
        _preview?.SelectedRows.Select(row => row.ExcelRowNumber).ToArray() ?? [];

    public async Task LoadExcelAsync(
        string path,
        CancellationToken cancellationToken = default)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(path);
        CancelPendingPreviewRefresh();
        ClearPreview();
        ExcelPath = Path.GetFullPath(path);
        OnPropertyChanged(nameof(ExcelFileName));
        ApplyAutomaticOutputDirectory(ExcelPath);
        WorksheetNames.Clear();
        SelectedWorksheet = null;

        try
        {
            var worksheetNames = await _excelReader.GetWorksheetNamesAsync(
                ExcelPath,
                cancellationToken);
            foreach (var worksheetName in worksheetNames)
            {
                WorksheetNames.Add(worksheetName);
            }

            SelectedWorksheet = worksheetNames.FirstOrDefault(name =>
                string.Equals(name, FixedWorksheetName, StringComparison.Ordinal))
                ?? worksheetNames.FirstOrDefault();

            if (SelectedWorksheet is null)
            {
                ExcelFormatStatusText = "Excel 格式不符：找不到可用工作表。";
                StatusText = ExcelFormatStatusText;
                return;
            }

            StatusText = $"已載入 Excel：{Path.GetFileName(ExcelPath)}";
            RequestPreviewRefresh();
        }
        catch (OperationCanceledException)
        {
            throw;
        }
        catch (TemplateDefinitionException exception)
        {
            ExcelFormatStatusText = "Excel 格式不符";
            StatusText = exception.UserMessage;
            _diagnosticsLogger.LogException("LoadExcel", exception.Code.ToString(), exception);
        }
        catch (Exception exception)
        {
            ExcelFormatStatusText = "Excel 格式不符";
            StatusText = "無法讀取 Excel 檔案，請確認檔案未損壞或被其他程式鎖定。";
            _diagnosticsLogger.LogException("LoadExcel", "ExcelReadFailed", exception);
        }
    }

    public async Task RefreshSelectionPreviewAsync(
        CancellationToken cancellationToken = default)
    {
        CancelPendingPreviewRefresh();
        var requestVersion = _previewRequestVersion;
        ClearPreview();

        if (string.IsNullOrWhiteSpace(ExcelPath)
            || string.IsNullOrWhiteSpace(SelectedWorksheet))
        {
            return;
        }

        var source = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
        _previewCancellationSource = source;
        _isPreviewRefreshing = true;
        OnPropertyChanged(nameof(IsPreviewRefreshing));
        ExcelFormatStatusText = "正在檢查 Excel 格式…";
        StatusText = "正在檢查 Excel 格式並更新資料預覽。";

        try
        {
            await Task.Delay(180, source.Token);
            var request = CreateSelectionRequest();
            var preview = await Task.Run(
                () => _batchService.PreviewAsync(
                    ExcelPath,
                    SelectedWorksheet!,
                    request,
                    source.Token),
                source.Token);
            source.Token.ThrowIfCancellationRequested();
            if (requestVersion != _previewRequestVersion)
            {
                return;
            }

            _preview = preview;
            _excelSchemaValid = preview.SchemaValid;
            ApplyPreviewStatus(preview);
        }
        catch (OperationCanceledException)
        {
            // A newer input change superseded this preview.
        }
        catch (Exception exception)
        {
            _excelSchemaValid = false;
            ExcelFormatStatusText = "Excel 格式不符";
            SelectionSummaryText = "無法計算資料選擇預覽。";
            StatusText = "無法計算資料選擇預覽，請確認 Excel 與選擇條件。";
            _diagnosticsLogger.LogException(
                "RefreshReceiptPreview",
                "ReceiptPreviewFailed",
                exception);
        }
        finally
        {
            if (ReferenceEquals(_previewCancellationSource, source))
            {
                _previewCancellationSource = null;
                source.Dispose();
            }

            if (requestVersion == _previewRequestVersion)
            {
                _isPreviewRefreshing = false;
                OnPropertyChanged(nameof(IsPreviewRefreshing));
            }

            OnPropertyChanged(nameof(CanGenerate));
        }
    }

    public async Task GenerateBatchAsync(
        CancellationToken cancellationToken = default)
    {
        if (!CanGenerate || SelectedWorksheet is null)
        {
            StatusText = "請先選擇有效的 Excel 資料。";
            return;
        }

        var request = CreateSelectionRequest();
        var source = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
        _batchCancellationSource = source;
        IsGenerating = true;
        BatchResults.Clear();
        ProgressValue = 0;
        ProgressMaximum = Math.Max(1, _preview?.TotalSelected ?? 1);
        ProgressText = $"正在產生 0 / {ProgressMaximum:0}";

        var progress = new Progress<ReceiptBatchProgress>(value =>
        {
            ProgressValue = value.Completed;
            ProgressMaximum = Math.Max(1, value.Total);
            ProgressText = $"正在產生 {value.Completed} / {value.Total}";
        });

        try
        {
            var result = await _batchService.GenerateBatchAsync(
                ExcelPath,
                SelectedWorksheet,
                request,
                OutputDirectory,
                progress,
                source.Token);

            foreach (var rowResult in result.Results)
            {
                BatchResults.Add(ReceiptResultItem.From(rowResult));
            }

            ProgressValue = result.Results.Count;
            ProgressMaximum = Math.Max(1, result.TotalSelected);
            if (result.IsCancelled)
            {
                ProgressText = $"已取消：{result.Results.Count} / {result.TotalSelected}";
                StatusText = $"已取消\n已完成：{result.SuccessCount}\n失敗：{result.FailureCount}\n未處理：{result.UnprocessedCount}";
            }
            else
            {
                ProgressText = $"完成 {result.Results.Count} / {result.TotalSelected}";
                StatusText = $"完成\n成功：{result.SuccessCount}\n失敗：{result.FailureCount}";
                if (!string.IsNullOrWhiteSpace(result.Notice))
                {
                    StatusText += $"\n{result.Notice}";
                }
            }
        }
        catch (ReceiptBatchGenerationException exception)
        {
            if (exception.Code == ReceiptBatchFatalErrorCode.InternalTemplateInvalid)
            {
                _internalTemplateAvailable = false;
                InternalTemplateStatusText = "內建收據模板無法使用，請重新安裝 ReceiptXcel。";
            }

            StatusText = exception.UserMessage;
            ProgressText = string.Empty;
            _diagnosticsLogger.LogException(
                "GenerateReceiptBatch",
                exception.Code.ToString(),
                exception);
        }
        catch (OperationCanceledException)
        {
            StatusText = "已取消\n已完成：0\n失敗：0\n未處理：尚未開始";
            ProgressText = string.Empty;
        }
        catch (Exception exception)
        {
            StatusText = "批次產生失敗，請查看 diagnostics log。";
            ProgressText = string.Empty;
            _diagnosticsLogger.LogException(
                "GenerateReceiptBatch",
                "ReceiptBatchFailed",
                exception);
        }
        finally
        {
            if (ReferenceEquals(_batchCancellationSource, source))
            {
                _batchCancellationSource = null;
            }

            source.Dispose();
            IsGenerating = false;
            SaveSettings();
        }
    }

    public void CancelBatch()
    {
        if (!CanCancel)
        {
            return;
        }

        StatusText = "正在取消，請稍候完成目前文件…";
        _batchCancellationSource!.Cancel();
    }

    public void SaveWindowSize(double width, double height)
    {
        if (width > 0)
        {
            _windowWidth = width;
        }

        if (height > 0)
        {
            _windowHeight = height;
        }

        SaveSettings();
    }

    private ReceiptSelectionRequest CreateSelectionRequest() =>
        SelectedMode.Mode switch
        {
            ReceiptSelectionMode.Latest => new(ReceiptSelectionMode.Latest, LatestCountText),
            ReceiptSelectionMode.ExcelRows => new(ReceiptSelectionMode.ExcelRows, RowExpressionText),
            _ => throw new InvalidOperationException("不支援的收據選擇方式。")
        };

    private void ApplyPreviewStatus(ReceiptBatchPreview preview)
    {
        if (!preview.SchemaValid)
        {
            ExcelFormatStatusText = "Excel 格式不符";
            SelectionSummaryText = preview.BlockingUserMessage ?? "Excel 格式不符。";
            StatusText = SelectionSummaryText;
            return;
        }

        ExcelFormatStatusText = "✓ Excel 格式正確";
        SelectionSummaryText =
            $"找到 {preview.AvailableRowCount} 筆收據資料\n"
            + $"符合條件：{preview.TotalSelected} 筆\n"
            + $"預計產生：{preview.TotalSelected} 份收據";

        if (!string.IsNullOrWhiteSpace(preview.Notice))
        {
            SelectionSummaryText += $"\n{preview.Notice}";
        }

        if (!string.IsNullOrWhiteSpace(preview.BlockingUserMessage))
        {
            SelectionSummaryText += $"\n{preview.BlockingUserMessage}";
            StatusText = preview.BlockingUserMessage;
        }
        else
        {
            StatusText = preview.TotalSelected == 0
                ? "沒有符合條件的資料。"
                : "預覽已更新，可以開始產生收據。";
        }
    }

    private void ApplyAutomaticOutputDirectory(string excelPath)
    {
        if (_isOutputDirectoryUserDefined)
        {
            return;
        }

        try
        {
            var excelDirectory = Path.GetDirectoryName(Path.GetFullPath(excelPath));
            if (string.IsNullOrWhiteSpace(excelDirectory))
            {
                return;
            }

            _isUpdatingAutomaticOutput = true;
            OutputDirectory = Path.Combine(excelDirectory, "ReceiptXcel輸出");
        }
        catch (Exception exception) when (exception is ArgumentException
            or IOException
            or NotSupportedException)
        {
            _diagnosticsLogger.LogException(
                "ApplyReceiptOutputDefault",
                "OutputDefaultPathFailed",
                exception);
        }
        finally
        {
            _isUpdatingAutomaticOutput = false;
        }
    }

    private void RequestPreviewRefresh()
    {
        CancelPendingPreviewRefresh();
        ClearPreview();
        if (string.IsNullOrWhiteSpace(ExcelPath)
            || string.IsNullOrWhiteSpace(SelectedWorksheet)
            || IsGenerating)
        {
            return;
        }

        _ = RefreshSelectionPreviewAsync();
    }

    private void CancelPendingPreviewRefresh()
    {
        _previewRequestVersion++;
        if (_previewCancellationSource is null)
        {
            return;
        }

        _previewCancellationSource.Cancel();
        _previewCancellationSource.Dispose();
        _previewCancellationSource = null;
    }

    private void ClearPreview()
    {
        _preview = null;
        _excelSchemaValid = null;
        _isPreviewRefreshing = false;
        OnPropertyChanged(nameof(ExcelSchemaValid));
        OnPropertyChanged(nameof(IsPreviewRefreshing));
        OnPropertyChanged(nameof(CanGenerate));
    }

    private void SaveSettings()
    {
        _settingsService.Save(new Models.ApplicationSettings
        {
            LastOutputDirectory = _isOutputDirectoryUserDefined
                && !string.IsNullOrWhiteSpace(OutputDirectory)
                ? OutputDirectory
                : null,
            LastSelectionMode = SelectedMode.Mode.ToString(),
            WindowWidth = _windowWidth,
            WindowHeight = _windowHeight
        });
    }

    private bool SetField<T>(
        ref T field,
        T value,
        [CallerMemberName] string? propertyName = null)
    {
        if (EqualityComparer<T>.Default.Equals(field, value))
        {
            return false;
        }

        field = value;
        OnPropertyChanged(propertyName);
        return true;
    }

    private void OnPropertyChanged([CallerMemberName] string? propertyName = null) =>
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
}

public sealed record ReceiptSelectionModeOption(
    ReceiptSelectionMode Mode,
    string DisplayName);

public sealed record ReceiptResultItem(
    string StatusIcon,
    string RowText,
    string ReceiptText,
    string Detail)
{
    public static ReceiptResultItem From(ReceiptRowGenerationResult result) =>
        result.Success
            ? new(
                "✓",
                $"Row {result.ExcelRowNumber}",
                result.ReceiptNumber ?? string.Empty,
                result.OutputPath is null ? string.Empty : Path.GetFileName(result.OutputPath))
            : new(
                "✕",
                $"Row {result.ExcelRowNumber}",
                result.ReceiptNumber ?? string.Empty,
                result.UserMessage ?? "產生失敗。" );
}
