using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Runtime.CompilerServices;
using XlsxDocxGenerator.Models;
using XlsxDocxGenerator.Services.Excel;
using XlsxDocxGenerator.Services.Errors;
using XlsxDocxGenerator.Services.Generation;
using XlsxDocxGenerator.Services.Diagnostics;
using XlsxDocxGenerator.Services.Settings;
using XlsxDocxGenerator.Services.Templates;
using XlsxDocxGenerator.Services.Word;

namespace XlsxDocxGenerator.ViewModels;

public sealed class MainWindowViewModel : INotifyPropertyChanged
{
    private readonly ExcelReader _excelReader = new();
    private readonly TemplateService _templateService;
    private readonly GenerationService _generationService;
    private readonly SelectionRuleParser _selectionRuleParser = new();
    private readonly SettingsService _settingsService;
    private readonly DiagnosticsLogger _diagnosticsLogger;
    private readonly FilenamePolicy _filenamePolicy = new();
    private string _excelPath = string.Empty;
    private string _wordTemplatePath = string.Empty;
    private string _outputDirectory = string.Empty;
    private string _headerRowText = "1";
    private string _markerRowText = string.Empty;
    private string _runtimeMarkerRowText = string.Empty;
    private string _markerTokensText = string.Empty;
    private string _rowExpressionText = string.Empty;
    private string _latestCountText = string.Empty;
    private string _outputPattern = string.Empty;
    private string? _selectedWorksheet;
    private string? _selectedMarkerColumn;
    private SelectionModeOption _selectedMode;
    private RowSelectionResult? _selectionPreview;
    private SelectionRule? _selectionRule;
    private bool _isGenerating;
    private string _statusText = "尚未載入 Excel\n尚未載入 Word 模板";
    private string _selectionSummaryText = "尚未建立資料選擇預覽。";
    private string _progressText = string.Empty;
    private double _progressValue;
    private double _progressMaximum = 1;
    private string _templateName = string.Empty;
    private string? _lastTemplateConfigPath;
    private CancellationTokenSource? _batchCancellationSource;
    private CancellationTokenSource? _previewRefreshCancellationSource;
    private int _previewRequestVersion;
    private bool? _excelSchemaValid;
    private bool _isPreviewRefreshing;
    private double? _windowWidth;
    private double? _windowHeight;
    private bool _isOutputDirectoryUserDefined;
    private bool _isOutputDirectoryAutomaticDefault;
    private bool _isOutputPatternUserDefined;
    private bool _isOutputPatternAutomaticDefault;
    private bool _isOutputPatternFromTemplate;
    private MainWindowMode _activeMode = MainWindowMode.GenerateDocuments;

    public MainWindowViewModel(
        SettingsService? settingsService = null,
        DiagnosticsLogger? diagnosticsLogger = null)
    {
        _settingsService = settingsService ?? new SettingsService();
        _diagnosticsLogger = diagnosticsLogger ?? new DiagnosticsLogger();
        _templateService = new TemplateService(_excelReader, new MarkerParser());
        _generationService = new GenerationService(_excelReader, new DocxGenerator());
        SelectionModes =
        [
            new(SelectionMode.Marker, "標記代號"),
            new(SelectionMode.ExcelRows, "Excel 列號"),
            new(SelectionMode.Latest, "最新資料")
        ];
        _selectedMode = SelectionModes[0];

        var settings = _settingsService.Load();
        _outputDirectory = settings.LastOutputDirectory ?? string.Empty;
        _outputPattern = settings.LastOutputFilenamePattern ?? string.Empty;
        _isOutputDirectoryUserDefined = !string.IsNullOrWhiteSpace(settings.LastOutputDirectory);
        _isOutputPatternUserDefined = !string.IsNullOrWhiteSpace(settings.LastOutputFilenamePattern);
        _windowWidth = settings.WindowWidth;
        _windowHeight = settings.WindowHeight;
        if (Enum.TryParse<SelectionMode>(settings.LastSelectionMode, true, out var savedMode))
        {
            _selectedMode = SelectionModes.FirstOrDefault(option => option.Mode == savedMode)
                ?? SelectionModes[0];
        }
        if (!string.IsNullOrWhiteSpace(settings.LastTemplateConfigPath))
        {
            _lastTemplateConfigPath = settings.LastTemplateConfigPath;
            RestoreLastTemplateAtStartup();
        }
    }

    public event PropertyChangedEventHandler? PropertyChanged;

    public MainWindowMode ActiveMode
    {
        get => _activeMode;
        set
        {
            if (SetField(ref _activeMode, value))
            {
                OnPropertyChanged(nameof(IsGenerateDocumentsMode));
                OnPropertyChanged(nameof(IsCreateTemplateMode));
                if (value == MainWindowMode.GenerateDocuments)
                {
                    RequestSelectionPreviewRefresh();
                }
                else
                {
                    InvalidateSelectionPreview();
                }
            }
        }
    }

    public bool IsGenerateDocumentsMode => ActiveMode == MainWindowMode.GenerateDocuments;

    public bool IsCreateTemplateMode => ActiveMode == MainWindowMode.CreateTemplate;

    public ObservableCollection<SelectionModeOption> SelectionModes { get; }

    public ObservableCollection<string> WorksheetNames { get; } = [];

    public ObservableCollection<string> SelectionMarkerColumns { get; } = [];

    public ObservableCollection<MappingPreviewItem> MappingPreview { get; } = [];

    public ObservableCollection<BatchResultItem> BatchResults { get; } = [];

    public TemplateDefinition? CurrentTemplateDefinition { get; private set; }

    public string TemplateName
    {
        get => _templateName;
        set
        {
            if (!SetField(ref _templateName, value))
            {
                return;
            }

            if (CurrentTemplateDefinition is not null && !string.IsNullOrWhiteSpace(value))
            {
                CurrentTemplateDefinition = CurrentTemplateDefinition with
                {
                    TemplateName = value.Trim()
                };
                OnPropertyChanged(nameof(CurrentTemplateDefinition));
            }
        }
    }

    public string TemplateSummaryText => CurrentTemplateDefinition is null
        ? "尚未建立或載入模板"
        : $"模板：{CurrentTemplateDefinition.TemplateName}｜Word 欄位：{CurrentTemplateDefinition.FieldMappings.Count}";

    public string TemplateDisplayName => CurrentTemplateDefinition?.TemplateName ?? "尚未載入模板";

    public bool HasTemplate => CurrentTemplateDefinition is not null;

    public string TemplateFieldCountText => CurrentTemplateDefinition is null
        ? "尚未載入模板"
        : $"{CurrentTemplateDefinition.FieldMappings.Count} 個";

    public string ExcelFileName => string.IsNullOrWhiteSpace(ExcelPath)
        ? "尚未選擇 Excel"
        : Path.GetFileName(ExcelPath);

    public string WordTemplateFileName => string.IsNullOrWhiteSpace(WordTemplatePath)
        ? "尚未選擇 Word"
        : Path.GetFileName(WordTemplatePath);

    public string ExcelFormatStatusText => string.IsNullOrWhiteSpace(ExcelPath)
        ? "請選擇 Excel 資料檔"
        : CurrentTemplateDefinition is null
            ? "已載入 Excel，請先載入模板"
            : _isPreviewRefreshing
                ? "正在檢查 Excel 格式…"
                : _excelSchemaValid == true
                    ? "Excel 格式符合目前模板"
                    : _excelSchemaValid == false
                        ? "Excel 欄位與目前模板不符"
                        : "已載入 Excel，正在準備資料預覽。";

    public string ExcelPath
    {
        get => _excelPath;
        set
        {
            if (SetField(ref _excelPath, value))
            {
                OnPropertyChanged(nameof(ExcelFileName));
                OnPropertyChanged(nameof(ExcelFormatStatusText));
                RequestSelectionPreviewRefresh();
                OnPropertyChanged(nameof(CanGenerateBatch));
            }
        }
    }

    public string WordTemplatePath
    {
        get => _wordTemplatePath;
        set
        {
            if (!SetField(ref _wordTemplatePath, value))
            {
                return;
            }

            OnPropertyChanged(nameof(WordTemplateFileName));

            if (CurrentTemplateDefinition is not null)
            {
                CurrentTemplateDefinition = CurrentTemplateDefinition with
                {
                    WordTemplatePath = value
                };
                OnPropertyChanged(nameof(CurrentTemplateDefinition));
            }

            if (!string.IsNullOrWhiteSpace(value)
                && !_isOutputPatternUserDefined
                && !_isOutputPatternFromTemplate
                && (_isOutputPatternAutomaticDefault
                    || string.IsNullOrWhiteSpace(CurrentTemplateDefinition?.OutputFileNamePattern)))
            {
                ApplyAutomaticOutputPattern(value);
            }

            RequestSelectionPreviewRefresh();
            OnPropertyChanged(nameof(CanGenerateBatch));
        }
    }

    public string? SelectedWorksheet
    {
        get => _selectedWorksheet;
        set
        {
            if (!SetField(ref _selectedWorksheet, value) || CurrentTemplateDefinition is null)
            {
                return;
            }

            if (!string.IsNullOrWhiteSpace(value))
            {
                CurrentTemplateDefinition = CurrentTemplateDefinition with
                {
                    PreferredWorksheetName = value
                };
                OnPropertyChanged(nameof(CurrentTemplateDefinition));
            }

            RequestSelectionPreviewRefresh();
        }
    }

    public string HeaderRowText
    {
        get => _headerRowText;
        set => SetField(ref _headerRowText, value);
    }

    public string MarkerRowText
    {
        get => _markerRowText;
        set => SetField(ref _markerRowText, value);
    }

    public string RuntimeMarkerRowText
    {
        get => _runtimeMarkerRowText;
        set
        {
            if (SetField(ref _runtimeMarkerRowText, value))
            {
                RequestSelectionPreviewRefresh();
            }
        }
    }

    public bool HasSingleMappingTemplate =>
        CurrentTemplateDefinition?.FieldMappings.Count == 1;

    public string RuntimeMarkerInstructionText => HasSingleMappingTemplate
        ? "此模板只有一個對應欄位，無法安全自動辨識欄位對應列。若 Excel 仍保留該列，請在「進階選項」中指定列號。"
        : "欄位對應列留白時，程式會依目前模板自動判斷。";

    public SelectionModeOption SelectedMode
    {
        get => _selectedMode;
        set
        {
            if (SetField(ref _selectedMode, value))
            {
                OnPropertyChanged(nameof(IsMarkerMode));
                OnPropertyChanged(nameof(IsExcelRowsMode));
                OnPropertyChanged(nameof(IsLatestMode));
                SaveSettings();
                RequestSelectionPreviewRefresh();
            }
        }
    }

    public string SelectedMarkerColumn
    {
        get => _selectedMarkerColumn ?? string.Empty;
        set
        {
            if (SetField(ref _selectedMarkerColumn, value))
            {
                RequestSelectionPreviewRefresh();
            }
        }
    }

    public string MarkerTokensText
    {
        get => _markerTokensText;
        set
        {
            if (SetField(ref _markerTokensText, value))
            {
                RequestSelectionPreviewRefresh();
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
                RequestSelectionPreviewRefresh();
            }
        }
    }

    public string LatestCountText
    {
        get => _latestCountText;
        set
        {
            if (SetField(ref _latestCountText, value))
            {
                RequestSelectionPreviewRefresh();
            }
        }
    }

    public string OutputDirectory
    {
        get => _outputDirectory;
        set
        {
            var wasUserDefined = _isOutputDirectoryUserDefined;
            _isOutputDirectoryUserDefined = !string.IsNullOrWhiteSpace(value);
            _isOutputDirectoryAutomaticDefault = false;

            if (SetField(ref _outputDirectory, value)
                || wasUserDefined != _isOutputDirectoryUserDefined)
            {
                SaveSettings();
                OnPropertyChanged(nameof(CanGenerateBatch));
            }
        }
    }

    public string OutputPattern
    {
        get => _outputPattern;
        set
        {
            var wasUserDefined = _isOutputPatternUserDefined;
            _isOutputPatternUserDefined = !string.IsNullOrWhiteSpace(value);
            _isOutputPatternAutomaticDefault = false;
            _isOutputPatternFromTemplate = false;

            if (!SetField(ref _outputPattern, value)
                && wasUserDefined == _isOutputPatternUserDefined)
            {
                return;
            }

            if (CurrentTemplateDefinition is not null)
            {
                CurrentTemplateDefinition = CurrentTemplateDefinition with
                {
                    OutputFileNamePattern = string.IsNullOrWhiteSpace(value) ? null : value.Trim()
                };
                OnPropertyChanged(nameof(CurrentTemplateDefinition));
            }
            SaveSettings();
            OnPropertyChanged(nameof(CanGenerateBatch));
        }
    }

    public string StatusText
    {
        get => _statusText;
        set => SetField(ref _statusText, value);
    }

    public string SelectionSummaryText
    {
        get => _selectionSummaryText;
        private set => SetField(ref _selectionSummaryText, value);
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

    public bool IsGenerating
    {
        get => _isGenerating;
        private set
        {
            if (SetField(ref _isGenerating, value))
            {
                OnPropertyChanged(nameof(CanGenerateBatch));
                OnPropertyChanged(nameof(CanCancelBatch));
                OnPropertyChanged(nameof(IsNotGenerating));
            }
        }
    }

    public bool IsNotGenerating => !IsGenerating;

    public bool CanCancelBatch => IsGenerating && _batchCancellationSource is not null;

    public double? SavedWindowWidth => _windowWidth;

    public double? SavedWindowHeight => _windowHeight;

    public bool IsMarkerMode => SelectedMode.Mode == SelectionMode.Marker;

    public bool IsExcelRowsMode => SelectedMode.Mode == SelectionMode.ExcelRows;

    public bool IsLatestMode => SelectedMode.Mode == SelectionMode.Latest;

    public bool CanGenerateBatch =>
        !IsGenerating
        && CurrentTemplateDefinition is not null
        && _selectionRule is not null
        && _selectionPreview is not null
        && _excelSchemaValid == true
        && _selectionPreview.TotalSelected > 0
        && File.Exists(ExcelPath)
        && File.Exists(WordTemplatePath)
        && !string.IsNullOrWhiteSpace(OutputDirectory)
        && IsOutputPatternValid();

    public async Task LoadExcelAsync(string path, CancellationToken cancellationToken = default)
    {
        var loadedTemplate = CurrentTemplateDefinition;
        RuntimeMarkerRowText = string.Empty;
        ExcelPath = path;
        ApplyAutomaticOutputDirectory(path);
        WorksheetNames.Clear();
        SelectionMarkerColumns.Clear();
        MappingPreview.Clear();
        BatchResults.Clear();
            CurrentTemplateDefinition = loadedTemplate;
            ClearSelectionPreview();
            OnPropertyChanged(nameof(CurrentTemplateDefinition));
            OnPropertyChanged(nameof(TemplateDisplayName));
            OnPropertyChanged(nameof(TemplateFieldCountText));
            OnPropertyChanged(nameof(ExcelFormatStatusText));

        try
        {
            var worksheetNames = await Task.Run(
                () => _excelReader.GetWorksheetNamesAsync(path, cancellationToken),
                cancellationToken);
            foreach (var worksheetName in worksheetNames)
            {
                WorksheetNames.Add(worksheetName);
            }

            if (CurrentTemplateDefinition is not null)
            {
                SelectedWorksheet = WorksheetNames.FirstOrDefault(name =>
                    string.Equals(name, CurrentTemplateDefinition.PreferredWorksheetName, StringComparison.OrdinalIgnoreCase))
                    ?? WorksheetNames.FirstOrDefault();
                await PopulateTemplatePreviewAsync(cancellationToken);
            }
            else
            {
                SelectedWorksheet = WorksheetNames.FirstOrDefault();
            }
            StatusText = worksheetNames.Count == 0
                ? "Excel 沒有可用的工作表。"
                : CurrentTemplateDefinition is null
                    ? IsCreateTemplateMode
                        ? $"已載入 Excel，共 {worksheetNames.Count} 個工作表，請讀取欄位對應。"
                        : $"已載入 Excel，共 {worksheetNames.Count} 個工作表，請先載入模板。"
                    : $"已載入 Excel，共 {worksheetNames.Count} 個工作表，正在檢查格式與資料預覽。";
            if (loadedTemplate is not null
                && !worksheetNames.Any(name => string.Equals(
                    name,
                    loadedTemplate.PreferredWorksheetName,
                    StringComparison.OrdinalIgnoreCase)))
            {
                StatusText = $"找不到模板指定的工作表「{loadedTemplate.PreferredWorksheetName}」，已選擇其他工作表，請確認欄位結構。";
            }

            RequestSelectionPreviewRefresh();
        }
        catch (TemplateDefinitionException exception)
        {
            WorksheetNames.Clear();
            SelectedWorksheet = null;
            _excelSchemaValid = false;
            OnPropertyChanged(nameof(ExcelFormatStatusText));
            OnPropertyChanged(nameof(CanGenerateBatch));
            StatusText = exception.UserMessage;
        }
        catch (Exception exception)
        {
            WorksheetNames.Clear();
            SelectedWorksheet = null;
            _excelSchemaValid = false;
            OnPropertyChanged(nameof(ExcelFormatStatusText));
            OnPropertyChanged(nameof(CanGenerateBatch));
            _diagnosticsLogger.LogException("LoadExcel", "ExcelLoadFailed", exception);
            StatusText = "讀取 Excel 失敗，請確認檔案格式與存取權限。";
        }
    }

    public async Task ReadMappingAsync(CancellationToken cancellationToken = default)
    {
        MappingPreview.Clear();
        BatchResults.Clear();
        CurrentTemplateDefinition = null;
        InvalidateSelectionPreview();
        OnPropertyChanged(nameof(CurrentTemplateDefinition));

        if (string.IsNullOrWhiteSpace(ExcelPath) || string.IsNullOrWhiteSpace(SelectedWorksheet))
        {
            StatusText = "請先選擇 Excel 檔案與工作表。";
            return;
        }

        if (!int.TryParse(HeaderRowText, out var headerRowNumber) || headerRowNumber <= 0)
        {
            StatusText = "Header row 必須是大於 0 的 Excel 列號。";
            return;
        }

        int? markerRowNumber = null;
        if (!string.IsNullOrWhiteSpace(MarkerRowText))
        {
            if (!int.TryParse(MarkerRowText, out var parsedMarkerRow) || parsedMarkerRow <= 0)
            {
                StatusText = "Mapping marker row 必須是大於 0 的 Excel 列號，或留白使用自動偵測。";
                return;
            }

            markerRowNumber = parsedMarkerRow;
        }

        try
        {
            var templateName = string.IsNullOrWhiteSpace(WordTemplatePath)
                ? Path.GetFileNameWithoutExtension(ExcelPath)
                : Path.GetFileNameWithoutExtension(WordTemplatePath);

            CurrentTemplateDefinition = await Task.Run(
                () => _templateService.CreateDefinitionAsync(
                    ExcelPath,
                    templateName,
                    WordTemplatePath,
                    SelectedWorksheet,
                    headerRowNumber,
                    markerRowNumber,
                    string.IsNullOrWhiteSpace(OutputPattern) ? null : OutputPattern.Trim(),
                cancellationToken),
                cancellationToken);
            TemplateName = CurrentTemplateDefinition.TemplateName;

            var worksheet = await Task.Run(
                () => _excelReader.ReadAsync(
                    ExcelPath,
                    SelectedWorksheet,
                    headerRowNumber,
                    cancellationToken),
                cancellationToken);

            foreach (var mapping in CurrentTemplateDefinition.FieldMappings)
            {
                MappingPreview.Add(new MappingPreviewItem(
                    $"{{{{{mapping.PlaceholderName}}}}}",
                    $"{mapping.ColumnLetter} - {mapping.HeaderName ?? "（無標題）"}"));
            }

            SelectionMarkerColumns.Clear();
            foreach (var header in worksheet.Headers.Where(header => !string.IsNullOrWhiteSpace(header)))
            {
                if (!SelectionMarkerColumns.Contains(header!, StringComparer.OrdinalIgnoreCase))
                {
                    SelectionMarkerColumns.Add(header!);
                }
            }

            SelectedMarkerColumn = SelectionMarkerColumns.FirstOrDefault() ?? string.Empty;
            OnPropertyChanged(nameof(CurrentTemplateDefinition));
            OnPropertyChanged(nameof(TemplateSummaryText));
            OnPropertyChanged(nameof(TemplateDisplayName));
            OnPropertyChanged(nameof(TemplateFieldCountText));
            OnPropertyChanged(nameof(ExcelFormatStatusText));
            StatusText = $"已解析 {MappingPreview.Count} 個欄位對應。請設定資料選擇方式並重新計算預覽。";
        }
        catch (TemplateDefinitionException exception)
        {
            StatusText = exception.UserMessage;
        }
        catch (Exception exception)
        {
            _diagnosticsLogger.LogException("ReadMapping", "MappingReadFailed", exception);
            StatusText = "建立模板定義失敗，請確認 Excel 內容。";
        }
    }

    private async Task PopulateTemplatePreviewAsync(CancellationToken cancellationToken)
    {
        if (CurrentTemplateDefinition is null || string.IsNullOrWhiteSpace(SelectedWorksheet))
        {
            return;
        }

        try
        {
            var worksheet = await Task.Run(
                () => _excelReader.ReadAsync(
                    ExcelPath,
                    SelectedWorksheet,
                    CurrentTemplateDefinition.HeaderRowNumber,
                    cancellationToken),
                cancellationToken);
            foreach (var mapping in CurrentTemplateDefinition.FieldMappings)
            {
                MappingPreview.Add(new MappingPreviewItem(
                    $"{{{{{mapping.PlaceholderName}}}}}",
                    $"{mapping.ColumnLetter} - {mapping.HeaderName ?? "（未命名欄位）"}"));
            }

            SelectionMarkerColumns.Clear();
            foreach (var header in worksheet.Headers.Where(header => !string.IsNullOrWhiteSpace(header)))
            {
                if (!SelectionMarkerColumns.Contains(header!, StringComparer.OrdinalIgnoreCase))
                {
                    SelectionMarkerColumns.Add(header!);
                }
            }

            SelectedMarkerColumn = SelectionMarkerColumns.FirstOrDefault() ?? string.Empty;
        }
        catch (TemplateDefinitionException exception)
        {
            StatusText = exception.UserMessage;
        }
    }

    public async Task CreateTemplateAsync(
        string savePath,
        CancellationToken cancellationToken = default)
    {
        if (CurrentTemplateDefinition is null)
        {
            await ReadMappingAsync(cancellationToken);
        }

        if (CurrentTemplateDefinition is null)
        {
            return;
        }

        if (string.IsNullOrWhiteSpace(WordTemplatePath) || !File.Exists(WordTemplatePath))
        {
            StatusText = "請先選擇有效的 Word 模板。";
            return;
        }

        try
        {
            var placeholders = await new DocxTemplateReader().ScanPlaceholdersAsync(
                WordTemplatePath,
                cancellationToken);
            new TemplateDefinitionValidator().ValidatePlaceholders(CurrentTemplateDefinition, placeholders);
            _templateService.SaveTemplate(savePath, CurrentTemplateDefinition);
            _lastTemplateConfigPath = savePath;
            SaveSettings();
            StatusText = $"模板建立完成\n欄位：{CurrentTemplateDefinition.FieldMappings.Count}\nWord placeholders：{placeholders.Count}";
        }
        catch (GenerationException exception)
        {
            StatusText = exception.UserMessage;
        }
        catch (TemplatePersistenceException exception)
        {
            StatusText = exception.UserMessage;
        }
        catch (Exception exception)
        {
            _diagnosticsLogger.LogException("CreateTemplate", "TemplateCreateFailed", exception);
            StatusText = $"模板建立失敗：{exception.Message}";
        }
    }

    public async Task LoadTemplateAsync(
        string path,
        CancellationToken cancellationToken = default)
    {
        try
        {
            var definition = _templateService.LoadTemplate(path);
            _lastTemplateConfigPath = path;
            SaveSettings();
            ApplyLoadedTemplate(definition);

            if (!File.Exists(definition.WordTemplatePath))
            {
                StatusText = "模板已載入，但找不到 Word 模板，請重新選擇。";
            }
            else
            {
                StatusText = $"模板已載入：{definition.TemplateName}";
            }

            if (!string.IsNullOrWhiteSpace(ExcelPath))
            {
                await LoadExcelAsync(ExcelPath, cancellationToken);
            }
        }
        catch (TemplatePersistenceException exception)
        {
            StatusText = exception.UserMessage;
        }
        catch (Exception exception)
        {
            _diagnosticsLogger.LogException("LoadTemplate", "TemplateLoadFailed", exception);
            StatusText = $"載入模板失敗：{exception.Message}";
        }
    }

    private void RestoreLastTemplateAtStartup()
    {
        if (string.IsNullOrWhiteSpace(_lastTemplateConfigPath))
        {
            return;
        }

        try
        {
            var definition = _templateService.LoadTemplate(_lastTemplateConfigPath);
            ApplyLoadedTemplate(definition);
            StatusText = File.Exists(definition.WordTemplatePath)
                ? "已自動載入最近使用的模板。"
                : "已載入最近使用的模板，但找不到 Word 模板，請重新指定。";
        }
        catch (TemplatePersistenceException exception)
        {
            _diagnosticsLogger.LogException(
                "RestoreLastTemplate",
                exception.Code.ToString(),
                exception);
            StatusText = File.Exists(_lastTemplateConfigPath)
                ? "無法載入最近使用的模板，請重新選擇模板。"
                : "最近使用的模板已不存在，請重新載入模板。";
        }
        catch (Exception exception)
        {
            _diagnosticsLogger.LogException("RestoreLastTemplate", "StartupTemplateRestoreFailed", exception);
            StatusText = "無法載入最近使用的模板，請重新選擇模板。";
        }
    }

    private void ApplyLoadedTemplate(TemplateDefinition definition)
    {
        CurrentTemplateDefinition = definition;
        TemplateName = definition.TemplateName;
        WordTemplatePath = definition.WordTemplatePath;
        HeaderRowText = definition.HeaderRowNumber.ToString();
        if (!string.IsNullOrWhiteSpace(definition.OutputFileNamePattern))
        {
            ApplyTemplateOutputPattern(definition.OutputFileNamePattern);
        }
        else if (!_isOutputPatternUserDefined)
        {
            ApplyAutomaticOutputPattern(definition.WordTemplatePath);
        }

        InvalidateSelectionPreview();
        MappingPreview.Clear();
        OnPropertyChanged(nameof(CurrentTemplateDefinition));
    }

    public async Task RefreshSelectionPreviewAsync(CancellationToken cancellationToken = default)
    {
        CancelPendingPreviewRefresh();
        await RefreshSelectionPreviewCoreAsync(cancellationToken, _previewRequestVersion);
    }

    private void RequestSelectionPreviewRefresh()
    {
        CancelPendingPreviewRefresh();
        ClearSelectionPreview();

        if (!IsGenerateDocumentsMode
            || CurrentTemplateDefinition is null
            || string.IsNullOrWhiteSpace(ExcelPath)
            || string.IsNullOrWhiteSpace(SelectedWorksheet))
        {
            return;
        }

        SelectionSummaryText = "正在更新資料選擇預覽…";
        StatusText = "正在檢查 Excel 格式並更新資料預覽。";
        var cancellationSource = new CancellationTokenSource();
        _previewRefreshCancellationSource = cancellationSource;
        _ = RefreshSelectionPreviewDebouncedAsync(cancellationSource, _previewRequestVersion);
    }

    private async Task RefreshSelectionPreviewDebouncedAsync(
        CancellationTokenSource cancellationSource,
        int requestVersion)
    {
        try
        {
            await Task.Delay(180, cancellationSource.Token);
            await RefreshSelectionPreviewCoreAsync(cancellationSource.Token, requestVersion);
        }
        catch (OperationCanceledException)
        {
            // A newer field change superseded this preview request.
        }
        finally
        {
            if (ReferenceEquals(_previewRefreshCancellationSource, cancellationSource))
            {
                _previewRefreshCancellationSource = null;
                cancellationSource.Dispose();
            }
        }
    }

    private async Task RefreshSelectionPreviewCoreAsync(
        CancellationToken cancellationToken,
        int requestVersion)
    {
        ClearSelectionPreview();

        if (CurrentTemplateDefinition is null || string.IsNullOrWhiteSpace(SelectedWorksheet))
        {
            SelectionSummaryText = "請先選擇模板、Excel 與工作表。";
            return;
        }

        _isPreviewRefreshing = true;
        OnPropertyChanged(nameof(ExcelFormatStatusText));
        try
        {
            if (!TryGetExplicitRuntimeMarkerRow(out var explicitRuntimeMarkerRowNumber))
            {
                return;
            }

            var selectionRule = _selectionRuleParser.Parse(
                SelectedMode.Mode,
                GetCurrentSelectionParameter(),
                SelectedMarkerColumn);

            var selectionPreview = await Task.Run(
                () => _generationService.PreviewSelectionAsync(
                    ExcelPath,
                    CurrentTemplateDefinition,
                    selectionRule,
                    cancellationToken,
                    explicitRuntimeMarkerRowNumber),
                cancellationToken);

            cancellationToken.ThrowIfCancellationRequested();
            if (requestVersion != _previewRequestVersion)
            {
                return;
            }

            _selectionRule = selectionRule;
            _selectionPreview = selectionPreview;

            var notice = string.IsNullOrWhiteSpace(_selectionPreview.Notice)
                ? string.Empty
                : $"\n{_selectionPreview.Notice}";
            SelectionSummaryText =
                $"找到 {_selectionPreview.AvailableRecordCount} 筆有效資料\n"
                + $"符合條件：{_selectionPreview.TotalSelected} 筆\n"
                + $"預計產生：{_selectionPreview.TotalSelected} 個 Word{notice}";
            _excelSchemaValid = true;
            StatusText = _selectionPreview.TotalSelected == 0
                ? "沒有符合條件的資料。"
                : "預覽已更新，可以開始批次產生。";
        }
        catch (OperationCanceledException)
        {
            return;
        }
        catch (SelectionValidationException exception)
        {
            SelectionSummaryText = exception.UserMessage;
            StatusText = exception.UserMessage;
        }
        catch (TemplateDefinitionException exception)
        {
            _excelSchemaValid = false;
            SelectionSummaryText = exception.UserMessage;
            StatusText = exception.UserMessage;
        }
        catch (Exception exception)
        {
            _diagnosticsLogger.LogException("RefreshSelectionPreview", "SelectionPreviewFailed", exception);
            SelectionSummaryText = "無法計算資料選擇預覽。";
            StatusText = "無法計算資料選擇預覽，請確認 Excel 與模板設定。";
        }
        finally
        {
            if (requestVersion == _previewRequestVersion)
            {
                _isPreviewRefreshing = false;
                OnPropertyChanged(nameof(ExcelFormatStatusText));
            }
        }

        if (requestVersion == _previewRequestVersion)
        {
            OnPropertyChanged(nameof(CanGenerateBatch));
        }
    }

    public async Task GenerateBatchAsync(CancellationToken cancellationToken = default)
    {
        if (CurrentTemplateDefinition is null || _selectionRule is null || _selectionPreview is null)
        {
            StatusText = "請先建立有效的資料選擇預覽。";
            return;
        }

        IsGenerating = true;
        using var linkedCancellation = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
        _batchCancellationSource = linkedCancellation;
        OnPropertyChanged(nameof(CanCancelBatch));
        BatchResults.Clear();
        ProgressValue = 0;
        ProgressMaximum = Math.Max(1, _selectionPreview.TotalSelected);
        ProgressText = $"正在產生 0 / {_selectionPreview.TotalSelected}";

        var progress = new Progress<BatchProgress>(value =>
        {
            ProgressValue = value.Completed;
            ProgressMaximum = Math.Max(1, value.Total);
            ProgressText = $"正在產生 {value.Completed} / {value.Total}";
        });

        try
        {
            if (!TryGetExplicitRuntimeMarkerRow(out var explicitRuntimeMarkerRowNumber))
            {
                return;
            }

            var result = await Task.Run(
                () => _generationService.GenerateBatchAsync(
                    ExcelPath,
                    CurrentTemplateDefinition,
                    _selectionRule,
                    OutputDirectory,
                    progress,
                    linkedCancellation.Token,
                    explicitRuntimeMarkerRowNumber),
                CancellationToken.None);

            if (result.IsCancelled)
            {
                foreach (var rowResult in result.Results)
                {
                    BatchResults.Add(BatchResultItem.From(rowResult));
                }

                ProgressValue = result.Results.Count;
                ProgressText = $"已取消：{result.Results.Count} / {result.TotalSelected}";
                StatusText = $"已取消\n已完成：{result.SuccessCount}\n失敗：{result.FailureCount}\n未處理：{result.UnprocessedCount}";
                return;
            }

            foreach (var rowResult in result.Results)
            {
                BatchResults.Add(BatchResultItem.From(rowResult));
            }

            ProgressValue = result.TotalSelected;
            ProgressText = $"完成 {result.TotalSelected} / {result.TotalSelected}";
            StatusText = $"完成\n成功：{result.SuccessCount}\n失敗：{result.FailureCount}"
                + (string.IsNullOrWhiteSpace(result.Notice) ? string.Empty : $"\n{result.Notice}");
        }
        catch (BatchGenerationException exception)
        {
            StatusText = $"批次產生停止：{exception.UserMessage}";
            ProgressText = string.Empty;
        }
        catch (OperationCanceledException)
        {
            StatusText = "批次產生已取消。";
            ProgressText = string.Empty;
        }
        catch (Exception exception)
        {
            _diagnosticsLogger.LogException("GenerateBatch", "BatchGenerationFailed", exception);
            StatusText = "批次產生失敗。";
            ProgressText = string.Empty;
        }
        finally
        {
            _batchCancellationSource = null;
            OnPropertyChanged(nameof(CanCancelBatch));
            IsGenerating = false;
            SaveSettings();
        }
    }

    public void CancelBatch()
    {
        if (!CanCancelBatch)
        {
            return;
        }

        StatusText = "正在取消，請稍候完成目前文件...";
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

    private string GetCurrentSelectionParameter() => SelectedMode.Mode switch
    {
        SelectionMode.Marker => MarkerTokensText,
        SelectionMode.ExcelRows => RowExpressionText,
        SelectionMode.Latest => LatestCountText,
        _ => string.Empty
    };

    private bool TryGetExplicitRuntimeMarkerRow(out int? rowNumber)
    {
        rowNumber = null;
        if (string.IsNullOrWhiteSpace(RuntimeMarkerRowText))
        {
            return true;
        }

        if (int.TryParse(RuntimeMarkerRowText.Trim(), out var parsedRow) && parsedRow > 0)
        {
            rowNumber = parsedRow;
            return true;
        }

        StatusText = "目前 Excel 欄位對應列必須是正整數，或留白。";
        SelectionSummaryText = StatusText;
        return false;
    }

    private bool IsOutputPatternValid()
    {
        if (CurrentTemplateDefinition is null)
        {
            return false;
        }

        try
        {
            var sampleRecord = new ExcelRecord
            {
                RowNumber = 1,
                Values = new Dictionary<string, string?>(),
                ColumnValues = CurrentTemplateDefinition.FieldMappings
                    .Select(mapping => mapping.ColumnIndex)
                    .Distinct()
                    .ToDictionary(columnIndex => columnIndex, _ => (string?)"sample")
            };
            _filenamePolicy.ResolveBaseName(
                CurrentTemplateDefinition.OutputFileNamePattern,
                CurrentTemplateDefinition,
                sampleRecord);
            return true;
        }
        catch (GenerationException)
        {
            return false;
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

            var defaultDirectory = Path.Combine(excelDirectory, "ReceiptXcel輸出");
            _isOutputDirectoryAutomaticDefault = true;
            if (SetField(ref _outputDirectory, defaultDirectory))
            {
                SaveSettings();
                OnPropertyChanged(nameof(CanGenerateBatch));
            }
        }
        catch (Exception exception) when (exception is ArgumentException or IOException or NotSupportedException)
        {
            _diagnosticsLogger.LogException("ApplyExcelOutputDefault", "OutputDefaultPathFailed", exception);
        }
    }

    private void ApplyAutomaticOutputPattern(string wordTemplatePath)
    {
        if (_isOutputPatternUserDefined || _isOutputPatternFromTemplate)
        {
            return;
        }

        var stem = Path.GetFileNameWithoutExtension(wordTemplatePath);
        if (string.IsNullOrWhiteSpace(stem))
        {
            return;
        }

        var defaultPattern = $"{stem}_{{{{ROW}}}}";
        _isOutputPatternAutomaticDefault = true;
        if (SetField(ref _outputPattern, defaultPattern))
        {
            UpdateTemplateOutputPattern(defaultPattern);
            SaveSettings();
            OnPropertyChanged(nameof(CanGenerateBatch));
        }
    }

    private void ApplyTemplateOutputPattern(string pattern)
    {
        _isOutputPatternUserDefined = false;
        _isOutputPatternAutomaticDefault = false;
        _isOutputPatternFromTemplate = true;
        if (SetField(ref _outputPattern, pattern))
        {
            UpdateTemplateOutputPattern(pattern);
            OnPropertyChanged(nameof(CanGenerateBatch));
        }
    }

    private void UpdateTemplateOutputPattern(string? pattern)
    {
        if (CurrentTemplateDefinition is null)
        {
            return;
        }

        CurrentTemplateDefinition = CurrentTemplateDefinition with
        {
            OutputFileNamePattern = string.IsNullOrWhiteSpace(pattern) ? null : pattern.Trim()
        };
        OnPropertyChanged(nameof(CurrentTemplateDefinition));
    }

    private void SaveSettings()
    {
        _settingsService.Save(new ApplicationSettings
        {
            LastTemplateConfigPath = _lastTemplateConfigPath,
            LastOutputDirectory = _isOutputDirectoryUserDefined
                && !_isOutputDirectoryAutomaticDefault
                && !string.IsNullOrWhiteSpace(OutputDirectory)
                ? OutputDirectory
                : null,
            LastSelectionMode = SelectedMode.Mode.ToString(),
            LastOutputFilenamePattern = _isOutputPatternUserDefined && !string.IsNullOrWhiteSpace(OutputPattern)
                ? OutputPattern
                : null,
            WindowWidth = _windowWidth,
            WindowHeight = _windowHeight
        });
    }

    private void CancelPendingPreviewRefresh()
    {
        _previewRequestVersion++;
        if (_previewRefreshCancellationSource is null)
        {
            return;
        }

        _previewRefreshCancellationSource.Cancel();
        _previewRefreshCancellationSource.Dispose();
        _previewRefreshCancellationSource = null;
    }

    private void InvalidateSelectionPreview()
    {
        CancelPendingPreviewRefresh();
        ClearSelectionPreview();
    }

    private void ClearSelectionPreview()
    {
        _selectionRule = null;
        _selectionPreview = null;
        _excelSchemaValid = null;
        _isPreviewRefreshing = false;
        BatchResults.Clear();
        SelectionSummaryText = "尚未建立資料選擇預覽。";
        OnPropertyChanged(nameof(ExcelFormatStatusText));
        OnPropertyChanged(nameof(CanGenerateBatch));
    }

    private bool SetField<T>(ref T field, T value, [CallerMemberName] string? propertyName = null)
    {
        if (EqualityComparer<T>.Default.Equals(field, value))
        {
            return false;
        }

        field = value;
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        return true;
    }

    private void OnPropertyChanged([CallerMemberName] string? propertyName = null)
    {
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        if (propertyName == nameof(CurrentTemplateDefinition))
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(TemplateSummaryText)));
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(HasTemplate)));
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(HasSingleMappingTemplate)));
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(RuntimeMarkerInstructionText)));
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(TemplateDisplayName)));
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(TemplateFieldCountText)));
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(ExcelFormatStatusText)));
        }
    }
}

public sealed record SelectionModeOption(SelectionMode Mode, string DisplayName);

public sealed record MappingPreviewItem(string Placeholder, string ExcelColumn);

public sealed record BatchResultItem(string StatusIcon, string RowText, string Detail)
{
    public static BatchResultItem From(RowGenerationResult result) =>
        result.Success
            ? new("✓", $"Row {result.ExcelRowNumber}", result.OutputPath ?? string.Empty)
            : new("✕", $"Row {result.ExcelRowNumber}", result.Message ?? "產生失敗");
}

public enum MainWindowMode
{
    GenerateDocuments,
    CreateTemplate
}
