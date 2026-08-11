using Microsoft.Win32;
using System.Diagnostics;
using System.IO;
using System.Windows;
using XlsxDocxGenerator.Services.Diagnostics;
using XlsxDocxGenerator.ViewModels;

namespace XlsxDocxGenerator;

public partial class MainWindow : Window
{
    private readonly MainWindowViewModel _viewModel;

    public MainWindow()
        : this(null)
    {
    }

    public MainWindow(MainWindowViewModel? viewModel)
    {
        InitializeComponent();
        _viewModel = viewModel ?? new MainWindowViewModel();
        DataContext = _viewModel;

        if (_viewModel.SavedWindowWidth is > 0)
        {
            Width = _viewModel.SavedWindowWidth.Value;
        }

        if (_viewModel.SavedWindowHeight is > 0)
        {
            Height = _viewModel.SavedWindowHeight.Value;
        }

        Closed += (_, _) => _viewModel.SaveWindowSize(Width, Height);
    }

    public bool IsStartupReady { get; private set; }

    public event EventHandler? StartupReady;

    private void Window_Loaded(object sender, RoutedEventArgs e)
    {
        IsStartupReady = true;
        StartupReady?.Invoke(this, EventArgs.Empty);
    }

    private void GenerateMode_Click(object sender, RoutedEventArgs e)
    {
        _viewModel.ActiveMode = MainWindowMode.GenerateDocuments;
    }

    private void CreateMode_Click(object sender, RoutedEventArgs e)
    {
        _viewModel.ActiveMode = MainWindowMode.CreateTemplate;
    }

    private void UseTemplate_Click(object sender, RoutedEventArgs e)
    {
        _viewModel.ActiveMode = MainWindowMode.GenerateDocuments;
    }

    private async void SelectExcel_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new OpenFileDialog
        {
            Title = "選擇 Excel 檔案",
            Filter = "Excel 檔案 (*.xlsx)|*.xlsx|所有檔案 (*.*)|*.*",
            CheckFileExists = true
        };

        if (dialog.ShowDialog(this) == true)
        {
            await _viewModel.LoadExcelAsync(dialog.FileName);
        }
    }

    private void SelectWordTemplate_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new OpenFileDialog
        {
            Title = "選擇 Word 模板",
            Filter = "Word 文件 (*.docx)|*.docx|所有檔案 (*.*)|*.*",
            CheckFileExists = true
        };

        if (dialog.ShowDialog(this) == true)
        {
            _viewModel.WordTemplatePath = dialog.FileName;
            _viewModel.StatusText = string.IsNullOrWhiteSpace(_viewModel.ExcelPath)
                ? $"已選擇 Word 模板：{dialog.SafeFileName}"
                : $"已選擇 Excel：{Path.GetFileName(_viewModel.ExcelPath)}\n已選擇 Word 模板：{dialog.SafeFileName}";
        }
    }

    private async void ReadMapping_Click(object sender, RoutedEventArgs e)
    {
        await _viewModel.ReadMappingAsync();
    }

    private async void RefreshPreview_Click(object sender, RoutedEventArgs e)
    {
        await _viewModel.RefreshSelectionPreviewAsync();
    }

    private async void GenerateWord_Click(object sender, RoutedEventArgs e)
    {
        await _viewModel.GenerateBatchAsync();
    }

    private void CancelBatch_Click(object sender, RoutedEventArgs e)
    {
        _viewModel.CancelBatch();
    }

    private async void CreateTemplate_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new SaveFileDialog
        {
            Title = "儲存模板設定",
            Filter = "DocXcel 模板 (*.docxcel.json)|*.docxcel.json",
            DefaultExt = ".docxcel.json",
            AddExtension = true
        };

        if (dialog.ShowDialog(this) == true)
        {
            await _viewModel.CreateTemplateAsync(dialog.FileName);
        }
    }

    private async void LoadTemplate_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new OpenFileDialog
        {
            Title = "載入模板設定",
            Filter = "DocXcel 模板 (*.docxcel.json)|*.docxcel.json|所有檔案 (*.*)|*.*",
            CheckFileExists = true
        };

        if (dialog.ShowDialog(this) == true)
        {
            await _viewModel.LoadTemplateAsync(dialog.FileName);
        }
    }

    private void SelectOutputDirectory_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new OpenFolderDialog
        {
            Title = "選擇輸出資料夾",
            Multiselect = false
        };

        if (dialog.ShowDialog(this) == true)
        {
            _viewModel.OutputDirectory = dialog.FolderName;
        }
    }

    private void OpenOutputDirectory_Click(object sender, RoutedEventArgs e)
    {
        if (!Directory.Exists(_viewModel.OutputDirectory))
        {
            _viewModel.StatusText = "輸出資料夾尚未建立，產生文件時會自動建立。";
            return;
        }

        try
        {
            Process.Start(new ProcessStartInfo
            {
                FileName = _viewModel.OutputDirectory,
                UseShellExecute = true
            });
        }
        catch (Exception exception)
        {
            _viewModel.StatusText = "無法開啟輸出資料夾。";
            new DiagnosticsLogger().LogException("OpenOutputDirectory", "ShellLaunchFailed", exception);
        }
    }
}
