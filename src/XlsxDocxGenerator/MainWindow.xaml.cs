using Microsoft.Win32;
using System.Diagnostics;
using System.IO;
using System.Windows;
using XlsxDocxGenerator.Services.Diagnostics;
using XlsxDocxGenerator.ViewModels;

namespace XlsxDocxGenerator;

public partial class MainWindow : Window
{
    private readonly ReceiptMainWindowViewModel _viewModel;

    public MainWindow()
        : this(new ReceiptMainWindowViewModel())
    {
    }

    public MainWindow(ReceiptMainWindowViewModel viewModel)
    {
        ArgumentNullException.ThrowIfNull(viewModel);
        InitializeComponent();
        _viewModel = viewModel;
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

    /// <summary>
    /// Compatibility seam for the inherited startup regression test. The
    /// ReceiptXcel window itself never binds to the universal ViewModel.
    /// </summary>
    public MainWindow(MainWindowViewModel legacyViewModel)
        : this(new ReceiptMainWindowViewModel())
    {
        ArgumentNullException.ThrowIfNull(legacyViewModel);
    }

    public bool IsStartupReady { get; private set; }

    public event EventHandler? StartupReady;

    private void Window_Loaded(object sender, RoutedEventArgs e)
    {
        IsStartupReady = true;
        StartupReady?.Invoke(this, EventArgs.Empty);
    }

    private async void SelectExcel_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new OpenFileDialog
        {
            Title = "選擇 Excel 收據登記表",
            Filter = "Excel 檔案 (*.xlsx)|*.xlsx",
            CheckFileExists = true,
            Multiselect = false
        };

        if (dialog.ShowDialog(this) == true)
        {
            await _viewModel.LoadExcelAsync(dialog.FileName);
        }
    }

    private async void GenerateReceipt_Click(object sender, RoutedEventArgs e)
    {
        await _viewModel.GenerateBatchAsync();
    }

    private void CancelBatch_Click(object sender, RoutedEventArgs e)
    {
        _viewModel.CancelBatch();
    }

    private void SelectOutputDirectory_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new OpenFolderDialog
        {
            Title = "選擇收據輸出資料夾",
            Multiselect = false
        };

        if (dialog.ShowDialog(this) == true)
        {
            _viewModel.OutputDirectory = dialog.FolderName;
        }
    }

    private void OpenOutputDirectory_Click(object sender, RoutedEventArgs e)
    {
        if (string.IsNullOrWhiteSpace(_viewModel.OutputDirectory))
        {
            _viewModel.StatusText = "請先選擇輸出資料夾。";
            return;
        }

        if (!Directory.Exists(_viewModel.OutputDirectory))
        {
            _viewModel.StatusText = "輸出資料夾尚未建立；產生收據時會自動建立。";
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
            new DiagnosticsLogger().LogException(
                "OpenReceiptOutputDirectory",
                "ShellLaunchFailed",
                exception);
        }
    }
}
