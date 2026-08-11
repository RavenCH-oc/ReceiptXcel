using System.Windows;
using System.Windows.Threading;
using XlsxDocxGenerator.Services.Diagnostics;

namespace XlsxDocxGenerator;

public partial class App : Application
{
    protected override void OnStartup(StartupEventArgs e)
    {
        DispatcherUnhandledException += OnDispatcherUnhandledException;
        base.OnStartup(e);
    }

    private static void OnDispatcherUnhandledException(
        object sender,
        DispatcherUnhandledExceptionEventArgs e)
    {
        new DiagnosticsLogger().LogException("UnhandledUiException", "UnhandledUiException", e.Exception);
        MessageBox.Show(
            "ReceiptXcel 啟動或操作失敗。\n請確認檔案未損壞或被其他程式鎖定後再試一次。",
            "ReceiptXcel",
            MessageBoxButton.OK,
            MessageBoxImage.Warning);
        e.Handled = true;
    }
}
