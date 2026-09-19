using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Threading;
using XlsxDocxGenerator.Services.Diagnostics;
using XlsxDocxGenerator.Services.Settings;

namespace XlsxDocxGenerator.Tests;

public sealed class Phase5CStartupWpfTests
{
    [Fact]
    public void MainWindowReachesLoadedWithoutDispatcherException()
    {
        Exception? failure = null;
        var startupReady = false;
        using var completed = new ManualResetEventSlim();
        var settingsDirectory = Path.Combine(
            Path.GetTempPath(),
            $"docxcel-phase5c-wpf-{Guid.NewGuid():N}");

        var thread = new Thread(() =>
        {
            var dispatcher = Dispatcher.CurrentDispatcher;
            Application? application = null;
            XlsxDocxGenerator.MainWindow? window = null;
            try
            {
                application = new XlsxDocxGenerator.App
                {
                    ShutdownMode = ShutdownMode.OnExplicitShutdown
                };
                Application.LoadComponent(
                    application,
                    new Uri($"/{typeof(XlsxDocxGenerator.App).Assembly.GetName().Name};component/App.xaml", UriKind.Relative));
                application.DispatcherUnhandledException += (_, args) =>
                {
                    failure = args.Exception;
                    args.Handled = true;
                    dispatcher.BeginInvokeShutdown(DispatcherPriority.Send);
                };

                Directory.CreateDirectory(settingsDirectory);
                var viewModel = new XlsxDocxGenerator.ViewModels.MainWindowViewModel(
                    new SettingsService(Path.Combine(settingsDirectory, "settings.json")),
                    new DiagnosticsLogger(Path.Combine(settingsDirectory, "Logs")));
                window = new XlsxDocxGenerator.MainWindow(viewModel);
                window.StartupReady += (_, _) => startupReady = true;
                window.Show();

                dispatcher.BeginInvoke(
                    DispatcherPriority.ApplicationIdle,
                    new Action(() =>
                    {
                        if (!startupReady || !window.IsLoaded || !window.IsStartupReady)
                        {
                            failure = new InvalidOperationException("MainWindow did not reach Loaded.");
                        }

                        if (window.Title != "ReceiptXcel｜自行收納款項收據產生工具")
                        {
                            failure = new InvalidOperationException("MainWindow title changed after assembly rename.");
                        }

                        window.Close();
                        dispatcher.BeginInvokeShutdown(DispatcherPriority.Background);
                    }));

                Dispatcher.Run();
            }
            catch (Exception exception)
            {
                failure = exception;
            }
            finally
            {
                window?.Close();
                application?.Shutdown();
                completed.Set();
            }
        });

        thread.SetApartmentState(ApartmentState.STA);
        thread.Start();

        Assert.True(completed.Wait(TimeSpan.FromSeconds(15)), "WPF startup test timed out.");
        Assert.True(startupReady, failure?.ToString() ?? "MainWindow did not reach Loaded.");
        Assert.Null(failure);

        if (Directory.Exists(settingsDirectory))
        {
            Directory.Delete(settingsDirectory, recursive: true);
        }
    }
}
