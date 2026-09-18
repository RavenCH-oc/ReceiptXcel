using System.Text;
using XlsxDocxGenerator.Services.Diagnostics;

namespace XlsxDocxGenerator.Tests;

public sealed class Phase5CStartupBindingTests
{
    [Fact]
    public void ProgressBarOutputBindingsAreExplicitlyOneWay()
    {
        var solutionDirectory = FindSolutionDirectory();
        var xamlPath = Path.Combine(
            solutionDirectory,
            "src",
            "XlsxDocxGenerator",
            "MainWindow.xaml");
        var xaml = File.ReadAllText(xamlPath);

        Assert.Contains(
            "Maximum=\"{Binding ProgressMaximum, Mode=OneWay}\"",
            xaml);
        Assert.Contains(
            "Value=\"{Binding ProgressValue, Mode=OneWay}\"",
            xaml);
        Assert.DoesNotContain(
            "Value=\"{Binding ProgressValue}\"",
            xaml);
    }

    [Fact]
    public void DiagnosticsLogIsUtf8AndDoesNotPersistExceptionMessages()
    {
        var directory = Path.Combine(
            Path.GetTempPath(),
            $"docxcel-phase5c-{Guid.NewGuid():N}");
        try
        {
            var logger = new DiagnosticsLogger(directory);
            const string sensitiveMessage = "測試例外訊息 PAYER_SENTINEL 6/6匯款";
            logger.LogException(
                "啟動驗證",
                "UnicodeTest",
                new InvalidOperationException(sensitiveMessage));

            var logPath = Directory.GetFiles(directory, "*.log").Single();
            var text = new UTF8Encoding(encoderShouldEmitUTF8Identifier: false, throwOnInvalidBytes: true)
                .GetString(File.ReadAllBytes(logPath));

            Assert.Contains("啟動驗證", text);
            Assert.Contains("ExceptionType: System.InvalidOperationException", text);
            Assert.DoesNotContain(sensitiveMessage, text, StringComparison.Ordinal);
        }
        finally
        {
            if (Directory.Exists(directory))
            {
                Directory.Delete(directory, recursive: true);
            }
        }
    }

    private static string FindSolutionDirectory()
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

        throw new DirectoryNotFoundException("Could not locate the solution directory.");
    }
}
