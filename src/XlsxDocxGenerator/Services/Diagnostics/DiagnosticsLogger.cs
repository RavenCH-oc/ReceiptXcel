using System.IO;
using System.Reflection;
using System.Text;

namespace XlsxDocxGenerator.Services.Diagnostics;

public sealed class DiagnosticsLogger
{
    private static readonly UTF8Encoding Utf8NoBom = new(encoderShouldEmitUTF8Identifier: false);

    public DiagnosticsLogger(string? logDirectory = null)
    {
        LogDirectory = logDirectory ?? Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "ReceiptXcel",
            "Logs");
    }

    public string LogDirectory { get; }

    public void LogException(
        string operation,
        string? internalErrorCode,
        Exception exception)
    {
        try
        {
            Directory.CreateDirectory(LogDirectory);
            var logPath = Path.Combine(LogDirectory, $"{DateTime.UtcNow:yyyy-MM-dd}.log");
            var builder = new StringBuilder()
                .AppendLine($"TimestampUtc: {DateTime.UtcNow:O}")
                .AppendLine($"ApplicationVersion: {Assembly.GetExecutingAssembly().GetName().Version?.ToString() ?? "0.1.0"}")
                .AppendLine($"Operation: {operation}")
                .AppendLine($"InternalErrorCode: {internalErrorCode ?? "(none)"}")
                .AppendLine($"ExceptionType: {exception.GetType().FullName}")
                .AppendLine("StackTrace:")
                // Exception.Message and Exception.ToString() can include Excel
                // values or generated receipt content. Diagnostics intentionally
                // retain only the type and call stack for release privacy.
                .AppendLine(exception.StackTrace ?? "(no stack trace)")
                .AppendLine(new string('-', 80));
            File.AppendAllText(logPath, builder.ToString(), Utf8NoBom);
        }
        catch
        {
            // Diagnostics must never become the primary application failure.
        }
    }
}
