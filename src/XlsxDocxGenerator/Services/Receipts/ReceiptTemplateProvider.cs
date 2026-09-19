using System.IO;
using XlsxDocxGenerator.Services.Diagnostics;

namespace XlsxDocxGenerator.Services.Receipts;

public interface IReceiptTemplateProvider
{
    ReceiptTemplateLease Acquire(CancellationToken cancellationToken = default);
}

/// <summary>
/// A template path is usable only while its lease is alive. Each embedded
/// lease owns one unique directory, never a shared cache or a caller's file.
/// </summary>
public sealed class ReceiptTemplateLease : IDisposable
{
    private readonly string? _ownedDirectory;
    private int _disposed;

    internal ReceiptTemplateLease(string templatePath, string? ownedDirectory = null)
    {
        TemplatePath = templatePath;
        _ownedDirectory = ownedDirectory;
    }

    public string TemplatePath { get; }

    public void Dispose()
    {
        if (Interlocked.Exchange(ref _disposed, 1) == 0 && _ownedDirectory is not null)
        {
            Cleanup(_ownedDirectory);
        }
    }

    internal static void Cleanup(string ownedDirectory)
    {
        try
        {
            if (Directory.Exists(ownedDirectory))
            {
                Directory.Delete(ownedDirectory, recursive: true);
            }
        }
        catch (Exception exception)
        {
            // Cleanup must never hide the batch outcome or crash the UI.
            new DiagnosticsLogger().LogException(
                "ReceiptTemplateCleanup", "TemporaryTemplateCleanupFailed", exception);
        }
    }
}

public sealed class EmbeddedReceiptTemplateProvider : IReceiptTemplateProvider
{
    public const string ResourceName = "ReceiptXcel.Templates.Receipt.docx";
    private readonly string _temporaryRoot;

    public EmbeddedReceiptTemplateProvider(string? temporaryRoot = null)
    {
        _temporaryRoot = Path.GetFullPath(
            temporaryRoot ?? Path.Combine(Path.GetTempPath(), "ReceiptXcel"));
    }

    public ReceiptTemplateLease Acquire(CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();
        var directory = Path.Combine(_temporaryRoot, Guid.NewGuid().ToString("N"));
        var path = Path.Combine(directory, "receipt-template.docx");
        try
        {
            using var resource = typeof(EmbeddedReceiptTemplateProvider).Assembly
                .GetManifestResourceStream(ResourceName)
                ?? throw new InvalidDataException("Embedded receipt template is missing.");
            Directory.CreateDirectory(directory);
            using (var output = new FileStream(path, FileMode.CreateNew, FileAccess.Write, FileShare.None))
            {
                resource.CopyTo(output);
            }
            cancellationToken.ThrowIfCancellationRequested();
            return new ReceiptTemplateLease(path, directory);
        }
        catch (OperationCanceledException)
        {
            ReceiptTemplateLease.Cleanup(directory);
            throw;
        }
        catch (Exception exception)
        {
            ReceiptTemplateLease.Cleanup(directory);
            throw new ReceiptGenerationException(
                ReceiptGenerationErrorCode.ReceiptTemplateInvalid,
                "內建收據模板無法使用，請重新安裝 ReceiptXcel。",
                exception);
        }
    }
}

// Preserve the explicit template override for contract/fixture tests. This
// provider never deletes the caller's file; production defaults to embedding.
internal sealed class FileReceiptTemplateProvider(string path) : IReceiptTemplateProvider
{
    private readonly string _path = Path.GetFullPath(path);

    public ReceiptTemplateLease Acquire(CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();
        return new ReceiptTemplateLease(_path);
    }
}
