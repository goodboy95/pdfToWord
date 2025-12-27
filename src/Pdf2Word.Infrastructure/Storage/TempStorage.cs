using Pdf2Word.Core.Services;

namespace Pdf2Word.Infrastructure.Storage;

public sealed class TempStorage : ITempStorage
{
    public string JobRoot { get; }
    public string PagesDir { get; }
    public string TablesDir { get; }
    public string IrDir { get; }
    public string LogsDir { get; }

    public TempStorage(string? root = null)
    {
        JobRoot = root ?? Path.Combine(Path.GetTempPath(), "Pdf2Word", $"job_{Guid.NewGuid():N}");
        PagesDir = Path.Combine(JobRoot, "pages");
        TablesDir = Path.Combine(JobRoot, "tables");
        IrDir = Path.Combine(JobRoot, "ir");
        LogsDir = Path.Combine(JobRoot, "logs");
    }

    public void EnsureCreated()
    {
        Directory.CreateDirectory(JobRoot);
        Directory.CreateDirectory(PagesDir);
        Directory.CreateDirectory(TablesDir);
        Directory.CreateDirectory(IrDir);
        Directory.CreateDirectory(LogsDir);
    }

    public string GetPageImagePath(int pageNumber, string suffix)
    {
        return Path.Combine(PagesDir, $"p{pageNumber:000}_{suffix}.png");
    }

    public string GetTableImagePath(int pageNumber, int tableIndex, string suffix)
    {
        return Path.Combine(TablesDir, $"p{pageNumber:000}_t{tableIndex:00}_{suffix}.png");
    }

    public string GetIrPath(string name)
    {
        return Path.Combine(IrDir, name);
    }

    public string GetLogPath(string name)
    {
        return Path.Combine(LogsDir, name);
    }

    public void Cleanup()
    {
        if (Directory.Exists(JobRoot))
        {
            Directory.Delete(JobRoot, true);
        }
    }
}
