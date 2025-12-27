using Pdf2Word.Core.Logging;

namespace Pdf2Word.App.Services;

public sealed class UiLogSink : ILogSink
{
    public event Action<LogEntry>? LogPublished;

    public void Publish(LogEntry entry)
    {
        LogPublished?.Invoke(entry);
    }
}
