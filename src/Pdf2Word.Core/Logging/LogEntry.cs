using Pdf2Word.Core.Models;

namespace Pdf2Word.Core.Logging;

public sealed class LogEntry
{
    public DateTime TimestampUtc { get; set; } = DateTime.UtcNow;
    public ErrorSeverity Severity { get; set; } = ErrorSeverity.Warning;
    public JobStage Stage { get; set; } = JobStage.Init;
    public int? PageNumber { get; set; }
    public int? TableIndex { get; set; }
    public string Message { get; set; } = string.Empty;
    public string? ErrorCode { get; set; }
    public int Attempt { get; set; } = 1;
    public long? ElapsedMs { get; set; }
}

public interface ILogSink
{
    void Publish(LogEntry entry);
}
