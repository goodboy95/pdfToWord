using System;
using System.IO;
using System.Text;
using Pdf2Word.Core.Logging;

namespace Pdf2Word.Infrastructure.Logging;

public sealed class InstallLogSink : ILogSink
{
    private readonly object _lock = new();
    private readonly string _logDir;
    private readonly string _logPath;

    public InstallLogSink()
    {
        _logDir = Path.Combine(AppContext.BaseDirectory, "log");
        _logPath = Path.Combine(_logDir, "app.log");
    }

    public void Publish(LogEntry entry)
    {
        try
        {
            Directory.CreateDirectory(_logDir);
            var line = new StringBuilder()
                .Append(entry.TimestampUtc.ToString("O"))
                .Append('\t').Append(entry.Severity)
                .Append('\t').Append(entry.Stage)
                .Append('\t').Append(entry.PageNumber?.ToString() ?? "-")
                .Append('\t').Append(entry.TableIndex?.ToString() ?? "-")
                .Append('\t').Append(entry.Attempt)
                .Append('\t').Append(entry.ErrorCode ?? string.Empty)
                .Append('\t').Append(entry.Message)
                .ToString();

            lock (_lock)
            {
                File.AppendAllText(_logPath, line + Environment.NewLine);
            }
        }
        catch
        {
            // log sink should never throw
        }
    }
}
