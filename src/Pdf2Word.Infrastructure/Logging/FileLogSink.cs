using System;
using System.IO;
using System.Text;
using Pdf2Word.Core.Logging;
using Pdf2Word.Core.Services;

namespace Pdf2Word.Infrastructure.Logging;

public sealed class FileLogSink : ILogSink
{
    private readonly ITempStorage _tempStorage;
    private readonly object _lock = new();

    public FileLogSink(ITempStorage tempStorage)
    {
        _tempStorage = tempStorage;
    }

    public void Publish(LogEntry entry)
    {
        try
        {
            _tempStorage.EnsureCreated();
            var path = _tempStorage.GetLogPath("job.log");
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
                File.AppendAllText(path, line + Environment.NewLine);
            }
        }
        catch
        {
            // log sink should never throw
        }
    }
}
