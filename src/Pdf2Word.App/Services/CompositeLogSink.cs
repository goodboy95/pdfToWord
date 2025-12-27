using System.Linq;
using Pdf2Word.Core.Logging;

namespace Pdf2Word.App.Services;

public sealed class CompositeLogSink : ILogSink
{
    private readonly IReadOnlyList<ILogSink> _sinks;

    public CompositeLogSink(params ILogSink[] sinks)
    {
        _sinks = sinks.Where(s => s != null).ToList();
    }

    public void Publish(LogEntry entry)
    {
        foreach (var sink in _sinks)
        {
            sink.Publish(entry);
        }
    }
}
