using Pdf2Word.Core.Options;

namespace Pdf2Word.Core.Models;

public sealed class ConvertJobRequest
{
    public string PdfPath { get; set; } = string.Empty;
    public string OutputPath { get; set; } = string.Empty;
    public string PageRangeText { get; set; } = string.Empty;
    public AppOptions Options { get; set; } = new();
    public string? GeminiApiKey { get; set; }
}

public sealed class ConvertResult
{
    public JobStatus Status { get; set; } = JobStatus.Pending;
    public string? OutputPath { get; set; }
    public List<FailureInfo> Failures { get; set; } = new();
    public TimeSpan Elapsed { get; set; }
}

public sealed class FailureInfo
{
    public int? PageNumber { get; set; }
    public int? TableIndex { get; set; }
    public string ErrorCode { get; set; } = string.Empty;
    public string Message { get; set; } = string.Empty;
    public ErrorSeverity Severity { get; set; }
    public JobStage Stage { get; set; }
    public int Attempt { get; set; } = 1;
}
