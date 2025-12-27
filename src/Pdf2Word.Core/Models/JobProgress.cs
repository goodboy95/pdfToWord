namespace Pdf2Word.Core.Models;

public sealed class JobProgress
{
    public int TotalPages { get; set; }
    public int CompletedPages { get; set; }
    public int? CurrentPage { get; set; }
    public JobStage Stage { get; set; }
    public string Message { get; set; } = string.Empty;
}
