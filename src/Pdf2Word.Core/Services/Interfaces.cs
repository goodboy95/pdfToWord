using Pdf2Word.Core.Models;
using Pdf2Word.Core.Models.Ir;
using Pdf2Word.Core.Options;

namespace Pdf2Word.Core.Services;

public interface IConversionService
{
    Task<ConvertResult> ConvertAsync(ConvertJobRequest request, IProgress<JobProgress> progress, CancellationToken ct);
}

public interface IPdfRenderer
{
    int GetPageCount(string pdfPath);
    System.Drawing.Bitmap RenderPage(string pdfPath, int pageIndex0Based, int dpi);
}

public interface IPageImagePipeline
{
    PageImageBundle Process(System.Drawing.Bitmap renderedPage, CropOptions crop, PreprocessOptions options, int pageNumber);
}

public sealed class CropOptions
{
    public HeaderFooterRemoveMode Mode { get; set; } = HeaderFooterRemoveMode.None;
    public double HeaderPercent { get; set; } = 0.06;
    public double FooterPercent { get; set; } = 0.06;
}

public interface ITableEngine
{
    IReadOnlyList<TableDetection> DetectTables(PageImageBundle bundle, TableDetectOptions options);
    System.Drawing.Bitmap MaskTables(System.Drawing.Bitmap source, IEnumerable<BBox> tableBoxes);
}

public interface IGeminiClient
{
    Task<TableCellOcrResult> RecognizeTableCellsAsync(System.Drawing.Bitmap tableImage, IReadOnlyList<CellBoxForOcr> cells, CancellationToken ct, bool strictJson, int attempt);
    Task<PageTextOcrResult> RecognizePageParagraphsAsync(System.Drawing.Bitmap pageImage, CancellationToken ct, bool strictJson, int attempt);
    Task<TableTextLinesResult> RecognizeTableAsLinesAsync(System.Drawing.Bitmap tableImage, CancellationToken ct, int attempt);
    Task<PageTextFallbackResult> RecognizePageAsSingleTextAsync(System.Drawing.Bitmap pageImage, CancellationToken ct, int attempt);
}

public sealed class CellBoxForOcr
{
    public string Id { get; init; } = string.Empty;
    public int X { get; init; }
    public int Y { get; init; }
    public int W { get; init; }
    public int H { get; init; }
}

public sealed class TableCellOcrResult
{
    public Dictionary<string, string> TextById { get; init; } = new();
    public string? RawJson { get; init; }
}

public sealed class PageTextOcrResult
{
    public List<ParagraphDto> Paragraphs { get; init; } = new();
    public string? RawJson { get; init; }
}

public sealed class ParagraphDto
{
    public string Role { get; init; } = "body";
    public string Text { get; init; } = string.Empty;
}

public sealed class TableTextLinesResult
{
    public List<string> Lines { get; init; } = new();
    public string? RawJson { get; init; }
}

public sealed class PageTextFallbackResult
{
    public string Text { get; init; } = string.Empty;
    public string? RawJson { get; init; }
}

public interface IDocxWriter
{
    Task WriteAsync(DocumentIr doc, DocxWriteOptions options, Stream output, CancellationToken ct);
}

public interface IApiKeyStore
{
    string? GetApiKey();
    void SaveApiKey(string apiKey);
}

public interface ITempStorage
{
    string JobRoot { get; }
    string PagesDir { get; }
    string TablesDir { get; }
    string IrDir { get; }
    string LogsDir { get; }

    string GetPageImagePath(int pageNumber, string suffix);
    string GetTableImagePath(int pageNumber, int tableIndex, string suffix);
    string GetIrPath(string name);
    string GetLogPath(string name);
    void EnsureCreated();
    void Cleanup();
}
