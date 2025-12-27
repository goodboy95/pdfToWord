using Pdf2Word.Core.Models;

namespace Pdf2Word.Core.Options;

public sealed class AppOptions
{
    public RenderOptions Render { get; init; } = new();
    public LayoutOptions Layout { get; init; } = new();
    public RuntimeOptions Runtime { get; init; } = new();
    public PreprocessOptions Preprocess { get; init; } = new();
    public TableDetectOptions TableDetect { get; init; } = new();
    public GeminiOptions Gemini { get; init; } = new();
    public DocxWriteOptions Docx { get; init; } = new();
    public DiagnosticsOptions Diagnostics { get; init; } = new();
    public ValidationOptions Validation { get; init; } = new();
    public OutputOptions Output { get; init; } = new();
}

public sealed class RenderOptions
{
    public int Dpi { get; set; } = 300;
    public RenderColorMode ColorMode { get; set; } = RenderColorMode.Color;
    public int MaxPreviewDpi { get; set; } = 150;
}

public enum RenderColorMode
{
    Color,
    Grayscale
}

public sealed class LayoutOptions
{
    public HeaderFooterRemoveMode HeaderFooterMode { get; set; } = HeaderFooterRemoveMode.None;
    public double HeaderPercent { get; set; } = 0.06;
    public double FooterPercent { get; set; } = 0.06;
    public double MaxCropTotalPercent { get; set; } = 0.30;
    public PageSizeMode PageSizeMode { get; set; } = PageSizeMode.FollowPdf;
    public int MarginTopTwips { get; set; } = 1134;
    public int MarginBottomTwips { get; set; } = 1134;
    public int MarginLeftTwips { get; set; } = 1134;
    public int MarginRightTwips { get; set; } = 1134;
}

public sealed class RuntimeOptions
{
    public int PageConcurrency { get; set; } = 2;
    public int GeminiConcurrency { get; set; } = 2;
}

public sealed class PreprocessOptions
{
    public bool EnableDeskew { get; set; } = true;
    public ContrastEnhanceMode ContrastEnhance { get; set; } = ContrastEnhanceMode.CLAHE;
    public ClaheOptions Clahe { get; set; } = new();
    public DenoiseMode Denoise { get; set; } = DenoiseMode.Median3;
    public BinarizeMode Binarize { get; set; } = BinarizeMode.Adaptive;
    public AdaptiveThresholdOptions Adaptive { get; set; } = new();
    public DeskewOptions Deskew { get; set; } = new();
}

public enum ContrastEnhanceMode
{
    None,
    Linear,
    CLAHE
}

public sealed class ClaheOptions
{
    public double ClipLimit { get; set; } = 2.5;
    public int TileGridSize { get; set; } = 8;
}

public enum DenoiseMode
{
    None,
    Gaussian3,
    Median3
}

public enum BinarizeMode
{
    Adaptive,
    Otsu
}

public sealed class AdaptiveThresholdOptions
{
    public int BlockSize { get; set; } = 41;
    public int C { get; set; } = 10;
}

public sealed class DeskewOptions
{
    public double MinAngleDeg { get; set; } = 0.3;
    public double MaxAngleDeg { get; set; } = 10.0;
}

public sealed class TableDetectOptions
{
    public bool Enable { get; set; } = true;
    public double MinAreaRatio { get; set; } = 0.005;
    public int MinWidthPx { get; set; } = 150;
    public int MinHeightPx { get; set; } = 80;
    public int MergeGapPx { get; set; } = 15;

    public int KernelDivisor { get; set; } = 30;
    public int MinKernelPx { get; set; } = 20;
    public int DilateKernelPx { get; set; } = 3;

    public int ClusterEpsPx { get; set; } = 4;
    public double MinCellSizeRatio { get; set; } = 0.02;

    public OwnerConflictPolicy OwnerConflictPolicy { get; set; } = OwnerConflictPolicy.FailAndFallbackText;
    public bool EnableTextTableFallback { get; set; } = true;
    public bool EnableImageTableFallback { get; set; } = false;
}

public enum OwnerConflictPolicy
{
    FailAndFallbackText,
    RetryWithTuning
}

public sealed class GeminiOptions
{
    public string BaseUrl { get; set; } = "";
    public string Model { get; set; } = "";
    public GeminiKeyStorage ApiKeyStorage { get; set; } = GeminiKeyStorage.DPAPI;
    public GeminiTimeouts TimeoutSeconds { get; set; } = new();
    public int MaxRetryCount { get; set; } = 1;
    public int BackoffBaseMs { get; set; } = 500;
    public bool ExpectStrictJson { get; set; } = true;
    public double TableMissingIdThreshold { get; set; } = 0.05;
    public double TableEmptyTextRateThreshold { get; set; } = 0.80;
    public int MinPageTextCharCount { get; set; } = 10;
    public GeminiImageOptions Image { get; set; } = new();
}

public enum GeminiKeyStorage
{
    DPAPI,
    None,
    EnvOnly
}

public sealed class GeminiTimeouts
{
    public int Table { get; set; } = 90;
    public int Page { get; set; } = 90;
}

public sealed class GeminiImageOptions
{
    public int MaxLongEdgePx { get; set; } = 2800;
    public int JpegQuality { get; set; } = 90;
}

public sealed class DocxWriteOptions
{
    public PageSizeMode PageSizeMode { get; set; } = PageSizeMode.FollowPdf;
    public int Dpi { get; set; } = 300;
    public string FontEastAsia { get; set; } = "微软雅黑";
    public string FontAscii { get; set; } = "Calibri";
    public int DefaultFontSizeHalfPoints { get; set; } = 24;
    public int MarginTopTwips { get; set; } = 1134;
    public int MarginBottomTwips { get; set; } = 1134;
    public int MarginLeftTwips { get; set; } = 1134;
    public int MarginRightTwips { get; set; } = 1134;
    public DocxTableOptions Table { get; set; } = new();
    public DocxParagraphOptions Paragraph { get; set; } = new();
    public DocxPageBreakOptions PageBreak { get; set; } = new();
    public bool UseStylesPart { get; set; } = false;
}

public sealed class DocxTableOptions
{
    public DocxTableWidthMode WidthMode { get; set; } = DocxTableWidthMode.FitPageContent;
    public bool SetBorders { get; set; } = true;
}

public enum DocxTableWidthMode
{
    Auto,
    FitPageContent
}

public sealed class DocxParagraphOptions
{
    public bool KeepLineBreaks { get; set; } = true;
}

public sealed class DocxPageBreakOptions
{
    public DocxPageBreakMode ModeA4 { get; set; } = DocxPageBreakMode.PageBreak;
    public DocxPageBreakMode ModeFollowPdf { get; set; } = DocxPageBreakMode.SectionBreakNextPage;
}

public enum DocxPageBreakMode
{
    PageBreak,
    SectionBreakNextPage
}

public sealed class DiagnosticsOptions
{
    public bool KeepTempFiles { get; set; } = false;
    public bool SaveRawGeminiJson { get; set; } = false;
    public bool ExportZip { get; set; } = true;
}

public sealed class ValidationOptions
{
    public bool Enable { get; set; } = true;
    public bool FailFast { get; set; } = false;
    public bool AllowSkipFailedPages { get; set; } = true;
    public TableFallbackPolicy TableStructureOkTextBad { get; set; } = TableFallbackPolicy.KeepStructureEmptyCells;
    public TableFallbackPolicy TableStructureBad { get; set; } = TableFallbackPolicy.FallbackTextTable;
    public TextFallbackPolicy TextEmpty { get; set; } = TextFallbackPolicy.SingleParagraph;
}

public enum TableFallbackPolicy
{
    KeepStructureEmptyCells,
    FallbackTextTable,
    InsertImage
}

public enum TextFallbackPolicy
{
    SingleParagraph,
    Empty
}

public sealed class OutputOptions
{
    public string OutputDirectory { get; set; } = "";
    public OutputFileNameMode FileNameMode { get; set; } = OutputFileNameMode.Auto;
    public bool Overwrite { get; set; } = false;
}

public enum OutputFileNameMode
{
    Auto,
    Custom
}
