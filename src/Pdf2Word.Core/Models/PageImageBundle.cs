using Pdf2Word.Core.Models.Ir;

namespace Pdf2Word.Core.Models;

public sealed class PageImageBundle
{
    public int PageNumber { get; init; }
    public required System.Drawing.Bitmap OriginalColor { get; init; }
    public required System.Drawing.Bitmap CroppedColor { get; init; }
    public required System.Drawing.Bitmap Gray { get; init; }
    public required System.Drawing.Bitmap Binary { get; init; }
    public required System.Drawing.Bitmap ColorForGemini { get; init; }
    public required System.Drawing.Bitmap BinaryForTable { get; init; }
    public required CropInfo CropInfo { get; init; }
    public required (int W, int H) OriginalSizePx { get; init; }
    public required (int W, int H) CroppedSizePx { get; init; }
}
