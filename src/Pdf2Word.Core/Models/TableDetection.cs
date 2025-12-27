namespace Pdf2Word.Core.Models;

public sealed class TableDetection
{
    public required BBox TableBBoxInPage { get; init; }
    public required System.Drawing.Bitmap TableImageColor { get; init; }
    public required System.Drawing.Bitmap TableBinary { get; init; }
    public List<int> ColLinesX { get; init; } = new();
    public List<int> RowLinesY { get; init; } = new();
    public List<CellBox> Cells { get; init; } = new();
    public int NCols { get; init; }
    public int NRows { get; init; }
    public string? DebugId { get; init; }
}

public sealed class CellBox
{
    public string Id { get; init; } = string.Empty;
    public BBox BBoxInTable { get; init; }
    public int Row { get; init; }
    public int Col { get; init; }
    public int Rowspan { get; init; } = 1;
    public int Colspan { get; init; } = 1;
}
