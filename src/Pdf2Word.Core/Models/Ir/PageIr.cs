using System.Text.Json.Serialization;
using Pdf2Word.Core.Models;

namespace Pdf2Word.Core.Models.Ir;

public sealed class PageIr
{
    [JsonPropertyName("pageNumber")]
    public int PageNumber { get; set; }

    [JsonPropertyName("originalWidthPx")]
    public int OriginalWidthPx { get; set; }

    [JsonPropertyName("originalHeightPx")]
    public int OriginalHeightPx { get; set; }

    [JsonPropertyName("widthPx")]
    public int WidthPx { get; set; }

    [JsonPropertyName("heightPx")]
    public int HeightPx { get; set; }

    [JsonPropertyName("crop")]
    public CropInfo Crop { get; set; } = new();

    [JsonPropertyName("blocks")]
    public List<BlockIr> Blocks { get; set; } = new();
}

public sealed class CropInfo
{
    [JsonPropertyName("mode")]
    [JsonConverter(typeof(JsonStringEnumConverter))]
    public HeaderFooterRemoveMode Mode { get; set; } = HeaderFooterRemoveMode.None;

    [JsonPropertyName("cropTopPx")]
    public int CropTopPx { get; set; }

    [JsonPropertyName("cropBottomPx")]
    public int CropBottomPx { get; set; }
}

[JsonDerivedType(typeof(ParagraphBlockIr), "paragraph")]
[JsonDerivedType(typeof(TableBlockIr), "table")]
public abstract class BlockIr
{
    [JsonPropertyName("type")]
    public abstract string Type { get; }

    [JsonPropertyName("bbox")]
    public BBox? BBox { get; set; }

    [JsonPropertyName("source")]
    public BlockSource Source { get; set; } = new();
}

public sealed class BlockSource
{
    [JsonPropertyName("producer")]
    public string Producer { get; set; } = string.Empty;

    [JsonPropertyName("attempt")]
    public int Attempt { get; set; } = 1;

    [JsonPropertyName("debugId")]
    public string? DebugId { get; set; }
}

public sealed class ParagraphBlockIr : BlockIr
{
    public override string Type => "paragraph";

    [JsonPropertyName("role")]
    [JsonConverter(typeof(JsonStringEnumConverter))]
    public ParagraphRole Role { get; set; } = ParagraphRole.Body;

    [JsonPropertyName("text")]
    public string Text { get; set; } = string.Empty;
}

public sealed class TableBlockIr : BlockIr
{
    public override string Type => "table";

    [JsonPropertyName("tableBBox")]
    public BBox TableBBox { get; set; }

    [JsonPropertyName("nCols")]
    public int NCols { get; set; }

    [JsonPropertyName("rows")]
    public List<TableRowIr> Rows { get; set; } = new();

    [JsonPropertyName("structureMeta")]
    public TableStructureMeta StructureMeta { get; set; } = new();
}

public sealed class TableStructureMeta
{
    [JsonPropertyName("engine")]
    public string Engine { get; set; } = "OpenCV";

    [JsonPropertyName("lineScore")]
    public double? LineScore { get; set; }

    [JsonPropertyName("detectedCellCount")]
    public int? DetectedCellCount { get; set; }
}

public sealed class TableRowIr
{
    [JsonPropertyName("cells")]
    public List<TableCellIr> Cells { get; set; } = new();
}

public sealed class TableCellIr
{
    [JsonPropertyName("text")]
    public string Text { get; set; } = string.Empty;

    [JsonPropertyName("rowspan")]
    public int Rowspan { get; set; } = 1;

    [JsonPropertyName("colspan")]
    public int Colspan { get; set; } = 1;

    [JsonPropertyName("cellBBoxInTableImage")]
    public BBox? CellBBoxInTableImage { get; set; }

    [JsonPropertyName("cellId")]
    public string? CellId { get; set; }
}
