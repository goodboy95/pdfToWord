using System.Text.Json.Serialization;
using Pdf2Word.Core.Models;

namespace Pdf2Word.Core.Models.Ir;

public sealed class DocumentIr
{
    [JsonPropertyName("version")]
    public string Version { get; set; } = "1.0";

    [JsonPropertyName("meta")]
    public DocumentMeta Meta { get; set; } = new();

    [JsonPropertyName("pages")]
    public List<PageIr> Pages { get; set; } = new();
}

public sealed class DocumentMeta
{
    [JsonPropertyName("sourcePdfPath")]
    public string SourcePdfPath { get; set; } = string.Empty;

    [JsonPropertyName("generatedAtUtc")]
    public DateTime GeneratedAtUtc { get; set; } = DateTime.UtcNow;

    [JsonPropertyName("options")]
    public ConvertOptionsSnapshot Options { get; set; } = new();
}

public sealed class ConvertOptionsSnapshot
{
    [JsonPropertyName("dpi")]
    public int Dpi { get; set; } = 300;

    [JsonPropertyName("pageRange")]
    public string PageRange { get; set; } = string.Empty;

    [JsonPropertyName("headerFooterMode")]
    [JsonConverter(typeof(JsonStringEnumConverter))]
    public HeaderFooterRemoveMode HeaderFooterMode { get; set; } = HeaderFooterRemoveMode.None;

    [JsonPropertyName("headerPercent")]
    public double HeaderPercent { get; set; } = 0.06;

    [JsonPropertyName("footerPercent")]
    public double FooterPercent { get; set; } = 0.06;

    [JsonPropertyName("pageSizeMode")]
    [JsonConverter(typeof(JsonStringEnumConverter))]
    public PageSizeMode PageSizeMode { get; set; } = PageSizeMode.FollowPdf;
}
