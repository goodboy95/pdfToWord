using Pdf2Word.Core.Models;
using Pdf2Word.Core.Models.Ir;

namespace Pdf2Word.Core.Services;

public static class IrBuilder
{
    public static PageIr BuildPage(int pageNumber,
        (int W, int H) originalSize,
        (int W, int H) croppedSize,
        CropInfo cropInfo,
        IReadOnlyList<ParagraphDto> paragraphs,
        IReadOnlyList<TableBlockIr> tables,
        string producer)
    {
        var pageIr = new PageIr
        {
            PageNumber = pageNumber,
            OriginalWidthPx = originalSize.W,
            OriginalHeightPx = originalSize.H,
            WidthPx = croppedSize.W,
            HeightPx = croppedSize.H,
            Crop = cropInfo
        };

        var paragraphBlocks = BuildParagraphBlocks(paragraphs, pageIr.WidthPx, pageIr.HeightPx, tables);
        foreach (var block in paragraphBlocks)
        {
            block.Source.Producer = producer;
        }

        foreach (var table in tables)
        {
            table.Source.Producer = "OpenCvTable+GeminiCells";
        }

        var allBlocks = new List<BlockIr>();
        allBlocks.AddRange(paragraphBlocks);
        allBlocks.AddRange(tables);

        pageIr.Blocks = SortBlocks(allBlocks, pageIr.WidthPx, pageIr.HeightPx);
        return pageIr;
    }

    private static List<ParagraphBlockIr> BuildParagraphBlocks(IReadOnlyList<ParagraphDto> paragraphs, int pageWidth, int pageHeight, IReadOnlyList<TableBlockIr> tables)
    {
        var blocks = new List<ParagraphBlockIr>();
        if (paragraphs.Count == 0)
        {
            return blocks;
        }

        var segments = BuildVerticalSegments(pageHeight, tables.Select(t => t.TableBBox).ToList());
        var totalSegments = segments.Count;
        var distribution = DistributeParagraphs(paragraphs.Count, segments.Select(s => s.Height).ToList());

        var index = 0;
        for (var s = 0; s < totalSegments; s++)
        {
            var count = distribution[s];
            if (count <= 0)
            {
                continue;
            }

            var segment = segments[s];
            var segmentSliceHeight = Math.Max(1, segment.Height / count);
            for (var i = 0; i < count && index < paragraphs.Count; i++)
            {
                var dto = paragraphs[index++];
                blocks.Add(new ParagraphBlockIr
                {
                    Role = dto.Role.Equals("title", StringComparison.OrdinalIgnoreCase) ? ParagraphRole.Title : ParagraphRole.Body,
                    Text = dto.Text,
                    BBox = new BBox(0, segment.Top + i * segmentSliceHeight, pageWidth, Math.Max(1, segmentSliceHeight))
                });
            }
        }

        while (index < paragraphs.Count)
        {
            var dto = paragraphs[index++];
            blocks.Add(new ParagraphBlockIr
            {
                Role = dto.Role.Equals("title", StringComparison.OrdinalIgnoreCase) ? ParagraphRole.Title : ParagraphRole.Body,
                Text = dto.Text,
                BBox = new BBox(0, pageHeight - 1, pageWidth, 1)
            });
        }

        return blocks;
    }

    private static List<BlockIr> SortBlocks(List<BlockIr> blocks, int pageWidth, int pageHeight)
    {
        foreach (var block in blocks)
        {
            if (!block.BBox.HasValue)
            {
                block.BBox = new BBox(0, 0, pageWidth, pageHeight);
            }
        }

        return blocks.OrderBy(b => b.BBox!.Value.Top)
            .ThenBy(b => b.BBox!.Value.Left)
            .ThenBy(b => b.Type)
            .ToList();
    }

    private static List<Segment> BuildVerticalSegments(int pageHeight, IReadOnlyList<BBox> tableBoxes)
    {
        var segments = new List<Segment>();
        if (tableBoxes.Count == 0)
        {
            segments.Add(new Segment(0, pageHeight));
            return segments;
        }

        var ordered = tableBoxes.OrderBy(b => b.Top).ToList();
        var currentTop = 0;
        foreach (var box in ordered)
        {
            if (box.Top > currentTop)
            {
                segments.Add(new Segment(currentTop, box.Top));
            }
            currentTop = Math.Max(currentTop, box.Bottom);
        }

        if (currentTop < pageHeight)
        {
            segments.Add(new Segment(currentTop, pageHeight));
        }

        if (segments.Count == 0)
        {
            segments.Add(new Segment(0, pageHeight));
        }

        return segments;
    }

    private static List<int> DistributeParagraphs(int totalParagraphs, List<int> segmentHeights)
    {
        var distribution = Enumerable.Repeat(0, segmentHeights.Count).ToList();
        if (segmentHeights.Count == 0 || totalParagraphs == 0)
        {
            return distribution;
        }

        var totalHeight = segmentHeights.Sum();
        if (totalHeight <= 0)
        {
            distribution[0] = totalParagraphs;
            return distribution;
        }

        var assigned = 0;
        for (var i = 0; i < segmentHeights.Count; i++)
        {
            var weight = (double)segmentHeights[i] / totalHeight;
            var count = (int)Math.Round(weight * totalParagraphs);
            distribution[i] = count;
            assigned += count;
        }

        var remaining = totalParagraphs - assigned;
        var idx = 0;
        while (remaining != 0)
        {
            distribution[idx % distribution.Count] += remaining > 0 ? 1 : -1;
            remaining += remaining > 0 ? -1 : 1;
            idx++;
        }

        return distribution;
    }

    private readonly record struct Segment(int Top, int Bottom)
    {
        public int Height => Math.Max(1, Bottom - Top);
    }
}
