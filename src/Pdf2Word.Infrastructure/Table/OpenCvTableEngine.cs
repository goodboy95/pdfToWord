using OpenCvSharp;
using OpenCvSharp.Extensions;
using Pdf2Word.Core.Models;
using Pdf2Word.Core.Options;
using Pdf2Word.Core.Services;

namespace Pdf2Word.Infrastructure.Table;

public sealed class OpenCvTableEngine : ITableEngine
{
    public IReadOnlyList<TableDetection> DetectTables(PageImageBundle bundle, TableDetectOptions options)
    {
        using var binaryMat = BitmapConverter.ToMat(bundle.BinaryForTable);
        using var binary = EnsureBinaryForeground(binaryMat);

        using var horizontal = ExtractLines(binary, true, options);
        using var vertical = ExtractLines(binary, false, options);
        using var gridMask = new Mat();
        Cv2.BitwiseOr(horizontal, vertical, gridMask);

        var bboxes = DetectTableBBoxes(gridMask, bundle.CroppedSizePx, options);
        var tables = new List<TableDetection>();

        var tableIndex = 0;
        foreach (var bbox in bboxes)
        {
            var table = BuildTable(bundle, binary, gridMask, horizontal, vertical, bbox, options, tableIndex, bundle.PageNumber);
            if (table != null)
            {
                tables.Add(table);
            }
            tableIndex++;
        }

        return tables;
    }

    public System.Drawing.Bitmap MaskTables(System.Drawing.Bitmap source, IEnumerable<BBox> tableBoxes)
    {
        var mat = BitmapConverter.ToMat(source);
        foreach (var box in tableBoxes)
        {
            var rect = new Rect(box.X, box.Y, box.W, box.H);
            Cv2.Rectangle(mat, rect, Scalar.White, -1);
        }

        return mat.ToBitmap();
    }

    private static Mat EnsureBinaryForeground(Mat binary)
    {
        var mean = Cv2.Mean(binary).Val0;
        if (mean < 128)
        {
            var inverted = new Mat();
            Cv2.BitwiseNot(binary, inverted);
            return inverted;
        }

        return binary.Clone();
    }

    private static Mat ExtractLines(Mat binary, bool horizontal, TableDetectOptions options)
    {
        var size = horizontal ? new Size(Math.Max(options.MinKernelPx, binary.Width / options.KernelDivisor), 1)
            : new Size(1, Math.Max(options.MinKernelPx, binary.Height / options.KernelDivisor));
        using var kernel = Cv2.GetStructuringElement(MorphShapes.Rect, size);
        var extracted = new Mat();
        Cv2.Erode(binary, extracted, kernel);
        Cv2.Dilate(extracted, extracted, kernel);
        if (options.DilateKernelPx > 0)
        {
            using var dilateKernel = Cv2.GetStructuringElement(MorphShapes.Rect, new Size(options.DilateKernelPx, options.DilateKernelPx));
            Cv2.Dilate(extracted, extracted, dilateKernel);
        }

        return extracted;
    }

    private static List<BBox> DetectTableBBoxes(Mat gridMask, (int W, int H) pageSize, TableDetectOptions options)
    {
        using var dilated = new Mat();
        using var kernel = Cv2.GetStructuringElement(MorphShapes.Rect, new Size(3, 3));
        Cv2.Dilate(gridMask, dilated, kernel);

        Cv2.FindContours(dilated, out var contours, out _, RetrievalModes.External, ContourApproximationModes.ApproxSimple);
        var candidates = new List<BBox>();
        var pageArea = pageSize.W * pageSize.H;
        foreach (var contour in contours)
        {
            var rect = Cv2.BoundingRect(contour);
            var area = rect.Width * rect.Height;
            if (area < pageArea * options.MinAreaRatio)
            {
                continue;
            }

            if (rect.Width < options.MinWidthPx || rect.Height < options.MinHeightPx)
            {
                continue;
            }

            candidates.Add(new BBox(rect.X, rect.Y, rect.Width, rect.Height));
        }

        return MergeBBoxes(candidates, options.MergeGapPx);
    }

    private static List<BBox> MergeBBoxes(List<BBox> boxes, int gap)
    {
        var merged = new List<BBox>(boxes);
        var changed = true;
        while (changed)
        {
            changed = false;
            for (var i = 0; i < merged.Count; i++)
            {
                for (var j = i + 1; j < merged.Count; j++)
                {
                    if (IsClose(merged[i], merged[j], gap))
                    {
                        var combined = Combine(merged[i], merged[j]);
                        merged[i] = combined;
                        merged.RemoveAt(j);
                        changed = true;
                        break;
                    }
                }
                if (changed)
                {
                    break;
                }
            }
        }

        return merged.OrderBy(b => b.Top).ToList();
    }

    private static bool IsClose(BBox a, BBox b, int gap)
    {
        var horizontalGap = Math.Max(0, Math.Max(a.Left - b.Right, b.Left - a.Right));
        var verticalGap = Math.Max(0, Math.Max(a.Top - b.Bottom, b.Top - a.Bottom));
        return horizontalGap <= gap && verticalGap <= gap;
    }

    private static BBox Combine(BBox a, BBox b)
    {
        var left = Math.Min(a.Left, b.Left);
        var top = Math.Min(a.Top, b.Top);
        var right = Math.Max(a.Right, b.Right);
        var bottom = Math.Max(a.Bottom, b.Bottom);
        return BBox.FromBounds(left, top, right, bottom);
    }

    private static TableDetection? BuildTable(PageImageBundle bundle, Mat binaryPage, Mat gridMask, Mat horizontalMask, Mat verticalMask, BBox bbox, TableDetectOptions options, int tableIndex, int pageNumber)
    {
        var rect = new Rect(bbox.X, bbox.Y, bbox.W, bbox.H);
        using var tableBinary = new Mat(binaryPage, rect);
        using var tableGrid = new Mat(gridMask, rect);
        using var tableHorizontal = new Mat(horizontalMask, rect);
        using var tableVertical = new Mat(verticalMask, rect);
        using var tableColor = BitmapConverter.ToMat(bundle.ColorForGemini);
        using var tableColorCrop = new Mat(tableColor, rect);

        var colLines = ExtractLineCoordinates(tableVertical, true, options);
        var rowLines = ExtractLineCoordinates(tableHorizontal, false, options);
        if (colLines.Count < 2 || rowLines.Count < 2)
        {
            return null;
        }

        var cells = DetectCells(tableGrid, colLines, rowLines, options, tableIndex, pageNumber);
        var detection = new TableDetection
        {
            TableBBoxInPage = bbox,
            TableImageColor = tableColorCrop.ToBitmap(),
            TableBinary = tableBinary.ToBitmap(),
            ColLinesX = colLines,
            RowLinesY = rowLines,
            Cells = cells,
            NCols = colLines.Count - 1,
            NRows = rowLines.Count - 1,
            DebugId = $"t{tableIndex:00}"
        };

        return detection;
    }

    private static List<int> ExtractLineCoordinates(Mat tableGrid, bool vertical, TableDetectOptions options)
    {
        var projection = new double[vertical ? tableGrid.Width : tableGrid.Height];
        for (var i = 0; i < projection.Length; i++)
        {
            projection[i] = vertical ? Cv2.Sum(tableGrid.Col(i)).Val0 : Cv2.Sum(tableGrid.Row(i)).Val0;
        }

        var length = vertical ? tableGrid.Height : tableGrid.Width;
        var threshold = length * 0.3 * 255;
        var indices = new List<int>();
        for (var i = 0; i < projection.Length; i++)
        {
            if (projection[i] >= threshold)
            {
                indices.Add(i);
            }
        }

        var clustered = ClusterIndices(indices, options.ClusterEpsPx);
        if (vertical)
        {
            clustered.Add(0);
            clustered.Add(tableGrid.Width - 1);
        }
        else
        {
            clustered.Add(0);
            clustered.Add(tableGrid.Height - 1);
        }

        return clustered.Distinct().OrderBy(i => i).ToList();
    }

    private static List<int> ClusterIndices(List<int> indices, int eps)
    {
        var clusters = new List<int>();
        if (indices.Count == 0)
        {
            return clusters;
        }

        indices.Sort();
        var start = indices[0];
        var sum = start;
        var count = 1;
        for (var i = 1; i < indices.Count; i++)
        {
            if (indices[i] - indices[i - 1] <= eps)
            {
                sum += indices[i];
                count++;
            }
            else
            {
                clusters.Add(sum / count);
                start = indices[i];
                sum = start;
                count = 1;
            }
        }

        clusters.Add(sum / count);
        return clusters;
    }

    private static List<CellBox> DetectCells(Mat tableGrid, List<int> colLines, List<int> rowLines, TableDetectOptions options, int tableIndex, int pageNumber)
    {
        using var dilated = new Mat();
        using var kernel = Cv2.GetStructuringElement(MorphShapes.Rect, new Size(3, 3));
        Cv2.Dilate(tableGrid, dilated, kernel);
        using var inverted = new Mat();
        Cv2.BitwiseNot(dilated, inverted);

        Cv2.FindContours(inverted, out var contours, out _, RetrievalModes.External, ContourApproximationModes.ApproxSimple);
        var cells = new List<CellBox>();
        var minW = tableGrid.Width * options.MinCellSizeRatio;
        var minH = tableGrid.Height * options.MinCellSizeRatio;

        foreach (var contour in contours)
        {
            var rect = Cv2.BoundingRect(contour);
            if (rect.Width < minW || rect.Height < minH)
            {
                continue;
            }

            if (rect.Width > tableGrid.Width * 0.98 && rect.Height > tableGrid.Height * 0.98)
            {
                continue;
            }

            var colRange = MapToLineRange(rect.Left, rect.Right, colLines, options.ClusterEpsPx);
            var rowRange = MapToLineRange(rect.Top, rect.Bottom, rowLines, options.ClusterEpsPx);
            if (colRange == null || rowRange == null)
            {
                continue;
            }

            var (colStart, colEnd) = colRange.Value;
            var (rowStart, rowEnd) = rowRange.Value;
            var cellId = $"p{pageNumber}_t{tableIndex:00}_r{rowStart}_c{colStart}";

            cells.Add(new CellBox
            {
                Id = cellId,
                BBoxInTable = new BBox(rect.X, rect.Y, rect.Width, rect.Height),
                Row = rowStart,
                Col = colStart,
                Rowspan = Math.Max(1, rowEnd - rowStart + 1),
                Colspan = Math.Max(1, colEnd - colStart + 1)
            });
        }

        return cells;
    }

    private static (int start, int end)? MapToLineRange(int start, int end, List<int> lines, int tolerance)
    {
        if (lines.Count < 2)
        {
            return null;
        }

        var startIndex = -1;
        var endIndex = -1;
        for (var i = 0; i < lines.Count - 1; i++)
        {
            var left = lines[i] - tolerance;
            var right = lines[i + 1] + tolerance;
            if (start >= left && start <= right && startIndex == -1)
            {
                startIndex = i;
            }
            if (end >= left && end <= right)
            {
                endIndex = i;
            }
        }

        if (startIndex == -1 || endIndex == -1)
        {
            return null;
        }

        if (endIndex < startIndex)
        {
            (startIndex, endIndex) = (endIndex, startIndex);
        }

        return (startIndex, endIndex);
    }
}
