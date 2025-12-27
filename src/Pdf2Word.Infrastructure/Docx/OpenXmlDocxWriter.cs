using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Pdf2Word.Core.Models;
using Pdf2Word.Core.Models.Ir;
using Pdf2Word.Core.Options;
using Pdf2Word.Core.Services;

namespace Pdf2Word.Infrastructure.Docx;

public sealed class OpenXmlDocxWriter : IDocxWriter
{
    public Task WriteAsync(DocumentIr doc, DocxWriteOptions options, Stream output, CancellationToken ct)
    {
        using var wordDoc = WordprocessingDocument.Create(output, WordprocessingDocumentType.Document);
        var mainPart = wordDoc.AddMainDocumentPart();
        mainPart.Document = new Document(new Body());
        var body = mainPart.Document.Body;

        for (var i = 0; i < doc.Pages.Count; i++)
        {
            ct.ThrowIfCancellationRequested();
            var page = doc.Pages[i];
            var pageSize = ResolvePageSize(page, options);
            var pageMargins = new PageMargin
            {
                Top = options.MarginTopTwips,
                Bottom = options.MarginBottomTwips,
                Left = options.MarginLeftTwips,
                Right = options.MarginRightTwips
            };

            var context = new PageContext(pageSize.Width, pageSize.Height, pageMargins);

            foreach (var block in page.Blocks)
            {
                switch (block)
                {
                    case ParagraphBlockIr paragraph:
                        body.AppendChild(BuildParagraph(paragraph, options));
                        break;
                    case TableBlockIr table:
                        body.AppendChild(BuildTable(table, context, options));
                        break;
                }
            }

            var isLast = i == doc.Pages.Count - 1;
            if (options.PageSizeMode == PageSizeMode.FollowPdf)
            {
                if (!isLast)
                {
                    body.AppendChild(BuildSectionBreak(pageSize, pageMargins));
                }
                else
                {
                    body.AppendChild(BuildSectionProperties(pageSize, pageMargins));
                }
            }
            else if (!isLast)
            {
                body.AppendChild(BuildPageBreakParagraph());
            }
        }

        if (options.PageSizeMode == PageSizeMode.A4)
        {
            body.AppendChild(BuildSectionProperties(GetA4Size(), new PageMargin
            {
                Top = options.MarginTopTwips,
                Bottom = options.MarginBottomTwips,
                Left = options.MarginLeftTwips,
                Right = options.MarginRightTwips
            }));
        }

        mainPart.Document.Save();
        return Task.CompletedTask;
    }

    private static Paragraph BuildParagraph(ParagraphBlockIr paragraph, DocxWriteOptions options)
    {
        var p = new Paragraph();
        var pPr = new ParagraphProperties();
        if (paragraph.Role == Pdf2Word.Core.Models.ParagraphRole.Title)
        {
            pPr.ParagraphStyleId = new ParagraphStyleId { Val = "Heading1" };
        }
        p.Append(pPr);

        var runs = BuildRuns(paragraph.Text, options);
        foreach (var run in runs)
        {
            p.Append(run);
        }

        return p;
    }

    private static IEnumerable<Run> BuildRuns(string text, DocxWriteOptions options)
    {
        var parts = text.Split('\n');
        for (var i = 0; i < parts.Length; i++)
        {
            var run = new Run();
            run.RunProperties = BuildRunProperties(options);
            run.AppendChild(new Text(parts[i]) { Space = SpaceProcessingModeValues.Preserve });
            if (i < parts.Length - 1)
            {
                run.AppendChild(new Break());
            }
            yield return run;
        }
    }

    private static RunProperties BuildRunProperties(DocxWriteOptions options)
    {
        return new RunProperties(
            new RunFonts { Ascii = options.FontAscii, EastAsia = options.FontEastAsia, HighAnsi = options.FontAscii },
            new FontSize { Val = options.DefaultFontSizeHalfPoints.ToString() });
    }

    private static Table BuildTable(TableBlockIr table, PageContext context, DocxWriteOptions options)
    {
        var tbl = new Table();
        var tblPr = new TableProperties();
        if (options.Table.SetBorders)
        {
            tblPr.TableBorders = new TableBorders(
                new TopBorder { Val = BorderValues.Single, Size = 4 },
                new BottomBorder { Val = BorderValues.Single, Size = 4 },
                new LeftBorder { Val = BorderValues.Single, Size = 4 },
                new RightBorder { Val = BorderValues.Single, Size = 4 },
                new InsideHorizontalBorder { Val = BorderValues.Single, Size = 4 },
                new InsideVerticalBorder { Val = BorderValues.Single, Size = 4 }
            );
        }

        tbl.AppendChild(tblPr);
        var grid = new TableGrid();
        var left = context.Margins.Left?.Value ?? 0;
        var right = context.Margins.Right?.Value ?? 0;
        var availableWidth = Math.Max(1, context.PageWidthTwips - left - right);
        var colWidth = table.NCols > 0 ? availableWidth / table.NCols : availableWidth;
        for (var i = 0; i < table.NCols; i++)
        {
            grid.AppendChild(new GridColumn { Width = colWidth.ToString() });
        }
        tbl.AppendChild(grid);

        var (owner, cellMap) = BuildOwnerMatrix(table);
        for (var r = 0; r < table.Rows.Count; r++)
        {
            var row = new TableRow();
            for (var c = 0; c < table.NCols; c++)
            {
                var cell = new TableCell();
                var cellProps = new TableCellProperties();
                if (owner[r, c] is { } ownerPos)
                {
                    if (ownerPos.r == r && ownerPos.c == c)
                    {
                        if (cellMap.TryGetValue((r, c), out var actualCell))
                        {
                            if (actualCell.Colspan > 1)
                            {
                                cellProps.GridSpan = new GridSpan { Val = actualCell.Colspan };
                            }
                            if (actualCell.Rowspan > 1)
                            {
                                cellProps.VerticalMerge = new VerticalMerge { Val = MergedCellValues.Restart };
                            }
                            cell.AppendChild(cellProps);
                            cell.AppendChild(BuildCellParagraph(actualCell.Text, options));
                        }
                        else
                        {
                            cell.AppendChild(cellProps);
                            cell.AppendChild(new Paragraph());
                        }
                    }
                    else
                    {
                        if (ownerPos.r < r)
                        {
                            cellProps.VerticalMerge = new VerticalMerge { Val = MergedCellValues.Continue };
                        }
                        cell.AppendChild(cellProps);
                        cell.AppendChild(new Paragraph());
                    }
                }
                else
                {
                    cell.AppendChild(new Paragraph());
                }

                row.AppendChild(cell);
            }
            tbl.AppendChild(row);
        }

        return tbl;
    }

    private static Paragraph BuildCellParagraph(string text, DocxWriteOptions options)
    {
        var paragraph = new Paragraph();
        foreach (var run in BuildRuns(text ?? string.Empty, options))
        {
            paragraph.Append(run);
        }
        return paragraph;
    }

    private static ((int r, int c)?[,] owner, Dictionary<(int r, int c), TableCellIr> cellMap) BuildOwnerMatrix(TableBlockIr table)
    {
        var rows = table.Rows.Count;
        var cols = table.NCols;
        var owner = new (int r, int c)?[rows, cols];
        var cellMap = new Dictionary<(int r, int c), TableCellIr>();
        for (var r = 0; r < rows; r++)
        {
            var colCursor = 0;
            foreach (var cell in table.Rows[r].Cells)
            {
                while (colCursor < cols && owner[r, colCursor].HasValue)
                {
                    colCursor++;
                }

                if (colCursor >= cols)
                {
                    break;
                }

                var rowspan = Math.Max(1, cell.Rowspan);
                var colspan = Math.Max(1, cell.Colspan);
                cellMap[(r, colCursor)] = cell;
                for (var rr = r; rr < Math.Min(rows, r + rowspan); rr++)
                {
                    for (var cc = colCursor; cc < Math.Min(cols, colCursor + colspan); cc++)
                    {
                        owner[rr, cc] = (r, colCursor);
                    }
                }

                colCursor += colspan;
            }
        }

        return (owner, cellMap);
    }

    private static Paragraph BuildPageBreakParagraph()
    {
        var paragraph = new Paragraph();
        paragraph.AppendChild(new Run(new Break { Type = BreakValues.Page }));
        return paragraph;
    }

    private static Paragraph BuildSectionBreak(PageSize size, PageMargin margin)
    {
        var paragraph = new Paragraph();
        var section = BuildSectionProperties(size, margin, true);
        paragraph.AppendChild(new ParagraphProperties(section));
        return paragraph;
    }

    private static SectionProperties BuildSectionProperties(PageSize size, PageMargin margin, bool nextPage = false)
    {
        var properties = new SectionProperties();
        if (nextPage)
        {
            properties.AppendChild(new SectionType { Val = SectionMarkValues.NextPage });
        }
        properties.AppendChild(size);
        properties.AppendChild(margin);
        return properties;
    }

    private static PageSize ResolvePageSize(PageIr page, DocxWriteOptions options)
    {
        if (options.PageSizeMode == PageSizeMode.A4)
        {
            return GetA4Size();
        }

        var widthTwips = TwipsConverter.PixelsToTwips(page.WidthPx, options.Dpi);
        var heightTwips = TwipsConverter.PixelsToTwips(page.HeightPx, options.Dpi);
        var orient = page.WidthPx > page.HeightPx ? PageOrientationValues.Landscape : PageOrientationValues.Portrait;
        return new PageSize { Width = (UInt32Value)(uint)widthTwips, Height = (UInt32Value)(uint)heightTwips, Orient = orient };
    }

    private static PageSize GetA4Size()
    {
        return new PageSize { Width = (UInt32Value)11906U, Height = (UInt32Value)16838U };
    }

    private readonly record struct PageContext(int PageWidthTwips, int PageHeightTwips, PageMargin Margins);
}
