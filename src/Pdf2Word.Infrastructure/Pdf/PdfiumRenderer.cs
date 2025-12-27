using Pdf2Word.Core.Services;
using PdfiumViewer;

namespace Pdf2Word.Infrastructure.Pdf;

public sealed class PdfiumRenderer : IPdfRenderer
{
    public int GetPageCount(string pdfPath)
    {
        using var document = PdfDocument.Load(pdfPath);
        return document.PageCount;
    }

    public System.Drawing.Bitmap RenderPage(string pdfPath, int pageIndex0Based, int dpi)
    {
        using var document = PdfDocument.Load(pdfPath);
        var size = document.PageSizes[pageIndex0Based];
        var width = (int)Math.Round(size.Width / 72.0 * dpi);
        var height = (int)Math.Round(size.Height / 72.0 * dpi);
        var bitmap = new System.Drawing.Bitmap(width, height, System.Drawing.Imaging.PixelFormat.Format24bppRgb);
        using var gfx = System.Drawing.Graphics.FromImage(bitmap);
        gfx.Clear(System.Drawing.Color.White);
        document.Render(pageIndex0Based, gfx, 0, 0, width, height, PdfRenderFlags.Annotations);
        return bitmap;
    }
}
