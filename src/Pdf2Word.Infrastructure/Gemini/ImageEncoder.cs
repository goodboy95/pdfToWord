using System.Drawing;
using System.Drawing.Imaging;

namespace Pdf2Word.Infrastructure.Gemini;

public static class ImageEncoder
{
    public static byte[] Encode(Bitmap bitmap, bool usePng, int maxLongEdgePx, int jpegQuality)
    {
        using var scaled = Scale(bitmap, maxLongEdgePx);
        using var ms = new MemoryStream();
        if (usePng)
        {
            scaled.Save(ms, ImageFormat.Png);
        }
        else
        {
            var encoder = ImageCodecInfo.GetImageEncoders().First(e => e.FormatID == ImageFormat.Jpeg.Guid);
            using var encoderParams = new EncoderParameters(1);
            encoderParams.Param[0] = new EncoderParameter(Encoder.Quality, Math.Clamp(jpegQuality, 60, 95));
            scaled.Save(ms, encoder, encoderParams);
        }

        return ms.ToArray();
    }

    private static Bitmap Scale(Bitmap source, int maxLongEdgePx)
    {
        if (maxLongEdgePx <= 0)
        {
            return (Bitmap)source.Clone();
        }

        var longEdge = Math.Max(source.Width, source.Height);
        if (longEdge <= maxLongEdgePx)
        {
            return (Bitmap)source.Clone();
        }

        var scale = (double)maxLongEdgePx / longEdge;
        var width = Math.Max(1, (int)Math.Round(source.Width * scale));
        var height = Math.Max(1, (int)Math.Round(source.Height * scale));
        var result = new Bitmap(width, height, source.PixelFormat);
        using var graphics = Graphics.FromImage(result);
        graphics.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
        graphics.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.HighQuality;
        graphics.DrawImage(source, 0, 0, width, height);
        return result;
    }
}
