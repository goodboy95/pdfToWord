namespace Pdf2Word.Core.Services;

public static class TwipsConverter
{
    public static int PixelsToTwips(int pixels, int dpi)
    {
        if (dpi <= 0)
        {
            throw new ArgumentOutOfRangeException(nameof(dpi));
        }

        return (int)Math.Round(pixels * 1440.0 / dpi);
    }
}
