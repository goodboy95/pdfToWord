namespace Pdf2Word.Core.Models;

public readonly record struct BBox(int X, int Y, int W, int H)
{
    public int Left => X;
    public int Top => Y;
    public int Right => X + W;
    public int Bottom => Y + H;

    public bool Intersects(BBox other)
    {
        return !(Right <= other.Left || other.Right <= Left || Bottom <= other.Top || other.Bottom <= Top);
    }

    public static BBox FromBounds(int left, int top, int right, int bottom)
    {
        return new BBox(left, top, Math.Max(0, right - left), Math.Max(0, bottom - top));
    }
}
