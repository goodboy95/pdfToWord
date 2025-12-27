using System.Windows.Media;

namespace Pdf2Word.App.Models;

public sealed class PreviewOverlay
{
    public double X { get; set; }
    public double Y { get; set; }
    public double Width { get; set; }
    public double Height { get; set; }
    public Brush Stroke { get; set; } = Brushes.Transparent;
    public double StrokeThickness { get; set; } = 1;
    public Brush Fill { get; set; } = Brushes.Transparent;
}
