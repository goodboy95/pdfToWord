using Pdf2Word.Core.Services;
using Xunit;

namespace Pdf2Word.Tests.Tests;

public class TwipsConverterTests
{
    [Fact]
    public void ConvertsPixelsToTwips()
    {
        var twips = TwipsConverter.PixelsToTwips(1440, 144);
        Assert.Equal(14400, twips);
    }
}
