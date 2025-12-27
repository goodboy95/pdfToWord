using Pdf2Word.Core.Services;
using Xunit;

namespace Pdf2Word.Tests.Tests;

public class TextCleanerTests
{
    [Fact]
    public void RemovesControlCharacters()
    {
        var input = "Hello\u0001World\nNext";
        var cleaned = TextCleaner.RemoveControlChars(input);
        Assert.Equal("HelloWorld\nNext", cleaned);
    }

    [Fact]
    public void CalculatesReplacementCharRate()
    {
        var input = "ab\uFFFDcd\uFFFD";
        var rate = TextCleaner.ReplacementCharRate(input);
        Assert.Equal(2d / input.Length, rate, 5);
    }
}
