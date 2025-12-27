using Pdf2Word.Core.Services;
using Xunit;

namespace Pdf2Word.Tests.Tests;

public class PageRangeParserTests
{
    [Fact]
    public void EmptyInputReturnsAllPages()
    {
        var result = PageRangeParser.Parse("", 5);
        Assert.False(result.HasError);
        Assert.Equal(new[] { 1, 2, 3, 4, 5 }, result.Pages);
    }

    [Fact]
    public void ParsesRangesAndLists()
    {
        var result = PageRangeParser.Parse("1,3,5-7", 10);
        Assert.False(result.HasError);
        Assert.Equal(new[] { 1, 3, 5, 6, 7 }, result.Pages);
    }

    [Fact]
    public void ReversedRangeIsCorrected()
    {
        var result = PageRangeParser.Parse("10-8", 12);
        Assert.False(result.HasError);
        Assert.Equal(new[] { 8, 9, 10 }, result.Pages);
    }

    [Fact]
    public void InvalidTokenReturnsError()
    {
        var result = PageRangeParser.Parse("1,a,3", 10);
        Assert.True(result.HasError);
    }

    [Fact]
    public void OutOfRangePagesAreIgnoredWithWarning()
    {
        var result = PageRangeParser.Parse("0,1,11", 10);
        Assert.False(result.HasError);
        Assert.Equal(new[] { 1 }, result.Pages);
        Assert.NotEmpty(result.Warnings);
    }
}
