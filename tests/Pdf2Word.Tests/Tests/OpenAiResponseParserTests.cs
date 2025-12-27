using Pdf2Word.Core.Services;
using Xunit;

namespace Pdf2Word.Tests.Tests;

public class OpenAiResponseParserTests
{
    [Fact]
    public void ExtractText_Returns_Message_Content_String()
    {
        var json = """
        {"choices":[{"message":{"content":"hello"}}]}
        """;

        var text = OpenAiResponseParser.ExtractText(json);

        Assert.Equal("hello", text);
    }

    [Fact]
    public void ExtractText_Combines_Message_Content_Array()
    {
        var json = """
        {"choices":[{"message":{"content":[{"type":"text","text":"hello "},{"type":"text","text":"world"}]}}]}
        """;

        var text = OpenAiResponseParser.ExtractText(json);

        Assert.Equal("hello world", text);
    }

    [Fact]
    public void ExtractText_Uses_Choice_Text_Fallback()
    {
        var json = """
        {"choices":[{"text":"fallback"}]}
        """;

        var text = OpenAiResponseParser.ExtractText(json);

        Assert.Equal("fallback", text);
    }

    [Fact]
    public void ExtractText_Uses_Root_Text_Fallback()
    {
        var json = """
        {"text":"root"}
        """;

        var text = OpenAiResponseParser.ExtractText(json);

        Assert.Equal("root", text);
    }
}
