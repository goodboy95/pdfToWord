using System.Text;
using System.Text.Json;

namespace Pdf2Word.Core.Services;

public static class OpenAiResponseParser
{
    public static string ExtractText(string responseJson)
    {
        using var doc = JsonDocument.Parse(responseJson);
        if (doc.RootElement.TryGetProperty("choices", out var choices)
            && choices.ValueKind == JsonValueKind.Array
            && choices.GetArrayLength() > 0)
        {
            var choice = choices[0];
            if (choice.TryGetProperty("message", out var message) && message.TryGetProperty("content", out var content))
            {
                if (content.ValueKind == JsonValueKind.String)
                {
                    return content.GetString() ?? string.Empty;
                }

                if (content.ValueKind == JsonValueKind.Array)
                {
                    var builder = new StringBuilder();
                    foreach (var part in content.EnumerateArray())
                    {
                        if (part.TryGetProperty("text", out var textProp))
                        {
                            builder.Append(textProp.GetString());
                        }
                    }
                    var combined = builder.ToString();
                    if (!string.IsNullOrWhiteSpace(combined))
                    {
                        return combined;
                    }
                }
            }

            if (choice.TryGetProperty("text", out var text))
            {
                return text.GetString() ?? string.Empty;
            }
        }

        if (doc.RootElement.TryGetProperty("text", out var rootText))
        {
            return rootText.GetString() ?? string.Empty;
        }

        return responseJson;
    }
}
