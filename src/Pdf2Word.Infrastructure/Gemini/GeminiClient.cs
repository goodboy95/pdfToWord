using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using Pdf2Word.Core.Options;
using Pdf2Word.Core.Services;

namespace Pdf2Word.Infrastructure.Gemini;

public sealed class GeminiClient : IGeminiClient
{
    private readonly HttpClient _httpClient;
    private readonly GeminiOptions _options;
    private readonly Func<string?> _apiKeyProvider;

    public GeminiClient(HttpClient httpClient, GeminiOptions options, Func<string?> apiKeyProvider)
    {
        _httpClient = httpClient;
        _options = options;
        _apiKeyProvider = apiKeyProvider;
    }

    public async Task<TableCellOcrResult> RecognizeTableCellsAsync(System.Drawing.Bitmap tableImage, IReadOnlyList<CellBoxForOcr> cells, CancellationToken ct, bool strictJson, int attempt)
    {
        var prompt = GeminiPrompts.BuildTablePrompt(cells, strictJson, attempt);
        var payload = BuildRequestPayload(prompt, tableImage, usePng: true);
        var json = await SendAsync(payload, _options.TimeoutSeconds.Table, ct).ConfigureAwait(false);
        var text = NormalizeJson(OpenAiResponseParser.ExtractText(json));
        return ParseTableCellResult(text, json);
    }

    public async Task<PageTextOcrResult> RecognizePageParagraphsAsync(System.Drawing.Bitmap pageImageMasked, CancellationToken ct, bool strictJson, int attempt)
    {
        var prompt = GeminiPrompts.BuildPagePrompt(strictJson, attempt);
        var payload = BuildRequestPayload(prompt, pageImageMasked, usePng: false);
        var json = await SendAsync(payload, _options.TimeoutSeconds.Page, ct).ConfigureAwait(false);
        var text = NormalizeJson(OpenAiResponseParser.ExtractText(json));
        return ParsePageParagraphs(text, json);
    }

    public async Task<TableTextLinesResult> RecognizeTableAsLinesAsync(System.Drawing.Bitmap tableImage, CancellationToken ct, int attempt)
    {
        var prompt = GeminiPrompts.BuildTableFallbackPrompt(attempt);
        var payload = BuildRequestPayload(prompt, tableImage, usePng: true);
        var json = await SendAsync(payload, _options.TimeoutSeconds.Table, ct).ConfigureAwait(false);
        var text = NormalizeJson(OpenAiResponseParser.ExtractText(json));
        return ParseTableLines(text, json);
    }

    public async Task<PageTextFallbackResult> RecognizePageAsSingleTextAsync(System.Drawing.Bitmap pageImage, CancellationToken ct, int attempt)
    {
        var prompt = GeminiPrompts.BuildPageFallbackPrompt(attempt);
        var payload = BuildRequestPayload(prompt, pageImage, usePng: false);
        var json = await SendAsync(payload, _options.TimeoutSeconds.Page, ct).ConfigureAwait(false);
        var text = NormalizeJson(OpenAiResponseParser.ExtractText(json));
        return ParsePageFallback(text, json);
    }

    private string BuildRequestPayload(string prompt, System.Drawing.Bitmap image, bool usePng)
    {
        var imageBytes = ImageEncoder.Encode(image, usePng, _options.Image.MaxLongEdgePx, _options.Image.JpegQuality);
        var mime = usePng ? "image/png" : "image/jpeg";
        var dataUrl = $"data:{mime};base64,{Convert.ToBase64String(imageBytes)}";

        var body = new
        {
            model = _options.Model,
            messages = new[]
            {
                new
                {
                    role = "user",
                    content = new object[]
                    {
                        new { type = "text", text = prompt },
                        new { type = "image_url", image_url = new { url = dataUrl } }
                    }
                }
            },
            temperature = 0
        };

        return JsonSerializer.Serialize(body);
    }

    private async Task<string> SendAsync(string payload, int timeoutSeconds, CancellationToken ct)
    {
        var apiKey = _apiKeyProvider();
        if (string.IsNullOrWhiteSpace(apiKey))
        {
            throw new InvalidOperationException("AI API Key is missing.");
        }

        var baseUrl = _options.BaseUrl;
        if (string.IsNullOrWhiteSpace(baseUrl))
        {
            throw new InvalidOperationException("AI Base URL is not configured.");
        }

        if (string.IsNullOrWhiteSpace(_options.Model))
        {
            throw new InvalidOperationException("AI model is not configured.");
        }

        var url = BuildChatCompletionsUrl(baseUrl);
        using var content = new StringContent(payload, Encoding.UTF8, "application/json");
        _httpClient.DefaultRequestHeaders.Accept.Clear();
        _httpClient.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
        _httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", apiKey);
        _httpClient.Timeout = TimeSpan.FromSeconds(timeoutSeconds);

        using var response = await _httpClient.PostAsync(url, content, ct).ConfigureAwait(false);
        if (!response.IsSuccessStatusCode)
        {
            var snippet = await ReadSnippetAsync(response, ct).ConfigureAwait(false);
            throw new GeminiHttpException(response.StatusCode, snippet);
        }
        return await response.Content.ReadAsStringAsync(ct).ConfigureAwait(false);
    }

    private static async Task<string?> ReadSnippetAsync(HttpResponseMessage response, CancellationToken ct)
    {
        try
        {
            var content = await response.Content.ReadAsStringAsync(ct).ConfigureAwait(false);
            if (string.IsNullOrWhiteSpace(content))
            {
                return null;
            }

            return content.Length <= 300 ? content : content[..300];
        }
        catch
        {
            return null;
        }
    }

    private static string BuildChatCompletionsUrl(string baseUrl)
    {
        var trimmed = baseUrl.TrimEnd('/');
        return $"{trimmed}/chat/completions";
    }

    private static string NormalizeJson(string text)
    {
        if (string.IsNullOrWhiteSpace(text))
        {
            return "{}";
        }

        var trimmed = text.Trim();
        if (trimmed.StartsWith("```", StringComparison.Ordinal))
        {
            var start = trimmed.IndexOf('\n');
            if (start >= 0)
            {
                trimmed = trimmed[(start + 1)..];
            }
            var endFence = trimmed.LastIndexOf("```");
            if (endFence >= 0)
            {
                trimmed = trimmed[..endFence];
            }
            trimmed = trimmed.Trim();
        }

        return trimmed;
    }

    private static TableCellOcrResult ParseTableCellResult(string outputText, string rawJson)
    {
        using var doc = JsonDocument.Parse(outputText);
        var result = new TableCellOcrResult { RawJson = rawJson };
        if (!doc.RootElement.TryGetProperty("cells", out var cells) || cells.ValueKind != JsonValueKind.Array)
        {
            return result;
        }

        foreach (var cell in cells.EnumerateArray())
        {
            if (!cell.TryGetProperty("id", out var idProp))
            {
                continue;
            }
            var id = idProp.GetString();
            if (string.IsNullOrWhiteSpace(id))
            {
                continue;
            }

            var text = cell.TryGetProperty("text", out var textProp) ? textProp.GetString() ?? string.Empty : string.Empty;
            result.TextById[id] = text;
        }

        return result;
    }

    private static PageTextOcrResult ParsePageParagraphs(string outputText, string rawJson)
    {
        using var doc = JsonDocument.Parse(outputText);
        var result = new PageTextOcrResult { RawJson = rawJson };
        if (!doc.RootElement.TryGetProperty("paragraphs", out var paragraphs) || paragraphs.ValueKind != JsonValueKind.Array)
        {
            return result;
        }

        foreach (var para in paragraphs.EnumerateArray())
        {
            var role = para.TryGetProperty("role", out var roleProp) ? roleProp.GetString() ?? "body" : "body";
            var text = para.TryGetProperty("text", out var textProp) ? textProp.GetString() ?? string.Empty : string.Empty;
            result.Paragraphs.Add(new ParagraphDto { Role = role, Text = text });
        }

        return result;
    }

    private static TableTextLinesResult ParseTableLines(string outputText, string rawJson)
    {
        using var doc = JsonDocument.Parse(outputText);
        var result = new TableTextLinesResult { RawJson = rawJson };
        if (doc.RootElement.TryGetProperty("lines", out var lines) && lines.ValueKind == JsonValueKind.Array)
        {
            foreach (var line in lines.EnumerateArray())
            {
                result.Lines.Add(line.GetString() ?? string.Empty);
            }
        }
        else if (doc.RootElement.TryGetProperty("text", out var text))
        {
            var value = text.GetString() ?? string.Empty;
            if (!string.IsNullOrWhiteSpace(value))
            {
                result.Lines.AddRange(value.Split('\n'));
            }
        }

        return result;
    }

    private static PageTextFallbackResult ParsePageFallback(string outputText, string rawJson)
    {
        using var doc = JsonDocument.Parse(outputText);
        var textValue = string.Empty;
        if (doc.RootElement.TryGetProperty("text", out var text))
        {
            textValue = text.GetString() ?? string.Empty;
        }

        return new PageTextFallbackResult
        {
            RawJson = rawJson,
            Text = textValue
        };
    }
}
