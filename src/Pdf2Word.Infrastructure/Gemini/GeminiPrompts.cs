using System.Text.Json;
using Pdf2Word.Core.Services;

namespace Pdf2Word.Infrastructure.Gemini;

public static class GeminiPrompts
{
    public static string BuildTablePrompt(IReadOnlyList<CellBoxForOcr> cells, bool strict, int attempt)
    {
        var payload = JsonSerializer.Serialize(cells.Select(c => new { id = c.Id, x = c.X, y = c.Y, w = c.W, h = c.H }));
        var builder = new List<string>
        {
            "你是一个 OCR 引擎。",
            "你将收到一张表格图片，以及一个 cells 列表，每个 cell 提供 id 和矩形坐标。",
            "识别每个矩形区域里的文字，并返回 JSON。",
            "必须严格输出 JSON，禁止输出任何解释、Markdown、代码块。",
            "不得猜测/补全，看不清就输出空字符串。",
            "必须保留原文：中文、英文字母、数字、标点、空格。",
            "识别时不要把表格线当成字符，只输出文字内容。",
            "输出格式：{\"cells\":[{\"id\":\"...\",\"text\":\"...\"}, ...]}"
        };
        if (strict)
        {
            builder.Add("如果输出包含任何非 JSON 字符，将被判为失败。必须包含输入中所有 id。");
        }

        builder.Add($"cells: {payload}");
        return string.Join("\n", builder);
    }

    public static string BuildPagePrompt(bool strict, int attempt)
    {
        var builder = new List<string>
        {
            "你是 OCR 引擎，识别图片中的正文文字。",
            "图片中可能存在被遮罩的空白区域（表格），请忽略空白区域。",
            "按阅读顺序输出段落列表。",
            "必须严格输出 JSON，禁止输出任何解释、Markdown、代码块。",
            "不得猜测/补全，看不清的字可用 ? 代替或省略该词。",
            "段内换行使用 \\n。",
            "若出现明显标题，将该段 role 标为 title，否则 body。",
            "输出格式：{\"paragraphs\":[{\"role\":\"title|body\",\"text\":\"...\"}, ...]}"
        };
        if (strict)
        {
            builder.Add("如果输出包含任何非 JSON 字符，将被判为失败。仅输出 JSON。");
        }

        return string.Join("\n", builder);
    }

    public static string BuildTableFallbackPrompt(int attempt)
    {
        return string.Join("\n", new[]
        {
            "你是 OCR 引擎。请识别整张表格图片，按行输出纯文本。",
            "必须严格输出 JSON，不要解释。",
            "输出格式：{\"lines\":[\"...\",\"...\"]}"
        });
    }

    public static string BuildPageFallbackPrompt(int attempt)
    {
        return string.Join("\n", new[]
        {
            "你是 OCR 引擎，请识别图片中的所有文字，输出为单段文本。",
            "必须严格输出 JSON，不要解释。",
            "输出格式：{\"text\":\"...\"}"
        });
    }
}
