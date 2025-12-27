using System.Text.RegularExpressions;

namespace Pdf2Word.Core.Services;

public sealed class PageRangeParseResult
{
    public List<int> Pages { get; } = new();
    public List<string> Warnings { get; } = new();
    public bool HasError { get; set; }
    public string? ErrorMessage { get; set; }
}

public static class PageRangeParser
{
    public static PageRangeParseResult Parse(string input, int totalPages)
    {
        var result = new PageRangeParseResult();
        if (totalPages <= 0)
        {
            result.HasError = true;
            result.ErrorMessage = "totalPages must be positive";
            return result;
        }

        var normalized = (input ?? string.Empty).Trim();
        if (string.IsNullOrWhiteSpace(normalized))
        {
            result.Pages.AddRange(Enumerable.Range(1, totalPages));
            return result;
        }

        if (normalized.Contains('，'))
        {
            normalized = normalized.Replace('，', ',');
            result.Warnings.Add("已将中文逗号替换为英文逗号。");
        }

        var tokens = normalized.Split(',', StringSplitOptions.None);
        var pages = new SortedSet<int>();
        foreach (var rawToken in tokens)
        {
            var token = rawToken.Trim();
            if (string.IsNullOrWhiteSpace(token))
            {
                result.HasError = true;
                result.ErrorMessage = "页码范围包含空片段。";
                return result;
            }

            if (token.Contains('-'))
            {
                var parts = token.Split('-', StringSplitOptions.None);
                if (parts.Length != 2 || string.IsNullOrWhiteSpace(parts[0]) || string.IsNullOrWhiteSpace(parts[1]) ||
                    !int.TryParse(parts[0].Trim(), out var start) || !int.TryParse(parts[1].Trim(), out var end))
                {
                    result.HasError = true;
                    result.ErrorMessage = "页码范围格式不正确。";
                    return result;
                }

                if (start > end)
                {
                    (start, end) = (end, start);
                    result.Warnings.Add($"区间 {parts[0]}-{parts[1]} 已自动纠正为 {start}-{end}。");
                }

                for (var i = start; i <= end; i++)
                {
                    if (i < 1 || i > totalPages)
                    {
                        result.Warnings.Add($"已忽略越界页码：{i}。");
                        continue;
                    }

                    pages.Add(i);
                }
            }
            else
            {
                if (!Regex.IsMatch(token, "^\\d+$") || !int.TryParse(token, out var page))
                {
                    result.HasError = true;
                    result.ErrorMessage = "页码范围格式不正确。";
                    return result;
                }

                if (page < 1 || page > totalPages)
                {
                    result.Warnings.Add($"已忽略越界页码：{page}。");
                    continue;
                }

                pages.Add(page);
            }
        }

        if (pages.Count == 0)
        {
            result.HasError = true;
            result.ErrorMessage = "页码范围为空或全部越界。";
            return result;
        }

        result.Pages.AddRange(pages);
        return result;
    }
}
