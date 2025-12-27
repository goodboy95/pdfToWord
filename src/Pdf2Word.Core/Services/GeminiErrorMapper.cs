using System.Net;

namespace Pdf2Word.Core.Services;

public readonly record struct GeminiErrorInfo(string ErrorCode, string Message, bool Retryable);

public static class GeminiErrorMapper
{
    public static GeminiErrorInfo FromException(Exception ex, bool cancellationRequested, string fallbackMessage)
    {
        if (ex is TaskCanceledException && !cancellationRequested)
        {
            return new GeminiErrorInfo("E_GEMINI_TIMEOUT", "识别服务请求超时", true);
        }

        if (ex is GeminiHttpException geminiHttp)
        {
            return FromStatusCode(geminiHttp.StatusCode);
        }

        if (ex is HttpRequestException http && http.StatusCode.HasValue)
        {
            return FromStatusCode(http.StatusCode.Value);
        }

        return new GeminiErrorInfo("E_GEMINI_TIMEOUT", fallbackMessage, true);
    }

    public static GeminiErrorInfo FromStatusCode(HttpStatusCode statusCode)
    {
        var code = (int)statusCode;
        if (code == 429)
        {
            return new GeminiErrorInfo("E_GEMINI_RATE_LIMITED", "识别服务繁忙（HTTP 429）", true);
        }

        if (code >= 500)
        {
            return new GeminiErrorInfo("E_GEMINI_SERVER_ERROR", $"识别服务暂时不可用（HTTP {code}）", true);
        }

        return new GeminiErrorInfo("E_GEMINI_HTTP_ERROR", $"识别服务返回 HTTP {code}，请检查 Base URL/模型/API Key 或网络限制", false);
    }
}
