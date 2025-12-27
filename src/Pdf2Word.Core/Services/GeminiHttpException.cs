using System.Net;

namespace Pdf2Word.Core.Services;

public sealed class GeminiHttpException : Exception
{
    public HttpStatusCode StatusCode { get; }
    public string? ResponseSnippet { get; }

    public GeminiHttpException(HttpStatusCode statusCode, string? responseSnippet = null, string? message = null, Exception? innerException = null)
        : base(message ?? $"Gemini request failed with HTTP {(int)statusCode}.", innerException)
    {
        StatusCode = statusCode;
        ResponseSnippet = responseSnippet;
    }
}
