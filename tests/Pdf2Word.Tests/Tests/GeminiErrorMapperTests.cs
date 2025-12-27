using System.Net;
using Pdf2Word.Core.Services;
using Xunit;

namespace Pdf2Word.Tests.Tests;

public class GeminiErrorMapperTests
{
    [Fact]
    public void TimeoutMapsToTimeoutError()
    {
        var error = GeminiErrorMapper.FromException(new TaskCanceledException(), cancellationRequested: false, fallbackMessage: "fallback");
        Assert.Equal("E_GEMINI_TIMEOUT", error.ErrorCode);
        Assert.True(error.Retryable);
    }

    [Fact]
    public void RateLimitMapsToRateLimitedError()
    {
        var error = GeminiErrorMapper.FromException(new GeminiHttpException(HttpStatusCode.TooManyRequests), cancellationRequested: false, fallbackMessage: "fallback");
        Assert.Equal("E_GEMINI_RATE_LIMITED", error.ErrorCode);
        Assert.True(error.Retryable);
    }

    [Fact]
    public void ServerErrorMapsToServerError()
    {
        var error = GeminiErrorMapper.FromException(new GeminiHttpException(HttpStatusCode.ServiceUnavailable), cancellationRequested: false, fallbackMessage: "fallback");
        Assert.Equal("E_GEMINI_SERVER_ERROR", error.ErrorCode);
        Assert.True(error.Retryable);
    }

    [Fact]
    public void ForbiddenMapsToHttpErrorAndNotRetryable()
    {
        var error = GeminiErrorMapper.FromException(new GeminiHttpException(HttpStatusCode.Forbidden), cancellationRequested: false, fallbackMessage: "fallback");
        Assert.Equal("E_GEMINI_HTTP_ERROR", error.ErrorCode);
        Assert.False(error.Retryable);
    }

    [Fact]
    public void HttpRequestExceptionStatusMapsToServerError()
    {
        var ex = new HttpRequestException("fail", null, HttpStatusCode.BadGateway);
        var error = GeminiErrorMapper.FromException(ex, cancellationRequested: false, fallbackMessage: "fallback");
        Assert.Equal("E_GEMINI_SERVER_ERROR", error.ErrorCode);
        Assert.True(error.Retryable);
    }
}
