using System.IO;
using System.Security.Cryptography;
using System.Text;
using Pdf2Word.Core.Options;
using Pdf2Word.Core.Services;

namespace Pdf2Word.App.Services;

public sealed class DpapiApiKeyStore : IApiKeyStore
{
    private readonly string _path;
    private readonly GeminiOptions _options;
    private string? _memoryKey;

    public DpapiApiKeyStore(AppOptions options)
    {
        _options = options.Gemini;
        var dir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Pdf2Word");
        Directory.CreateDirectory(dir);
        _path = Path.Combine(dir, "gemini.key");
    }

    public string? GetApiKey()
    {
        if (!string.IsNullOrWhiteSpace(_memoryKey))
        {
            return _memoryKey;
        }

        if (_options.ApiKeyStorage == GeminiKeyStorage.EnvOnly)
        {
            return ReadEnvApiKey();
        }

        if (_options.ApiKeyStorage == GeminiKeyStorage.None)
        {
            return ReadEnvApiKey();
        }

        if (!File.Exists(_path))
        {
            return ReadEnvApiKey();
        }

        var protectedBytes = File.ReadAllBytes(_path);
        var bytes = ProtectedData.Unprotect(protectedBytes, null, DataProtectionScope.CurrentUser);
        return Encoding.UTF8.GetString(bytes);
    }

    public void SaveApiKey(string apiKey)
    {
        if (string.IsNullOrWhiteSpace(apiKey))
        {
            return;
        }

        _memoryKey = apiKey;

        if (_options.ApiKeyStorage != GeminiKeyStorage.DPAPI)
        {
            return;
        }

        var bytes = Encoding.UTF8.GetBytes(apiKey);
        var protectedBytes = ProtectedData.Protect(bytes, null, DataProtectionScope.CurrentUser);
        File.WriteAllBytes(_path, protectedBytes);
    }

    private static string? ReadEnvApiKey()
    {
        return Environment.GetEnvironmentVariable("AI_API_KEY")
            ?? Environment.GetEnvironmentVariable("OPENAI_API_KEY")
            ?? Environment.GetEnvironmentVariable("GEMINI_API_KEY");
    }
}
