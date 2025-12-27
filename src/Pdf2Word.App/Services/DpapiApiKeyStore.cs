using System.Security.Cryptography;
using System.Text;
using Pdf2Word.Core.Services;

namespace Pdf2Word.App.Services;

public sealed class DpapiApiKeyStore : IApiKeyStore
{
    private readonly string _path;

    public DpapiApiKeyStore()
    {
        var dir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Pdf2Word");
        Directory.CreateDirectory(dir);
        _path = Path.Combine(dir, "gemini.key");
    }

    public string? GetApiKey()
    {
        if (!File.Exists(_path))
        {
            return Environment.GetEnvironmentVariable("GEMINI_API_KEY");
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

        var bytes = Encoding.UTF8.GetBytes(apiKey);
        var protectedBytes = ProtectedData.Protect(bytes, null, DataProtectionScope.CurrentUser);
        File.WriteAllBytes(_path, protectedBytes);
    }
}
