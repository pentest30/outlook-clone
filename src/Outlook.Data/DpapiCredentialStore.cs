using System.Runtime.Versioning;
using System.Security.Cryptography;
using System.Text;
using Outlook.Core.Services;

namespace Outlook.Data;

/// <summary>
/// Stores account passwords encrypted with Windows DPAPI, so only the signed-in Windows user can
/// decrypt them. macOS Keychain / libsecret backends come later.
/// </summary>
[SupportedOSPlatform("windows")]
public class DpapiCredentialStore(string directory) : ICredentialStore
{
    private static readonly byte[] Entropy = "OutlookClone.Credentials.v1"u8.ToArray();

    public void Save(int accountId, string secret)
    {
        Directory.CreateDirectory(directory);
        var encrypted = ProtectedData.Protect(Encoding.UTF8.GetBytes(secret), Entropy, DataProtectionScope.CurrentUser);
        File.WriteAllBytes(PathFor(accountId), encrypted);
    }

    public string? Load(int accountId)
    {
        var path = PathFor(accountId);
        if (!File.Exists(path))
            return null;
        var plain = ProtectedData.Unprotect(File.ReadAllBytes(path), Entropy, DataProtectionScope.CurrentUser);
        return Encoding.UTF8.GetString(plain);
    }

    public void Delete(int accountId)
    {
        var path = PathFor(accountId);
        if (File.Exists(path))
            File.Delete(path);
    }

    private string PathFor(int accountId) => Path.Combine(directory, $"account-{accountId}.bin");
}

/// <summary>Keeps secrets in memory only. Used by tests and on platforms without a secure store yet.</summary>
public class InMemoryCredentialStore : ICredentialStore
{
    private readonly Dictionary<int, string> _secrets = [];

    public void Save(int accountId, string secret) => _secrets[accountId] = secret;
    public string? Load(int accountId) => _secrets.GetValueOrDefault(accountId);
    public void Delete(int accountId) => _secrets.Remove(accountId);
}
