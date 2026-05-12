using System;
using System.Text;

namespace TailorMail.Helpers;

public static class CredentialHelper
{
    private static readonly byte[] _entropy = Encoding.UTF8.GetBytes("TailorMail_SecureV1");

    public static string Protect(string plainText)
    {
        if (string.IsNullOrEmpty(plainText)) return string.Empty;
        var data = Encoding.UTF8.GetBytes(plainText);
        var protectedData = System.Security.Cryptography.ProtectedData.Protect(
            data, _entropy, System.Security.Cryptography.DataProtectionScope.CurrentUser);
        return Convert.ToBase64String(protectedData);
    }

    public static string Unprotect(string protectedText)
    {
        if (string.IsNullOrEmpty(protectedText)) return string.Empty;
        try
        {
            var data = Convert.FromBase64String(protectedText);
            var unprotectedData = System.Security.Cryptography.ProtectedData.Unprotect(
                data, _entropy, System.Security.Cryptography.DataProtectionScope.CurrentUser);
            return Encoding.UTF8.GetString(unprotectedData);
        }
        catch
        {
            return string.Empty;
        }
    }
}