using System;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using Microsoft.Win32;

namespace Neurow.Pictos
{
    internal sealed class PairingStore
    {
        private readonly string path;
        private static readonly byte[] Entropy = Encoding.UTF8.GetBytes("Neurow.Pictos.Pairing.v1");
        internal PairingStore(string directory) { Directory.CreateDirectory(directory); path = Path.Combine(directory, "pairing.dat"); }
        internal string LoadOrCreate()
        {
            if (!File.Exists(path)) return Renew();
            byte[] secret = ProtectedData.Unprotect(File.ReadAllBytes(path), Entropy, DataProtectionScope.CurrentUser);
            if (secret.Length != 32) throw new CryptographicException("Invalid pairing record.");
            try { return Hex(secret); } finally { Array.Clear(secret, 0, secret.Length); }
        }
        internal string Renew()
        {
            byte[] secret = new byte[32];
            using (var random = RandomNumberGenerator.Create()) random.GetBytes(secret);
            string temporary = path + "." + Guid.NewGuid().ToString("N") + ".tmp";
            try
            {
                byte[] encrypted = ProtectedData.Protect(secret, Entropy, DataProtectionScope.CurrentUser);
                using (var file = new FileStream(temporary, FileMode.CreateNew, FileAccess.Write, FileShare.None))
                { file.Write(encrypted, 0, encrypted.Length); file.Flush(true); }
                if (File.Exists(path)) File.Replace(temporary, path, null); else File.Move(temporary, path);
                return Hex(secret);
            }
            finally { Array.Clear(secret, 0, secret.Length); if (File.Exists(temporary)) File.Delete(temporary); }
        }
        private static string Hex(byte[] bytes) { return BitConverter.ToString(bytes).Replace("-", "").ToLowerInvariant(); }
    }

    internal sealed class StartupRegistration
    {
        private readonly string keyPath;
        private const string ValueName = "Neurow.Pictos";
        internal StartupRegistration(string keyPath = @"Software\Microsoft\Windows\CurrentVersion\Run") { this.keyPath = keyPath; }
        internal static string Command(string executable)
        {
            if (executable.IndexOf('"') >= 0) throw new ArgumentException("Invalid executable path.");
            string value = "\"" + executable + "\" --startup";
            if (value.Length > 260) throw new ArgumentException("Déplacez le compagnon dans un dossier au chemin plus court.");
            return value;
        }
        internal bool IsEnabled(string executable)
        {
            using (var key = Registry.CurrentUser.OpenSubKey(keyPath))
                return string.Equals(key?.GetValue(ValueName) as string, Command(executable), StringComparison.OrdinalIgnoreCase);
        }
        internal void SetEnabled(string executable, bool enabled)
        {
            using (var key = Registry.CurrentUser.CreateSubKey(keyPath))
            {
                if (enabled) key.SetValue(ValueName, Command(executable), RegistryValueKind.String);
                else key.DeleteValue(ValueName, false);
            }
        }
    }
}
