using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using System.Threading;
using System.Web.Script.Serialization;
using Microsoft.Win32;

namespace Neurow.Pictos
{
    internal static class SelfTests
    {
        internal static int Run(string[] args)
        {
            string report = Argument(args, "--report"), node = Argument(args, "--node");
            if (string.IsNullOrEmpty(report) || !File.Exists(node)) return 2;
            string temporary = Path.Combine(Path.GetTempPath(), "pictos-native-test-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(temporary);
            var checks = new List<object>(); int failures = 0;
            Action<string, Action> check = (name, action) => {
                try { action(); checks.Add(new { name, passed = true }); }
                catch (Exception error) { failures++; checks.Add(new { name, passed = false, error = error.GetType().Name }); }
            };
            try
            {
                check("DPAPI: code stable, chiffré et renouvelable", () => {
                    string directory = Path.Combine(temporary, "pairing");
                    var store = new PairingStore(directory); string first = store.LoadOrCreate();
                    Require(first.Length == 64 && first == new PairingStore(directory).LoadOrCreate());
                    byte[] encrypted = File.ReadAllBytes(Path.Combine(directory, "pairing.dat"));
                    Require(Encoding.UTF8.GetString(encrypted).IndexOf(first, StringComparison.Ordinal) < 0);
                    string next = store.Renew(); Require(first != next && next == store.LoadOrCreate());
                    File.WriteAllBytes(Path.Combine(directory, "pairing.dat"), new byte[] { 1, 2, 3 });
                    bool refused = false; try { store.LoadOrCreate(); } catch (CryptographicException) { refused = true; }
                    Require(refused && File.ReadAllBytes(Path.Combine(directory, "pairing.dat")).Length == 3);
                });
                check("Démarrage utilisateur: chemin cité, activation et suppression ciblées", () => {
                    string key = @"Software\NeurowPictosSelfTests\" + Guid.NewGuid().ToString("N");
                    try
                    {
                        using (var other = Registry.CurrentUser.CreateSubKey(key)) other.SetValue("Keep", "untouched");
                        var registration = new StartupRegistration(key);
                        const string executable = @"C:\Users\Test User\Pictos\Neurow.Pictos.exe";
                        Require(StartupRegistration.Command(executable) == "\"" + executable + "\" --startup");
                        registration.SetEnabled(executable, true); Require(registration.IsEnabled(executable));
                        registration.SetEnabled(executable, false); Require(!registration.IsEnabled(executable));
                        using (var other = Registry.CurrentUser.OpenSubKey(key)) Require((string)other.GetValue("Keep") == "untouched");
                    }
                    finally { Registry.CurrentUser.DeleteSubKeyTree(key, false); }
                });
                string root = Path.Combine(temporary, "portable with spaces");
                Directory.CreateDirectory(Path.Combine(root, "runtime")); Directory.CreateDirectory(Path.Combine(root, "bridge"));
                File.Copy(node, Path.Combine(root, "runtime", "node.exe"));
                File.WriteAllText(Path.Combine(root, "bridge", "server.js"),
                    "const{spawn}=require('node:child_process');let b='',started=false;process.stdin.setEncoding('utf8');process.stdin.on('data',d=>{b+=d;let n;while((n=b.indexOf('\\n'))>=0){const m=JSON.parse(b.slice(0,n));b=b.slice(n+1);if(!started){if(!/^[a-f0-9]{64}$/.test(m.pairingToken))process.exit(9);started=true;const child=spawn(process.execPath,['-e','setInterval(()=>{},1000)'],{stdio:'ignore',windowsHide:true});console.log(JSON.stringify({event:'ready',child:child.pid}));}else if(m.command==='stop')process.exit(0);}});", new UTF8Encoding(false));
                check("Processus caché: secret par canal privé et arrêt des descendants", () => ExerciseHost(root, temporary, false));
                check("Fermeture brutale du lanceur: arrêt de Node et de ses descendants", () => ExerciseHost(root, temporary, true));
            }
            catch (Exception error) { failures++; checks.Add(new { name = "Préparation", passed = false, error = error.GetType().Name }); }
            finally
            {
                try { Directory.Delete(temporary, true); } catch { }
                File.WriteAllText(report, new JavaScriptSerializer().Serialize(new { passed = failures == 0, checks }), new UTF8Encoding(false));
            }
            return failures == 0 ? 0 : 1;
        }
        private static void ExerciseHost(string root, string temporary, bool abrupt)
        {
            using (var ready = new ManualResetEvent(false))
            using (var host = new BridgeHost())
            {
                string token = new string('a', 64), output = ""; int descendant = 0;
                host.Output += line => {
                    output += line;
                    var message = new JavaScriptSerializer().Deserialize<Dictionary<string, object>>(line);
                    if (message.TryGetValue("child", out object child)) { descendant = Convert.ToInt32(child); ready.Set(); }
                };
                host.Start(root, temporary, token);
                Require(ready.WaitOne(10000));
                int parent = host.Process.Id;
                Require(host.Process.MainWindowHandle == IntPtr.Zero);
                Require(!host.Process.StartInfo.Arguments.Contains(token));
                Require(!output.Contains(token));
                foreach (string value in host.Process.StartInfo.EnvironmentVariables.Values) Require(value != token);
                if (abrupt) host.Dispose(); else host.StopAsync().GetAwaiter().GetResult();
                var deadline = DateTime.UtcNow.AddSeconds(5);
                while ((Exists(parent) || Exists(descendant)) && DateTime.UtcNow < deadline) Thread.Sleep(50);
                Require(!Exists(parent) && !Exists(descendant));
            }
        }
        private static bool Exists(int id) { try { using (var process = Process.GetProcessById(id)) return !process.HasExited; } catch (ArgumentException) { return false; } }
        private static string Argument(string[] args, string name) { int i = Array.IndexOf(args, name); return i >= 0 && i + 1 < args.Length ? args[i + 1] : null; }
        private static void Require(bool value) { if (!value) throw new InvalidOperationException("Self-test assertion failed."); }
    }
}
