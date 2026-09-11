using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.IO;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Security.Principal;
using System.Threading;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using System.Windows.Forms;

namespace Neurow.Pictos
{
    internal static class Program
    {
        [STAThread]
        private static int Main(string[] args)
        {
            if (Array.IndexOf(args, "--self-test") >= 0) return SelfTests.Run(args);
            string user = WindowsIdentity.GetCurrent().User.Value;
            using (var instance = new Mutex(true, @"Local\Neurow.Pictos." + user, out bool created))
            {
                if (!created) return 0;
                try
                {
                    Application.EnableVisualStyles(); Application.SetCompatibleTextRenderingDefault(false);
                    using (var tray = new TrayContext())
                    {
                        Application.ThreadException += (_, __) => tray.ReportFailure();
                        Application.Run(tray);
                    }
                    return 0;
                }
                catch { return 1; }
                finally { instance.ReleaseMutex(); }
            }
        }
    }

    internal sealed class TrayContext : ApplicationContext
    {
        private const string BaseUrl = "http://127.0.0.1:43129";
        private readonly string executable = Assembly.GetExecutingAssembly().Location;
        private readonly string root = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
        private readonly string directory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "AtelierPictos");
        private readonly Control dispatcher = new Control();
        private readonly NotifyIcon tray;
        private readonly ContextMenuStrip menu = new ContextMenuStrip();
        private readonly ToolStripMenuItem status = new ToolStripMenuItem("Démarrage…") { Enabled = false };
        private readonly ToolStripMenuItem open = new ToolStripMenuItem("Connexion et liaison PowerPoint…");
        private readonly ToolStripMenuItem copy = new ToolStripMenuItem("Copier le code de liaison");
        private readonly ToolStripMenuItem startup = new ToolStripMenuItem("Démarrer avec Windows · cet utilisateur");
        private readonly StartupRegistration registration = new StartupRegistration();
        private readonly System.Windows.Forms.Timer timer = new System.Windows.Forms.Timer { Interval = 2000 };
        private readonly HttpClient http = new HttpClient(new HttpClientHandler { UseProxy = false, AllowAutoRedirect = false }) { Timeout = TimeSpan.FromSeconds(10) };
        private readonly Icon waiting = MakeIcon(Color.FromArgb(164, 105, 16));
        private readonly Icon connected = MakeIcon(Color.FromArgb(40, 117, 65));
        private readonly Icon failed = MakeIcon(Color.FromArgb(184, 24, 49));
        private PairingStore pairing;
        private BridgeHost host;
        private string token, lastError;
        private bool ready, quitting, polling, restarting;
        private int generation, failures;
        private DateTime nextRestart = DateTime.MaxValue, startedAt;

        internal TrayContext()
        {
            var handle = dispatcher.Handle; // Hidden message target; never creates a visible window.
            tray = new NotifyIcon { Icon = waiting, Text = "Neurow.Pictos — démarrage", ContextMenuStrip = menu, Visible = true };
            menu.Items.Add(new ToolStripMenuItem("Neurow.Pictos") { Enabled = false });
            menu.Items.Add(status); menu.Items.Add(new ToolStripSeparator());
            menu.Items.Add(open); menu.Items.Add(copy); menu.Items.Add(new ToolStripSeparator());
            menu.Items.Add(startup);
            var restart = new ToolStripMenuItem("Relancer le compagnon"); menu.Items.Add(restart);
            var more = new ToolStripMenuItem("Options de liaison");
            var renew = new ToolStripMenuItem("Créer un nouveau code…"); more.DropDownItems.Add(renew); menu.Items.Add(more);
            menu.Items.Add(new ToolStripSeparator());
            var quit = new ToolStripMenuItem("Quitter"); menu.Items.Add(quit);
            open.Click += async (_, __) => await OpenAsync();
            copy.Click += (_, __) => CopyCode();
            startup.Click += (_, __) => ToggleStartup();
            restart.Click += async (_, __) => await RestartAsync();
            renew.Click += async (_, __) => await RenewAsync();
            quit.Click += async (_, __) => await QuitAsync();
            menu.Opening += (_, __) => { SyncStartup(); copy.Text = "Copier le code de liaison"; };
            tray.MouseUp += (_, e) => { if (e.Button == MouseButtons.Left) menu.Show(Cursor.Position); };
            timer.Tick += async (_, __) => {
                if (DateTime.UtcNow >= nextRestart) { nextRestart = DateTime.MaxValue; await RestartAsync(false); }
                else if (ready) await RefreshAsync();
            };
            try { pairing = new PairingStore(directory); token = pairing.LoadOrCreate(); StartHost(); }
            catch { SetStatus("Liaison à réparer · Options de liaison", failed); }
            SyncStartup(); timer.Start();
        }
        private void Ui(Action action)
        {
            if (quitting || dispatcher.IsDisposed) return;
            try { dispatcher.BeginInvoke(new Action(() => { if (!quitting) action(); })); } catch (InvalidOperationException) { }
        }
        private void StartHost()
        {
            if (quitting || string.IsNullOrEmpty(token)) return;
            int current = ++generation; ready = false; lastError = null; startedAt = DateTime.UtcNow;
            SetStatus("Démarrage du compagnon…", waiting);
            host = new BridgeHost();
            host.Output += line => Ui(() => {
                if (current != generation) return;
                try
                {
                    var message = new JavaScriptSerializer().Deserialize<Dictionary<string, object>>(line);
                    if (Value(message, "event") == "ready" && Value(message, "url") == BaseUrl)
                    { ready = true; SetStatus("Vérification de ChatGPT…", waiting); _ = RefreshAsync(); }
                    else if (Value(message, "event") == "error")
                    { lastError = Value(message, "code") == "EADDRINUSE" ? "Ancien compagnon ouvert · fermez son terminal" : "Démarrage impossible · relancez le compagnon"; SetStatus(lastError, failed); }
                }
                catch { /* Do not display or retain arbitrary child output. */ }
            });
            host.Exited += () => Ui(() => {
                if (current != generation || restarting) return;
                ready = false;
                SetStatus(lastError ?? "Compagnon arrêté · relance en cours", failed);
                if (DateTime.UtcNow - startedAt > TimeSpan.FromMinutes(5)) failures = 0;
                failures++;
                if (failures <= 3) { nextRestart = DateTime.UtcNow.AddSeconds(5 * failures); timer.Interval = 2000; }
                else SetStatus(lastError ?? "Compagnon arrêté · utilisez Relancer", failed);
            });
            try { host.Start(root, directory, token); }
            catch (FileNotFoundException) { lastError = "Dossier incomplet · décompressez toute l’archive"; SetStatus(lastError, failed); }
            catch { lastError = "Démarrage impossible · vérifiez les autorisations du poste"; SetStatus(lastError, failed); }
        }
        private async Task RefreshAsync()
        {
            if (polling || !ready || quitting) return;
            polling = true; int current = generation;
            try
            {
                var state = await RequestAsync("connection");
                if (current != generation || quitting) return;
                bool online = Flag(state, "connected");
                bool pending = state.TryGetValue("login", out object login) && login is Dictionary<string, object> details && Flag(details, "pending");
                timer.Interval = pending ? 3000 : 30000;
                SetStatus(online ? (Flag(state, "busy") ? "Recherche en cours" : "Connecté à ChatGPT") : pending ? "Connexion ChatGPT en cours…" : "ChatGPT à connecter", online ? connected : waiting);
            }
            catch { if (current == generation && !quitting) SetStatus("Connexion à vérifier · ouvrez la liaison", failed); }
            finally { polling = false; }
        }
        private async Task<Dictionary<string, object>> RequestAsync(string route, string body = null)
        {
            using (var request = new HttpRequestMessage(body == null ? HttpMethod.Get : HttpMethod.Post, BaseUrl + "/v1/" + route))
            {
                request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
                if (body != null) request.Content = new StringContent(body, System.Text.Encoding.UTF8, "application/json");
                using (var response = await http.SendAsync(request))
                {
                    response.EnsureSuccessStatusCode();
                    string text = await response.Content.ReadAsStringAsync();
                    return new JavaScriptSerializer { MaxJsonLength = 65536 }.Deserialize<Dictionary<string, object>>(text);
                }
            }
        }
        private async Task OpenAsync()
        {
            if (!ready || quitting) return;
            open.Enabled = false;
            try
            {
                // A short-lived, single-use ticket keeps the permanent code out of browser arguments.
                var result = await RequestAsync("dashboard-ticket", "{}");
                string ticket = Value(result, "ticket");
                if (!System.Text.RegularExpressions.Regex.IsMatch(ticket, "^[a-f0-9]{48}$")) throw new InvalidOperationException();
                Process.Start(new ProcessStartInfo(BaseUrl + "/#link=" + ticket) { UseShellExecute = true });
            }
            catch { SetStatus("Ouverture impossible · relancez le compagnon", failed); }
            finally { open.Enabled = ready && !quitting; }
        }
        private void CopyCode()
        {
            if (string.IsNullOrEmpty(token)) return;
            try { Clipboard.SetText(token); copy.Text = "Code copié ✓"; }
            catch (ExternalException) { copy.Text = "Presse-papiers occupé · réessayez"; }
        }
        private void ToggleStartup()
        {
            try { registration.SetEnabled(executable, !registration.IsEnabled(executable)); SyncStartup(); }
            catch { SetStatus("Démarrage automatique non autorisé pour ce dossier", failed); }
        }
        private void SyncStartup()
        {
            try { startup.Checked = registration.IsEnabled(executable); } catch { startup.Checked = false; }
        }
        private async Task RestartAsync(bool resetFailures = true)
        {
            if (restarting || quitting) return;
            restarting = true; ++generation; ready = false; nextRestart = DateTime.MaxValue;
            if (resetFailures) failures = 0;
            if (host != null) { await host.StopAsync(); host = null; }
            restarting = false;
            if (!quitting) StartHost();
        }
        private async Task RenewAsync()
        {
            if (MessageBox.Show("Le code actuel cessera de fonctionner. Il faudra coller le nouveau code dans PowerPoint. Continuer ?", "Nouvelle liaison PowerPoint", MessageBoxButtons.OKCancel, MessageBoxIcon.Information) != DialogResult.OK) return;
            try { pairing = pairing ?? new PairingStore(directory); token = pairing.Renew(); await RestartAsync(); CopyCode(); }
            catch { SetStatus("Impossible de créer un nouveau code", failed); }
        }
        private async Task QuitAsync()
        {
            if (quitting) return;
            quitting = true; ++generation; timer.Stop(); tray.Visible = false;
            if (host != null) { await host.StopAsync(); host = null; }
            ExitThread();
        }
        private void SetStatus(string text, Icon icon)
        {
            status.Text = text;
            string tooltip = "Neurow.Pictos — " + text;
            tray.Text = tooltip.Length > 63 ? tooltip.Substring(0, 63) : tooltip;
            tray.Icon = icon; open.Enabled = ready; copy.Enabled = !string.IsNullOrEmpty(token);
        }
        internal void ReportFailure() { SetStatus("Action interrompue · réessayez depuis le menu", failed); }
        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                quitting = true; timer.Dispose(); host?.Dispose(); http.Dispose();
                tray.Visible = false; tray.Dispose(); menu.Dispose(); dispatcher.Dispose();
                waiting.Dispose(); connected.Dispose(); failed.Dispose();
            }
            base.Dispose(disposing);
        }
        private static string Value(Dictionary<string, object> value, string key) { return value != null && value.TryGetValue(key, out object result) ? result as string ?? "" : ""; }
        private static bool Flag(Dictionary<string, object> value, string key) { return value.TryGetValue(key, out object result) && result is bool yes && yes; }
        [DllImport("user32.dll")] private static extern bool DestroyIcon(IntPtr handle);
        private static Icon MakeIcon(Color color)
        {
            using (var bitmap = new Bitmap(32, 32))
            using (var graphics = Graphics.FromImage(bitmap))
            using (var brush = new SolidBrush(color))
            using (var pen = new Pen(Color.White, 3) { LineJoin = LineJoin.Round, StartCap = LineCap.Round, EndCap = LineCap.Round })
            {
                graphics.SmoothingMode = SmoothingMode.AntiAlias;
                graphics.FillEllipse(brush, 1, 1, 30, 30);
                graphics.DrawLine(pen, 11, 24, 11, 8);
                graphics.DrawArc(pen, 7, 8, 15, 11, 270, 180);
                graphics.DrawLine(pen, 11, 8, 14, 8); graphics.DrawLine(pen, 11, 19, 14, 19);
                IntPtr handle = bitmap.GetHicon();
                try { using (var icon = Icon.FromHandle(handle)) return (Icon)icon.Clone(); }
                finally { DestroyIcon(handle); }
            }
        }
    }
}
