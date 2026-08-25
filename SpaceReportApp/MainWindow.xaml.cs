using System;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Security.Principal;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Threading.Tasks;
using System.Windows;
using Microsoft.Web.WebView2.Core;

namespace SpaceReportApp
{
    /// <summary>
    /// A window hosting the UI, and a thin bridge to Get-SpaceReport.ps1.
    ///
    /// Deliberately thin: this app contains NO knowledge of what is safe to
    /// delete. Classification and every deletion rail live in the script, so
    /// there is one source of truth. The app asks the script what it found,
    /// and hands back a list of paths the user ticked; the script decides what
    /// it is willing to remove and reports anything it refused.
    /// </summary>
    public partial class MainWindow : Window
    {
        private string _scriptPath;
        private bool _busy;

        // Small settings file. Chromium restricts localStorage depending on the
        // page's origin, so preferences live on this side where they are certain
        // to persist rather than silently failing to save.
        private static string SettingsPath => Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "SpaceReport", "settings.json");

        private static bool GetSkipAdminPrompt()
        {
            try
            {
                if (!File.Exists(SettingsPath)) return false;
                using var doc = JsonDocument.Parse(File.ReadAllText(SettingsPath));
                return doc.RootElement.TryGetProperty("skipAdminPrompt", out var v) &&
                       v.ValueKind == JsonValueKind.True;
            }
            catch { return false; }
        }

        private static void SetSkipAdminPrompt(bool value)
        {
            try
            {
                Directory.CreateDirectory(Path.GetDirectoryName(SettingsPath));
                File.WriteAllText(SettingsPath,
                    "{\"skipAdminPrompt\":" + (value ? "true" : "false") + "}");
            }
            catch { }
        }

        private string _uiPath;

        public MainWindow()
        {
            InitializeComponent();
            PrepareRuntimeFiles();
            Loaded += OnLoaded;
        }

        /// <summary>
        /// Locates ui.html and Get-SpaceReport.ps1.
        ///
        /// A normal build copies both next to the exe; those are preferred so that
        /// editing either one and rebuilding takes effect immediately. A published
        /// single-file build has no loose files, so the embedded copies are
        /// unpacked into LocalAppData instead - which is what lets the download be
        /// a single executable rather than a folder people must keep together.
        /// </summary>
        private void PrepareRuntimeFiles()
        {
            string beside = AppContext.BaseDirectory;
            string localScript = Path.Combine(beside, "Get-SpaceReport.ps1");
            string localUi     = Path.Combine(beside, "ui.html");
            if (File.Exists(localScript) && File.Exists(localUi))
            {
                _scriptPath = localScript;
                _uiPath     = localUi;
                return;
            }

            string dir = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "SpaceReport", "runtime");
            Directory.CreateDirectory(dir);
            _scriptPath = Unpack("Get-SpaceReport.ps1", dir) ?? localScript;
            _uiPath     = Unpack("ui.html", dir) ?? localUi;
        }

        /// <summary>
        /// Writes an embedded resource to disk, skipping the write when the file
        /// is already byte-identical. That keeps upgrades correct while avoiding a
        /// write on every launch - and avoids fighting a second running instance
        /// over the same file.
        /// </summary>
        private static string Unpack(string resourceName, string dir)
        {
            string target = Path.Combine(dir, resourceName);
            try
            {
                var asm = typeof(MainWindow).Assembly;
                using var src = asm.GetManifestResourceStream(resourceName);
                if (src == null) return null;

                using var ms = new MemoryStream();
                src.CopyTo(ms);
                byte[] wanted = ms.ToArray();

                if (File.Exists(target))
                {
                    try
                    {
                        byte[] have = File.ReadAllBytes(target);
                        if (have.Length == wanted.Length &&
                            have.AsSpan().SequenceEqual(wanted)) return target;
                    }
                    catch { }
                }

                File.WriteAllBytes(target, wanted);
                return target;
            }
            catch
            {
                // If it is already there from a previous run, use it rather than failing.
                return File.Exists(target) ? target : null;
            }
        }

        private async void OnLoaded(object sender, RoutedEventArgs e)
        {
            try
            {
            // Keep the browser profile beside the user's data, not next to the exe,
            // so the app works from a read-only location. Elevated runs get their
            // own profile: a WebView2 user-data folder cannot be shared between two
            // live processes, and during "Run as admin" both briefly overlap.
            string udf = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "SpaceReport", IsElevated() ? "WebView2-admin" : "WebView2");
            Directory.CreateDirectory(udf);

            var env = await CoreWebView2Environment.CreateAsync(null, udf);
            await Web.EnsureCoreWebView2Async(env);

            var s = Web.CoreWebView2.Settings;
            s.AreDefaultContextMenusEnabled = false;
            s.IsStatusBarEnabled = false;
            s.AreDevToolsEnabled = true;   // handy if something misbehaves

            Web.CoreWebView2.WebMessageReceived += OnWebMessage;
            Web.CoreWebView2.NewWindowRequested += (o, a) => a.Handled = true;

            if (string.IsNullOrEmpty(_uiPath) || !File.Exists(_uiPath))
            {
                MessageBox.Show("The user interface could not be unpacked.\n\nExpected: " +
                    (_uiPath ?? "(not resolved)"),
                    "Space Report", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }
            Web.CoreWebView2.Navigate(new Uri(_uiPath).AbsoluteUri);

            Web.CoreWebView2.DOMContentLoaded += async (o, a) =>
            {
                await SendAsync(new { type = "ready",
                                      elevated = IsElevated(),
                                      skipAdminPrompt = GetSkipAdminPrompt(),
                                      script = _scriptPath,
                                      scriptFound = File.Exists(_scriptPath),
                                      auto = ReadAutoScanArgs() });
                await SendDrivesAsync();
            };
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "The embedded browser could not start.\n\n" +
                    ex.GetType().Name + ": " + ex.Message +
                    "\n\nThe WebView2 runtime is required. If this happened right after " +
                    "'Run as admin', close any other copy of Space Report and try again.",
                    "Space Report", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        // --- Recycle Bin ----------------------------------------------------
        // Recycled files still occupy the drive, so the app has to be able to
        // show what is sitting in there and let you go and empty it.
        [StructLayout(LayoutKind.Sequential)]
        private struct SHQUERYRBINFO
        {
            public int cbSize;
            public long i64Size;
            public long i64NumItems;
        }

        [DllImport("shell32.dll", CharSet = CharSet.Unicode)]
        private static extern int SHQueryRecycleBin(string pszRootPath, ref SHQUERYRBINFO info);

        private async Task SendRecycleAsync()
        {
            long size = 0, items = 0;
            try
            {
                var info = new SHQUERYRBINFO();
                info.cbSize = Marshal.SizeOf(typeof(SHQUERYRBINFO));
                // null = every drive's bin combined
                if (SHQueryRecycleBin(null, ref info) == 0)
                {
                    size  = info.i64Size;
                    items = info.i64NumItems;
                }
            }
            catch { }
            await SendAsync(new { type = "recycle", size, items });
        }

        private static void OpenRecycleBin()
        {
            try
            {
                Process.Start(new ProcessStartInfo("explorer.exe", "shell:RecycleBinFolder")
                    { UseShellExecute = true });
            }
            catch { }
        }

        /// <summary>
        /// Optional launch arguments, so a scan can be started without clicking:
        ///     SpaceReport.exe --scan C:\ --min 500 --view system
        /// --view accepts "system" or a verdict such as SAFE. Used for capturing
        /// documentation screenshots, and handy for a shortcut that always scans
        /// the same drive.
        /// </summary>
        private static object ReadAutoScanArgs()
        {
            try
            {
                var a = Environment.GetCommandLineArgs();
                string path = null, view = null;
                int min = 500;
                for (int i = 1; i < a.Length - 1; i++)
                {
                    if (a[i].Equals("--scan", StringComparison.OrdinalIgnoreCase)) path = a[i + 1];
                    else if (a[i].Equals("--min", StringComparison.OrdinalIgnoreCase))
                        int.TryParse(a[i + 1], out min);
                    else if (a[i].Equals("--view", StringComparison.OrdinalIgnoreCase)) view = a[i + 1];
                }
                if (path == null) return null;
                return new { path, minMB = min, view };
            }
            catch { return null; }
        }

        private static bool IsElevated()
        {
            try
            {
                using var id = WindowsIdentity.GetCurrent();
                return new WindowsPrincipal(id).IsInRole(WindowsBuiltInRole.Administrator);
            }
            catch { return false; }
        }

        private Task SendAsync(object payload)
        {
            string json = JsonSerializer.Serialize(payload);
            return Dispatcher.InvokeAsync(() =>
            {
                try { Web.CoreWebView2.PostWebMessageAsJson(json); } catch { }
            }).Task;
        }

        private async Task SendDrivesAsync()
        {
            var list = new JsonArray();
            foreach (var d in DriveInfo.GetDrives())
            {
                try
                {
                    if (d.DriveType != DriveType.Fixed || !d.IsReady) continue;
                    list.Add(new JsonObject
                    {
                        ["name"]  = d.Name,
                        ["label"] = string.IsNullOrWhiteSpace(d.VolumeLabel) ? d.Name : d.VolumeLabel,
                        ["size"]  = d.TotalSize,
                        ["free"]  = d.AvailableFreeSpace
                    });
                }
                catch { }
            }
            await SendAsync(new { type = "drives", drives = JsonSerializer.Deserialize<JsonElement>(list.ToJsonString()) });
        }

        private async void OnWebMessage(object sender, CoreWebView2WebMessageReceivedEventArgs e)
        {
            JsonElement msg;
            try
            {
                using var doc = JsonDocument.Parse(e.WebMessageAsJson);
                msg = doc.RootElement.Clone();   // Clone: RootElement dies with the document.

                // Accept a JSON string too, in case something posts a stringified
                // payload - unwrap it rather than throwing.
                if (msg.ValueKind == JsonValueKind.String)
                {
                    using var inner = JsonDocument.Parse(msg.GetString());
                    msg = inner.RootElement.Clone();
                }
            }
            catch { return; }

            if (msg.ValueKind != JsonValueKind.Object) return;
            string cmd = msg.TryGetProperty("cmd", out var c) && c.ValueKind == JsonValueKind.String
                       ? c.GetString() : null;
            if (cmd == null) return;

            switch (cmd)
            {
                case "scan":   await DoScanAsync(msg);   break;
                case "delete": await DoDeleteAsync(msg); break;
                case "drives": await SendDrivesAsync();  break;
                case "recycle": await SendRecycleAsync(); break;
                case "openrecycle":
                    OpenRecycleBin();
                    await SendRecycleAsync();
                    break;
                case "setskipadmin":
                    SetSkipAdminPrompt(msg.TryGetProperty("value", out var sv) &&
                                       sv.ValueKind == JsonValueKind.True);
                    break;
                case "title":
                    // Lets an external script tell when a scan has finished,
                    // by watching the window title rather than guessing a delay.
                    if (msg.TryGetProperty("text", out var tt) && tt.ValueKind == JsonValueKind.String)
                    {
                        string t = tt.GetString();
                        await Dispatcher.InvokeAsync(() => Title = t);
                    }
                    break;
                case "reveal": Reveal(msg);              break;
                case "elevate": Elevate();               break;
            }
        }

        private static void Reveal(JsonElement msg)
        {
            try
            {
                string p = msg.GetProperty("path").GetString();
                if (File.Exists(p))
                    Process.Start(new ProcessStartInfo("explorer.exe", "/select,\"" + p + "\"")
                        { UseShellExecute = true });
            }
            catch { }
        }

        private void Elevate()
        {
            try
            {
                string exe = Process.GetCurrentProcess().MainModule.FileName;
                Process.Start(new ProcessStartInfo(exe) { UseShellExecute = true, Verb = "runas" });
                Application.Current.Shutdown();
            }
            catch { /* user declined the UAC prompt */ }
        }

        private async Task DoScanAsync(JsonElement msg)
        {
            if (_busy) return;
            _busy = true;
            try
            {
                string path  = msg.TryGetProperty("path", out var p) ? p.GetString() : "C:\\";
                int    minMB = msg.TryGetProperty("minMB", out var m) && m.TryGetInt32(out var mv) ? mv : 500;
                if (minMB < 1) minMB = 1;

                string outFile = Path.Combine(Path.GetTempPath(),
                    "spacereport-scan-" + Guid.NewGuid().ToString("N") + ".json");

                // Bytes already used on the target drive is the denominator for the
                // progress bar. It only means anything when scanning a whole drive;
                // for a subfolder we report bytes seen and leave the bar indeterminate.
                long usedBytes = 0;
                bool wholeDrive = false;
                try
                {
                    string rootOf = Path.GetPathRoot(path);
                    wholeDrive = !string.IsNullOrEmpty(rootOf) &&
                                 string.Equals(Path.TrimEndingDirectorySeparator(path),
                                               Path.TrimEndingDirectorySeparator(rootOf),
                                               StringComparison.OrdinalIgnoreCase);
                    var di = new DriveInfo(rootOf);
                    usedBytes = di.TotalSize - di.AvailableFreeSpace;
                }
                catch { }

                var (ok, err) = await RunScriptAsync(new[]
                {
                    "-Path", path,
                    "-MinSizeMB", minMB.ToString(),
                    "-Top", "0",
                    "-TopFolders", "12",
                    "-Progress",
                    "-Json", outFile
                },
                line =>
                {
                    // "##P <bytes> <files> <folder>"
                    if (line.Length < 4 || line[0] != '#' || line[1] != '#' || line[2] != 'P') return;
                    var bits = line.Split('\t');
                    if (bits.Length < 4) return;
                    if (!long.TryParse(bits[1], out long seen)) return;
                    long.TryParse(bits[2], out long nfiles);

                    int pct = -1;
                    if (wholeDrive && usedBytes > 0)
                        pct = (int)Math.Min(99, seen * 100 / usedBytes);

                    _ = SendAsync(new { type = "progress", pct, seen, files = nfiles, folder = bits[3] });
                });

                if (!ok && !File.Exists(outFile))
                {
                    await SendAsync(new { type = "error", stage = "scan", message = err });
                    return;
                }

                string json = File.ReadAllText(outFile, Encoding.UTF8);
                try { File.Delete(outFile); } catch { }

                await SendAsync(new { type = "scan",
                                      data = JsonSerializer.Deserialize<JsonElement>(json) });
            }
            catch (Exception ex)
            {
                await SendAsync(new { type = "error", stage = "scan", message = ex.Message });
            }
            finally { _busy = false; }
        }

        private async Task DoDeleteAsync(JsonElement msg)
        {
            if (_busy) return;
            _busy = true;
            try
            {
                if (!msg.TryGetProperty("paths", out var arr) || arr.ValueKind != JsonValueKind.Array
                    || arr.GetArrayLength() == 0)
                {
                    await SendAsync(new { type = "error", stage = "delete", message = "No files selected." });
                    return;
                }

                bool permanent = msg.TryGetProperty("permanent", out var pm) &&
                                 pm.ValueKind == JsonValueKind.True;

                string outFile = Path.Combine(Path.GetTempPath(),
                    "spacereport-del-" + Guid.NewGuid().ToString("N") + ".json");

                var args = new System.Collections.Generic.List<string>();
                // -Force skips the script's own typed confirmation because the app
                // has already shown its own. Every other rail still applies: the
                // script re-classifies each path and refuses what it must.
                args.Add("-Force");
                if (permanent) args.Add("-Permanent");
                args.Add("-Json"); args.Add(outFile);
                args.Add("-Paths");
                foreach (var el in arr.EnumerateArray())
                {
                    string v = el.GetString();
                    if (!string.IsNullOrWhiteSpace(v)) args.Add(v);
                }

                var (ok, err) = await RunScriptAsync(args.ToArray());

                if (!File.Exists(outFile))
                {
                    await SendAsync(new { type = "error", stage = "delete",
                                          message = string.IsNullOrWhiteSpace(err)
                                              ? "The script produced no result file."
                                              : err });
                    return;
                }

                string json = File.ReadAllText(outFile, Encoding.UTF8);
                try { File.Delete(outFile); } catch { }

                await SendAsync(new { type = "deleted",
                                      data = JsonSerializer.Deserialize<JsonElement>(json) });
                await SendDrivesAsync();     // free space has moved
                await SendRecycleAsync();    // and the bin has grown
            }
            catch (Exception ex)
            {
                await SendAsync(new { type = "error", stage = "delete", message = ex.Message });
            }
            finally { _busy = false; }
        }

        /// <summary>
        /// Runs Get-SpaceReport.ps1 out of process. Arguments are passed as a real
        /// argument list (never concatenated into a command line), so paths with
        /// spaces, quotes or ampersands cannot alter what is executed.
        /// </summary>
        private async Task<(bool ok, string err)> RunScriptAsync(string[] scriptArgs,
                                                                Action<string> onLine = null)
        {
            if (!File.Exists(_scriptPath))
                return (false, "Get-SpaceReport.ps1 was not found next to the application.");

            var psi = new ProcessStartInfo
            {
                FileName               = ResolveShell(),
                UseShellExecute        = false,
                CreateNoWindow         = true,
                RedirectStandardOutput = true,
                RedirectStandardError  = true,
                StandardOutputEncoding = Encoding.UTF8,
                StandardErrorEncoding  = Encoding.UTF8
            };
            psi.ArgumentList.Add("-NoLogo");
            psi.ArgumentList.Add("-NonInteractive");
            psi.ArgumentList.Add("-ExecutionPolicy");
            psi.ArgumentList.Add("Bypass");
            psi.ArgumentList.Add("-File");
            psi.ArgumentList.Add(_scriptPath);
            foreach (var a in scriptArgs) psi.ArgumentList.Add(a);

            using var proc = new Process { StartInfo = psi };
            var stderr = new StringBuilder();

            // Read line by line rather than ReadToEnd: ReadToEnd only returns once
            // the process exits, which makes live progress impossible.
            proc.OutputDataReceived += (s, a) => { if (a.Data != null) onLine?.Invoke(a.Data); };
            proc.ErrorDataReceived  += (s, a) => { if (a.Data != null) stderr.AppendLine(a.Data); };

            proc.Start();
            proc.BeginOutputReadLine();
            proc.BeginErrorReadLine();
            await Task.Run(() => proc.WaitForExit());

            return (proc.ExitCode == 0, stderr.ToString().Trim());
        }

        private static string ResolveShell()
        {
            foreach (var candidate in new[]
            {
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles),
                             "PowerShell", "7", "pwsh.exe"),
                "pwsh.exe"
            })
            {
                if (File.Exists(candidate)) return candidate;
            }
            return "powershell.exe";   // Windows PowerShell fallback
        }
    }
}
