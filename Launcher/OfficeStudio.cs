using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Security.Cryptography;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Markup;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Threading;
using Microsoft.Win32;

[assembly: AssemblyTitle("Office Studio")]
[assembly: AssemblyDescription("Choose, download, and install Microsoft Office. Created and provided by KAREER.")]
[assembly: AssemblyCompany("KAREER")]
[assembly: AssemblyProduct("Office Studio")]
[assembly: AssemblyVersion("1.2.2.0")]

namespace Kareer.OfficeStudio
{
    public sealed class Plan
    {
        public string PlanId { get; set; }
        public string Branch { get; set; }
        public string Channel { get; set; }
        public string[] Components { get; set; }
        public string Architecture { get; set; }
        public string Language { get; set; }
        public string DownloadPath { get; set; }
    }
    public sealed class Preferences
    {
        public int Edition { get; set; }
        public string Folder { get; set; }
        public string Language { get; set; }
        public string Architecture { get; set; }
        public string[] Components { get; set; }
    }
    public sealed class Choice
    {
        public string Label { get; set; }
        public string Value { get; set; }
        public override string ToString() { return Label; }
    }
    public static class Program
    {
        [STAThread]
        public static int Main(string[] args)
        {
            bool test = args.Length == 2 && args[0] == "--self-test";
            bool inventoryTest = args.Length == 2 && args[0] == "--inventory-test";
            bool startupTest = args.Length == 2 && args[0] == "--startup-test";
            try
            {
                bool created;
                using (var mutex = new Mutex(true, "Local\\KAREER.OfficeStudio" + (test || inventoryTest || startupTest ? ".Test" : ""), out created))
                {
                    if (!created) { MessageBox.Show("Office Studio is already open.", "Office Studio"); return 1; }
                    var app = new Application();
                    var studio = new Studio(test || inventoryTest || startupTest, startupTest);
                    if (test) studio.Window.ContentRendered += delegate { studio.SelfTest(args[1]); };
                    if (inventoryTest) studio.Window.ContentRendered += async delegate { await studio.InventoryTest(args[1]); };
                    if (startupTest) studio.Window.ContentRendered += async delegate { await studio.StartupTest(args[1]); };
                    app.Run(studio.Window);
                    return studio.ExitCode;
                }
            }
            catch (Exception ex)
            {
                if (test || inventoryTest || startupTest) File.WriteAllText(Path.Combine(args[1], "fatal.txt"), ex.ToString());
                else MessageBox.Show("Office Studio couldn't start.\n\n" + ex.Message, "Office Studio", MessageBoxButton.OK, MessageBoxImage.Error);
                return 1;
            }
        }
    }
    public sealed partial class Studio
    {
        public Window Window { get; private set; }
        public int ExitCode { get; private set; }
        readonly JavaScriptSerializer json = new JavaScriptSerializer();
        readonly Dictionary<string, CheckBox> apps = new Dictionary<string, CheckBox>();
        readonly string[] branches = { "ProPlus2019Retail", "ProPlus2021Volume", "ProPlus2024Volume", "O365ProPlusRetail" };
        readonly string[] channels = { "Current", "PerpetualVL2021", "PerpetualVL2024", "Current" };
        readonly string[] editions = { "Office 2019", "Office 2021", "Office 2024", "Microsoft 365" };
        readonly string[] notes = {
            "Professional Plus Retail · Current channel. Office 2019 reached end of support in October 2025.",
            "Professional Plus LTSC · PerpetualVL2021. Requires a volume license; support ends October 2026.",
            "Professional Plus LTSC · PerpetualVL2024. Requires a volume license. Publisher is unavailable.",
            "Microsoft 365 Apps for enterprise · Current channel. Requires a qualifying subscription. Publisher is retired."
        };
        readonly bool testMode;
        readonly bool startupProbe;
        Task savedSizeRefresh;
        readonly string stateRoot = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "KAREER", "Office Studio");
        readonly StringBuilder activity = new StringBuilder();
        RadioButton[] profiles;
        bool initialized, busy, ready, updating, licenseReset;
        string baseFolder, activeLog, runtimeDirectory, lastError;
        bool workerCompleted;
        Process worker;
        DownloadJob downloadJob;
        string activeAction;
        bool cancelRequested;
        int cacheSizeRequest;
        Window logWindow;
        TextBox logBox;
        DispatcherTimer timer;
        Stopwatch clock = new Stopwatch();
        int edition;

        T Find<T>(string name) where T : class { return Window.FindName(name) as T; }
        TextBlock Status { get { return Find<TextBlock>("StatusTitle"); } }
        TextBlock Detail { get { return Find<TextBlock>("StatusDetail"); } }
        ProgressBar Progress { get { return Find<ProgressBar>("ProgressBar"); } }
        Button Download { get { return Find<Button>("DownloadButton"); } }
        Button Install { get { return Find<Button>("InstallButton"); } }
        CheckBox License { get { return Find<CheckBox>("LicenseCheck"); } }
        ComboBox Language { get { return Find<ComboBox>("LanguageCombo"); } }
        ComboBox Architecture { get { return Find<ComboBox>("ArchitectureCombo"); } }

        public Studio(bool test, bool probeStartup = false)
        {
            testMode = test;
            startupProbe = probeStartup;
            using (Stream stream = Assembly.GetExecutingAssembly().GetManifestResourceStream("Kareer.MainWindow.xaml"))
                Window = (Window)XamlReader.Load(stream);
            using (Stream stream = Assembly.GetExecutingAssembly().GetManifestResourceStream("Kareer.ico"))
                if (stream != null) Window.Icon = BitmapFrame.Create(stream, BitmapCreateOptions.None, BitmapCacheOption.OnLoad);
            if (!testMode)
            {
                Window.Height = Math.Min(Window.Height, SystemParameters.WorkArea.Height - 20);
                Window.Width = Math.Min(Window.Width, SystemParameters.WorkArea.Width - 20);
                Window.MinHeight = Math.Min(Window.MinHeight, Window.Height);
                Window.MinWidth = Math.Min(Window.MinWidth, Window.Width);
            }
            profiles = new[] { Find<RadioButton>("Profile2019"), Find<RadioButton>("Profile2021"), Find<RadioButton>("Profile2024"), Find<RadioButton>("Profile365") };
            AddApp("Word", "W", "Documents", "#2366BC");
            AddApp("Excel", "X", "Spreadsheets", "#21824C");
            AddApp("PowerPoint", "P", "Presentations", "#CE5839");
            AddApp("Outlook", "O", "Email & calendar", "#1685BC");
            AddApp("OneNote", "N", "Notes & ideas", "#8651BB");
            AddApp("Access", "A", "Databases", "#B44657");
            AddApp("OneDrive", "D", "Cloud files · add-on", "#2B90CC");
            AddApp("Teams", "T", "Meetings · add-on", "#6966AE");
            AddApp("Publisher", "P", "Page layouts", "#258783");
            string[] languageValues = { "en-us", "en-gb", "fr-fr", "es-es", "de-de", "it-it", "pt-br", "hi-in" };
            string[] languageLabels = { "English (United States)", "English (United Kingdom)", "French", "Spanish", "German", "Italian", "Portuguese (Brazil)", "Hindi" };
            for (int i = 0; i < languageValues.Length; i++) Language.Items.Add(new Choice { Label = languageLabels[i], Value = languageValues[i] });
            Language.SelectedIndex = 0;
            if (Environment.Is64BitOperatingSystem) Architecture.Items.Add(new Choice { Label = "64-bit (recommended)", Value = "64" });
            Architecture.Items.Add(new Choice { Label = "32-bit (legacy)", Value = "32" });
            Architecture.SelectedIndex = 0;
            baseFolder = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Office Downloads");
            apps["Word"].IsChecked = true;
            apps["Excel"].IsChecked = true;
            if (!testMode) LoadPreferences();
            for (int i = 0; i < profiles.Length; i++)
            {
                int value = i;
                profiles[i].Checked += delegate { edition = value; SelectionChanged(); };
            }
            Language.SelectionChanged += delegate { SelectionChanged(); };
            Architecture.SelectionChanged += delegate { SelectionChanged(); };
            Find<Button>("SelectAllButton").Click += delegate { SelectApps(true); };
            Find<Button>("ClearButton").Click += delegate { SelectApps(false); };
            Find<Button>("BrowseButton").Click += delegate { Browse(); };
            Find<Button>("OpenFolderButton").Click += delegate { OpenFolder(); };
            Find<Button>("LogsButton").Click += delegate { ShowActivity(); };
            Download.Click += async delegate { await RunWorker("Download"); };
            Install.Click += async delegate { await RunWorker("Install"); };
            Find<Button>("AbortButton").Click += delegate { AbortDownload(); };
            License.Checked += delegate { UpdateButtons(); };
            License.Unchecked += delegate { UpdateButtons(); };
            Find<System.Windows.Documents.Hyperlink>("LicenseTerms").RequestNavigate += delegate(object sender, System.Windows.Navigation.RequestNavigateEventArgs e)
            {
                try { Process.Start(new ProcessStartInfo(e.Uri.AbsoluteUri) { UseShellExecute = true }); }
                catch (Exception ex) { Log("Couldn't open license terms: " + ex.Message); }
                e.Handled = true;
            };
            Window.Closing += delegate(object sender, System.ComponentModel.CancelEventArgs e)
            {
                if (busy && activeAction == "Download")
                {
                    e.Cancel = true;
                    if (!cancelRequested && MessageBox.Show(Window, "Abort this download? Partial Office files will be kept for a retry. You can close the app once it has stopped.", "Abort download", MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.Yes) AbortDownload();
                }
                else if (busy || maintenanceRunning) { e.Cancel = true; MessageBox.Show(Window, "Setup is still running. Keep Office Studio open until the current operation finishes.\n\nFor installation, use Cancel in Microsoft's setup window if offered. You can see progress in Activity.", "Operation in progress", MessageBoxButton.OK, MessageBoxImage.Information); }
                else { SavePreferences(); if (logWindow != null) logWindow.Close(); if (maintenanceWindow != null) maintenanceWindow.Close(); CleanRuntime(); }
            };
            timer = new DispatcherTimer { Interval = TimeSpan.FromSeconds(1) };
            timer.Tick += delegate { if (busy) Find<TextBlock>("ProgressLabel").Text = clock.Elapsed.ToString(@"hh\:mm\:ss"); };
            initialized = true;
            SelectionChanged();
            Log("Created and provided by KAREER. Select your edition and apps to begin.");
            InitializeMaintenance();
        }

        void AddApp(string name, string letter, string subtitle, string color)
        {
            var content = new StackPanel { Orientation = Orientation.Horizontal };
            var icon = new Border { Width = 33, Height = 35, CornerRadius = new CornerRadius(7), Background = (Brush)new BrushConverter().ConvertFrom(color), Margin = new Thickness(0, 0, 10, 0) };
            icon.Child = new TextBlock { Text = letter, Foreground = Brushes.White, FontSize = 19, FontWeight = FontWeights.SemiBold, HorizontalAlignment = HorizontalAlignment.Center, VerticalAlignment = VerticalAlignment.Center };
            content.Children.Add(icon);
            var labels = new StackPanel { VerticalAlignment = VerticalAlignment.Center };
            labels.Children.Add(new TextBlock { Text = name, FontSize = 13, FontWeight = FontWeights.SemiBold });
            labels.Children.Add(new TextBlock { Text = subtitle, FontSize = 9, Foreground = (Brush)Window.FindResource("Muted"), Margin = new Thickness(0, 3, 0, 0) });
            content.Children.Add(labels);
            var check = new CheckBox { Content = content, Style = (Style)Window.FindResource("AppTile"), ToolTip = name + " — " + subtitle };
            System.Windows.Automation.AutomationProperties.SetName(check, name);
            check.Checked += delegate { SelectionChanged(); };
            check.Unchecked += delegate { SelectionChanged(); };
            apps.Add(name, check);
            Find<System.Windows.Controls.Primitives.UniformGrid>("AppsPanel").Children.Add(check);
        }

        void SelectApps(bool value)
        {
            updating = true;
            foreach (var app in apps.Values) app.IsChecked = value && app.IsEnabled;
            updating = false;
            SelectionChanged();
        }

        void SelectionChanged()
        {
            if (!initialized || updating || busy) return;
            updating = true;
            apps["Publisher"].IsEnabled = edition < 2;
            if (edition >= 2) apps["Publisher"].IsChecked = false;
            updating = false;
            Find<TextBlock>("EditionNote").Text = notes[edition];
            Find<TextBox>("FolderBox").Text = baseFolder;
            Find<TextBox>("FolderBox").ToolTip = MakePlan().DownloadPath;
            // Consent is tied to the chosen package, not remembered across changes/restarts.
            if (!licenseReset) { licenseReset = true; License.IsChecked = false; licenseReset = false; }
            ready = LooksReady(MakePlan());
            Progress.IsIndeterminate = false;
            Progress.Value = ready ? 100 : 0;
            Find<TextBlock>("ProgressLabel").Text = ready ? "DOWNLOADED" : "";
            Status.Text = ready ? "Your apps are ready to install" : "Ready to download";
            Detail.Text = ready ? "Saved download found. Accept the license terms, then choose Install Office." : Summary();
            if (!HasSuiteApp()) Detail.Text = "Choose at least one Office app. OneDrive and Teams are optional add-ons.";
            UpdateButtons();
            RefreshSavedSize();
            SavePreferences();
        }

        bool HasSuiteApp() { return apps.Any(p => p.Value.IsChecked == true && p.Key != "OneDrive" && p.Key != "Teams"); }
        void UpdateButtons()
        {
            Download.IsEnabled = !busy && !maintenanceRunning && HasSuiteApp();
            Install.IsEnabled = !busy && !maintenanceRunning && ready && License.IsChecked == true && HasSuiteApp();
            Download.Content = "Download";
            var abort = Find<Button>("AbortButton");
            abort.Visibility = busy && activeAction == "Download" ? Visibility.Visible : Visibility.Collapsed;
            abort.IsEnabled = downloadJob != null && !cancelRequested;
            abort.Content = cancelRequested ? "Stopping…" : "Abort download";
            int step = busy ? (activeAction == "Install" ? 3 : 2) : (ready ? 3 : 1);
            string[] steps = { "StepChoose", "StepDownload", "StepInstall" };
            for (int i = 0; i < steps.Length; i++)
            {
                var label = Find<TextBlock>(steps[i]);
                label.Foreground = (Brush)new BrushConverter().ConvertFrom(i + 1 == step ? "#7DE2BB" : "#B4C9CF");
                label.FontWeight = i + 1 == step ? FontWeights.SemiBold : FontWeights.Normal;
            }
            UpdateMaintenanceButtons();
        }
        void RefreshSavedSize() { savedSizeRefresh = RefreshSavedSizeAsync(); }
        async Task RefreshSavedSizeAsync()
        {
            int request = ++cacheSizeRequest;
            Plan plan = MakePlan();
            var label = Find<TextBlock>("SavedSize");
            label.Text = "Checking existing files…";
            string message;
            try
            {
                long bytes = testMode && !startupProbe ? SavedOfficeBytes(plan) : await Task.Run(delegate { return SavedOfficeBytes(plan); }).ConfigureAwait(false);
                message = "Existing files: " + (bytes / 1048576.0).ToString("N1") + " MB on disk";
            }
            catch { message = "Existing file size unavailable"; }
            // The first call runs in the constructor, before Application.Run has
            // installed a WPF synchronization context. Always dispatch UI writes,
            // including the error path, instead of trusting an await continuation.
            Action update = delegate { if (request == cacheSizeRequest) label.Text = message; };
            if (Window.Dispatcher.CheckAccess()) update();
            else if (!Window.Dispatcher.HasShutdownStarted)
            {
                try { await Window.Dispatcher.InvokeAsync(update); }
                catch (TaskCanceledException) { /* Window closed before the queued update ran. */ }
            }
        }
        static long SavedOfficeBytes(Plan plan)
        {
            long bytes = FolderBytes(Path.Combine(plan.DownloadPath, "Office"));
            if (!Directory.Exists(plan.DownloadPath)) return bytes;
            foreach (var stage in new DirectoryInfo(plan.DownloadPath).GetDirectories(".kareer-download-*").OrderByDescending(d => d.LastWriteTimeUtc))
            {
                if (!System.Text.RegularExpressions.Regex.IsMatch(stage.Name, @"^\.kareer-download-[a-f0-9]{32}$") || (stage.Attributes & FileAttributes.ReparsePoint) != 0) continue;
                string config = Path.Combine(stage.FullName, "Config.xml");
                if (!File.Exists(config) || (File.GetAttributes(config) & FileAttributes.ReparsePoint) != 0) continue;
                bool matches = false;
                try
                {
                    var document = new System.Xml.XmlDocument();
                    document.XmlResolver = null;
                    using (var reader = System.Xml.XmlReader.Create(config, new System.Xml.XmlReaderSettings { DtdProcessing = System.Xml.DtdProcessing.Prohibit, XmlResolver = null })) document.Load(reader);
                    var add = document.SelectSingleNode("/Configuration/Add") as System.Xml.XmlElement;
                    var products = document.SelectNodes("/Configuration/Add/Product");
                    var languages = document.SelectNodes("/Configuration/Add/Product/Language");
                    matches = add != null && products.Count == 1 && languages.Count == 1 &&
                        string.Equals(add.GetAttribute("SourcePath"), stage.FullName, StringComparison.OrdinalIgnoreCase) &&
                        add.GetAttribute("OfficeClientEdition") == plan.Architecture && add.GetAttribute("Channel") == plan.Channel &&
                        ((System.Xml.XmlElement)products[0]).GetAttribute("ID") == plan.Branch && ((System.Xml.XmlElement)languages[0]).GetAttribute("ID") == plan.Language;
                }
                catch (System.Xml.XmlException) { }
                if (matches) { bytes += FolderBytes(Path.Combine(stage.FullName, "Office")); break; }
            }
            return bytes;
        }
        public async Task StartupTest(string outputDirectory)
        {
            Directory.CreateDirectory(outputDirectory);
            var results = new List<string>();
            try
            {
                // This task was started by the constructor, before Application.Run,
                // using the same asynchronous file scan as a normal launch.
                await savedSizeRefresh;
                string text = Find<TextBlock>("SavedSize").Text;
                if (!text.StartsWith("Existing files: ") || !text.EndsWith(" MB on disk")) throw new Exception("Startup file-size update failed: " + text);
                results.Add("PASS: Background file scan started before the UI event loop safely updates the window.");
                results.Add("PASS: " + text);
                if (Convert.ToString(Download.Content) != "Download") throw new Exception("Download button label changed.");
                results.Add("PASS: Download button remains Download.");
                Capture(Path.Combine(outputDirectory, "09-startup.png"));
                ExitCode = 0;
            }
            catch (Exception ex) { results.Add(ex.ToString()); ExitCode = 1; }
            finally { File.WriteAllLines(Path.Combine(outputDirectory, "startup-test-results.txt"), results); Window.Close(); }
        }
        static long FolderBytes(string path)
        {
            if (!Directory.Exists(path)) return 0;
            // Avoid traversing linked cache folders, including linked ancestors.
            for (var parent = new DirectoryInfo(path); parent != null; parent = parent.Parent)
                if ((parent.Attributes & FileAttributes.ReparsePoint) != 0) return 0;
            long bytes = 0;
            var folders = new Queue<DirectoryInfo>(); folders.Enqueue(new DirectoryInfo(path));
            while (folders.Count > 0)
                foreach (var entry in folders.Dequeue().GetFileSystemInfos())
                {
                    if ((entry.Attributes & FileAttributes.ReparsePoint) != 0) continue;
                    var directory = entry as DirectoryInfo;
                    if (directory != null) folders.Enqueue(directory);
                    else { var file = (FileInfo)entry; file.Refresh(); bytes += file.Length; }
                }
            return bytes;
        }
        void AbortDownload()
        {
            if (!busy || activeAction != "Download" || downloadJob == null || cancelRequested) return;
            try
            {
                downloadJob.Abort();
                cancelRequested = true;
                Status.Text = "Stopping your download…";
                Detail.Text = "Stopping this download's processes. Keeping partial files for a retry.";
                Log("Abort requested for this download.");
                UpdateButtons();
            }
            catch (Exception ex) { Log("Could not abort: " + ex.Message); ShowActivity(); }
        }
        string Summary() { return string.Join(" + ", apps.Where(p => p.Value.IsChecked == true).Select(p => p.Key)) + " · " + editions[edition]; }
        Plan MakePlan()
        {
            string architecture = ((Choice)Architecture.SelectedItem).Value;
            string language = ((Choice)Language.SelectedItem).Value;
            var plan = new Plan { Branch = branches[edition], Channel = channels[edition], Components = apps.Where(p => p.Value.IsChecked == true).Select(p => p.Key).OrderBy(s => s, StringComparer.Ordinal).ToArray(), Architecture = architecture, Language = language, DownloadPath = Path.GetFullPath(Path.Combine(baseFolder, branches[edition] + "-x" + architecture + "-" + language)) };
            string identity = plan.Branch + "|" + plan.Channel + "|" + plan.Architecture + "|" + plan.Language + "|" + plan.DownloadPath.TrimEnd('\\').ToUpperInvariant() + "|" + string.Join(",", plan.Components);
            using (var sha = SHA256.Create()) plan.PlanId = BitConverter.ToString(sha.ComputeHash(Encoding.UTF8.GetBytes(identity))).Replace("-", "").ToLowerInvariant();
            return plan;
        }
        bool LooksReady(Plan plan)
        {
            try
            {
                string marker = Path.Combine(plan.DownloadPath, "ready.json");
                if (!File.Exists(marker) || !File.Exists(Path.Combine(plan.DownloadPath, "Config.xml")) || !File.Exists(Path.Combine(plan.DownloadPath, "setup.exe"))) return false;
                var data = json.Deserialize<Dictionary<string, object>>(File.ReadAllText(marker));
                if (!data.ContainsKey("PlanId") || Convert.ToString(data["PlanId"]) != plan.PlanId) return false;
                string version = data.ContainsKey("OfficeVersion") ? Convert.ToString(data["OfficeVersion"]) : "";
                if (!System.Text.RegularExpressions.Regex.IsMatch(version, @"^16\.0\.\d+\.\d+$")) return false;
                string office = Path.Combine(plan.DownloadPath, "Office", "Data", version);
                string platform = plan.Architecture == "64" ? "x64" : "x86";
                return File.Exists(Path.Combine(office, "stream." + platform + ".x-none.dat")) && File.Exists(Path.Combine(office, "stream." + platform + "." + plan.Language + ".dat"));
            }
            catch { return false; }
        }
        void Browse()
        {
            using (var dialog = new System.Windows.Forms.FolderBrowserDialog { Description = "Choose where Office Studio saves downloaded packages", SelectedPath = baseFolder, ShowNewFolderButton = true })
            {
                if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK) { baseFolder = dialog.SelectedPath; SelectionChanged(); }
            }
        }
        void OpenFolder()
        {
            try
            {
                string folder = MakePlan().DownloadPath;
                if (!Directory.Exists(folder)) folder = Directory.Exists(baseFolder) ? baseFolder : AppDomain.CurrentDomain.BaseDirectory;
                Process.Start(new ProcessStartInfo("explorer.exe", Quote(folder)) { UseShellExecute = true });
            }
            catch (Exception ex) { Log("Couldn't open folder: " + ex.Message); }
        }

        async Task RunWorker(string action)
        {
            if (busy || maintenanceRunning || !HasSuiteApp() || (action == "Install" && (!ready || License.IsChecked != true))) return;
            if (testMode) throw new InvalidOperationException("Real deployment is disabled in self-test mode.");
            if (action == "Install" && !await CheckBeforeInstall()) return;
            var plan = MakePlan();
            busy = true; workerCompleted = false; lastError = null; lastCompletion = null; Detail.ToolTip = null;
            activeAction = action; cancelRequested = false;
            ready = false;
            Find<StackPanel>("OptionsPanel").IsEnabled = false;
            License.IsEnabled = false;
            UpdateButtons();
            Progress.IsIndeterminate = true;
            Status.Text = action == "Download" ? "Preparing your download…" : "Installing your Office…";
            Detail.Text = action == "Download" ? "Connecting to Microsoft. Watch the activity log for details." : "Verifying your saved package. Microsoft setup will show installation progress.";
            clock.Restart(); timer.Start();
            try
            {
                Directory.CreateDirectory(Path.Combine(stateRoot, "Logs"));
                activeLog = Path.Combine(stateRoot, "Logs", DateTime.Now.ToString("yyyyMMdd-HHmmss-fff") + "-" + action.ToLowerInvariant() + ".log");
                File.WriteAllText(activeLog, activity.ToString(), Encoding.UTF8);
                CleanRuntime();
                runtimeDirectory = Path.Combine(stateRoot, "Runtime", Guid.NewGuid().ToString("N"));
                Directory.CreateDirectory(runtimeDirectory);
                string script = Path.Combine(runtimeDirectory, "DeploymentWorker.ps1");
                using (var resource = Assembly.GetExecutingAssembly().GetManifestResourceStream("Kareer.DeploymentWorker.ps1"))
                using (var file = File.Create(script)) resource.CopyTo(file);
                string planPath = Path.Combine(runtimeDirectory, "plan.json");
                File.WriteAllText(planPath, json.Serialize(plan), new UTF8Encoding(false));
                Log(action + " requested: " + Summary());
                Log("Package folder: " + plan.DownloadPath);
                string systemFolder = Environment.GetFolderPath(Environment.SpecialFolder.System);
                // AnyCPU prefers 64-bit. Sysnative also supports a 32-bit host on 64-bit Windows.
                if (Environment.Is64BitOperatingSystem && !Environment.Is64BitProcess) systemFolder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Sysnative");
                string startSignal = Path.Combine(runtimeDirectory, "start.signal");
                if (action == "Download") downloadJob = new DownloadJob();
                var start = new ProcessStartInfo(Path.Combine(systemFolder, "WindowsPowerShell", "v1.0", "powershell.exe"), "-NoLogo -NoProfile -NonInteractive -ExecutionPolicy Bypass -File " + Quote(script) + " -Action " + action + " -PlanPath " + Quote(planPath) + (action == "Download" ? " -StartSignal " + Quote(startSignal) : ""))
                {
                    UseShellExecute = false, CreateNoWindow = true, WindowStyle = ProcessWindowStyle.Hidden,
                    RedirectStandardOutput = true, RedirectStandardError = true, StandardOutputEncoding = Encoding.UTF8, StandardErrorEncoding = Encoding.UTF8,
                    WorkingDirectory = runtimeDirectory
                };
                worker = new Process { StartInfo = start };
                worker.OutputDataReceived += delegate(object sender, DataReceivedEventArgs e) { if (e.Data != null) PostOutput(e.Data, false); };
                worker.ErrorDataReceived += delegate(object sender, DataReceivedEventArgs e) { if (e.Data != null) PostOutput(e.Data, true); };
                worker.Start(); worker.BeginOutputReadLine(); worker.BeginErrorReadLine();
                if (downloadJob != null)
                {
                    try { downloadJob.Assign(worker); File.WriteAllText(startSignal, "start"); }
                    catch { if (!worker.HasExited) { worker.Kill(); worker.WaitForExit(); } throw; }
                }
                UpdateButtons();
                await Task.Run(delegate { worker.WaitForExit(); });
                int code = worker.ExitCode;
                // Drain the queued output callbacks before deciding whether the worker completed.
                await Window.Dispatcher.InvokeAsync(delegate { }, DispatcherPriority.Background);
                if (cancelRequested && (code != 0 || !workerCompleted)) throw new OperationCanceledException();
                if (code != 0 || !workerCompleted) throw new InvalidOperationException(lastError ?? "Microsoft setup did not finish successfully (exit code " + code + "). Open Activity for details and retry.");
                Progress.IsIndeterminate = false; Progress.Value = 100;
                if (action == "Download")
                {
                    ready = LooksReady(plan);
                    if (!ready) throw new InvalidOperationException("The download finished without a matching complete package. Open Activity, then retry Download.");
                    Status.Text = "Your apps are ready to install";
                    Detail.Text = "Download complete. Accept the license terms, then choose Install Office.";
                    if (!string.IsNullOrEmpty(lastCompletion)) Detail.Text = lastCompletion;
                    Find<TextBlock>("ProgressLabel").Text = "DOWNLOADED";
                }
                else
                {
                    Status.Text = "Office setup completed";
                    Detail.Text = "Your selected apps are installed. Open Office and sign in or activate with your license.";
                    Find<TextBlock>("ProgressLabel").Text = "COMPLETE";
                    if (!string.IsNullOrEmpty(lastCompletion)) Detail.Text = lastCompletion;
                }
                Log(action + " completed successfully.");
            }
            catch (OperationCanceledException)
            {
                Progress.IsIndeterminate = false; Progress.Value = 0; ready = false;
                Status.Text = "Download aborted";
                Detail.Text = "Partial files kept. Choose Download with the same edition, language and folder to retry.";
                Find<TextBlock>("ProgressLabel").Text = "STOPPED";
                Log("Download aborted. No installation was started.");
            }
            catch (Exception ex)
            {
                Progress.IsIndeterminate = false; Progress.Value = 0;
                Status.Text = action == "Install" ? "Installation needs attention" : "Download needs attention";
                Detail.Text = ex.Message;
                Detail.ToolTip = ex.Message;
                Find<TextBlock>("ProgressLabel").Text = "SEE ACTIVITY";
                Log("ERROR: " + ex.Message);
                // Let a failed installation retry only if the original package still appears ready.
                ready = action == "Install" && LooksReady(plan);
                ShowActivity();
            }
            finally
            {
                timer.Stop(); clock.Stop(); busy = false;
                if (downloadJob != null) { downloadJob.Dispose(); downloadJob = null; }
                activeAction = null;
                if (worker != null) { worker.Dispose(); worker = null; }
                Find<StackPanel>("OptionsPanel").IsEnabled = true;
                License.IsEnabled = true;
                UpdateButtons();
                RefreshSavedSize();
                SavePreferences();
                CleanRuntime();
            }
            if (action == "Install") await RefreshOfficeInventory();
        }
        string lastCompletion;
        static string Quote(string text) { return "\"" + text.Replace("\"", "") + "\""; }
        void PostOutput(string line, bool error)
        {
            Window.Dispatcher.BeginInvoke(new Action(delegate { HandleOutput(line, error); }));
        }
        void HandleOutput(string line, bool error)
        {
            if (line.StartsWith("@@", StringComparison.Ordinal))
            {
                try
                {
                    var data = json.Deserialize<Dictionary<string, object>>(line.Substring(2));
                    string kind = data.ContainsKey("kind") ? Convert.ToString(data["kind"]) : "log";
                    string message = data.ContainsKey("message") ? Convert.ToString(data["message"]) : "";
                    HandleMaintenanceEvent(kind, message, data);
                    if (!inventoryScanning && !cancelRequested && (kind == "status" || kind == "progress"))
                    {
                        Detail.Text = message;
                        if (busy && activeAction == "Download") Status.Text = message.StartsWith("Checking", StringComparison.OrdinalIgnoreCase) || message.StartsWith("Verifying file", StringComparison.OrdinalIgnoreCase) ? "Checking your downloaded files…" : "Downloading your Office…";
                    }
                    if (!inventoryScanning && !cancelRequested && kind == "progress" && data.ContainsKey("percent"))
                    {
                        double value = Convert.ToDouble(data["percent"]);
                        Progress.IsIndeterminate = value < 0;
                        if (value >= 0) Progress.Value = Math.Max(0, Math.Min(100, value));
                    }
                    if (kind == "error") lastError = message;
                    if (kind == "complete") { workerCompleted = true; lastCompletion = message; }
                    Log(message);
                    return;
                }
                catch { /* Plain logging preserves malformed output for troubleshooting. */ }
            }
            Log((error ? "PowerShell: " : "") + line);
        }
        void Log(string message)
        {
            string line = "[" + DateTime.Now.ToString("HH:mm:ss") + "] " + message + Environment.NewLine;
            activity.Append(line);
            // Keep the UI bounded; full activity remains on disk.
            if (activity.Length > 180000) activity.Remove(0, activity.Length - 140000);
            if (activeLog != null) { try { File.AppendAllText(activeLog, line, Encoding.UTF8); } catch { } }
            if (logBox != null) { logBox.Text = activity.ToString(); logBox.ScrollToEnd(); }
        }
        void ShowActivity()
        {
            if (logWindow != null) { logWindow.Activate(); return; }
            logWindow = new Window { Owner = Window, Title = "KAREER | Activity", Width = 850, Height = 490, MinWidth = 600, MinHeight = 300, WindowStartupLocation = WindowStartupLocation.CenterOwner, Background = new SolidColorBrush(Color.FromRgb(18, 42, 49)) };
            var grid = new Grid { Margin = new Thickness(18) };
            grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            grid.RowDefinitions.Add(new RowDefinition());
            grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            grid.Children.Add(new TextBlock { Text = "LIVE ACTIVITY", FontFamily = new FontFamily("Segoe UI"), Foreground = new SolidColorBrush(Color.FromRgb(125, 226, 187)), FontWeight = FontWeights.Bold, Margin = new Thickness(0, 0, 0, 12) });
            logBox = new TextBox { Text = activity.ToString(), IsReadOnly = true, TextWrapping = TextWrapping.Wrap, VerticalScrollBarVisibility = ScrollBarVisibility.Auto, Background = new SolidColorBrush(Color.FromRgb(18, 42, 49)), Foreground = new SolidColorBrush(Color.FromRgb(215, 230, 232)), BorderThickness = new Thickness(0), FontFamily = new FontFamily("Consolas"), FontSize = 12 };
            Grid.SetRow(logBox, 1); grid.Children.Add(logBox);
            var save = new Button { Content = "Save activity log", HorizontalAlignment = HorizontalAlignment.Right, Margin = new Thickness(0, 14, 0, 0), Padding = new Thickness(15, 7, 15, 7) };
            save.Click += delegate
            {
                var dialog = new SaveFileDialog { FileName = "KAREER-Office-activity.log", Filter = "Log file|*.log|Text file|*.txt" };
                if (dialog.ShowDialog(logWindow) == true)
                {
                    try { File.WriteAllText(dialog.FileName, activeLog != null && File.Exists(activeLog) ? File.ReadAllText(activeLog) : activity.ToString()); }
                    catch (Exception ex) { MessageBox.Show(logWindow, ex.Message, "Could not save log"); }
                }
            };
            Grid.SetRow(save, 2); grid.Children.Add(save); logWindow.Content = grid;
            logWindow.Closed += delegate { logWindow = null; logBox = null; };
            logWindow.Show(); logBox.ScrollToEnd();
        }
        void LoadPreferences()
        {
            try
            {
                string file = Path.Combine(stateRoot, "preferences.json");
                if (!File.Exists(file)) return;
                var pref = json.Deserialize<Preferences>(File.ReadAllText(file));
                if (pref.Edition >= 0 && pref.Edition < profiles.Length) { edition = pref.Edition; profiles[edition].IsChecked = true; }
                if (!string.IsNullOrWhiteSpace(pref.Folder) && Path.IsPathRooted(pref.Folder)) baseFolder = Path.GetFullPath(pref.Folder);
                foreach (Choice item in Language.Items) if (item.Value == pref.Language) Language.SelectedItem = item;
                foreach (Choice item in Architecture.Items) if (item.Value == pref.Architecture) Architecture.SelectedItem = item;
                if (pref.Components != null) foreach (var pair in apps) pair.Value.IsChecked = pref.Components.Contains(pair.Key);
            }
            catch { /* A corrupt preference file must not prevent setup. */ }
        }
        void SavePreferences()
        {
            if (testMode || !initialized) return;
            try
            {
                Directory.CreateDirectory(stateRoot);
                var pref = new Preferences { Edition = edition, Folder = baseFolder, Language = ((Choice)Language.SelectedItem).Value, Architecture = ((Choice)Architecture.SelectedItem).Value, Components = apps.Where(p => p.Value.IsChecked == true).Select(p => p.Key).ToArray() };
                File.WriteAllText(Path.Combine(stateRoot, "preferences.json"), json.Serialize(pref));
            }
            catch { }
        }
        void CleanRuntime()
        {
            if (runtimeDirectory == null) return;
            try
            {
                // Only remove files explicitly created by this instance. Never recurse through user folders.
                File.Delete(Path.Combine(runtimeDirectory, "DeploymentWorker.ps1"));
                File.Delete(Path.Combine(runtimeDirectory, "OfficeMaintenance.ps1"));
                File.Delete(Path.Combine(runtimeDirectory, "plan.json"));
                File.Delete(Path.Combine(runtimeDirectory, "request.json"));
                File.Delete(Path.Combine(runtimeDirectory, "start.signal"));
                Directory.Delete(runtimeDirectory, false);
            }
            catch { }
            runtimeDirectory = null;
        }
        public void SelfTest(string outputDirectory)
        {
            var results = new List<string>();
            try
            {
                Directory.CreateDirectory(outputDirectory);
                Action<bool, string> check = delegate(bool condition, string message) { if (!condition) throw new Exception("FAILED: " + message); results.Add("PASS: " + message); };
                check(HasSuiteApp() && apps["Word"].IsChecked == true && apps["Excel"].IsChecked == true, "Word and Excel selected by default");
                check(!Install.IsEnabled && Download.IsEnabled, "Install gated before download");
                Capture(Path.Combine(outputDirectory, "01-start.png"));
                string first = MakePlan().PlanId;
                string originalFolder = baseFolder;
                string fixtureRoot = Path.Combine(outputDirectory, "readiness-fixture-" + Guid.NewGuid().ToString("N"));
                var fixtureFiles = new List<string>();
                var fixtureDirectories = new HashSet<string>();
                try
                {
                    baseFolder = fixtureRoot;
                    foreach (Choice architecture in Architecture.Items)
                    {
                        Architecture.SelectedItem = architecture;
                        Plan fixture = MakePlan();
                        string version = "16.0.12345.67890";
                        string payload = Path.Combine(fixture.DownloadPath, "Office", "Data", version);
                        Directory.CreateDirectory(payload);
                        string parent = payload;
                        while (parent != outputDirectory) { fixtureDirectories.Add(parent); parent = Path.GetDirectoryName(parent); }
                        string platform = architecture.Value == "64" ? "x64" : "x86";
                        var fixtureData = new Dictionary<string, string> {
                            { Path.Combine(fixture.DownloadPath, "ready.json"), json.Serialize(new { PlanId = fixture.PlanId, OfficeVersion = version }) },
                            { Path.Combine(fixture.DownloadPath, "setup.exe"), "test fixture, never executed" },
                            { Path.Combine(fixture.DownloadPath, "Config.xml"), "<Configuration/>" },
                            { Path.Combine(payload, "stream." + platform + ".x-none.dat"), "test" },
                            { Path.Combine(payload, "stream." + platform + "." + fixture.Language + ".dat"), "test" }
                        };
                        foreach (var entry in fixtureData) { File.WriteAllText(entry.Key, entry.Value); fixtureFiles.Add(entry.Key); }
                        check(LooksReady(fixture), "Saved " + architecture.Value + "-bit package is detected using " + platform);
                        fixture.PlanId = new string('0', 64);
                        check(!LooksReady(fixture), "Mismatched saved " + architecture.Value + "-bit plan is rejected");
                        fixture = MakePlan();
                        File.Delete(Path.Combine(payload, "stream." + platform + "." + fixture.Language + ".dat"));
                        check(!LooksReady(fixture), "Missing " + architecture.Value + "-bit language payload is rejected");
                    }
                }
                finally
                {
                    foreach (string path in fixtureFiles) File.Delete(path);
                    foreach (string path in fixtureDirectories.OrderByDescending(s => s.Length)) Directory.Delete(path, false);
                    baseFolder = originalFolder; Architecture.SelectedIndex = 0; SelectionChanged();
                }
                SelectApps(false); check(!HasSuiteApp() && !Download.IsEnabled, "Empty selection cannot download");
                apps["Teams"].IsChecked = true; check(!Download.IsEnabled, "Add-on only selection cannot download");
                for (int i = 0; i < 4; i++)
                {
                    profiles[i].IsChecked = true; SelectApps(true); Plan plan = MakePlan();
                    check(plan.Branch == branches[i] && plan.Channel == channels[i], "Edition/channel pair " + editions[i]);
                    check(apps["Publisher"].IsEnabled == (i < 2) && (apps["Publisher"].IsChecked == true) == (i < 2), "Publisher availability " + editions[i]);
                    check(plan.Components.Length == (i < 2 ? 9 : 8), "Select all available apps " + editions[i]);
                    check(plan.PlanId.Length == 64 && plan.PlanId != first, "Plan fingerprint changes " + editions[i]);
                }
                profiles[2].IsChecked = true;
                Capture(Path.Combine(outputDirectory, "02-all-apps.png"));
                ready = true; License.IsChecked = false; UpdateButtons(); check(!Install.IsEnabled, "License consent gates ready installation");
                License.IsChecked = true; check(Install.IsEnabled, "Ready download and consent enable installation");
                apps["Word"].IsChecked = false; check(!Install.IsEnabled && License.IsChecked == false, "Selection changes invalidate readiness and consent");
                busy = true; activeAction = "Download"; downloadJob = new DownloadJob(); UpdateButtons();
                Find<StackPanel>("OptionsPanel").IsEnabled = false; License.IsEnabled = false;
                check(Find<Button>("AbortButton").IsEnabled && Find<Button>("AbortButton").Visibility == Visibility.Visible, "Download exposes Abort button");
                check(Find<TextBlock>("StepDownload").FontWeight == FontWeights.SemiBold && Find<TextBlock>("StepChoose").FontWeight == FontWeights.Normal, "Download highlights sidebar step 02");
                HandleOutput("@@{\"kind\":\"progress\",\"message\":\"Downloading shared Office files · 1.24 GB received\",\"percent\":-1}", false);
                check(Progress.IsIndeterminate, "Unknown progress is indeterminate");
                Status.Text = "Downloading your Office…";
                Capture(Path.Combine(outputDirectory, "03-downloading.png"));
                Window.Width = 990; Window.Height = 730; Capture(Path.Combine(outputDirectory, "03-downloading-compact.png"));
                Window.Width = 1160; Window.Height = 880;
                AbortDownload();
                check(cancelRequested && !Find<Button>("AbortButton").IsEnabled && Status.Text == "Stopping your download…", "Abort handler enters stopping state and disables repeated clicks");
                HandleOutput("@@{\"kind\":\"progress\",\"message\":\"Late progress\",\"percent\":90}", false);
                check(Status.Text == "Stopping your download…" && Detail.Text != "Late progress", "Late worker output cannot overwrite stopping status");
                cancelRequested = false;
                activeAction = "Install"; UpdateButtons();
                check(Find<Button>("AbortButton").Visibility == Visibility.Collapsed && Find<TextBlock>("StepInstall").FontWeight == FontWeights.SemiBold, "Installation highlights step 03 without forced termination");
                downloadJob.Dispose(); downloadJob = null; busy = false; activeAction = null; UpdateButtons();
                Find<StackPanel>("OptionsPanel").IsEnabled = true; License.IsEnabled = true;
                TestDownloadCancellation(outputDirectory, check);
                HandleOutput("@@{\"kind\":\"error\",\"message\":\"Test failure\"}", false);
                check(lastError == "Test failure", "Structured backend error preserved");
                HandleOutput("@@{\"kind\":\"complete\",\"message\":\"Complete\"}", false);
                check(workerCompleted, "Backend completion requires explicit completion event");
                HandleOutput("@@not-json", false); check(activity.ToString().Contains("@@not-json"), "Malformed backend output retained in activity");
                Progress.IsIndeterminate = false; Progress.Value = 100;
                Status.Text = "Your apps are ready to install"; Detail.Text = "Download complete. Accept the license terms, then choose Install Office.";
                ready = true; License.IsChecked = true; UpdateButtons();
                check(Find<TextBlock>("StepInstall").FontWeight == FontWeights.SemiBold, "Ready package highlights sidebar step 03");
                Capture(Path.Combine(outputDirectory, "04-ready.png"));
                Window.Width = 990; Window.Height = 730; Capture(Path.Combine(outputDirectory, "05-compact.png"));
                TestMaintenanceUi(outputDirectory, check);
                results.Add("No Microsoft download, installation, registry changes, or preferences written by this test.");
                ExitCode = 0;
            }
            catch (Exception ex) { results.Add(ex.ToString()); ExitCode = 1; }
            finally { File.WriteAllLines(Path.Combine(outputDirectory, "ui-test-results.txt"), results); busy = false; Window.Close(); }
        }
        void Capture(string path)
        {
            Window.UpdateLayout();
            var content = (FrameworkElement)Window.Content;
            content.Measure(new Size(Window.Width - 16, Window.Height - 39));
            content.Arrange(new Rect(new Size(Window.Width - 16, Window.Height - 39)));
            content.UpdateLayout();
            var image = new RenderTargetBitmap((int)content.ActualWidth, (int)content.ActualHeight, 96, 96, PixelFormats.Pbgra32);
            image.Render(content);
            var encoder = new PngBitmapEncoder(); encoder.Frames.Add(BitmapFrame.Create(image));
            using (var file = File.Create(path)) encoder.Save(file);
        }

        void TestDownloadCancellation(string outputDirectory, Action<bool, string> check)
        {
            string id = Guid.NewGuid().ToString("N");
            string scriptPath = Path.Combine(outputDirectory, "cancel-test-" + id + ".ps1");
            string signal = Path.Combine(outputDirectory, "cancel-test-" + id + ".signal");
            string pidPath = Path.Combine(outputDirectory, "cancel-test-" + id + ".pid");
            string powershell = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.System), "WindowsPowerShell", "v1.0", "powershell.exe");
            string script = "param($Signal,$PidPath)\nwhile (-not [IO.File]::Exists($Signal)) { Start-Sleep -Milliseconds 50 }\n" +
                "$info = New-Object Diagnostics.ProcessStartInfo\n$info.FileName = (Get-Process -Id $PID).Path\n" +
                "$info.Arguments = '-NoProfile -NonInteractive -Command \"Start-Sleep -Seconds 60\"'\n$info.UseShellExecute = $false\n$info.CreateNoWindow = $true\n" +
                "$child = [Diagnostics.Process]::Start($info)\n[IO.File]::WriteAllText($PidPath, [string]$child.Id)\n$child.WaitForExit()\n";
            File.WriteAllText(scriptPath, script);
            Process parent = null, child = null, unrelated = null;
            try
            {
                unrelated = Process.Start(new ProcessStartInfo(powershell, "-NoProfile -NonInteractive -Command \"Start-Sleep -Seconds 60\"") { UseShellExecute = false, CreateNoWindow = true });
                using (var job = new DownloadJob())
                {
                    parent = Process.Start(new ProcessStartInfo(powershell, "-NoProfile -NonInteractive -ExecutionPolicy Bypass -File " + Quote(scriptPath) + " -Signal " + Quote(signal) + " -PidPath " + Quote(pidPath)) { UseShellExecute = false, CreateNoWindow = true });
                    job.Assign(parent);
                    Thread.Sleep(200);
                    check(!File.Exists(pidPath), "Download start gate prevents launching children before release");
                    File.WriteAllText(signal, "start");
                    var wait = Stopwatch.StartNew();
                    while (!File.Exists(pidPath) && wait.ElapsedMilliseconds < 10000) Thread.Sleep(50);
                    check(File.Exists(pidPath), "Cancellation fixture launches owned child");
                    child = Process.GetProcessById(int.Parse(File.ReadAllText(pidPath)));
                    job.Abort();
                    check(parent.WaitForExit(5000) && child.WaitForExit(5000), "Abort stops owned worker and child process");
                    check(!unrelated.HasExited, "Abort leaves unrelated processes running");
                }
            }
            finally
            {
                foreach (var process in new[] { parent, child, unrelated })
                    if (process != null) { try { if (!process.HasExited) { process.Kill(); process.WaitForExit(5000); } } finally { process.Dispose(); } }
                File.Delete(scriptPath); File.Delete(signal); File.Delete(pidPath);
            }
        }
    }
}
