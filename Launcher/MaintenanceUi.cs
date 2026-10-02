using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Markup;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Threading;

namespace Kareer.OfficeStudio
{
    public sealed class InstalledOfficeItem
    {
        public string Id { get; set; }
        public string Name { get; set; }
        public string Version { get; set; }
        public string Architecture { get; set; }
        public string Kind { get; set; }
        public string[] ProductIds { get; set; }
        public bool Removable { get; set; }
        public string Details { get; set; }
        public string DisplayName
        {
            get
            {
                if (Kind != "ClickToRun" || ProductIds == null || ProductIds.Length == 0) return Name;
                return string.Join(" + ", ProductIds.Select(FriendlyProduct).Distinct());
            }
        }
        static string FriendlyProduct(string id)
        {
            if (id.StartsWith("O365", StringComparison.OrdinalIgnoreCase)) return "Microsoft 365";
            if (id.StartsWith("OneNote", StringComparison.OrdinalIgnoreCase)) return "OneNote";
            if (id.StartsWith("Visio", StringComparison.OrdinalIgnoreCase)) return "Visio";
            if (id.StartsWith("Project", StringComparison.OrdinalIgnoreCase)) return "Project";
            foreach (string year in new[] { "2024", "2021", "2019" }) if (id.Contains(year)) return "Office " + year;
            return id;
        }
        public string ProductSummary { get { return ProductIds == null || ProductIds.Length == 0 ? "" : "Registered products: " + string.Join(", ", ProductIds); } }
        public string Summary { get { return string.Join("  ·  ", new[] { Kind == "ClickToRun" ? "Click-to-Run" : Kind, Architecture, string.IsNullOrWhiteSpace(Version) ? "" : "Build " + Version }.Where(s => !string.IsNullOrWhiteSpace(s))); } }
    }
    public sealed class OfficeInventory
    {
        public InstalledOfficeItem[] Items { get; set; }
        public string[] Warnings { get; set; }
        public string Fingerprint { get; set; }
    }
    public sealed partial class Studio
    {
        bool maintenanceRunning, inventoryScanning, inventoryKnown, receivedInventory, receivedTerms;
        OfficeInventory inventory;
        Window maintenanceWindow;
        string inventoryFailure;
        T Maintenance<T>(string name) where T : class { return maintenanceWindow == null ? null : maintenanceWindow.FindName(name) as T; }
        bool DeepSelected { get { var button = Maintenance<RadioButton>("DeepRemoval"); return button != null && button.IsChecked == true; } }
        bool ScanIsComplete { get { return inventoryKnown && inventory != null && inventory.Warnings != null && inventory.Warnings.Length == 0 && !string.IsNullOrEmpty(inventory.Fingerprint); } }

        void InitializeMaintenance()
        {
            Find<Button>("ManageOfficeButton").Click += delegate { ShowMaintenance(); };
            Window.ContentRendered += async delegate { if (!testMode) await RefreshOfficeInventory(); };
            UpdateInventoryDisplay();
        }
        void UpdateMaintenanceButtons()
        {
            var manage = Find<Button>("ManageOfficeButton");
            if (manage != null) manage.IsEnabled = !busy;
            if (maintenanceWindow == null) return;
            bool free = !busy && !maintenanceRunning;
            Maintenance<Button>("RefreshOfficeButton").IsEnabled = free;
            Maintenance<Button>("CleanupToolTerms").IsEnabled = free;
            Maintenance<RadioButton>("StandardRemoval").IsEnabled = free;
            Maintenance<RadioButton>("DeepRemoval").IsEnabled = free;
            Maintenance<CheckBox>("RemovalConsent").IsEnabled = free;
            bool hasTargets = inventory != null && inventory.Items != null && inventory.Items.Length > 0;
            bool standardSupported = hasTargets && inventory.Items.All(i => i.Removable);
            Maintenance<Button>("UninstallOfficeButton").IsEnabled = free && ScanIsComplete && (DeepSelected || standardSupported) && Maintenance<CheckBox>("RemovalConsent").IsChecked == true;
            Maintenance<Button>("UninstallOfficeButton").Content = DeepSelected ? "Run deep cleanup" : "Uninstall Office";
        }
        void UpdateInventoryDisplay()
        {
            string summary;
            if (inventoryScanning) summary = "Checking installed Office…";
            else if (!inventoryKnown) summary = inventoryFailure == null ? "Ready to check installed Office" : "Detection needs attention";
            else if (inventory.Items.Length == 0) summary = inventory.Warnings.Length == 0 ? "No installed Office found" : "Scan incomplete";
            else summary = inventory.Items.Length == 1 ? inventory.Items[0].DisplayName : inventory.Items.Length + " Office products detected";
            Find<TextBlock>("ComputerOfficeSummary").Text = summary;
            Find<TextBlock>("ComputerOfficeSummary").ToolTip = inventory != null && inventory.Items.Length > 0 ? string.Join(Environment.NewLine, inventory.Items.Select(i => i.Name + " · " + i.Summary)) : summary;
            if (maintenanceWindow != null)
            {
                Maintenance<TextBlock>("InventoryHeading").Text = summary;
                Maintenance<ItemsControl>("InstalledOfficeList").ItemsSource = inventory == null ? null : inventory.Items;
                Maintenance<TextBlock>("EmptyOfficeNote").Visibility = inventoryKnown && inventory.Items.Length == 0 ? Visibility.Visible : Visibility.Collapsed;
                string note = "Checks Click-to-Run, Windows Installer (MSI), and Microsoft Store Office installations.";
                if (inventoryScanning) note = "Reading Windows installation records. No changes are made.";
                else if (!inventoryKnown) note = inventoryFailure ?? "Choose Refresh to check this PC before uninstalling.";
                else if (inventory.Warnings.Length > 0) note = "Scan incomplete: " + string.Join(" ", inventory.Warnings);
                else if (inventory.Items.Any(i => !i.Removable)) note = "Some detected products need Microsoft's deep cleanup tool or their own uninstaller. Review the product details.";
                Maintenance<TextBlock>("InventoryNote").Text = note;
            }
            UpdateMaintenanceButtons();
        }
        void ShowMaintenance()
        {
            if (maintenanceWindow != null) { maintenanceWindow.Activate(); return; }
            using (Stream stream = Assembly.GetExecutingAssembly().GetManifestResourceStream("Kareer.MaintenanceWindow.xaml")) maintenanceWindow = (Window)XamlReader.Load(stream);
            maintenanceWindow.Owner = Window;
            maintenanceWindow.Icon = Window.Icon;
            maintenanceWindow.Resources.MergedDictionaries.Add(Window.Resources);
            if (!testMode)
            {
                maintenanceWindow.Height = Math.Min(maintenanceWindow.Height, SystemParameters.WorkArea.Height - 20);
                maintenanceWindow.MinHeight = Math.Min(maintenanceWindow.MinHeight, maintenanceWindow.Height);
            }
            Maintenance<Button>("RefreshOfficeButton").Click += async delegate { await RefreshOfficeInventory(); };
            Maintenance<Button>("CleanupToolTerms").Click += async delegate { await RunMaintenance("ToolTerms", "Deep"); };
            Maintenance<Button>("MaintenanceActivity").Click += delegate { ShowActivity(); };
            Maintenance<Button>("OfficeCleanupHelp").Click += delegate
            {
                try { Process.Start(new ProcessStartInfo("https://learn.microsoft.com/en-us/troubleshoot/microsoft-365/admin/miscellaneous/get-help-office-uninstall") { UseShellExecute = true }); }
                catch (Exception ex) { Log("Couldn't open Microsoft help: " + ex.Message); }
            };
            Maintenance<RadioButton>("StandardRemoval").Checked += delegate { RemovalModeChanged(); };
            Maintenance<RadioButton>("DeepRemoval").Checked += delegate { RemovalModeChanged(); };
            Maintenance<CheckBox>("RemovalConsent").Checked += delegate { UpdateMaintenanceButtons(); };
            Maintenance<CheckBox>("RemovalConsent").Unchecked += delegate { UpdateMaintenanceButtons(); };
            Maintenance<Button>("UninstallOfficeButton").Click += async delegate
            {
                if (Maintenance<Button>("UninstallOfficeButton").IsEnabled) await RunMaintenance("Uninstall", DeepSelected ? "Deep" : "Standard");
            };
            maintenanceWindow.Closing += delegate(object sender, System.ComponentModel.CancelEventArgs e) { if (busy && maintenanceRunning) e.Cancel = true; };
            maintenanceWindow.Closed += delegate { maintenanceWindow = null; };
            UpdateInventoryDisplay();
            RemovalModeChanged();
            maintenanceWindow.Show();
        }
        void RemovalModeChanged()
        {
            if (maintenanceWindow == null) return;
            Maintenance<CheckBox>("RemovalConsent").IsChecked = false;
            Maintenance<TextBlock>("RemovalDescription").Text = DeepSelected
                ? "Downloads Microsoft's current Get Help cleanup tool to remove Office, Project, Visio, and broken Office remnants. Office settings and licensing may be reset. A restart may be required."
                : "Removes all detected Office products, including Project, Visio, and Office language packs, using Microsoft's uninstallers.";
            Maintenance<TextBlock>("RemovalConsentText").Text = DeepSelected
                ? "I want Microsoft's tool to remove all Office from this PC, including Project, Visio, and broken remnants, and accept the tool license terms."
                : "I want to remove all Office products shown above, including any Project and Visio, and accept Microsoft's tool license terms.";
            Maintenance<TextBlock>("RemovalStatus").Text = "Review the products above, then tick the confirmation to enable removal.";
            UpdateMaintenanceButtons();
        }
        async Task<bool> RefreshOfficeInventory()
        {
            if (busy || maintenanceRunning) return false;
            return await RunMaintenance("Inventory", "Standard");
        }
        async Task<bool> CheckBeforeInstall()
        {
            if (!await RefreshOfficeInventory() || !ScanIsComplete)
            {
                Status.Text = "Check installed Office first";
                Detail.Text = "The installed Office scan is incomplete. Open Review & uninstall for details, then Refresh.";
                ShowMaintenance();
                return false;
            }
            string selected = ((Choice)Architecture.SelectedItem).Value;
            string conflict = InstallationConflict(inventory, selected);
            if (conflict != null)
            {
                Status.Text = "Review the existing Office installation";
                Detail.Text = conflict;
                ShowMaintenance();
                return false;
            }
            if (inventory.Items.Length > 0) Log("Existing Office detected before installation: " + string.Join("; ", inventory.Items.Select(i => i.Name + " (" + i.Summary + ")")));
            return true;
        }
        static string InstallationConflict(OfficeInventory snapshot, string selected)
        {
            if (snapshot.Items.Any(i => i.Kind == "MSI" || i.Kind == "Store" || i.Kind == "Unsupported"))
                return "An existing Office package uses another installation type. Review and remove it before installing with this launcher.";
            foreach (var item in snapshot.Items.Where(i => i.Kind == "ClickToRun"))
            {
                var platforms = (item.Architecture ?? "").Split(',').Select(s => s.Trim().ToLowerInvariant()).Select(s => s == "x64" || s == "64-bit" || s == "64" ? "64" : s == "x86" || s == "32-bit" || s == "32" ? "32" : "unknown").Distinct().ToArray();
                if (platforms.Length != 1 || platforms[0] == "unknown") return "The installed Office architecture is unknown or conflicting. Review the existing installation before proceeding.";
                if (platforms[0] != selected) return "Installed Office uses a different architecture. Match its 32/64-bit choice, or uninstall it first.";
            }
            return null;
        }
        void HandleMaintenanceEvent(string kind, string message, Dictionary<string, object> data)
        {
            if (kind == "terms" && data.ContainsKey("text"))
            {
                string terms = Convert.ToString(data["text"]);
                if (string.IsNullOrWhiteSpace(terms)) throw new InvalidDataException("The Microsoft tool returned empty license terms.");
                receivedTerms = true;
                if (!testMode)
                {
                    var termsWindow = new Window { Owner = Window, Title = "Microsoft Get Help | Tool license terms", Width = 800, Height = Math.Min(650, SystemParameters.WorkArea.Height - 40), WindowStartupLocation = WindowStartupLocation.CenterOwner, Icon = Window.Icon };
                    termsWindow.Content = new TextBox { Text = terms, IsReadOnly = true, TextWrapping = TextWrapping.Wrap, VerticalScrollBarVisibility = ScrollBarVisibility.Auto, FontFamily = new FontFamily("Segoe UI"), FontSize = 13, Padding = new Thickness(20), BorderThickness = new Thickness(0) };
                    termsWindow.Show();
                }
            }
            if (kind == "inventory" && data.ContainsKey("inventory"))
            {
                var snapshot = json.ConvertToType<OfficeInventory>(data["inventory"]);
                if (snapshot == null || snapshot.Items == null || snapshot.Warnings == null || !System.Text.RegularExpressions.Regex.IsMatch(snapshot.Fingerprint ?? "", "^[A-Fa-f0-9]{64}$")) throw new InvalidDataException("Invalid installed Office scan response.");
                inventory = snapshot; inventoryKnown = true; receivedInventory = true; inventoryFailure = null;
                if (maintenanceWindow != null) Maintenance<CheckBox>("RemovalConsent").IsChecked = false;
                UpdateInventoryDisplay();
            }
            if (maintenanceWindow != null && maintenanceRunning && (kind == "status" || kind == "progress" || kind == "error" || kind == "complete"))
            {
                Maintenance<TextBlock>("RemovalStatus").Text = message;
                Maintenance<TextBlock>("RemovalStatus").ToolTip = message;
                Maintenance<ProgressBar>("RemovalProgress").IsIndeterminate = kind != "complete" && kind != "error";
                if (kind == "complete") Maintenance<ProgressBar>("RemovalProgress").Value = 100;
            }
        }
        async Task<bool> RunMaintenance(string action, string mode)
        {
            if (busy || maintenanceRunning) return false;
            if (testMode && action != "Inventory") throw new InvalidOperationException("Only read-only inventory is allowed in diagnostic mode. Cleanup and network actions are disabled.");
            bool removal = action == "Uninstall";
            bool termsOnly = action == "ToolTerms";
            if (removal && (!ScanIsComplete || maintenanceWindow == null || Maintenance<CheckBox>("RemovalConsent").IsChecked != true)) return false;
            string fingerprint = inventory == null ? null : inventory.Fingerprint;
            maintenanceRunning = true; inventoryScanning = action == "Inventory"; receivedInventory = false; receivedTerms = false;
            workerCompleted = false; lastError = null; lastCompletion = null;
            if (removal) { busy = true; Progress.IsIndeterminate = true; Status.Text = mode == "Deep" ? "Running Office deep cleanup…" : "Uninstalling Office…"; clock.Restart(); timer.Start(); }
            else if (!termsOnly) { inventoryKnown = false; inventoryFailure = null; }
            Find<StackPanel>("OptionsPanel").IsEnabled = false;
            License.IsEnabled = false;
            UpdateInventoryDisplay(); UpdateButtons();
            bool success = false;
            try
            {
                Directory.CreateDirectory(Path.Combine(stateRoot, "Logs"));
                activeLog = Path.Combine(stateRoot, "Logs", DateTime.Now.ToString("yyyyMMdd-HHmmss-fff") + "-" + action.ToLowerInvariant() + ".log");
                File.WriteAllText(activeLog, activity.ToString(), Encoding.UTF8);
                CleanRuntime();
                runtimeDirectory = Path.Combine(stateRoot, "Runtime", Guid.NewGuid().ToString("N"));
                Directory.CreateDirectory(runtimeDirectory);
                foreach (string name in new[] { "DeploymentWorker.ps1", "OfficeMaintenance.ps1" })
                    using (Stream resource = Assembly.GetExecutingAssembly().GetManifestResourceStream("Kareer." + name))
                    using (var file = File.Create(Path.Combine(runtimeDirectory, name))) resource.CopyTo(file);
                string script = Path.Combine(runtimeDirectory, "OfficeMaintenance.ps1");
                string arguments = "-NoLogo -NoProfile -NonInteractive -ExecutionPolicy Bypass -File " + Quote(script) + " -Action " + action;
                if (removal || termsOnly)
                {
                    string request = Path.Combine(runtimeDirectory, "request.json");
                    string workPath = Path.Combine(stateRoot, "Maintenance", Guid.NewGuid().ToString("N"));
                    object requestData;
                    if (removal) requestData = new { Fingerprint = fingerprint, Confirmed = true, Mode = mode, WorkPath = workPath };
                    else requestData = new { WorkPath = workPath };
                    File.WriteAllText(request, json.Serialize(requestData), new UTF8Encoding(false));
                    arguments += " -RequestPath " + Quote(request);
                    if (removal) Log("User confirmed " + mode.ToLowerInvariant() + " removal of: " + string.Join("; ", inventory.Items.Select(i => i.Name)));
                }
                string systemFolder = Environment.GetFolderPath(Environment.SpecialFolder.System);
                if (Environment.Is64BitOperatingSystem && !Environment.Is64BitProcess) systemFolder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Sysnative");
                var start = new ProcessStartInfo(Path.Combine(systemFolder, "WindowsPowerShell", "v1.0", "powershell.exe"), arguments)
                {
                    UseShellExecute = false, CreateNoWindow = true, WindowStyle = ProcessWindowStyle.Hidden,
                    RedirectStandardOutput = true, RedirectStandardError = true, StandardOutputEncoding = Encoding.UTF8, StandardErrorEncoding = Encoding.UTF8,
                    WorkingDirectory = runtimeDirectory
                };
                worker = new Process { StartInfo = start };
                worker.OutputDataReceived += delegate(object sender, DataReceivedEventArgs e) { if (e.Data != null) PostOutput(e.Data, false); };
                worker.ErrorDataReceived += delegate(object sender, DataReceivedEventArgs e) { if (e.Data != null) PostOutput(e.Data, true); };
                worker.Start(); worker.BeginOutputReadLine(); worker.BeginErrorReadLine();
                await Task.Run(delegate { worker.WaitForExit(); });
                int code = worker.ExitCode;
                await Window.Dispatcher.InvokeAsync(delegate { }, DispatcherPriority.Background);
                if (code != 0 || !workerCompleted || (termsOnly ? !receivedTerms : !receivedInventory)) throw new InvalidOperationException(lastError ?? "The requested Office operation did not complete (exit code " + code + "). Open Activity for details.");
                if (removal)
                {
                    Status.Text = "Office removal finished";
                    Detail.Text = lastCompletion;
                    Detail.ToolTip = lastCompletion;
                    Progress.IsIndeterminate = false; Progress.Value = 100;
                    Find<TextBlock>("ProgressLabel").Text = "REVIEW RESULT";
                }
                success = true;
            }
            catch (Exception ex)
            {
                Log("ERROR: " + ex.Message);
                if (!termsOnly) { inventoryKnown = false; inventoryFailure = ex.Message + " Refresh to verify the current installation state."; }
                if (removal) { Status.Text = "Office removal needs attention"; Detail.Text = ex.Message; Detail.ToolTip = ex.Message; Progress.IsIndeterminate = false; Progress.Value = 0; Find<TextBlock>("ProgressLabel").Text = "SEE ACTIVITY"; }
                if (maintenanceWindow != null)
                {
                    Maintenance<TextBlock>("RemovalStatus").Text = ex.Message;
                    Maintenance<TextBlock>("RemovalStatus").ToolTip = ex.Message;
                    Maintenance<ProgressBar>("RemovalProgress").IsIndeterminate = false;
                }
                if (removal) ShowActivity();
            }
            finally
            {
                if (worker != null) { worker.Dispose(); worker = null; }
                maintenanceRunning = false; inventoryScanning = false;
                if (removal) { busy = false; timer.Stop(); clock.Stop(); }
                Find<StackPanel>("OptionsPanel").IsEnabled = true; License.IsEnabled = true;
                if (maintenanceWindow != null) { Maintenance<CheckBox>("RemovalConsent").IsChecked = false; Maintenance<ProgressBar>("RemovalProgress").IsIndeterminate = false; }
                CleanRuntime(); UpdateInventoryDisplay(); UpdateButtons();
            }
            return success;
        }

        void TestMaintenanceUi(string outputDirectory, Action<bool, string> check)
        {
            var fixture = new OfficeInventory {
                Fingerprint = new string('a', 64), Warnings = new string[0], Items = new[] {
                    new InstalledOfficeItem { Id = "c2r", Name = "Microsoft Office Professional Plus 2021", Version = "16.0.14332.21008", Architecture = "64-bit", Kind = "ClickToRun", Removable = true, ProductIds = new[] { "ProPlus2021Volume", "VisioPro2021Volume" }, Details = "Includes Office applications and Visio. All installed Office languages are included in removal." }
                }
            };
            HandleOutput("@@" + json.Serialize(new { kind = "inventory", message = "Test inventory", inventory = fixture }), false);
            check(inventoryKnown && inventory.Items.Length == 1, "Installed Office inventory message is displayed");
            check(InstallationConflict(fixture, "64") == null && InstallationConflict(fixture, "32") != null, "Installation preflight detects a 32/64-bit mismatch");
            fixture.Items[0].Architecture = "x64, x86"; check(InstallationConflict(fixture, "64") != null, "Conflicting architecture registrations cannot bypass preflight");
            fixture.Items[0].Architecture = ""; check(InstallationConflict(fixture, "64") != null, "Unknown installed architecture requires review");
            fixture.Items[0].Architecture = "64-bit";
            HandleOutput("@@" + json.Serialize(new { kind = "terms", message = "Fixture terms", text = "Test license only" }), false);
            check(receivedTerms, "Tool license text is received separately from cleanup");
            ShowMaintenance();
            check(!Maintenance<Button>("UninstallOfficeButton").IsEnabled, "Uninstall requires explicit confirmation");
            Maintenance<CheckBox>("RemovalConsent").IsChecked = true;
            check(Maintenance<Button>("UninstallOfficeButton").IsEnabled, "Supported complete scan and consent enable uninstall");
            Maintenance<RadioButton>("DeepRemoval").IsChecked = true;
            check(Maintenance<CheckBox>("RemovalConsent").IsChecked == false, "Changing cleanup mode clears consent");
            CaptureMaintenance(Path.Combine(outputDirectory, "06-installed-office.png"));
            fixture.Warnings = new[] { "Store inventory could not be read." }; inventory = fixture; UpdateInventoryDisplay();
            Maintenance<CheckBox>("RemovalConsent").IsChecked = true;
            check(!Maintenance<Button>("UninstallOfficeButton").IsEnabled, "An incomplete inventory blocks destructive removal");
            fixture.Warnings = new string[0]; fixture.Items = new InstalledOfficeItem[0]; inventory = fixture; UpdateInventoryDisplay();
            Maintenance<RadioButton>("StandardRemoval").IsChecked = true; Maintenance<CheckBox>("RemovalConsent").IsChecked = true;
            check(!Maintenance<Button>("UninstallOfficeButton").IsEnabled, "Standard uninstall requires detected Office");
            Maintenance<RadioButton>("DeepRemoval").IsChecked = true; Maintenance<CheckBox>("RemovalConsent").IsChecked = true;
            check(Maintenance<Button>("UninstallOfficeButton").IsEnabled, "Explicit deep cleanup can handle unregistered remnants");
            maintenanceRunning = true; UpdateMaintenanceButtons(); check(!Maintenance<Button>("UninstallOfficeButton").IsEnabled, "Concurrent operations cannot start removal"); maintenanceRunning = false;
            maintenanceWindow.Close();
            inventoryKnown = false; inventory = null; UpdateInventoryDisplay();
        }
        public async Task InventoryTest(string outputDirectory)
        {
            try
            {
                Directory.CreateDirectory(outputDirectory);
                bool success = await RefreshOfficeInventory();
                if (!success || !inventoryKnown) throw new Exception(inventoryFailure ?? "The live inventory did not return a valid snapshot.");
                File.WriteAllText(Path.Combine(outputDirectory, "inventory-live.json"), json.Serialize(inventory), Encoding.UTF8);
                File.WriteAllText(Path.Combine(outputDirectory, "inventory-live-result.txt"), "Read-only end-to-end inventory: PASS" + Environment.NewLine + "Detected product groups: " + inventory.Items.Length + Environment.NewLine + "Scan warnings: " + inventory.Warnings.Length + Environment.NewLine + "No uninstall, cleanup, application closing, or network download was requested.");
                ShowMaintenance();
                CaptureMaintenance(Path.Combine(outputDirectory, "07-live-installed-office.png"));
                Capture(Path.Combine(outputDirectory, "08-live-home.png"));
                ExitCode = 0;
            }
            catch (Exception ex) { File.WriteAllText(Path.Combine(outputDirectory, "inventory-live-result.txt"), ex.ToString()); ExitCode = 1; }
            finally { Window.Close(); }
        }
        void CaptureMaintenance(string path)
        {
            maintenanceWindow.UpdateLayout();
            var content = (FrameworkElement)maintenanceWindow.Content;
            content.Measure(new Size(maintenanceWindow.Width - 16, maintenanceWindow.Height - 39));
            content.Arrange(new Rect(new Size(maintenanceWindow.Width - 16, maintenanceWindow.Height - 39)));
            content.UpdateLayout();
            var bitmap = new RenderTargetBitmap((int)content.ActualWidth, (int)content.ActualHeight, 96, 96, PixelFormats.Pbgra32); bitmap.Render(content);
            var encoder = new PngBitmapEncoder(); encoder.Frames.Add(BitmapFrame.Create(bitmap)); using (var file = File.Create(path)) encoder.Save(file);
        }
    }
}
