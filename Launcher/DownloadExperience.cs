using System;
using System.Collections.Generic;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Interop;
using System.Windows.Media;

namespace Kareer.OfficeStudio
{
    public sealed partial class Studio
    {
        string downloadStage = "";
        long lastDownloadBytes = -1;
        double bytesChangedAt, lastDownloadHeartbeat;
        bool readyNotificationSent;
        int readyNotificationCount;
        System.Windows.Forms.NotifyIcon downloadTray;

        void BeginDownloadExperience()
        {
            Progress.IsIndeterminate = true; Progress.Value = 0;
            readyNotificationSent = false;
            lastDownloadBytes = -1;
            bytesChangedAt = lastDownloadHeartbeat = 0;
            SetDownloadStage("preparing");
            if (!testMode) EnsureDownloadTray();
        }

        void SetDownloadStage(string stage)
        {
            if (stage != "preparing" && stage != "checking-saved" && stage != "downloading" && stage != "reusing" && stage != "checking" && stage != "ready") return;
            if (downloadStage != stage) { lastDownloadBytes = -1; bytesChangedAt = clock.Elapsed.TotalSeconds; lastDownloadHeartbeat = 0; }
            downloadStage = stage;
            Status.Text = stage == "ready" ? "Your download is ready" : stage == "checking-saved" ? "Checking your saved files…" :
                stage == "checking" ? "Checking downloaded files…" : stage == "preparing" ? "Preparing your download…" : "Downloading Office files…";
            RenderDownloadExperience();
        }

        void RenderDownloadExperience()
        {
            bool downloading = busy && activeAction == "Download";
            bool show = !cancelRequested && (downloading || (!busy && !maintenanceRunning && ready));
            Find<StackPanel>("DownloadExperience").Visibility = show ? Visibility.Visible : Visibility.Collapsed;
            if (!show) return;
            string stage = downloading ? downloadStage : "ready";
            int current = stage == "ready" ? 4 : stage == "checking" ? 3 : (stage == "downloading" || stage == "reusing") ? 2 : 1;
            for (int index = 1; index <= 4; index++)
            {
                var chip = Find<Border>("DownloadStage" + index);
                chip.Background = (Brush)new BrushConverter().ConvertFrom(index == current ? "#147C71" : "#F0F3F7");
                ((TextBlock)chip.Child).Foreground = index == current ? Brushes.White : (Brush)Window.FindResource("Muted");
            }
            Find<TextBlock>("DownloadReassurance").Text = DownloadReassurance(stage, clock.Elapsed.TotalSeconds);
        }

        string DownloadReassurance(string stage, double elapsedSeconds)
        {
            if (stage == "ready") return License.IsChecked == true ? "All set. Choose Install Office when you're ready." : "All set. Accept Microsoft's license terms below, then choose Install Office when you're ready.";
            if (stage == "checking" || stage == "checking-saved") return stage == "checking-saved" ?
                "Checking and reusing your saved files. This can take a little while, even when no download is needed." :
                "The Office files are downloaded. Checking large files can take a while. You can minimize this window.";
            if (stage == "preparing") return "Getting everything ready. You can minimize this window; we'll let you know when your download is ready.";
            if (lastDownloadBytes >= 0 && elapsedSeconds - bytesChangedAt >= 60 && elapsedSeconds - lastDownloadHeartbeat <= 25)
                return "Microsoft setup is still running. Files may be downloading or being checked, even when the size stays the same.";
            if (stage == "reusing") return "Reusing the files already on this PC. Microsoft may download incomplete files again. You can minimize this window.";
            return "Office can take a while to download. You can minimize this window; we'll let you know when it's ready.";
        }

        void UpdateDownloadBytes(long bytes, double elapsedSeconds)
        {
            if (bytes < 0 || cancelRequested || !busy || activeAction != "Download") return;
            if (lastDownloadBytes != bytes) { lastDownloadBytes = bytes; bytesChangedAt = elapsedSeconds; }
            lastDownloadHeartbeat = elapsedSeconds;
            Find<TextBlock>("SavedSize").Text = "Office files on disk: " + (bytes / 1048576.0).ToString("N1") + " MB";
            RenderDownloadExperience();
        }

        void RestoreDownloadWindow()
        {
            if (Window.WindowState == WindowState.Minimized) Window.WindowState = WindowState.Normal;
            Window.Show(); Window.Activate();
        }

        void EnsureDownloadTray(bool preview = false)
        {
            if (downloadTray != null || (testMode && !preview)) return;
            try
            {
                using (var stream = Assembly.GetExecutingAssembly().GetManifestResourceStream("Kareer.ico"))
                using (var icon = new System.Drawing.Icon(stream))
                    downloadTray = new System.Windows.Forms.NotifyIcon { Icon = (System.Drawing.Icon)icon.Clone(), Text = "Office Studio", Visible = true };
                downloadTray.Click += delegate { RestoreDownloadWindow(); };
                downloadTray.BalloonTipClicked += delegate { RestoreDownloadWindow(); };
            }
            catch (Exception ex) { DisposeDownloadTray(); Log("Desktop notification unavailable: " + ex.Message); }
        }

        void NotifyDownloadReady()
        {
            // A worker message alone is not success: RunWorker calls this only
            // after exit-code and saved-package checks have passed.
            if (!ready || !workerCompleted || cancelRequested || activeAction != "Download" || readyNotificationSent) return;
            readyNotificationSent = true; readyNotificationCount++;
            if (testMode) return;
            ShowDownloadNotification();
        }

        bool ShowDownloadNotification(bool preview = false)
        {
            EnsureDownloadTray(preview);
            try
            {
                if (downloadTray == null) return false;
                downloadTray.ShowBalloonTip(10000,
                    preview ? "Office Studio — notification test" : "Office Studio — your download is ready",
                    preview ? "Testing the ready notification. No Office download or installation was started." : "Your selected apps are ready to install. Click here to return to Office Studio.", System.Windows.Forms.ToolTipIcon.Info);
                if (!Window.IsActive || Window.WindowState == WindowState.Minimized)
                {
                    var flash = new FlashInfo { Size = (uint)Marshal.SizeOf(typeof(FlashInfo)), Window = new WindowInteropHelper(Window).Handle, Flags = 2 | 12, Count = 3, Timeout = 0 };
                    FlashWindowEx(ref flash);
                }
                return true;
            }
            catch (Exception ex) { Log("Desktop notification unavailable; your package is still ready: " + ex.Message); return false; }
        }

        public async Task NotificationTest(string outputDirectory)
        {
            var results = new List<string>();
            System.IO.Directory.CreateDirectory(outputDirectory);
            try
            {
                Window.WindowState = WindowState.Minimized;
                if (!ShowDownloadNotification(true)) throw new Exception("The native notification request failed.");
                if (downloadTray == null || !downloadTray.Visible) throw new Exception("Windows notification-area icon was not initialized.");
                results.Add("PASS: Native Windows notification requested while the window was minimized.");
                await Task.Delay(1500);
                RestoreDownloadWindow();
                if (Window.WindowState == WindowState.Minimized || !Window.IsVisible) throw new Exception("Notification navigation did not restore the window.");
                if (worker != null || busy || ready) throw new Exception("Notification preview started a deployment operation.");
                results.Add("PASS: Notification navigation restores the window without starting installation.");
                DisposeDownloadTray();
                if (downloadTray != null) throw new Exception("Notification icon was not disposed.");
                results.Add("PASS: Notification icon cleaned up.");
                results.Add("Windows notification settings may suppress banners; this check verifies the notification request and window behavior.");
                ExitCode = 0;
            }
            catch (Exception ex) { results.Add(ex.ToString()); ExitCode = 1; }
            finally { System.IO.File.WriteAllLines(System.IO.Path.Combine(outputDirectory, "notification-test-results.txt"), results); Window.Close(); }
        }

        void DisposeDownloadTray()
        {
            if (downloadTray == null) return;
            var icon = downloadTray.Icon;
            downloadTray.Visible = false; downloadTray.Dispose(); downloadTray = null;
            if (icon != null) icon.Dispose();
        }

        [StructLayout(LayoutKind.Sequential)]
        struct FlashInfo { public uint Size; public IntPtr Window; public uint Flags, Count, Timeout; }
        [DllImport("user32.dll")] static extern bool FlashWindowEx(ref FlashInfo info);

        void TestDownloadExperience(string outputDirectory, Action<bool, string> check)
        {
            busy = true; activeAction = "Download"; ready = false; cancelRequested = false; workerCompleted = false;
            downloadJob = new DownloadJob();
            Find<StackPanel>("OptionsPanel").IsEnabled = false; License.IsEnabled = false;
            BeginDownloadExperience(); UpdateButtons();
            check(Status.Text == "Preparing your download…", "Download starts with explicit preparation stage");
            HandleOutput("@@" + json.Serialize(new { kind = "stage", stage = "checking-saved", message = "Checking saved files" }), false);
            check(Status.Text == "Checking your saved files…", "Cache verification has its own status");
            HandleOutput("@@" + json.Serialize(new { kind = "stage", stage = "reusing", message = "Reusing saved files" }), false);
            check(downloadStage == "reusing" && Find<TextBlock>("DownloadReassurance").Text.Contains("Reusing"), "Partial downloads explain automatic file reuse");
            HandleOutput("@@" + json.Serialize(new { kind = "stage", stage = "downloading", message = "Downloading Office files" }), false);
            HandleOutput("@@" + json.Serialize(new { kind = "progress", percent = 85, message = "Addon progress", bytesOnDisk = 1073741824L }), false);
            check(Progress.IsIndeterminate, "Addon percentage is never presented as overall Office progress");
            check(lastDownloadBytes == 1073741824L && Find<TextBlock>("SavedSize").Text.Contains("MB"), "Structured file bytes update the visible MB counter");
            check(Status.Text == "Downloading Office files…", "Progress log wording cannot change the stage");
            UpdateDownloadBytes(1073741824L, 65);
            check(DownloadReassurance("downloading", 65).Contains("still running"), "Unchanged size with a recent heartbeat gets a reassuring message");
            check(!DownloadReassurance("downloading", 120).Contains("still running"), "An old heartbeat cannot claim Microsoft setup is still running");
            UpdateDownloadBytes(1073741825L, 66);
            check(!DownloadReassurance("downloading", 66).Contains("still running"), "A changed size clears the unchanged-size message");
            HandleOutput("@@" + json.Serialize(new { kind = "stage", stage = "checking", message = "Checking file hashes" }), false);
            check(Status.Text == "Checking downloaded files…" && !ready, "Checking files is distinct from ready to install");
            Capture(System.IO.Path.Combine(outputDirectory, "10-checking-files.png"));
            int previousNotifications = readyNotificationCount;
            HandleOutput("@@{\"kind\":\"complete\",\"message\":\"Worker finished\"}", false);
            NotifyDownloadReady();
            check(readyNotificationCount == previousNotifications, "A completion event without a verified ready package cannot notify");
            ready = true; cancelRequested = true; NotifyDownloadReady();
            check(readyNotificationCount == previousNotifications, "Cancelled downloads never notify ready");
            cancelRequested = false; activeAction = "Install"; NotifyDownloadReady();
            check(readyNotificationCount == previousNotifications, "Install events never trigger a download-ready notification");
            activeAction = "Download"; SetDownloadStage("ready"); NotifyDownloadReady(); NotifyDownloadReady();
            check(readyNotificationCount == previousNotifications + 1, "A verified successful download notifies exactly once");
            busy = false; activeAction = null; License.IsChecked = false; UpdateButtons();
            downloadJob.Dispose(); downloadJob = null;
            Find<StackPanel>("OptionsPanel").IsEnabled = true; License.IsEnabled = true;
            check(!Install.IsEnabled && Find<TextBlock>("DownloadReassurance").Text.Contains("license"), "Ready guidance preserves the license-consent requirement");
            License.IsChecked = true; UpdateButtons();
            check(Install.IsEnabled, "Install becomes available after ready state and consent");
        }
    }
}
