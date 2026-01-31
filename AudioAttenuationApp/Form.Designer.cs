using AudioAttenuationApp.Helper;
using AudioAttenuationApp.Models;
using CSCore.CoreAudioAPI;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace AudioAttenuationApp
{
    partial class Form
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;
        private NotifyIcon notifyIcon;
        private System.Windows.Forms.Button refreshButton;
        private System.Windows.Forms.TextBox delayTextBox;
        private Label delayLabel;
        private System.Windows.Forms.TrackBar lowVolumeTrackBar;
        private Label lowVolumeLabel;
        private ListBox processListBox;

        private CancellationTokenSource processCts;
        private CancellationTokenSource detectCts;

        private CancellationTokenSource adjustCts;

        private int? selectedProcessId = null;
        private const int DEFAULT_LOW_VOLUME = 30;
        private float lowerVolumeLimit = DEFAULT_LOW_VOLUME / 100.0f;

        private const int DEFAULT_SILENT_MS = 1000;
        private const int MIN_DELAY_TIME_MS = 300;
        private const int MAX_DELAY_TIME_MS = 5000;
        private int silentThresholdMs = DEFAULT_SILENT_MS;
        private readonly Stopwatch silentWatch = Stopwatch.StartNew();

        private AdjustingStage PreviousStage;

        private bool isClosingHandled = false;
        private ToolTip errorToolTip = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        private async void Form_Load(object sender, EventArgs e)
        {
            this.delayTextBox.Text = silentThresholdMs.ToString();
            this.lowVolumeLabel.Text = formatVolumeLabel(DEFAULT_LOW_VOLUME);
            this.lowVolumeTrackBar.Value = DEFAULT_LOW_VOLUME;
            this.errorToolTip = new ToolTip
            {
                ToolTipIcon = ToolTipIcon.Error,
                ToolTipTitle = "Invalid value"
            };
            await GetProcessAsync();
        }

        private async void Form_Closing(object sender, FormClosingEventArgs e)
        {
            if (isClosingHandled) return;
            e.Cancel = true;
            try
            {
                Debug.WriteLine("Form closing, resetting volumes...");
                selectedProcessId = null;
                processCts?.Cancel();
                processCts?.Dispose();

                using var resetCts = new CancellationTokenSource(TimeSpan.FromSeconds(2));
                await Task.Run(() => AdjustVolume(AdjustingStage.Up, resetCts.Token));
            }
            catch (OperationCanceledException)
            {
                Debug.WriteLine("Failed to reset volume.");
            }
            finally
            {
                isClosingHandled = true;
                Close();
            }
        }

        private void Form_Resize(object sender, EventArgs e)
        {
            switch (this.WindowState)
            {
                case FormWindowState.Minimized:
                    notifyIcon.Visible = true;
                    this.ShowInTaskbar = false;
                    this.Hide();
                    break;
                case FormWindowState.Normal:
                    notifyIcon.Visible = false;
                    this.ShowInTaskbar = true;
                    this.Show();
                    break;
            }
        }

        #region Process Detection

        private async Task GetProcessAsync()
        {
            refreshButton.Enabled = false;
            refreshButton.Text = "Listing...";
            processListBox.Items.Clear();

            processCts?.Cancel();
            processCts?.Dispose();

            processCts = new CancellationTokenSource();

            try
            {
                var list = await Task.Run(() => GetProcessList(processCts.Token), processCts.Token);
                foreach (var item in list)
                    processListBox.Items.Add(
                        new ProcessItem(item.Id, item.Name, item.Icon)
                    );

                int matchIndex = list.FindIndex(p => p.Id == selectedProcessId);
                if (matchIndex >= 0)
                    processListBox.SelectedIndex = matchIndex;
            }
            catch (OperationCanceledException)
            {
                Debug.WriteLine("Process listing cancelled.");
            }
            finally
            {
                refreshButton.Enabled = true;
                refreshButton.Text = "Refresh";
            }
        }

        private List<ProcessItem> GetProcessList(CancellationToken token)
        {
            var result = new List<ProcessItem>();

            using var sessionManager = GetDefaultAudioSessionManager2(DataFlow.Render);
            using var sessions = sessionManager.GetSessionEnumerator();

            foreach (var s in sessions)
            {
                token.ThrowIfCancellationRequested();

                using var session = s.QueryInterface<AudioSessionControl2>();
                result.Add(new ProcessItem(
                    session.ProcessID,
                    GetSessionName(session),
                    GetSessionIcon(session)));
            }

            return result;
        }

        #endregion

        #region Sound Detection and Volume Adjustment
        private void runDetectAudioBackgroundWorker()
        {
            if (detectCts != null) return;

            detectCts = new CancellationTokenSource();
            Task.Run(() => DetectAudioLoop(detectCts.Token), detectCts.Token);
        }

        private async Task DetectAudioLoop(CancellationToken token)
        {
            AudioMeterInformation trackedMeter = null;
            HashSet<string> previousSessions = new();

            try
            {
                while (!token.IsCancellationRequested)
                {
                    HashSet<string> currentSessions = new();

                    using var sessionManager = GetDefaultAudioSessionManager2(DataFlow.Render);
                    using var sessions = sessionManager.GetSessionEnumerator();

                    foreach (var s in sessions)
                    {
                        token.ThrowIfCancellationRequested();

                        using var session = s.QueryInterface<AudioSessionControl2>();

                        currentSessions.Add(session.SessionInstanceIdentifier);

                        if (session.ProcessID == selectedProcessId)
                        {
                            trackedMeter?.Dispose();
                            trackedMeter = session.QueryInterface<AudioMeterInformation>();
                        }
                    }

                    if (trackedMeter != null)
                    {
                        bool isNoSound = trackedMeter.GetPeakValue() < 0.001f;

                        if (isNoSound)
                            silentWatch.Start();
                        else
                            silentWatch.Reset();

                        AdjustingStage currentStage;

                        if (isNoSound && silentWatch.ElapsedMilliseconds >= silentThresholdMs)
                            currentStage = AdjustingStage.Up;
                        else
                            currentStage = AdjustingStage.Down;

                        if ((currentStage != PreviousStage && currentStage != AdjustingStage.None)
                            || !previousSessions.SetEquals(currentSessions))
                        {
                            adjustCts?.Cancel();
                            adjustCts?.Dispose();

                            adjustCts = CancellationTokenSource.CreateLinkedTokenSource(token);

                            _ = Task.Run(() => AdjustVolume(currentStage, adjustCts.Token));

                            Debug.WriteLine($"Adjusting Stage: {currentStage}");

                            PreviousStage = currentStage;
                            previousSessions = currentSessions;
                        }
                    }

                    await Task.Delay(300, token).ConfigureAwait(false);
                }
            }
            catch (OperationCanceledException)
            {
                Debug.WriteLine("Audio detection loop cancelled.");
            }
            finally
            {
                trackedMeter?.Dispose();

                adjustCts?.Cancel();
                adjustCts?.Dispose();
                adjustCts = null;
            }
        }

        private async Task AdjustVolume(AdjustingStage stage, CancellationToken token)
        {
            var tasks = new List<Task>();

            using var sessionManager = GetDefaultAudioSessionManager2(DataFlow.Render);
            using var sessions = sessionManager.GetSessionEnumerator();

            foreach (var s in sessions)
            {
                token.ThrowIfCancellationRequested();

                using var session = s.QueryInterface<AudioSessionControl2>();

                if (session.ProcessID == selectedProcessId)
                    continue;

                // Capture everything per-iteration
                var simpleVolume = session.QueryInterface<SimpleAudioVolume>();

                tasks.Add(Task.Run(async () =>
                {
                    using (simpleVolume)
                    using (var fader = new VolumeFader(simpleVolume))
                    {
                        switch (stage)
                        {
                            case AdjustingStage.Up:
                                await fader.FadeToAsync(1.0f, 240, 15)
                                           .ConfigureAwait(false);
                                break;

                            case AdjustingStage.Down:
                                await fader.FadeToAsync(lowerVolumeLimit, 200, 13)
                                           .ConfigureAwait(false);
                                break;
                        }
                    }
                }, token));
            }

            await Task.WhenAll(tasks).ConfigureAwait(false);
        }

        #endregion

        #region Helper functions

        private AudioSessionManager2 GetDefaultAudioSessionManager2(DataFlow dataFlow, Role role = Role.Multimedia)
        {
            var enumerator = new MMDeviceEnumerator();
            var device = enumerator.GetDefaultAudioEndpoint(dataFlow, role);
            enumerator.Dispose();
            return AudioSessionManager2.FromMMDevice(device);
        }

        private string GetSessionName(AudioSessionControl2 session)
        {
            // System sound
            if (session.IsSystemSoundSession)
                return "System Sounds";
            // Process name
            if (!string.IsNullOrEmpty(session.Process.ProcessName))
                return session.Process.ProcessName;
            // Display name
            if (!string.IsNullOrEmpty(session.DisplayName))
                return session.DisplayName;

            return "Unknown";
        }

        private Icon GetSessionIcon(AudioSessionControl2 session)
        {
            // Session icon
            if (!string.IsNullOrEmpty(session.IconPath))
                return ShellIconExtractor.LoadIndirectIcon(session.IconPath);

            // Process icon
            try
            {
                return Icon.ExtractAssociatedIcon(session.Process.MainModule.FileName);
            }
            catch
            {
                Debug.WriteLine("Failed to get process icon for " + session.ProcessID);
            }

            // Fallback
            return SystemIcons.Application;
        }

        private string formatVolumeLabel(int volume)
        {
            return "Low: " + volume + "%";
        }

        #endregion

        #region Form component functions

        private void notifyIcon_MouseDoubleClick(object sender, EventArgs e)
        {
            this.Show();
            this.WindowState = FormWindowState.Normal;
        }

        private async void refreshButton_Click(object sender, EventArgs e)
        {
            await GetProcessAsync();
        }

        private void processListBox_SelectedValueChanged(object sender, EventArgs e)
        {
            if (processListBox.SelectedItem == null)
            {
                return;
            }

            var selectedItem = (ProcessItem)processListBox.SelectedItem;
            selectedProcessId = selectedItem.Id;
            runDetectAudioBackgroundWorker();
        }

        private void delayTextBox_TextChanged(object sender, EventArgs e)
        {
            if (!int.TryParse(delayTextBox.Text, out int value) ||
                value < MIN_DELAY_TIME_MS || value > MAX_DELAY_TIME_MS)
            {
                errorToolTip.Show(
                        $"Must between {MIN_DELAY_TIME_MS}-{MAX_DELAY_TIME_MS} ms",
                        delayTextBox,
                        -delayTextBox.Width,
                        -delayTextBox.Height * 2,
                        2000
                );
                return;
            }

            silentThresholdMs = value;
        }

        private void delayTextBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar))
            {
                e.Handled = true;
            }
        }

        private void lowVolumnTrackBar_Scroll(object sender, EventArgs e)
        {
            lowerVolumeLimit = lowVolumeTrackBar.Value / 100.0f;
            lowVolumeLabel.Text = formatVolumeLabel(lowVolumeTrackBar.Value);
        }

        private void processListBox_DrawItem(object sender, DrawItemEventArgs e)
        {
            e.DrawBackground();

            if (e.Index < 0 || e.Index >= processListBox.Items.Count)
                return;

            var item = (ProcessItem)processListBox.Items[e.Index];
            var g = e.Graphics;

            int iconSize = 32;  // increase icon size here
            int padding = 6;

            // Center icon vertically within the item bounds
            int iconTop = e.Bounds.Top + (e.Bounds.Height - iconSize) / 2;

            if (item.Icon != null)
            {
                g.DrawIcon(item.Icon, new Rectangle(e.Bounds.Left + padding, iconTop, iconSize, iconSize));
            }

            // Text position starts after icon + padding
            int textLeft = e.Bounds.Left + padding + iconSize + padding;
            int textTop = e.Bounds.Top + (e.Bounds.Height - e.Font.Height) / 2;

            using (var brush = new SolidBrush(e.ForeColor))
            {
                g.DrawString($"{item.Id} - {item.Name}", e.Font, brush, textLeft, textTop);
            }

            e.DrawFocusRectangle();
        }

        #endregion

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            components = new Container();
            ComponentResourceManager resources = new ComponentResourceManager(typeof(Form));
            notifyIcon = new NotifyIcon(components);
            refreshButton = new Button();
            delayTextBox = new TextBox();
            delayLabel = new Label();
            lowVolumeTrackBar = new TrackBar();
            lowVolumeLabel = new Label();
            processListBox = new ListBox();
            ((ISupportInitialize)lowVolumeTrackBar).BeginInit();
            SuspendLayout();
            // 
            // notifyIcon
            // 
            notifyIcon.Icon = (Icon)resources.GetObject("notifyIcon.Icon");
            notifyIcon.Text = "AudioAttenuation";
            notifyIcon.DoubleClick += notifyIcon_MouseDoubleClick;
            // 
            // refreshButton
            // 
            refreshButton.Anchor = AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            refreshButton.Location = new Point(7, 302);
            refreshButton.Margin = new Padding(2);
            refreshButton.MaximumSize = new Size(92, 23);
            refreshButton.MinimumSize = new Size(92, 23);
            refreshButton.Name = "refreshButton";
            refreshButton.Size = new Size(92, 23);
            refreshButton.TabIndex = 1;
            refreshButton.Text = "Refresh";
            refreshButton.UseVisualStyleBackColor = true;
            refreshButton.Click += refreshButton_Click;
            // 
            // delayTextBox
            // 
            delayTextBox.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
            delayTextBox.Location = new Point(443, 314);
            delayTextBox.Margin = new Padding(2, 1, 2, 1);
            delayTextBox.MaximumSize = new Size(83, 39);
            delayTextBox.Name = "delayTextBox";
            delayTextBox.Size = new Size(71, 23);
            delayTextBox.TabIndex = 9;
            delayTextBox.TextChanged += delayTextBox_TextChanged;
            // 
            // delayLabel
            // 
            delayLabel.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
            delayLabel.Location = new Point(374, 317);
            delayLabel.Margin = new Padding(2, 0, 2, 0);
            delayLabel.MaximumSize = new Size(65, 15);
            delayLabel.Name = "delayLabel";
            delayLabel.Size = new Size(65, 15);
            delayLabel.TabIndex = 5;
            delayLabel.Text = "Delay(ms)";
            delayLabel.TextAlign = ContentAlignment.MiddleCenter;
            // 
            // lowVolumeTrackBar
            // 
            lowVolumeTrackBar.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
            lowVolumeTrackBar.LargeChange = 10;
            lowVolumeTrackBar.Location = new Point(267, 292);
            lowVolumeTrackBar.Margin = new Padding(2, 1, 2, 1);
            lowVolumeTrackBar.Maximum = 100;
            lowVolumeTrackBar.Name = "lowVolumeTrackBar";
            lowVolumeTrackBar.Size = new Size(100, 45);
            lowVolumeTrackBar.TabIndex = 6;
            lowVolumeTrackBar.TickStyle = TickStyle.None;
            lowVolumeTrackBar.Scroll += lowVolumnTrackBar_Scroll;
            // 
            // lowVolumeLabel
            // 
            lowVolumeLabel.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
            lowVolumeLabel.Location = new Point(272, 321);
            lowVolumeLabel.Margin = new Padding(2, 0, 2, 0);
            lowVolumeLabel.Name = "lowVolumeLabel";
            lowVolumeLabel.Size = new Size(90, 16);
            lowVolumeLabel.TabIndex = 7;
            lowVolumeLabel.TextAlign = ContentAlignment.MiddleCenter;
            // 
            // processListBox
            // 
            processListBox.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            processListBox.DrawMode = DrawMode.OwnerDrawFixed;
            processListBox.ItemHeight = 64;
            processListBox.Location = new Point(7, 8);
            processListBox.Name = "processListBox";
            processListBox.Size = new Size(513, 260);
            processListBox.TabIndex = 1;
            processListBox.DrawItem += processListBox_DrawItem;
            processListBox.SelectedValueChanged += processListBox_SelectedValueChanged;
            // 
            // Form
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(525, 345);
            Controls.Add(processListBox);
            Controls.Add(lowVolumeLabel);
            Controls.Add(lowVolumeTrackBar);
            Controls.Add(delayLabel);
            Controls.Add(delayTextBox);
            Controls.Add(refreshButton);
            Margin = new Padding(2);
            MinimumSize = new Size(277, 255);
            Name = "Form";
            Text = "AudioAttenuation";
            FormClosing += Form_Closing;
            Load += Form_Load;
            Resize += Form_Resize;
            ((ISupportInitialize)lowVolumeTrackBar).EndInit();
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion
    }
}
