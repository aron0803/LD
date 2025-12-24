using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using OpenCvSharp;
using OpenCvSharp.Extensions;

namespace TraceLD
{
    public class MainForm : Form
    {
        // UI Controls
        private TextBox txtLdPath;
        private TextBox txtScreenshotDir;
        private NumericUpDown numWaitSeconds;
        private ListBox lstImages;
        private CheckedListBox chkLstEmulators;
        private Label lblStatus;
        private Button btnStart;
        private Button btnStop;
        private CheckBox chkDebugMode;

        // Logic
        private bool isRunning = false;
        private Thread workerThread;
        private LdPlayerManager ldManager;

        public MainForm()
        {
            InitializeComponent();
            LoadConfig();
            LoadImages();
            LoadEmulatorList();
        }

        private void InitializeComponent()
        {
            this.Text = "LDPlayer Trace 腳本 (C# .NET 4.0)";
            this.Size = new System.Drawing.Size(500, 650);
            this.StartPosition = FormStartPosition.CenterScreen;

            TableLayoutPanel mainLayout = new TableLayoutPanel();
            mainLayout.Dock = DockStyle.Fill;
            mainLayout.Padding = new Padding(10);
            mainLayout.RowCount = 5;
            mainLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize)); // Settings
            mainLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 350)); // Lists (Emu + Image)
            mainLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize)); // Action Controls
            mainLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize)); // Status
            this.Controls.Add(mainLayout);

            // --- Settings Group ---
            GroupBox grpSettings = new GroupBox { Text = "設定 (Settings)", Dock = DockStyle.Fill, AutoSize = true };
            mainLayout.Controls.Add(grpSettings, 0, 0);

            TableLayoutPanel settingsLayout = new TableLayoutPanel { Dock = DockStyle.Top, AutoSize = true, RowCount = 4, ColumnCount = 3 };
            settingsLayout.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            settingsLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            settingsLayout.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            grpSettings.Controls.Add(settingsLayout);

            // LD Path
            settingsLayout.Controls.Add(new Label { Text = "LDPlayer 路徑:", Anchor = AnchorStyles.Left, AutoSize = true }, 0, 0);
            txtLdPath = new TextBox { Anchor = AnchorStyles.Left | AnchorStyles.Right };
            settingsLayout.Controls.Add(txtLdPath, 1, 0);
            Button btnBrowseLd = new Button { Text = "瀏覽", AutoSize = true };
            btnBrowseLd.Click += (s, e) => {
                OpenFileDialog ofd = new OpenFileDialog { Filter = "Executable|*.exe" };
                if (ofd.ShowDialog() == DialogResult.OK) txtLdPath.Text = ofd.FileName;
            };
            settingsLayout.Controls.Add(btnBrowseLd, 2, 0);

            // Screenshot Dir
            settingsLayout.Controls.Add(new Label { Text = "截圖資料夾:", Anchor = AnchorStyles.Left, AutoSize = true }, 0, 1);
            txtScreenshotDir = new TextBox { Anchor = AnchorStyles.Left | AnchorStyles.Right };
            settingsLayout.Controls.Add(txtScreenshotDir, 1, 1);
            Button btnBrowseDir = new Button { Text = "瀏覽", AutoSize = true };
            btnBrowseDir.Click += (s, e) => {
                FolderBrowserDialog fbd = new FolderBrowserDialog();
                if (fbd.ShowDialog() == DialogResult.OK) txtScreenshotDir.Text = fbd.SelectedPath;
            };
            settingsLayout.Controls.Add(btnBrowseDir, 2, 1);

            // Wait Seconds
            settingsLayout.Controls.Add(new Label { Text = "比對間隔 (秒):", Anchor = AnchorStyles.Left, AutoSize = true }, 0, 2);
            numWaitSeconds = new NumericUpDown { Minimum = 0.1m, Maximum = 60, Increment = 0.5m, DecimalPlaces = 1, Value = 1.0m, Width = 60 };
            settingsLayout.Controls.Add(numWaitSeconds, 1, 2);

            // --- Lists Container (Side by Side) ---
            TableLayoutPanel listsContainer = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, RowCount = 1 };
            listsContainer.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));
            listsContainer.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));
            mainLayout.Controls.Add(listsContainer, 0, 1);

            // --- Emulator List (Left) ---
            GroupBox grpEmulators = new GroupBox { Text = "模擬器選擇", Dock = DockStyle.Fill };
            listsContainer.Controls.Add(grpEmulators, 0, 0);

            TableLayoutPanel emuLayout = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2 };
            emuLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            emuLayout.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            grpEmulators.Controls.Add(emuLayout);

            chkLstEmulators = new CheckedListBox { Dock = DockStyle.Fill, CheckOnClick = true };
            emuLayout.Controls.Add(chkLstEmulators, 0, 0);

            FlowLayoutPanel emuButtons = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.TopDown, AutoSize = true };
            emuLayout.Controls.Add(emuButtons, 1, 0);

            Button btnSelectAll = new Button { Text = "全選", Width = 70 };
            btnSelectAll.Click += (s, e) => { for (int i = 0; i < chkLstEmulators.Items.Count; i++) chkLstEmulators.SetItemChecked(i, true); };
            emuButtons.Controls.Add(btnSelectAll);

            Button btnDeselectAll = new Button { Text = "全不選", Width = 70 };
            btnDeselectAll.Click += (s, e) => { for (int i = 0; i < chkLstEmulators.Items.Count; i++) chkLstEmulators.SetItemChecked(i, false); };
            emuButtons.Controls.Add(btnDeselectAll);

            Button btnRefreshEmu = new Button { Text = "重新整理", Width = 70, Height = 30 };
            btnRefreshEmu.Click += (s, e) => LoadEmulatorList();
            emuButtons.Controls.Add(btnRefreshEmu);

            // --- Image List (Right) ---
            GroupBox grpImages = new GroupBox { Text = "比對清單 (優先順序)", Dock = DockStyle.Fill };
            listsContainer.Controls.Add(grpImages, 1, 0);

            TableLayoutPanel imgLayout = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2 };
            imgLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            imgLayout.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            grpImages.Controls.Add(imgLayout);

            lstImages = new ListBox { Dock = DockStyle.Fill, AllowDrop = true };
            lstImages.MouseDown += LstImages_MouseDown;
            lstImages.DragOver += LstImages_DragOver;
            lstImages.DragDrop += LstImages_DragDrop;
            imgLayout.Controls.Add(lstImages, 0, 0);

            FlowLayoutPanel imgButtons = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.TopDown, AutoSize = true };
            imgLayout.Controls.Add(imgButtons, 1, 0);

            Button btnUp = new Button { Text = "上移", Width = 60 };
            btnUp.Click += (s, e) => MoveItem(-1);
            imgButtons.Controls.Add(btnUp);

            Button btnDown = new Button { Text = "下移", Width = 60 };
            btnDown.Click += (s, e) => MoveItem(1);
            imgButtons.Controls.Add(btnDown);

            Button btnRefreshImg = new Button { Text = "整理", Width = 60 };
            btnRefreshImg.Click += (s, e) => LoadImages();
            imgButtons.Controls.Add(btnRefreshImg);

            // --- Action Controls ---
            FlowLayoutPanel pnlActions = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true, FlowDirection = FlowDirection.LeftToRight };
            mainLayout.Controls.Add(pnlActions, 0, 2);

            btnStart = new Button { Text = "開始 (Start)", Width = 100, Height = 40, BackColor = Color.LightGreen };
            btnStart.Click += (s, e) => StartScript();
            pnlActions.Controls.Add(btnStart);

            btnStop = new Button { Text = "停止 (Stop)", Width = 100, Height = 40, BackColor = Color.LightPink, Enabled = false };
            btnStop.Click += (s, e) => StopScript();
            pnlActions.Controls.Add(btnStop);

            Button btnSave = new Button { Text = "儲存設定", Width = 100, Height = 40 };
            btnSave.Click += (s, e) => SaveConfig();
            pnlActions.Controls.Add(btnSave);

            chkDebugMode = new CheckBox { Text = "Debug 模式", AutoSize = true, Margin = new Padding(10, 12, 0, 0) };
            pnlActions.Controls.Add(chkDebugMode);

            // --- Status Bar ---
            lblStatus = new Label { Text = "狀態：準備就緒", Dock = DockStyle.Fill, AutoSize = true, BorderStyle = BorderStyle.Fixed3D, TextAlign = ContentAlignment.MiddleLeft, Padding = new Padding(5) };
            mainLayout.Controls.Add(lblStatus, 0, 3);
        }

        #region Drag and Drop Support
        private void LstImages_MouseDown(object sender, MouseEventArgs e)
        {
            if (lstImages.SelectedItem == null) return;
            lstImages.DoDragDrop(lstImages.SelectedItem, DragDropEffects.Move);
        }

        private void LstImages_DragOver(object sender, DragEventArgs e)
        {
            e.Effect = DragDropEffects.Move;
        }

        private void LstImages_DragDrop(object sender, DragEventArgs e)
        {
            System.Drawing.Point point = lstImages.PointToClient(new System.Drawing.Point(e.X, e.Y));
            int index = lstImages.IndexFromPoint(point);
            if (index < 0) index = lstImages.Items.Count - 1;

            object data = e.Data.GetData(typeof(string));
            if (data != null)
            {
                lstImages.Items.Remove(data);
                lstImages.Items.Insert(index, data);
                lstImages.SelectedIndex = index;
            }
        }
        #endregion

        private void MoveItem(int direction)
        {
            if (lstImages.SelectedItem == null || lstImages.SelectedIndex < 0) return;

            int newIndex = lstImages.SelectedIndex + direction;
            if (newIndex < 0 || newIndex >= lstImages.Items.Count) return;

            object selected = lstImages.SelectedItem;
            lstImages.Items.Remove(selected);
            lstImages.Items.Insert(newIndex, selected);
            lstImages.SetSelected(newIndex, true);
        }

        private void LoadConfig()
        {
            string cfgLdPath = IniFile.ReadValue("Settings", "ld_path", "");
            string cfgScreenshotDir = IniFile.ReadValue("Settings", "screenshot_dir", "");
            
            string autoLdPath, autoScreenshotDir;
            AutoDetectPaths(out autoLdPath, out autoScreenshotDir);

            txtLdPath.Text = !string.IsNullOrEmpty(cfgLdPath) && File.Exists(cfgLdPath) ? cfgLdPath : (autoLdPath ?? @"D:\LDPlayer\LDPlayer9\ld.exe");
            txtScreenshotDir.Text = !string.IsNullOrEmpty(cfgScreenshotDir) && Directory.Exists(cfgScreenshotDir) ? cfgScreenshotDir : (autoScreenshotDir ?? @"D:\Screenshots");

            decimal wait = 1.0m;
            decimal.TryParse(IniFile.ReadValue("Settings", "wait_seconds", "1.0"), out wait);
            numWaitSeconds.Value = Math.Max(0.1m, Math.Min(60m, wait));

            chkDebugMode.Checked = IniFile.ReadValue("Settings", "debug_mode", "False").Equals("True", StringComparison.OrdinalIgnoreCase);
        }

        private bool AutoDetectPaths(out string ldPath, out string screenshotDir)
        {
            ldPath = null;
            screenshotDir = null;
            bool foundAny = false;

            string[] commonPaths = {
                @"C:\LDPlayer\LDPlayer9\ld.exe",
                @"D:\LDPlayer\LDPlayer9\ld.exe",
                @"E:\LDPlayer\LDPlayer9\ld.exe",
                @"F:\LDPlayer\LDPlayer9\ld.exe",
                @"C:\Program Files\LDPlayer\LDPlayer9\ld.exe",
                @"C:\XuanZhi\LDPlayer9\ld.exe"
            };

            foreach (string path in commonPaths)
            {
                if (File.Exists(path)) { ldPath = path; foundAny = true; break; }
            }

            if (string.IsNullOrEmpty(ldPath))
            {
                string pathEnv = Environment.GetEnvironmentVariable("PATH");
                if (!string.IsNullOrEmpty(pathEnv))
                {
                    foreach (string p in pathEnv.Split(Path.PathSeparator))
                    {
                        string potentialPath = Path.Combine(p, "ld.exe");
                        if (File.Exists(potentialPath)) { ldPath = potentialPath; foundAny = true; break; }
                    }
                }
            }

            if (!string.IsNullOrEmpty(ldPath))
            {
                try
                {
                    string ldDir = Path.GetDirectoryName(ldPath);
                    string configPath = Path.Combine(ldDir, "leidian.config");
                    if (File.Exists(configPath))
                    {
                        string[] lines = File.ReadAllLines(configPath);
                        foreach (string line in lines)
                        {
                            if (line.Contains("\"picturePath\""))
                            {
                                int start = line.IndexOf(":");
                                if (start > 0)
                                {
                                    string val = line.Substring(start + 1).Trim().Trim('"', ',', ' ').Replace("\\\\", "\\");
                                    if (Directory.Exists(val)) { screenshotDir = val; foundAny = true; break; }
                                }
                            }
                        }
                    }
                }
                catch { }
            }

            if (string.IsNullOrEmpty(screenshotDir))
            {
                string myDocuments = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                string xuanZhiPictures = Path.Combine(myDocuments, "XuanZhi9", "Pictures");
                if (Directory.Exists(xuanZhiPictures)) { screenshotDir = xuanZhiPictures; foundAny = true; }
            }

            if (string.IsNullOrEmpty(screenshotDir))
            {
                string myPictures = Environment.GetFolderPath(Environment.SpecialFolder.MyPictures);
                string ldPictures = Path.Combine(myPictures, "LDPlayer");
                if (Directory.Exists(ldPictures)) { screenshotDir = ldPictures; foundAny = true; }
            }

            return foundAny;
        }

        private void SaveConfig()
        {
            IniFile.WriteValue("Settings", "ld_path", txtLdPath.Text);
            IniFile.WriteValue("Settings", "screenshot_dir", txtScreenshotDir.Text);
            IniFile.WriteValue("Settings", "wait_seconds", numWaitSeconds.Value.ToString());
            IniFile.WriteValue("Settings", "debug_mode", chkDebugMode.Checked.ToString());

            List<string> selectedEmus = new List<string>();
            foreach (var item in chkLstEmulators.CheckedItems) selectedEmus.Add(item.ToString());
            IniFile.WriteValue("Settings", "selected_emulators", string.Join("|", selectedEmus.ToArray()));

            List<string> items = new List<string>();
            foreach (var item in lstImages.Items) items.Add(item.ToString());
            IniFile.WriteValue("Settings", "image_order", string.Join("|", items.ToArray()));

            MessageBox.Show("設定已儲存");
        }

        private void LoadImages()
        {
            string traceDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "trace");
            if (!Directory.Exists(traceDir))
            {
                try { Directory.CreateDirectory(traceDir); } catch { }
                return;
            }

            string savedOrder = IniFile.ReadValue("Settings", "image_order", "");
            string[] savedItems = savedOrder.Split(new char[] { '|' }, StringSplitOptions.RemoveEmptyEntries);
            
            HashSet<string> currentFiles = new HashSet<string>();
            foreach (string file in Directory.GetFiles(traceDir, "*.png"))
            {
                currentFiles.Add(Path.GetFileName(file));
            }

            lstImages.Items.Clear();
            foreach (string item in savedItems)
            {
                if (currentFiles.Contains(item))
                {
                    lstImages.Items.Add(item);
                    currentFiles.Remove(item);
                }
            }
            foreach (string file in currentFiles)
            {
                lstImages.Items.Add(file);
            }
        }

        private void UpdateStatus(string msg)
        {
            if (this.InvokeRequired)
            {
                this.Invoke(new Action<string>(UpdateStatus), msg);
            }
            else
            {
                lblStatus.Text = "狀態：" + msg;
            }
        }

        private void LoadEmulatorList()
        {
            if (!File.Exists(txtLdPath.Text)) return;

            string ldDir = Path.GetDirectoryName(txtLdPath.Text);
            string consolePath = Path.Combine(ldDir, "ldconsole.exe");
            if (!File.Exists(consolePath)) consolePath = txtLdPath.Text;

            try
            {
                ProcessStartInfo psi = new ProcessStartInfo
                {
                    FileName = consolePath,
                    Arguments = "list2",
                    UseShellExecute = false,
                    RedirectStandardOutput = true,
                    CreateNoWindow = true,
                    StandardOutputEncoding = Encoding.GetEncoding("UTF-8")
                };

                using (Process p = Process.Start(psi))
                {
                    string output = p.StandardOutput.ReadToEnd();
                    p.WaitForExit(3000);

                    string[] lines = output.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.RemoveEmptyEntries);

                    chkLstEmulators.Items.Clear();
                    foreach (string line in lines)
                    {
                        // Format: index,title,top_hwnd,bind_hwnd,android_state,pid,vbox_pid
                        string[] parts = line.Split(',');
                        if (parts.Length >= 5)
                        {
                            string itemText = string.Format("{0}: {1}", parts[0], parts[1]);
                            bool isRunning = parts[4] == "1";
                            
                            chkLstEmulators.Items.Add(itemText);
                            if (isRunning)
                            {
                                chkLstEmulators.SetItemChecked(chkLstEmulators.Items.Count - 1, true);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("讀取模擬器清單失敗: " + ex.Message);
            }
        }

        private bool CheckEmulatorStatus(int index)
        {
            if (!File.Exists(txtLdPath.Text)) return false;

            string ldDir = Path.GetDirectoryName(txtLdPath.Text);
            string consolePath = Path.Combine(ldDir, "ldconsole.exe");
            if (!File.Exists(consolePath)) consolePath = txtLdPath.Text;

            try
            {
                ProcessStartInfo psi = new ProcessStartInfo
                {
                    FileName = consolePath,
                    Arguments = "list2",
                    UseShellExecute = false,
                    RedirectStandardOutput = true,
                    CreateNoWindow = true,
                    StandardOutputEncoding = Encoding.GetEncoding("UTF-8")
                };

                using (Process p = Process.Start(psi))
                {
                    string output = p.StandardOutput.ReadToEnd();
                    p.WaitForExit(3000);

                    string[] lines = output.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.RemoveEmptyEntries);
                    foreach (string line in lines)
                    {
                        string[] parts = line.Split(',');
                        if (parts.Length >= 5 && parts[0] == index.ToString())
                        {
                            return parts[4] == "1";
                        }
                    }
                }
            }
            catch { }
            return false;
        }

        private void StartScript()
        {
            if (isRunning) return;

            if (!File.Exists(txtLdPath.Text))
            {
                MessageBox.Show("找不到 LDPlayer 執行檔 (ld.exe)！", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (chkLstEmulators.CheckedItems.Count == 0)
            {
                MessageBox.Show("請至少勾選一個模擬器！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            Logger.IsDebugMode = chkDebugMode.Checked;
            Logger.ClearLog();
            
            UpdateStatus("載入圖片快取...");
            ImageUtils.LoadTemplateCache();
            
            ldManager = new LdPlayerManager(txtLdPath.Text, txtScreenshotDir.Text, 0);
            isRunning = true;
            btnStart.Enabled = false;
            btnStop.Enabled = true;
            
            workerThread = new Thread(RunAutomation);
            workerThread.IsBackground = true;
            workerThread.Start();
        }

        private void StopScript()
        {
            isRunning = false;
            btnStart.Enabled = true;
            btnStop.Enabled = false;
            UpdateStatus("已停止");
        }

        private void RunAutomation()
        {
            try
            {
                UpdateStatus("開始執行...");
                
                while (isRunning)
                {
                    List<string> images = new List<string>();
                    List<int> selectedIndices = new List<int>();
                    double waitSec = 1.0;

                    this.Invoke(new Action(() => {
                        foreach (var item in lstImages.Items) images.Add(item.ToString());
                        foreach (var item in chkLstEmulators.CheckedItems)
                        {
                            string text = item.ToString();
                            int colonIdx = text.IndexOf(':');
                            if (colonIdx > 0)
                            {
                                int idx;
                                if (int.TryParse(text.Substring(0, colonIdx), out idx)) selectedIndices.Add(idx);
                            }
                        }
                        waitSec = (double)numWaitSeconds.Value;
                    }));

                    if (selectedIndices.Count == 0)
                    {
                        UpdateStatus("未選擇任何模擬器");
                        break;
                    }

                    foreach (int emuIndex in selectedIndices)
                    {
                        if (!isRunning) break;
                        ldManager.Index = emuIndex;
                        UpdateStatus(string.Format("[{0}] 截圖中...", emuIndex));

                        Bitmap capBitmap = ldManager.Screencap();
                        if (capBitmap != null)
                        {
                            using (Mat capMat = BitmapConverter.ToMat(capBitmap))
                            using (Mat capBGR = ImageUtils.ConvertToBGR(capMat, true))
                            {
                                capBitmap.Dispose();
                                
                                bool matched = false;
                                System.Drawing.Rectangle searchRange = new System.Drawing.Rectangle(180, 120, 430, 60);

                                foreach (string imgName in images)
                                {
                                    UpdateStatus(string.Format("[{0}] 比對: {1}", emuIndex, imgName));
                                    MatchResult result = ImageUtils.FindImage(capBGR, imgName, 0.90, searchRange);
                                    if (result.Success)
                                    {
                                        UpdateStatus(string.Format("[{0}] 匹配: {1} -> 點擊 ({2}, 320)", emuIndex, imgName, result.Location.X));
                                        ldManager.Click(result.Location.X, 320);
                                        matched = true;
                                        break;
                                    }
                                }
                                
                                if (!matched) UpdateStatus(string.Format("[{0}] 無匹配項", emuIndex));
                            }
                        }
                        else
                        {
                            UpdateStatus(string.Format("[{0}] 截圖失敗，檢查狀態...", emuIndex));
                            if (!CheckEmulatorStatus(emuIndex))
                            {
                                UpdateStatus(string.Format("[{0}] 模擬器已關閉，取消勾選", emuIndex));
                                this.Invoke(new Action(() => {
                                    for (int i = 0; i < chkLstEmulators.Items.Count; i++)
                                    {
                                        if (chkLstEmulators.Items[i].ToString().StartsWith(emuIndex + ":"))
                                        {
                                            chkLstEmulators.SetItemChecked(i, false);
                                            break;
                                        }
                                    }
                                }));
                            }
                        }
                    }

                    if (isRunning) Thread.Sleep((int)(waitSec * 1000));
                }
            }
            catch (Exception ex)
            {
                Logger.Log("RunAutomation Error: " + ex.ToString());
                UpdateStatus("錯誤: " + ex.Message);
                isRunning = false;
                this.Invoke(new Action(() => {
                    btnStart.Enabled = true;
                    btnStop.Enabled = false;
                }));
            }
        }

        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new MainForm());
        }
    }

    public class LdPlayerManager
    {
        public string LdPath { get; private set; }
        public string ScreenshotDir { get; private set; }
        public int Index { get; set; }

        public LdPlayerManager(string path, string dir, int index)
        {
            LdPath = path;
            ScreenshotDir = dir;
            Index = index;
        }

        public void Click(int x, int y)
        {
            Logger.Log(string.Format("[Click] ({0}, {1})", x, y));
            RunCommand("input", "tap", x.ToString(), y.ToString());
        }

        public Bitmap Screencap()
        {
            string filename = string.Format("trace_cap_{0}.png", Index);
            string localPath = Path.Combine(ScreenshotDir, filename);
            string remotePath = "/sdcard/Pictures/" + filename;
            
            try
            {
                RunCommand("screencap", remotePath);
                int retries = 50; // 5 seconds
                while (retries > 0)
                {
                    if (File.Exists(localPath))
                    {
                        using (var fs = new FileStream(localPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                        {
                            return new Bitmap(fs);
                        }
                    }
                    Thread.Sleep(100);
                    retries--;
                }
                return null;
            }
            catch { return null; }
        }

        private void RunCommand(params string[] args)
        {
            try
            {
                ProcessStartInfo psi = new ProcessStartInfo();
                psi.FileName = LdPath;
                psi.Arguments = string.Format("-s {0} {1}", Index, string.Join(" ", args));
                psi.UseShellExecute = false;
                psi.CreateNoWindow = true;
                using (Process p = Process.Start(psi)) { p.WaitForExit(5000); }
            }
            catch (Exception ex) { Logger.Log("RunCommand Error: " + ex.Message); }
        }
    }

    public struct MatchResult
    {
        public bool Success;
        public System.Drawing.Point Location;
        public double Score;
    }

    public static class ImageUtils
    {
        private static Dictionary<string, Mat> templateCache = new Dictionary<string, Mat>();
        private static object cacheLock = new object();

        public static void LoadTemplateCache()
        {
            lock (cacheLock)
            {
                foreach (var kvp in templateCache) if (kvp.Value != null) kvp.Value.Dispose();
                templateCache.Clear();
                
                string traceDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "trace");
                if (Directory.Exists(traceDir))
                {
                    foreach (string file in Directory.GetFiles(traceDir, "*.png"))
                    {
                        try
                        {
                            using (Bitmap bmp = new Bitmap(file))
                            using (Mat mat = BitmapConverter.ToMat(bmp))
                            {
                                templateCache[Path.GetFileName(file)] = ConvertToBGR(mat, true);
                            }
                        }
                        catch { }
                    }
                }
            }
        }

        public static Mat ConvertToBGR(Mat input, bool returnClone = false)
        {
            if (input.Channels() == 3) return returnClone ? input.Clone() : input;
            Mat output = new Mat();
            if (input.Channels() == 4) Cv2.CvtColor(input, output, ColorConversionCodes.BGRA2BGR);
            else if (input.Channels() == 1) Cv2.CvtColor(input, output, ColorConversionCodes.GRAY2BGR);
            return output;
        }

        public static MatchResult FindImage(Mat sourceMat, string templateName, double threshold, System.Drawing.Rectangle searchRegion = default(System.Drawing.Rectangle))
        {
            Mat templateMat = null;
            lock (cacheLock) { if (templateCache.ContainsKey(templateName)) templateMat = templateCache[templateName]; }
            if (templateMat == null) return new MatchResult { Success = false };

            // 如果有指定搜尋區域，則進行裁切
            Mat processedSource = sourceMat;
            int offsetX = 0;
            int offsetY = 0;

            if (searchRegion != default(System.Drawing.Rectangle) && searchRegion.Width > 0 && searchRegion.Height > 0)
            {
                if (searchRegion.X >= 0 && searchRegion.Y >= 0 && 
                    searchRegion.X + searchRegion.Width <= sourceMat.Width &&
                    searchRegion.Y + searchRegion.Height <= sourceMat.Height)
                {
                    offsetX = searchRegion.X;
                    offsetY = searchRegion.Y;
                    OpenCvSharp.Rect roi = new OpenCvSharp.Rect(searchRegion.X, searchRegion.Y, searchRegion.Width, searchRegion.Height);
                    processedSource = new Mat(sourceMat, roi);
                }
            }

            using (Mat result = new Mat())
            {
                Mat sourceProcessed = ConvertToBGR(processedSource, false);
                try
                {
                    Cv2.MatchTemplate(sourceProcessed, templateMat, result, TemplateMatchModes.CCoeffNormed);
                    double minVal, maxVal;
                    OpenCvSharp.Point minLoc, maxLoc;
                    Cv2.MinMaxLoc(result, out minVal, out maxVal, out minLoc, out maxLoc);
                    
                    bool success = maxVal >= threshold;
                    return new MatchResult 
                    { 
                        Success = success, 
                        Location = new System.Drawing.Point(maxLoc.X + offsetX, maxLoc.Y + offsetY), 
                        Score = maxVal 
                    };
                }
                finally 
                { 
                    if (sourceProcessed != processedSource) sourceProcessed.Dispose();
                    if (processedSource != sourceMat) processedSource.Dispose();
                }
            }
        }
    }

    public static class Logger
    {
        private static string logPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "trace_log.txt");
        public static bool IsDebugMode { get; set; }
        public static void ClearLog() { try { if (File.Exists(logPath)) File.Delete(logPath); } catch { } }
        public static void Log(string message)
        {
            if (!IsDebugMode) return;
            try { File.AppendAllText(logPath, string.Format("[{0}] {1}\r\n", DateTime.Now.ToString("HH:mm:ss.fff"), message)); } catch { }
        }
    }

    public static class IniFile
    {
        [DllImport("kernel32", CharSet = CharSet.Unicode)]
        private static extern long WritePrivateProfileString(string section, string key, string val, string filePath);
        [DllImport("kernel32", CharSet = CharSet.Unicode)]
        private static extern int GetPrivateProfileString(string section, string key, string def, StringBuilder retVal, int size, string filePath);

        public static string ReadValue(string section, string key, string def)
        {
            StringBuilder temp = new StringBuilder(1024);
            string path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "trace.ini");
            GetPrivateProfileString(section, key, def, temp, 1024, path);
            return temp.ToString();
        }

        public static void WriteValue(string section, string key, string value)
        {
            string path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "trace.ini");
            WritePrivateProfileString(section, key, value, path);
        }
    }
}
