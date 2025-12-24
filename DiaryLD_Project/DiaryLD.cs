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

namespace DiaryLD
{
    public class MainForm : Form
    {
        // UI Controls
        private TextBox txtLdPath;
        private TextBox txtScreenshotDir;
        private NumericUpDown numIndex;
        private NumericUpDown numDelayMultiplier;
        private FlowLayoutPanel pnlImages;
        private Label lblStatus;
        private Button btnStart;
        private Button btnStop;
        private CheckBox chkDebugMode;
        private Dictionary<string, CheckBox> imageCheckBoxes = new Dictionary<string, CheckBox>();

        // Logic
        private bool isRunning = false;
        private Thread workerThread;
        private LdPlayerManager ldManager;

        public MainForm()
        {
            InitializeComponent();
            LoadConfig();
            LoadImages();
        }

        private void InitializeComponent()
        {
            this.Text = "LDPlayer 自動化腳本 (C# .NET 4.0)";
            this.Size = new System.Drawing.Size(600, 550);
            this.MinimumSize = new System.Drawing.Size(600, 100); // Allow height to shrink if needed
            this.StartPosition = FormStartPosition.CenterScreen;
            this.AutoSize = true;
            this.AutoSizeMode = AutoSizeMode.GrowAndShrink;

            // Main Layout
            TableLayoutPanel mainLayout = new TableLayoutPanel();
            mainLayout.Dock = DockStyle.Fill;
            mainLayout.Padding = new Padding(10);
            mainLayout.RowCount = 4;
            mainLayout.AutoSize = true;
            mainLayout.AutoSizeMode = AutoSizeMode.GrowAndShrink;
            mainLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize)); // Settings
            mainLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize)); // Images (Changed from Percent to AutoSize)
            mainLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize)); // Controls
            mainLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize)); // Status
            this.Controls.Add(mainLayout);

            // --- Settings Group ---
            GroupBox grpSettings = new GroupBox();
            grpSettings.Text = "設定 (Settings)";
            grpSettings.Dock = DockStyle.Fill;
            grpSettings.AutoSize = true;
            grpSettings.AutoSizeMode = AutoSizeMode.GrowAndShrink;
            mainLayout.Controls.Add(grpSettings, 0, 0);

            TableLayoutPanel settingsLayout = new TableLayoutPanel();
            settingsLayout.Dock = DockStyle.Top;
            settingsLayout.AutoSize = true;
            settingsLayout.AutoSizeMode = AutoSizeMode.GrowAndShrink;
            settingsLayout.RowCount = 3;
            settingsLayout.ColumnCount = 3;
            settingsLayout.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            settingsLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            settingsLayout.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            grpSettings.Controls.Add(settingsLayout);

            // LD Path
            settingsLayout.Controls.Add(new Label { Text = "LDPlayer 路徑:", Anchor = AnchorStyles.Left | AnchorStyles.Right, AutoSize = true, TextAlign = ContentAlignment.MiddleLeft }, 0, 0);
            txtLdPath = new TextBox { Anchor = AnchorStyles.Left | AnchorStyles.Right };
            settingsLayout.Controls.Add(txtLdPath, 1, 0);
            Button btnBrowseLd = new Button { Text = "瀏覽", AutoSize = true };
            btnBrowseLd.Click += (s, e) => {
                OpenFileDialog ofd = new OpenFileDialog { Filter = "Executable|*.exe" };
                if (ofd.ShowDialog() == DialogResult.OK) txtLdPath.Text = ofd.FileName;
            };
            settingsLayout.Controls.Add(btnBrowseLd, 2, 0);

            // Screenshot Dir
            settingsLayout.Controls.Add(new Label { Text = "截圖資料夾:", Anchor = AnchorStyles.Left | AnchorStyles.Right, AutoSize = true, TextAlign = ContentAlignment.MiddleLeft }, 0, 1);
            txtScreenshotDir = new TextBox { Anchor = AnchorStyles.Left | AnchorStyles.Right };
            settingsLayout.Controls.Add(txtScreenshotDir, 1, 1);
            Button btnBrowseDir = new Button { Text = "瀏覽", AutoSize = true };
            btnBrowseDir.Click += (s, e) => {
                FolderBrowserDialog fbd = new FolderBrowserDialog();
                if (fbd.ShowDialog() == DialogResult.OK) txtScreenshotDir.Text = fbd.SelectedPath;
            };
            settingsLayout.Controls.Add(btnBrowseDir, 2, 1);

            // Index
            settingsLayout.Controls.Add(new Label { Text = "模擬器 Index (0~30):", Anchor = AnchorStyles.Left | AnchorStyles.Right, AutoSize = true, TextAlign = ContentAlignment.MiddleLeft }, 0, 2);
            numIndex = new NumericUpDown { Minimum = 0, Maximum = 30, Width = 60 };
            numIndex.ValueChanged += (s, e) => SaveIndexToConfig(s, e);
            settingsLayout.Controls.Add(numIndex, 1, 2);

            // Delay Multiplier
            settingsLayout.RowCount = 4; // Increase row count
            settingsLayout.Controls.Add(new Label { Text = "延遲倍率 (0.9~1.5):", Anchor = AnchorStyles.Left | AnchorStyles.Right, AutoSize = true, TextAlign = ContentAlignment.MiddleLeft }, 0, 3);
            numDelayMultiplier = new NumericUpDown { Minimum = 0.9m, Maximum = 1.5m, Increment = 0.1m, DecimalPlaces = 1, Value = 1.0m, Width = 60 };
            settingsLayout.Controls.Add(numDelayMultiplier, 1, 3);

            // --- Target Images Group ---
            GroupBox grpImages = new GroupBox();
            grpImages.Text = "目標圖片 (Target Images)";
            grpImages.Dock = DockStyle.Fill;
            grpImages.AutoSize = true;
            grpImages.AutoSizeMode = AutoSizeMode.GrowAndShrink;
            // Force width to fit 5 items (150+10 margin)*5 = 800 + padding
            grpImages.MinimumSize = new System.Drawing.Size(830, 0);
            grpImages.MaximumSize = new System.Drawing.Size(830, 9999); // Constrain width, allow height expansion
            mainLayout.Controls.Add(grpImages, 0, 1);

            pnlImages = new FlowLayoutPanel();
            pnlImages.Dock = DockStyle.Top; // Dock Top to allow AutoSize to grow downwards
            pnlImages.AutoSize = true;
            pnlImages.AutoSizeMode = AutoSizeMode.GrowAndShrink;
            pnlImages.WrapContents = true;
            pnlImages.MaximumSize = new System.Drawing.Size(820, 0); // Constrain width to force wrap, 0 means unlimited height
            // pnlImages.AutoScroll = true; // No longer needed if we auto-size the window
            grpImages.Controls.Add(pnlImages);

            // --- Control Group ---
            FlowLayoutPanel pnlControls = new FlowLayoutPanel();
            pnlControls.Dock = DockStyle.Fill;
            pnlControls.AutoSize = true;
            pnlControls.FlowDirection = FlowDirection.LeftToRight;
            // pnlControls.Alignment = FlowDirection.Center; // Property does not exist in WinForms
            mainLayout.Controls.Add(pnlControls, 0, 2);

            btnStart = new Button { Text = "開始執行 (Start)", Width = 120, Height = 40, BackColor = Color.LightGreen };
            btnStart.Click += (s, e) => StartScript();
            pnlControls.Controls.Add(btnStart);

            btnStop = new Button { Text = "停止執行 (Stop)", Width = 120, Height = 40, BackColor = Color.LightPink, Enabled = false };
            btnStop.Click += (s, e) => StopScript();
            pnlControls.Controls.Add(btnStop);

            Button btnSave = new Button { Text = "儲存設定 (Save)", Width = 120, Height = 40 };
            btnSave.Click += (s, e) => SaveConfig();
            pnlControls.Controls.Add(btnSave);

            chkDebugMode = new CheckBox { Text = "Debug 模式 (記錄比對)", AutoSize = true, Margin = new Padding(10, 10, 5, 5) };
            pnlControls.Controls.Add(chkDebugMode);

            // --- Status Bar ---
            lblStatus = new Label { Text = "狀態：準備就緒", Dock = DockStyle.Fill, AutoSize = true, BorderStyle = BorderStyle.Fixed3D, TextAlign = ContentAlignment.MiddleLeft, Padding = new Padding(5) };
            mainLayout.Controls.Add(lblStatus, 0, 3);
        }

        private void LoadConfig()
        {
            string cfgLdPath = IniFile.ReadValue("Settings", "ld_path", "");
            string cfgScreenshotDir = IniFile.ReadValue("Settings", "screenshot_dir", "");
            
            // Auto-detect if config is empty or invalid
            string autoLdPath, autoScreenshotDir;
            bool autoDetected = AutoDetectPaths(out autoLdPath, out autoScreenshotDir);

            // Logic for LD Path: Config > Auto > Default
            if (!string.IsNullOrEmpty(cfgLdPath) && File.Exists(cfgLdPath))
            {
                txtLdPath.Text = cfgLdPath;
            }
            else if (!string.IsNullOrEmpty(autoLdPath))
            {
                txtLdPath.Text = autoLdPath;
                if (string.IsNullOrEmpty(cfgLdPath)) 
                    UpdateStatus("已自動偵測到 LDPlayer 路徑");
                else
                    UpdateStatus("設定路徑無效，已切換至自動偵測路徑");
            }
            else
            {
                txtLdPath.Text = @"D:\LDPlayer\LDPlayer9\ld.exe"; // Fallback default
            }

            // Logic for Screenshot Dir: Config > Auto > Default
            if (!string.IsNullOrEmpty(cfgScreenshotDir) && Directory.Exists(cfgScreenshotDir))
            {
                txtScreenshotDir.Text = cfgScreenshotDir;
            }
            else if (!string.IsNullOrEmpty(autoScreenshotDir))
            {
                txtScreenshotDir.Text = autoScreenshotDir;
            }
            else
            {
                txtScreenshotDir.Text = @"D:\Screenshots"; // Fallback default
            }

            int idx = 0;
            int.TryParse(IniFile.ReadValue("Settings", "ld_index", "0"), out idx);
            
            // 暫時移除事件處理器，避免在載入時觸發儲存
            numIndex.ValueChanged -= SaveIndexToConfig;
            numIndex.Value = idx;
            numIndex.ValueChanged += SaveIndexToConfig;

            // Load Debug Mode (Default to FALSE as requested)
            string debugModeStr = IniFile.ReadValue("Settings", "debug_mode", "False");
            bool debugMode;
            if (bool.TryParse(debugModeStr, out debugMode))
            {
                chkDebugMode.Checked = debugMode;
            }
            else
            {
                chkDebugMode.Checked = false;
            }

            // Load Delay Multiplier
            string delayMultStr = IniFile.ReadValue("Settings", "delay_multiplier", "1.0");
            decimal delayMult;
            if (decimal.TryParse(delayMultStr, out delayMult))
            {
                numDelayMultiplier.Value = Math.Max(0.9m, Math.Min(1.5m, delayMult));
            }
            else
            {
                numDelayMultiplier.Value = 1.0m;
            }
        }

        private bool AutoDetectPaths(out string ldPath, out string screenshotDir)
        {
            ldPath = null;
            screenshotDir = null;
            bool foundAny = false;

            // 1. Try to find LDPlayer path
            // Check common installation paths
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
                if (File.Exists(path))
                {
                    ldPath = path;
                    foundAny = true;
                    break;
                }
            }

            // If not found in common paths, check PATH environment variable
            if (string.IsNullOrEmpty(ldPath))
            {
                string pathEnv = Environment.GetEnvironmentVariable("PATH");
                if (!string.IsNullOrEmpty(pathEnv))
                {
                    string[] paths = pathEnv.Split(Path.PathSeparator);
                    foreach (string p in paths)
                    {
                        string potentialPath = Path.Combine(p, "ld.exe");
                        if (File.Exists(potentialPath))
                        {
                            ldPath = potentialPath;
                            foundAny = true;
                            break;
                        }
                    }
                }
            }

            // 2. Try to find/guess Screenshot directory
            
            // Priority 1: Check leidian.config if we found the LDPlayer path
            if (!string.IsNullOrEmpty(ldPath))
            {
                try
                {
                    string ldDir = Path.GetDirectoryName(ldPath);
                    string configPath = Path.Combine(ldDir, "leidian.config");
                    if (File.Exists(configPath))
                    {
                        // Simple parsing for "picturePath" : "..."
                        // The file is usually JSON-like
                        string[] lines = File.ReadAllLines(configPath);
                        foreach (string line in lines)
                        {
                            if (line.Contains("\"picturePath\""))
                            {
                                int start = line.IndexOf(":");
                                if (start > 0)
                                {
                                    string val = line.Substring(start + 1).Trim().Trim('"', ',', ' ');
                                    // Handle escaped backslashes if present (e.g. C:\\Users...)
                                    val = val.Replace("\\\\", "\\");
                                    if (Directory.Exists(val))
                                    {
                                        screenshotDir = val;
                                        foundAny = true;
                                        UpdateStatus("從 leidian.config 讀取到截圖路徑");
                                        break;
                                    }
                                }
                            }
                        }
                    }
                }
                catch { /* Ignore config parsing errors */ }
            }

            // Priority 2: Check Documents\XuanZhi9\Pictures
            if (string.IsNullOrEmpty(screenshotDir))
            {
                string myDocuments = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                string xuanZhiPictures = Path.Combine(myDocuments, "XuanZhi9", "Pictures");
                if (Directory.Exists(xuanZhiPictures))
                {
                    screenshotDir = xuanZhiPictures;
                    foundAny = true;
                }
            }

            // Priority 3: Check MyPictures\LDPlayer
            if (string.IsNullOrEmpty(screenshotDir))
            {
                string myPictures = Environment.GetFolderPath(Environment.SpecialFolder.MyPictures);
                string ldPictures = Path.Combine(myPictures, "LDPlayer");
                if (Directory.Exists(ldPictures))
                {
                    screenshotDir = ldPictures;
                    foundAny = true;
                }
            }

            // Priority 4: Check D:\Screenshots (Legacy/User specific)
            if (string.IsNullOrEmpty(screenshotDir))
            {
                if (Directory.Exists(@"D:\Screenshots"))
                {
                    screenshotDir = @"D:\Screenshots";
                    foundAny = true;
                }
            }

            return foundAny;
        }

        private void SaveIndexToConfig(object sender, EventArgs e)
        {
            // 自動儲存 index 到 config.ini
            IniFile.WriteValue("Settings", "ld_index", numIndex.Value.ToString());
        }

        private void SaveConfig()
        {
            IniFile.WriteValue("Settings", "ld_path", txtLdPath.Text);
            IniFile.WriteValue("Settings", "screenshot_dir", txtScreenshotDir.Text);
            IniFile.WriteValue("Settings", "ld_index", numIndex.Value.ToString());
            IniFile.WriteValue("Settings", "debug_mode", chkDebugMode.Checked.ToString());
            IniFile.WriteValue("Settings", "delay_multiplier", numDelayMultiplier.Value.ToString());

            // Save selected images
            List<string> selected = new List<string>();
            foreach (var kvp in imageCheckBoxes)
            {
                if (kvp.Value.Checked) selected.Add(kvp.Key);
            }
            IniFile.WriteValue("Settings", "selected_images", string.Join("|", selected.ToArray()));

            MessageBox.Show("設定已儲存");
        }

        private void LoadImages()
        {
            pnlImages.Controls.Clear();
            imageCheckBoxes.Clear();
            string picDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "pic");
            if (!Directory.Exists(picDir))
            {
                try { Directory.CreateDirectory(picDir); } catch { }
                return;
            }

            // Load saved selections
            string savedStr = IniFile.ReadValue("Settings", "selected_images", "");
            HashSet<string> savedSet = new HashSet<string>(savedStr.Split(new char[] { '|' }, StringSplitOptions.RemoveEmptyEntries));

            foreach (string file in Directory.GetFiles(picDir, "*.png"))
            {
                string fileName = Path.GetFileName(file);
                CheckBox chk = new CheckBox();
                chk.Text = fileName;
                chk.Appearance = Appearance.Button; // Make it look like a button
                chk.AutoSize = false;
                chk.Size = new System.Drawing.Size(150, 40); // Fixed size for grid alignment
                chk.TextAlign = ContentAlignment.MiddleCenter;
                chk.Margin = new Padding(5);
                chk.FlatStyle = FlatStyle.Flat;
                chk.FlatAppearance.CheckedBackColor = Color.LightSkyBlue; // Highlight when checked
                
                if (savedSet.Contains(fileName)) chk.Checked = true;

                pnlImages.Controls.Add(chk);
                imageCheckBoxes[fileName] = chk;
            }
        }

        private void UpdateStatus(string msg)
        {
            try
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
            catch (Exception ex)
            {
                // Log UpdateStatus errors
                Logger.Log(string.Format("[UpdateStatus] 錯誤: {0}\r\n堆疊追蹤: {1}", ex.Message, ex.StackTrace));
            }
        }

        private void WriteDebugLog(string imageName, double score, bool success)
        {
            if (!chkDebugMode.Checked) return;
            Logger.Log(string.Format("[FindImage] 圖片: {0,-20} 分數: {1:F6}  結果: {2}",
                imageName, score, success ? "成功" : "失敗"));
        }

        private void StartScript()
        {
            if (isRunning) return;

            if (!File.Exists(txtLdPath.Text))
            {
                MessageBox.Show("找不到 LDPlayer 執行檔 (ld.exe)，請檢查路徑！", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // Clear old log file at the start of each run
            Logger.IsDebugMode = chkDebugMode.Checked; // Set debug mode flag
            Logger.ClearLog();
            
            // Pre-load all template images into cache for faster matching
            UpdateStatus("載入圖片快取...");
            ImageUtils.LoadTemplateCache();
            
            ldManager = new LdPlayerManager(txtLdPath.Text, txtScreenshotDir.Text, (int)numIndex.Value);
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
                UpdateStatus(string.Format("執行模擬器 Index: {0}", ldManager.Index));
                
                // Get delay multiplier from UI (thread-safe)
                double delayMultiplier = 1.0;
                this.Invoke(new Action(() => {
                    delayMultiplier = (double)numDelayMultiplier.Value;
                }));
                
                // Get selected images
                List<string> selectedImages = new List<string>();
                this.Invoke(new Action(() => {
                    foreach (var kvp in imageCheckBoxes)
                    {
                        if (kvp.Value.Checked) selectedImages.Add(kvp.Key);
                    }
                }));

                while (isRunning)
                {
                    UpdateStatus("正在掃描... (Scanning)");
                    
                    // 1. Click initial point
                    ldManager.Click(400, 380);
                    Thread.Sleep((int)(1000 * delayMultiplier));

                    // 2. Check diary.png - Convert to Mat once and reuse
                    Bitmap capBitmap = ldManager.Screencap();
                    if (capBitmap == null)
                    {
                        UpdateStatus("截圖失敗 (Screenshot Failed) - Check Path/Shared Folder");
                        Thread.Sleep((int)(1000 * delayMultiplier));
                        continue;
                    }

                    // Convert Bitmap -> Mat -> BGR once for all matches on this screenshot
                    using (Mat capMat = BitmapConverter.ToMat(capBitmap))
                    using (Mat capBGR = ImageUtils.ConvertToBGR(capMat, true))
                    {
                        capBitmap.Dispose();
                        
                        MatchResult result = ImageUtils.FindImage(capBGR, "diary.png", 0.90);
                        WriteDebugLog("diary.png", result.Score, result.Success);

                        if (result.Success)
                        {
                            UpdateStatus(string.Format("找到日記本 (Score: {0:F4})", result.Score));
                            ldManager.Click(result.Location.X, result.Location.Y);
                            Thread.Sleep((int)(1500 * delayMultiplier));
                        }
                        else
                        {
                            UpdateStatus(string.Format("掃描中... (Max Score: {0:F4})", result.Score));
                        }
                    }

                    // 3. Check ball.png
                    Thread.Sleep((int)(500 * delayMultiplier)); // Wait for screen transition
                    capBitmap = ldManager.Screencap();
                    if (capBitmap != null)
                    {
                        using (Mat capMat = BitmapConverter.ToMat(capBitmap))
                        using (Mat capBGR = ImageUtils.ConvertToBGR(capMat, true))
                        {
                            capBitmap.Dispose();
                            
                            MatchResult result = ImageUtils.FindImage(capBGR, "ball.png", 0.96);
                            WriteDebugLog("ball.png", result.Score, result.Success);

                            if (result.Success)
                            {
                                UpdateStatus(string.Format("找到神秘珠子 (ball.png Score: {0:F4})", result.Score));
                                ldManager.Click(result.Location.X, result.Location.Y);
                                Thread.Sleep((int)(1500 * delayMultiplier));
                            }
                            else
                            {
                                UpdateStatus(string.Format("未找到神秘珠子 (ball.png Max Score: {0:F4})", result.Score));
                            }
                        }
                    }

                    // 4. Check check1.png
                    Thread.Sleep((int)(500 * delayMultiplier)); // Wait for screen transition
                    capBitmap = ldManager.Screencap();
                    if (capBitmap != null)
                    {
                        using (Mat capMat = BitmapConverter.ToMat(capBitmap))
                        using (Mat capBGR = ImageUtils.ConvertToBGR(capMat, true))
                        {
                            capBitmap.Dispose();
                            
                            MatchResult result = ImageUtils.FindImage(capBGR, "check1.png", 0.98);
                            WriteDebugLog("check1.png", result.Score, result.Success);

                            if (result.Success)
                            {
                                UpdateStatus("確認");
                                ldManager.Click(470, 370);
                                Thread.Sleep((int)(1000 * delayMultiplier));
                            }
                        }
                    }

                    // 5. Check check2.png
                    Thread.Sleep((int)(500 * delayMultiplier)); // Wait for screen transition
                    capBitmap = ldManager.Screencap();
                    if (capBitmap != null)
                    {
                        using (Mat capMat = BitmapConverter.ToMat(capBitmap))
                        using (Mat capBGR = ImageUtils.ConvertToBGR(capMat, true))
                        {
                            capBitmap.Dispose();
                            
                            MatchResult result = ImageUtils.FindImage(capBGR, "check2.png", 0.98);
                            WriteDebugLog("check2.png", result.Score, result.Success);

                            if (result.Success)
                            {
                                // Inner loop for crafting
                                while (isRunning)
                                {
                                    bool foundAny = false;
                                    // Define equipment search region (265,196)-(536,246) for faster matching
                                    Rectangle equipmentRegion = new Rectangle(265, 196, 271, 50);
                                    
                                    // Convert screenshot once and check all equipment
                                    capBitmap = ldManager.Screencap();
                                    if (capBitmap == null) break;
                                    
                                    using (Mat capMat3 = BitmapConverter.ToMat(capBitmap))
                                    using (Mat capBGR3 = ImageUtils.ConvertToBGR(capMat3, true))
                                    {
                                        capBitmap.Dispose();
                                        
                                        // Verify check2.png exists before checking equipment
                                        MatchResult check2Res = ImageUtils.FindImage(capBGR3, "check2.png", 0.98);
                                        if (!check2Res.Success)
                                        {
                                            UpdateStatus("check2.png 消失，停止循環");
                                            break; // Exit the while loop
                                        }
                                        
                                        foreach (string pngFile in selectedImages)
                                        {
                                            UpdateStatus(string.Format("判斷 {0}", pngFile));
                                            
                                            MatchResult matchRes = ImageUtils.FindImage(capBGR3, pngFile, 0.98, equipmentRegion);
                                            WriteDebugLog(pngFile, matchRes.Score, matchRes.Success);

                                            if (matchRes.Success)
                                            {
                                                UpdateStatus(string.Format("製作 {0}", pngFile));
                                                foundAny = true;
                                                // Click craft button 6 times
                                                for (int i = 0; i < 6; i++)
                                                {
                                                    ldManager.Click(405, 357);
                                                    Thread.Sleep((int)(1000 * delayMultiplier));
                                                }
                                                break; 
                                            }
                                        }
                                    }

                                    if (foundAny) break; 

                                    UpdateStatus("變更");
                                    ldManager.Click(470, 300);
                                    Thread.Sleep((int)(1500 * delayMultiplier));
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                // Log detailed error information
                string errorMsg = string.Format("錯誤: {0}", ex.Message);
                Logger.Log(string.Format("[RunAutomation] 發生例外狀況\r\n訊息: {0}\r\n類型: {1}\r\n堆疊追蹤:\r\n{2}", 
                    ex.Message, ex.GetType().Name, ex.StackTrace));
                
                UpdateStatus(errorMsg);
                
                try
                {
                    this.Invoke(new Action(() => {
                        btnStart.Enabled = true;
                        btnStop.Enabled = false;
                    }));
                }
                catch (Exception invokeEx)
                {
                    Logger.Log(string.Format("[RunAutomation] UI 更新失敗: {0}", invokeEx.Message));
                }
                
                isRunning = false;
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
        public int Index { get; private set; }

        public LdPlayerManager(string path, string dir, int index)
        {
            LdPath = path;
            ScreenshotDir = dir;
            Index = index;
        }

        public void Click(int x, int y)
        {
            Logger.Log(string.Format("[Click] 座標: ({0}, {1})", x, y));
            RunCommand("input", "tap", x.ToString(), y.ToString());
        }

        public Bitmap Screencap()
        {
            string filename = string.Format("cap_{0}.png", Index);
            string localPath = Path.Combine(ScreenshotDir, filename);
            string remotePath = "/sdcard/Pictures/" + filename;

            Logger.Log(string.Format("[Screencap] 開始截圖 Index: {0}, 路徑: {1}", Index, remotePath));
            
            try
            {
                // Capture to sdcard
                RunCommand("screencap", remotePath);
                
                // Wait briefly for file to sync if using shared folder mechanism
                // Note: Python script assumes the file magically appears in ScreenshotDir.
                // This usually implies LDPlayer shared folder setup: /sdcard/Pictures -> D:\Screenshots
                
                int retries = 10;
                while (retries > 0)
                {
                    try
                    {
                        if (File.Exists(localPath))
                        {
                            Logger.Log(string.Format("[Screencap] 截圖成功: {0}", localPath));
                            // Load copy to avoid locking
                            using (var fs = new FileStream(localPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                            {
                                return new Bitmap(fs);
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Logger.Log(string.Format("[Screencap] 讀取檔案失敗 (重試 {0}/10): {1}", 11 - retries, ex.Message));
                    }
                    Thread.Sleep(100);
                    retries--;
                }
                
                Logger.Log(string.Format("[Screencap] 截圖失敗: 超時或檔案不存在 - {0}", localPath));
                return null; // Or throw exception
            }
            catch (Exception ex)
            {
                Logger.Log(string.Format("[Screencap] 截圖過程發生例外狀況\r\n訊息: {0}\r\n堆疊追蹤: {1}", ex.Message, ex.StackTrace));
                return null;
            }
        }

        private void RunCommand(params string[] args)
        {
            try
            {
                ProcessStartInfo psi = new ProcessStartInfo();
                psi.FileName = LdPath;
                string argsStr = string.Format("-s {0} {1}", Index, string.Join(" ", args));
                psi.Arguments = argsStr;
                psi.UseShellExecute = false;
                psi.CreateNoWindow = true;
                using (Process p = Process.Start(psi))
                {
                    p.WaitForExit();
                }
            }
            catch (Exception ex)
            {
                Logger.Log(string.Format("[RunCommand] 執行命令失敗\r\n命令: {0}\r\n訊息: {1}\r\n堆疊追蹤: {2}", 
                    string.Join(" ", args), ex.Message, ex.StackTrace));
            }
        }
    }

    public struct MatchResult
    {
        public bool Success;
        public System.Drawing.Point Location;  // Explicitly use System.Drawing.Point
        public double Score;
    }

    public static class ImageUtils
    {
        // Cache for pre-loaded template images (Mat format)
        // Key: template filename, Value: template Mat in BGR format
        private static Dictionary<string, Mat> templateCache = new Dictionary<string, Mat>();
        private static object cacheLock = new object();
        
        /// <summary>
        /// Pre-load all template images into memory as Mat objects
        /// Call this once at startup for better performance
        /// </summary>
        public static void LoadTemplateCache()
        {
            lock (cacheLock)
            {
                Logger.Log("[TemplateCache] 開始載入圖片快取...");
                
                // Clear existing cache
                foreach (var kvp in templateCache)
                {
                    if (kvp.Value != null) kvp.Value.Dispose();
                }
                templateCache.Clear();
                
                string baseDir = AppDomain.CurrentDomain.BaseDirectory;
                int loadedCount = 0;
                
                // Load base images from root directory
                string[] baseImages = { "diary.png", "ball.png", "check1.png", "check2.png" };
                foreach (string fileName in baseImages)
                {
                    string path = Path.Combine(baseDir, fileName);
                    if (File.Exists(path))
                    {
                        try
                        {
                            using (Bitmap bmp = new Bitmap(path))
                            using (Mat mat = BitmapConverter.ToMat(bmp))
                            {
                                // Convert to BGR format and cache
                                Mat cached = ConvertToBGR(mat, true);
                                templateCache[fileName] = cached;
                                loadedCount++;
                                Logger.Log(string.Format("[TemplateCache] 已載入主目錄圖片: {0} ({1}x{2})", 
                                    fileName, cached.Width, cached.Height));
                            }
                        }
                        catch (Exception ex)
                        {
                            Logger.Log(string.Format("[TemplateCache] 載入失敗: {0} - {1}", fileName, ex.Message));
                        }
                    }
                }
                
                // Load equipment images from pic directory
                string picDir = Path.Combine(baseDir, "pic");
                if (Directory.Exists(picDir))
                {
                    foreach (string file in Directory.GetFiles(picDir, "*.png"))
                    {
                        string fileName = Path.GetFileName(file);
                        try
                        {
                            using (Bitmap bmp = new Bitmap(file))
                            using (Mat mat = BitmapConverter.ToMat(bmp))
                            {
                                // Convert to BGR format and cache
                                Mat cached = ConvertToBGR(mat, true);
                                templateCache[fileName] = cached;
                                loadedCount++;
                                Logger.Log(string.Format("[TemplateCache] 已載入 pic 目錄圖片: {0} ({1}x{2})", 
                                    fileName, cached.Width, cached.Height));
                            }
                        }
                        catch (Exception ex)
                        {
                            Logger.Log(string.Format("[TemplateCache] 載入失敗: {0} - {1}", fileName, ex.Message));
                        }
                    }
                }
                
                Logger.Log(string.Format("[TemplateCache] 快取載入完成，共 {0} 張圖片", loadedCount));
            }
        }
        
        /// <summary>
        /// Convert Mat to BGR format (3 channels, 8-bit)
        /// If returnClone is true and input is already BGR, returns a clone.
        /// If false and input is already BGR, returns the input directly (caller must manage lifetime)
        /// </summary>
        public static Mat ConvertToBGR(Mat input, bool returnClone = false)
        {
            if (input.Channels() == 3)
            {
                // Already BGR - avoid unnecessary clone unless requested
                return returnClone ? input.Clone() : input;
            }
            else if (input.Channels() == 4)
            {
                // BGRA to BGR
                Mat output = new Mat();
                Cv2.CvtColor(input, output, ColorConversionCodes.BGRA2BGR);
                return output;
            }
            else if (input.Channels() == 1)
            {
                // Grayscale to BGR
                Mat output = new Mat();
                Cv2.CvtColor(input, output, ColorConversionCodes.GRAY2BGR);
                return output;
            }
            return returnClone ? input.Clone() : input;
        }
        /// <summary>
        /// Find image template in source bitmap (legacy method - converts to Mat internally)
        /// For better performance, use the Mat overload directly
        /// </summary>
        public static MatchResult FindImage(Bitmap source, string templateName, double threshold, Rectangle searchRegion = default(Rectangle))
        {
            if (source == null)
            {
                Logger.Log(string.Format("[FindImage] 錯誤: 來源圖片為 null - {0}", templateName));
                return new MatchResult { Success = false };
            }
            
            // Convert to Mat and use the optimized Mat version
            using (Mat sourceMat = BitmapConverter.ToMat(source))
            {
                return FindImage(sourceMat, templateName, threshold, searchRegion);
            }
        }
        
        /// <summary>
        /// Find image template in source Mat (optimized version - avoids redundant conversions)
        /// </summary>
        public static MatchResult FindImage(Mat sourceMat, string templateName, double threshold, Rectangle searchRegion = default(Rectangle))
        {
            if (sourceMat == null || sourceMat.Empty())
            {
                Logger.Log(string.Format("[FindImage] 錯誤: 來源圖片為 null 或空 - {0}", templateName));
                return new MatchResult { Success = false };
            }
            
            // If searchRegion is specified, crop the source to that region
            Mat processedSource = sourceMat;
            bool needsSourceDisposal = false;
            int offsetX = 0;
            int offsetY = 0;
            
            if (searchRegion != default(Rectangle) && searchRegion.Width > 0 && searchRegion.Height > 0)
            {
                // Validate region bounds
                if (searchRegion.X < 0 || searchRegion.Y < 0 || 
                    searchRegion.X + searchRegion.Width > sourceMat.Width ||
                    searchRegion.Y + searchRegion.Height > sourceMat.Height)
                {
                    Logger.Log(string.Format("[FindImage] 警告: 搜尋區域超出圖片範圍，使用完整圖片 - {0}", templateName));
                }
                else
                {
                    // Crop to region using Mat ROI (no memory copy)
                    offsetX = searchRegion.X;
                    offsetY = searchRegion.Y;
                    OpenCvSharp.Rect roi = new OpenCvSharp.Rect(searchRegion.X, searchRegion.Y, searchRegion.Width, searchRegion.Height);
                    processedSource = new Mat(sourceMat, roi);
                    needsSourceDisposal = true;
                    Logger.Log(string.Format("[FindImage] 使用搜尋區域: ({0},{1})-({2},{3}) - {4}",
                        searchRegion.X, searchRegion.Y, 
                        searchRegion.X + searchRegion.Width, searchRegion.Y + searchRegion.Height,
                        templateName));
                }
            }

            // Get cached template or load it
            Mat templateMat = null;
            lock (cacheLock)
            {
                if (templateCache.ContainsKey(templateName))
                {
                    templateMat = templateCache[templateName];
                }
            }
            
            // If not in cache, load it (fallback for backward compatibility)
            bool needsTemplateDisposal = false;
            if (templateMat == null)
            {
                Logger.Log(string.Format("[FindImage] 快取未命中，動態載入: {0}", templateName));
                
                string templatePath;
                string[] baseImages = { "diary.png", "ball.png", "check1.png", "check2.png" };
                bool isBaseImage = false;
                foreach (string baseName in baseImages)
                {
                    if (templateName.Equals(baseName, StringComparison.OrdinalIgnoreCase))
                    {
                        isBaseImage = true;
                        break;
                    }
                }
                
                if (isBaseImage)
                {
                    templatePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, templateName);
                }
                else
                {
                    templatePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "pic", templateName);
                }
                
                if (!File.Exists(templatePath))
                {
                    Logger.Log(string.Format("[FindImage] 錯誤: 找不到模板圖片 - {0}", templatePath));
                    return new MatchResult { Success = false };
                }
                
                using (Bitmap bmp = new Bitmap(templatePath))
                using (Mat mat = BitmapConverter.ToMat(bmp))
                {
                    templateMat = ConvertToBGR(mat, true); // Force clone for safety
                    needsTemplateDisposal = true;
                }
            }
            else
            {
                Logger.Log(string.Format("[FindImage] 使用快取圖片: {0}", templateName));
            }

            // Perform template matching
            using (Mat result = new Mat())
            {
                // Convert source to BGR if needed (avoid clone if already BGR)
                Mat sourceProcessed = ConvertToBGR(processedSource, false);
                bool needsSourceProcessedDisposal = (sourceProcessed != processedSource);
                
                try
                {
                    // Perform template matching using Normalized Cross-Correlation
                    // TM_CCOEFF_NORMED matches Python's cv2.TM_CCOEFF_NORMED
                    Cv2.MatchTemplate(sourceProcessed, templateMat, result, TemplateMatchModes.CCoeffNormed);
                    
                    // Find the best match location and score
                    double minVal, maxVal;
                    OpenCvSharp.Point minLoc, maxLoc;
                    Cv2.MinMaxLoc(result, out minVal, out maxVal, out minLoc, out maxLoc);
                    
                    bool success = maxVal >= threshold;
                    // Add offset if we used a cropped region
                    int finalX = maxLoc.X + offsetX;
                    int finalY = maxLoc.Y + offsetY;
                    
                    Logger.Log(string.Format("[FindImage] {0} - 分數: {1:F6}, 閾值: {2:F2}, 位置: ({3}, {4}){5}, {6}",
                        templateName, maxVal, threshold, finalX, finalY,
                        (offsetX != 0 || offsetY != 0) ? string.Format(" [區域內: ({0},{1})]", maxLoc.X, maxLoc.Y) : "",
                        success ? "✓ 成功" : "✗ 失敗"));
                    
                    return new MatchResult
                    {
                        Success = success,
                        Location = new System.Drawing.Point(finalX, finalY),
                        Score = maxVal
                    };
                }
                finally
                {
                    // Clean up converted source (only if conversion created a new Mat)
                    if (needsSourceProcessedDisposal && sourceProcessed != null)
                        sourceProcessed.Dispose();
                    
                    // Clean up cropped source ROI
                    if (needsSourceDisposal && processedSource != null)
                        processedSource.Dispose();
                    
                    // Clean up template if it wasn't from cache
                    if (needsTemplateDisposal && templateMat != null)
                        templateMat.Dispose();
                }
            }
        }
    }

    public static class Logger
    {
        private static object logLock = new object();
        private static string logPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "operation_log.txt");
        
        // Static flag to control logging based on debug mode
        public static bool IsDebugMode { get; set; }

        /// <summary>
        /// Clear the log file at the start of a new run
        /// </summary>
        public static void ClearLog()
        {
            lock (logLock)
            {
                try
                {
                    if (File.Exists(logPath))
                    {
                        File.Delete(logPath);
                    }
                }
                catch
                {
                    // Ignore errors when clearing log
                }
            }
        }

        public static void Log(string message)
        {
            // Only log if debug mode is enabled
            if (!IsDebugMode) return;
            
            lock (logLock)
            {
                try
                {
                    string timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff");
                    string logLine = string.Format("[{0}] {1}\r\n", timestamp, message);
                    File.AppendAllText(logPath, logLine);
                }
                catch { }
            }
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
            StringBuilder temp = new StringBuilder(255);
            string path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "config.ini");
            int i = GetPrivateProfileString(section, key, def, temp, 255, path);
            return temp.ToString();
        }

        public static void WriteValue(string section, string key, string value)
        {
            string path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "config.ini");
            WritePrivateProfileString(section, key, value, path);
        }
    }
}
