using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Windows.Forms;
using System.Drawing;
using System.Xml.Serialization;
using System.Linq;

namespace MusicBeePlugin
{
    public partial class Plugin
    {
        private MusicBeeApiInterface mbApiInterface;
        private PluginInfo about = new PluginInfo();
        private Settings config = new Settings();
        private System.Timers.Timer backupTimer;
        private string settingsFilePath;

        public PluginInfo Initialise(IntPtr apiInterfacePtr)
        {
            mbApiInterface = new MusicBeeApiInterface();
            mbApiInterface.Initialise(apiInterfacePtr);
            about.PluginInfoVersion = PluginInfoVersion;
            about.Name = "播放列表自动备份";
            about.Description = "现代化 UI，支持静态/动态列表分离与相对路径管理。";
            about.Author = "AI Assistant";
            about.TargetApplication = "";
            about.Type = PluginType.General;
            about.VersionMajor = 3;
            about.VersionMinor = 3; // Final UI Polish
            about.Revision = 1;
            about.MinInterfaceVersion = MinInterfaceVersion;
            about.MinApiRevision = MinApiRevision;
            about.ReceiveNotifications = (ReceiveNotificationFlags.PlayerEvents | ReceiveNotificationFlags.TagEvents);
            about.ConfigurationPanelHeight = 0;

            string dataPath = mbApiInterface.Setting_GetPersistentStoragePath();
            settingsFilePath = Path.Combine(dataPath, "mb_AutoBackup_V3.xml");

            LoadSettings();
            SetupTimer();

            return about;
        }

        public bool Configure(IntPtr panelHandle)
        {
            using (var form = new ConfigWindow(this, config, mbApiInterface))
            {
                if (form.ShowDialog() == DialogResult.OK)
                {
                    this.config = form.GetUpdatedConfig();
                    SaveSettings();
                    SetupTimer();
                }
            }
            return true;
        }

        public void SaveSettings()
        {
            try
            {
                XmlSerializer serializer = new XmlSerializer(typeof(Settings));
                using (StreamWriter writer = new StreamWriter(settingsFilePath))
                {
                    serializer.Serialize(writer, config);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("保存设置失败: " + ex.Message);
            }
        }

        private void LoadSettings()
        {
            if (File.Exists(settingsFilePath))
            {
                try
                {
                    XmlSerializer serializer = new XmlSerializer(typeof(Settings));
                    using (StreamReader reader = new StreamReader(settingsFilePath))
                    {
                        config = (Settings)serializer.Deserialize(reader);
                    }
                }
                catch { config = new Settings(); }
            }
        }

        private void SetupTimer()
        {
            if (backupTimer != null)
            {
                backupTimer.Stop();
                backupTimer.Dispose();
                backupTimer = null;
            }

            if (config.EnableIntervalBackup && config.IntervalMinutes > 0)
            {
                backupTimer = new System.Timers.Timer(config.IntervalMinutes * 60 * 1000);
                backupTimer.Elapsed += (s, e) => PerformBackup("Interval");
                backupTimer.AutoReset = true;
                backupTimer.Start();
            }
        }

        public void Close(PluginCloseReason reason)
        {
            if (config.BackupOnShutdown)
            {
                PerformBackup("Shutdown");
            }
            if (backupTimer != null) backupTimer.Dispose();
        }

        public void Uninstall()
        {
            if (File.Exists(settingsFilePath)) File.Delete(settingsFilePath);
        }

        public void ReceiveNotification(string sourceFileUrl, NotificationType type) { }

        // --- 核心逻辑 ---
        public void PerformBackup(string triggerSource)
        {
            if (config.Playlists == null || config.Playlists.Count == 0) return;

            Dictionary<string, string> validPlaylists = new Dictionary<string, string>();
            if (mbApiInterface.Playlist_QueryPlaylists())
            {
                string url;
                while ((url = mbApiInterface.Playlist_QueryGetNextPlaylist()) != null)
                {
                    string name = mbApiInterface.Playlist_GetName(url);
                    if (!validPlaylists.ContainsKey(name)) validPlaylists.Add(name, url);
                }
            }

            foreach (var plSetting in config.Playlists)
            {
                if (!plSetting.Enabled) continue;
                if (!validPlaylists.ContainsKey(plSetting.Name)) continue;

                string plUrl = validPlaylists[plSetting.Name];

                string exportDir = plSetting.CustomExportPath;
                if (string.IsNullOrWhiteSpace(exportDir)) exportDir = config.DefaultExportPath;
                exportDir = ResolvePath(exportDir);

                string baseDir = plSetting.CustomRootPath;
                bool useRelative = !string.IsNullOrWhiteSpace(baseDir);
                if (useRelative) baseDir = ResolvePath(baseDir);

                if (!Directory.Exists(exportDir))
                {
                    try { Directory.CreateDirectory(exportDir); } catch { continue; }
                }

                List<string> files = new List<string>();
                if (mbApiInterface.Playlist_QueryFiles(plUrl))
                {
                    string file;
                    while ((file = mbApiInterface.Playlist_QueryGetNextFile()) != null)
                    {
                        files.Add(file);
                    }
                }

                string simpleName = GetLeafName(plSetting.Name);
                string m3uPath = Path.Combine(exportDir, ReplaceInvalidChars(simpleName) + ".m3u8");

                try
                {
                    using (StreamWriter sw = new StreamWriter(m3uPath, false, Encoding.UTF8))
                    {
                        sw.WriteLine("#EXTM3U");
                        foreach (string audioFile in files)
                        {
                            string writePath = audioFile;
                            if (useRelative)
                            {
                                writePath = CalculateRelativePath(baseDir, audioFile);
                            }
                            sw.WriteLine(writePath);
                        }
                    }
                }
                catch { }
            }
        }

        private string GetLeafName(string fullPath)
        {
            if (string.IsNullOrEmpty(fullPath)) return "";
            int idx = fullPath.LastIndexOf('\\');
            if (idx >= 0 && idx < fullPath.Length - 1)
                return fullPath.Substring(idx + 1);
            return fullPath;
        }

        private string ResolvePath(string inputPath)
        {
            if (string.IsNullOrWhiteSpace(inputPath)) return "";
            if (inputPath.StartsWith(".\\") || inputPath.StartsWith("./"))
            {
                string appPath = System.AppDomain.CurrentDomain.BaseDirectory;
                return Path.Combine(appPath, inputPath.Substring(2));
            }
            if (!Path.IsPathRooted(inputPath))
            {
                return Path.Combine(System.AppDomain.CurrentDomain.BaseDirectory, inputPath);
            }
            return inputPath;
        }

        private string CalculateRelativePath(string fromPath, string toPath)
        {
            try
            {
                if (string.IsNullOrEmpty(fromPath) || string.IsNullOrEmpty(toPath)) return toPath;
                string from = Path.GetFullPath(fromPath).TrimEnd(Path.DirectorySeparatorChar);
                string to = Path.GetFullPath(toPath);

                string fromRoot = Path.GetPathRoot(from);
                string toRoot = Path.GetPathRoot(to);
                if (!string.Equals(fromRoot, toRoot, StringComparison.OrdinalIgnoreCase)) return toPath;

                string[] fromDirs = from.Split(Path.DirectorySeparatorChar);
                string[] toDirs = to.Split(Path.DirectorySeparatorChar);
                int length = Math.Min(fromDirs.Length, toDirs.Length);
                int lastCommonRoot = -1;

                for (int i = 0; i < length; i++)
                {
                    if (string.Equals(fromDirs[i], toDirs[i], StringComparison.OrdinalIgnoreCase)) lastCommonRoot = i;
                    else break;
                }
                if (lastCommonRoot == -1) return toPath;

                StringBuilder sb = new StringBuilder();
                for (int i = lastCommonRoot + 1; i < fromDirs.Length; i++) sb.Append("..\\");
                for (int i = lastCommonRoot + 1; i < toDirs.Length; i++)
                {
                    sb.Append(toDirs[i]);
                    if (i < toDirs.Length - 1) sb.Append("\\");
                }
                string result = sb.ToString();
                if (!result.StartsWith("..")) return ".\\" + result;
                return result;
            }
            catch { return toPath; }
        }

        private string ReplaceInvalidChars(string filename)
        {
            return string.Join("_", filename.Split(Path.GetInvalidFileNameChars()));
        }

        public class Settings
        {
            public string Language { get; set; } = "CN";
            public bool UseSkinTheme { get; set; } = false;
            public bool BackupOnShutdown { get; set; } = false;
            public bool EnableIntervalBackup { get; set; } = false;
            public int IntervalMinutes { get; set; } = 1440;
            public string DefaultExportPath { get; set; } = ".\\PlaylistsBackup";
            public List<PlaylistSetting> Playlists { get; set; } = new List<PlaylistSetting>();
        }

        public class PlaylistSetting
        {
            public string Name { get; set; }
            public bool Enabled { get; set; } = false;
            public string CustomExportPath { get; set; } = "";
            public string CustomRootPath { get; set; } = "";
        }

        // --- UI ---
        public class ConfigWindow : Form
        {
            private Settings _config;
            private Plugin _plugin;
            private MusicBeeApiInterface _api;

            // 控件 - 将所有带文字的控件提升为字段以便翻译
            private DataGridView gridStatic, gridAuto;
            private Panel panelAutoContainer;
            private Button btnToggleAuto;
            private CheckBox chkShutdown, chkInterval, chkTheme;
            private NumericUpDown numDay, numHour, numMin;
            private TextBox txtDefaultPath;
            private Button btnSave, btnCancel;
            private ComboBox comboLang;

            // 需要动态翻译的容器和标签
            private GroupBox grpSettings, grpStatic;
            private Label lblStrategy, lblDefPath;
            private Label lblD, lblH, lblM;

            private Color clrBg, clrFg, clrPanel, clrBorder, clrAccent;

            public ConfigWindow(Plugin plugin, Settings config, MusicBeeApiInterface api)
            {
                _plugin = plugin;
                _config = config;
                _api = api;

                clrBg = Color.FromArgb(245, 245, 245);
                clrFg = Color.FromArgb(30, 30, 30);
                clrPanel = Color.White;
                clrBorder = Color.FromArgb(210, 210, 210);
                clrAccent = Color.FromArgb(0, 120, 215);

                if (_config.UseSkinTheme) GetSkinColors();

                InitializeComponent();
                LoadData();
                ApplyTheme();
            }

            private void GetSkinColors()
            {
                try
                {
                    int bg = _api.Setting_GetSkinElementColour(SkinElement.SkinSubPanel, ElementState.ElementStateDefault, ElementComponent.ComponentBackground);
                    int fg = _api.Setting_GetSkinElementColour(SkinElement.SkinSubPanel, ElementState.ElementStateDefault, ElementComponent.ComponentForeground);
                    if (bg != 0)
                    {
                        clrBg = Color.FromArgb(bg);
                        clrPanel = ControlPaint.Light(clrBg);
                    }
                    if (fg != 0) clrFg = Color.FromArgb(fg);
                    clrBorder = ControlPaint.Dark(clrBg);
                    clrAccent = clrFg;
                }
                catch { }
            }

            private void InitializeComponent()
            {
                this.Size = new Size(950, 750);
                this.StartPosition = FormStartPosition.CenterParent;
                this.FormBorderStyle = FormBorderStyle.Sizable;
                this.MinimumSize = new Size(850, 600);
                this.Font = new Font("Segoe UI", 9F, FontStyle.Regular, GraphicsUnit.Point, 0);

                TableLayoutPanel layout = new TableLayoutPanel { Dock = DockStyle.Fill, Padding = new Padding(20) };
                layout.RowCount = 4;
                layout.RowStyles.Add(new RowStyle(SizeType.AutoSize)); // Top Options
                layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100f)); // Static Grid
                layout.RowStyles.Add(new RowStyle(SizeType.AutoSize)); // Auto Grid Container (AutoSize)
                layout.RowStyles.Add(new RowStyle(SizeType.AutoSize)); // Footer
                this.Controls.Add(layout);

                // 1. 顶部设置区
                grpSettings = new GroupBox { Text = "设置", Dock = DockStyle.Top, AutoSize = true, Padding = new Padding(10) };
                TableLayoutPanel panelTop = new TableLayoutPanel { Dock = DockStyle.Top, AutoSize = true, ColumnCount = 1 };

                // Line 1: 备份策略
                FlowLayoutPanel flowStrategy = new FlowLayoutPanel { AutoSize = true, WrapContents = true, FlowDirection = FlowDirection.LeftToRight };
                lblStrategy = new Label { Text = "备份策略:", AutoSize = true, Margin = new Padding(0, 8, 10, 0), Font = new Font(this.Font, FontStyle.Bold) };

                chkShutdown = new CheckBox { Text = "软件关闭时", AutoSize = true, Margin = new Padding(0, 5, 20, 0) };
                chkInterval = new CheckBox { Text = "定时备份:", AutoSize = true, Margin = new Padding(0, 5, 5, 0) };
                numDay = CreateNum(0, 365); lblD = new Label { Text = "天", AutoSize = true, Margin = new Padding(0, 8, 5, 0) };
                numHour = CreateNum(0, 23); lblH = new Label { Text = "时", AutoSize = true, Margin = new Padding(0, 8, 5, 0) };
                numMin = CreateNum(0, 59); lblM = new Label { Text = "分", AutoSize = true, Margin = new Padding(0, 8, 20, 0) };

                flowStrategy.Controls.AddRange(new Control[] { lblStrategy, chkShutdown, chkInterval, numDay, lblD, numHour, lblH, numMin, lblM });

                // Line 2: 通用设置
                FlowLayoutPanel flowGeneral = new FlowLayoutPanel { AutoSize = true, WrapContents = true, FlowDirection = FlowDirection.LeftToRight, Margin = new Padding(0, 10, 0, 0) };

                comboLang = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = 80, Margin = new Padding(0, 5, 20, 0) };
                comboLang.Items.AddRange(new object[] { "EN", "CN" });
                comboLang.SelectedIndexChanged += (s, e) => ApplyLanguage();

                chkTheme = new CheckBox { Text = "跟随主题色", AutoSize = true, Margin = new Padding(0, 5, 20, 0) };
                chkTheme.CheckedChanged += (s, e) => { _config.UseSkinTheme = chkTheme.Checked; ApplyTheme(); };

                lblDefPath = new Label { Text = "默认导出位置:", AutoSize = true, Margin = new Padding(0, 8, 5, 0) };
                txtDefaultPath = new TextBox { Width = 250, Margin = new Padding(0, 5, 5, 0) };
                Button btnDefBrowse = new Button { Text = "...", Width = 30, Height = 25, Margin = new Padding(0, 4, 0, 0) };
                btnDefBrowse.Click += (s, e) => {
                    FolderBrowserDialog fbd = new FolderBrowserDialog();
                    if (fbd.ShowDialog() == DialogResult.OK) txtDefaultPath.Text = fbd.SelectedPath;
                };

                flowGeneral.Controls.AddRange(new Control[] { comboLang, chkTheme, lblDefPath, txtDefaultPath, btnDefBrowse });

                panelTop.Controls.Add(flowStrategy);
                panelTop.Controls.Add(flowGeneral);
                grpSettings.Controls.Add(panelTop);
                layout.Controls.Add(grpSettings, 0, 0);

                // 2. 静态列表
                grpStatic = new GroupBox { Text = "静态播放列表", Dock = DockStyle.Fill };
                gridStatic = CreateModernGrid();

                FlowLayoutPanel panelBatchStatic = CreateBatchPanel(gridStatic);
                grpStatic.Controls.Add(panelBatchStatic);
                grpStatic.Controls.Add(gridStatic);

                layout.Controls.Add(grpStatic, 0, 1);

                // 3. 动态列表 (可折叠)
                panelAutoContainer = new Panel { Dock = DockStyle.Top, AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink };

                btnToggleAuto = new Button { Text = "▼ 动态播放列表", Dock = DockStyle.Top, Height = 30, FlatStyle = FlatStyle.Flat, TextAlign = ContentAlignment.MiddleLeft };
                btnToggleAuto.FlatAppearance.BorderSize = 0;
                btnToggleAuto.Click += (s, e) => ToggleAutoGrid();

                // 修复点：关闭 AutoSize，并将高度减小为 160
                GroupBox grpAuto = new GroupBox { Dock = DockStyle.Top, Text = "", Visible = false, Height = 160 };
                grpAuto.AutoSize = false;

                gridAuto = CreateModernGrid();
                FlowLayoutPanel panelBatchAuto = CreateBatchPanel(gridAuto);

                grpAuto.Controls.Add(gridAuto);
                grpAuto.Controls.Add(panelBatchAuto);
                gridAuto.BringToFront();

                panelAutoContainer.Controls.Add(grpAuto);
                panelAutoContainer.Controls.Add(btnToggleAuto);
                layout.Controls.Add(panelAutoContainer, 0, 2);

                // 4. 底部按钮
                FlowLayoutPanel panelFooter = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.RightToLeft, Padding = new Padding(0, 15, 0, 0) };
                btnCancel = CreateMainButton("取消", false);
                btnSave = CreateMainButton("保存配置", true);

                btnCancel.Click += (s, e) => { this.DialogResult = DialogResult.Cancel; this.Close(); };
                btnSave.Click += BtnSave_Click;

                panelFooter.Controls.Add(btnCancel);
                panelFooter.Controls.Add(btnSave);
                layout.Controls.Add(panelFooter, 0, 3);
            }

            private FlowLayoutPanel CreateBatchPanel(DataGridView dgv)
            {
                FlowLayoutPanel p = new FlowLayoutPanel { Dock = DockStyle.Bottom, AutoSize = true, Height = 40, FlowDirection = FlowDirection.LeftToRight };

                Button btnAppPath = CreateFlatButton("应用到同类");
                Button btnClrPath = CreateFlatButton("清除选中");
                Label sep = new Label { Text = "|", AutoSize = true, Margin = new Padding(5, 10, 5, 0), ForeColor = Color.Gray };

                Button btnAppRoot = CreateFlatButton("应用到同类");
                Button btnClrRoot = CreateFlatButton("清除选中");

                btnAppPath.Click += (s, e) => BatchApply(dgv, 2);
                btnClrPath.Click += (s, e) => BatchClear(dgv, 2);

                btnAppRoot.Click += (s, e) => BatchApply(dgv, 4);
                btnClrRoot.Click += (s, e) => BatchClear(dgv, 4);

                p.Controls.AddRange(new Control[] {
                    new Label{Text="导出路径:", AutoSize=true, Margin=new Padding(0,10,0,0)}, btnAppPath, btnClrPath,
                    sep,
                    new Label{Text="前缀:", AutoSize=true, Margin=new Padding(0,10,0,0)}, btnAppRoot, btnClrRoot
                });
                return p;
            }

            private void ToggleAutoGrid()
            {
                GroupBox g = panelAutoContainer.Controls.OfType<GroupBox>().FirstOrDefault();
                if (g == null) return;

                bool isVisible = !g.Visible;
                g.Visible = isVisible;

                string suffix = (comboLang.SelectedItem.ToString() == "CN" ? " 动态播放列表" : " Auto Playlists");
                btnToggleAuto.Text = (isVisible ? "▲" : "▼") + suffix;
            }

            private NumericUpDown CreateNum(int min, int max)
            {
                return new NumericUpDown { Minimum = min, Maximum = max, Width = 45, Margin = new Padding(0, 5, 0, 0) };
            }

            private Button CreateFlatButton(string text)
            {
                Button b = new Button { Text = text, AutoSize = true, Height = 26, FlatStyle = FlatStyle.Flat, Margin = new Padding(3, 5, 3, 3) };
                b.FlatAppearance.BorderColor = Color.LightGray;
                b.Font = new Font(this.Font.FontFamily, 8f);
                return b;
            }

            private Button CreateMainButton(string text, bool primary)
            {
                Button b = new Button { Text = text, Width = 100, Height = 35, FlatStyle = FlatStyle.Flat, Cursor = Cursors.Hand };
                b.FlatAppearance.BorderSize = 0;
                return b;
            }

            private DataGridView CreateModernGrid()
            {
                var dgv = new DataGridView();
                dgv.Dock = DockStyle.Fill;
                dgv.AutoGenerateColumns = false;
                dgv.AllowUserToAddRows = false;
                dgv.AllowUserToDeleteRows = false;
                dgv.AllowUserToResizeRows = false;
                dgv.RowHeadersVisible = false;
                dgv.SelectionMode = DataGridViewSelectionMode.CellSelect;
                dgv.BackgroundColor = Color.White;
                dgv.BorderStyle = BorderStyle.FixedSingle;
                dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
                dgv.RowTemplate.Height = 32;

                dgv.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing;
                dgv.ColumnHeadersHeight = 35;

                // 列
                var colCheck = new DataGridViewCheckBoxColumn { DataPropertyName = "Enabled", Width = 40, HeaderText = "✔" };
                colCheck.AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
                colCheck.MinimumWidth = 40;

                var colName = new DataGridViewTextBoxColumn { DataPropertyName = "DisplayName", ReadOnly = true, HeaderText = "列表名称" };
                colName.MinimumWidth = 100;

                var colPath = new DataGridViewTextBoxColumn { DataPropertyName = "Path", HeaderText = "导出路径" };
                colPath.MinimumWidth = 100;

                var colBtnPath = new DataGridViewButtonColumn { Text = "...", UseColumnTextForButtonValue = true, Width = 30 };
                colBtnPath.AutoSizeMode = DataGridViewAutoSizeColumnMode.None;

                var colRoot = new DataGridViewTextBoxColumn { DataPropertyName = "RootPath", HeaderText = "相对路径前缀" };
                colRoot.MinimumWidth = 100;

                var colBtnRoot = new DataGridViewButtonColumn { Text = "...", UseColumnTextForButtonValue = true, Width = 30 };
                colBtnRoot.AutoSizeMode = DataGridViewAutoSizeColumnMode.None;

                dgv.Columns.AddRange(colCheck, colName, colPath, colBtnPath, colRoot, colBtnRoot);
                dgv.CellContentClick += Grid_CellContentClick;
                return dgv;
            }

            private void Grid_CellContentClick(object sender, DataGridViewCellEventArgs e)
            {
                var dgv = sender as DataGridView;
                if (e.RowIndex < 0 || dgv == null) return;

                if (e.ColumnIndex == 3)
                {
                    FolderBrowserDialog fbd = new FolderBrowserDialog();
                    if (fbd.ShowDialog() == DialogResult.OK) dgv.Rows[e.RowIndex].Cells[2].Value = fbd.SelectedPath;
                }
                else if (e.ColumnIndex == 5)
                {
                    FolderBrowserDialog fbd = new FolderBrowserDialog();
                    if (fbd.ShowDialog() == DialogResult.OK) dgv.Rows[e.RowIndex].Cells[4].Value = fbd.SelectedPath;
                }
            }

            // 批量应用
            private void BatchApply(DataGridView dgv, int valueColIndex)
            {
                HashSet<string> values = new HashSet<string>();
                List<int> checkedRowIndices = new List<int>();

                foreach (DataGridViewRow row in dgv.Rows)
                {
                    bool enabled = Convert.ToBoolean(row.Cells[0].Value);
                    if (enabled)
                    {
                        checkedRowIndices.Add(row.Index);
                        string val = row.Cells[valueColIndex].Value?.ToString();
                        if (!string.IsNullOrWhiteSpace(val)) values.Add(val);
                    }
                }

                if (values.Count > 1)
                {
                    MessageBox.Show("所选项目值不一致，无法自动应用。", "冲突", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                if (values.Count == 0) return;

                string targetValue = values.First();
                foreach (int idx in checkedRowIndices) dgv.Rows[idx].Cells[valueColIndex].Value = targetValue;
            }

            private void BatchClear(DataGridView dgv, int valueColIndex)
            {
                foreach (DataGridViewRow row in dgv.Rows)
                {
                    bool enabled = Convert.ToBoolean(row.Cells[0].Value);
                    if (enabled) row.Cells[valueColIndex].Value = "";
                }
            }

            private void BtnSave_Click(object sender, EventArgs e)
            {
                CheckPathConflict(gridStatic);
                CheckPathConflict(gridAuto);
                this.DialogResult = DialogResult.OK;
                this.Close();
            }

            private void CheckPathConflict(DataGridView dgv)
            {
                foreach (DataGridViewRow row in dgv.Rows)
                {
                    bool enabled = Convert.ToBoolean(row.Cells[0].Value);
                    if (!enabled) continue;

                    string path = row.Cells[2].Value?.ToString() ?? "";
                    string root = row.Cells[4].Value?.ToString() ?? "";

                    if (!string.IsNullOrWhiteSpace(path) && !string.IsNullOrWhiteSpace(root))
                    {
                        try
                        {
                            string rootDrive = Path.GetPathRoot(Path.GetFullPath(root));
                            string pathDrive = Path.GetPathRoot(Path.GetFullPath(path));
                            if (rootDrive != pathDrive)
                            {
                                MessageBox.Show(
                                    $"列表 '{row.Cells[1].Value}' 的导出路径与前缀不在同盘符，将失效。",
                                    "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                return;
                            }
                        }
                        catch { }
                    }
                }
            }

            private class GridItem
            {
                public bool Enabled { get; set; }
                public string Name { get; set; }
                public string DisplayName { get; set; }
                public string Path { get; set; }
                public string RootPath { get; set; }
            }

            private void LoadData()
            {
                comboLang.SelectedItem = _config.Language;
                chkShutdown.Checked = _config.BackupOnShutdown;
                chkInterval.Checked = _config.EnableIntervalBackup;
                chkTheme.Checked = _config.UseSkinTheme;
                txtDefaultPath.Text = _config.DefaultExportPath;

                int totalMin = _config.IntervalMinutes;
                numDay.Value = totalMin / (24 * 60);
                numHour.Value = (totalMin % (24 * 60)) / 60;
                numMin.Value = totalMin % 60;

                var listStatic = new List<GridItem>();
                var listAuto = new List<GridItem>();

                if (_api.Playlist_QueryPlaylists())
                {
                    string url;
                    while ((url = _api.Playlist_QueryGetNextPlaylist()) != null)
                    {
                        string fullName = _api.Playlist_GetName(url);
                        PlaylistFormat fmt = _api.Playlist_GetType(url);

                        string simpleName = fullName;
                        int idx = fullName.LastIndexOf('\\');
                        if (idx >= 0 && idx < fullName.Length - 1) simpleName = fullName.Substring(idx + 1);

                        var saved = _config.Playlists.FirstOrDefault(p => p.Name == fullName);

                        var item = new GridItem
                        {
                            Enabled = saved != null ? saved.Enabled : false,
                            Name = fullName,
                            DisplayName = simpleName,
                            Path = saved != null ? saved.CustomExportPath : "",
                            RootPath = saved != null ? saved.CustomRootPath : ""
                        };

                        if (fmt == PlaylistFormat.Auto) listAuto.Add(item);
                        else listStatic.Add(item);
                    }
                }

                gridStatic.DataSource = listStatic;
                gridAuto.DataSource = listAuto;
                ApplyLanguage();
            }

            public Settings GetUpdatedConfig()
            {
                _config.Language = comboLang.SelectedItem.ToString();
                _config.BackupOnShutdown = chkShutdown.Checked;
                _config.EnableIntervalBackup = chkInterval.Checked;
                _config.IntervalMinutes = (int)((numDay.Value * 24 * 60) + (numHour.Value * 60) + numMin.Value);
                _config.DefaultExportPath = txtDefaultPath.Text;
                _config.UseSkinTheme = chkTheme.Checked;

                _config.Playlists.Clear();
                SaveGrid(gridStatic);
                SaveGrid(gridAuto);
                return _config;
            }

            private void SaveGrid(DataGridView dgv)
            {
                var items = dgv.DataSource as List<GridItem>;
                if (items == null) return;
                foreach (var item in items)
                {
                    _config.Playlists.Add(new PlaylistSetting
                    {
                        Name = item.Name,
                        Enabled = item.Enabled,
                        CustomExportPath = item.Path,
                        CustomRootPath = item.RootPath
                    });
                }
            }

            private void ApplyTheme()
            {
                if (_config.UseSkinTheme) GetSkinColors();
                else
                {
                    clrBg = Color.WhiteSmoke;
                    clrFg = Color.Black;
                    clrPanel = Color.White;
                    clrBorder = Color.LightGray;
                    clrAccent = Color.DodgerBlue;
                }

                this.BackColor = clrBg;
                this.ForeColor = clrFg;
                UpdateColorRec(this);
                StyleGrid(gridStatic);
                StyleGrid(gridAuto);

                btnSave.BackColor = clrAccent;
                btnSave.ForeColor = Color.White;
                btnCancel.BackColor = clrPanel;
                btnToggleAuto.BackColor = ControlPaint.Light(clrBg);
            }

            private void UpdateColorRec(Control p)
            {
                foreach (Control c in p.Controls)
                {
                    if (c is GroupBox) c.ForeColor = clrFg;
                    if (c is CheckBox || c is Label) c.ForeColor = clrFg;
                    if (c is TextBox || c is NumericUpDown || c is ComboBox)
                    {
                        c.BackColor = clrPanel;
                        c.ForeColor = clrFg;
                    }
                    if (c is Button && c != btnSave && c != btnToggleAuto)
                    {
                        c.BackColor = clrPanel;
                        c.ForeColor = clrFg;
                    }
                    if (c.HasChildren) UpdateColorRec(c);
                }
            }

            private void StyleGrid(DataGridView g)
            {
                g.BackgroundColor = clrPanel;
                g.DefaultCellStyle.BackColor = clrPanel;
                g.DefaultCellStyle.ForeColor = clrFg;
                g.DefaultCellStyle.SelectionBackColor = clrAccent;
                g.ColumnHeadersDefaultCellStyle.BackColor = clrBg;
                g.ColumnHeadersDefaultCellStyle.ForeColor = clrFg;
                g.EnableHeadersVisualStyles = false;
                g.GridColor = clrBorder;
            }

            private void ApplyLanguage()
            {
                bool isCn = comboLang.SelectedItem.ToString() == "CN";

                this.Text = isCn ? "播放列表自动备份" : "Playlist Auto Backup";

                string hName = isCn ? "列表名称" : "Name";
                string hPath = isCn ? "导出路径" : "Export Path";
                string hRoot = isCn ? "相对路径前缀" : "Rel. Prefix";

                gridStatic.Columns[1].HeaderText = hName;
                gridStatic.Columns[2].HeaderText = hPath;
                gridStatic.Columns[4].HeaderText = hRoot;

                gridAuto.Columns[1].HeaderText = hName;
                gridAuto.Columns[2].HeaderText = hPath;
                gridAuto.Columns[4].HeaderText = hRoot;

                // Fix: 动态翻译所有控件
                grpSettings.Text = isCn ? "设置" : "Settings";
                lblStrategy.Text = isCn ? "备份策略:" : "Backup Strategy:";
                grpStatic.Text = isCn ? "静态播放列表" : "Static Playlists";
                lblDefPath.Text = isCn ? "默认导出位置:" : "Default Export Path:";
                lblD.Text = isCn ? "天" : "d";
                lblH.Text = isCn ? "时" : "h";
                lblM.Text = isCn ? "分" : "m";

                chkShutdown.Text = isCn ? "软件关闭时" : "On Shutdown";
                chkInterval.Text = isCn ? "定时备份:" : "Interval:";
                chkTheme.Text = isCn ? "跟随主题色" : "Follow Theme";

                string suffix = isCn ? " 动态播放列表" : " Auto Playlists";
                if (gridAuto != null)
                {
                    btnToggleAuto.Text = (gridAuto.Visible ? "▲" : "▼") + suffix;
                }

                btnSave.Text = isCn ? "保存配置" : "Save";
                btnCancel.Text = isCn ? "取消" : "Cancel";

                UpdateBatchButtonText(this, isCn);
            }

            private void UpdateBatchButtonText(Control parent, bool isCn)
            {
                foreach (Control c in parent.Controls)
                {
                    if (c is FlowLayoutPanel)
                    {
                        foreach (Control btn in c.Controls)
                        {
                            if (btn is Button)
                            {
                                if (btn.Text.Contains("应用") || btn.Text.Contains("Apply"))
                                    btn.Text = isCn ? "应用到同类" : "Apply to All";
                                if (btn.Text.Contains("清除") || btn.Text.Contains("Clear"))
                                    btn.Text = isCn ? "清除选中" : "Clear Sel.";
                            }
                            if (btn is Label)
                            {
                                if (btn.Text.Contains("导出") || btn.Text.Contains("Export"))
                                    btn.Text = isCn ? "导出路径:" : "Export Path:";
                                if (btn.Text.Contains("前缀") || btn.Text.Contains("Prefix"))
                                    btn.Text = isCn ? "前缀:" : "Prefix:";
                            }
                        }
                    }
                    if (c.HasChildren) UpdateBatchButtonText(c, isCn);
                }
            }
        }
    }
}