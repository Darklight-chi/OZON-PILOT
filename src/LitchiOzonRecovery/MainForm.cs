using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.IO;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;
using Newtonsoft.Json.Linq;

namespace LitchiOzonRecovery
{
    public sealed class MainForm : Form
    {
        private const string DefaultOzonClientId = "3401155";
        private const string DefaultOzonApiKey = "639a0e1c-4299-4fbb-a59d-e972df07402e";
        private static readonly Color ShellBack = Color.FromArgb(247, 245, 237);
        private static readonly Color CardBack = Color.FromArgb(242, 255, 252, 246);
        private static readonly Color SoftCardBack = Color.FromArgb(236, 255, 255, 250);
        private static readonly Color LineWarm = Color.FromArgb(224, 214, 196);
        private static readonly Color TextStrong = Color.FromArgb(22, 33, 48);
        private static readonly Color TextMuted = Color.FromArgb(112, 127, 145);
        private static readonly Color PilotGreen = Color.FromArgb(14, 139, 100);
        private static readonly Color PilotGreenDark = Color.FromArgb(9, 113, 91);
        private static readonly Color PilotGreenSoft = Color.FromArgb(225, 247, 240);

        private sealed class FeeRuleDisplayRow
        {
            public FeeRule Rule { get; set; }
            public int Id { get; set; }
            public long CategoryId1 { get; set; }
            public long CategoryId2 { get; set; }
            public string Category1 { get; set; }
            public string Category2 { get; set; }
            public decimal FBS { get; set; }
            public decimal FBS1500 { get; set; }
            public decimal FBS5000 { get; set; }
            public decimal FBP { get; set; }
            public decimal FBP1500 { get; set; }
            public decimal FBP5000 { get; set; }
            public decimal FBO { get; set; }
            public decimal FBO1500 { get; set; }
            public decimal FBO5000 { get; set; }
        }

        private sealed class RoundedPanel : Panel
        {
            public int Radius { get; set; }
            public Color BorderColor { get; set; }
            public Color FillColor { get; set; }
            public Color ShadowColor { get; set; }
            public bool DrawShadow { get; set; }

            public RoundedPanel()
            {
                Radius = 22;
                BorderColor = Color.FromArgb(210, 224, 214, 198);
                FillColor = SoftCardBack;
                ShadowColor = Color.FromArgb(20, 87, 76, 55);
                DrawShadow = true;
                BackColor = Color.Transparent;
                SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer | ControlStyles.UserPaint | ControlStyles.ResizeRedraw, true);
            }

            protected override void OnPaint(PaintEventArgs e)
            {
                e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
                Rectangle body = ClientRectangle;
                body.Width -= 1;
                body.Height -= 1;

                if (DrawShadow && body.Width > 12 && body.Height > 12)
                {
                    Rectangle shadow = body;
                    shadow.Inflate(-2, -2);
                    shadow.Offset(0, 5);
                    using (GraphicsPath shadowPath = CreateRoundPath(shadow, Radius))
                    using (SolidBrush shadowBrush = new SolidBrush(ShadowColor))
                    {
                        e.Graphics.FillPath(shadowBrush, shadowPath);
                    }
                }

                using (GraphicsPath path = CreateRoundPath(body, Radius))
                using (SolidBrush brush = new SolidBrush(FillColor))
                using (Pen pen = new Pen(BorderColor))
                {
                    e.Graphics.FillPath(brush, path);
                    e.Graphics.DrawPath(pen, path);
                }

                base.OnPaint(e);
            }
        }

        private sealed class GradientPanel : Panel
        {
            public Color StartColor { get; set; }
            public Color EndColor { get; set; }
            public float Angle { get; set; }

            public GradientPanel()
            {
                StartColor = Color.FromArgb(250, 248, 241);
                EndColor = Color.FromArgb(239, 245, 238);
                Angle = 90f;
                SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer | ControlStyles.UserPaint | ControlStyles.ResizeRedraw, true);
            }

            protected override void OnPaint(PaintEventArgs e)
            {
                using (LinearGradientBrush brush = new LinearGradientBrush(ClientRectangle, StartColor, EndColor, Angle))
                {
                    e.Graphics.FillRectangle(brush, ClientRectangle);
                }

                base.OnPaint(e);
            }
        }

        private readonly AppPaths _paths;
        private readonly ProductAutomationService _automationService;
        private readonly Random _random;
        private CancellationTokenSource _currentProcessCancel;
        private string _currentProcessName;
        private AssetSnapshot _snapshot;
        private SourcingResult _lastSourcingResult;
        private string _fullAutoReport;

        private TabControl _mainTabs;
        private TextBox _overviewBox;
        private Label _cardCategoryValue;
        private Label _cardFeeValue;
        private Label _cardPluginValue;
        private ComboBox _languageComboBox;
        private TextBox _autoLoopCountBox;
        private PropertyGrid _configGrid;
        private TreeView _categoryTree;
        private DataGridView _feeGrid;
        private TextBox _assetSearchBox;
        private TextBox _browserUrlBox;
        private Label _browserStatusLabel;
        private WebView2 _browser;
        private bool _browserExtensionReady;
        private TextBox _autoKeywordsBox;
        private ComboBox _autoProviderBox;
        private TextBox _autoApiKeyBox;
        private TextBox _autoApiSecretBox;
        private TextBox _autoPerKeywordBox;
        private TextBox _autoDetailLimitBox;
        private TextBox _autoRubRateBox;
        private TextBox _autoCategoryIdBox;
        private TextBox _autoTypeIdBox;
        private TextBox _autoPriceMultiplierBox;
        private TextBox _ozonClientIdBox;
        private TextBox _ozonApiKeyBox;
        private DataGridView _autoResultGrid;
        private TextBox _autoLogBox;
        private StatusStrip _statusStrip;
        private ToolStripStatusLabel _statusLabel;
        private string _assetErrorMessage;
        private string _uiLanguage = "zh";

        public MainForm()
        {
            _paths = AppPaths.Discover();
            _automationService = new ProductAutomationService();
            _random = new Random();
            LoadUiLanguagePreference();

            Text = "OZON-PILOT";
            Width = 1460;
            Height = 960;
            StartPosition = FormStartPosition.CenterScreen;
            MinimumSize = new Size(1280, 820);
            Font = new Font("Microsoft YaHei UI", 9F, FontStyle.Regular, GraphicsUnit.Point, 134);
            BackColor = ShellBack;
            ApplyWindowIcon();

            InitializeControls();
            LoadAll();
            ApplyOzonSellerDefaults();
            Shown += delegate { InitializeBrowser(null, EventArgs.Empty); };
        }

        private void ApplyWindowIcon()
        {
            try
            {
                Icon icon = Icon.ExtractAssociatedIcon(Application.ExecutablePath);
                if (icon != null)
                {
                    Icon = icon;
                }
            }
            catch
            {
            }
        }

        private void InitializeControls()
        {
            Panel header = BuildHeaderPanel();

            _mainTabs = new TabControl();
            _mainTabs.Dock = DockStyle.Fill;
            _mainTabs.ItemSize = new Size(156, 56);
            _mainTabs.SizeMode = TabSizeMode.Fixed;
            _mainTabs.Font = new Font("Microsoft YaHei UI", 10F, FontStyle.Bold, GraphicsUnit.Point, 134);
            _mainTabs.Appearance = TabAppearance.FlatButtons;
            _mainTabs.DrawMode = TabDrawMode.OwnerDrawFixed;
            _mainTabs.DrawItem += DrawMainTab;
            _mainTabs.TabPages.Add(BuildOverviewTab());
            _mainTabs.TabPages.Add(BuildConfigTab());
            _mainTabs.TabPages.Add(BuildAssetsTab());
            _mainTabs.TabPages.Add(BuildLanguageTab());
            _mainTabs.TabPages.Add(BuildAutomationTab());
            _mainTabs.TabPages.Add(BuildBrowserTab());

            _statusStrip = new StatusStrip();
            _statusStrip.SizingGrip = false;
            _statusStrip.BackColor = Color.FromArgb(39, 50, 43);
            _statusLabel = new ToolStripStatusLabel();
            _statusLabel.Text = "Ready";
            _statusLabel.ForeColor = Color.White;
            _statusStrip.Items.Add(_statusLabel);

            Controls.Add(_mainTabs);
            Controls.Add(header);
            Controls.Add(_statusStrip);
        }

        private void DrawMainTab(object sender, DrawItemEventArgs e)
        {
            TabControl tabs = sender as TabControl;
            if (tabs == null || e.Index < 0 || e.Index >= tabs.TabPages.Count)
            {
                return;
            }

            e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
            Rectangle bounds = tabs.GetTabRect(e.Index);
            Rectangle clear = bounds;
            clear.Inflate(2, 8);
            using (SolidBrush clearBrush = new SolidBrush(ShellBack))
            {
                e.Graphics.FillRectangle(clearBrush, clear);
            }

            bounds.Inflate(-14, -10);
            bool selected = e.Index == tabs.SelectedIndex;

            Color fill = selected ? Color.FromArgb(232, 255, 248, 241) : Color.FromArgb(120, 255, 255, 252);
            Color border = selected ? Color.FromArgb(182, 137, 207, 189) : Color.Transparent;
            Color text = selected ? TextStrong : TextMuted;

            using (GraphicsPath path = CreateRoundPath(bounds, 22))
            using (SolidBrush brush = new SolidBrush(fill))
            using (Pen pen = new Pen(border))
            {
                e.Graphics.FillPath(brush, path);
                if (selected)
                {
                    e.Graphics.DrawPath(pen, path);
                }
            }

            string icon = string.Empty;
            if (e.Index == 0) icon = "⚙  ";
            else if (e.Index == 1) icon = "🛠  ";
            else if (e.Index == 2) icon = "📦  ";
            else if (e.Index == 3) icon = "🌐  ";
            else if (e.Index == 4) icon = "🚚  ";
            else if (e.Index == 5) icon = "🔎  ";

            TextRenderer.DrawText(
                e.Graphics,
                icon + tabs.TabPages[e.Index].Text,
                tabs.Font,
                bounds,
                text,
                TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter | TextFormatFlags.EndEllipsis);
        }

        private Panel BuildHeaderPanel()
        {
            GradientPanel panel = new GradientPanel();
            panel.Dock = DockStyle.Top;
            panel.Height = 76;
            panel.StartColor = Color.FromArgb(252, 250, 244);
            panel.EndColor = Color.FromArgb(244, 247, 239);
            panel.Angle = 0f;
            panel.Padding = new Padding(28, 10, 20, 10);

            RoundedPanel badge = new RoundedPanel();
            badge.Left = 30;
            badge.Top = 10;
            badge.Width = 54;
            badge.Height = 52;
            badge.Radius = 18;
            badge.FillColor = Color.FromArgb(255, 16, 128, 111);
            badge.BorderColor = Color.FromArgb(255, 16, 128, 111);
            badge.ShadowColor = Color.FromArgb(24, 16, 128, 111);

            Label badgeText = new Label();
            badgeText.Text = "OP";
            badgeText.ForeColor = Color.White;
            badgeText.BackColor = Color.Transparent;
            badgeText.Font = new Font("Microsoft YaHei UI", 11F, FontStyle.Bold, GraphicsUnit.Point, 134);
            badgeText.AutoSize = false;
            badgeText.Dock = DockStyle.Fill;
            badgeText.TextAlign = ContentAlignment.MiddleCenter;
            badge.Controls.Add(badgeText);

            Panel liveDot = new Panel();
            liveDot.Left = 104;
            liveDot.Top = 31;
            liveDot.Width = 12;
            liveDot.Height = 12;
            liveDot.BackColor = Color.FromArgb(94, 187, 144);
            SetRoundedRegion(liveDot, 6);

            Label title = new Label();
            title.Text = "Ozon Pilot";
            title.Font = new Font("Microsoft YaHei UI", 14F, FontStyle.Bold, GraphicsUnit.Point, 134);
            title.ForeColor = TextStrong;
            title.BackColor = Color.Transparent;
            title.AutoSize = true;
            title.Location = new Point(136, 18);

            Label subtitle = new Label();
            subtitle.Text = "管理控制台";
            subtitle.ForeColor = TextMuted;
            subtitle.BackColor = Color.Transparent;
            subtitle.AutoSize = true;
            subtitle.Font = new Font("Microsoft YaHei UI", 10F, FontStyle.Regular, GraphicsUnit.Point, 134);
            subtitle.Location = new Point(264, 21);

            Label account = new Label();
            account.Text = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "      488059663@qq.com";
            account.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            account.AutoSize = true;
            account.ForeColor = Color.FromArgb(119, 136, 154);
            account.BackColor = Color.Transparent;
            account.Font = new Font("Consolas", 9.5F, FontStyle.Regular, GraphicsUnit.Point, 0);
            account.Location = new Point(840, 24);

            Label logout = new Label();
            logout.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            logout.Text = "退出";
            logout.AutoSize = false;
            logout.Width = 74;
            logout.Height = 34;
            logout.Left = 1278;
            logout.Top = 16;
            logout.TextAlign = ContentAlignment.MiddleCenter;
            logout.ForeColor = Color.FromArgb(100, 114, 130);
            logout.BackColor = Color.FromArgb(253, 251, 246);
            logout.BorderStyle = BorderStyle.FixedSingle;
            SetRoundedRegion(logout, 12);

            panel.Controls.Add(logout);
            panel.Controls.Add(account);
            panel.Controls.Add(title);
            panel.Controls.Add(subtitle);
            panel.Controls.Add(liveDot);
            panel.Controls.Add(badge);
            return panel;
        }

        private TabPage BuildOverviewTab()
        {
            TabPage tab = CreateTabPage(T("overview"));

            FlowLayoutPanel actions = CreateActionBar();
            actions.Controls.Add(CreateButton(T("reload"), delegate { LoadAll(); }, true));

            actions.Controls.Add(CreateButton(T("brake"), EmergencyStopCurrentProcess, false));

            Label loopLabel = new Label();
            loopLabel.Text = T("loopCount");
            loopLabel.AutoSize = true;
            loopLabel.Margin = new Padding(20, 9, 4, 0);
            actions.Controls.Add(loopLabel);
            _autoLoopCountBox = new TextBox();
            _autoLoopCountBox.Width = 54;
            _autoLoopCountBox.Text = "1";
            actions.Controls.Add(_autoLoopCountBox);
            actions.Controls.Add(CreateButton(T("fullAuto"), RunFullAutoLoop, true));

            TableLayoutPanel cards = new TableLayoutPanel();
            cards.Dock = DockStyle.Top;
            cards.Height = 124;
            cards.ColumnCount = 3;
            cards.RowCount = 1;
            cards.Padding = new Padding(12, 0, 12, 0);
            cards.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 33.33F));
            cards.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 33.33F));
            cards.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 33.34F));

            cards.Controls.Add(CreateStatCard(T("categoryNodes"), "0", T("categoryNodesDesc"), PilotGreen, out _cardCategoryValue), 0, 0);
            cards.Controls.Add(CreateStatCard(T("feeRules"), "0", T("feeRulesDesc"), Color.FromArgb(42, 126, 190), out _cardFeeValue), 1, 0);
            cards.Controls.Add(CreateStatCard(T("pluginFiles"), "0", T("pluginFilesDesc"), Color.FromArgb(205, 127, 30), out _cardPluginValue), 2, 0);

            SplitContainer split = new SplitContainer();
            split.Dock = DockStyle.Fill;
            split.SplitterDistance = 800;
            split.Panel1.Padding = new Padding(12, 8, 6, 12);
            split.Panel2.Padding = new Padding(6, 8, 12, 12);

            _overviewBox = new TextBox();
            _overviewBox.Multiline = true;
            _overviewBox.ReadOnly = true;
            _overviewBox.ScrollBars = ScrollBars.Vertical;
            _overviewBox.BackColor = Color.FromArgb(255, 253, 248);
            _overviewBox.BorderStyle = BorderStyle.FixedSingle;
            _overviewBox.Font = new Font("Microsoft YaHei UI", 9F);

            TextBox quickGuide = new TextBox();
            quickGuide.Multiline = true;
            quickGuide.ReadOnly = true;
            quickGuide.ScrollBars = ScrollBars.Vertical;
            quickGuide.BackColor = Color.FromArgb(255, 253, 248);
            quickGuide.BorderStyle = BorderStyle.FixedSingle;
            quickGuide.Font = new Font("Microsoft YaHei UI", 9F);
            quickGuide.Text = T("quickGuide");

            split.Panel1.Controls.Add(WrapWithGroup(T("overviewDetails"), _overviewBox));
            split.Panel2.Controls.Add(WrapWithGroup(T("quickStart"), quickGuide));

            tab.Controls.Add(split);
            tab.Controls.Add(cards);
            tab.Controls.Add(actions);
            return tab;
        }

        private TabPage BuildConfigTab()
        {
            TabPage tab = CreateTabPage(T("config"));

            FlowLayoutPanel actions = CreateActionBar();
            actions.Controls.Add(CreateButton(T("reloadConfig"), delegate { LoadConfig(); UpdateOverview(); }, true));
            actions.Controls.Add(CreateButton(T("saveConfig"), SaveConfig, false));

            _configGrid = new PropertyGrid();
            _configGrid.Dock = DockStyle.Fill;
            _configGrid.HelpVisible = true;
            _configGrid.ToolbarVisible = false;
            _configGrid.PropertySort = PropertySort.Categorized;
            _configGrid.BackColor = Color.FromArgb(255, 253, 248);
            _configGrid.ViewBackColor = Color.FromArgb(255, 253, 248);

            Panel body = new Panel();
            body.Dock = DockStyle.Fill;
            body.Padding = new Padding(12);
            body.Controls.Add(WrapWithGroup(T("filterConfig"), _configGrid));

            tab.Controls.Add(body);
            tab.Controls.Add(actions);
            return tab;
        }

        private TabPage BuildAssetsTab()
        {
            TabPage tab = CreateTabPage(T("assets"));

            FlowLayoutPanel actions = CreateActionBar();
            actions.Controls.Add(CreateButton(T("reloadAssets"), delegate { LoadAssets(); }, true));

            Label searchLabel = new Label();
            searchLabel.Text = T("searchAssets");
            searchLabel.AutoSize = true;
            searchLabel.Margin = new Padding(20, 9, 4, 0);
            actions.Controls.Add(searchLabel);

            _assetSearchBox = new TextBox();
            _assetSearchBox.Width = 260;
            _assetSearchBox.TextChanged += FilterAssetsChanged;
            actions.Controls.Add(_assetSearchBox);

            SplitContainer split = new SplitContainer();
            split.Dock = DockStyle.Fill;
            split.SplitterDistance = 460;
            split.Panel1.Padding = new Padding(12, 10, 6, 12);
            split.Panel2.Padding = new Padding(6, 10, 12, 12);

            _categoryTree = new TreeView();
            _categoryTree.Dock = DockStyle.Fill;
            _categoryTree.BackColor = Color.FromArgb(255, 253, 248);
            _categoryTree.NodeMouseDoubleClick += UseSelectedCategoryForAutomation;

            _feeGrid = CreateGrid();
            _feeGrid.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            _feeGrid.MultiSelect = false;
            _feeGrid.CellDoubleClick += UseSelectedFeeRule;

            split.Panel1.Controls.Add(WrapWithGroup(T("categoryTree"), _categoryTree));
            split.Panel2.Controls.Add(WrapWithGroup(T("feeRuleTable"), _feeGrid));

            tab.Controls.Add(split);
            tab.Controls.Add(actions);
            return tab;
        }

        private TabPage BuildLanguageTab()
        {
            TabPage tab = CreateTabPage(T("language"));

            Panel body = new Panel();
            body.Dock = DockStyle.Fill;
            body.Padding = new Padding(18);

            RoundedPanel card = new RoundedPanel();
            card.Dock = DockStyle.Top;
            card.Height = 250;
            card.Padding = new Padding(26, 24, 26, 24);
            card.FillColor = Color.FromArgb(250, 252, 255);
            card.BorderColor = Color.FromArgb(215, 226, 242);

            Label title = new Label();
            title.Text = T("languageTitle");
            title.Font = new Font("Microsoft YaHei UI", 15F, FontStyle.Bold, GraphicsUnit.Point, 134);
            title.ForeColor = Color.FromArgb(28, 45, 73);
            title.AutoSize = true;
            title.Location = new Point(24, 18);

            Label desc = new Label();
            desc.Text = T("languageDesc");
            desc.ForeColor = Color.FromArgb(93, 105, 126);
            desc.AutoSize = false;
            desc.Width = 780;
            desc.Height = 48;
            desc.Location = new Point(24, 56);

            Label comboLabel = new Label();
            comboLabel.Text = T("languageCurrent");
            comboLabel.ForeColor = Color.FromArgb(60, 72, 97);
            comboLabel.AutoSize = true;
            comboLabel.Location = new Point(24, 118);

            _languageComboBox = new ComboBox();
            _languageComboBox.DropDownStyle = ComboBoxStyle.DropDownList;
            _languageComboBox.Width = 260;
            _languageComboBox.Location = new Point(24, 144);
            _languageComboBox.Items.AddRange(new object[] { "简体中文", "English", "Русский" });
            _languageComboBox.SelectedIndex = LanguageIndexFromCode(_uiLanguage);

            Button apply = CreateButton(T("applyLanguage"), ApplyLanguageSelection, true);
            apply.Location = new Point(304, 140);

            Label note = new Label();
            note.Text = T("languageNote");
            note.ForeColor = Color.FromArgb(123, 132, 150);
            note.AutoSize = false;
            note.Width = 900;
            note.Height = 52;
            note.Location = new Point(24, 188);

            card.Controls.Add(title);
            card.Controls.Add(desc);
            card.Controls.Add(comboLabel);
            card.Controls.Add(_languageComboBox);
            card.Controls.Add(apply);
            card.Controls.Add(note);

            body.Controls.Add(card);
            tab.Controls.Add(body);
            return tab;
        }
        private TabPage BuildAutomationTab()
        {
            TabPage tab = CreateTabPage("Auto Sourcing");

            FlowLayoutPanel actions = CreateActionBar();
            actions.Controls.Add(CreateButton("Run 1688 selection", RunAutoSourcing, true));

            SplitContainer main = new SplitContainer();
            main.Dock = DockStyle.Fill;
            main.SplitterDistance = 420;
            main.Panel1.Padding = new Padding(12);
            main.Panel2.Padding = new Padding(0, 12, 12, 12);

            Panel editor = new Panel();
            editor.Dock = DockStyle.Fill;
            editor.BackColor = Color.FromArgb(255, 253, 248);
            editor.Padding = new Padding(16);

            int top = 18;
            int labelLeft = 16;
            int inputLeft = 150;
            int labelWidth = 120;
            int inputWidth = 230;

            editor.Controls.Add(CreateFormLabel("Keywords", labelLeft, top, labelWidth));
            _autoKeywordsBox = CreateTextBox(inputLeft, top, inputWidth, "家居收纳\r\n宠物慢食碗\r\n厨房置物架");
            _autoKeywordsBox.Multiline = true;
            _autoKeywordsBox.Height = 88;
            editor.Controls.Add(_autoKeywordsBox);
            top += 100;

            Label browserMode = new Label();
            browserMode.Text = "1688 Source";
            browserMode.Left = labelLeft;
            browserMode.Top = top + 6;
            browserMode.Width = labelWidth;
            editor.Controls.Add(browserMode);

            Label browserModeValue = new Label();
            browserModeValue.Text = "Plugin browser session";
            browserModeValue.Left = inputLeft;
            browserModeValue.Top = top + 6;
            browserModeValue.Width = inputWidth;
            browserModeValue.ForeColor = PilotGreen;
            editor.Controls.Add(browserModeValue);

            _autoProviderBox = new ComboBox();
            _autoProviderBox.DropDownStyle = ComboBoxStyle.DropDownList;
            _autoProviderBox.Left = inputLeft;
            _autoProviderBox.Top = top;
            _autoProviderBox.Width = inputWidth;
            _autoProviderBox.Items.AddRange(new object[] { "browser" });
            _autoProviderBox.SelectedIndex = 0;
            top += 34;

            FlowLayoutPanel quickActions = new FlowLayoutPanel();
            quickActions.Left = inputLeft;
            quickActions.Top = top;
            quickActions.Width = 260;
            quickActions.Height = 74;
            quickActions.WrapContents = true;
            quickActions.Controls.Add(CreateButton("Run", RunAutoSourcing, true));
            quickActions.Controls.Add(CreateButton("Upload", UploadSelectedToOzon, false));
            quickActions.Controls.Add(CreateButton("Export", ExportAutoCandidates, false));
            editor.Controls.Add(quickActions);
            top += 78;

            _autoApiKeyBox = CreateTextBox(inputLeft, top, inputWidth, string.Empty);
            _autoApiKeyBox.Visible = false;

            _autoApiSecretBox = CreateTextBox(inputLeft, top, inputWidth, string.Empty);
            _autoApiSecretBox.Visible = false;

            editor.Controls.Add(CreateFormLabel("Per keyword", labelLeft, top, labelWidth));
            _autoPerKeywordBox = CreateTextBox(inputLeft, top, inputWidth, "5");
            editor.Controls.Add(_autoPerKeywordBox);
            top += 34;

            editor.Controls.Add(CreateFormLabel("Detail limit", labelLeft, top, labelWidth));
            _autoDetailLimitBox = CreateTextBox(inputLeft, top, inputWidth, "12");
            editor.Controls.Add(_autoDetailLimitBox);
            top += 34;

            editor.Controls.Add(CreateFormLabel("RUB/CNY", labelLeft, top, labelWidth));
            _autoRubRateBox = CreateTextBox(inputLeft, top, inputWidth, "12.5");
            editor.Controls.Add(_autoRubRateBox);
            top += 34;

            editor.Controls.Add(CreateFormLabel("Ozon Category", labelLeft, top, labelWidth));
            _autoCategoryIdBox = CreateTextBox(inputLeft, top, inputWidth, "0");
            editor.Controls.Add(_autoCategoryIdBox);
            top += 34;

            editor.Controls.Add(CreateFormLabel("Ozon Type", labelLeft, top, labelWidth));
            _autoTypeIdBox = CreateTextBox(inputLeft, top, inputWidth, "0");
            editor.Controls.Add(_autoTypeIdBox);
            top += 34;

            editor.Controls.Add(CreateFormLabel("Min price x", labelLeft, top, labelWidth));
            _autoPriceMultiplierBox = CreateTextBox(inputLeft, top, inputWidth, "2.2");
            editor.Controls.Add(_autoPriceMultiplierBox);
            top += 34;

            editor.Controls.Add(CreateFormLabel("Ozon Client-Id", labelLeft, top, labelWidth));
            _ozonClientIdBox = CreateTextBox(inputLeft, top, inputWidth, DefaultOzonClientId);
            editor.Controls.Add(_ozonClientIdBox);
            top += 34;

            editor.Controls.Add(CreateFormLabel("Ozon Api-Key", labelLeft, top, labelWidth));
            _ozonApiKeyBox = CreateTextBox(inputLeft, top, inputWidth, DefaultOzonApiKey);
            editor.Controls.Add(_ozonApiKeyBox);
            top += 42;

            Label hint = new Label();
            hint.Text = "先在插件浏览器登录 1688，再点 Run。程序会复用同一个浏览器会话搜索并抓详情。Ozon 字段只在上传时需要。";
            hint.Left = labelLeft;
            hint.Top = top;
            hint.Width = 360;
            hint.Height = 60;
            hint.ForeColor = Color.FromArgb(96, 98, 102);
            editor.Controls.Add(hint);

            SplitContainer right = new SplitContainer();
            right.Dock = DockStyle.Fill;
            right.Orientation = Orientation.Horizontal;
            right.SplitterDistance = 430;

            _autoResultGrid = CreateGrid();
            _autoResultGrid.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            _autoResultGrid.MultiSelect = true;

            _autoLogBox = new TextBox();
            _autoLogBox.Multiline = true;
            _autoLogBox.ReadOnly = true;
            _autoLogBox.ScrollBars = ScrollBars.Vertical;
            _autoLogBox.BackColor = Color.FromArgb(255, 253, 248);
            _autoLogBox.BorderStyle = BorderStyle.FixedSingle;
            _autoLogBox.Font = new Font("Consolas", 9F, FontStyle.Regular, GraphicsUnit.Point, 0);

            right.Panel1.Controls.Add(WrapWithGroup("Candidates", _autoResultGrid));
            right.Panel2.Controls.Add(WrapWithGroup("Automation log", _autoLogBox));
            main.Panel1.Controls.Add(WrapWithGroup("Automation settings", editor));
            main.Panel2.Controls.Add(right);

            tab.Controls.Add(main);
            tab.Controls.Add(actions);
            return tab;
        }

        private TabPage BuildBrowserTab()
        {
            TabPage tab = CreateTabPage(T("browser"));

            FlowLayoutPanel actions = CreateActionBar();
            actions.Controls.Add(CreateButton(T("initBrowser"), InitializeBrowser, true));
            actions.Controls.Add(CreateButton(T("openUrl"), NavigateBrowser, false));

            Label urlLabel = new Label();
            urlLabel.Text = "网址：";
            urlLabel.AutoSize = true;
            urlLabel.Margin = new Padding(20, 9, 4, 0);
            actions.Controls.Add(urlLabel);

            _browserUrlBox = new TextBox();
            _browserUrlBox.Width = 420;
            _browserUrlBox.Text = "https://www.1688.com/";
            actions.Controls.Add(_browserUrlBox);

            Panel summary = new Panel();
            summary.Dock = DockStyle.Top;
            summary.Height = 54;
            summary.Padding = new Padding(16, 12, 16, 0);
            summary.BackColor = ShellBack;

            _browserStatusLabel = new Label();
            _browserStatusLabel.Dock = DockStyle.Top;
            _browserStatusLabel.Height = 24;
            _browserStatusLabel.ForeColor = Color.FromArgb(96, 98, 102);
            _browserStatusLabel.Text = "浏览器尚未初始化。";
            summary.Controls.Add(_browserStatusLabel);

            Panel body = new Panel();
            body.Dock = DockStyle.Fill;
            body.Padding = new Padding(12);

            _browser = new WebView2();
            body.Controls.Add(WrapWithGroup("1688 插件运行区", _browser));

            tab.Controls.Add(body);
            tab.Controls.Add(summary);
            tab.Controls.Add(actions);
            return tab;
        }

        private void LoadAll()
        {
            _assetErrorMessage = null;

            SetStatus("正在加载恢复工作台...");
            LoadConfig();

            try
            {
                LoadAssets();
            }
            catch (Exception ex)
            {
                _assetErrorMessage = ex.Message;
                ClearAssetViews();
            }

            UpdateOverview();

            if (!string.IsNullOrEmpty(_assetErrorMessage))
            {
                SetStatus("类目/规则读取失败：" + _assetErrorMessage);
                return;
            }

            SetStatus("恢复工作台加载完成。");
        }

        private void LoadConfig()
        {
            EnsureSnapshot();
            _snapshot.Config = ConfigService.Load(_paths.ConfigFile);
            _uiLanguage = NormalizeUiLanguage(_snapshot.Config == null ? null : _snapshot.Config.UiLanguage);
            _configGrid.SelectedObject = _snapshot.Config;
            if (_languageComboBox != null)
            {
                _languageComboBox.SelectedIndex = LanguageIndexFromCode(_uiLanguage);
            }
            ApplyOzonSellerDefaults();
        }

        private void ApplyOzonSellerDefaults()
        {
            if (_ozonClientIdBox != null && string.IsNullOrWhiteSpace(_ozonClientIdBox.Text))
            {
                _ozonClientIdBox.Text = DefaultOzonClientId;
            }

            if (_ozonApiKeyBox != null && string.IsNullOrWhiteSpace(_ozonApiKeyBox.Text))
            {
                _ozonApiKeyBox.Text = DefaultOzonApiKey;
            }
        }

        private void SaveConfig(object sender, EventArgs e)
        {
            try
            {
                AppConfig config = _configGrid.SelectedObject as AppConfig;
                if (config == null)
                {
                    return;
                }

                config.UiLanguage = _uiLanguage;
                ConfigService.Save(_paths.ConfigFile, config);
                UpdateOverview();
                SetStatus("配置已保存。");
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, ex.ToString(), "保存失败", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void LoadAssets()
        {
            EnsureSnapshot();
            _snapshot.Categories = AssetCatalogService.LoadCategories(_paths.CategoryFile);
            _snapshot.FeeRules = AssetCatalogService.LoadFeeRules(_paths.FeeFile);
            _assetErrorMessage = null;
            ApplyAssetFilter();
            UpdateOverview();
        }

        private void UpdateOverview()
        {
            if (_snapshot == null)
            {
                return;
            }

            int categoryCount = _snapshot.Categories == null ? 0 : AssetCatalogService.CountCategories(_snapshot.Categories);
            int feeCount = _snapshot.FeeRules == null ? 0 : _snapshot.FeeRules.Count;
            int pluginFileCount = Directory.Exists(_paths.Plugin1688Folder)
                ? Directory.GetFiles(_paths.Plugin1688Folder, "*", SearchOption.AllDirectories).Length
                : 0;

            if (_cardCategoryValue != null) _cardCategoryValue.Text = categoryCount.ToString();
            if (_cardFeeValue != null) _cardFeeValue.Text = feeCount.ToString();
            if (_cardPluginValue != null) _cardPluginValue.Text = pluginFileCount.ToString();

            string updaterPath = _paths.FindUpdaterExecutable();
            StringBuilder builder = new StringBuilder();
            builder.AppendLine(T("resourcePaths"));
            builder.AppendLine("--------");
            builder.AppendLine(T("workRoot") + _paths.WorkRoot);
            builder.AppendLine(T("baselineRoot") + _paths.BaselineRoot);
            builder.AppendLine(T("configFile") + _paths.ConfigFile);
            builder.AppendLine(T("plugin1688") + _paths.Plugin1688Folder);
            builder.AppendLine(T("updaterPath") + SafeValue(updaterPath));
            builder.AppendLine();
            builder.AppendLine(T("recoveryStatus"));
            builder.AppendLine("--------");
            builder.AppendLine(T("categoryCountLine") + categoryCount);
            builder.AppendLine(T("feeCountLine") + feeCount);
            builder.AppendLine(T("pluginCountLine") + pluginFileCount);

            if (!string.IsNullOrEmpty(_assetErrorMessage))
            {
                builder.AppendLine(T("assetLoadFailed"));
                builder.AppendLine(T("failureReason") + _assetErrorMessage);
            }

            builder.AppendLine();
            builder.AppendLine(T("currentFilters"));
            builder.AppendLine("------------");
            builder.AppendLine(T("saveDir") + SafeValue(_snapshot.Config == null ? null : _snapshot.Config.SaveUrl));
            builder.AppendLine(T("priceRange") + (_snapshot.Config == null ? 0 : _snapshot.Config.MinPirce) + " ~ " + (_snapshot.Config == null ? 0 : _snapshot.Config.MaxPrice));
            builder.AppendLine(T("minProfit") + (_snapshot.Config == null ? 0 : _snapshot.Config.MinProfitPer) + "%");
            builder.AppendLine(T("defaultShipping") + (_snapshot.Config == null ? 0 : _snapshot.Config.DeliveryFee));
            builder.AppendLine(T("enable1688") + YesNo(_snapshot.Config != null && _snapshot.Config.Is1688));
            builder.AppendLine(T("autoExport") + YesNo(_snapshot.Config != null && _snapshot.Config.IsAutoExport));
            builder.AppendLine(T("cloudFilter") + YesNo(_snapshot.Config != null && _snapshot.Config.IsCloudFilter));

            if (!string.IsNullOrEmpty(_fullAutoReport))
            {
                builder.AppendLine();
                builder.AppendLine("全链路自动循环简报");
                builder.AppendLine("----------------");
                builder.AppendLine(_fullAutoReport);
            }

            _overviewBox.Text = builder.ToString();
        }

        private void ExportFeeRules(object sender, EventArgs e)
        {
            if (_snapshot == null || _snapshot.FeeRules == null || _snapshot.FeeRules.Count == 0)
            {
                SetStatus("当前没有可导出的运费规则。");
                return;
            }

            string path = Path.Combine(_paths.WorkRoot, "运费规则导出.xlsx");
            AssetCatalogService.ExportFeeRulesToExcel(_snapshot.FeeRules, path);
            _paths.OpenPath(path);
            SetStatus("运费规则已导出到：" + path);
        }

        private void FilterAssetsChanged(object sender, EventArgs e)
        {
            ApplyAssetFilter();
        }

        private void ApplyAssetFilter()
        {
            if (_snapshot == null || _snapshot.Categories == null || _snapshot.FeeRules == null)
            {
                return;
            }

            string keyword = _assetSearchBox == null ? string.Empty : _assetSearchBox.Text;
            List<CategoryNode> categories = AssetCatalogService.FilterCategories(_snapshot.Categories, keyword);
            List<FeeRule> rules = AssetCatalogService.FilterFeeRules(_snapshot.FeeRules, keyword);

            _categoryTree.BeginUpdate();
            _categoryTree.Nodes.Clear();

            int i;
            for (i = 0; i < categories.Count; i++)
            {
                _categoryTree.Nodes.Add(BuildTreeNode(categories[i]));
            }

            if (_categoryTree.Nodes.Count > 0)
            {
                _categoryTree.Nodes[0].Expand();
            }

            _categoryTree.EndUpdate();
            _feeGrid.DataSource = null;
            _feeGrid.DataSource = BuildFeeRuleDisplayRows(rules);
        }

        private void UseSelectedFeeRule(object sender, EventArgs e)
        {
            FeeRule rule = GetSelectedFeeRule();
            if (rule == null)
            {
                SetStatus("请先在运费规则表中选中一条规则。");
                return;
            }

            CategoryNode mappedNode = FindBestLeafCategoryForFeeRule(_snapshot == null ? null : _snapshot.Categories, rule);
            if (mappedNode == null)
            {
                SetStatus("这条运费规则没有匹配到可上传的 Ozon 类目，请从左侧类目树双击一个叶子类目。");
                return;
            }

            string categoryId = ResolveUploadCategoryId(mappedNode);
            string typeId = string.IsNullOrEmpty(mappedNode.DescriptionTypeId) ? "0" : mappedNode.DescriptionTypeId;
            string keyword = !string.IsNullOrEmpty(mappedNode.DescriptionTypeName)
                ? mappedNode.DescriptionTypeName
                : (!string.IsNullOrEmpty(rule.Category2) ? rule.Category2 : mappedNode.DescriptionCategoryName);
            FillAutomationCategory(categoryId, typeId, keyword);
            _mainTabs.SelectedIndex = 4;
            SetStatus(T("ruleAppliedToAuto") + " [" + categoryId + "/" + typeId + "]");
        }

        private void UseSelectedCategoryForAutomation(object sender, TreeNodeMouseClickEventArgs e)
        {
            CategoryNode node = e == null ? null : e.Node.Tag as CategoryNode;
            if (node == null)
            {
                return;
            }

            if (string.IsNullOrEmpty(node.DescriptionTypeId) || node.DescriptionTypeId == "0")
            {
                node = FindFirstOzonLeafCategory(node);
                if (node == null)
                {
                    SetStatus("Selected category has no uploadable Ozon type.");
                    return;
                }
            }

            string categoryId = ResolveUploadCategoryId(node);
            string typeId = string.IsNullOrEmpty(node.DescriptionTypeId) ? "0" : node.DescriptionTypeId;
            string keyword = !string.IsNullOrEmpty(node.DescriptionTypeName)
                ? node.DescriptionTypeName
                : node.DescriptionCategoryName;
            FillAutomationCategory(categoryId, typeId, keyword);
            SetStatus("已把类目树选中的 Category/Type 填入 Auto Sourcing：" + categoryId + " / " + typeId);
        }

        private void FillAutomationCategory(string categoryId, string typeId)
        {
            FillAutomationCategory(categoryId, typeId, null);
        }

        private void FillAutomationCategory(string categoryId, string typeId, string keyword)
        {
            if (_autoCategoryIdBox != null && !string.IsNullOrEmpty(categoryId))
            {
                _autoCategoryIdBox.Text = categoryId;
            }

            if (_autoTypeIdBox != null && !string.IsNullOrEmpty(typeId))
            {
                _autoTypeIdBox.Text = typeId;
            }

            if (_autoKeywordsBox != null && !string.IsNullOrEmpty(keyword))
            {
                _autoKeywordsBox.Text = keyword;
            }
            else if (_autoKeywordsBox != null)
            {
                _autoKeywordsBox.Text = string.Empty;
            }
        }

        private List<FeeRuleDisplayRow> BuildFeeRuleDisplayRows(IList<FeeRule> rules)
        {
            List<FeeRuleDisplayRow> rows = new List<FeeRuleDisplayRow>();
            for (int i = 0; rules != null && i < rules.Count; i++)
            {
                FeeRule rule = rules[i];
                if (rule == null)
                {
                    continue;
                }

                rows.Add(new FeeRuleDisplayRow
                {
                    Rule = rule,
                    Id = rule.Id,
                    CategoryId1 = rule.CategoryId1,
                    CategoryId2 = rule.CategoryId2,
                    Category1 = BuildBilingualCategoryName(rule.Category1),
                    Category2 = BuildBilingualCategoryName(rule.Category2),
                    FBS = rule.FBS,
                    FBS1500 = rule.FBS1500,
                    FBS5000 = rule.FBS5000,
                    FBP = rule.FBP,
                    FBP1500 = rule.FBP1500,
                    FBP5000 = rule.FBP5000,
                    FBO = rule.FBO,
                    FBO1500 = rule.FBO1500,
                    FBO5000 = rule.FBO5000
                });
            }

            return rows;
        }

        private string BuildBilingualCategoryName(string name)
        {
            string text = string.IsNullOrEmpty(name) ? string.Empty : name.Trim();
            if (string.IsNullOrEmpty(text) || IsAsciiKeyword(text))
            {
                return text;
            }

            string english = BuildEnglishKeywordFromRules(text);
            if (string.IsNullOrEmpty(english))
            {
                english = MakeFallbackKeywordFromCategory(text);
            }

            return text + " / " + english;
        }

        private string BuildEnglishKeyword(string keyword)
        {
            string text = string.IsNullOrEmpty(keyword) ? string.Empty : keyword.Trim();
            if (string.IsNullOrEmpty(text))
            {
                return string.Empty;
            }

            if (IsAsciiKeyword(text))
            {
                return text;
            }

            string ruleKeyword = BuildEnglishKeywordFromRules(text);
            if (!string.IsNullOrEmpty(ruleKeyword))
            {
                return ruleKeyword;
            }

            string aiKeyword = _automationService.GenerateEnglishCategoryKeyword(text);
            if (!string.IsNullOrEmpty(aiKeyword) &&
                aiKeyword.IndexOf("general merchandise", StringComparison.OrdinalIgnoreCase) < 0)
            {
                return aiKeyword;
            }

            return MakeFallbackKeywordFromCategory(text);
        }

        private string BuildEnglishKeywordFromRules(string text)
        {
            if (string.IsNullOrEmpty(text))
            {
                return string.Empty;
            }

            if (ContainsAny(text, "\u60c5\u8da3", "\u6210\u4eba", "\u6027\u7231", "\u907f\u5b55", "\u5b89\u5168\u5957", "\u79c1\u5904"))
            {
                return "adult wellness product";
            }

            string[,] map = new string[,]
            {
                { "\u5c0f\u5de5\u5177", "hand tool accessory" },
                { "\u5546\u4e1a\u8bbe\u5907", "commercial equipment" },
                { "\u6253\u5370\u8017\u6750", "printer supplies" },
                { "\u58a8\u76d2", "printer ink cartridge" },
                { "\u7852\u9f13", "printer toner cartridge" },
                { "\u624b\u673a", "phone accessory" },
                { "\u5e73\u677f", "tablet accessory" },
                { "\u7535\u8111", "computer accessory" },
                { "\u6570\u7801", "electronics accessory" },
                { "\u7535\u5b50", "electronics accessory" },
                { "\u6c7d\u8f66", "car accessory" },
                { "\u6469\u6258", "motorcycle accessory" },
                { "\u81ea\u884c\u8f66", "bicycle accessory" },
                { "\u5ba0\u7269", "pet supplies" },
                { "\u732b", "cat supplies" },
                { "\u72d7", "dog supplies" },
                { "\u53a8\u623f", "kitchen organizer" },
                { "\u9910\u5177", "kitchen utensil" },
                { "\u70d8\u7119", "baking tool" },
                { "\u6536\u7eb3", "storage organizer" },
                { "\u6d74\u5ba4", "bathroom organizer" },
                { "\u536b\u6d74", "bathroom accessory" },
                { "\u5bb6\u5c45", "home improvement product" },
                { "\u4f4f\u5b85", "home improvement product" },
                { "\u56ed\u827a", "garden tool" },
                { "\u82b1\u56ed", "garden tool" },
                { "\u529e\u516c", "office organizer" },
                { "\u6587\u5177", "stationery supplies" },
                { "\u4e66\u5199", "writing supplies" },
                { "\u65c5\u884c", "travel organizer" },
                { "\u884c\u674e", "travel organizer" },
                { "\u8fd0\u52a8", "sports accessory" },
                { "\u5065\u8eab", "fitness accessory" },
                { "\u6237\u5916", "outdoor accessory" },
                { "\u513f\u7ae5", "kids product" },
                { "\u5a74\u513f", "baby product" },
                { "\u6bcd\u5a74", "baby product" },
                { "\u73a9\u5177", "kids toy" },
                { "\u6e05\u6d01", "cleaning tool" },
                { "\u5de5\u5177", "tool accessory" },
                { "\u670d\u88c5", "clothing accessory" },
                { "\u978b", "shoe accessory" },
                { "\u5185\u8863", "underwear organizer" },
                { "\u914d\u9970", "fashion accessory" },
                { "\u706f", "led light" },
                { "\u8282\u5e86", "party decoration" },
                { "\u88c5\u9970", "home decoration" },
                { "\u7f8e\u5bb9", "beauty tool" },
                { "\u5316\u5986", "makeup organizer" },
                { "\u62a4\u80a4", "skin care tool" },
                { "\u5065\u5eb7", "health care product" },
                { "\u533b\u7597", "health care accessory" },
                { "\u98df\u54c1", "food storage container" },
                { "\u5305\u88c5", "packing supplies" },
                { "\u5c55\u793a", "display stand" },
                { "\u652f\u67b6", "holder stand" },
                { "\u67b6", "storage rack" },
                { "\u76d2", "storage box" },
                { "\u888b", "storage bag" },
                { "\u7845\u80f6", "silicone product" },
                { "\u5851\u6599", "plastic product" }
            };

            for (int i = 0; i < map.GetLength(0); i++)
            {
                if (text.IndexOf(map[i, 0], StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    return map[i, 1];
                }
            }

            return string.Empty;
        }

        private string MakeFallbackKeywordFromCategory(string category)
        {
            string text = category ?? string.Empty;
            if (ContainsAny(text, "\u5c0f\u5de5\u5177"))
            {
                return "hand tool accessory";
            }
            if (ContainsAny(text, "\u5546\u4e1a\u8bbe\u5907"))
            {
                return "commercial equipment";
            }
            if (ContainsAny(text, "\u7535\u5b50\u4ea7\u54c1"))
            {
                return "electronics accessory";
            }
            if (ContainsAny(text, "\u65e5\u7528\u54c1"))
            {
                return "household supplies";
            }

            return "category specific product";
        }

        private bool ContainsAny(string text, params string[] needles)
        {
            for (int i = 0; i < needles.Length; i++)
            {
                if (text.IndexOf(needles[i], StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    return true;
                }
            }

            return false;
        }

        private bool IsAsciiKeyword(string text)
        {
            for (int i = 0; i < text.Length; i++)
            {
                if (text[i] > 127)
                {
                    return false;
                }
            }

            return true;
        }

        private bool StartCurrentProcess(string name)
        {
            if (_currentProcessCancel != null && !_currentProcessCancel.IsCancellationRequested)
            {
                SetStatus("Another process is already running: " + _currentProcessName);
                return false;
            }

            _currentProcessCancel = new CancellationTokenSource();
            _currentProcessName = name;
            return true;
        }

        private void FinishCurrentProcess()
        {
            if (_currentProcessCancel != null)
            {
                _currentProcessCancel.Dispose();
                _currentProcessCancel = null;
            }

            _currentProcessName = null;
        }

        private void ThrowIfCurrentProcessStopped()
        {
            if (_currentProcessCancel != null && _currentProcessCancel.IsCancellationRequested)
            {
                throw new OperationCanceledException();
            }
        }

        private void EmergencyStopCurrentProcess(object sender, EventArgs e)
        {
            if (_currentProcessCancel != null)
            {
                _currentProcessCancel.Cancel();
            }

            try
            {
                if (_browser != null && _browser.CoreWebView2 != null)
                {
                    _browser.CoreWebView2.Stop();
                }
            }
            catch
            {
            }

            _fullAutoReport = "Emergency brake pressed: " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + Environment.NewLine +
                "Current process: " + SafeValue(_currentProcessName) + Environment.NewLine +
                "The running browser/navigation task was asked to stop.";
            _mainTabs.SelectedIndex = 0;
            UpdateOverview();
            SetStatus("Emergency brake requested.");
        }

        private List<string> GetSeedKeywords(IList<SourcingSeed> seeds)
        {
            List<string> values = new List<string>();
            for (int i = 0; seeds != null && i < seeds.Count; i++)
            {
                if (seeds[i] != null && !string.IsNullOrEmpty(seeds[i].Keyword))
                {
                    values.Add(seeds[i].Keyword);
                }
            }

            return values;
        }

        private async void RunAutoSourcing(object sender, EventArgs e)
        {
            if (!StartCurrentProcess("Run 1688 selection"))
            {
                return;
            }

            StringBuilder report = new StringBuilder();
            try
            {
                EnsureSnapshot();
                SourcingOptions options = ReadSourcingOptions();
                List<SourcingSeed> seeds = ReadSourcingSeeds();
                if (_browser == null || _browser.CoreWebView2 == null)
                {
                    SetStatus("Please initialize the plugin browser and login to 1688 first.");
                    _mainTabs.SelectedTab = _mainTabs.TabPages[_mainTabs.TabPages.Count - 1];
                    return;
                }

                _mainTabs.SelectedIndex = 5;
                report.AppendLine("Manual 1688 selection brief");
                report.AppendLine("---------------------------");
                report.AppendLine("Start: " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
                report.AppendLine("Keywords: " + string.Join(", ", GetSeedKeywords(seeds).ToArray()));
                SetStatus("Running 1688 selection...");
                ThrowIfCurrentProcessStopped();
                _lastSourcingResult = await Collect1688CandidatesFromBrowser(seeds, _snapshot.Config, options);
                ThrowIfCurrentProcessStopped();
                _autoResultGrid.DataSource = _lastSourcingResult.Products;
                WriteAutomationLog(_lastSourcingResult.Logs);
                report.AppendLine("Candidates: " + _lastSourcingResult.Products.Count);
                report.AppendLine("End: " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
                _fullAutoReport = report.ToString();
                _mainTabs.SelectedIndex = 0;
                UpdateOverview();
                SetStatus("1688 selection finished: " + _lastSourcingResult.Products.Count + " candidates.");
            }
            catch (OperationCanceledException)
            {
                report.AppendLine("Stopped by emergency brake: " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
                _fullAutoReport = report.ToString();
                _mainTabs.SelectedIndex = 0;
                UpdateOverview();
                SetStatus("Current process stopped.");
            }
            catch (Exception ex)
            {
                AppendAutomationLog("ERROR: " + ex.Message);
                report.AppendLine("ERROR: " + ex.Message);
                _fullAutoReport = report.ToString();
                _mainTabs.SelectedIndex = 0;
                UpdateOverview();
                MessageBox.Show(this, ex.ToString(), "Auto sourcing failed", MessageBoxButtons.OK, MessageBoxIcon.Error);
                SetStatus("Auto sourcing failed: " + ex.Message);
            }
            finally
            {
                FinishCurrentProcess();
            }
        }

        private async void RunFullAutoLoop(object sender, EventArgs e)
        {
            if (!StartCurrentProcess("Full auto loop"))
            {
                return;
            }

            StringBuilder report = new StringBuilder();
            try
            {
                EnsureSnapshot();
                if (_browser == null || _browser.CoreWebView2 == null)
                {
                    _mainTabs.SelectedIndex = 5;
                    SetStatus("请先初始化插件浏览器并登录 1688，然后再启动全链路循环。");
                    return;
                }

                int loopCount = (int)ParseLong(_autoLoopCountBox == null ? "1" : _autoLoopCountBox.Text);
                if (loopCount <= 0)
                {
                    loopCount = 1;
                }

                report.AppendLine("启动时间：" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
                report.AppendLine("计划循环：" + loopCount + " 次");
                SetStatus("全链路自动循环开始...");

                for (int i = 0; i < loopCount; i++)
                {
                    ThrowIfCurrentProcessStopped();
                    CategoryNode category = PickRandomOzonLeafCategory();
                    if (category == null)
                    {
                        report.AppendLine("Round " + (i + 1) + ": no usable Ozon leaf category; stopped.");
                        break;
                    }

                    string categoryKeyword = !string.IsNullOrEmpty(category.DescriptionTypeName) ? category.DescriptionTypeName : category.DescriptionCategoryName;
                    _mainTabs.SelectedIndex = 4;
                    FillAutomationCategory(ResolveUploadCategoryId(category), category.DescriptionTypeId, categoryKeyword);
                    await Task.Delay(300);

                    report.AppendLine();
                    report.AppendLine("Round " + (i + 1));
                    report.AppendLine("Ozon category: " + category.DescriptionCategoryName + " / " + category.DescriptionTypeName + " [" + ResolveUploadCategoryId(category) + "/" + category.DescriptionTypeId + "]");
                    report.AppendLine("Keyword: " + categoryKeyword);

                    _mainTabs.SelectedIndex = 5;
                    SourcingOptions options = ReadSourcingOptions();
                    List<SourcingSeed> seeds = ReadSourcingSeeds();
                    ThrowIfCurrentProcessStopped();
                    _lastSourcingResult = await Collect1688CandidatesFromBrowser(seeds, _snapshot.Config, options);
                    ThrowIfCurrentProcessStopped();
                    _autoResultGrid.DataSource = _lastSourcingResult.Products;
                    WriteAutomationLog(_lastSourcingResult.Logs);
                    report.AppendLine("选品候选：" + _lastSourcingResult.Products.Count + " 个");
                    _fullAutoReport = report.ToString();
                    _mainTabs.SelectedIndex = 0;
                    UpdateOverview();
                    SetStatus("1688 抓取完成，正在提交 Ozon 上传...");

                    List<SourceProduct> uploadProducts = PickProductsForUpload(_lastSourcingResult.Products);
                    if (uploadProducts.Count == 0)
                    {
                        report.AppendLine("上传：跳过，没有可上传候选。");
                        continue;
                    }

                    try
                    {
                        OzonImportResult uploadResult = _automationService.UploadToOzon(
                            uploadProducts,
                            options,
                            _ozonClientIdBox.Text.Trim(),
                            _ozonApiKeyBox.Text.Trim());

                        if (uploadResult.Success)
                        {
                            report.AppendLine("Ozon task submitted: " + uploadProducts.Count + " items, task_id=" + SafeValue(uploadResult.TaskId));

                            ThrowIfCurrentProcessStopped();
                            OzonImportResult importResult = await Task.Run(() => _automationService.WaitForOzonImportInfo(
                                uploadResult.TaskId,
                                _ozonClientIdBox.Text.Trim(),
                                _ozonApiKeyBox.Text.Trim(),
                                12,
                                10000));

                            report.AppendLine("Ozon import result:");
                            report.AppendLine(importResult.ImportSummary);
                            importResult = await RetryFailedOzonImportsOnce(uploadProducts, options, importResult, _ozonClientIdBox.Text.Trim(), _ozonApiKeyBox.Text.Trim(), delegate(string line)
                            {
                                report.AppendLine(line);
                            });
                            if (importResult.AcceptedOfferIds.Count > 0)
                            {
                                if (!importResult.Success)
                                {
                                    report.AppendLine("Ozon import had partial failures, but accepted offers will still continue to SKU/stock update.");
                                }

                                try
                                {
                                    List<string> readyOfferIds = await WaitForSkuCreationWithBrief(report, importResult.AcceptedOfferIds, 20, 30000);
                                    if (readyOfferIds.Count == 0)
                                    {
                                        report.AppendLine("Ozon SKU creation is still pending; stock update skipped until SKU exists.");
                                    }
                                    else
                                    {
                                        report.AppendLine("SKU 已创建，正在设置库存 100...");
                                        _fullAutoReport = report.ToString();
                                        UpdateOverview();
                                        string stockResponse = await Task.Run(() => _automationService.SetOzonStockTo100(
                                            readyOfferIds,
                                            _ozonClientIdBox.Text.Trim(),
                                            _ozonApiKeyBox.Text.Trim()));
                                        report.AppendLine("Ozon stock set to 100: " + stockResponse);
                                    }
                                }
                                catch (Exception stockEx)
                                {
                                    report.AppendLine("Ozon stock update failed: " + stockEx.Message);
                                    string stockLogPath = WriteExceptionLog("ozon-stock", stockEx, report.ToString());
                                    if (!string.IsNullOrEmpty(stockLogPath))
                                    {
                                        report.AppendLine("库存异常日志：" + stockLogPath);
                                    }
                                }
                            }

                            if (!importResult.Success)
                            {
                                report.AppendLine("Ozon import was not accepted. Check missing attributes/category requirements above.");
                            }
                        }
                        else
                        {
                            report.AppendLine("Ozon upload failed: " + uploadResult.ErrorMessage);
                            report.AppendLine("本轮未成功上架到 Ozon，请以上传失败/校验失败信息为准。");
                        }
                    }
                    catch (Exception uploadEx)
                    {
                        report.AppendLine("Ozon upload exception: " + uploadEx.Message);
                        string uploadLogPath = WriteExceptionLog("ozon-upload", uploadEx, report.ToString());
                        if (!string.IsNullOrEmpty(uploadLogPath))
                        {
                            report.AppendLine("上传异常日志：" + uploadLogPath);
                        }
                        report.AppendLine("本轮未成功上架到 Ozon，请先处理上述异常后再重试。");
                    }

                    _fullAutoReport = report.ToString();
                    UpdateOverview();
                }

                report.AppendLine();
                report.AppendLine("结束时间：" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
                _fullAutoReport = report.ToString();
                _mainTabs.SelectedIndex = 0;
                UpdateOverview();
                SetStatus("全链路自动循环完成。");
            }
            catch (OperationCanceledException)
            {
                report.AppendLine();
                report.AppendLine("Stopped by emergency brake: " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
                _fullAutoReport = report.ToString();
                _mainTabs.SelectedIndex = 0;
                UpdateOverview();
                SetStatus("Current process stopped.");
            }
            catch (Exception ex)
            {
                report.AppendLine();
                report.AppendLine("异常：" + ex.Message);
                string fullRunLogPath = WriteExceptionLog("full-auto", ex, report.ToString());
                if (!string.IsNullOrEmpty(fullRunLogPath))
                {
                    report.AppendLine("完整异常日志：" + fullRunLogPath);
                }
                _fullAutoReport = report.ToString();
                _mainTabs.SelectedIndex = 0;
                UpdateOverview();
                MessageBox.Show(this, ex.ToString(), "全链路自动循环失败", MessageBoxButtons.OK, MessageBoxIcon.Error);
                SetStatus("全链路自动循环失败：" + ex.Message);
            }
            finally
            {
                FinishCurrentProcess();
            }
        }

        private FeeRule PickRandomSecondLevelRule()
        {
            if (_snapshot == null || _snapshot.FeeRules == null || _snapshot.FeeRules.Count == 0)
            {
                return null;
            }

            List<FeeRule> candidates = new List<FeeRule>();
            for (int i = 0; i < _snapshot.FeeRules.Count; i++)
            {
                FeeRule rule = _snapshot.FeeRules[i];
                if (rule != null && rule.CategoryId1 > 0 && rule.CategoryId2 > 0 && !string.IsNullOrEmpty(rule.Category2))
                {
                    candidates.Add(rule);
                }
            }

            if (candidates.Count == 0)
            {
                return null;
            }

            return candidates[_random.Next(candidates.Count)];
        }

        private async Task<List<string>> WaitForSkuCreationWithBrief(StringBuilder report, IList<string> offerIds, int attempts, int delayMs)
        {
            List<string> ready = new List<string>();
            if (offerIds == null || offerIds.Count == 0)
            {
                report.AppendLine("SKU wait skipped: no accepted offer_id values.");
                return ready;
            }

            _mainTabs.SelectedIndex = 0;
            report.AppendLine("绛夊緟 Ozon 鍒涘缓 SKU...");
            _fullAutoReport = report.ToString();
            UpdateOverview();

            int maxAttempts = attempts <= 0 ? 20 : attempts;
            int wait = delayMs <= 0 ? 30000 : delayMs;
            for (int attempt = 0; attempt < maxAttempts; attempt++)
            {
                ThrowIfCurrentProcessStopped();
                if (attempt > 0)
                {
                    await Task.Delay(wait);
                }

                Dictionary<string, string> skuByOffer = await Task.Run(() => _automationService.GetOzonSkuMap(
                    offerIds,
                    _ozonClientIdBox.Text.Trim(),
                    _ozonApiKeyBox.Text.Trim()));

                ready.Clear();
                for (int i = 0; i < offerIds.Count; i++)
                {
                    string offerId = offerIds[i];
                    if (skuByOffer.ContainsKey(offerId) && !string.IsNullOrEmpty(skuByOffer[offerId]))
                    {
                        ready.Add(offerId);
                    }
                }

                report.AppendLine("SKU 检查 " + (attempt + 1) + "/" + maxAttempts + "：" + ready.Count + "/" + offerIds.Count + " 已创建");
                int shown = 0;
                foreach (KeyValuePair<string, string> pair in skuByOffer)
                {
                    if (shown >= 5)
                    {
                        break;
                    }

                    report.AppendLine("  " + pair.Key + " -> " + (string.IsNullOrEmpty(pair.Value) ? "绛夊緟 SKU" : "SKU " + pair.Value));
                    shown += 1;
                }

                _fullAutoReport = report.ToString();
                UpdateOverview();
                SetStatus("等待 Ozon SKU 创建：" + ready.Count + "/" + offerIds.Count + " 已创建");

                if (ready.Count >= offerIds.Count)
                {
                    report.AppendLine("SKU 创建完成，可以开始设置库存。");
                    _fullAutoReport = report.ToString();
                    UpdateOverview();
                    return new List<string>(ready);
                }
            }

            report.AppendLine("SKU 等待超时：Ozon 仍在处理，已创建 " + ready.Count + "/" + offerIds.Count + "。");
            _fullAutoReport = report.ToString();
            UpdateOverview();
            return new List<string>(ready);
        }

        private CategoryNode PickRandomOzonLeafCategory()
        {
            EnsureSnapshot();
            List<CategoryNode> candidates = new List<CategoryNode>();
            CollectOzonLeafCategories(_snapshot == null ? null : _snapshot.Categories, candidates);
            if (candidates.Count == 0)
            {
                return null;
            }

            return candidates[_random.Next(candidates.Count)];
        }

        private void CollectOzonLeafCategories(IList<CategoryNode> nodes, IList<CategoryNode> output)
        {
            for (int i = 0; nodes != null && i < nodes.Count; i++)
            {
                CategoryNode node = nodes[i];
                if (node == null || node.Disabled)
                {
                    continue;
                }

                if (IsHiddenFromAutoLoopCategory(node))
                {
                    continue;
                }

                if (!string.IsNullOrEmpty(node.DescriptionCategoryId) &&
                    !string.IsNullOrEmpty(node.DescriptionTypeId) &&
                    node.DescriptionCategoryId != "0" &&
                    node.DescriptionTypeId != "0")
                {
                    output.Add(node);
                }

                CollectOzonLeafCategories(node.Children, output);
            }
        }

        private CategoryNode FindFirstOzonLeafCategory(CategoryNode node)
        {
            if (node == null || node.Disabled)
            {
                return null;
            }

            if (!string.IsNullOrEmpty(node.DescriptionCategoryId) &&
                !string.IsNullOrEmpty(node.DescriptionTypeId) &&
                node.DescriptionCategoryId != "0" &&
                node.DescriptionTypeId != "0")
            {
                return node;
            }

            for (int i = 0; node.Children != null && i < node.Children.Count; i++)
            {
                CategoryNode found = FindFirstOzonLeafCategory(node.Children[i]);
                if (found != null)
                {
                    return found;
                }
            }

            return null;
        }

        private static bool IsHiddenFromAutoLoopCategory(CategoryNode node)
        {
            if (node == null)
            {
                return false;
            }

            string text = ((node.DescriptionCategoryName ?? string.Empty) + " " + (node.DescriptionTypeName ?? string.Empty)).ToLowerInvariant();
            string[] hiddenTerms = new string[]
            {
                "成人", "情趣", "避孕", "安全套", "成人用品", "adult", "sex", "sexual", "condom", "эрот"
            };

            for (int i = 0; i < hiddenTerms.Length; i++)
            {
                if (text.IndexOf(hiddenTerms[i], StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    return true;
                }
            }

            return false;
        }

        private static string ResolveUploadCategoryId(CategoryNode node)
        {
            if (node == null)
            {
                return "0";
            }

            string categoryId = !string.IsNullOrEmpty(node.UploadCategoryId)
                ? node.UploadCategoryId
                : node.DescriptionCategoryId;
            return string.IsNullOrEmpty(categoryId) ? "0" : categoryId;
        }

        private CategoryNode FindBestLeafCategoryForFeeRule(IList<CategoryNode> nodes, FeeRule rule)
        {
            if (rule == null)
            {
                return null;
            }

            CategoryNode bySecondLevel = FindLeafCategoryByCategoryId(nodes, rule.CategoryId2, new List<long>());
            if (bySecondLevel != null)
            {
                return bySecondLevel;
            }

            return FindLeafCategoryByCategoryId(nodes, rule.CategoryId1, new List<long>());
        }

        private CategoryNode FindLeafCategoryByCategoryId(IList<CategoryNode> nodes, long targetCategoryId, IList<long> ancestors)
        {
            if (targetCategoryId <= 0)
            {
                return null;
            }

            for (int i = 0; nodes != null && i < nodes.Count; i++)
            {
                CategoryNode node = nodes[i];
                if (node == null || node.Disabled)
                {
                    continue;
                }

                List<long> nextAncestors = new List<long>(ancestors);
                long nodeCategoryId = ParseLong(node.DescriptionCategoryId);
                if (nodeCategoryId > 0 && !nextAncestors.Contains(nodeCategoryId))
                {
                    nextAncestors.Add(nodeCategoryId);
                }

                string uploadCategoryIdText = ResolveUploadCategoryId(node);
                long uploadCategoryId = ParseLong(uploadCategoryIdText);
                bool matches = nodeCategoryId == targetCategoryId ||
                    uploadCategoryId == targetCategoryId ||
                    nextAncestors.Contains(targetCategoryId);

                bool isLeaf = !string.IsNullOrEmpty(node.DescriptionTypeId) &&
                    node.DescriptionTypeId != "0" &&
                    !string.IsNullOrEmpty(node.DescriptionCategoryId) &&
                    node.DescriptionCategoryId != "0";
                if (matches && isLeaf)
                {
                    return node;
                }

                CategoryNode found = FindLeafCategoryByCategoryId(node.Children, targetCategoryId, nextAncestors);
                if (found != null)
                {
                    return found;
                }
            }

            return null;
        }

        private List<SourceProduct> PickProductsForUpload(IList<SourceProduct> products)
        {
            List<SourceProduct> selected = new List<SourceProduct>();
            if (products == null)
            {
                return selected;
            }

            for (int i = 0; i < products.Count && selected.Count < 3; i++)
            {
                SourceProduct product = products[i];
                if (product != null &&
                    !IsComplianceRestrictedProduct(product) &&
                    string.Equals(product.Decision, "Go", StringComparison.OrdinalIgnoreCase) &&
                    IsStrongKeywordMatch(product))
                {
                    selected.Add(product);
                }
            }

            if (selected.Count == 0)
            {
                for (int i = 0; i < products.Count; i++)
                {
                    SourceProduct product = products[i];
                    if (product != null &&
                        !IsComplianceRestrictedProduct(product) &&
                        string.Equals(product.Decision, "Watch", StringComparison.OrdinalIgnoreCase) &&
                        IsStrongKeywordMatch(product))
                    {
                        selected.Add(product);
                        break;
                    }
                }
            }

            return selected;
        }

        private static bool IsComplianceRestrictedProduct(SourceProduct product)
        {
            if (product == null)
            {
                return false;
            }

            StringBuilder builder = new StringBuilder();
            builder.Append(product.Title).Append(' ')
                .Append(product.Keyword).Append(' ')
                .Append(product.Reason).Append(' ')
                .Append(product.ShopName);

            foreach (KeyValuePair<string, string> pair in product.Attributes)
            {
                builder.Append(' ').Append(pair.Key).Append(' ').Append(pair.Value);
            }

            string text = builder.ToString().ToLowerInvariant();
            string[] restrictedTerms = new string[]
            {
                "轮椅", "助行", "拐杖", "康复", "护理床", "病床", "矫形", "残疾", "残障",
                "医疗", "医用", "药品", "药物", "保健", "wheelchair", "walker", "crutch",
                "rehabilitation", "orthopedic", "disabled", "invalid", "medical", "medicine",
                "инвалид", "коляск", "медицин",
                "杀虫", "杀虫剂", "除虫", "灭虫", "驱虫", "农药", "气雾剂", "insecticide", "pesticide", "repellent",
                "инсектицид", "пестицид", "аэрозоль", "яд"
            };

            for (int i = 0; i < restrictedTerms.Length; i++)
            {
                if (text.IndexOf(restrictedTerms[i], StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    return true;
                }
            }

            return false;
        }

        private async Task<SourcingResult> Collect1688CandidatesFromBrowser(IList<SourcingSeed> seeds, AppConfig config, SourcingOptions options)
        {
            SourcingResult result = new SourcingResult();
            Dictionary<string, SourceProduct> byOfferId = new Dictionary<string, SourceProduct>();
            int perKeywordLimit = options.PerKeywordLimit <= 0 ? 5 : options.PerKeywordLimit;
            int detailLimit = options.DetailLimit <= 0 ? 12 : options.DetailLimit;

            for (int i = 0; seeds != null && i < seeds.Count; i++)
            {
                SourcingSeed seed = seeds[i];
                if (seed == null || string.IsNullOrEmpty(seed.Keyword))
                {
                    continue;
                }

                AppendAutomationLog("Search 1688 in browser: " + seed.Keyword);
                string url = "https://s.1688.com/selloffer/offer_search.htm?keywords=" + Encode1688Keyword(seed.Keyword);
                await NavigateAndWait(url, 1500);
                bool rendered = await WaitForSearchResults(seed.Keyword, 35000);
                result.Logs.Add(seed.Keyword + " render wait: " + (rendered ? "ready" : "timeout"));
                List<SourceProduct> cards = await ScrapeSearchPage(seed.Keyword, perKeywordLimit);
                if (cards.Count == 0)
                {
                    AppendAutomationLog("Search cards not ready, retrying after extra wait: " + seed.Keyword);
                    await Task.Delay(8000);
                    cards = await ScrapeSearchPage(seed.Keyword, perKeywordLimit);
                }
                result.Logs.Add(seed.Keyword + " search cards: " + cards.Count);

                for (int j = 0; j < cards.Count; j++)
                {
                    SourceProduct card = cards[j];
                    if (!string.IsNullOrEmpty(card.OfferId) && !byOfferId.ContainsKey(card.OfferId))
                    {
                        byOfferId[card.OfferId] = card;
                    }
                }
            }

            List<SourceProduct> candidates = new List<SourceProduct>(byOfferId.Values);
            int maxDetails = Math.Min(detailLimit, candidates.Count);
            for (int i = 0; i < maxDetails; i++)
            {
                SourceProduct candidate = candidates[i];
                AppendAutomationLog("Detail " + (i + 1) + "/" + maxDetails + ": " + candidate.OfferId);
                await NavigateAndWait(candidate.SourceUrl, 5500);
                SourceProduct detail = await ScrapeDetailPage(candidate);
                MergeBrowserProduct(candidate, detail);
                ScoreBrowserCandidate(candidate, config);
            }

            for (int i = maxDetails; i < candidates.Count; i++)
            {
                ScoreBrowserCandidate(candidates[i], config);
            }

            candidates.Sort(delegate(SourceProduct left, SourceProduct right)
            {
                int decision = BrowserDecisionRank(right.Decision).CompareTo(BrowserDecisionRank(left.Decision));
                if (decision != 0) return decision;
                return right.Score.CompareTo(left.Score);
            });

            result.Products.AddRange(candidates);
            result.Logs.Add("Browser scrape finished: " + result.Products.Count + " candidates.");
            return result;
        }

        private async Task NavigateAndWait(string url, int waitAfterMs)
        {
            TaskCompletionSource<bool> completion = new TaskCompletionSource<bool>();
            EventHandler<CoreWebView2NavigationCompletedEventArgs> handler = null;
            handler = delegate(object sender, CoreWebView2NavigationCompletedEventArgs e)
            {
                _browser.CoreWebView2.NavigationCompleted -= handler;
                completion.TrySetResult(true);
            };

            _browser.CoreWebView2.NavigationCompleted += handler;
            _browser.CoreWebView2.Navigate(url);
            Task finished = await Task.WhenAny(completion.Task, Task.Delay(25000));
            if (finished != completion.Task)
            {
                _browser.CoreWebView2.NavigationCompleted -= handler;
            }

            await Task.Delay(waitAfterMs);
        }

        private async Task<bool> WaitForSearchResults(string keyword, int timeoutMs)
        {
            DateTime deadline = DateTime.Now.AddMilliseconds(timeoutMs <= 0 ? 30000 : timeoutMs);
            string script = @"
(function(){
  var links = document.querySelectorAll('a[href*=""offer""], a[href*=""detail.1688.com""]');
  if (links && links.length > 0) return true;
  return /offer\/\d+\.html|offerId=\d+|object_id=\d+/.test(document.body ? document.body.innerHTML : '');
})();";

            while (DateTime.Now < deadline)
            {
                ThrowIfCurrentProcessStopped();
                try
                {
                    string value = await _browser.CoreWebView2.ExecuteScriptAsync(script);
                    if (string.Equals(value, "true", StringComparison.OrdinalIgnoreCase))
                    {
                        return true;
                    }
                }
                catch
                {
                }

                AppendAutomationLog("Waiting 1688 render: " + keyword);
                await Task.Delay(1000);
            }

            return false;
        }

        private async Task<List<SourceProduct>> ScrapeSearchPage(string keyword, int limit)
        {
            string script = @"
(function(){
  function text(node){ return (node && (node.innerText || node.textContent) || '').replace(/\s+/g,' ').trim(); }
  function attr(node, name){ return node ? (node.getAttribute(name) || '') : ''; }
  function abs(url){ try { return new URL(url, location.href).href; } catch(e) { return url || ''; } }
  function offerId(url){ var m=String(url||'').match(/offer\/(\d+)\.html|offerId=(\d+)|object_id=(\d+)/i); return m ? (m[1]||m[2]||m[3]||'') : ''; }
  var anchors = Array.prototype.slice.call(document.querySelectorAll('a[href*=""offer""], a[href*=""detail.1688.com""]'));
  var seen = {};
  var items = [];
  anchors.forEach(function(a){
    var href = abs(a.href);
    var id = offerId(href);
    if(!id || seen[id]) return;
    seen[id] = true;
    var card = a;
    for(var i=0;i<6 && card && card.parentElement;i++){
      card = card.parentElement;
      if(text(card).length > 30) break;
    }
    var img = card ? card.querySelector('img') : null;
    var title = attr(a,'title') || attr(img,'alt') || text(a) || text(card).slice(0,80);
    var blockText = text(card);
    var priceMatch = blockText.match(/[¥￥]\s*([0-9]+(?:\.[0-9]+)?)/) || blockText.match(/([0-9]+(?:\.[0-9]+)?)\s*元/);
    var salesMatch = blockText.match(/([0-9.]+)\s*(万)?\s*(?:人付款|成交|已售|销量)/);
    var sales = 0;
    if(salesMatch){ sales = parseFloat(salesMatch[1] || '0') || 0; if(salesMatch[2]) sales *= 10000; }
    items.push({
      OfferId:id,
      Title:title,
      SourceUrl:'https://detail.1688.com/offer/' + id + '.html',
      PriceText:priceMatch ? priceMatch[1] : '',
      PriceCny:priceMatch ? parseFloat(priceMatch[1]) || 0 : 0,
      SalesCount:Math.round(sales),
      ShopName:'',
      ShopUrl:'',
      MainImage:img ? abs(img.currentSrc || img.src || attr(img,'data-src')) : '',
      Images:img ? [abs(img.currentSrc || img.src || attr(img,'data-src'))] : [],
      Keyword:'" + EscapeJavaScript(keyword) + @"',
      Attributes:{}
    });
  });
  return items.slice(0," + limit + @");
})();";
            string json = await _browser.CoreWebView2.ExecuteScriptAsync(script);
            return ParseBrowserProducts(json);
        }

        private async Task<SourceProduct> ScrapeDetailPage(SourceProduct fallback)
        {
            string script = @"
(function(){
  function text(node){ return (node && (node.innerText || node.textContent) || '').replace(/\s+/g,' ').trim(); }
  function attr(node, name){ return node ? (node.getAttribute(name) || '') : ''; }
  function abs(url){ try { return new URL(url, location.href).href; } catch(e) { return url || ''; } }
  function offerId(url){ var m=String(url||location.href).match(/offer\/(\d+)\.html|offerId=(\d+)/i); return m ? (m[1]||m[2]||'') : ''; }
  var title = attr(document.querySelector('meta[property=""og:title""]'),'content') || text(document.querySelector('h1')) || document.title;
  var body = text(document.body);
  var priceMatch = body.match(/[¥￥]\s*([0-9]+(?:\.[0-9]+)?)/) || body.match(/价格\s*[:：]?\s*([0-9]+(?:\.[0-9]+)?)/);
  var salesMatch = body.match(/([0-9.]+)\s*(万)?\s*(?:人付款|成交|已售|销量)/);
  var sales = 0;
  if(salesMatch){ sales = parseFloat(salesMatch[1] || '0') || 0; if(salesMatch[2]) sales *= 10000; }
  function bestImg(node){
    var srcset = attr(node,'srcset');
    if(srcset){
      var first = srcset.split(',')[0].trim().split(' ')[0];
      if(first) return abs(first);
    }
    return abs(node.currentSrc || node.src || attr(node,'data-src') || attr(node,'data-lazy-src') || attr(node,'data-ks-lazyload') || attr(node,'data-img-url'));
  }
  var imgs = Array.prototype.slice.call(document.querySelectorAll('img, source'))
    .map(function(img){ return bestImg(img); })
    .filter(function(url){ return /alicdn|cbu|1688/i.test(url) && !/avatar|logo|icon|lazyload/i.test(url); });
  var unique = [];
  imgs.forEach(function(url){ if(url && unique.indexOf(url) < 0) unique.push(url.split('?')[0]); });
  var attrs = {};
  function addAttr(k, v){
    k = (k || '').replace(/[：:]+$/,'').trim();
    v = (v || '').trim();
    if(!k || !v || k.length > 30 || v.length > 120) return;
    if(Object.keys(attrs).length >= 40) return;
    if(!attrs[k]) attrs[k] = v;
  }
  Array.prototype.slice.call(document.querySelectorAll('tr')).forEach(function(row){
    var cells = row.querySelectorAll('th,td');
    if(cells.length >= 2){
      addAttr(text(cells[0]), text(cells[cells.length - 1]));
    }
  });
  Array.prototype.slice.call(document.querySelectorAll('dl')).forEach(function(row){
    addAttr(text(row.querySelector('dt')), text(row.querySelector('dd')));
  });
  Array.prototype.slice.call(document.querySelectorAll('li, .offer-attr, .mod-detail-attributes')).slice(0,120).forEach(function(row){
    var t = text(row);
    var m = t.match(/^(.{1,30})[:：]\s*(.{1,120})$/);
    if(m) addAttr(m[1], m[2]);
  });
  return {
    OfferId:offerId(location.href),
    Title:title.replace(/\s*[-_].*1688.*$/i,'').slice(0,160),
    SourceUrl:location.href,
    PriceText:priceMatch ? priceMatch[1] : '',
    PriceCny:priceMatch ? parseFloat(priceMatch[1]) || 0 : 0,
    SalesCount:Math.round(sales),
    ShopName:text(document.querySelector('[class*=""shop""] a, [class*=""company""] a')).slice(0,80),
    ShopUrl:abs(attr(document.querySelector('a[href*=""shop.1688.com""]'),'href')),
    MainImage:unique[0] || '',
    Images:unique.slice(0,12),
    Attributes:attrs
  };
})();";
            string json = await _browser.CoreWebView2.ExecuteScriptAsync(script);
            List<SourceProduct> products = ParseBrowserProducts("[" + json + "]");
            return products.Count > 0 ? products[0] : fallback;
        }

        private List<SourceProduct> ParseBrowserProducts(string json)
        {
            List<SourceProduct> products = new List<SourceProduct>();
            if (string.IsNullOrEmpty(json) || json == "null")
            {
                return products;
            }

            JArray array = JArray.Parse(json);
            for (int i = 0; i < array.Count; i++)
            {
                JObject item = array[i] as JObject;
                if (item == null)
                {
                    continue;
                }

                SourceProduct product = new SourceProduct();
                product.OfferId = JsonString(item, "OfferId");
                product.Title = JsonString(item, "Title");
                product.SourceUrl = JsonString(item, "SourceUrl");
                product.PriceText = JsonString(item, "PriceText");
                product.PriceCny = JsonDecimal(item, "PriceCny");
                product.SalesCount = (int)JsonDecimal(item, "SalesCount");
                product.ShopName = JsonString(item, "ShopName");
                product.ShopUrl = JsonString(item, "ShopUrl");
                product.MainImage = JsonString(item, "MainImage");
                product.Keyword = JsonString(item, "Keyword");

                JArray images = item["Images"] as JArray;
                for (int j = 0; images != null && j < images.Count; j++)
                {
                    string image = Convert.ToString(images[j]);
                    if (!string.IsNullOrEmpty(image) && !product.Images.Contains(image))
                    {
                        product.Images.Add(image);
                    }
                }

                JObject attrs = item["Attributes"] as JObject;
                if (attrs != null)
                {
                    foreach (JProperty property in attrs.Properties())
                    {
                        product.Attributes[property.Name] = Convert.ToString(property.Value);
                    }
                }

                if (!string.IsNullOrEmpty(product.OfferId))
                {
                    products.Add(product);
                }
            }

            return products;
        }

        private void MergeBrowserProduct(SourceProduct target, SourceProduct detail)
        {
            if (target == null || detail == null)
            {
                return;
            }

            if (!string.IsNullOrEmpty(detail.Title)) target.Title = detail.Title;
            if (!string.IsNullOrEmpty(detail.SourceUrl)) target.SourceUrl = detail.SourceUrl;
            if (detail.PriceCny > 0) target.PriceCny = detail.PriceCny;
            if (!string.IsNullOrEmpty(detail.PriceText)) target.PriceText = detail.PriceText;
            if (detail.SalesCount > 0) target.SalesCount = detail.SalesCount;
            if (!string.IsNullOrEmpty(detail.ShopName)) target.ShopName = detail.ShopName;
            if (!string.IsNullOrEmpty(detail.ShopUrl)) target.ShopUrl = detail.ShopUrl;
            if (!string.IsNullOrEmpty(detail.MainImage)) target.MainImage = detail.MainImage;
            if (detail.Images.Count > 0) target.Images = detail.Images;
            foreach (KeyValuePair<string, string> pair in detail.Attributes)
            {
                target.Attributes[pair.Key] = pair.Value;
            }
        }

        private void ScoreBrowserCandidate(SourceProduct product, AppConfig config)
        {
            decimal score = 0m;
            List<string> reasons = new List<string>();
            if (product.PriceCny > 0) score += 20m; else reasons.Add("missing price");
            if (product.SalesCount > 0) score += Math.Min(25m, product.SalesCount / 20m); else reasons.Add("missing sales");
            if (!string.IsNullOrEmpty(product.MainImage) || product.Images.Count > 0) score += 20m; else reasons.Add("missing image");
            if (product.Attributes.Count > 0) score += Math.Min(15m, product.Attributes.Count * 2m);
            if (!string.IsNullOrEmpty(product.ShopName)) score += 10m;
            int relevance = ComputeKeywordRelevance(product);
            if (relevance >= 55)
            {
                score += 20m;
            }
            else if (relevance >= 35)
            {
                score += 10m;
            }
            else
            {
                reasons.Add("weak keyword relevance");
            }

            if (config != null && config.MinSaleNum > 0 && product.SalesCount > 0 && product.SalesCount < config.MinSaleNum) reasons.Add("sales below config");
            if (IsComplianceRestrictedProduct(product)) reasons.Add("restricted compliance category");
            product.Score = Math.Round(score, 2);
            product.Decision = reasons.Count == 0 && score >= 45m ? "Go" : (score >= 30m ? "Watch" : "No-Go");
            product.Reason = reasons.Count == 0 ? "browser scrape signals look usable" : string.Join("; ", reasons.ToArray());
        }

        private static bool IsStrongKeywordMatch(SourceProduct product)
        {
            return ComputeKeywordRelevance(product) >= 35;
        }

        private static int ComputeKeywordRelevance(SourceProduct product)
        {
            if (product == null)
            {
                return 0;
            }

            string keyword = Convert.ToString(product.Keyword ?? string.Empty).Trim();
            if (string.IsNullOrEmpty(keyword))
            {
                return 100;
            }

            string haystack = BuildKeywordHaystack(product);
            string normalizedHaystack = NormalizeKeywordText(haystack);
            string normalizedTitle = NormalizeKeywordText(product.Title);
            string normalizedKeyword = NormalizeKeywordText(keyword);
            if (string.IsNullOrEmpty(normalizedHaystack) || string.IsNullOrEmpty(normalizedKeyword))
            {
                return 0;
            }

            int score = 0;
            if (normalizedTitle.IndexOf(normalizedKeyword, StringComparison.OrdinalIgnoreCase) >= 0)
            {
                score += 60;
            }
            else if (normalizedHaystack.IndexOf(normalizedKeyword, StringComparison.OrdinalIgnoreCase) >= 0)
            {
                score += 45;
            }

            List<string> tokens = BuildKeywordTokens(keyword);
            for (int i = 0; i < tokens.Count; i++)
            {
                string token = tokens[i];
                string normalizedToken = NormalizeKeywordText(token);
                if (string.IsNullOrEmpty(normalizedToken))
                {
                    continue;
                }

                if (normalizedTitle.IndexOf(normalizedToken, StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    score += normalizedToken.Length >= 4 ? 18 : 12;
                }
                else if (normalizedHaystack.IndexOf(normalizedToken, StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    score += normalizedToken.Length >= 4 ? 10 : 6;
                }
            }

            return Math.Min(100, score);
        }

        private static string BuildKeywordHaystack(SourceProduct product)
        {
            StringBuilder builder = new StringBuilder();
            builder.Append(product.Title).Append(' ')
                .Append(product.Keyword).Append(' ')
                .Append(product.ShopName).Append(' ');

            foreach (KeyValuePair<string, string> pair in product.Attributes)
            {
                builder.Append(pair.Key).Append(' ').Append(pair.Value).Append(' ');
            }

            return builder.ToString();
        }

        private static List<string> BuildKeywordTokens(string keyword)
        {
            List<string> tokens = new List<string>();
            Dictionary<string, bool> unique = new Dictionary<string, bool>(StringComparer.OrdinalIgnoreCase);
            string[] pieces = (keyword ?? string.Empty)
                .Replace("/", " ")
                .Replace("\\", " ")
                .Replace(",", " ")
                .Replace("，", " ")
                .Replace("|", " ")
                .Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);

            for (int i = 0; i < pieces.Length; i++)
            {
                string piece = pieces[i].Trim();
                AddKeywordToken(tokens, unique, piece);
                if (ContainsChinese(piece) && piece.Length >= 4)
                {
                    for (int j = 0; j <= piece.Length - 2; j++)
                    {
                        AddKeywordToken(tokens, unique, piece.Substring(j, 2));
                    }
                }
            }

            return tokens;
        }

        private static void AddKeywordToken(List<string> tokens, Dictionary<string, bool> unique, string token)
        {
            string normalized = NormalizeKeywordText(token);
            if (string.IsNullOrEmpty(normalized) || normalized.Length < 2 || unique.ContainsKey(normalized))
            {
                return;
            }

            unique[normalized] = true;
            tokens.Add(token);
        }

        private static string NormalizeKeywordText(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return string.Empty;
            }

            StringBuilder builder = new StringBuilder(value.Length);
            string lowered = value.Trim().ToLowerInvariant();
            for (int i = 0; i < lowered.Length; i++)
            {
                char c = lowered[i];
                if (char.IsLetterOrDigit(c) || (c >= 0x4E00 && c <= 0x9FFF))
                {
                    builder.Append(c);
                }
            }

            return builder.ToString();
        }

        private static bool ContainsChinese(string value)
        {
            for (int i = 0; !string.IsNullOrEmpty(value) && i < value.Length; i++)
            {
                char c = value[i];
                if (c >= 0x4E00 && c <= 0x9FFF)
                {
                    return true;
                }
            }

            return false;
        }

        private int BrowserDecisionRank(string decision)
        {
            if (string.Equals(decision, "Go", StringComparison.OrdinalIgnoreCase)) return 3;
            if (string.Equals(decision, "Watch", StringComparison.OrdinalIgnoreCase)) return 2;
            return 1;
        }

        private static string JsonString(JObject obj, string name)
        {
            JToken token = obj[name];
            return token == null || token.Type == JTokenType.Null ? string.Empty : Convert.ToString(token);
        }

        private static decimal JsonDecimal(JObject obj, string name)
        {
            decimal value;
            return decimal.TryParse(JsonString(obj, name), out value) ? value : 0m;
        }

        private static string EscapeJavaScript(string value)
        {
            return (value ?? string.Empty).Replace("\\", "\\\\").Replace("'", "\\'");
        }

        private static string Encode1688Keyword(string value)
        {
            byte[] bytes = Encoding.GetEncoding("GB18030").GetBytes(value ?? string.Empty);
            StringBuilder builder = new StringBuilder(bytes.Length * 3);
            for (int i = 0; i < bytes.Length; i++)
            {
                byte b = bytes[i];
                bool keep = (b >= (byte)'A' && b <= (byte)'Z') ||
                    (b >= (byte)'a' && b <= (byte)'z') ||
                    (b >= (byte)'0' && b <= (byte)'9') ||
                    b == (byte)'-' || b == (byte)'_' || b == (byte)'.' || b == (byte)'~';
                if (keep)
                {
                    builder.Append((char)b);
                }
                else if (b == (byte)' ')
                {
                    builder.Append("%20");
                }
                else
                {
                    builder.Append('%');
                    builder.Append(b.ToString("X2"));
                }
            }

            return builder.ToString();
        }

        private async void UploadSelectedToOzon(object sender, EventArgs e)
        {
            try
            {
                List<SourceProduct> selected = GetSelectedSourceProducts();
                if (selected.Count == 0)
                {
                    SetStatus("No selected candidates.");
                    return;
                }

                SourcingOptions options = ReadSourcingOptions();
                OzonImportResult result = _automationService.UploadToOzon(
                    selected,
                    options,
                    _ozonClientIdBox.Text.Trim(),
                    _ozonApiKeyBox.Text.Trim());

                if (!result.Success)
                {
                    AppendAutomationLog("Ozon upload failed: " + result.ErrorMessage);
                    SetStatus("Ozon upload failed: " + result.ErrorMessage);
                    return;
                }

                AppendAutomationLog("Ozon upload response: " + result.RawResponse);
                if (!string.IsNullOrEmpty(result.TaskId))
                {
                    AppendAutomationLog("Ozon task id: " + result.TaskId);
                    Clipboard.SetText(result.TaskId);
                    SetStatus("Ozon task submitted; checking import result...");

                    OzonImportResult importResult = await Task.Run(() => _automationService.WaitForOzonImportInfo(
                        result.TaskId,
                        _ozonClientIdBox.Text.Trim(),
                        _ozonApiKeyBox.Text.Trim(),
                        12,
                        10000));

                    AppendAutomationLog("Ozon import result: " + importResult.ImportSummary);
                    importResult = await RetryFailedOzonImportsOnce(selected, options, importResult, _ozonClientIdBox.Text.Trim(), _ozonApiKeyBox.Text.Trim(), delegate(string line)
                    {
                        AppendAutomationLog(line);
                    });
                    if (importResult.AcceptedOfferIds.Count > 0)
                    {
                        if (!importResult.Success)
                        {
                            AppendAutomationLog("Ozon import had partial failures; continuing stock update for accepted offers.");
                        }

                        try
                        {
                            AppendAutomationLog("Waiting for Ozon SKU creation...");
                            OzonSkuWaitResult skuResult = await Task.Run(() => _automationService.WaitForOzonSkuCreation(
                                importResult.AcceptedOfferIds,
                                _ozonClientIdBox.Text.Trim(),
                                _ozonApiKeyBox.Text.Trim(),
                                20,
                                30000));
                            AppendAutomationLog(skuResult.Summary);

                            if (skuResult.ReadyOfferIds.Count == 0)
                            {
                                AppendAutomationLog("Ozon SKU creation is still pending; stock update skipped until SKU exists.");
                            }
                            else
                            {
                            string stockResponse = await Task.Run(() => _automationService.SetOzonStockTo100(
                                skuResult.ReadyOfferIds,
                                _ozonClientIdBox.Text.Trim(),
                                _ozonApiKeyBox.Text.Trim()));
                            AppendAutomationLog("Ozon stock set to 100: " + stockResponse);
                            }
                        }
                        catch (Exception stockEx)
                        {
                            AppendAutomationLog("Ozon stock update failed: " + stockEx.Message);
                        }
                    }

                    if (!importResult.Success)
                    {
                        SetStatus("Ozon import returned errors. See automation log.");
                        return;
                    }
                }

                SetStatus("Ozon import accepted. Products may still need Ozon moderation.");
            }
            catch (Exception ex)
            {
                AppendAutomationLog("OZON ERROR: " + ex.Message);
                MessageBox.Show(this, ex.ToString(), "Ozon upload failed", MessageBoxButtons.OK, MessageBoxIcon.Error);
                SetStatus("Ozon upload failed: " + ex.Message);
            }
        }

        private void ExportAutoCandidates(object sender, EventArgs e)
        {
            if (_lastSourcingResult == null || _lastSourcingResult.Products.Count == 0)
            {
                SetStatus("No candidates to export.");
                return;
            }

            string path = Path.Combine(_paths.WorkRoot, "1688-ozon-candidates.json");
            _automationService.ExportCandidates(path, _lastSourcingResult.Products);
            _paths.OpenPath(path);
            SetStatus("Candidates exported: " + path);
        }

        private SourcingOptions ReadSourcingOptions()
        {
            long categoryId = ParseLong(_autoCategoryIdBox.Text);
            long typeId = ParseLong(_autoTypeIdBox.Text);
            return new SourcingOptions
            {
                Provider = Convert.ToString(_autoProviderBox.SelectedItem),
                ApiKey = _autoApiKeyBox.Text.Trim(),
                ApiSecret = _autoApiSecretBox.Text.Trim(),
                PerKeywordLimit = (int)ParseLong(_autoPerKeywordBox.Text),
                DetailLimit = (int)ParseLong(_autoDetailLimitBox.Text),
                RubPerCny = ParseDecimal(_autoRubRateBox.Text),
                OzonCategoryId = categoryId,
                OzonTypeId = typeId,
                OzonCategoryCandidateIds = BuildOzonCategoryCandidateIds(categoryId, typeId),
                PriceMultiplier = ParseDecimal(_autoPriceMultiplierBox.Text),
                MinOzonPrice = 0m,
                CurrencyCode = "CNY",
                Vat = "0",
                Config = _snapshot == null ? null : _snapshot.Config,
                FeeRules = _snapshot == null ? new List<FeeRule>() : _snapshot.FeeRules,
                FulfillmentMode = _snapshot != null && _snapshot.Config != null && _snapshot.Config.IsFbo ? "FBO" : "FBS"
            };
        }

        private async Task<OzonImportResult> RetryFailedOzonImportsOnce(IList<SourceProduct> products, SourcingOptions options, OzonImportResult importResult, string clientId, string apiKey, Action<string> log)
        {
            if (importResult == null || string.IsNullOrEmpty(importResult.ImportInfoResponse))
            {
                return importResult;
            }

            List<SourceProduct> retryProducts;
            string repairSummary = _automationService.PrepareRetryForFailedImports(products, options, importResult.ImportInfoResponse, clientId, apiKey, out retryProducts);
            if (retryProducts == null || retryProducts.Count == 0)
            {
                return importResult;
            }

            if (log != null)
            {
                log("Ozon auto-repair prepared for failed items:");
                if (!string.IsNullOrEmpty(repairSummary))
                {
                    string[] lines = repairSummary.Replace("\r", string.Empty).Split(new char[] { '\n' }, StringSplitOptions.RemoveEmptyEntries);
                    for (int i = 0; i < lines.Length; i++)
                    {
                        log("  " + lines[i]);
                    }
                }
            }

            OzonImportResult retrySubmit = _automationService.UploadToOzon(retryProducts, options, clientId, apiKey);
            if (!retrySubmit.Success)
            {
                if (log != null)
                {
                    log("Ozon auto-repair retry submit failed: " + retrySubmit.ErrorMessage);
                }

                return importResult;
            }

            if (log != null)
            {
                log("Ozon auto-repair retry task submitted: " + retryProducts.Count + " items, task_id=" + SafeValue(retrySubmit.TaskId));
            }

            OzonImportResult retryImportResult = await Task.Run(() => _automationService.WaitForOzonImportInfo(
                retrySubmit.TaskId,
                clientId,
                apiKey,
                12,
                10000));

            if (log != null)
            {
                log("Ozon auto-repair retry result:");
                log(retryImportResult.ImportSummary);
            }

            return MergeOzonImportResults(importResult, retryImportResult);
        }

        private static OzonImportResult MergeOzonImportResults(OzonImportResult original, OzonImportResult retry)
        {
            if (original == null)
            {
                return retry;
            }

            if (retry == null)
            {
                return original;
            }

            OzonImportResult merged = new OzonImportResult();
            merged.Success = original.Success || retry.Success;
            merged.TaskId = string.IsNullOrEmpty(retry.TaskId) ? original.TaskId : retry.TaskId;
            merged.RawResponse = string.IsNullOrEmpty(retry.RawResponse) ? original.RawResponse : retry.RawResponse;
            merged.ErrorMessage = string.IsNullOrEmpty(retry.ErrorMessage) ? original.ErrorMessage : retry.ErrorMessage;
            merged.ImportInfoResponse = string.IsNullOrEmpty(retry.ImportInfoResponse) ? original.ImportInfoResponse : retry.ImportInfoResponse;
            merged.ImportSummary = original.ImportSummary;
            if (!string.IsNullOrEmpty(retry.ImportSummary))
            {
                merged.ImportSummary += Environment.NewLine + "Auto-repair retry:" + Environment.NewLine + retry.ImportSummary;
            }

            AppendUniqueOfferIds(merged.AcceptedOfferIds, original.AcceptedOfferIds);
            AppendUniqueOfferIds(merged.AcceptedOfferIds, retry.AcceptedOfferIds);
            return merged;
        }

        private static void AppendUniqueOfferIds(IList<string> target, IList<string> source)
        {
            if (target == null || source == null)
            {
                return;
            }

            for (int i = 0; i < source.Count; i++)
            {
                string offerId = source[i];
                if (string.IsNullOrEmpty(offerId) || target.Contains(offerId))
                {
                    continue;
                }

                target.Add(offerId);
            }
        }

        private List<long> BuildOzonCategoryCandidateIds(long categoryId, long typeId)
        {
            List<long> candidates = new List<long>();
            if (categoryId > 0)
            {
                candidates.Add(categoryId);
            }

            EnsureSnapshot();
            AppendCategoryCandidatesByType(_snapshot == null ? null : _snapshot.Categories, typeId, candidates, new List<long>());
            return candidates;
        }

        private void AppendCategoryCandidatesByType(IList<CategoryNode> nodes, long targetTypeId, IList<long> output, IList<long> ancestors)
        {
            for (int i = 0; nodes != null && i < nodes.Count; i++)
            {
                CategoryNode node = nodes[i];
                if (node == null)
                {
                    continue;
                }

                List<long> nextAncestors = new List<long>(ancestors);
                long nodeCategoryId = ParseLong(node.DescriptionCategoryId);
                if (nodeCategoryId > 0 && !nextAncestors.Contains(nodeCategoryId))
                {
                    nextAncestors.Add(nodeCategoryId);
                }

                if (ParseLong(node.DescriptionTypeId) == targetTypeId)
                {
                    AppendUniqueLong(output, nodeCategoryId);
                    for (int ancestorIndex = nextAncestors.Count - 1; ancestorIndex >= 0; ancestorIndex--)
                    {
                        AppendUniqueLong(output, nextAncestors[ancestorIndex]);
                    }
                }

                AppendCategoryCandidatesByType(node.Children, targetTypeId, output, nextAncestors);
            }
        }

        private static void AppendUniqueLong(IList<long> values, long value)
        {
            if (value <= 0 || values == null)
            {
                return;
            }

            for (int i = 0; i < values.Count; i++)
            {
                if (values[i] == value)
                {
                    return;
                }
            }

            values.Add(value);
        }

        private List<SourcingSeed> ReadSourcingSeeds()
        {
            string text = _autoKeywordsBox == null ? string.Empty : _autoKeywordsBox.Text;
            string[] lines = text.Replace("\r", "\n").Split(new char[] { '\n' }, StringSplitOptions.RemoveEmptyEntries);
            List<SourcingSeed> seeds = new List<SourcingSeed>();
            Dictionary<string, bool> unique = new Dictionary<string, bool>();
            for (int i = 0; i < lines.Length; i++)
            {
                string keyword = lines[i].Trim();
                if (string.IsNullOrEmpty(keyword) || unique.ContainsKey(keyword))
                {
                    continue;
                }

                unique[keyword] = true;
                seeds.Add(new SourcingSeed { Keyword = keyword });
            }

            return seeds;
        }

        private List<SourceProduct> GetSelectedSourceProducts()
        {
            List<SourceProduct> products = new List<SourceProduct>();
            if (_autoResultGrid == null)
            {
                return products;
            }

            for (int i = 0; i < _autoResultGrid.SelectedRows.Count; i++)
            {
                SourceProduct product = _autoResultGrid.SelectedRows[i].DataBoundItem as SourceProduct;
                if (product != null && !products.Contains(product))
                {
                    products.Add(product);
                }
            }

            return products;
        }

        private void WriteAutomationLog(IList<string> lines)
        {
            StringBuilder builder = new StringBuilder();
            for (int i = 0; lines != null && i < lines.Count; i++)
            {
                builder.AppendLine(lines[i]);
            }

            _autoLogBox.Text = builder.ToString();
        }

        private void AppendAutomationLog(string line)
        {
            if (_autoLogBox == null)
            {
                return;
            }

            if (!string.IsNullOrEmpty(_autoLogBox.Text))
            {
                _autoLogBox.AppendText(Environment.NewLine);
            }

            _autoLogBox.AppendText(DateTime.Now.ToString("HH:mm:ss") + " " + line);
        }

        private string WriteExceptionLog(string stage, Exception ex, string reportSnapshot)
        {
            try
            {
                string logDirectory = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "logs");
                Directory.CreateDirectory(logDirectory);

                string safeStage = string.IsNullOrEmpty(stage) ? "error" : stage.Replace(' ', '-');
                string path = Path.Combine(logDirectory, DateTime.Now.ToString("yyyyMMdd-HHmmss") + "-" + safeStage + ".log");

                StringBuilder builder = new StringBuilder();
                builder.AppendLine("Stage: " + safeStage);
                builder.AppendLine("Time: " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
                builder.AppendLine();
                builder.AppendLine("Exception");
                builder.AppendLine("---------");
                builder.AppendLine(ex == null ? "(null)" : ex.ToString());

                if (!string.IsNullOrEmpty(reportSnapshot))
                {
                    builder.AppendLine();
                    builder.AppendLine("Report Snapshot");
                    builder.AppendLine("---------------");
                    builder.AppendLine(reportSnapshot);
                }

                File.WriteAllText(path, builder.ToString(), new UTF8Encoding(false));
                return path;
            }
            catch
            {
                return null;
            }
        }

        private void InitializeBrowser(object sender, EventArgs e)
        {
            try
            {
                _browserExtensionReady = false;
                _browser.CoreWebView2InitializationCompleted -= BrowserInitializationCompleted;
                _browser.CoreWebView2InitializationCompleted += BrowserInitializationCompleted;
                CoreWebView2Environment env = BrowserBootstrap.CreateEnvironment(_paths);
                if (env == null)
                {
                    _browserStatusLabel.Text = "未找到 1688 插件目录。";
                    return;
                }

                _browser.EnsureCoreWebView2Async(env);
                _browserStatusLabel.Text = "正在初始化插件浏览器...";
            }
            catch (Exception ex)
            {
                string message = ex.ToString();
                if (message.IndexOf("WebView2", StringComparison.OrdinalIgnoreCase) >= 0 && File.Exists(_paths.WebViewRuntimeInstaller))
                {
                    message += Environment.NewLine + "可尝试运行本地安装包：" + _paths.WebViewRuntimeInstaller;
                }

                _browserStatusLabel.Text = "初始化失败：" + ex.Message;
                MessageBox.Show(this, message, "浏览器初始化失败", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private async void BrowserInitializationCompleted(object sender, CoreWebView2InitializationCompletedEventArgs e)
        {
            if (!e.IsSuccess)
            {
                _browserStatusLabel.Text = "浏览器初始化失败。";
                string message = e.InitializationException == null ? "未知错误。" : e.InitializationException.ToString();
                if (File.Exists(_paths.WebViewRuntimeInstaller))
                {
                    message += Environment.NewLine + "可尝试运行本地安装包：" + _paths.WebViewRuntimeInstaller;
                }

                MessageBox.Show(this, message, "浏览器初始化失败", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            try
            {
                if (!_browserExtensionReady)
                {
                    await _browser.CoreWebView2.Profile.AddBrowserExtensionAsync(_paths.Plugin1688Folder);
                    _browserExtensionReady = true;
                }

                _browserStatusLabel.Text = "浏览器已就绪，1688 插件已挂载。";
                if (!string.IsNullOrEmpty(_browserUrlBox.Text))
                {
                    _browser.CoreWebView2.Navigate(_browserUrlBox.Text.Trim());
                }
            }
            catch (Exception ex)
            {
                _browserStatusLabel.Text = "插件挂载失败：" + ex.Message;
                MessageBox.Show(this, ex.ToString(), "插件挂载失败", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void NavigateBrowser(object sender, EventArgs e)
        {
            try
            {
                if (_browser.CoreWebView2 == null)
                {
                    _browserStatusLabel.Text = "请先初始化插件浏览器。";
                    return;
                }

                _browser.CoreWebView2.Navigate(_browserUrlBox.Text.Trim());
            }
            catch (Exception ex)
            {
                _browserStatusLabel.Text = ex.Message;
            }
        }

        private void EnsureSnapshot()
        {
            if (_snapshot == null)
            {
                _snapshot = new AssetSnapshot();
            }
        }

        private void ClearAssetViews()
        {
            EnsureSnapshot();
            _snapshot.Categories = new List<CategoryNode>();
            _snapshot.FeeRules = new List<FeeRule>();
            _categoryTree.Nodes.Clear();
            _feeGrid.DataSource = null;
        }

        private FeeRule GetSelectedFeeRule()
        {
            if (_feeGrid.CurrentRow == null)
            {
                return null;
            }

            FeeRuleDisplayRow display = _feeGrid.CurrentRow.DataBoundItem as FeeRuleDisplayRow;
            if (display != null)
            {
                return display.Rule;
            }

            return _feeGrid.CurrentRow.DataBoundItem as FeeRule;
        }

        private TreeNode BuildTreeNode(CategoryNode source)
        {
            string text = BuildBilingualCategoryName(source.DescriptionCategoryName);
            if (!string.IsNullOrEmpty(source.DescriptionTypeName))
            {
                text += " / " + BuildBilingualCategoryName(source.DescriptionTypeName);
            }

            if (!string.IsNullOrEmpty(source.DescriptionCategoryId))
            {
                text += " [" + source.DescriptionCategoryId + "]";
            }

            TreeNode node = new TreeNode(text);
            node.Tag = source;
            int i;
            for (i = 0; i < source.Children.Count; i++)
            {
                node.Nodes.Add(BuildTreeNode(source.Children[i]));
            }

            return node;
        }

        private TabPage CreateTabPage(string title)
        {
            TabPage tab = new TabPage(title);
            tab.BackColor = ShellBack;
            return tab;
        }

        private FlowLayoutPanel CreateActionBar()
        {
            FlowLayoutPanel panel = new FlowLayoutPanel();
            panel.Dock = DockStyle.Top;
            panel.Height = 68;
            panel.WrapContents = false;
            panel.AutoScroll = true;
            panel.Padding = new Padding(26, 14, 16, 0);
            panel.BackColor = ShellBack;
            return panel;
        }

        private Control CreateStatCard(string title, string value, string description, Color accent, out Label valueLabel)
        {
            RoundedPanel card = new RoundedPanel();
            card.Dock = DockStyle.Fill;
            card.Margin = new Padding(10);
            card.FillColor = CardBack;
            card.BorderColor = LineWarm;
            card.ShadowColor = Color.FromArgb(22, 116, 95, 62);
            card.Radius = 26;
            card.Padding = new Padding(16, 14, 16, 14);

            Panel line = new Panel();
            line.BackColor = accent;
            line.Width = 6;
            line.Dock = DockStyle.Left;
            line.Margin = new Padding(0, 12, 0, 12);

            Label titleLabel = new Label();
            titleLabel.Text = title;
            titleLabel.AutoSize = true;
            titleLabel.ForeColor = TextMuted;
            titleLabel.Location = new Point(18, 14);

            valueLabel = new Label();
            valueLabel.Text = value;
            valueLabel.AutoSize = true;
            valueLabel.Font = new Font("Microsoft YaHei UI", 23F, FontStyle.Bold, GraphicsUnit.Point, 134);
            valueLabel.ForeColor = TextStrong;
            valueLabel.Location = new Point(18, 36);

            Label descriptionLabel = new Label();
            descriptionLabel.Text = description;
            descriptionLabel.AutoSize = false;
            descriptionLabel.Width = 240;
            descriptionLabel.Height = 36;
            descriptionLabel.ForeColor = TextMuted;
            descriptionLabel.Location = new Point(18, 78);

            card.Controls.Add(descriptionLabel);
            card.Controls.Add(valueLabel);
            card.Controls.Add(titleLabel);
            card.Controls.Add(line);
            return card;
        }

        private static Control WrapWithGroup(string title, Control inner)
        {
            RoundedPanel group = new RoundedPanel();
            group.Dock = DockStyle.Fill;
            group.Padding = new Padding(16, 42, 16, 16);
            group.FillColor = CardBack;
            group.BorderColor = LineWarm;
            group.ShadowColor = Color.FromArgb(20, 95, 82, 58);
            group.Radius = 28;

            Label titleLabel = new Label();
            titleLabel.Text = title;
            titleLabel.AutoSize = true;
            titleLabel.Left = 16;
            titleLabel.Top = 12;
            titleLabel.ForeColor = TextStrong;
            titleLabel.Font = new Font("Microsoft YaHei UI", 9.5F, FontStyle.Bold, GraphicsUnit.Point, 134);

            Panel content = new Panel();
            content.Dock = DockStyle.Fill;
            content.BackColor = Color.FromArgb(255, 253, 248);
            inner.Font = new Font("Microsoft YaHei UI", 9F, FontStyle.Regular, GraphicsUnit.Point, 134);
            inner.Dock = DockStyle.Fill;
            content.Controls.Add(inner);
            group.Controls.Add(content);
            group.Controls.Add(titleLabel);
            return group;
        }

        private DataGridView CreateGrid()
        {
            DataGridView grid = new DataGridView();
            grid.ReadOnly = true;
            grid.AllowUserToAddRows = false;
            grid.AllowUserToDeleteRows = false;
            grid.AllowUserToResizeRows = false;
            grid.RowHeadersVisible = false;
            grid.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            grid.MultiSelect = false;
            grid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells;
            grid.BackgroundColor = Color.FromArgb(255, 253, 248);
            grid.BorderStyle = BorderStyle.None;
            grid.EnableHeadersVisualStyles = false;
            grid.RowTemplate.Height = 34;
            grid.GridColor = Color.FromArgb(231, 223, 207);
            grid.ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.Single;
            grid.CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal;
            grid.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(244, 239, 228);
            grid.ColumnHeadersDefaultCellStyle.ForeColor = Color.FromArgb(64, 73, 82);
            grid.ColumnHeadersDefaultCellStyle.Font = new Font("Microsoft YaHei UI", 9F, FontStyle.Bold, GraphicsUnit.Point, 134);
            grid.DefaultCellStyle.Font = new Font("Microsoft YaHei UI", 9F, FontStyle.Regular, GraphicsUnit.Point, 134);
            grid.DefaultCellStyle.BackColor = Color.FromArgb(255, 253, 248);
            grid.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(249, 246, 238);
            grid.DefaultCellStyle.SelectionBackColor = Color.FromArgb(219, 244, 235);
            grid.DefaultCellStyle.SelectionForeColor = TextStrong;
            return grid;
        }

        private Button CreateButton(string text, EventHandler handler, bool primary)
        {
            Button button = new Button();
            button.Text = text;
            button.AutoSize = false;
            button.Height = 38;
            button.Width = Math.Max(104, Math.Min(188, text.Length * 14 + 34));
            button.Margin = new Padding(5, 0, 7, 9);
            button.Padding = new Padding(12, 5, 12, 5);
            button.FlatStyle = FlatStyle.Flat;
            button.FlatAppearance.BorderSize = 1;
            button.Cursor = Cursors.Hand;
            Color normalBack = primary ? PilotGreen : Color.FromArgb(252, 250, 244);
            Color hoverBack = primary ? Color.FromArgb(19, 159, 116) : Color.FromArgb(232, 247, 240);
            Color downBack = primary ? PilotGreenDark : Color.FromArgb(212, 237, 225);
            Color normalBorder = primary ? PilotGreen : Color.FromArgb(222, 213, 195);
            Color hoverBorder = primary ? Color.FromArgb(76, 190, 150) : Color.FromArgb(161, 207, 187);
            button.FlatAppearance.BorderColor = normalBorder;
            button.BackColor = normalBack;
            button.ForeColor = primary ? Color.White : TextStrong;
            button.FlatAppearance.MouseOverBackColor = hoverBack;
            button.FlatAppearance.MouseDownBackColor = downBack;
            button.Font = new Font("Microsoft YaHei UI", 9F, FontStyle.Bold, GraphicsUnit.Point, 134);
            Padding normalMargin = button.Margin;
            Padding pressedMargin = new Padding(normalMargin.Left, normalMargin.Top + 2, normalMargin.Right, Math.Max(0, normalMargin.Bottom - 2));
            button.MouseEnter += delegate
            {
                button.BackColor = hoverBack;
                button.FlatAppearance.BorderColor = hoverBorder;
            };
            button.MouseLeave += delegate
            {
                button.BackColor = normalBack;
                button.FlatAppearance.BorderColor = normalBorder;
                button.Margin = normalMargin;
            };
            button.MouseDown += delegate
            {
                button.BackColor = downBack;
                button.Margin = pressedMargin;
            };
            button.MouseUp += delegate
            {
                button.BackColor = hoverBack;
                button.Margin = normalMargin;
            };
            button.Resize += delegate { SetRoundedRegion(button, 18); };
            SetRoundedRegion(button, 18);
            button.Click += handler;
            return button;
        }

        private Label CreateFormLabel(string text, int left, int top, int width)
        {
            Label label = new Label();
            label.Text = text;
            label.Left = left;
            label.Top = top + 6;
            label.Width = width;
            return label;
        }

        private TextBox CreateTextBox(int left, int top, int width, string value)
        {
            TextBox textBox = new TextBox();
            textBox.Left = left;
            textBox.Top = top;
            textBox.Width = width;
            textBox.Text = value;
            textBox.BorderStyle = BorderStyle.FixedSingle;
            textBox.BackColor = Color.FromArgb(255, 253, 248);
            return textBox;
        }

        private static void SetRoundedRegion(Control control, int radius)
        {
            if (control.Width <= 0 || control.Height <= 0)
            {
                return;
            }

            if (control.Region != null)
            {
                control.Region.Dispose();
            }

            control.Region = new Region(CreateRoundPath(new Rectangle(0, 0, control.Width, control.Height), radius));
        }

        private static GraphicsPath CreateRoundPath(Rectangle bounds, int radius)
        {
            GraphicsPath path = new GraphicsPath();
            int diameter = Math.Max(1, radius * 2);
            Rectangle arc = new Rectangle(bounds.Location, new Size(diameter, diameter));
            path.AddArc(arc, 180, 90);
            arc.X = bounds.Right - diameter;
            path.AddArc(arc, 270, 90);
            arc.Y = bounds.Bottom - diameter;
            path.AddArc(arc, 0, 90);
            arc.X = bounds.Left;
            path.AddArc(arc, 90, 90);
            path.CloseFigure();
            return path;
        }

        private decimal ParseDecimal(string text)
        {
            decimal value;
            return decimal.TryParse(text, out value) ? value : 0m;
        }

        private long ParseLong(string text)
        {
            long value;
            return long.TryParse(text, out value) ? value : 0L;
        }

        private void LoadUiLanguagePreference()
        {
            try
            {
                AppConfig config = ConfigService.Load(_paths.ConfigFile);
                _uiLanguage = NormalizeUiLanguage(config == null ? null : config.UiLanguage);
            }
            catch
            {
                _uiLanguage = "zh";
            }
        }

        private static string NormalizeUiLanguage(string language)
        {
            string text = Convert.ToString(language ?? string.Empty).Trim().ToLowerInvariant();
            if (text == "en" || text == "ru")
            {
                return text;
            }

            return "zh";
        }

        private int LanguageIndexFromCode(string language)
        {
            switch (NormalizeUiLanguage(language))
            {
                case "en":
                    return 1;
                case "ru":
                    return 2;
                default:
                    return 0;
            }
        }

        private string LanguageCodeFromIndex(int index)
        {
            switch (index)
            {
                case 1:
                    return "en";
                case 2:
                    return "ru";
                default:
                    return "zh";
            }
        }

        private Dictionary<string, string> CaptureUiState()
        {
            Dictionary<string, string> state = new Dictionary<string, string>();
            state["tab"] = _mainTabs == null ? "0" : _mainTabs.SelectedIndex.ToString();
            state["keywords"] = _autoKeywordsBox == null ? string.Empty : _autoKeywordsBox.Text;
            state["categoryId"] = _autoCategoryIdBox == null ? string.Empty : _autoCategoryIdBox.Text;
            state["typeId"] = _autoTypeIdBox == null ? string.Empty : _autoTypeIdBox.Text;
            state["clientId"] = _ozonClientIdBox == null ? string.Empty : _ozonClientIdBox.Text;
            state["apiKey"] = _ozonApiKeyBox == null ? string.Empty : _ozonApiKeyBox.Text;
            state["loop"] = _autoLoopCountBox == null ? "1" : _autoLoopCountBox.Text;
            state["browserUrl"] = _browserUrlBox == null ? string.Empty : _browserUrlBox.Text;
            return state;
        }

        private void RestoreUiState(Dictionary<string, string> state)
        {
            if (state == null)
            {
                return;
            }

            if (_autoKeywordsBox != null && state.ContainsKey("keywords"))
            {
                _autoKeywordsBox.Text = state["keywords"];
            }

            if (_autoCategoryIdBox != null && state.ContainsKey("categoryId"))
            {
                _autoCategoryIdBox.Text = state["categoryId"];
            }

            if (_autoTypeIdBox != null && state.ContainsKey("typeId"))
            {
                _autoTypeIdBox.Text = state["typeId"];
            }

            if (_ozonClientIdBox != null && state.ContainsKey("clientId") && !string.IsNullOrWhiteSpace(state["clientId"]))
            {
                _ozonClientIdBox.Text = state["clientId"];
            }

            if (_ozonApiKeyBox != null && state.ContainsKey("apiKey") && !string.IsNullOrWhiteSpace(state["apiKey"]))
            {
                _ozonApiKeyBox.Text = state["apiKey"];
            }

            if (_autoLoopCountBox != null && state.ContainsKey("loop"))
            {
                _autoLoopCountBox.Text = state["loop"];
            }

            if (_browserUrlBox != null && state.ContainsKey("browserUrl") && !string.IsNullOrWhiteSpace(state["browserUrl"]))
            {
                _browserUrlBox.Text = state["browserUrl"];
            }

            if (_languageComboBox != null)
            {
                _languageComboBox.SelectedIndex = LanguageIndexFromCode(_uiLanguage);
            }

            int selectedTab = 0;
            if (state.ContainsKey("tab"))
            {
                int.TryParse(state["tab"], out selectedTab);
            }

            if (_mainTabs != null && selectedTab >= 0 && selectedTab < _mainTabs.TabPages.Count)
            {
                _mainTabs.SelectedIndex = selectedTab;
            }
        }

        private void ApplyLanguageSelection(object sender, EventArgs e)
        {
            string selected = LanguageCodeFromIndex(_languageComboBox == null ? 0 : _languageComboBox.SelectedIndex);
            if (selected == _uiLanguage)
            {
                SetStatus(T("languageNoChange"));
                return;
            }

            Dictionary<string, string> state = CaptureUiState();
            _uiLanguage = selected;

            AppConfig config = _snapshot == null ? ConfigService.Load(_paths.ConfigFile) : _snapshot.Config;
            if (config == null)
            {
                config = AppConfig.CreateDefault();
            }

            config.UiLanguage = _uiLanguage;
            ConfigService.Save(_paths.ConfigFile, config);

            if (_browser != null)
            {
                try
                {
                    _browser.Dispose();
                }
                catch
                {
                }
            }

            SuspendLayout();
            Controls.Clear();
            InitializeControls();
            LoadAll();
            ApplyOzonSellerDefaults();
            RestoreUiState(state);
            ResumeLayout();
            PerformLayout();
            SetStatus(T("languageApplied"));
        }

        private string T(string key)
        {
            switch (NormalizeUiLanguage(_uiLanguage))
            {
                case "en":
                    switch (key)
                    {
                        case "subtitle": return "1688 sourcing to Ozon upload workspace";
                        case "overview": return "Overview";
                        case "config": return "Config Center";
                        case "assets": return "Categories & Rules";
                        case "language": return "Language";
                        case "browser": return "Plugin Browser";
                        case "reload": return "Reload";
                        case "openBase": return "Open Assets";
                        case "openPlugin": return "Open 1688 Plugin";
                        case "runUpdater": return "Run Updater";
                        case "brake": return "Emergency Stop";
                        case "loopCount": return "Loop count";
                        case "fullAuto": return "Full Auto Loop";
                        case "categoryNodes": return "Category Nodes";
                        case "categoryNodesDesc": return "Total nodes loaded from category.txt";
                        case "feeRules": return "Fee Rules";
                        case "feeRulesDesc": return "Total rules loaded from fee.txt";
                        case "pluginFiles": return "Plugin Files";
                        case "pluginFilesDesc": return "Files detected in the 1688 plugin folder";
                        case "overviewDetails": return "Workspace Summary";
                        case "resourcePaths": return "Resource Paths";
                        case "workRoot": return "Workspace: ";
                        case "baselineRoot": return "Baseline: ";
                        case "configFile": return "Config file: ";
                        case "plugin1688": return "1688 plugin: ";
                        case "updaterPath": return "Updater: ";
                        case "recoveryStatus": return "Recovery Status";
                        case "categoryCountLine": return "Category nodes: ";
                        case "feeCountLine": return "Fee rules: ";
                        case "pluginCountLine": return "Plugin files: ";
                        case "assetLoadFailed": return "Category/rule status: load failed";
                        case "failureReason": return "Reason: ";
                        case "currentFilters": return "Current Filters";
                        case "saveDir": return "Save directory: ";
                        case "priceRange": return "Price range: ";
                        case "minProfit": return "Minimum profit: ";
                        case "defaultShipping": return "Default shipping: ";
                        case "enable1688": return "Enable 1688: ";
                        case "autoExport": return "Auto export: ";
                        case "cloudFilter": return "Cloud filter: ";
                        case "yes": return "Yes";
                        case "no": return "No";
                        case "quickStart": return "Quick Start";
                        case "quickGuide": return "1. Open Plugin Browser and sign in to 1688.\r\n2. Double-click a second-level category in Categories & Rules.\r\n3. Run Auto Sourcing. The browser will search 1688 with the Chinese keyword.\r\n4. Upload to Ozon and watch the SKU and stock brief on Overview.\r\n5. Keywords always stay in Chinese and are not translated.";
                        case "reloadConfig": return "Reload Config";
                        case "saveConfig": return "Save Config";
                        case "initBrowser": return "Initialize Browser";
                        case "openUrl": return "Open URL";
                        case "filterConfig": return "Filter Config";
                        case "reloadAssets": return "Reload Assets";
                        case "exportFeeRules": return "Export Fee Rules";
                        case "useRuleForAuto": return "Send Rule to Auto";
                        case "searchAssets": return "Search categories/rules";
                        case "categoryTree": return "Category Tree";
                        case "feeRuleTable": return "Fee Rule Table";
                        case "languageTitle": return "Interface Language";
                        case "languageDesc": return "Switch the interface language here. The page text updates immediately after applying the new language.";
                        case "languageCurrent": return "Current language";
                        case "applyLanguage": return "Apply Language";
                        case "languageNote": return "Keywords in Auto Sourcing always stay Chinese. This switch only changes the UI language.";
                        case "languageApplied": return "Interface language updated.";
                        case "languageNoChange": return "Interface language is already active.";
                        case "ruleAppliedToAuto": return "Selected rule was sent to Auto Sourcing.";
                        default: return key;
                    }
                case "ru":
                    switch (key)
                    {
                        case "subtitle": return "Рабочее место для отбора 1688 и загрузки в Ozon";
                        case "overview": return "Обзор";
                        case "config": return "Настройки";
                        case "assets": return "Категории и правила";
                        case "language": return "Язык";
                        case "browser": return "Браузер плагина";
                        case "reload": return "Перезагрузить";
                        case "openBase": return "Открыть ресурсы";
                        case "openPlugin": return "Открыть плагин 1688";
                        case "runUpdater": return "Запустить обновление";
                        case "brake": return "Стоп";
                        case "loopCount": return "Циклов";
                        case "fullAuto": return "Полный автоцикл";
                        case "categoryNodes": return "Категории";
                        case "categoryNodesDesc": return "Всего узлов из category.txt";
                        case "feeRules": return "Правила доставки";
                        case "feeRulesDesc": return "Всего правил из fee.txt";
                        case "pluginFiles": return "Файлы плагина";
                        case "pluginFilesDesc": return "Файлы в папке плагина 1688";
                        case "overviewDetails": return "Сводка рабочего места";
                        case "resourcePaths": return "Пути ресурсов";
                        case "workRoot": return "Рабочая папка: ";
                        case "baselineRoot": return "Базовая папка: ";
                        case "configFile": return "Файл настроек: ";
                        case "plugin1688": return "Плагин 1688: ";
                        case "updaterPath": return "Обновление: ";
                        case "recoveryStatus": return "Статус восстановления";
                        case "categoryCountLine": return "Категории: ";
                        case "feeCountLine": return "Правила доставки: ";
                        case "pluginCountLine": return "Файлы плагина: ";
                        case "assetLoadFailed": return "Статус категорий/правил: ошибка загрузки";
                        case "failureReason": return "Причина: ";
                        case "currentFilters": return "Текущие фильтры";
                        case "saveDir": return "Папка сохранения: ";
                        case "priceRange": return "Диапазон цены: ";
                        case "minProfit": return "Минимальная прибыль: ";
                        case "defaultShipping": return "Доставка по умолчанию: ";
                        case "enable1688": return "1688 включен: ";
                        case "autoExport": return "Автоэкспорт: ";
                        case "cloudFilter": return "Облачный фильтр: ";
                        case "yes": return "Да";
                        case "no": return "Нет";
                        case "quickStart": return "Быстрый старт";
                        case "quickGuide": return "1. Откройте браузер плагина и войдите в 1688.\r\n2. Дважды щёлкните подкатегорию в разделе категорий.\r\n3. Запустите Auto Sourcing. Браузер будет искать по китайскому ключевому слову.\r\n4. Загрузите товары в Ozon и следите за брифом SKU и остатков на обзоре.\r\n5. Ключевые слова в Auto Sourcing всегда остаются китайскими.";
                        case "reloadConfig": return "Обновить настройки";
                        case "saveConfig": return "Сохранить настройки";
                        case "initBrowser": return "Запустить браузер";
                        case "openUrl": return "Открыть сайт";
                        case "filterConfig": return "Параметры фильтра";
                        case "reloadAssets": return "Обновить ресурсы";
                        case "exportFeeRules": return "Экспорт правил";
                        case "useRuleForAuto": return "Передать правило в Auto";
                        case "searchAssets": return "Поиск категорий/правил";
                        case "categoryTree": return "Дерево категорий";
                        case "feeRuleTable": return "Таблица правил";
                        case "languageTitle": return "Язык интерфейса";
                        case "languageDesc": return "Здесь можно переключать язык интерфейса. После применения текст страницы обновляется сразу.";
                        case "languageCurrent": return "Текущий язык";
                        case "applyLanguage": return "Применить язык";
                        case "languageNote": return "Ключевые слова в Auto Sourcing всегда остаются на китайском. Переключатель меняет только интерфейс.";
                        case "languageApplied": return "Язык интерфейса обновлён.";
                        case "languageNoChange": return "Этот язык уже активен.";
                        case "ruleAppliedToAuto": return "Выбранное правило передано в Auto Sourcing.";
                        default: return key;
                    }
                default:
                    switch (key)
                    {
                        case "subtitle": return "1688 自动选品到 Ozon 上架工作台";
                        case "overview": return "总览";
                        case "config": return "配置中心";
                        case "assets": return "类目与规则";
                        case "language": return "语言切换";
                        case "browser": return "插件浏览器";
                        case "reload": return "重新载入全部";
                        case "openBase": return "打开基础资源";
                        case "openPlugin": return "打开 1688 插件";
                        case "runUpdater": return "运行更新器";
                        case "brake": return "紧急刹车";
                        case "loopCount": return "全自动循环次数";
                        case "fullAuto": return "全链路自动循环";
                        case "categoryNodes": return "类目节点";
                        case "categoryNodesDesc": return "从 category.txt 读取到的类目树节点总数";
                        case "feeRules": return "运费规则";
                        case "feeRulesDesc": return "从 fee.txt 读取到的运费规则总数";
                        case "pluginFiles": return "插件文件";
                        case "pluginFilesDesc": return "1688 插件目录中的文件总数";
                        case "overviewDetails": return "恢复资产明细";
                        case "resourcePaths": return "资源路径";
                        case "workRoot": return "工作区目录：";
                        case "baselineRoot": return "基础目录：";
                        case "configFile": return "配置文件：";
                        case "plugin1688": return "1688 插件：";
                        case "updaterPath": return "更新器程序：";
                        case "recoveryStatus": return "恢复状态";
                        case "categoryCountLine": return "类目节点：";
                        case "feeCountLine": return "运费规则：";
                        case "pluginCountLine": return "插件文件：";
                        case "assetLoadFailed": return "类目/规则状态：加载失败";
                        case "failureReason": return "失败原因：";
                        case "currentFilters": return "当前筛选配置";
                        case "saveDir": return "保存目录：";
                        case "priceRange": return "售价范围：";
                        case "minProfit": return "最低利润率：";
                        case "defaultShipping": return "默认运费：";
                        case "enable1688": return "启用 1688：";
                        case "autoExport": return "自动导出：";
                        case "cloudFilter": return "云筛选：";
                        case "yes": return "是";
                        case "no": return "否";
                        case "quickStart": return "上手说明";
                        case "quickGuide": return "1. 打开插件浏览器，先登录 1688，保持浏览器会话在线。\r\n2. 在类目与规则里双击一个二级类目，系统会把类目 ID、类型 ID 和中文关键词带入 Auto Sourcing。\r\n3. 在 Auto Sourcing 点 Run，程序会跳到插件浏览器搜索 1688，抓取商品、详情、价格、图片和属性。\r\n4. 点 Upload 或在总览页跑全链路循环，系统会生成俄文标题和文案，再按 Ozon Seller API 上传。\r\n5. SKU 创建和库存写入过程会实时回写到总览简报里，关键词始终保持中文。";
                        case "reloadConfig": return "重新读取配置";
                        case "saveConfig": return "保存当前配置";
                        case "initBrowser": return "初始化插件浏览器";
                        case "openUrl": return "打开网址";
                        case "filterConfig": return "筛选配置";
                        case "reloadAssets": return "重新读取资源";
                        case "exportFeeRules": return "导出运费规则";
                        case "useRuleForAuto": return "把规则带入 Auto";
                        case "searchAssets": return "搜索类目/规则";
                        case "categoryTree": return "类目树";
                        case "feeRuleTable": return "运费规则表";
                        case "languageTitle": return "界面语言";
                        case "languageDesc": return "这里可以切换中文、英文、俄文三套界面。应用后会立即刷新当前窗口。";
                        case "languageCurrent": return "当前界面语言";
                        case "applyLanguage": return "应用语言";
                        case "languageNote": return "注意：Auto Sourcing 里的关键词始终保持中文，不会被语言切换改成英文或俄文。";
                        case "languageApplied": return "界面语言已切换。";
                        case "languageNoChange": return "当前已经是这个界面语言。";
                        case "ruleAppliedToAuto": return "已把选中规则带入 Auto Sourcing。";
                        default: return key;
                    }
            }
        }

        private string SafeValue(string text)
        {
            return string.IsNullOrEmpty(text) ? "未找到" : text;
        }

        private string YesNo(bool value)
        {
            return value ? T("yes") : T("no");
        }

        private void SetStatus(string message)
        {
            _statusLabel.Text = message;
        }
    }
}



