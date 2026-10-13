// Main window: pick the NSF and the PST, run export (Notes) then import (PST) as child processes of this exe,
// show progress and the verification summary (numbers only, written to <work>\summary.txt).
// Three views in one window: Setup -> Converting -> Done. Traditional Chinese / English, light / dark theme.
using System.Diagnostics;
using System.Drawing.Drawing2D;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;

namespace Nsf2Pst;

internal sealed class MainForm : Form
{
    // ---- language / theme (chosen before the controls are created; switching rebuilds the window)
    static bool En;
    static string L(string zh, string en) => En ? en : zh;

    /// <summary>Sets language and theme from the saved settings (or Windows) before a window is created.</summary>
    public static void Prepare()
    {
        Prefs.Load();
        En = Prefs.Lang == "en" || (Prefs.Lang == "" && !CultureInfo.CurrentUICulture.Name.StartsWith("zh", StringComparison.OrdinalIgnoreCase));
        bool dark = Prefs.Theme == "dark" || (Prefs.Theme == "" && Native.SystemUsesDarkTheme());
        Th.Apply(dark);
    }

    /// <summary>Set when the window closes only to rebuild itself in another language or theme.</summary>
    public bool RestartRequested { get; private set; }
    static (string nsf, string pst, string server, string user, string id, string pwd, string dir, bool anon, bool adv)? carry;

    // ---- controls used by the logic (names = UI Automation ids for the automated tests)
    readonly TextBox txtNsf = new(), txtPst = new(), txtServer = new(), txtUser = new(), txtLog = new(), txtId = new(), txtDir = new(),
                     txtPwd = new() { UseSystemPasswordChar = true };
    readonly CheckBox chkAnon = new() { Text = L("摘要隱藏資料夾名稱（方便對外提供）", "Hide folder names in the summary (safe to share)"), AutoSize = true };
    readonly Pill btnStart = new(L("開始轉換", "Start conversion"), primary: true, icon: Ico.Forward, iconAfter: true), btnStop = new(L("停止", "Stop"), icon: Ico.Stop);
    readonly Pill btnOpenOut = new(L("開啟 PST 位置", "Show PST in folder"), primary: true, icon: Ico.Folder), btnOpenLog = new(L("逐封記錄", "Message log"), icon: Ico.Document);
    readonly Pill btnSummary = new(L("摘要檔", "Summary"), icon: Ico.List), btnAgain = new(L("轉換另一個郵箱", "Convert another mailbox"), icon: Ico.Refresh);
    readonly Pill btnNsf = new(L("瀏覽檔案", "Browse files"), small: true), btnPst = new(L("更改", "Change"), small: true),
                  btnId = new(L("瀏覽", "Browse"), small: true), btnDir = new(L("瀏覽", "Browse"), small: true);
    readonly Pill btnLang = new(En ? "中文" : "English", small: true, icon: Ico.Globe), btnTheme = new("", small: true, icon: Th.Dark ? Ico.Sun : Ico.Moon);
    readonly Label lblStatus = new() { Text = L("請選擇 NSF 檔案。", "Choose an NSF file."), AutoSize = true };
    readonly Bar progress = new();
    Process? child;
    bool stopRequested;
    string workDir = "", statusBase = "";

    // ---- view parts
    readonly Panel host = new() { Dock = DockStyle.Fill, AutoScroll = true };
    TableLayoutPanel viewSetup = null!, viewRun = null!, viewDone = null!;
    readonly Stepper stepper = new();
    readonly Label lblPercent = new(), lblRunTitle = new(), lblRunFile = new(), lblElapsed = new();
    readonly Label lblDoneIcon = new(), lblDoneTitle = new(), lblDoneSub = new();
    readonly Card hero = new();
    readonly Dictionary<string, Label> tiles = new();
    readonly Dictionary<string, Card> tileCards = new();
    readonly DataGridView grid = new();
    readonly Card dropZone = new() { Dashed = true };
    readonly Panel dropEmpty = new(), dropFile = new();
    readonly Label lblFileName = new(), lblFileMeta = new(), lblEnv = new(), lblEnvWarn = new(), lblAdvSummary = new(), lblAdvChevron = new();
    readonly Card advCard = new();
    readonly Panel logPanel = new();
    readonly LinkLabel btnLog = new();
    readonly ToolTip tips = new();
    readonly System.Windows.Forms.Timer clock = new() { Interval = 1000 };
    DateTime runStart;
    List<(string name, int exp, int pst, int expU, int pstU, bool ok)> folderRows = new();
    bool envOk = true;

    static float k = 1;                       // DPI scale
    static int S(int px) => (int)Math.Round(px * k);

    public MainForm()
    {
        Text = L("Nsf2Pst - Notes 郵箱 (NSF) 轉 Outlook PST", "Nsf2Pst - Notes mailbox (NSF) to Outlook PST");
        Font = Th.Ui(10f);
        AutoScaleMode = AutoScaleMode.None;
        k = DeviceDpi / 96f;
        BackColor = Th.Back;
        try { Icon = Icon.ExtractAssociatedIcon(Environment.ProcessPath!); } catch { }
        var work = Screen.FromPoint(Cursor.Position).WorkingArea;
        MinimumSize = new Size(Math.Min(S(780), work.Width), Math.Min(S(640), work.Height));
        Size = new Size(Math.Min(S(980), work.Width), Math.Min(S(880), work.Height));
        StartPosition = FormStartPosition.CenterScreen;
        AllowDrop = true;
        KeyPreview = true;

        foreach (var b in new[] { btnStart, btnStop, btnOpenOut, btnOpenLog, btnSummary, btnAgain, btnNsf, btnPst, btnId, btnDir, btnLang, btnTheme }) b.Fit();
        BuildSetup(); BuildRun(); BuildDone();
        host.BackColor = Th.Back;
        host.Resize += (_, _) => { int side = Math.Max(S(28), (host.ClientSize.Width - S(820)) / 2); host.Padding = new Padding(side, S(26), side, S(26)); };
        ShowView(viewSetup);

        // log panel (hidden until asked for)
        txtLog.Multiline = true; txtLog.ReadOnly = true; txtLog.ScrollBars = ScrollBars.Both; txtLog.WordWrap = false;
        txtLog.Font = new Font("Consolas", 9.25f); txtLog.Dock = DockStyle.Fill; txtLog.BorderStyle = BorderStyle.None;
        txtLog.BackColor = Th.LogBack; txtLog.ForeColor = Th.LogText;
        logPanel.Dock = DockStyle.Bottom; logPanel.Height = S(210); logPanel.BackColor = Th.LogBack; logPanel.Padding = new Padding(S(16), S(10), S(8), S(8));
        logPanel.Controls.Add(txtLog); logPanel.Visible = false;

        Controls.Add(host);
        Controls.Add(logPanel);
        Controls.Add(BuildFooter());
        Controls.Add(BuildHeader());

        txtNsf.Name = "txtNsf"; txtPst.Name = "txtPst"; txtServer.Name = "txtServer"; txtUser.Name = "txtUser"; txtLog.Name = "txtLog"; txtId.Name = "txtId"; txtPwd.Name = "txtPwd"; txtDir.Name = "txtDir";
        chkAnon.Name = "chkAnon"; btnStart.Name = "btnStart"; btnStop.Name = "btnStop"; lblStatus.Name = "lblStatus"; btnLog.Name = "btnLog";
        btnOpenOut.Name = "btnOpenOut"; btnOpenLog.Name = "btnOpenLog"; btnAgain.Name = "btnAgain"; btnLang.Name = "btnLang"; btnTheme.Name = "btnTheme";

        btnNsf.Click += (_, _) => PickNsf();
        btnPst.Click += (_, _) => PickPst();
        btnDir.Click += (_, _) => { using var d = new OpenFileDialog { Filter = L("Domino 通訊錄 (*.nsf)|*.nsf|所有檔案|*.*", "Domino directory (*.nsf)|*.nsf|All files|*.*") }; if (d.ShowDialog(this) == DialogResult.OK) txtDir.Text = d.FileName; };
        btnId.Click += (_, _) => { using var d = new OpenFileDialog { Filter = L("Notes ID (*.id)|*.id|所有檔案|*.*", "Notes ID (*.id)|*.id|All files|*.*") }; if (d.ShowDialog(this) == DialogResult.OK) txtId.Text = d.FileName; };
        btnStart.Click += async (_, _) => await RunAsync();
        btnStop.Click += (_, _) => StopChild();
        btnOpenOut.Click += (_, _) => { if (File.Exists(txtPst.Text)) Process.Start("explorer.exe", $"/select,\"{txtPst.Text}\""); };
        btnOpenLog.Click += (_, _) => OpenFile(Path.Combine(workDir, "export_log.csv"));
        btnSummary.Click += (_, _) => OpenFile(Path.Combine(workDir, "summary.txt"));
        btnAgain.Click += (_, _) => { txtNsf.Text = ""; txtPst.Text = ""; ShowView(viewSetup); SetStatus(L("請選擇 NSF 檔案。", "Choose an NSF file.")); Taskbar.Set(Handle, 0, 0); };
        btnLang.Click += (_, _) => Rebuild(lang: En ? "zh" : "en");
        btnTheme.Click += (_, _) => Rebuild(theme: Th.Dark ? "light" : "dark");
        tips.SetToolTip(btnTheme, L("切換深色 / 淺色", "Switch dark / light"));
        tips.SetToolTip(btnLang, L("Switch to English", "切換到中文"));
        txtNsf.TextChanged += (_, _) => NsfChanged();
        foreach (var t in new Control[] { txtServer, txtDir, txtUser, txtId, txtPwd }) t.TextChanged += (_, _) => UpdateAdvSummary();
        chkAnon.CheckedChanged += (_, _) => UpdateAdvSummary();
        clock.Tick += (_, _) => UpdateElapsed();
        DragEnter += OnDragEnter; DragDrop += OnDragDrop; DragLeave += (_, _) => DropHighlight(false);
        KeyDown += (_, e) => { if (e.Control && e.KeyCode == Keys.O && child == null) { e.SuppressKeyPress = true; PickNsf(); } };
        AcceptButton = btnStart;
        FormClosing += (_, e) =>
        {
            if (child != null && MessageBox.Show(this, L("轉換仲進行緊，確定要停止並關閉？", "A conversion is still running. Stop it and close?"), Text, MessageBoxButtons.YesNo) == DialogResult.No) e.Cancel = true;
            else StopChild();
        };
        Shown += (_, _) => CheckPrerequisites();
        HandleCreated += (_, _) =>
        {
            Native.StyleTitleBar(Handle, Th.Surface, Th.Line, Th.Dark);
            Native.DarkScrollbars(txtLog.Handle, true);
            if (Th.Dark) { Native.DarkScrollbars(host.Handle, true); Native.DarkScrollbars(grid.Handle, true); }
        };

        if (carry is { } c)
        {
            txtNsf.Text = c.nsf; txtPst.Text = c.pst; txtServer.Text = c.server; txtUser.Text = c.user; txtId.Text = c.id; txtPwd.Text = c.pwd; txtDir.Text = c.dir; chkAnon.Checked = c.anon;
            if (c.adv) ToggleAdvanced();
            carry = null;
        }
        NsfChanged(); UpdateAdvSummary();
    }

    /// <summary>Saves the choice and closes; Program opens a new window with the same field values.</summary>
    void Rebuild(string? lang = null, string? theme = null)
    {
        if (child != null) return;
        if (lang != null) Prefs.Lang = lang;
        if (theme != null) Prefs.Theme = theme;
        Prefs.Save();
        carry = (txtNsf.Text, txtPst.Text, txtServer.Text, txtUser.Text, txtId.Text, txtPwd.Text, txtDir.Text, chkAnon.Checked, advCard.Visible);
        RestartRequested = true;
        Close();
    }

    // =====================================================================================================
    // layout
    // =====================================================================================================
    Control BuildHeader()
    {
        var bar = new Panel { Dock = DockStyle.Top, Height = S(68), BackColor = Th.Surface };
        bar.Paint += (_, e) => { using var p = new Pen(Th.Line); e.Graphics.DrawLine(p, 0, bar.Height - 1, bar.Width, bar.Height - 1); };
        var logo = new IconBadge(Ico.Mail) { Size = new Size(S(38), S(38)), Location = new Point(S(24), S(15)) };
        var title = new Label { Text = "Nsf2Pst", Font = Th.Num(15f), ForeColor = Th.Text, AutoSize = true, Location = new Point(S(74), S(12)), BackColor = Th.Surface };
        var sub = new Label { Text = L("Notes 郵箱 (NSF) 轉 Outlook PST", "Notes mailbox (NSF) to Outlook PST"), Font = Th.Ui(9f), ForeColor = Th.Muted, AutoSize = true, Location = new Point(S(76), S(40)), BackColor = Th.Surface };
        lblEnv.AutoSize = true; lblEnv.Font = Th.Ui(9f); lblEnv.ForeColor = Th.Muted; lblEnv.BackColor = Th.Surface; lblEnv.Text = L("檢查環境中 ...", "Checking environment ...");
        bar.Controls.AddRange(new Control[] { logo, title, sub, lblEnv, btnLang, btnTheme });
        void Place()
        {
            int x = bar.Width - S(24) - lblEnv.Width;
            lblEnv.Location = new Point(x, (bar.Height - lblEnv.Height) / 2);
            x -= S(20) + btnLang.Width; btnLang.Location = new Point(x, (bar.Height - btnLang.Height) / 2);
            x -= S(8) + btnTheme.Width; btnTheme.Location = new Point(x, (bar.Height - btnTheme.Height) / 2);
        }
        bar.Resize += (_, _) => Place();
        lblEnv.SizeChanged += (_, _) => Place();
        return bar;
    }

    Control BuildFooter()
    {
        var bar = new Panel { Dock = DockStyle.Bottom, Height = S(40), BackColor = Th.Surface, Padding = new Padding(S(24), 0, S(24), 0) };
        bar.Paint += (_, e) => { using var p = new Pen(Th.Line); e.Graphics.DrawLine(p, 0, 0, bar.Width, 0); };
        lblStatus.Font = Th.Ui(9.5f, FontStyle.Bold); lblStatus.BackColor = Th.Surface; lblStatus.ForeColor = Th.Text; lblStatus.Location = new Point(S(24), S(11));
        btnLog.Text = L("顯示詳細記錄", "Show details"); btnLog.AutoSize = true; btnLog.Font = Th.Ui(9.5f); btnLog.LinkColor = Th.Accent; btnLog.ActiveLinkColor = Th.AccentDark;
        btnLog.LinkBehavior = LinkBehavior.HoverUnderline; btnLog.BackColor = Th.Surface;
        btnLog.LinkClicked += (_, _) => ToggleLog();
        bar.Controls.Add(lblStatus); bar.Controls.Add(btnLog);
        bar.Resize += (_, _) => btnLog.Location = new Point(bar.Width - btnLog.Width - S(24), S(11));
        btnLog.SizeChanged += (_, _) => btnLog.Location = new Point(bar.Width - btnLog.Width - S(24), S(11));
        return bar;
    }

    void ToggleLog()
    {
        logPanel.Visible = !logPanel.Visible;
        btnLog.Text = logPanel.Visible ? L("隱藏詳細記錄", "Hide details") : L("顯示詳細記錄", "Show details");
        if (logPanel.Visible) { txtLog.SelectionStart = txtLog.TextLength; txtLog.ScrollToCaret(); }
    }

    static TableLayoutPanel Column()
    {
        var t = new TableLayoutPanel { Dock = DockStyle.Top, AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, ColumnCount = 1, BackColor = Th.Back, Margin = new Padding(0) };
        t.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        return t;
    }

    static Label Heading(string text, float size = 17f) =>
        new() { Text = text, Font = Th.Ui(size, FontStyle.Bold), ForeColor = Th.Text, AutoSize = true, Margin = new Padding(0, 0, 0, S(2)), BackColor = Th.Back };

    static Label SubText(string text) =>
        new() { Text = text, Font = Th.Ui(9.75f), ForeColor = Th.Muted, AutoSize = true, Margin = new Padding(0, 0, 0, S(18)), BackColor = Th.Back };

    /// <summary>A rounded input with a borderless TextBox (and an optional button on the right).</summary>
    static Card Field(TextBox box, string placeholder, Control? right = null)
    {
        var c = new Card { Radius = 8, Height = S(44), Dock = DockStyle.Top, Padding = new Padding(S(12), S(11), S(right == null ? 12 : 6), S(6)), Margin = new Padding(0, 0, 0, S(12)) };
        box.BorderStyle = BorderStyle.None; box.BackColor = Th.Surface; box.ForeColor = Th.Text; box.Font = Th.Ui(10f);
        box.PlaceholderText = placeholder; box.Dock = DockStyle.Fill;
        if (right != null) { right.Dock = DockStyle.Right; c.Controls.Add(right); c.Padding = new Padding(S(12), S(6), S(6), S(6)); var inner = new Panel { Dock = DockStyle.Fill, Padding = new Padding(0, S(5), S(8), 0), BackColor = Th.Surface }; inner.Controls.Add(box); c.Controls.Add(inner); inner.BringToFront(); }
        else c.Controls.Add(box);
        return c;
    }

    static Label FieldLabel(string text) =>
        new() { Text = text, Font = Th.Ui(9.5f, FontStyle.Bold), ForeColor = Th.Text, AutoSize = true, Margin = new Padding(0, 0, 0, S(6)), BackColor = Th.Back };

    void BuildSetup()
    {
        viewSetup = Column();
        viewSetup.Controls.Add(Heading(L("選擇要轉換嘅郵箱", "Choose a mailbox to convert")));
        viewSetup.Controls.Add(SubText(L("郵件、資料夾、已讀/未讀、聯絡人、日曆同加密郵件，一次過轉入 Outlook PST。",
                                         "Mail, folders, read/unread state, contacts, calendar and encrypted mail, all into one Outlook PST.")));

        lblEnvWarn.AutoSize = false; lblEnvWarn.Dock = DockStyle.Top; lblEnvWarn.Height = S(40); lblEnvWarn.Visible = false;
        lblEnvWarn.BackColor = Th.BadSoft; lblEnvWarn.ForeColor = Th.Bad; lblEnvWarn.Font = Th.Ui(9.5f, FontStyle.Bold);
        lblEnvWarn.TextAlign = ContentAlignment.MiddleLeft; lblEnvWarn.Padding = new Padding(S(14), 0, 0, 0); lblEnvWarn.Margin = new Padding(0, 0, 0, S(14));
        viewSetup.Controls.Add(lblEnvWarn);

        // drop zone: empty state / chosen file state
        dropZone.Dock = DockStyle.Top; dropZone.Margin = new Padding(0, 0, 0, S(12)); dropZone.AllowDrop = true;
        dropZone.DragEnter += OnDragEnter; dropZone.DragDrop += OnDragDrop; dropZone.DragLeave += (_, _) => DropHighlight(false);
        dropEmpty.Dock = DockStyle.Fill; dropEmpty.BackColor = Th.Surface;
        var upIcon = new Label { Text = Ico.Upload, Font = Ico.Font(26f), ForeColor = Th.Accent, AutoSize = true, BackColor = Th.Surface };
        var upText = new Label { Text = L("將 .nsf 檔案拖放到呢度", "Drop an .nsf file here"), Font = Th.Ui(12.5f, FontStyle.Bold), ForeColor = Th.Text, AutoSize = true, BackColor = Th.Surface };
        var upHint = new Label { Text = L("建議使用郵箱副本 · 本機或者 Domino server 上嘅檔案都得", "Use a copy of the mailbox · local files and Domino server paths both work"), Font = Th.Ui(9.25f), ForeColor = Th.Muted, AutoSize = true, BackColor = Th.Surface };
        dropEmpty.Controls.AddRange(new Control[] { upIcon, upText, upHint, btnNsf });
        dropEmpty.Resize += (_, _) =>
        {
            int cx = dropEmpty.Width / 2, total = upIcon.Height + upText.Height + upHint.Height + btnNsf.Height + S(16), y = Math.Max(S(8), (dropEmpty.Height - total) / 2);
            upIcon.Location = new Point(cx - upIcon.Width / 2, y); y += upIcon.Height + S(4);
            upText.Location = new Point(cx - upText.Width / 2, y); y += upText.Height + S(2);
            upHint.Location = new Point(cx - upHint.Width / 2, y); y += upHint.Height + S(10);
            btnNsf.Location = new Point(cx - btnNsf.Width / 2, y);
        };
        dropFile.Dock = DockStyle.Fill; dropFile.BackColor = Th.Surface; dropFile.Visible = false;
        var fileIcon = new IconBadge(Ico.Mail) { Size = new Size(S(52), S(52)), Location = new Point(S(24), S(26)) };
        lblFileName.AutoSize = true; lblFileName.Font = Th.Ui(13f, FontStyle.Bold); lblFileName.ForeColor = Th.Text; lblFileName.BackColor = Th.Surface; lblFileName.Location = new Point(S(90), S(24));
        lblFileMeta.AutoSize = true; lblFileMeta.Font = Th.Ui(9.25f); lblFileMeta.ForeColor = Th.Muted; lblFileMeta.BackColor = Th.Surface; lblFileMeta.Location = new Point(S(91), S(54));
        var change = new LinkLabel { Text = L("更換檔案", "Choose another file"), AutoSize = true, Font = Th.Ui(9.5f), LinkColor = Th.Accent, ActiveLinkColor = Th.AccentDark, LinkBehavior = LinkBehavior.HoverUnderline, BackColor = Th.Surface, Location = new Point(S(91), S(78)) };
        change.LinkClicked += (_, _) => PickNsf();
        dropFile.Controls.AddRange(new Control[] { fileIcon, lblFileName, lblFileMeta, change });
        var pathBox = Field(txtNsf, L("或者直接貼上 NSF 路徑", "...or paste the NSF path"));
        pathBox.Dock = DockStyle.Bottom; pathBox.Margin = new Padding(0); pathBox.Height = S(40); pathBox.Fill = Th.Back; txtNsf.BackColor = Th.Back;
        var pathWrap = new Panel { Dock = DockStyle.Bottom, Height = S(52), Padding = new Padding(S(14), 0, S(14), S(12)), BackColor = Th.Surface };
        pathWrap.Controls.Add(pathBox);
        dropZone.Controls.Add(dropEmpty); dropZone.Controls.Add(dropFile); dropZone.Controls.Add(pathWrap);
        dropZone.Padding = new Padding(S(2));
        dropZone.Height = S(232);
        viewSetup.Controls.Add(dropZone);

        viewSetup.Controls.Add(FieldLabel(L("輸出 PST", "Output PST")));
        viewSetup.Controls.Add(Field(txtPst, L("新 PST 檔案位置（揀咗 NSF 會自動填寫）", "New PST file (filled in when you pick the NSF)"), btnPst));

        // advanced settings expander
        var advHead = new Card { Radius = 10, Height = S(50), Dock = DockStyle.Top, Margin = new Padding(0, S(2), 0, S(8)), Cursor = Cursors.Hand, Name = "btnAdvanced" };
        var gear = new Label { Text = Ico.Settings, Font = Ico.Font(13f), ForeColor = Th.Muted, AutoSize = true, BackColor = Th.Surface, Location = new Point(S(16), S(15)) };
        var advTitle = new Label { Text = L("進階設定", "Advanced settings"), Font = Th.Ui(10.5f, FontStyle.Bold), ForeColor = Th.Text, AutoSize = true, BackColor = Th.Surface, Location = new Point(S(46), S(14)) };
        lblAdvSummary.AutoSize = true; lblAdvSummary.Font = Th.Ui(9.25f); lblAdvSummary.ForeColor = Th.Muted; lblAdvSummary.BackColor = Th.Surface;
        lblAdvChevron.Text = Ico.ChevronDown; lblAdvChevron.Font = Ico.Font(10f); lblAdvChevron.ForeColor = Th.Muted; lblAdvChevron.AutoSize = true; lblAdvChevron.BackColor = Th.Surface;
        advHead.Controls.AddRange(new Control[] { gear, advTitle, lblAdvSummary, lblAdvChevron });
        advHead.Resize += (_, _) => { lblAdvChevron.Location = new Point(advHead.Width - S(34), S(17)); lblAdvSummary.Location = new Point(advHead.Width - S(46) - lblAdvSummary.Width, S(16)); };
        lblAdvSummary.SizeChanged += (_, _) => lblAdvSummary.Location = new Point(advHead.Width - S(46) - lblAdvSummary.Width, S(16));
        foreach (Control c in new Control[] { advHead, gear, advTitle, lblAdvSummary, lblAdvChevron }) { c.Click += (_, _) => ToggleAdvanced(); c.Cursor = Cursors.Hand; }
        viewSetup.Controls.Add(advHead);

        var advGrid = Column(); advGrid.BackColor = Th.Surface; advGrid.Dock = DockStyle.Top;
        void Add(string label, string hint, TextBox box, string placeholder, Control? right = null)
        {
            var l = FieldLabel(label); l.BackColor = Th.Surface; advGrid.Controls.Add(l);
            var f = Field(box, placeholder, right); f.Fill = Th.Back; box.BackColor = Th.Back; foreach (Control c in f.Controls) if (c is Panel p) p.BackColor = Th.Back;
            f.Margin = new Padding(0, 0, 0, S(2)); advGrid.Controls.Add(f);
            advGrid.Controls.Add(new Label { Text = hint, Font = Th.Ui(8.75f), ForeColor = Th.Muted, AutoSize = true, BackColor = Th.Surface, Margin = new Padding(S(2), 0, 0, S(12)) });
        }
        string opt = L("選填", "Optional");
        Add("Domino server", L("NSF 喺 server 上先填，留空即係本機檔案。", "Only if the NSF is on a server; leave empty for a local file."), txtServer, L("例如 mail01/ExampleCorp", "e.g. mail01/ExampleCorp"));
        Add(L("Domino 通訊錄", "Domino directory"), L("用嚟將 Notes 名轉成 email 地址；可以填 server!!names.nsf 或者揀 names.nsf 副本。",
            "Turns Notes names into email addresses; enter server!!names.nsf or pick a copy of names.nsf."), txtDir, opt, btnDir);
        Add(L("已讀 / 未讀用戶", "Read / unread user"), L("留空會自動用郵箱主人嘅已讀記錄。", "Leave empty to use the mailbox owner's read marks."), txtUser, L("例如 CN=名/O=機構", "e.g. CN=Name/O=Org"));
        Add(L("郵箱主人 Notes ID", "Mailbox owner's Notes ID"), L("只有加密郵件先需要；留空用而家登入嘅 ID。", "Only needed for encrypted mail; leave empty to use the current ID."), txtId, opt, btnId);
        Add(L("ID 密碼", "ID password"), L("只會傳畀轉換程序，唔會寫入記錄。", "Passed only to the converter process, never written to a log."), txtPwd, opt);
        chkAnon.BackColor = Th.Surface; chkAnon.ForeColor = Th.Text; chkAnon.Font = Th.Ui(9.5f); chkAnon.Margin = new Padding(0, S(2), 0, S(4));
        advGrid.Controls.Add(chkAnon);
        advCard.Radius = 10; advCard.Dock = DockStyle.Top; advCard.AutoSize = true; advCard.AutoSizeMode = AutoSizeMode.GrowAndShrink;
        advCard.Padding = new Padding(S(20), S(16), S(20), S(12)); advCard.Margin = new Padding(0, 0, 0, S(14)); advCard.Visible = false;
        advCard.Controls.Add(advGrid);
        viewSetup.Controls.Add(advCard);

        btnStart.Height = S(52); btnStart.Dock = DockStyle.Top; btnStart.Margin = new Padding(0, S(10), 0, S(10)); btnStart.Font = Th.Ui(12f, FontStyle.Bold);
        viewSetup.Controls.Add(btnStart);
        viewSetup.Controls.Add(new Label
        {
            Text = L("轉換期間 Notes 會更新 NSF 入面嘅索引，所以建議用郵箱嘅副本。　Enter = 開始　·　Ctrl+O = 揀檔案",
                     "Notes updates indexes inside the NSF while converting, so work on a copy.    Enter = start  ·  Ctrl+O = open"),
            Font = Th.Ui(8.75f), ForeColor = Th.Muted, AutoSize = true, BackColor = Th.Back, Margin = new Padding(0, 0, 0, S(6)),
        });
    }

    void ToggleAdvanced()
    {
        advCard.Visible = !advCard.Visible;
        lblAdvChevron.Text = advCard.Visible ? Ico.ChevronUp : Ico.ChevronDown;
    }

    void BuildRun()
    {
        viewRun = Column();
        lblRunTitle.Text = L("轉換中", "Converting"); lblRunTitle.Font = Th.Ui(17f, FontStyle.Bold); lblRunTitle.ForeColor = Th.Text; lblRunTitle.AutoSize = true; lblRunTitle.BackColor = Th.Back;
        viewRun.Controls.Add(lblRunTitle);
        lblRunFile.Font = Th.Ui(9.75f); lblRunFile.ForeColor = Th.Muted; lblRunFile.AutoSize = true; lblRunFile.BackColor = Th.Back; lblRunFile.Margin = new Padding(0, 0, 0, S(18));
        viewRun.Controls.Add(lblRunFile);

        var card = new Card { Dock = DockStyle.Top, Height = S(264), Padding = new Padding(S(28), S(24), S(28), S(24)), Margin = new Padding(0, 0, 0, S(16)) };
        stepper.Dock = DockStyle.Top; stepper.Height = S(70);
        lblPercent.Text = "0%"; lblPercent.Font = Th.Num(40f); lblPercent.ForeColor = Th.Text; lblPercent.AutoSize = true; lblPercent.BackColor = Th.Surface;
        progress.Dock = DockStyle.Top; progress.Height = S(10);
        var detail = new Label { Name = "lblDetail", AutoSize = true, Font = Th.Ui(10f), ForeColor = Th.Muted, BackColor = Th.Surface };
        lblElapsed.AutoSize = true; lblElapsed.Font = Th.Ui(9.25f); lblElapsed.ForeColor = Th.Muted; lblElapsed.BackColor = Th.Surface;
        var mid = new Panel { Dock = DockStyle.Top, Height = S(92), BackColor = Th.Surface };
        mid.Controls.AddRange(new Control[] { lblPercent, lblElapsed });
        mid.Resize += (_, _) => { lblPercent.Location = new Point(0, S(14)); lblElapsed.Location = new Point(mid.Width - lblElapsed.Width, S(52)); };
        lblElapsed.SizeChanged += (_, _) => lblElapsed.Location = new Point(mid.Width - lblElapsed.Width, S(52));
        var below = new Panel { Dock = DockStyle.Top, Height = S(40), BackColor = Th.Surface, Padding = new Padding(0, S(12), 0, 0) };
        below.Controls.Add(detail); detail.Location = new Point(0, S(14));
        progress.ValueChanged += () =>
        {
            lblPercent.Text = $"{progress.Value / 10}%";
            if (child != null && IsHandleCreated) Taskbar.Set(Handle, progress.Value, 1000);
        };
        lblStatus.TextChanged += (_, _) => detail.Text = lblStatus.Text;
        card.Controls.Add(below); card.Controls.Add(progress); card.Controls.Add(mid); card.Controls.Add(stepper);
        viewRun.Controls.Add(card);

        var actions = new FlowLayoutPanel { AutoSize = true, Dock = DockStyle.Top, WrapContents = false, BackColor = Th.Back, Margin = new Padding(0) };
        actions.Controls.Add(btnStop);
        viewRun.Controls.Add(actions);
    }

    void BuildDone()
    {
        viewDone = Column();
        hero.Dock = DockStyle.Top; hero.Height = S(112); hero.Margin = new Padding(0, 0, 0, S(16));
        lblDoneIcon.AutoSize = true; lblDoneIcon.Font = Ico.Font(34f); lblDoneIcon.Location = new Point(S(28), S(26));
        lblDoneTitle.AutoSize = true; lblDoneTitle.Font = Th.Ui(16f, FontStyle.Bold); lblDoneTitle.ForeColor = Th.Text; lblDoneTitle.Location = new Point(S(96), S(24));
        lblDoneSub.AutoSize = true; lblDoneSub.Font = Th.Ui(9.75f); lblDoneSub.ForeColor = Th.Muted; lblDoneSub.Location = new Point(S(97), S(60));
        hero.Controls.AddRange(new Control[] { lblDoneIcon, lblDoneTitle, lblDoneSub });
        viewDone.Controls.Add(hero);

        var tileRow = new TableLayoutPanel { Dock = DockStyle.Top, Height = S(96), ColumnCount = 5, BackColor = Th.Back, Margin = new Padding(0, 0, 0, S(16)) };
        foreach (var (key, label, glyph) in new[] { ("mail", L("郵件", "Mail"), Ico.Mail), ("contact", L("聯絡人", "Contacts"), Ico.Contact), ("cal", L("日曆項目", "Calendar"), Ico.Calendar),
                                                   ("unread", L("未讀", "Unread"), Ico.Unread), ("fail", L("失敗", "Failed"), Ico.Warning) })
        {
            tileRow.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 20));
            var t = new Card { Dock = DockStyle.Fill, Margin = new Padding(key == "mail" ? 0 : S(6), 0, key == "fail" ? 0 : S(6), 0), Padding = new Padding(S(16), S(12), S(8), S(8)) };
            var ic = new Label { Text = glyph, Font = Ico.Font(11f), ForeColor = Th.Accent, AutoSize = true, BackColor = Th.Surface, Location = new Point(S(16), S(14)) };
            var cap = new Label { Text = label, Font = Th.Ui(9f), ForeColor = Th.Muted, AutoSize = true, BackColor = Th.Surface, Location = new Point(S(38), S(13)) };
            var v = new Label { Text = "–", Font = Th.Num(19f), ForeColor = Th.Text, AutoSize = true, BackColor = Th.Surface, Location = new Point(S(14), S(38)) };
            t.Controls.AddRange(new Control[] { ic, cap, v });
            tiles[key] = v; tileCards[key] = t;
            tileRow.Controls.Add(t);
        }
        viewDone.Controls.Add(tileRow);

        var tableCard = new Card { Dock = DockStyle.Top, Height = S(300), Padding = new Padding(S(20), S(16), S(20), S(16)), Margin = new Padding(0, 0, 0, S(16)), Name = "tableCard" };
        var tTitle = new Label { Text = L("資料夾核對", "Folder check"), Font = Th.Ui(11f, FontStyle.Bold), ForeColor = Th.Text, AutoSize = true, BackColor = Th.Surface, Dock = DockStyle.Top, Padding = new Padding(0, 0, 0, S(10)) };
        StyleGrid();
        tableCard.Controls.Add(grid); tableCard.Controls.Add(tTitle);
        viewDone.Controls.Add(tableCard);

        var actions = new FlowLayoutPanel { AutoSize = true, Dock = DockStyle.Top, WrapContents = true, BackColor = Th.Back, Margin = new Padding(0) };
        foreach (var b in new[] { btnOpenOut, btnOpenLog, btnSummary, btnAgain }) { b.Margin = new Padding(0, 0, S(10), S(10)); actions.Controls.Add(b); }
        viewDone.Controls.Add(actions);
    }

    void StyleGrid()
    {
        grid.Dock = DockStyle.Fill; grid.BorderStyle = BorderStyle.None; grid.BackgroundColor = Th.Surface; grid.GridColor = Th.Line;
        grid.CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal; grid.ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.None;
        grid.RowHeadersVisible = false; grid.AllowUserToAddRows = false; grid.AllowUserToDeleteRows = false; grid.AllowUserToResizeRows = false;
        grid.ReadOnly = true; grid.SelectionMode = DataGridViewSelectionMode.FullRowSelect; grid.MultiSelect = false;
        grid.EnableHeadersVisualStyles = false; grid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
        grid.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing; grid.ColumnHeadersHeight = S(34);
        grid.RowTemplate.Height = S(34);
        grid.ColumnHeadersDefaultCellStyle = new DataGridViewCellStyle { BackColor = Th.Surface, ForeColor = Th.Muted, Font = Th.Ui(9f, FontStyle.Bold), SelectionBackColor = Th.Surface, SelectionForeColor = Th.Muted, Padding = new Padding(S(4), 0, 0, 0) };
        grid.DefaultCellStyle = new DataGridViewCellStyle { BackColor = Th.Surface, ForeColor = Th.Text, Font = Th.Ui(9.5f), SelectionBackColor = Th.AccentSoft, SelectionForeColor = Th.Text, Padding = new Padding(S(4), 0, 0, 0) };
        grid.Columns.Add("folder", L("資料夾", "Folder")); grid.Columns.Add("exp", "NSF"); grid.Columns.Add("pst", "PST"); grid.Columns.Add("unread", L("未讀 NSF / PST", "Unread NSF / PST")); grid.Columns.Add("ok", L("核對", "Check"));
        grid.Columns["folder"]!.FillWeight = 46; grid.Columns["exp"]!.FillWeight = 12; grid.Columns["pst"]!.FillWeight = 12; grid.Columns["unread"]!.FillWeight = 18; grid.Columns["ok"]!.FillWeight = 12;
        foreach (var c in new[] { "exp", "pst", "unread" }) { grid.Columns[c]!.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight; grid.Columns[c]!.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleRight; }
        grid.Columns["ok"]!.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
        foreach (DataGridViewColumn col in grid.Columns) col.SortMode = DataGridViewColumnSortMode.NotSortable;
        grid.SelectionChanged += (_, _) => grid.ClearSelection();
        grid.Columns["ok"]!.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
        grid.Columns["ok"]!.DefaultCellStyle.Font = Ico.Font(10f);
    }

    void ShowView(Control view)
    {
        host.SuspendLayout();
        host.Controls.Clear();
        host.Controls.Add(view);
        host.ResumeLayout();
        host.AutoScrollPosition = new Point(0, 0);
        AcceptButton = view == viewSetup ? btnStart : null;
        btnLang.Enabled = btnTheme.Enabled = view != viewRun;
    }

    enum Tone { Normal, Ok, Warn, Bad }

    void SetStatus(string text, Tone tone = Tone.Normal)
    {
        statusBase = text;
        lblStatus.ForeColor = tone switch { Tone.Ok => Th.Ok, Tone.Warn => Th.Warn, Tone.Bad => Th.Bad, _ => Th.Text };
        lblStatus.Text = text;
    }

    // =====================================================================================================
    // setup interactions
    // =====================================================================================================
    void OnDragEnter(object? s, DragEventArgs e)
    {
        bool ok = e.Data?.GetData(DataFormats.FileDrop) is string[] f && f.Length > 0 && f[0].EndsWith(".nsf", StringComparison.OrdinalIgnoreCase) && child == null;
        e.Effect = ok ? DragDropEffects.Copy : DragDropEffects.None;
        DropHighlight(ok);
    }

    void OnDragDrop(object? s, DragEventArgs e)
    {
        DropHighlight(false);
        if (e.Data?.GetData(DataFormats.FileDrop) is string[] f && f.Length > 0) { if (host.Controls.Contains(viewDone)) ShowView(viewSetup); txtPst.Text = ""; txtNsf.Text = f[0]; }
    }

    void DropHighlight(bool on)
    {
        if (dropZone.Highlight == on) return;
        dropZone.Highlight = on; dropZone.Invalidate();
    }

    void NsfChanged()
    {
        string p = txtNsf.Text.Trim().Trim('"');
        bool exists = p.Length > 0 && File.Exists(p);
        if (exists)
        {
            var fi = new FileInfo(p);
            lblFileName.Text = fi.Name;
            lblFileMeta.Text = $"{SizeText(fi.Length)}　·　{L("修改於", "modified")} {fi.LastWriteTime:yyyy-MM-dd HH:mm}";
            if (txtPst.Text.Trim().Length == 0) txtPst.Text = Path.Combine(fi.DirectoryName!, Path.GetFileNameWithoutExtension(fi.Name) + ".pst");
            if (envOk) SetStatus(L("準備好，可以開始轉換。", "Ready to convert."));
        }
        else if (p.Length > 0) SetStatus(txtServer.Text.Trim().Length > 0 ? L("會喺 Domino server 開啟呢個郵箱。", "The mailbox will be opened on the Domino server.") : L("搵唔到呢個 NSF 檔案。", "NSF file not found."),
                                         txtServer.Text.Trim().Length > 0 ? Tone.Normal : Tone.Warn);
        dropEmpty.Visible = !exists; dropFile.Visible = exists;
        dropZone.Dashed = !exists; dropZone.Height = S(exists ? 168 : 232); dropZone.Invalidate();
        btnStart.Enabled = p.Length > 0 && child == null;
    }

    void UpdateAdvSummary()
    {
        var parts = new List<string>();
        if (txtServer.Text.Trim().Length > 0) parts.Add("Server");
        if (txtDir.Text.Trim().Length > 0) parts.Add(L("通訊錄", "Directory"));
        if (txtUser.Text.Trim().Length > 0) parts.Add(L("已讀用戶", "Read user"));
        if (txtId.Text.Trim().Length > 0) parts.Add("Notes ID");
        if (chkAnon.Checked) parts.Add(L("隱藏資料夾名", "Hidden folder names"));
        lblAdvSummary.Text = parts.Count == 0 ? L("全部選填 · 用預設值", "All optional · using defaults") : string.Join(" · ", parts);
        lblAdvSummary.ForeColor = parts.Count == 0 ? Th.Muted : Th.Accent;
    }

    static string SizeText(long b) => b >= 1L << 30 ? $"{b / (double)(1L << 30):N2} GB" : b >= 1L << 20 ? $"{b / (double)(1L << 20):N1} MB" : $"{b / 1024.0:N0} KB";

    static string Duration(TimeSpan t) => t.TotalHours >= 1 ? t.ToString(@"h\:mm\:ss") : t.ToString(@"m\:ss");

    string Elapsed() => Duration(DateTime.Now - runStart);

    /// <summary>Elapsed time plus a rough estimate of the remaining time once enough progress is known.</summary>
    void UpdateElapsed()
    {
        var used = DateTime.Now - runStart;
        string text = L("已用時間 ", "Elapsed ") + Duration(used);
        int v = progress.Value;
        if (v >= 30 && v < 1000 && used.TotalSeconds >= 5)
        {
            var left = TimeSpan.FromSeconds(used.TotalSeconds * (1000 - v) / v);
            text += L("　·　約剩 ", "  ·  about ") + Duration(left) + L("", " left");
        }
        lblElapsed.Text = text;
    }

    void CheckPrerequisites()
    {
        bool notes = Microsoft.Win32.Registry.GetValue(@"HKEY_CLASSES_ROOT\CLSID\{29131539-2EED-1069-BF5D-00DD011186B7}\InprocServer32", "", null) != null;
        bool outlook = new[] { @"%ProgramFiles%\Microsoft Office\root\Office16\OLMAPI32.DLL", @"%ProgramFiles%\Microsoft Office\Office16\OLMAPI32.DLL" }
            .Any(p => File.Exists(Environment.ExpandEnvironmentVariables(p)));
        envOk = notes && outlook;
        lblEnv.Text = $"{(notes ? "●" : "○")} Notes      {(outlook ? "●" : "○")} Outlook 64-bit";
        lblEnv.ForeColor = envOk ? Th.Ok : Th.Bad;
        Log(L($"環境檢查：Notes client {(notes ? "✓" : "✗ 搵唔到（需要 64-bit HCL Notes）")}，Outlook 64-bit {(outlook ? "✓" : "✗ 搵唔到（需要 64-bit Outlook）")}",
              $"Environment: Notes client {(notes ? "✓" : "✗ not found (64-bit HCL Notes required)")}, Outlook 64-bit {(outlook ? "✓" : "✗ not found (64-bit Outlook required)")}"));
        if (!envOk)
        {
            var missing = string.Join(L("、", ", "), new[] { notes ? null : "64-bit HCL Notes", outlook ? null : "64-bit Outlook" }.Where(x => x != null));
            lblEnvWarn.Text = L($"  缺少必要軟件：{missing}。未能轉換。", $"  Missing required software: {missing}. Cannot convert.");
            lblEnvWarn.Visible = true;
            SetStatus(L("缺少必要軟件，未能轉換。", "Missing required software; cannot convert."), Tone.Bad);
        }
    }

    void PickNsf()
    {
        using var d = new OpenFileDialog { Filter = L("Notes 資料庫 (*.nsf)|*.nsf|所有檔案|*.*", "Notes database (*.nsf)|*.nsf|All files|*.*"), Title = L("選擇 NSF（建議用副本）", "Choose an NSF (preferably a copy)") };
        if (d.ShowDialog(this) != DialogResult.OK) return;
        txtPst.Text = "";
        txtNsf.Text = d.FileName;
    }

    void PickPst()
    {
        using var d = new SaveFileDialog { Filter = L("Outlook 資料檔 (*.pst)|*.pst", "Outlook data file (*.pst)|*.pst"), Title = L("輸出 PST", "Output PST"), OverwritePrompt = false, FileName = Path.GetFileName(txtPst.Text) };
        if (d.ShowDialog(this) == DialogResult.OK) txtPst.Text = d.FileName;
    }

    // =====================================================================================================
    // conversion
    // =====================================================================================================
    async Task RunAsync()
    {
        string nsf = txtNsf.Text.Trim().Trim('"'), pst = txtPst.Text.Trim().Trim('"'), server = txtServer.Text.Trim(), user = txtUser.Text.Trim(), id = txtId.Text.Trim();
        if (id.Length > 0 && !File.Exists(id)) { MessageBox.Show(this, L("搵唔到 Notes ID 檔案。", "Notes ID file not found."), Text); return; }
        if (server.Length == 0 && !File.Exists(nsf)) { MessageBox.Show(this, L("搵唔到 NSF 檔案。", "NSF file not found."), Text); return; }
        if (nsf.Length == 0 || pst.Length == 0) { MessageBox.Show(this, L("請填寫 NSF 同輸出 PST。", "Enter the NSF and the output PST."), Text); return; }
        pst = Path.GetFullPath(pst);
        if (!pst.EndsWith(".pst", StringComparison.OrdinalIgnoreCase)) pst += ".pst";
        if (File.Exists(pst))
        {
            MessageBox.Show(this, L("輸出 PST 已經存在，請用新檔名（再寫入同一個 PST 會令信件重複）。",
                                    "The output PST already exists. Use a new file name (writing into the same PST again would duplicate the mail)."), Text);
            return;
        }
        txtPst.Text = pst;
        workDir = Path.Combine(Path.GetDirectoryName(pst)!, Path.GetFileNameWithoutExtension(pst) + "_work");
        Directory.CreateDirectory(workDir);

        SetRunning(true);
        txtLog.Clear();
        foreach (var t in tiles.Values) { t.Text = "–"; t.ForeColor = Th.Text; }
        lblRunTitle.Text = L("轉換中", "Converting");
        lblRunFile.Text = $"{Path.GetFileName(nsf)}  →  {Path.GetFileName(pst)}";
        stepper.Current = 0;
        runStart = DateTime.Now; UpdateElapsed(); clock.Start();
        ShowView(viewRun);
        Taskbar.Set(Handle, 0, 1000);
        bool finished = false, allOk = false, stopped = false;
        string failText = "";
        try
        {
            Log(L("開始：", "Start: ") + $"{nsf} → {pst}");
            var exportArgs = new List<string> { "export", nsf, workDir };
            if (server.Length > 0) exportArgs.AddRange(new[] { "--server", server });
            if (user.Length > 0) exportArgs.AddRange(new[] { "--user", user });
            if (id.Length > 0) exportArgs.AddRange(new[] { "--id", id });
            if (txtDir.Text.Trim().Length > 0) exportArgs.AddRange(new[] { "--directory", txtDir.Text.Trim() });
            SetStatus(L("第 1 步：由 Notes 匯出信件 ...", "Step 1: exporting mail from Notes ..."));
            int rc1 = await RunChild(exportArgs, 0, 700, password: id.Length > 0 ? txtPwd.Text : null);
            if (stopRequested) { SetStatus(L("已停止。", "Stopped."), Tone.Warn); stopped = true; return; }
            if (rc1 != 0 && rc1 != 1) { SetStatus(L("匯出失敗，請睇記錄。", "Export failed - see the log."), Tone.Bad); failText = L("匯出失敗", "export failed"); return; }   // 1 = some mails failed, still continue

            stepper.Current = 1;
            SetStatus(L("第 2 步：寫入 PST ...", "Step 2: writing the PST ..."));
            int rc2 = await RunChild(new List<string> { "import", pst, Path.Combine(workDir, "manifest.jsonl") }, 700, 950);
            if (stopRequested) { SetStatus(L("已停止。", "Stopped."), Tone.Warn); stopped = true; return; }
            if (rc2 != 0 && rc2 != 1) { SetStatus(L("寫入 PST 失敗，請睇記錄。", "Writing the PST failed - see the log."), Tone.Bad); failText = L("寫入 PST 失敗", "writing the PST failed"); return; }

            stepper.Current = 2;
            SetStatus(L("第 3 步：核對 ...", "Step 3: verifying ..."));
            var counts = new List<string>();
            await RunChild(new List<string> { "counts", pst }, 950, 1000, counts);
            var summary = BuildSummary(nsf, pst, counts, (DateTime.Now - runStart).TotalSeconds, chkAnon.Checked, out allOk);
            File.WriteAllLines(Path.Combine(workDir, "summary.txt"), summary, new UTF8Encoding(true));
            Log("");
            foreach (var l in summary) Log(l);
            UpdateTiles(summary);
            Log("");
            Log(L("摘要檔（只有數字，可以對外提供）：", "Summary file (numbers only, safe to share): ") + Path.Combine(workDir, "summary.txt"));
            progress.Value = 1000;
            stepper.Current = 3;
            if (allOk) SetStatus(L("完成 ✓ 所有 folder 數量核對正確。", "Done ✓ All folder counts match."), Tone.Ok);
            else SetStatus(L("完成，但有 folder 數量對唔上，請睇摘要。", "Done, but some folder counts do not match - see the summary."), Tone.Warn);
            finished = true;
        }
        catch (Exception ex)
        {
            Log(L("錯誤：", "Error: ") + ex.Message);
            SetStatus(L("出錯，請睇記錄。", "Error - see the log."), Tone.Bad);
            failText = L("出錯", "error");
        }
        finally
        {
            clock.Stop();
            SetRunning(false);
            ShowDone(finished, allOk, stopped, failText, pst);
        }
    }

    void ShowDone(bool finished, bool allOk, bool stopped, string failText, string pst)
    {
        Color tint;
        if (finished)
        {
            lblDoneIcon.Text = allOk ? Ico.Completed : Ico.Warning;
            lblDoneIcon.ForeColor = allOk ? Th.Ok : Th.Warn;
            tint = allOk ? Th.OkSoft : Th.WarnSoft;
            lblDoneTitle.Text = allOk ? L("轉換完成", "Conversion complete") : L("轉換完成，但有資料夾數量對唔上", "Converted, but some folder counts do not match");
            long size = File.Exists(pst) ? new FileInfo(pst).Length : 0;
            lblDoneSub.Text = $"{L("用時", "Took")} {Elapsed()}　·　PST {SizeText(size)}　·　{Path.GetFileName(pst)}";
        }
        else
        {
            lblDoneIcon.Text = stopped ? Ico.Stop : Ico.ErrorBadge;
            lblDoneIcon.ForeColor = stopped ? Th.Warn : Th.Bad;
            tint = stopped ? Th.WarnSoft : Th.BadSoft;
            lblDoneTitle.Text = stopped ? L("轉換已停止", "Conversion stopped") : L($"轉換未完成：{failText}", $"Conversion not finished: {failText}");
            lblDoneSub.Text = L("撳下面嘅「顯示詳細記錄」睇原因。", "Open \"Show details\" below to see why.");
        }
        hero.Fill = tint; hero.Border = Color.Empty; hero.Invalidate();
        foreach (Control c in hero.Controls) c.BackColor = tint;

        grid.Rows.Clear();
        foreach (var r in folderRows)
        {
            int i = grid.Rows.Add(r.name == "(default folders)" ? L("預設資料夾（收件匣、寄件備份 …）", "Default folders (Inbox, Sent ...)") : r.name,
                                  r.exp.ToString("N0"), r.pst.ToString("N0"), $"{r.expU:N0} / {r.pstU:N0}", r.ok ? Ico.CheckMark : Ico.Cancel);
            grid.Rows[i].Cells["ok"].Style.ForeColor = r.ok ? Th.Ok : Th.Bad;
        }
        var tableCard = viewDone.Controls["tableCard"]!;
        tableCard.Height = S(66) + grid.ColumnHeadersHeight + Math.Min(Math.Max(folderRows.Count, 1), 10) * grid.RowTemplate.Height;
        tableCard.Visible = finished;
        btnOpenOut.Enabled = btnOpenLog.Enabled = btnSummary.Enabled = finished;
        ShowView(viewDone);
        if (!finished && !logPanel.Visible) ToggleLog();

        // taskbar: full green bar on success, red / yellow otherwise; flash the button if the window is in the background
        if (finished && allOk) Taskbar.Clear(Handle); else Taskbar.Set(Handle, 1000, 1000, finished || stopped ? Taskbar.Paused : Taskbar.Error);
        if (Form.ActiveForm != this) Native.Flash(Handle);
    }

    /// <summary>Runs this exe with the given arguments, streams its output to the log and maps PROGRESS lines to the bar.</summary>
    Task<int> RunChild(List<string> args, int from, int to, List<string>? capture = null, string? password = null)
    {
        var psi = new ProcessStartInfo(Environment.ProcessPath!)
        {
            UseShellExecute = false, CreateNoWindow = true,
            RedirectStandardOutput = true, RedirectStandardError = true,
            StandardOutputEncoding = Encoding.UTF8, StandardErrorEncoding = Encoding.UTF8,
        };
        foreach (var a in args) psi.ArgumentList.Add(a);
        if (password != null) psi.Environment["NOTES_ID_PASSWORD"] = password;   // never on the command line
        var tcs = new TaskCompletionSource<int>();
        var p = new Process { StartInfo = psi, EnableRaisingEvents = true };
        void OnLine(string? line)
        {
            if (line == null) return;
            if (IsNoise(line)) return;
            BeginInvoke(() =>
            {
                if (line.StartsWith("PROGRESS "))
                {
                    var parts = line.Split(' ');
                    if (parts.Length == 3 && int.TryParse(parts[1], out int done) && int.TryParse(parts[2], out int total) && total > 0)
                    {
                        progress.Value = Math.Min(1000, from + (int)((long)(to - from) * done / total));
                        lblStatus.Text = $"{statusBase}{L("（", " (")}{done:N0} / {total:N0}{L("）", ")")}";
                    }
                }
                else if (capture != null) capture.Add(line);
                else Log("  " + line);
            });
        }
        p.OutputDataReceived += (_, e) => OnLine(e.Data);
        p.ErrorDataReceived += (_, e) => OnLine(e.Data);
        p.Exited += (_, _) => { p.WaitForExit(); BeginInvoke(() => { child = null; tcs.TrySetResult(p.ExitCode); }); };
        child = p;
        p.Start();
        p.BeginOutputReadLine();
        p.BeginErrorReadLine();
        return tcs.Task;
    }

    // Notes writes housekeeping messages to the console (consistency checks, certificate warnings ...)
    static bool IsNoise(string l) =>
        string.IsNullOrWhiteSpace(l) || l.Contains("consistency check") || l.Contains("ReorgFolder") || l.Contains("certificate")
        || l.StartsWith("details") || l.Contains("leaked") || l.Contains("Free block") || l.Contains("Leaked block") || l.Contains("CSRF") || l.Contains("iNotes");

    void StopChild()
    {
        stopRequested = true;
        try { child?.Kill(true); } catch { }
    }

    void SetRunning(bool running)
    {
        stopRequested = false;
        btnStart.Enabled = !running && txtNsf.Text.Trim().Length > 0;
        btnNsf.Enabled = btnPst.Enabled = btnId.Enabled = btnDir.Enabled = !running;
        txtNsf.ReadOnly = txtPst.ReadOnly = txtServer.ReadOnly = txtUser.ReadOnly = txtId.ReadOnly = txtPwd.ReadOnly = txtDir.ReadOnly = running;
        btnStop.Enabled = running;
        if (running) progress.Value = 0;
    }

    void Log(string s) { txtLog.AppendText(s + Environment.NewLine); }

    static void OpenFile(string path) { if (File.Exists(path)) Process.Start(new ProcessStartInfo(path) { UseShellExecute = true }); }

    /// <summary>Fills the result tiles from the summary line "documents N | mails exported N | contacts N | ...".</summary>
    void UpdateTiles(IEnumerable<string> summary)
    {
        string line = summary.FirstOrDefault(l => l.StartsWith("documents ")) ?? "";
        string Num(string key) { var m = Regex.Match(line, Regex.Escape(key) + @" (\d+)"); return m.Success ? int.Parse(m.Groups[1].Value).ToString("N0") : "–"; }
        tiles["mail"].Text = Num("mails exported");
        tiles["contact"].Text = Num("contacts");
        tiles["cal"].Text = int.TryParse(Num("calendar entries").Replace(",", ""), out int ce) && int.TryParse(Num("invitations in calendar").Replace(",", ""), out int ic) ? (ce + ic).ToString("N0") : Num("calendar entries");
        tiles["unread"].Text = Num("| unread");
        tiles["fail"].Text = Num("| failed");
        bool failed = tiles["fail"].Text is not ("0" or "–");
        tiles["fail"].ForeColor = failed ? Th.Bad : Th.Text;
        var fc = tileCards["fail"];
        fc.Fill = failed ? Th.BadSoft : Th.Surface; fc.Invalidate();
        foreach (Control c in fc.Controls) c.BackColor = fc.Fill;
    }

    /// <summary>Expected (manifest) vs actual (PST) per folder; numbers only.</summary>
    List<string> BuildSummary(string nsf, string pst, List<string> counts, double seconds, bool anonymize, out bool allOk)
    {
        var expected = new Dictionary<string, (int total, int unread)>();
        foreach (var line in File.ReadLines(Path.Combine(workDir, "manifest.jsonl"), Encoding.UTF8))
        {
            if (string.IsNullOrWhiteSpace(line)) continue;
            using var j = JsonDocument.Parse(line);
            string f = j.RootElement.GetProperty("folder").GetString() ?? "";
            if (f.StartsWith('@')) f = "(default folders)";
            bool unread = j.RootElement.TryGetProperty("read", out var r) && !r.GetBoolean();
            var e = expected.GetValueOrDefault(f);
            expected[f] = (e.total + 1, e.unread + (unread ? 1 : 0));
        }
        var actual = new Dictionary<string, (int total, int unread)>();
        foreach (var row in counts)
        {
            var c = row.Split('\t');
            if (c.Length < 3 || !int.TryParse(c[1], out int t) || !int.TryParse(c[2], out int u)) continue;
            string f = expected.ContainsKey(c[0]) ? c[0] : "(default folders)";
            var a = actual.GetValueOrDefault(f);
            actual[f] = (a.total + t, a.unread + u);
        }

        var log = ReadLog();
        int Count(Func<string[], bool> pred) => log.Count(pred);
        var outp = new List<string>
        {
            $"NSF -> PST summary  {DateTime.Now:yyyy-MM-dd HH:mm}  (numbers only)",
            $"time {seconds:N0}s | NSF size {(File.Exists(nsf) ? new FileInfo(nsf).Length : 0):N0} | PST size {new FileInfo(pst).Length:N0}",
            $"documents {log.Count} | mails exported {Count(l => l[3].StartsWith("exported"))} | contacts {Count(l => l[3] == "contact")} | calendar entries {Count(l => l[3].StartsWith("appointment"))} | invitations in calendar {Count(l => l[3].Contains("calendar (tentative)"))} | encrypted ok {Count(l => l[3] == "exported (encrypted)")} | skipped (not mail) {Count(l => l[3].StartsWith("skipped"))} | failed {Count(l => l[3].StartsWith("FAILED"))} (encrypted, no key {Count(l => l[3].Contains("encrypted -"))}) | body rebuilt {Count(l => l[4] == "yes")} | unread {Count(l => l[5] == "yes")}",
        };
        string stats = Path.Combine(workDir, "address_stats.txt");
        if (File.Exists(stats) && File.ReadAllText(stats).Split('\t') is { Length: 3 } st)
            outp.Add($"addresses: {st[0]} Notes name(s) -> SMTP | {st[1]} not resolvable ({st[2]} distinct, list in unresolved_names.txt)");
        var forms = log.Where(l => l[3].StartsWith("skipped")).GroupBy(l => l[2]).Select(g => $"{(g.Key == "" ? "(none)" : g.Key)}={g.Count()}").ToList();
        if (forms.Count > 0) outp.Add("skipped forms: " + string.Join(", ", forms));
        var fails = log.Where(l => l[3].StartsWith("FAILED")).GroupBy(l => l[3][8..]).Select(g => $"{g.Count()} x {g.Key}").ToList();
        if (fails.Count > 0) outp.Add("failure reasons: " + string.Join(" ; ", fails));
        outp.Add("");
        outp.Add($"{"folder",-40}{"expected",10}{"in PST",10}{"unread",10}{"unreadPST",10}  check");
        int bad = 0, n = 0;
        folderRows = new();
        foreach (var f in expected.Keys.Union(actual.Keys).OrderBy(x => x, StringComparer.Ordinal))
        {
            n++;
            var e = expected.GetValueOrDefault(f); var a = actual.GetValueOrDefault(f);
            bool ok = e == a; if (!ok) bad++;
            string name = anonymize && f != "(default folders)" ? $"folder{n:D3}" : f;
            outp.Add($"{name,-40}{e.total,10}{a.total,10}{e.unread,10}{a.unread,10}  {(ok ? "OK" : "MISMATCH")}");
            folderRows.Add((name, e.total, a.total, e.unread, a.unread, ok));
        }
        outp.Add("");
        outp.Add(bad == 0 ? "RESULT: all folders match" : $"RESULT: {bad} folder(s) do not match");
        allOk = bad == 0;
        return outp;
    }

    /// <summary>export_log.csv rows as string arrays (NoteID, UNID, Form, Status, BodyRebuilt, Unread, Folders, Subject).</summary>
    List<string[]> ReadLog()
    {
        var rows = new List<string[]>();
        foreach (var line in File.ReadLines(Path.Combine(workDir, "export_log.csv"), Encoding.UTF8).Skip(1))
        {
            var cells = new List<string>(); var cur = new StringBuilder(); bool q = false;
            for (int i = 0; i < line.Length; i++)
            {
                char c = line[i];
                if (q) { if (c == '"' && i + 1 < line.Length && line[i + 1] == '"') { cur.Append('"'); i++; } else if (c == '"') q = false; else cur.Append(c); }
                else if (c == '"') q = true;
                else if (c == ',') { cells.Add(cur.ToString()); cur.Clear(); }
                else cur.Append(c);
            }
            cells.Add(cur.ToString());
            if (cells.Count >= 8) rows.Add(cells.ToArray());
        }
        return rows;
    }

    // =====================================================================================================
    // settings, theme, icons and custom-drawn controls
    // =====================================================================================================
    /// <summary>Language and theme choice in %LOCALAPPDATA%\Nsf2Pst\settings.txt (empty = follow Windows).</summary>
    static class Prefs
    {
        public static string Lang = "", Theme = "";
        static string FilePath => Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Nsf2Pst", "settings.txt");
        public static void Load()
        {
            try
            {
                if (!File.Exists(FilePath)) return;
                foreach (var line in File.ReadAllLines(FilePath))
                {
                    var kv = line.Split('=', 2);
                    if (kv.Length != 2) continue;
                    if (kv[0] == "lang") Lang = kv[1].Trim();
                    if (kv[0] == "theme") Theme = kv[1].Trim();
                }
            }
            catch { }
        }
        public static void Save()
        {
            try { Directory.CreateDirectory(Path.GetDirectoryName(FilePath)!); File.WriteAllLines(FilePath, new[] { "lang=" + Lang, "theme=" + Theme }); } catch { }
        }
    }

    static class Th
    {
        public static bool Dark;
        public static Color Back, Surface, Line, Text, Muted, Accent, AccentDark, AccentSoft, AccentLine, Accent2, DashLine;
        public static Color Ok, OkSoft, Warn, WarnSoft, Bad, BadSoft, LogBack, LogText, Disabled, DisabledText;
        static Color C(int r, int g, int b) => Color.FromArgb(r, g, b);
        public static void Apply(bool dark)
        {
            Dark = dark;
            if (!dark)
            {
                Back = C(245, 246, 250); Surface = Color.White; Line = C(228, 230, 238); Text = C(28, 31, 42); Muted = C(107, 114, 128);
                Accent = C(91, 91, 214); AccentDark = C(70, 70, 190); AccentSoft = C(238, 238, 252); AccentLine = C(190, 190, 236); Accent2 = C(124, 92, 246); DashLine = C(178, 180, 230);
                Ok = C(18, 150, 84); OkSoft = C(234, 248, 240); Warn = C(214, 140, 8); WarnSoft = C(254, 246, 230); Bad = C(214, 52, 52); BadSoft = C(253, 236, 236);
                Disabled = C(200, 202, 222); DisabledText = Color.White;
            }
            else
            {
                Back = C(19, 20, 26); Surface = C(29, 31, 40); Line = C(50, 53, 66); Text = C(234, 236, 242); Muted = C(150, 156, 172);
                Accent = C(132, 128, 255); AccentDark = C(112, 108, 240); AccentSoft = C(45, 44, 80); AccentLine = C(96, 94, 170); Accent2 = C(160, 120, 255); DashLine = C(88, 88, 150);
                Ok = C(64, 200, 128); OkSoft = C(24, 46, 36); Warn = C(240, 176, 52); WarnSoft = C(52, 42, 22); Bad = C(242, 96, 96); BadSoft = C(60, 30, 34);
                Disabled = C(56, 58, 82); DisabledText = C(140, 142, 170);
            }
            LogBack = C(14, 15, 20); LogText = C(206, 212, 224);
        }
        public static Font Ui(float size, FontStyle style = FontStyle.Regular) => new(En ? "Segoe UI" : "Microsoft JhengHei UI", size, style);
        public static Font Num(float size) => new(FontFamily.Families.Any(f => f.Name == "Segoe UI Variable Display") ? "Segoe UI Variable Display" : "Segoe UI", size, FontStyle.Bold);
    }

    /// <summary>Glyphs of the Windows icon font (Segoe Fluent Icons on Windows 11, Segoe MDL2 Assets on Windows 10).</summary>
    static class Ico
    {
        public const string Mail = "", Upload = "", Settings = "", ChevronDown = "", ChevronUp = "", Forward = "",
            Stop = "", Folder = "", Document = "", List = "", Refresh = "", Contact = "", Calendar = "",
            Unread = "", Warning = "", Completed = "", ErrorBadge = "", CheckMark = "", Cancel = "",
            Globe = "", Moon = "", Sun = "";
        static readonly string family = FontFamily.Families.Any(f => f.Name == "Segoe Fluent Icons") ? "Segoe Fluent Icons" : "Segoe MDL2 Assets";
        public static Font Font(float size) => new(family, size);
    }

    static GraphicsPath Round(RectangleF r, float radius)
    {
        var p = new GraphicsPath();
        float d = Math.Min(radius * 2, Math.Min(r.Width, r.Height));
        if (d <= 0) { p.AddRectangle(r); return p; }
        p.AddArc(r.X, r.Y, d, d, 180, 90); p.AddArc(r.Right - d, r.Y, d, d, 270, 90);
        p.AddArc(r.Right - d, r.Bottom - d, d, d, 0, 90); p.AddArc(r.X, r.Bottom - d, d, d, 90, 90);
        p.CloseFigure();
        return p;
    }

    /// <summary>Rounded card with a hairline border (dashed for the drop zone, accent while a file is dragged over it).</summary>
    sealed class Card : Panel
    {
        public int Radius = 12;
        public bool Dashed, Highlight;
        public Color Fill = Th.Surface, Border = Th.Line;
        public Card() { DoubleBuffered = true; ResizeRedraw = true; BackColor = Th.Surface; SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.UserPaint | ControlStyles.OptimizedDoubleBuffer, true); }
        protected override void OnPaintBackground(PaintEventArgs e)
        {
            var g = e.Graphics;
            g.Clear(Parent?.BackColor ?? Th.Back);
            g.SmoothingMode = SmoothingMode.AntiAlias;
            var r = new RectangleF(0.5f, 0.5f, Width - 1.5f, Height - 1.5f);
            using var path = Round(r, S(Radius));
            using var b = new SolidBrush(Fill);
            g.FillPath(b, path);
            Color border = Highlight ? Th.Accent : Dashed ? Th.DashLine : Border;
            if (border.IsEmpty) return;
            using var pen = new Pen(border, Highlight || Dashed ? S(2) : 1);
            if (Dashed && !Highlight) pen.DashPattern = new[] { 4f, 3f };
            g.DrawPath(pen, path);
        }
    }

    /// <summary>Rounded button: filled accent (primary) or outline (secondary); text, icon or both.</summary>
    sealed class Pill : Button
    {
        readonly bool primary, small, iconAfter;
        readonly string? icon;
        bool hover, down;
        public Pill(string text, bool primary = false, bool small = false, string? icon = null, bool iconAfter = false)
        {
            this.primary = primary; this.small = small; this.icon = icon; this.iconAfter = iconAfter;
            Text = text;
            SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw, true);
            FlatStyle = FlatStyle.Flat; FlatAppearance.BorderSize = 0; Cursor = Cursors.Hand;
            Font = Th.Ui(small ? 9.25f : 10f, primary ? FontStyle.Bold : FontStyle.Regular);
            Fit();
        }
        /// <summary>Size from the text; called again once the window knows its DPI.</summary>
        public void Fit()
        {
            Height = S(small ? 34 : 42);
            int w = Text.Length == 0 ? Height : TextRenderer.MeasureText(Text, Font).Width + S(small ? 32 : 44) + (icon != null ? S(26) : 0);
            MinimumSize = new Size(w, 0); Width = w;
        }
        protected override void OnMouseEnter(EventArgs e) { hover = true; Invalidate(); base.OnMouseEnter(e); }
        protected override void OnMouseLeave(EventArgs e) { hover = down = false; Invalidate(); base.OnMouseLeave(e); }
        protected override void OnMouseDown(MouseEventArgs e) { down = true; Invalidate(); base.OnMouseDown(e); }
        protected override void OnMouseUp(MouseEventArgs e) { down = false; Invalidate(); base.OnMouseUp(e); }
        protected override void OnEnabledChanged(EventArgs e) { Invalidate(); base.OnEnabledChanged(e); }
        protected override void OnPaint(PaintEventArgs e)
        {
            var g = e.Graphics;
            g.Clear(Parent?.BackColor ?? Th.Back);
            g.SmoothingMode = SmoothingMode.AntiAlias;
            var r = new RectangleF(0.5f, 0.5f, Width - 1.5f, Height - 1.5f);
            using var path = Round(r, S(small ? 8 : 10));
            Color fill = !Enabled ? (primary ? Th.Disabled : Th.Back) : primary ? (down || hover ? Th.AccentDark : Th.Accent) : (down ? Th.Line : hover ? Th.AccentSoft : Th.Surface);
            using (var b = new SolidBrush(fill)) g.FillPath(b, path);
            if (!primary) { using var pen = new Pen(Enabled && hover ? Th.AccentLine : Th.Line); g.DrawPath(pen, path); }
            Color fg = primary ? (Enabled ? Color.White : Th.DisabledText) : Enabled ? Th.Text : Th.Muted;
            var tsize = TextRenderer.MeasureText(g, Text, Font, Size.Empty, TextFormatFlags.NoPadding);
            if (icon == null) { TextRenderer.DrawText(g, Text, Font, ClientRectangle, fg, TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter | TextFormatFlags.NoPadding); return; }
            using var iconFont = Ico.Font(Font.Size * 0.95f);
            if (Text.Length == 0) { TextRenderer.DrawText(g, icon, iconFont, ClientRectangle, fg, TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter | TextFormatFlags.NoPadding); return; }
            var isize = TextRenderer.MeasureText(g, icon, iconFont, Size.Empty, TextFormatFlags.NoPadding);
            int gap = S(8), total = isize.Width + gap + tsize.Width, x = (Width - total) / 2;
            int ty = (Height - tsize.Height) / 2, iy = (Height - isize.Height) / 2 + S(1);
            if (iconAfter)
            {
                TextRenderer.DrawText(g, Text, Font, new Point(x, ty), fg, TextFormatFlags.NoPadding);
                TextRenderer.DrawText(g, icon, iconFont, new Point(x + tsize.Width + gap, iy), fg, TextFormatFlags.NoPadding);
            }
            else
            {
                TextRenderer.DrawText(g, icon, iconFont, new Point(x, iy), fg, TextFormatFlags.NoPadding);
                TextRenderer.DrawText(g, Text, Font, new Point(x + isize.Width + gap, ty), fg, TextFormatFlags.NoPadding);
            }
        }
    }

    /// <summary>Rounded progress bar, 0-1000.</summary>
    sealed class Bar : Control
    {
        int value;
        public event Action? ValueChanged;
        public int Value { get => value; set { int v = Math.Clamp(value, 0, 1000); if (v == this.value) return; this.value = v; Invalidate(); ValueChanged?.Invoke(); } }
        public Bar() { SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw, true); }
        protected override void OnPaint(PaintEventArgs e)
        {
            var g = e.Graphics;
            g.Clear(Parent?.BackColor ?? Th.Surface);
            g.SmoothingMode = SmoothingMode.AntiAlias;
            var r = new RectangleF(0, 0, Width - 1, Height - 1);
            using (var track = Round(r, Height / 2f)) using (var b = new SolidBrush(Th.AccentSoft)) g.FillPath(b, track);
            if (value <= 0) return;
            var f = new RectangleF(0, 0, Math.Max(Height, (Width - 1) * value / 1000f), Height - 1);
            using var fill = Round(f, Height / 2f);
            using var lg = new LinearGradientBrush(new RectangleF(0, 0, Width, Height), Th.Accent2, Th.Accent, LinearGradientMode.Horizontal);
            g.FillPath(lg, fill);
        }
    }

    /// <summary>Three-step indicator: export -> write PST -> verify.</summary>
    sealed class Stepper : Control
    {
        int current;
        public int Current { get => current; set { current = value; Invalidate(); } }
        public Stepper() { SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw, true); }
        protected override void OnPaint(PaintEventArgs e)
        {
            var g = e.Graphics;
            g.Clear(Parent?.BackColor ?? Th.Surface);
            g.SmoothingMode = SmoothingMode.AntiAlias;
            string[] names = { L("匯出郵件", "Export mail"), L("寫入 PST", "Write PST"), L("核對結果", "Verify") };
            int n = names.Length, d = S(28), top = S(4);
            float slot = (Width - 1f) / n;
            using var numFont = Th.Ui(10f, FontStyle.Bold);
            using var lblFont = Th.Ui(9.5f);
            using var lblBold = Th.Ui(9.5f, FontStyle.Bold);
            using var icon = Ico.Font(10f);
            for (int i = 0; i < n; i++)
            {
                float cx = slot * i + slot / 2;
                if (i < n - 1)
                {
                    using var pen = new Pen(i < current ? Th.Accent : Th.Line, S(2));
                    g.DrawLine(pen, cx + d / 2f + S(6), top + d / 2f, cx + slot - d / 2f - S(6), top + d / 2f);
                }
                var c = new RectangleF(cx - d / 2f, top, d, d);
                bool done = i < current, active = i == current;
                using (var b = new SolidBrush(done || active ? Th.Accent : Th.Surface)) g.FillEllipse(b, c);
                if (!done && !active) { using var pen = new Pen(Th.Line, S(2)); g.DrawEllipse(pen, c); }
                if (active) { using var halo = new Pen(Color.FromArgb(60, Th.Accent), S(4)); g.DrawEllipse(halo, RectangleF.Inflate(c, S(3), S(3))); }
                var cr = Rectangle.Round(c);
                if (done) TextRenderer.DrawText(g, Ico.CheckMark, icon, cr, Color.White, TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter);
                else TextRenderer.DrawText(g, (i + 1).ToString(), numFont, cr, active ? Color.White : Th.Muted, TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter);
                var lr = new Rectangle((int)(cx - slot / 2), top + d + S(8), (int)slot, S(22));
                TextRenderer.DrawText(g, names[i], active ? lblBold : lblFont, lr, active || done ? Th.Text : Th.Muted, TextFormatFlags.HorizontalCenter);
            }
        }
    }

    /// <summary>Accent-tinted rounded square with an icon glyph.</summary>
    sealed class IconBadge : Control
    {
        readonly string glyph;
        public IconBadge(string glyph) { this.glyph = glyph; SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw, true); }
        protected override void OnPaint(PaintEventArgs e)
        {
            var g = e.Graphics;
            g.Clear(Parent?.BackColor ?? Th.Surface);
            g.SmoothingMode = SmoothingMode.AntiAlias;
            using var path = Round(new RectangleF(0, 0, Width - 1, Height - 1), Width / 3.6f);
            using var lg = new LinearGradientBrush(new RectangleF(0, 0, Width, Height), Color.FromArgb(124, 92, 246), Color.FromArgb(91, 91, 214), LinearGradientMode.ForwardDiagonal);
            g.FillPath(lg, path);
            using var f = Ico.Font(Height * 0.36f * 72f / DeviceDpi);
            TextRenderer.DrawText(g, glyph, f, ClientRectangle, Color.White, TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter);
        }
    }

    /// <summary>Progress on the taskbar button (ITaskbarList3).</summary>
    static class Taskbar
    {
        public const int NoProgress = 0, Normal = 2, Error = 4, Paused = 8;
        [ComImport, Guid("ea1afb91-9e28-4b86-90e9-9e9f8a5eefaf"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        interface ITaskbarList3
        {
            void HrInit(); void AddTab(IntPtr h); void DeleteTab(IntPtr h); void ActivateTab(IntPtr h); void SetActiveAlt(IntPtr h);
            void MarkFullscreenWindow(IntPtr h, [MarshalAs(UnmanagedType.Bool)] bool full);
            void SetProgressValue(IntPtr h, ulong done, ulong total);
            void SetProgressState(IntPtr h, int flags);
        }
        [ComImport, Guid("56FDF344-FD6D-11d0-958A-006097C9A090"), ClassInterface(ClassInterfaceType.None)]
        class TaskbarInstance { }
        static ITaskbarList3? list;
        static ITaskbarList3? Get()
        {
            if (list != null) return list;
            try { list = (ITaskbarList3)new TaskbarInstance(); list.HrInit(); } catch { list = null; }
            return list;
        }
        public static void Set(IntPtr h, int done, int total, int state = Normal)
        {
            try
            {
                var l = Get(); if (l == null) return;
                if (total <= 0) { l.SetProgressState(h, NoProgress); return; }
                l.SetProgressState(h, state); l.SetProgressValue(h, (ulong)done, (ulong)total);
            }
            catch { }
        }
        public static void Clear(IntPtr h) => Set(h, 0, 0);
    }

    static class Native
    {
        [DllImport("dwmapi.dll")] static extern int DwmSetWindowAttribute(IntPtr hwnd, int attr, ref int value, int size);
        [DllImport("uxtheme.dll", CharSet = CharSet.Unicode)] static extern int SetWindowTheme(IntPtr hwnd, string? app, string? idList);
        [DllImport("user32.dll")] static extern bool FlashWindowEx(ref FLASHWINFO info);
        [StructLayout(LayoutKind.Sequential)] struct FLASHWINFO { public uint cbSize; public IntPtr hwnd; public uint dwFlags; public uint uCount; public uint dwTimeout; }

        static int Rgb(Color c) => c.R | c.G << 8 | c.B << 16;

        /// <summary>Windows 11: title bar in the header colour, dark caption text in dark mode, rounded corners (ignored on older Windows).</summary>
        public static void StyleTitleBar(IntPtr hwnd, Color caption, Color border, bool dark)
        {
            try
            {
                int d = dark ? 1 : 0; DwmSetWindowAttribute(hwnd, 20, ref d, 4);   // DWMWA_USE_IMMERSIVE_DARK_MODE
                int c = Rgb(caption); DwmSetWindowAttribute(hwnd, 35, ref c, 4);    // DWMWA_CAPTION_COLOR
                int b = Rgb(border); DwmSetWindowAttribute(hwnd, 34, ref b, 4);     // DWMWA_BORDER_COLOR
                int round = 2; DwmSetWindowAttribute(hwnd, 33, ref round, 4);       // DWMWA_WINDOW_CORNER_PREFERENCE = round
            }
            catch { }
        }

        public static void DarkScrollbars(IntPtr hwnd, bool dark) { try { SetWindowTheme(hwnd, dark ? "DarkMode_Explorer" : null, null); } catch { } }

        public static bool SystemUsesDarkTheme()
        {
            try { return Microsoft.Win32.Registry.GetValue(@"HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize", "AppsUseLightTheme", 1) is int v && v == 0; }
            catch { return false; }
        }

        /// <summary>Flashes the taskbar button until the window comes to the foreground.</summary>
        public static void Flash(IntPtr hwnd)
        {
            try { var f = new FLASHWINFO { cbSize = (uint)Marshal.SizeOf<FLASHWINFO>(), hwnd = hwnd, dwFlags = 3 | 12, uCount = 0, dwTimeout = 0 }; FlashWindowEx(ref f); } catch { }
        }
    }
}
