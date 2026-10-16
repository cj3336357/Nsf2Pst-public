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
    internal static bool En;
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
    static Point? carryBounds;   // window position kept across a rebuild after a DPI change
    bool dpiMoved;

    // ---- controls used by the logic (names = UI Automation ids for the automated tests)
    readonly TextBox txtNsf = new(), txtPst = new(), txtServer = new(), txtUser = new(), txtLog = new(), txtId = new(), txtDir = new(),
                     txtPwd = new() { UseSystemPasswordChar = true };
    readonly CheckBox chkAnon = new() { Text = L("報告唔顯示資料夾名稱同錯誤詳情（方便交俾其他人）", "Leave folder names and error details out of the report (for sharing)"), AutoSize = true };
    readonly Pill btnStart = new(L("開始轉換", "Start conversion"), primary: true, icon: Ico.Forward, iconAfter: true), btnStop = new(L("停止", "Stop"), icon: Ico.Stop);
    readonly Pill btnOpenOut = new(L("開啟 PST 位置", "Show PST in folder"), primary: true, icon: Ico.Folder), btnOpenLog = new(L("逐項記錄 (CSV)", "Item log (CSV)"), icon: Ico.Document);
    readonly Pill btnSummary = new(L("圖表報告", "Visual report"), icon: Ico.Chart), btnAgain = new(L("轉換另一個郵箱", "Convert another mailbox"), icon: Ico.Refresh);
    readonly Pill btnNsf = new(L("瀏覽檔案", "Browse files"), small: true), btnPst = new(L("更改", "Change"), small: true),
                  btnId = new(L("瀏覽", "Browse"), small: true), btnDir = new(L("瀏覽", "Browse"), small: true);
    readonly Pill btnLang = new(En ? "中文" : "English", small: true, icon: Ico.Globe), btnTheme = new("", small: true, icon: Th.Dark ? Ico.Sun : Ico.Moon);
    readonly Label lblStatus = new() { Text = L("請選擇 NSF 檔案。", "Choose an NSF file."), AutoSize = true };
    readonly Bar progress = new();
    Process? child;
    bool stopRequested;
    string workDir = "", statusBase = "";
    string? lastFatal;   // last "fatal: ..." line of a child: shown as the reason when a run fails

    /// <summary>The reason of a failed step in the UI language (known causes), else the child's own text.</summary>
    static string FriendlyFatal(string f) =>
        f.Contains("in use by another program") ? L("PST 檔被其他程式使用緊（例如 Outlook），請先關閉 Outlook 或者用 PST 嘅副本。", "The PST is in use by another program (e.g. Outlook). Close Outlook or use a copy of the PST.")
        : f.Contains("not an Outlook data file") ? L("呢個檔案唔係 Outlook 資料檔（PST）。", "This file is not an Outlook data file (PST).")
        : f.Contains("not a Notes database") ? L("呢個檔案唔係 Notes 資料庫（NSF）。", "This file is not a Notes database (NSF).")
        : f.Contains("cannot open the NSF") ? L("開唔到 NSF：檔案可能已損壞，或者而家個 Notes ID 冇權限。", "The NSF cannot be opened: it may be damaged, or this Notes ID has no access.")
        : f.Contains("path is too long") ? L("NSF 嘅路徑太長，Notes 開唔到，請用較短嘅資料夾或者檔名。", "The NSF path is too long for Notes. Use a shorter folder or file name.")
        : f.Contains("not found") ? L("搵唔到檔案。", "File not found.")
        : f;

    // ---- direction: NSF -> PST (default) or PST -> NSF. txtNsf / txtPst keep their names (UI Automation ids) and always
    // hold the SOURCE / the OUTPUT file: in PST -> NSF mode txtNsf holds the PST and txtPst the new NSF.
    static bool Rev;
    static string D(string nsfZh, string nsfEn, string pstZh, string pstEn) => Rev ? L(pstZh, pstEn) : L(nsfZh, nsfEn);
    readonly Segmented seg = new(L("NSF → PST", "NSF → PST"), L("PST → NSF", "PST → NSF")) { Name = "segDirection" };
    readonly Label lblHeading = new(), lblSub = new(), lblUpText = new(), lblUpHint = new(), lblOutLabel = new(), lblFoot = new(), lblSubTitle = new();
    readonly Label hintServer = new(), hintUser = new();
    readonly List<Control> nsfOnly = new();   // advanced rows that only apply to NSF -> PST (directory, ID, password)
    Wordmark wordmark = null!;
    Action layoutDrop = () => { };

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
    readonly TableLayoutPanel tileRow = new();
    bool unreadFailed;   // PST -> NSF: the unread marks could not be written
    readonly List<string> writeFails = new();   // "FAILED line N (...): reason" of the writing step (import / nsf-import)
    readonly Card dropZone = new() { Dashed = true };
    readonly Panel dropEmpty = new(), dropFile = new();
    readonly Label lblFileName = new(), lblFileMeta = new(), lblEnvWarn = new();
    readonly Expander advHead = new();
    readonly Card advCard = new();
    readonly Panel logPanel = new();
    readonly LinkLabel btnLog = new();
    readonly ToolTip tips = new();
    readonly System.Windows.Forms.Timer clock = new() { Interval = 1000 };
    DateTime runStart;
    List<(string name, int exp, int pst, int expU, int pstU, bool ok)> folderRows = new();
    sealed record Stats(int Docs, int Mails, int Contacts, int Calendar, int Invites, int Encrypted, int Skipped, int Failed, int Unread,
                        double Seconds, long NsfSize, long PstSize, int AddrOk, int AddrBad, List<(string reason, int count)> Fails);
    Stats? report;
    bool envOk = true;

    static float k = 1;                       // DPI scale
    static int S(int px) => (int)Math.Round(px * k);

    public MainForm()
    {
        Text = L("Nsf2Pst - Notes 郵箱 (NSF) 轉 Outlook PST", "Nsf2Pst - Notes mailbox (NSF) to Outlook PST");
        Font = Th.Ui(Th.Body);
        AutoScaleMode = AutoScaleMode.None;
        // the scale of the screen the window will open on (after a DPI change: the screen it was moved to)
        var openAt = carryBounds ?? Cursor.Position;
        k = Native.MonitorDpi(openAt) / 96f;
        BackColor = Th.Back;
        try { Icon = Icon.ExtractAssociatedIcon(Environment.ProcessPath!); } catch { }
        var work = Screen.FromPoint(openAt).WorkingArea;
        MinimumSize = new Size(Math.Min(S(780), work.Width), Math.Min(S(640), work.Height));
        Size = new Size(Math.Min(S(980), work.Width), Math.Min(S(880), work.Height));
        StartPosition = FormStartPosition.CenterScreen;
        if (carryBounds is { } at) { StartPosition = FormStartPosition.Manual; Location = at; carryBounds = null; }
        AllowDrop = true;
        KeyPreview = true;

        foreach (var b in new[] { btnStart, btnStop, btnOpenOut, btnOpenLog, btnSummary, btnAgain, btnNsf, btnPst, btnId, btnDir, btnLang, btnTheme }) b.Fit();
        seg.Fit();
        BuildSetup(); BuildRun(); BuildDone();
        host.BackColor = Th.Back;
        host.Resize += (_, _) => { int side = Math.Max(S(28), (host.ClientSize.Width - S(820)) / 2); host.Padding = new Padding(side, S(26), side, S(26)); };
        ShowView(viewSetup);

        // log panel (hidden until asked for)
        txtLog.Multiline = true; txtLog.ReadOnly = true; txtLog.ScrollBars = ScrollBars.Vertical; txtLog.WordWrap = true;   // long paths wrap instead of running off the edge
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
        btnSummary.Click += (_, _) => OpenFile(Path.Combine(workDir, "summary.html"));
        btnAgain.Click += (_, _) => { txtNsf.Text = ""; txtPst.Text = ""; ShowView(viewSetup); NsfChanged(); Taskbar.Set(Handle, 0, 0); };
        btnLang.Click += (_, _) => Rebuild(lang: En ? "zh" : "en");
        btnTheme.Click += (_, _) => Rebuild(theme: Th.Dark ? "light" : "dark");
        tips.SetToolTip(btnTheme, L("切換深色／淺色", "Switch dark / light"));
        btnTheme.AccessibleName = L("切換深色／淺色", "Switch dark / light");
        progress.AccessibleRole = AccessibleRole.ProgressBar; stepper.AccessibleRole = AccessibleRole.StaticText;
        lblDoneIcon.AccessibleRole = AccessibleRole.Graphic;
        tips.SetToolTip(btnLang, L("Switch to English", "切換到中文"));
        txtNsf.TextChanged += (_, _) => NsfChanged();
        foreach (var t in new Control[] { txtServer, txtDir, txtUser, txtId, txtPwd }) t.TextChanged += (_, _) => UpdateAdvSummary();
        chkAnon.CheckedChanged += (_, _) => UpdateAdvSummary();
        clock.Tick += (_, _) => UpdateElapsed();
        DragEnter += OnDragEnter; DragDrop += OnDragDrop; DragLeave += (_, _) => DropHighlight(false);
        KeyDown += (_, e) =>
        {
            if (e.Control && e.KeyCode == Keys.O && child == null) { e.SuppressKeyPress = true; PickNsf(); }
            else if (e.KeyCode == Keys.Escape && child != null) { e.SuppressKeyPress = true; StopChild(); }   // Esc = Stop
        };
        // the window is laid out for one DPI: moved to a screen with another scale it rebuilds itself once the move ends
        DpiChanged += (_, _) => dpiMoved = true;
        ResizeEnd += (_, _) => { if (dpiMoved && child == null) { dpiMoved = false; carryBounds = Location; Rebuild(); } };
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
            Native.DarkScrollbars(txtLog.Handle);
            if (Th.Dark) { Native.DarkScrollbars(host.Handle); Native.DarkScrollbars(grid.Handle); }
        };

        if (carry is { } c)
        {
            txtNsf.Text = c.nsf; txtPst.Text = c.pst; txtServer.Text = c.server; txtUser.Text = c.user; txtId.Text = c.id; txtPwd.Text = c.pwd; txtDir.Text = c.dir; chkAnon.Checked = c.anon;
            if (c.adv) ToggleAdvanced();
            carry = null;
        }
        ApplyDirection();
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
        var logo = new IconBadge(Ico.Mail, badge: true) { Size = new Size(S(40), S(40)), Location = new Point(S(24), S(14)) };
        var title = wordmark = new Wordmark { Location = new Point(S(76), S(11)), BackColor = Th.Surface };
        var sub = lblSubTitle;
        sub.Font = Th.Ui(Th.Small); sub.ForeColor = Th.Muted; sub.AutoSize = true; sub.Location = new Point(S(76), S(40)); sub.BackColor = Th.Surface;
        // the environment check only shows up when something is missing (banner in the setup view)
        bar.Controls.AddRange(new Control[] { logo, title, sub, btnLang, btnTheme });
        void Place()
        {
            int x = bar.Width - S(24) - btnLang.Width; btnLang.Location = new Point(x, (bar.Height - btnLang.Height) / 2);
            x -= S(8) + btnTheme.Width; btnTheme.Location = new Point(x, (bar.Height - btnTheme.Height) / 2);
        }
        bar.Resize += (_, _) => Place();
        return bar;
    }

    Control BuildFooter()
    {
        var bar = new Panel { Dock = DockStyle.Bottom, Height = S(40), BackColor = Th.Surface, Padding = new Padding(S(24), 0, S(24), 0) };
        bar.Paint += (_, e) => { using var p = new Pen(Th.Line); e.Graphics.DrawLine(p, 0, 0, bar.Width, 0); };
        lblStatus.Font = Th.Ui(Th.Body, FontStyle.Bold); lblStatus.BackColor = Th.Surface; lblStatus.ForeColor = Th.Text; lblStatus.Location = new Point(S(24), S(11));
        btnLog.Text = L("顯示詳細記錄", "Show details"); btnLog.AutoSize = true; btnLog.Font = Th.Ui(Th.Body); btnLog.LinkColor = Th.Accent; btnLog.ActiveLinkColor = Th.AccentDark;
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

    /// <summary>A rounded input with a borderless TextBox (and an optional button on the right).</summary>
    static Card Field(TextBox box, string placeholder, Control? right = null)
    {
        var c = new Card { Height = S(44), Dock = DockStyle.Top, Padding = new Padding(S(12), S(11), S(right == null ? 12 : 6), S(6)), Margin = new Padding(0, 0, 0, S(12)) };
        box.BorderStyle = BorderStyle.None; box.BackColor = Th.Surface; box.ForeColor = Th.Text; box.Font = Th.Ui(Th.Body);
        box.PlaceholderText = placeholder; box.Dock = DockStyle.Fill;
        if (right != null) { right.Dock = DockStyle.Right; c.Controls.Add(right); c.Padding = new Padding(S(12), S(6), S(6), S(6)); var inner = new Panel { Dock = DockStyle.Fill, Padding = new Padding(0, S(5), S(8), 0), BackColor = Th.Surface }; inner.Controls.Add(box); c.Controls.Add(inner); inner.BringToFront(); }
        else c.Controls.Add(box);
        return c;
    }

    static Label FieldLabel(string text) =>
        new() { Text = text, Font = Th.Ui(Th.Body, FontStyle.Bold), ForeColor = Th.Text, AutoSize = true, Margin = new Padding(0, 0, 0, S(6)), BackColor = Th.Back };

    void BuildSetup()
    {
        viewSetup = Column();
        seg.Margin = new Padding(0, 0, 0, S(18));
        seg.Changed += () => { if (child == null) { Rev = seg.Index == 1; txtNsf.Text = ""; txtPst.Text = ""; ApplyDirection(); } };
        viewSetup.Controls.Add(seg);
        lblHeading.Font = Th.Ui(Th.Heading, FontStyle.Bold); lblHeading.ForeColor = Th.Text; lblHeading.AutoSize = true; lblHeading.Margin = new Padding(0, 0, 0, S(2)); lblHeading.BackColor = Th.Back;
        lblSub.Font = Th.Ui(Th.Body); lblSub.ForeColor = Th.Muted; lblSub.AutoSize = true; lblSub.Margin = new Padding(0, 0, 0, S(18)); lblSub.BackColor = Th.Back;
        viewSetup.Controls.Add(lblHeading);
        viewSetup.Controls.Add(lblSub);

        lblEnvWarn.AutoSize = false; lblEnvWarn.Dock = DockStyle.Top; lblEnvWarn.Height = S(40); lblEnvWarn.Visible = false;
        lblEnvWarn.BackColor = Th.BadSoft; lblEnvWarn.ForeColor = Th.Bad; lblEnvWarn.Font = Th.Ui(Th.Body, FontStyle.Bold);
        lblEnvWarn.TextAlign = ContentAlignment.MiddleLeft; lblEnvWarn.Padding = new Padding(S(14), 0, 0, 0); lblEnvWarn.Margin = new Padding(0, 0, 0, S(14));
        viewSetup.Controls.Add(lblEnvWarn);

        // drop zone: empty state / chosen file state
        dropZone.Dock = DockStyle.Top; dropZone.Margin = new Padding(0, 0, 0, S(12)); dropZone.AllowDrop = true;
        dropZone.DragEnter += OnDragEnter; dropZone.DragDrop += OnDragDrop; dropZone.DragLeave += (_, _) => DropHighlight(false);
        dropEmpty.Dock = DockStyle.Fill; dropEmpty.BackColor = Th.Surface;
        var upIcon = new Label { Text = Ico.Upload, Font = Ico.Font(26f), ForeColor = Th.Accent, AutoSize = true, BackColor = Th.Surface };
        var upText = lblUpText; upText.Font = Th.Ui(Th.Title, FontStyle.Bold); upText.ForeColor = Th.Text; upText.AutoSize = true; upText.BackColor = Th.Surface;
        var upHint = lblUpHint; upHint.Font = Th.Ui(Th.Body); upHint.ForeColor = Th.Muted; upHint.AutoSize = true; upHint.BackColor = Th.Surface;
        dropEmpty.Controls.AddRange(new Control[] { upIcon, upText, upHint, btnNsf });
        layoutDrop = () =>
        {
            int cx = dropEmpty.Width / 2, total = upIcon.Height + upText.Height + upHint.Height + btnNsf.Height + S(16), y = Math.Max(S(8), (dropEmpty.Height - total) / 2);
            upIcon.Location = new Point(cx - upIcon.Width / 2, y); y += upIcon.Height + S(4);
            upText.Location = new Point(cx - upText.Width / 2, y); y += upText.Height + S(2);
            upHint.Location = new Point(cx - upHint.Width / 2, y); y += upHint.Height + S(10);
            btnNsf.Location = new Point(cx - btnNsf.Width / 2, y);
        };
        dropEmpty.Resize += (_, _) => layoutDrop();
        dropFile.Dock = DockStyle.Fill; dropFile.BackColor = Th.Surface; dropFile.Visible = false;
        var fileIcon = new IconBadge(Ico.Mail) { Size = new Size(S(52), S(52)), Location = new Point(S(24), S(26)) };
        lblFileName.AutoSize = true; lblFileName.Font = Th.Ui(Th.Title, FontStyle.Bold); lblFileName.ForeColor = Th.Text; lblFileName.BackColor = Th.Surface; lblFileName.Location = new Point(S(90), S(24));
        lblFileMeta.AutoSize = true; lblFileMeta.Font = Th.Ui(Th.Body); lblFileMeta.ForeColor = Th.Muted; lblFileMeta.BackColor = Th.Surface; lblFileMeta.Location = new Point(S(91), S(54));
        var change = new LinkLabel { Text = L("更換檔案", "Choose another file"), AutoSize = true, Font = Th.Ui(Th.Body), LinkColor = Th.Accent, ActiveLinkColor = Th.AccentDark, LinkBehavior = LinkBehavior.HoverUnderline, BackColor = Th.Surface, Location = new Point(S(91), S(78)) };
        change.LinkClicked += (_, _) => PickNsf();
        dropFile.Controls.AddRange(new Control[] { fileIcon, lblFileName, lblFileMeta, change });
        var pathBox = Field(txtNsf, "");
        pathBox.Dock = DockStyle.Bottom; pathBox.Margin = new Padding(0); pathBox.Height = S(40); pathBox.Fill = Th.Back; txtNsf.BackColor = Th.Back;
        var pathWrap = new Panel { Dock = DockStyle.Bottom, Height = S(52), Padding = new Padding(S(14), 0, S(14), S(12)), BackColor = Th.Surface };
        pathWrap.Controls.Add(pathBox);
        dropZone.Controls.Add(dropEmpty); dropZone.Controls.Add(dropFile); dropZone.Controls.Add(pathWrap);
        dropZone.Padding = new Padding(S(2));
        dropZone.Height = S(232);
        viewSetup.Controls.Add(dropZone);

        lblOutLabel.Font = Th.Ui(Th.Body, FontStyle.Bold); lblOutLabel.ForeColor = Th.Text; lblOutLabel.AutoSize = true; lblOutLabel.Margin = new Padding(0, 0, 0, S(6)); lblOutLabel.BackColor = Th.Back;
        viewSetup.Controls.Add(lblOutLabel);
        viewSetup.Controls.Add(Field(txtPst, "", btnPst));

        // advanced settings expander
        advHead.Name = "btnAdvanced"; advHead.Text = L("進階設定", "Advanced settings");
        advHead.Height = S(50); advHead.Dock = DockStyle.Top; advHead.Margin = new Padding(0, S(2), 0, S(8));
        advHead.Click += (_, _) => ToggleAdvanced();
        viewSetup.Controls.Add(advHead);

        var advGrid = Column(); advGrid.BackColor = Th.Surface; advGrid.Dock = DockStyle.Top;
        // one advanced row = label, field, hint; returns the three controls (rows that only apply to NSF -> PST are hidden in PST -> NSF mode)
        Control[] Add(string label, string hint, TextBox box, string placeholder, Control? right = null, Label? hintLabel = null)
        {
            var l = FieldLabel(label); l.BackColor = Th.Surface; advGrid.Controls.Add(l);
            var f = Field(box, placeholder, right); f.Fill = Th.Back; box.BackColor = Th.Back; foreach (Control c in f.Controls) if (c is Panel p) p.BackColor = Th.Back;
            f.Margin = new Padding(0, 0, 0, S(2)); advGrid.Controls.Add(f);
            var h = hintLabel ?? new Label();
            h.Text = hint; h.Font = Th.Ui(Th.Small); h.ForeColor = Th.Muted; h.AutoSize = true; h.BackColor = Th.Surface; h.Margin = new Padding(S(2), 0, 0, S(12));
            advGrid.Controls.Add(h);
            return new Control[] { l, f, h };
        }
        string opt = L("選填", "Optional");
        Add(L("郵件伺服器（Domino）", "Mail server (Domino)"), "", txtServer, L("例如 mail01/ExampleCorp", "e.g. mail01/ExampleCorp"), hintLabel: hintServer);
        nsfOnly.AddRange(Add(L("公司通訊錄（names.nsf）", "Company directory (names.nsf)"), L("用嚟將 Notes 名轉做電郵地址。揀一個 names.nsf 副本，或者填「伺服器名稱!!names.nsf」。",
            "Turns Notes names into email addresses. Pick a copy of names.nsf, or enter \"server name!!names.nsf\"."), txtDir, opt, btnDir));
        Add(L("用邊個人嘅已讀記錄", "Whose read marks to use"), "", txtUser, L("Notes 全名，例如 陳大文/IT/ExampleCorp", "Full Notes name, e.g. John Smith/IT/ExampleCorp"), hintLabel: hintUser);
        nsfOnly.AddRange(Add(L("郵箱主人 Notes ID", "Mailbox owner's Notes ID"), L("只有加密郵件先需要；留空用而家登入嘅 ID。", "Only needed for encrypted mail; leave empty to use the current ID."), txtId, opt, btnId));
        nsfOnly.AddRange(Add(L("ID 密碼", "ID password"), L("只會傳畀轉換程序，唔會寫入記錄。", "Passed only to the converter process, never written to a log."), txtPwd, opt));
        chkAnon.BackColor = Th.Surface; chkAnon.ForeColor = Th.Text; chkAnon.Font = Th.Ui(Th.Body); chkAnon.Margin = new Padding(0, S(2), 0, S(4));
        advGrid.Controls.Add(chkAnon);
        advCard.Dock = DockStyle.Top; advCard.AutoSize = true; advCard.AutoSizeMode = AutoSizeMode.GrowAndShrink;
        advCard.Padding = new Padding(S(20), S(16), S(20), S(12)); advCard.Margin = new Padding(0, 0, 0, S(14)); advCard.Visible = false;
        advCard.Controls.Add(advGrid);
        viewSetup.Controls.Add(advCard);

        btnStart.Height = S(52); btnStart.Dock = DockStyle.Top; btnStart.Margin = new Padding(0, S(10), 0, S(10)); btnStart.Font = Th.Ui(Th.Title, FontStyle.Bold);
        viewSetup.Controls.Add(btnStart);
        lblFoot.Font = Th.Ui(Th.Small); lblFoot.ForeColor = Th.Muted; lblFoot.AutoSize = true; lblFoot.BackColor = Th.Back; lblFoot.Margin = new Padding(0, 0, 0, S(6));
        viewSetup.Controls.Add(lblFoot);
    }

    /// <summary>Puts every direction-dependent text of the window in place (NSF -> PST or PST -> NSF).</summary>
    void ApplyDirection()
    {
        seg.Index = Rev ? 1 : 0;
        Text = D("Nsf2Pst - Notes 郵箱 (NSF) 轉 Outlook PST", "Nsf2Pst - Notes mailbox (NSF) to Outlook PST",
                 "Nsf2Pst - Outlook PST 轉 Notes 郵箱 (NSF)", "Nsf2Pst - Outlook PST to Notes mailbox (NSF)");
        lblSubTitle.Text = D("Notes 郵箱 (NSF) 轉 Outlook PST", "Notes mailbox (NSF) to Outlook PST", "Outlook PST 轉 Notes 郵箱 (NSF)", "Outlook PST to Notes mailbox (NSF)");
        lblHeading.Text = D("選擇要轉換嘅郵箱", "Choose a mailbox to convert", "選擇要轉換嘅 PST", "Choose a PST to convert");
        lblSub.Text = D("郵件、資料夾、已讀／未讀、聯絡人同日曆轉入一個 Outlook PST；加密郵件需要郵箱主人嘅 Notes ID。",
                        "Mail, folders, read/unread state, contacts and calendar into one Outlook PST; encrypted mail needs the owner's Notes ID.",
                        "郵件、資料夾、已讀／未讀、聯絡人、日曆（包括重複會議）同工作，轉入一個新嘅 Notes 郵箱 (NSF)。",
                        "Mail, folders, read/unread state, contacts, calendar (including repeating meetings) and tasks, into a new Notes mailbox (NSF).");
        lblUpText.Text = D("將 .nsf 檔案拖放到呢度", "Drop an .nsf file here", "將 .pst 檔案拖放到呢度", "Drop a .pst file here");
        lblUpHint.Text = D("建議使用郵箱副本　·　本機或者郵件伺服器上嘅檔案都得", "Use a copy of the mailbox · local files and mail server paths both work",
                           "建議使用 PST 副本　·　轉換時唔好喺 Outlook 開住佢", "Use a copy of the PST · do not have it open in Outlook while converting");
        txtNsf.PlaceholderText = D("或者直接貼上 NSF 路徑", "...or paste the NSF path", "或者直接貼上 PST 路徑", "...or paste the PST path");
        lblOutLabel.Text = D("輸出 PST", "Output PST", "輸出 NSF", "Output NSF");
        txtPst.PlaceholderText = D("新 PST 檔案位置（揀咗 NSF 會自動填寫）", "New PST file (filled in when you pick the NSF)",
                                   "新 NSF 檔案位置（揀咗 PST 會自動填寫）", "New NSF file (filled in when you pick the PST)");
        hintServer.Text = D("NSF 喺伺服器上先填；留空即係本機檔案。", "Only if the NSF is on a server; leave empty for a local file.",
                            "要喺伺服器上建立 NSF 先填；留空即係本機檔案。", "Only to create the NSF on a server; leave empty for a local file.");
        hintUser.Text = D("留空會自動用郵箱主人嘅已讀記錄。", "Leave empty to use the mailbox owner's read marks.",
                          "未讀標記會寫俾呢個用戶；留空用而家登入 Notes 嘅用戶。", "Unread marks are written for this user; leave empty for the current Notes user.");
        lblFoot.Text = D("轉換期間 Notes 會改動原本嘅 NSF，所以建議用副本。　Enter：開始　·　Ctrl+O：揀檔案",
                         "Notes changes the NSF while converting, so work on a copy.    Enter: start  ·  Ctrl+O: open",
                         "新 NSF 會用 Notes 標準郵箱範本建立；建議用 PST 副本。　Enter：開始　·　Ctrl+O：揀檔案",
                         "The new NSF is created from the standard Notes mail template; work on a copy of the PST.    Enter: start  ·  Ctrl+O: open");
        btnOpenOut.Text = D("開啟 PST 位置", "Show PST in folder", "開啟 NSF 位置", "Show NSF in folder"); btnOpenOut.Fit();
        foreach (var c in nsfOnly) c.Visible = !Rev;
        wordmark?.Invalidate(); stepper.Invalidate();
        layoutDrop();
        NsfChanged(); UpdateAdvSummary();
    }

    void ToggleAdvanced()
    {
        advCard.Visible = !advCard.Visible;
        advHead.Open = advCard.Visible;
    }

    void BuildRun()
    {
        viewRun = Column();
        lblRunTitle.Text = L("轉換中", "Converting"); lblRunTitle.Font = Th.Ui(Th.Heading, FontStyle.Bold); lblRunTitle.ForeColor = Th.Text; lblRunTitle.AutoSize = true; lblRunTitle.BackColor = Th.Back;
        viewRun.Controls.Add(lblRunTitle);
        lblRunFile.Font = Th.Ui(Th.Body); lblRunFile.ForeColor = Th.Muted; lblRunFile.AutoSize = true; lblRunFile.BackColor = Th.Back; lblRunFile.Margin = new Padding(0, 0, 0, S(18));
        viewRun.Controls.Add(lblRunFile);

        var card = new Card { Dock = DockStyle.Top, Height = S(264), Padding = new Padding(S(28), S(24), S(28), S(24)), Margin = new Padding(0, 0, 0, S(16)) };
        stepper.Dock = DockStyle.Top; stepper.Height = S(70);
        lblPercent.Text = "0%"; lblPercent.Font = Th.Num(40f); lblPercent.ForeColor = Th.Text; lblPercent.AutoSize = true; lblPercent.BackColor = Th.Surface;
        progress.Dock = DockStyle.Top; progress.Height = S(10);
        var detail = new Label { Name = "lblDetail", AutoSize = true, Font = Th.Ui(Th.Body), ForeColor = Th.Muted, BackColor = Th.Surface };
        lblElapsed.AutoSize = true; lblElapsed.Font = Th.Ui(Th.Body); lblElapsed.ForeColor = Th.Muted; lblElapsed.BackColor = Th.Surface;
        var mid = new Panel { Dock = DockStyle.Top, Height = S(92), BackColor = Th.Surface };
        mid.Controls.AddRange(new Control[] { lblPercent, lblElapsed });
        mid.Resize += (_, _) => { lblPercent.Location = new Point(0, S(14)); lblElapsed.Location = new Point(mid.Width - lblElapsed.Width, S(52)); };
        lblElapsed.SizeChanged += (_, _) => lblElapsed.Location = new Point(mid.Width - lblElapsed.Width, S(52));
        var below = new Panel { Dock = DockStyle.Top, Height = S(40), BackColor = Th.Surface, Padding = new Padding(0, S(12), 0, 0) };
        below.Controls.Add(detail); detail.Location = new Point(0, S(14));
        progress.ValueChanged += () =>
        {
            lblPercent.Text = $"{progress.Value / 10}%";
            progress.AccessibleName = L("進度 ", "Progress ") + lblPercent.Text;
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
        lblDoneTitle.AutoSize = true; lblDoneTitle.Font = Th.Ui(Th.Heading, FontStyle.Bold); lblDoneTitle.ForeColor = Th.Text; lblDoneTitle.Location = new Point(S(96), S(24));
        lblDoneSub.AutoSize = true; lblDoneSub.Font = Th.Ui(Th.Body); lblDoneSub.ForeColor = Th.Muted; lblDoneSub.Location = new Point(S(97), S(60));
        hero.Controls.AddRange(new Control[] { lblDoneIcon, lblDoneTitle, lblDoneSub });
        // long reasons wrap inside the card, and the card grows with them
        void FitSub()
        {
            lblDoneSub.MaximumSize = new Size(Math.Max(S(200), hero.Width - lblDoneSub.Left - S(24)), 0);
            hero.Height = Math.Max(S(112), lblDoneSub.Bottom + S(22));
        }
        hero.Resize += (_, _) => { if (lblDoneSub.MaximumSize.Width != Math.Max(S(200), hero.Width - lblDoneSub.Left - S(24))) FitSub(); };
        lblDoneSub.TextChanged += (_, _) => FitSub();
        viewDone.Controls.Add(hero);

        // the actions come right after the result, before the details (the log panel can cover the bottom of the window)
        var actions = new FlowLayoutPanel { AutoSize = true, Dock = DockStyle.Top, WrapContents = true, BackColor = Th.Back, Margin = new Padding(0, 0, 0, S(6)) };
        foreach (var b in new[] { btnOpenOut, btnSummary, btnOpenLog, btnAgain }) { b.Margin = new Padding(0, 0, S(8), S(8)); actions.Controls.Add(b); }
        viewDone.Controls.Add(actions);

        tileRow.Dock = DockStyle.Top; tileRow.Height = S(96); tileRow.ColumnCount = 5; tileRow.BackColor = Th.Back; tileRow.Margin = new Padding(0, 0, 0, S(16));
        foreach (var (key, label, glyph) in new[] { ("mail", L("郵件", "Mail"), Ico.Mail), ("contact", L("聯絡人", "Contacts"), Ico.Contact), ("cal", L("日曆項目", "Calendar"), Ico.Calendar),
                                                   ("unread", L("未讀", "Unread"), Ico.Unread), ("fail", L("失敗", "Failed"), Ico.Warning) })
        {
            tileRow.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 20));
            var t = new Card { Dock = DockStyle.Fill, Margin = new Padding(key == "mail" ? 0 : S(6), 0, key == "fail" ? 0 : S(6), 0), Padding = new Padding(S(16), S(12), S(8), S(8)) };
            var ic = new Label { Text = glyph, Font = Ico.Font(11f), ForeColor = Th.Accent, AutoSize = true, BackColor = Th.Surface, Location = new Point(S(16), S(14)) };
            var cap = new Label { Text = label, Font = Th.Ui(Th.Small), ForeColor = Th.Muted, AutoSize = true, BackColor = Th.Surface, Location = new Point(S(38), S(13)) };
            var v = new Label { Text = "–", Font = Th.Num(19f), ForeColor = Th.Text, AutoSize = true, BackColor = Th.Surface, Location = new Point(S(14), S(38)) };
            t.Controls.AddRange(new Control[] { ic, cap, v });
            tiles[key] = v; tileCards[key] = t;
            tileRow.Controls.Add(t);
        }
        viewDone.Controls.Add(tileRow);

        var tableCard = new Card { Dock = DockStyle.Top, Height = S(300), Padding = new Padding(S(20), S(16), S(20), S(16)), Margin = new Padding(0, 0, 0, S(16)), Name = "tableCard" };
        var tTitle = new Label { Text = L("資料夾核對", "Folder check"), Font = Th.Ui(Th.Title, FontStyle.Bold), ForeColor = Th.Text, AutoSize = true, BackColor = Th.Surface, Dock = DockStyle.Top, Padding = new Padding(0, 0, 0, S(10)) };
        StyleGrid();
        tableCard.Controls.Add(grid); tableCard.Controls.Add(tTitle);
        viewDone.Controls.Add(tableCard);
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
        grid.ColumnHeadersDefaultCellStyle = new DataGridViewCellStyle { BackColor = Th.Surface, ForeColor = Th.Muted, Font = Th.Ui(Th.Small, FontStyle.Bold), SelectionBackColor = Th.Surface, SelectionForeColor = Th.Muted, Padding = new Padding(S(4), 0, 0, 0) };
        grid.DefaultCellStyle = new DataGridViewCellStyle { BackColor = Th.Surface, ForeColor = Th.Text, Font = Th.Ui(Th.Body), SelectionBackColor = Th.AccentSoft, SelectionForeColor = Th.Text, Padding = new Padding(S(4), 0, 0, 0) };
        grid.Columns.Add("folder", L("資料夾", "Folder")); grid.Columns.Add("exp", "NSF"); grid.Columns.Add("pst", "PST"); grid.Columns.Add("unread", L("未讀 NSF / PST", "Unread NSF / PST")); grid.Columns.Add("ok", L("一致", "Match"));
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
        lblStatus.Visible = view != viewRun;   // the run card shows the same status in large
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
        bool ok = e.Data?.GetData(DataFormats.FileDrop) is string[] f && f.Length > 0 && Path.GetExtension(f[0]).ToLowerInvariant() is ".nsf" or ".pst" && child == null;
        e.Effect = ok ? DragDropEffects.Copy : DragDropEffects.None;
        DropHighlight(ok);
    }

    void OnDragDrop(object? s, DragEventArgs e)
    {
        DropHighlight(false);
        if (e.Data?.GetData(DataFormats.FileDrop) is not string[] f || f.Length == 0) return;
        if (host.Controls.Contains(viewDone)) ShowView(viewSetup);
        bool pstFile = f[0].EndsWith(".pst", StringComparison.OrdinalIgnoreCase);
        if (pstFile != Rev) { Rev = pstFile; ApplyDirection(); }   // a .pst dropped in NSF -> PST mode switches the direction
        txtPst.Text = ""; txtNsf.Text = f[0];
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
            if (txtPst.Text.Trim().Length == 0) txtPst.Text = Path.Combine(fi.DirectoryName!, Path.GetFileNameWithoutExtension(fi.Name) + (Rev ? ".nsf" : ".pst"));
        }
        if (!envOk) SetStatus(L("缺少必要軟件，未能轉換。", "Missing required software; cannot convert."), Tone.Bad);
        else if (exists) SetStatus(L("準備好，可以開始轉換。", "Ready to convert."));
        else if (p.Length > 0 && Rev) SetStatus(L("搵唔到呢個 PST 檔案。", "PST file not found."), Tone.Warn);
        else if (p.Length > 0) SetStatus(txtServer.Text.Trim().Length > 0 ? L("會喺郵件伺服器開啟呢個郵箱。", "The mailbox will be opened on the mail server.") : L("搵唔到呢個 NSF 檔案。", "NSF file not found."),
                                         txtServer.Text.Trim().Length > 0 ? Tone.Normal : Tone.Warn);
        else SetStatus(D("請選擇 NSF 檔案。", "Choose an NSF file.", "請選擇 PST 檔案。", "Choose a PST file."));
        dropEmpty.Visible = !exists; dropFile.Visible = exists;
        dropZone.Dashed = !exists; dropZone.Height = S(exists ? 168 : 232); dropZone.Invalidate();
        btnStart.Enabled = envOk && p.Length > 0 && child == null;
    }

    void UpdateAdvSummary()
    {
        var parts = new List<string>();
        if (txtServer.Text.Trim().Length > 0) parts.Add(L("伺服器", "Server"));
        if (!Rev && txtDir.Text.Trim().Length > 0) parts.Add(L("通訊錄", "Directory"));
        if (txtUser.Text.Trim().Length > 0) parts.Add(L("已讀用戶", "Read user"));
        if (!Rev && txtId.Text.Trim().Length > 0) parts.Add("Notes ID");
        if (chkAnon.Checked) parts.Add(L("報告唔顯示資料夾名", "Report without folder names"));
        advHead.Summary = parts.Count == 0 ? L("全部選填　·　用預設值", "All optional · using defaults") : string.Join(L("　·　", " · "), parts);
        advHead.SummarySet = parts.Count > 0;
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
        Log(L($"環境檢查：Notes {(notes ? "✓" : "✗ 搵唔到（需要 64 位元 HCL Notes）")}，Outlook {(outlook ? "✓" : "✗ 搵唔到（需要 64 位元 Outlook）")}",
              $"Environment: Notes {(notes ? "✓" : "✗ not found (64-bit HCL Notes required)")}, Outlook {(outlook ? "✓" : "✗ not found (64-bit Outlook required)")}"));
        if (!envOk)
        {
            var missing = string.Join(L("、", ", "), new[] { notes ? null : L("64 位元 HCL Notes", "64-bit HCL Notes"), outlook ? null : L("64 位元 Outlook", "64-bit Outlook") }.Where(x => x != null));
            lblEnvWarn.Text = L($"缺少必要軟件：{missing}。裝好之後再開過呢個程式。", $"Missing required software: {missing}. Install it, then reopen this program.");
            lblEnvWarn.Visible = true;
            SetStatus(L("缺少必要軟件，未能轉換。", "Missing required software; cannot convert."), Tone.Bad);
        }
        NsfChanged();
    }

    void PickNsf()
    {
        using var d = Rev
            ? new OpenFileDialog { Filter = L("Outlook 資料檔 (*.pst)|*.pst|所有檔案|*.*", "Outlook data file (*.pst)|*.pst|All files|*.*"), Title = L("選擇 PST（建議用副本）", "Choose a PST (preferably a copy)") }
            : new OpenFileDialog { Filter = L("Notes 資料庫 (*.nsf)|*.nsf|所有檔案|*.*", "Notes database (*.nsf)|*.nsf|All files|*.*"), Title = L("選擇 NSF（建議用副本）", "Choose an NSF (preferably a copy)") };
        if (d.ShowDialog(this) != DialogResult.OK) return;
        txtPst.Text = "";
        txtNsf.Text = d.FileName;
    }

    void PickPst()
    {
        using var d = new SaveFileDialog
        {
            Filter = Rev ? L("Notes 資料庫 (*.nsf)|*.nsf", "Notes database (*.nsf)|*.nsf") : L("Outlook 資料檔 (*.pst)|*.pst", "Outlook data file (*.pst)|*.pst"),
            Title = D("輸出 PST", "Output PST", "輸出 NSF", "Output NSF"), OverwritePrompt = false, FileName = Path.GetFileName(txtPst.Text),
        };
        if (d.ShowDialog(this) == DialogResult.OK) txtPst.Text = d.FileName;
    }

    // =====================================================================================================
    // conversion
    // =====================================================================================================
    async Task RunAsync()
    {
        // src = the file chosen in the drop zone, dst = the new output file (NSF -> PST: nsf / pst; PST -> NSF: pst / nsf)
        string src = txtNsf.Text.Trim().Trim('"'), dst = txtPst.Text.Trim().Trim('"'), server = txtServer.Text.Trim(), user = txtUser.Text.Trim(), id = Rev ? "" : txtId.Text.Trim();
        string dstExt = Rev ? ".nsf" : ".pst";
        bool dstOnServer = Rev && server.Length > 0;   // PST -> NSF can create the NSF on a Domino server
        if (id.Length > 0 && !File.Exists(id)) { MessageBox.Show(this, L("搵唔到 Notes ID 檔案。", "Notes ID file not found."), Text); return; }
        if (src.Length == 0 || dst.Length == 0) { MessageBox.Show(this, D("請填寫 NSF 同輸出 PST。", "Enter the NSF and the output PST.", "請填寫 PST 同輸出 NSF。", "Enter the PST and the output NSF."), Text); return; }
        if ((Rev || server.Length == 0) && !File.Exists(src)) { MessageBox.Show(this, D("搵唔到 NSF 檔案。", "NSF file not found.", "搵唔到 PST 檔案。", "PST file not found."), Text); return; }
        if (!dstOnServer) dst = Path.GetFullPath(dst);
        if (!dst.EndsWith(dstExt, StringComparison.OrdinalIgnoreCase)) dst += dstExt;
        if (!dstOnServer && File.Exists(dst))
        {
            MessageBox.Show(this, D("輸出 PST 已經存在，請用新檔名（再寫入同一個 PST 會令信件重複）。",
                                    "The output PST already exists. Use a new file name (writing into the same PST again would duplicate the mail).",
                                    "輸出 NSF 已經存在，請用新檔名（每次都會建立一個新嘅 NSF）。",
                                    "The output NSF already exists. Use a new file name (a new NSF is always created)."), Text);
            return;
        }
        // Notes cannot open a database whose path is longer than about 255 bytes (a Chinese character takes 3)
        string nsfPath = Rev ? dst : src;
        if ((Rev ? !dstOnServer : server.Length == 0) && Encoding.UTF8.GetByteCount(nsfPath) > NsfExport.NotesApi.MaxPathBytes)
        {
            MessageBox.Show(this, L($"NSF 嘅路徑太長，Notes 開唔到（上限大約 {NsfExport.NotesApi.MaxPathBytes} 位元組，每個中文字佔 3 個）。請將檔案放去較短嘅資料夾或者用較短嘅檔名。\n\n{nsfPath}",
                                    $"The NSF path is too long for Notes (about {NsfExport.NotesApi.MaxPathBytes} bytes; a Chinese character takes 3). Use a shorter folder or file name.\n\n{nsfPath}"), Text,
                            MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }
        txtPst.Text = dst;
        string workBase = dstOnServer ? Path.Combine(Path.GetDirectoryName(Path.GetFullPath(src))!, Path.GetFileNameWithoutExtension(dst)) : Path.ChangeExtension(dst, null);
        workDir = workBase + "_work";
        // PST -> NSF: the EML files of an earlier run must not mix in, so an existing work folder gets a new name (never deleted)
        if (Rev && Directory.Exists(workDir)) workDir = workBase + "_work_" + DateTime.Now.ToString("yyyyMMdd_HHmmss");
        try
        {
            Directory.CreateDirectory(workDir);
            string probe = Path.Combine(workDir, ".write_test");   // the work folder and the output go to the same place
            File.WriteAllText(probe, ""); File.Delete(probe);
        }
        catch (Exception ex) when (ex is UnauthorizedAccessException or IOException)
        {
            MessageBox.Show(this, L($"冇權限寫入輸出位置，請揀另一個資料夾。\n\n{ex.Message}", $"The output location cannot be written to. Choose another folder.\n\n{ex.Message}"), Text,
                            MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }
        string nsf = Rev ? dst : src, pst = Rev ? src : dst;

        SetRunning(true);
        txtLog.Clear();
        lastFatal = null; unreadFailed = false; writeFails.Clear(); report = null; folderRows = new();
        foreach (var t in tiles.Values) { t.Text = "–"; t.ForeColor = Th.Text; }
        lblRunTitle.Text = L("轉換中", "Converting");
        lblRunFile.Text = $"{Path.GetFileName(src)}  →  {Path.GetFileName(dst)}";
        stepper.Current = 0;
        runStart = DateTime.Now; UpdateElapsed(); clock.Start();
        ShowView(viewRun);
        Taskbar.Set(Handle, 0, 1000);
        bool finished = false, allOk = false, stopped = false;
        string failText = "";
        try
        {
            Log(L("開始：", "Start: ") + $"{src} → {dst}");
            if (Rev) { (finished, allOk, stopped, failText) = await RunPstToNsf(src, dst, server, user); return; }
            var exportArgs = new List<string> { "export", nsf, workDir };
            if (server.Length > 0) exportArgs.AddRange(new[] { "--server", server });
            if (user.Length > 0) exportArgs.AddRange(new[] { "--user", user });
            if (id.Length > 0) exportArgs.AddRange(new[] { "--id", id });
            if (txtDir.Text.Trim().Length > 0) exportArgs.AddRange(new[] { "--directory", txtDir.Text.Trim() });
            SetStatus(L("第 1 步：由 Notes 匯出信件…", "Step 1: exporting mail from Notes…"));
            int rc1 = await RunChild(exportArgs, 0, 700, password: id.Length > 0 ? txtPwd.Text : null);
            if (stopRequested) { SetStatus(L("已停止。", "Stopped."), Tone.Warn); stopped = true; return; }
            if (rc1 != 0 && rc1 != 1) { SetStatus(L("匯出失敗。", "Export failed."), Tone.Bad); failText = L("匯出失敗", "export failed"); return; }   // 1 = some mails failed, still continue

            stepper.Current = 1;
            SetStatus(L("第 2 步：寫入 PST…", "Step 2: writing the PST…"));
            int rc2 = await RunChild(new List<string> { "import", pst, Path.Combine(workDir, "manifest.jsonl") }, 700, 950, fails: writeFails);
            if (stopRequested) { SetStatus(L("已停止。", "Stopped."), Tone.Warn); stopped = true; return; }
            if (rc2 != 0 && rc2 != 1) { SetStatus(L("寫入 PST 失敗。", "Writing the PST failed."), Tone.Bad); failText = L("寫入 PST 失敗", "writing the PST failed"); return; }

            stepper.Current = 2;
            SetStatus(L("第 3 步：核對數量…", "Step 3: verifying…"));
            var counts = new List<string>();
            await RunChild(new List<string> { "counts", pst }, 950, 1000, counts);
            allOk = FinishSummary(nsf, pst, counts);
            finished = true;
        }
        catch (Exception ex)
        {
            Log(L("錯誤：", "Error: ") + ex.Message);
            SetStatus(L("出錯。", "Error."), Tone.Bad);
            failText = L("出錯", "error");
        }
        finally
        {
            clock.Stop();
            // a stopped or failed run leaves a half-written output file: rename it so that it is not taken for a
            // complete mailbox (it is kept for checking; the output did not exist before this run)
            string? incomplete = !finished && !dstOnServer ? await MoveIncomplete(dst) : null;
            SetRunning(false);
            ShowDone(finished, allOk, stopped, failText, dst, incomplete);
        }
    }

    async Task<string?> MoveIncomplete(string path)
    {
        if (!File.Exists(path)) return null;
        string target = Path.Combine(Path.GetDirectoryName(path)!,
            Path.GetFileNameWithoutExtension(path) + "_incomplete_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + Path.GetExtension(path));
        for (int i = 0; i < 10; i++)   // the stopped child may hold the file for a moment
        {
            try { File.Move(path, target); Log(L("未完成嘅輸出檔已改名：", "The incomplete output file was renamed: ") + target); return target; }
            catch (IOException) { await Task.Delay(500); }
            catch (UnauthorizedAccessException) { await Task.Delay(500); }
        }
        Log(L("警告：未完成嘅輸出檔改唔到名，請唔好使用：", "Warning: the incomplete output file could not be renamed, do not use it: ") + path);
        return null;
    }

    /// <summary>PST -> NSF: read the PST (Outlook side), write a new NSF (Notes side), set unread marks, verify.</summary>
    async Task<(bool finished, bool allOk, bool stopped, string failText)> RunPstToNsf(string pst, string nsf, string server, string user)
    {
        var srvArgs = server.Length > 0 ? new[] { "--server", server } : Array.Empty<string>();
        SetStatus(L("第 1 步：由 PST 讀取信件…", "Step 1: reading mail from the PST…"));
        int rc1 = await RunChild(new List<string> { "pst-export", pst, workDir }, 0, 450);
        if (stopRequested) { SetStatus(L("已停止。", "Stopped."), Tone.Warn); return (false, false, true, ""); }
        if (rc1 != 0 && rc1 != 1) { SetStatus(L("讀取 PST 失敗。", "Reading the PST failed."), Tone.Bad); return (false, false, false, L("讀取 PST 失敗", "reading the PST failed")); }

        stepper.Current = 1;
        SetStatus(L("第 2 步：寫入 NSF…", "Step 2: writing the NSF…"));
        var importArgs = new List<string> { "nsf-import", nsf, Path.Combine(workDir, "manifest.jsonl") };
        importArgs.AddRange(srvArgs);
        if (user.Length > 0) importArgs.AddRange(new[] { "--user", user });
        int rc2 = await RunChild(importArgs, 450, 900, fails: writeFails);
        if (stopRequested) { SetStatus(L("已停止。", "Stopped."), Tone.Warn); return (false, false, true, ""); }
        if (rc2 != 0 && rc2 != 1) { SetStatus(L("寫入 NSF 失敗。", "Writing the NSF failed."), Tone.Bad); return (false, false, false, L("寫入 NSF 失敗", "writing the NSF failed")); }
        // unread marks: a separate C-API-only process (see NsfImportTool.MarkUnread)
        var markArgs = new List<string> { "nsf-mark-unread", nsf, Path.Combine(workDir, "unread_ids.txt") };
        markArgs.AddRange(srvArgs);
        int rc3 = await RunChild(markArgs, 900, 950);
        unreadFailed = rc3 != 0;
        if (unreadFailed) Log(L("警告：未讀標記寫入失敗。", "Warning: the unread marks could not be written."));

        stepper.Current = 2;
        SetStatus(L("第 3 步：核對數量…", "Step 3: verifying…"));
        var counts = new List<string>();
        var countArgs = new List<string> { "nsf-counts", nsf };
        countArgs.AddRange(srvArgs);
        if (user.Length > 0) countArgs.AddRange(new[] { "--user", user });
        await RunChild(countArgs, 950, 1000, counts);
        bool allOk = FinishSummary(nsf, pst, counts) && !unreadFailed;
        return (true, allOk, false, "");
    }

    /// <summary>Writes summary.txt / summary.html, fills the log and the tiles; returns whether every folder matches.</summary>
    bool FinishSummary(string nsf, string pst, List<string> counts)
    {
        // the counts output may contain Notes console messages; only "folder TAB items TAB unread" lines count
        Trace("summary start");
        var summary = BuildSummary(nsf, pst, counts.Where(c => c.Contains('\t')).ToList(), (DateTime.Now - runStart).TotalSeconds, chkAnon.Checked, out bool allOk);
        File.WriteAllLines(Path.Combine(workDir, "summary.txt"), summary, new UTF8Encoding(true));
        try { File.WriteAllText(Path.Combine(workDir, "summary.html"), ReportHtml(allOk, chkAnon.Checked), new UTF8Encoding(false)); } catch (Exception ex) { Log("summary.html: " + ex.Message); }
        Log("");
        foreach (var l in summary) Log(l);
        UpdateTiles(summary);
        Log("");
        Log(L("摘要檔：", "Summary files: ") + Path.Combine(workDir, "summary.txt") + " / summary.html");
        Trace("summary written");
        progress.Value = 1000;
        Trace("progress 1000 set");
        stepper.Current = 3;
        if (allOk) SetStatus(L("完成：每個資料夾數量一致，冇失敗項目。", "Done: every folder matches, nothing failed."), Tone.Ok);
        else SetStatus(L("完成，但有項目未轉換或者數量唔一致。", "Done, but some items were not converted or counts do not match."), Tone.Warn);
        return allOk;
    }

    void ShowDone(bool finished, bool allOk, bool stopped, string failText, string pst, string? incomplete = null)
    {
        Trace("ShowDone start");
        Color tint;
        if (finished)
        {
            lblDoneIcon.Text = allOk ? Ico.Completed : Ico.Warning;
            lblDoneIcon.ForeColor = allOk ? Th.Ok : Th.Warn;
            tint = allOk ? Th.OkSoft : Th.WarnSoft;
            int failed = report?.Failed ?? 0, badFolders = folderRows.Count(r => !r.ok);
            lblDoneTitle.Text = allOk ? L("全部轉換完成", "Everything was converted")
                : failed > 0 ? L($"已轉換，但有 {failed:N0} 項未能轉換", $"Converted, but {failed:N0} item(s) could not be converted")
                : L($"已轉換，但有 {badFolders:N0} 個資料夾數量唔一致", $"Converted, but {badFolders:N0} folder(s) do not match");
            long size = File.Exists(pst) ? new FileInfo(pst).Length : 0;   // pst = the output file (an NSF in PST -> NSF mode)
            var lines = new List<string> { $"{L("用時", "Took")} {Elapsed()}　·　{(Rev ? "NSF" : "PST")} {SizeText(size)}　·　{Path.GetFileName(pst)}" };
            // the reasons of failed items, most frequent first (the full list is in the item log)
            if (report != null && failed > 0)
            {
                foreach (var (reason, count) in report.Fails.Take(3)) lines.Add($"{count:N0} × {reason}");
                if (report.Fails.Count > 3) lines.Add(L($"…其餘原因見「逐項記錄」。", "…see the item log for the other reasons."));
            }
            if (badFolders > 0) lines.Add(L("數量唔一致嘅資料夾喺下面標咗 ✗。", "Folders that do not match are marked ✗ below."));
            if (unreadFailed) lines.Add(L("未讀標記寫入失敗：所有信件會顯示為已讀。", "The unread marks could not be written: all mail will show as read."));
            lblDoneSub.Text = string.Join("\n", lines);
        }
        else
        {
            lblDoneIcon.Text = stopped ? Ico.Warning : Ico.ErrorBadge;   // the stop glyph (an empty square) read as an unticked box
            lblDoneIcon.ForeColor = stopped ? Th.Warn : Th.Bad;
            tint = stopped ? Th.WarnSoft : Th.BadSoft;
            lblDoneTitle.Text = stopped ? L("轉換已停止", "Conversion stopped") : L($"轉換未完成：{failText}", $"Conversion not finished: {failText}");
            var lines = new List<string>();
            if (!stopped && lastFatal != null) lines.Add(FriendlyFatal(lastFatal));
            if (incomplete != null) lines.Add(L($"未完成嘅檔案已改名做 {Path.GetFileName(incomplete)}，唔好當完整郵箱使用。", $"The incomplete file was renamed to {Path.GetFileName(incomplete)} - do not use it as a complete mailbox."));
            if (lines.Count == 0) lines.Add(stopped ? L("轉換喺完成之前停咗，冇產生完整嘅輸出檔。", "The conversion was stopped before it finished; no complete output file was made.")
                                                    : L("原因見下面嘅詳細記錄。", "The details below show why."));
            lblDoneSub.Text = string.Join("\n", lines);
        }
        hero.Fill = tint; hero.Border = Color.Empty; hero.Invalidate();
        foreach (Control c in hero.Controls) c.BackColor = tint;

        grid.Columns["exp"]!.HeaderText = Rev ? "PST" : "NSF"; grid.Columns["pst"]!.HeaderText = Rev ? "NSF" : "PST";
        grid.Columns["unread"]!.HeaderText = D("未讀 NSF / PST", "Unread NSF / PST", "未讀 PST / NSF", "Unread PST / NSF");
        grid.Rows.Clear();
        foreach (var r in folderRows)
        {
            int i = grid.Rows.Add(r.name == "(default folders)" ? L("預設資料夾（收件匣、寄件備份 …）", "Default folders (Inbox, Sent ...)") : r.name,
                                  r.exp.ToString("N0"), r.pst.ToString("N0"), $"{r.expU:N0} / {r.pstU:N0}", r.ok ? Ico.CheckMark : Ico.Cancel);
            grid.Rows[i].Cells["ok"].Style.ForeColor = r.ok ? Th.Ok : Th.Bad;
        }
        var tableCard = viewDone.Controls["tableCard"]!;
        tableCard.Height = S(66) + grid.ColumnHeadersHeight + Math.Min(Math.Max(folderRows.Count, 1), 10) * grid.RowTemplate.Height;
        // a run that did not finish has no numbers: show only what can be acted on (no empty tiles, no dead buttons)
        tableCard.Visible = tileRow.Visible = finished;
        btnOpenOut.Visible = btnSummary.Visible = finished;
        btnOpenOut.Enabled = finished && File.Exists(pst);   // an NSF created on a server has no local file to show
        btnOpenLog.Visible = File.Exists(Path.Combine(workDir, "export_log.csv"));
        ShowView(viewDone);
        Trace("done view shown");
        if (!finished && !logPanel.Visible) ToggleLog();

        // taskbar: full green bar on success, red / yellow otherwise; flash the button if the window is in the background
        if (finished && allOk) Taskbar.Clear(Handle); else Taskbar.Set(Handle, 1000, 1000, finished || stopped ? Taskbar.Paused : Taskbar.Error);
        Trace("ShowDone taskbar set");
        if (Form.ActiveForm != this) Native.Flash(Handle);
        Trace("ShowDone end");
    }

    /// <summary>Runs this exe with the given arguments, streams its output to the log and maps PROGRESS lines to the bar.</summary>
    Task<int> RunChild(List<string> args, int from, int to, List<string>? capture = null, string? password = null, List<string>? fails = null)
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
        // progress arrives every few items: only the latest value is shown, at most 10 times a second (one repaint instead
        // of one per line); the other lines are rare and go straight to the log
        (int done, int total)? latest = null, shown = null;
        var gate = new object();
        var pump = new System.Windows.Forms.Timer { Interval = 100 };
        void ShowProgress()
        {
            (int done, int total)? now; lock (gate) now = latest;
            if (now is not { } v || now == shown) return;
            shown = now;
            progress.Value = Math.Min(1000, from + (int)((long)(to - from) * v.done / v.total));
            lblStatus.Text = $"{statusBase}{L("（", " (")}{v.done:N0} / {v.total:N0}{L("）", ")")}";
        }
        pump.Tick += (_, _) => ShowProgress();
        pump.Start();
        void OnLine(string? line)
        {
            if (line == null) return;
            if (IsNoise(line)) return;
            if (line.StartsWith("PROGRESS "))
            {
                var parts = line.Split(' ');
                if (parts.Length == 3 && int.TryParse(parts[1], out int done) && int.TryParse(parts[2], out int total) && total > 0)
                    lock (gate) latest = (done, total);
                return;
            }
            BeginInvoke(() =>
            {
                if (line.StartsWith("fatal: ")) { lastFatal = line[7..]; Log("  " + line); }
                else if (fails != null && line.StartsWith("FAILED ")) { fails.Add(line); Log("  " + line); }
                else if (capture != null) capture.Add(line);
                else Log("  " + line);
            });
        }
        p.OutputDataReceived += (_, e) => OnLine(e.Data);
        p.ErrorDataReceived += (_, e) => OnLine(e.Data);
        p.Exited += (_, _) =>
        {
            Trace($"child {args[0]} exited");
            p.WaitForExit();   // all output read
            Trace($"child {args[0]} output closed");
            BeginInvoke(() => { Trace($"child {args[0]} done on UI"); pump.Dispose(); ShowProgress(); child = null; int rc = p.ExitCode; p.Dispose(); tcs.TrySetResult(rc); });
        };
        child = p;
        Trace($"child {args[0]} start");
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
        btnStart.Enabled = !running && envOk && txtNsf.Text.Trim().Length > 0;
        btnNsf.Enabled = btnPst.Enabled = btnId.Enabled = btnDir.Enabled = !running;
        txtNsf.ReadOnly = txtPst.ReadOnly = txtServer.ReadOnly = txtUser.ReadOnly = txtId.ReadOnly = txtPwd.ReadOnly = txtDir.ReadOnly = running;
        btnStop.Enabled = running;
        if (running) progress.Value = 0;
    }

    void Log(string s) { txtLog.AppendText(s + Environment.NewLine); }

    // NSF2PST_TRACE (tests only): timestamps of the run steps, to find where the window waits
    static readonly string? tracePath = Environment.GetEnvironmentVariable("NSF2PST_TRACE");
    static readonly object traceGate = new();
    static void Trace(string what)
    {
        if (tracePath == null) return;
        lock (traceGate)   // the taskbar thread and the output readers write too
            try { File.AppendAllText(tracePath, $"{DateTime.Now:HH:mm:ss.fff} [{Environment.CurrentManagedThreadId}] {what}{Environment.NewLine}"); } catch { }
    }

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
            if (j.RootElement.TryGetProperty("type", out var ty) && ty.GetString() == "folder") continue;   // folder list, not an item
            string f = j.RootElement.GetProperty("folder").GetString() ?? "";
            if (f.StartsWith('@')) f = "(default folders)";
            bool unread = j.RootElement.TryGetProperty("read", out var r) && !r.GetBoolean();
            // PST -> NSF: each modified occurrence of a repeating appointment becomes an entry of its own
            int items = 1 + (j.RootElement.TryGetProperty("modified", out var mods) && mods.ValueKind == JsonValueKind.Array ? mods.GetArrayLength() : 0);
            var e = expected.GetValueOrDefault(f);
            expected[f] = (e.total + items, e.unread + (unread ? 1 : 0));
        }
        // PST -> NSF renames folders that clash with the mail template ("Calendar" -> "Calendar (Outlook)"): compare by the PST name
        var renames = new List<(string from, string to)>();
        string renameFile = Path.Combine(workDir, "folder_renames.txt");
        if (Rev && File.Exists(renameFile))
            foreach (var l in File.ReadAllLines(renameFile, Encoding.UTF8)) if (l.Split('\t') is { Length: 2 } p) renames.Add((p[0], p[1]));
        // an exact rename first (a sub folder can be renamed again, e.g. shortened), else the longest renamed parent
        string Original(string name)
        {
            foreach (var (from, to) in renames)
                if (name.Equals(to, StringComparison.OrdinalIgnoreCase)) return from;
            foreach (var (from, to) in renames.OrderByDescending(x => x.to.Length))
                if (name.StartsWith(to + "\\", StringComparison.OrdinalIgnoreCase)) return from + name[to.Length..];
            return name;
        }
        var actual = new Dictionary<string, (int total, int unread)>();
        foreach (var row in counts)
        {
            var c = row.Split('\t');
            if (c.Length < 3 || !int.TryParse(c[1], out int t) || !int.TryParse(c[2], out int u)) continue;
            string name = Original(c[0]);
            string f = expected.ContainsKey(name) ? name : "(default folders)";
            var a = actual.GetValueOrDefault(f);
            actual[f] = (a.total + t, a.unread + u);
        }

        var log = ReadLog();
        int Count(Func<string[], bool> pred) => log.Count(pred);
        // items read fine but not written (second step): only on its console, as "FAILED line N (what): reason"
        var writeReasons = writeFails.Select(l => Regex.Match(l, @"^FAILED line \d+(?: \([^)]*\))?: (.*)$") is { Success: true } m ? m.Groups[1].Value : l).ToList();
        int failedAll = Count(IsFail) + writeReasons.Count;
        var failGroups = log.Where(IsFail).Select(FailReason).Concat(writeReasons).GroupBy(x => x).Select(g => (g.Key, g.Count())).OrderByDescending(x => x.Item2).ToList();
        string statsFile = Path.Combine(workDir, "address_stats.txt");
        int addrOk = 0, addrBad = 0;
        if (File.Exists(statsFile) && File.ReadAllText(statsFile).Split('\t') is { Length: 3 } ast) { int.TryParse(ast[0], out addrOk); int.TryParse(ast[1], out addrBad); }
        report = new Stats(log.Count, Count(l => l[3].StartsWith("exported")), Count(l => l[3] == "contact" || l[3].StartsWith("group (")), Count(l => l[3].StartsWith("appointment")),
            Count(l => l[3].Contains("calendar (tentative)")), Count(l => l[3] == "exported (encrypted)"), Count(l => l[3].StartsWith("skipped")), failedAll,
            Count(l => l[5] == "yes"), seconds, File.Exists(nsf) ? new FileInfo(nsf).Length : 0, new FileInfo(pst).Length, addrOk, addrBad, failGroups);
        var outp = new List<string>
        {
            $"{(Rev ? "PST -> NSF" : "NSF -> PST")} summary  {DateTime.Now:yyyy-MM-dd HH:mm}  (numbers only)",
            $"time {seconds:N0}s | NSF size {(File.Exists(nsf) ? new FileInfo(nsf).Length : 0):N0} | PST size {new FileInfo(pst).Length:N0}",
            $"documents {log.Count} | mails exported {Count(l => l[3].StartsWith("exported"))} | contacts {Count(l => l[3] == "contact" || l[3].StartsWith("group ("))} | calendar entries {Count(l => l[3].StartsWith("appointment"))} | tasks {Count(l => l[3] == "task")} | groups {Count(l => l[3].StartsWith("group ("))} | invitations in calendar {Count(l => l[3].Contains("calendar (tentative)"))} | encrypted ok {Count(l => l[3] == "exported (encrypted)")} | skipped (not mail) {Count(l => l[3].StartsWith("skipped"))} | failed {failedAll} (encrypted, no key {Count(l => l[3].Contains("encrypted -"))}) | body rebuilt {Count(l => l[4] == "yes")} | unread {Count(l => l[5] == "yes")}",
        };
        string stats = Path.Combine(workDir, "address_stats.txt");
        if (File.Exists(stats) && File.ReadAllText(stats).Split('\t') is { Length: 3 } st)
            outp.Add(Rev ? $"addresses: {st[0]} Exchange address(es) -> SMTP"
                         : $"addresses: {st[0]} Notes name(s) -> SMTP | {st[1]} not resolvable ({st[2]} distinct, list in unresolved_names.txt)");
        var forms = log.Where(l => l[3].StartsWith("skipped")).GroupBy(l => l[2]).Select(g => $"{(g.Key == "" ? "(none)" : g.Key)}={g.Count()}").ToList();
        if (forms.Count > 0) outp.Add("skipped forms: " + string.Join(", ", forms));
        // failure reasons are raw error texts (may contain names or paths): left out when the report is meant for sharing
        if (failGroups.Count > 0) outp.Add(anonymize ? $"failure reasons: {failedAll} item(s), details in the logs of the work folder"
                                                     : "failure reasons: " + string.Join(" ; ", failGroups.Select(g => $"{g.Item2} x {g.Item1}")));
        outp.Add("");
        outp.Add($"{"folder",-40}{"expected",10}{(Rev ? "in NSF" : "in PST"),10}{"unread",10}{(Rev ? "unreadNSF" : "unreadPST"),10}  check");
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
        // complete = every folder matches AND nothing failed (failed items never reach the manifest, so the folder
        // check alone cannot see them)
        int failed = report.Failed;
        outp.Add(bad == 0 && failed == 0 ? "RESULT: complete - all folders match, nothing failed"
                 : $"RESULT: NOT complete - {bad} folder(s) do not match, {failed} item(s) failed");
        allOk = bad == 0 && failed == 0;
        return outp;
    }

    // =====================================================================================================
    // visual report (summary.html): the numbers of summary.txt as one self-contained page, no scripts, no external files
    // =====================================================================================================
    const string ReportCss = """
        :root{--bg:#f5f6fa;--card:#fff;--line:#e4e6ee;--text:#1c1f2a;--muted:#6b7280;--accent:#5b5bd6;--soft:#eeeefc;
              --ok:#129654;--oksoft:#eaf8f0;--warn:#d68c08;--warnsoft:#fef6e6;--bad:#d63434;--badsoft:#fdecec}
        @media (prefers-color-scheme:dark){:root{--bg:#13141a;--card:#1d1f28;--line:#323542;--text:#eaecf2;--muted:#969cac;--accent:#8480ff;
              --soft:#2d2c50;--ok:#40c880;--oksoft:#182e24;--warn:#f0b034;--warnsoft:#342a16;--bad:#f26060;--badsoft:#3c1e22}}
        *{box-sizing:border-box}
        body{margin:0;background:var(--bg);color:var(--text);font:15px/1.5 "Segoe UI","Microsoft JhengHei UI","PingFang TC",sans-serif}
        .wrap{max-width:980px;margin:0 auto;padding:32px 20px 48px}
        header{display:flex;align-items:center;gap:14px}
        .logo{width:46px;height:46px;border-radius:13px;background:linear-gradient(135deg,#8b5cf6,#4f46e5);display:grid;place-items:center;
              box-shadow:0 6px 18px rgba(91,91,214,.35)}
        .brand{font-size:24px;font-weight:700;letter-spacing:.2px}.brand b{color:var(--accent);font-weight:700;margin:0 4px}
        .sub{color:var(--muted);font-size:13px}
        .hero{display:flex;align-items:center;gap:18px;border-radius:18px;padding:22px 26px;margin:24px 0 16px}
        .hero.ok{background:var(--oksoft)}.hero.warn{background:var(--warnsoft)}
        .hero .mark{width:54px;height:54px;border-radius:50%;display:grid;place-items:center;border:3px solid currentColor;font-size:26px;font-weight:700;flex:none}
        .hero.ok .mark{color:var(--ok)}.hero.warn .mark{color:var(--warn)}
        .hero h1{margin:0;font-size:22px}.hero p{margin:2px 0 0;color:var(--muted)}
        .kpis{display:grid;grid-template-columns:repeat(auto-fit,minmax(140px,1fr));gap:12px}
        .card{background:var(--card);border:1px solid var(--line);border-radius:16px;padding:16px 18px}
        .kpi .l{color:var(--muted);font-size:13px}.kpi .v{font-size:30px;font-weight:700;line-height:1.2}.kpi.bad{background:var(--badsoft)}.kpi.bad .v{color:var(--bad)}
        .grid2{display:grid;grid-template-columns:1.15fr 1fr;gap:16px;margin-top:16px}.full{margin-top:16px}
        @media (max-width:780px){.grid2{grid-template-columns:1fr}}
        h2{font-size:15px;margin:0 0 12px}
        .donut{display:flex;gap:18px;align-items:center}
        .legend{list-style:none;margin:0;padding:0;font-size:13px}.legend li{display:flex;align-items:center;gap:8px;margin:4px 0}
        .legend i{width:10px;height:10px;border-radius:3px;display:inline-block}.legend span{margin-left:auto;font-weight:600;padding-left:12px}
        .meter{height:12px;border-radius:7px;background:var(--badsoft);overflow:hidden;margin:10px 0 6px}.meter i{display:block;height:100%;background:linear-gradient(90deg,#2dd4bf,#0d9488)}
        .facts{display:grid;grid-template-columns:1fr auto;gap:4px 12px;font-size:13px;color:var(--muted)}.facts b{color:var(--text);text-align:right}
        .row{display:grid;grid-template-columns:minmax(160px,2fr) 3fr 100px 100px 26px;gap:12px;align-items:center;padding:9px 0;border-top:1px solid var(--line);font-size:14px}
        .row.h{border-top:none;color:var(--muted);font-size:12px;font-weight:600;padding-top:0}
        .row .n{text-align:right;font-variant-numeric:tabular-nums}.name{overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
        .bar{position:relative;height:10px;border-radius:6px;background:var(--soft)}
        .bar .nsf{position:absolute;inset:0 auto 0 0;border-radius:6px;border:1.5px dashed var(--muted);opacity:.55}
        .bar .pst{position:absolute;inset:0 auto 0 0;border-radius:6px;background:linear-gradient(90deg,#8b5cf6,#5b5bd6)}
        .ok{color:var(--ok)}.badc{color:var(--bad)}
        .reasons{margin:0;padding-left:18px;font-size:13px}.reasons li{margin:4px 0}
        footer{margin-top:22px;color:var(--muted);font-size:12px;text-align:center}
        @media print{body{background:#fff}.card{break-inside:avoid}}
        """;

    /// <summary>Builds summary.html from the last run (numbers only; folder names follow the "hide folder names" option).</summary>
    string ReportHtml(bool allOk, bool anonymize)
    {
        var s = report!;
        static string E(string x) => System.Net.WebUtility.HtmlEncode(x);
        int bad = folderRows.Count(r => !r.ok);
        var h = new StringBuilder();
        h.Append($"<!doctype html><html lang=\"{(En ? "en" : "zh-Hant")}\"><head><meta charset=\"utf-8\"><meta name=\"viewport\" content=\"width=device-width,initial-scale=1\">");
        h.Append($"<title>Nsf2Pst · {E(L("轉換報告", "Conversion report"))}</title><style>{ReportCss}</style></head><body><div class=\"wrap\">");
        h.Append("<header><div class=\"logo\"><svg viewBox=\"0 0 24 24\" width=\"24\" height=\"24\" fill=\"none\" stroke=\"#fff\" stroke-width=\"1.8\" stroke-linecap=\"round\" stroke-linejoin=\"round\">");
        h.Append("<rect x=\"3\" y=\"5.5\" width=\"18\" height=\"13\" rx=\"2.5\"/><path d=\"M4 7.5l8 5.5 8-5.5\"/></svg></div>");
        h.Append($"<div><div class=\"brand\">Nsf<b>&rarr;</b>Pst</div><div class=\"sub\">{E(anonymize ? L("轉換報告 · 只有數量（資料夾名同錯誤詳情已隱藏）", "Conversion report · counts only (folder names and error details hidden)") : L("轉換報告 · 數量、資料夾名同錯誤原因", "Conversion report · counts, folder names and error reasons"))} · {DateTime.Now:yyyy-MM-dd HH:mm}</div></div></header>");

        string took = Duration(TimeSpan.FromSeconds(s.Seconds));
        h.Append($"<section class=\"hero {(allOk ? "ok" : "warn")}\"><div class=\"mark\">{(allOk ? "&#10003;" : "!")}</div><div>");
        h.Append($"<h1>{E(allOk ? L("全部轉換完成", "Everything was converted") : s.Failed > 0 ? L($"有 {s.Failed:N0} 項未能轉換", $"{s.Failed:N0} item(s) could not be converted") : L($"{bad} 個資料夾數量唔一致", $"{bad} folder(s) do not match"))}</h1>");
        h.Append($"<p>{E(L("用時", "Took"))} {took} · {(Rev ? $"PST {SizeText(s.PstSize)} → NSF {SizeText(s.NsfSize)}" : $"NSF {SizeText(s.NsfSize)} → PST {SizeText(s.PstSize)}")} · {folderRows.Count} {E(L("個資料夾", "folders"))}</p></div></section>");

        void Kpi(string label, int v, bool warnWhenNonZero = false) =>
            h.Append($"<div class=\"card kpi{(warnWhenNonZero && v > 0 ? " bad" : "")}\"><div class=\"l\">{E(label)}</div><div class=\"v\">{v:N0}</div></div>");
        h.Append("<section class=\"kpis\">");
        Kpi(L("郵件", "Mail"), s.Mails); Kpi(L("聯絡人", "Contacts"), s.Contacts); Kpi(L("日曆項目", "Calendar"), s.Calendar + s.Invites);
        Kpi(L("未讀", "Unread"), s.Unread); if (!Rev) Kpi(L("加密郵件已解密", "Encrypted, decrypted"), s.Encrypted); Kpi(L("失敗", "Failed"), s.Failed, true);
        h.Append("</section>");

        // donut: what the documents in the NSF were
        var segs = new (string label, int v, string color)[]
        {
            (L("郵件", "Mail"), s.Mails, "#6366f1"), (L("聯絡人", "Contacts"), s.Contacts, "#14b8a6"), (L("日曆", "Calendar"), s.Calendar, "#f59e0b"),
            (L("略過（非郵件）", "Skipped (not mail)"), s.Skipped, "#94a3b8"), (L("失敗", "Failed"), s.Failed, "#ef4444"),
        };
        int total = Math.Max(1, segs.Sum(x => x.v));
        const double R = 54, C = 2 * Math.PI * R;
        var inv = CultureInfo.InvariantCulture;
        h.Append("<section class=\"grid2\"><div class=\"card\"><h2>").Append(E(D("NSF 入面有咩", "What was in the NSF", "PST 入面有咩", "What was in the PST"))).Append("</h2><div class=\"donut\">");
        h.Append("<svg viewBox=\"0 0 140 140\" width=\"140\" height=\"140\"><circle cx=\"70\" cy=\"70\" r=\"54\" fill=\"none\" stroke=\"var(--soft)\" stroke-width=\"18\"/>");
        double off = 0;
        foreach (var (_, v, color) in segs.Where(x => x.v > 0))
        {
            double len = C * v / total;
            h.Append(string.Format(inv, "<circle cx=\"70\" cy=\"70\" r=\"54\" fill=\"none\" stroke=\"{0}\" stroke-width=\"18\" stroke-dasharray=\"{1:F2} {2:F2}\" stroke-dashoffset=\"{3:F2}\" transform=\"rotate(-90 70 70)\"/>",
                color, Math.Max(0, len - 1.5), C - Math.Max(0, len - 1.5), -off));
            off += len;
        }
        h.Append($"<text x=\"70\" y=\"68\" text-anchor=\"middle\" font-size=\"24\" font-weight=\"700\" fill=\"var(--text)\">{s.Docs:N0}</text>");
        h.Append($"<text x=\"70\" y=\"88\" text-anchor=\"middle\" font-size=\"11\" fill=\"var(--muted)\">{E(L("份文件", "documents"))}</text></svg><ul class=\"legend\">");
        foreach (var (label, v, color) in segs) h.Append($"<li><i style=\"background:{color}\"></i>{E(label)}<span>{v:N0}</span></li>");
        h.Append("</ul></div>");
        if (s.Invites > 0) h.Append($"<p class=\"sub\">{E(L($"另外 {s.Invites:N0} 封未回覆邀請已加入日曆（暫定）。", $"{s.Invites:N0} unanswered invitation(s) were also added to the calendar as tentative."))}</p>");
        h.Append("</div>");

        // run facts + address resolution
        h.Append($"<div class=\"card\"><h2>{E(L("轉換資料", "Run details"))}</h2><div class=\"facts\">");
        void Fact(string k, string v) => h.Append($"<span>{E(k)}</span><b>{E(v)}</b>");
        Fact(L("用時", "Time"), took); Fact(L("NSF 大小", "NSF size"), SizeText(s.NsfSize)); Fact(L("PST 大小", "PST size"), SizeText(s.PstSize)); Fact(L("方向", "Direction"), Rev ? "PST → NSF" : "NSF → PST");
        Fact(L("資料夾", "Folders"), folderRows.Count.ToString("N0")); Fact(L("未讀", "Unread"), s.Unread.ToString("N0"));
        if (!Rev) Fact(L("加密郵件已解密", "Encrypted mail decrypted"), s.Encrypted.ToString("N0"));
        h.Append("</div>");
        if (s.AddrOk + s.AddrBad > 0)
        {
            int pct = (int)Math.Round(100.0 * s.AddrOk / (s.AddrOk + s.AddrBad));
            h.Append($"<h2 style=\"margin-top:18px\">{E(L("Notes 名轉 email 地址", "Notes names to email addresses"))}</h2><div class=\"meter\"><i style=\"width:{pct}%\"></i></div>");
            h.Append($"<div class=\"facts\"><span>{E(L("已轉換", "Resolved"))}</span><b>{s.AddrOk:N0} ({pct}%)</b><span>{E(L("轉換唔到", "Not resolvable"))}</span><b>{s.AddrBad:N0}</b></div>");
        }
        h.Append("</div></section>");

        // folders: NSF (dashed outline) vs PST (filled), unread, check
        int max = Math.Max(1, folderRows.Count == 0 ? 1 : folderRows.Max(r => Math.Max(r.exp, r.pst)));
        h.Append($"<section class=\"card full\"><h2>{E(L("資料夾核對", "Folder check"))}</h2>");
        h.Append($"<div class=\"row h\"><span>{E(L("資料夾", "Folder"))}</span><span>{E(D("數量（虛線 NSF · 實色 PST）", "Items (dashed NSF, filled PST)", "數量（虛線 PST · 實色 NSF）", "Items (dashed PST, filled NSF)"))}</span><span class=\"n\">{(Rev ? "PST / NSF" : "NSF / PST")}</span><span class=\"n\">{E(L("未讀", "Unread"))}</span><span></span></div>");
        foreach (var r in folderRows)
        {
            string name = r.name == "(default folders)" ? L("預設資料夾（收件匣、寄件備份 …）", "Default folders (Inbox, Sent ...)") : r.name;
            h.Append(string.Format(inv, "<div class=\"row\"><span class=\"name\" title=\"{0}\">{0}</span><div class=\"bar\"><div class=\"nsf\" style=\"width:{1:F1}%\"></div><div class=\"pst\" style=\"width:{2:F1}%\"></div></div>",
                E(name), 100.0 * r.exp / max, 100.0 * r.pst / max));
            h.Append($"<span class=\"n\">{r.exp:N0} / {r.pst:N0}</span><span class=\"n\">{r.expU:N0} / {r.pstU:N0}</span><span class=\"{(r.ok ? "ok" : "badc")}\">{(r.ok ? "&#10003;" : "&#10007;")}</span></div>");
        }
        h.Append("</section>");

        if (s.Fails.Count > 0)
        {
            h.Append($"<section class=\"card\" style=\"margin-top:16px\"><h2>{E(L("失敗原因", "Failure reasons"))}</h2><ul class=\"reasons\">");
            if (anonymize) h.Append($"<li><b>{s.Failed:N0}</b> × {E(L("詳情見 export_log.csv（呢份報告冇列出）", "details in export_log.csv (not listed in this report)"))}</li>");
            else foreach (var (reason, count) in s.Fails) h.Append($"<li><b>{count:N0}</b> × {E(reason)}</li>");
            h.Append("</ul></section>");
        }
        h.Append($"<footer>Nsf2Pst · {E(anonymize ? L("呢份報告只有數量，冇郵件內容、資料夾名或者錯誤詳情。", "This report has counts only: no mail content, folder names or error details.") : L("呢份報告有數量、資料夾名同錯誤原因，冇郵件內容。", "This report has counts, folder names and error reasons, but no mail content."))}</footer></div></body></html>");
        return h.ToString();
    }

    // a row failed completely ("FAILED: reason") or partly (an entry was made but e.g. its attachments failed: "... FAILED ...")
    static bool IsFail(string[] l) => l[3].Contains("FAILED");
    static string FailReason(string[] l) => l[3].StartsWith("FAILED: ") ? l[3][8..] : l[3];

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
        // NSF2PST_SETTINGS (tests only): another settings file, so automated runs never touch the user's own choice
        static string FilePath => Environment.GetEnvironmentVariable("NSF2PST_SETTINGS") is { Length: > 0 } t ? t
            : Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Nsf2Pst", "settings.txt");
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
        public static Color Ok, OkSoft, Warn, WarnSoft, Bad, BadSoft, LogBack, LogText, Disabled, DisabledText, OnAccent;
        static Color C(int r, int g, int b) => Color.FromArgb(r, g, b);
        // every text colour is >= 4.5:1 on the surfaces it is drawn on (Back, Surface and the matching *Soft tint)
        public static void Apply(bool dark)
        {
            Dark = dark;
            if (!dark)
            {
                Back = C(245, 246, 250); Surface = Color.White; Line = C(228, 230, 238); Text = C(28, 31, 42); Muted = C(95, 102, 117);
                Accent = C(91, 91, 214); AccentDark = C(70, 70, 190); AccentSoft = C(238, 238, 252); AccentLine = C(190, 190, 236); Accent2 = C(124, 92, 246); DashLine = C(178, 180, 230);
                Ok = C(11, 122, 67); OkSoft = C(234, 248, 240); Warn = C(161, 92, 0); WarnSoft = C(254, 246, 230); Bad = C(180, 35, 35); BadSoft = C(253, 236, 236);
                Disabled = C(228, 230, 238); DisabledText = C(95, 102, 117); OnAccent = Color.White;
            }
            else
            {
                Back = C(19, 20, 26); Surface = C(29, 31, 40); Line = C(50, 53, 66); Text = C(234, 236, 242); Muted = C(150, 156, 172);
                Accent = C(132, 128, 255); AccentDark = C(150, 146, 255); AccentSoft = C(45, 44, 80); AccentLine = C(96, 94, 170); Accent2 = C(160, 120, 255); DashLine = C(88, 88, 150);
                Ok = C(64, 200, 128); OkSoft = C(24, 46, 36); Warn = C(240, 176, 52); WarnSoft = C(52, 42, 22); Bad = C(242, 96, 96); BadSoft = C(60, 30, 34);
                Disabled = C(56, 58, 82); DisabledText = C(176, 179, 202); OnAccent = C(19, 20, 26);
            }
            LogBack = C(14, 15, 20); LogText = C(206, 212, 224);
        }
        // type scale: 9 (captions), 10 (body), 13 (titles), 17 (page heading); numbers in the display face
        public const float Small = 9f, Body = 10f, Title = 13f, Heading = 17f;
        public static Font Ui(float size, FontStyle style = FontStyle.Regular) => new(En ? "Segoe UI" : "Microsoft JhengHei UI", size, style);
        static readonly string numFamily = FontFamily.Families.Any(f => f.Name == "Segoe UI Variable Display") ? "Segoe UI Variable Display" : "Segoe UI";
        public static Font Num(float size) => new(numFamily, size, FontStyle.Bold);
    }

    /// <summary>Glyphs of the Windows icon font (Segoe Fluent Icons on Windows 11, Segoe MDL2 Assets on Windows 10).</summary>
    static class Ico
    {
        public const string Mail = "\uE715", Upload = "\uE898", Settings = "\uE713", ChevronDown = "\uE70D", ChevronUp = "\uE70E", Forward = "\uE72A",
            Stop = "\uE71A", Folder = "\uE8B7", Document = "\uE8A5", Refresh = "\uE72C", Contact = "\uE77B", Calendar = "\uE787",
            Unread = "\uE715", Warning = "\uE7BA", Completed = "\uE930", ErrorBadge = "\uEA39", CheckMark = "\uE73E", Cancel = "\uE711",
            Globe = "\uE774", Moon = "\uE708", Sun = "\uE706", Chart = "\uE9D2";
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
        public int Radius = 10;
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

    /// <summary>The keyboard focus indicator of the custom-drawn controls: a 2 px accent ring just inside the edge.</summary>
    static void FocusRing(Graphics g, RectangleF r, float radius, Color? color = null)
    {
        g.SmoothingMode = SmoothingMode.AntiAlias;
        var inner = RectangleF.Inflate(r, -S(2), -S(2));
        using var path = Round(inner, Math.Max(0, radius - S(2)));
        using var pen = new Pen(color ?? Th.Accent, S(2));
        g.DrawPath(pen, path);
    }

    /// <summary>Rounded button: filled accent (primary) or outline (secondary); text, icon or both.</summary>
    sealed class Pill : Button
    {
        readonly bool primary, small, iconAfter;
        readonly string? icon;
        Font? iconFont;
        bool hover, down;
        public Pill(string text, bool primary = false, bool small = false, string? icon = null, bool iconAfter = false)
        {
            this.primary = primary; this.small = small; this.icon = icon; this.iconAfter = iconAfter;
            Text = text;
            SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw, true);
            FlatStyle = FlatStyle.Flat; FlatAppearance.BorderSize = 0; Cursor = Cursors.Hand;
            Font = Th.Ui(small ? Th.Small : Th.Body, primary ? FontStyle.Bold : FontStyle.Regular);
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
        protected override void OnGotFocus(EventArgs e) { Invalidate(); base.OnGotFocus(e); }
        protected override void OnLostFocus(EventArgs e) { Invalidate(); base.OnLostFocus(e); }
        protected override void OnPaint(PaintEventArgs e)
        {
            Draw(e.Graphics);
            if (Focused && ShowFocusCues && Enabled) FocusRing(e.Graphics, new RectangleF(0.5f, 0.5f, Width - 1.5f, Height - 1.5f), S(10), primary ? Th.OnAccent : Th.Accent);
        }
        void Draw(Graphics g)
        {
            g.Clear(Parent?.BackColor ?? Th.Back);
            g.SmoothingMode = SmoothingMode.AntiAlias;
            var r = new RectangleF(0.5f, 0.5f, Width - 1.5f, Height - 1.5f);
            using var path = Round(r, S(10));
            Color fill = !Enabled ? (primary ? Th.Disabled : Th.Back) : primary ? (down || hover ? Th.AccentDark : Th.Accent) : (down ? Th.Line : hover ? Th.AccentSoft : Th.Surface);
            using (var b = new SolidBrush(fill)) g.FillPath(b, path);
            if (!primary) { using var pen = new Pen(Enabled && hover ? Th.AccentLine : Th.Line); g.DrawPath(pen, path); }
            Color fg = primary ? (Enabled ? Th.OnAccent : Th.DisabledText) : Enabled ? Th.Text : Th.Muted;
            var tsize = TextRenderer.MeasureText(g, Text, Font, Size.Empty, TextFormatFlags.NoPadding);
            if (icon == null) { TextRenderer.DrawText(g, Text, Font, ClientRectangle, fg, TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter | TextFormatFlags.NoPadding); return; }
            var iconFont = this.iconFont ??= Ico.Font(Font.Size * 0.95f);   // created once, not on every repaint
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

    /// <summary>
    /// The "Advanced settings" header: one real button (Tab, Enter / Space, UI Automation Invoke) drawn as a card with a
    /// gear, the title, a summary of what is filled in and a chevron.
    /// </summary>
    sealed class Expander : Button
    {
        string summary = "";
        bool open, summarySet, hover;
        Font? icons, small, chev;   // made on the first paint, kept
        public string Summary { get => summary; set { summary = value; Invalidate(); } }
        public bool SummarySet { get => summarySet; set { summarySet = value; Invalidate(); } }
        public bool Open { get => open; set { open = value; AccessibleDescription = value ? L("已展開", "Expanded") : L("已收起", "Collapsed"); Invalidate(); } }
        public Expander()
        {
            SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw, true);
            FlatStyle = FlatStyle.Flat; FlatAppearance.BorderSize = 0; Cursor = Cursors.Hand; Font = Th.Ui(Th.Title, FontStyle.Bold);
            Open = false;
        }
        protected override void OnMouseEnter(EventArgs e) { hover = true; Invalidate(); base.OnMouseEnter(e); }
        protected override void OnMouseLeave(EventArgs e) { hover = false; Invalidate(); base.OnMouseLeave(e); }
        protected override void OnGotFocus(EventArgs e) { Invalidate(); base.OnGotFocus(e); }
        protected override void OnLostFocus(EventArgs e) { Invalidate(); base.OnLostFocus(e); }
        protected override void OnPaint(PaintEventArgs e)
        {
            var g = e.Graphics;
            g.Clear(Parent?.BackColor ?? Th.Back);
            g.SmoothingMode = SmoothingMode.AntiAlias;
            var r = new RectangleF(0.5f, 0.5f, Width - 1.5f, Height - 1.5f);
            using (var path = Round(r, S(10)))
            {
                using (var b = new SolidBrush(hover ? Th.AccentSoft : Th.Surface)) g.FillPath(b, path);
                using var pen = new Pen(hover ? Th.AccentLine : Th.Line); g.DrawPath(pen, path);
            }
            var icons = this.icons ??= Ico.Font(Th.Title);
            var small = this.small ??= Th.Ui(Th.Body);
            var flags = TextFormatFlags.VerticalCenter | TextFormatFlags.NoPadding | TextFormatFlags.SingleLine;
            int x = S(16);
            var gs = TextRenderer.MeasureText(g, Ico.Settings, icons, Size.Empty, TextFormatFlags.NoPadding);
            TextRenderer.DrawText(g, Ico.Settings, icons, new Rectangle(x, 0, gs.Width, Height), Th.Muted, flags);
            x += gs.Width + S(12);
            var ts = TextRenderer.MeasureText(g, Text, Font, Size.Empty, TextFormatFlags.NoPadding);
            TextRenderer.DrawText(g, Text, Font, new Rectangle(x, 0, ts.Width, Height), Th.Text, flags);
            var chev = this.chev ??= Ico.Font(Th.Body);
            int cx = Width - S(16) - S(16);
            TextRenderer.DrawText(g, open ? Ico.ChevronUp : Ico.ChevronDown, chev, new Rectangle(cx, 0, S(16), Height), Th.Muted, flags | TextFormatFlags.HorizontalCenter);
            int sx = x + ts.Width + S(16), sw = cx - S(12) - sx;
            if (sw > S(40))
                TextRenderer.DrawText(g, summary, small, new Rectangle(sx, 0, sw, Height), summarySet ? Th.Accent : Th.Muted, flags | TextFormatFlags.Right | TextFormatFlags.EndEllipsis);
            if (Focused && ShowFocusCues) FocusRing(g, r, S(10));
        }
    }

    /// <summary>Two-option segmented switch (conversion direction). The options are real buttons, so UI Automation can press them.</summary>
    sealed class Segmented : Panel
    {
        readonly Seg[] segs;
        int index;
        public event Action? Changed;
        public int Index { get => index; set { index = value; foreach (var s in segs) s.Invalidate(); } }
        public Segmented(params string[] items)
        {
            DoubleBuffered = true; ResizeRedraw = true; BackColor = Th.Surface;
            segs = items.Select((t, i) => new Seg(this, i, t) { Name = "segOption" + i }).ToArray();
            Controls.AddRange(segs);
            Fit();
        }
        /// <summary>Size from the DPI; called again once the window knows its DPI.</summary>
        public void Fit()
        {
            Size = new Size(S(330), S(44)); Padding = new Padding(S(4));
            int w = (Width - S(8)) / segs.Length;
            for (int i = 0; i < segs.Length; i++) segs[i].Bounds = new Rectangle(S(4) + i * w, S(4), w, Height - S(8));
        }
        void Select(int i) { if (i == index || !Enabled) return; Index = i; Changed?.Invoke(); }
        protected override void OnPaintBackground(PaintEventArgs e)
        {
            var g = e.Graphics;
            g.Clear(Parent?.BackColor ?? Th.Back);
            g.SmoothingMode = SmoothingMode.AntiAlias;
            using var path = Round(new RectangleF(0.5f, 0.5f, Width - 1.5f, Height - 1.5f), S(10));
            using (var b = new SolidBrush(Th.Surface)) g.FillPath(b, path);
            using var pen = new Pen(Th.Line); g.DrawPath(pen, path);
        }

        sealed class Seg : Button
        {
            readonly Segmented owner; readonly int i; bool hover;
            public Seg(Segmented owner, int i, string text)
            {
                this.owner = owner; this.i = i; Text = text;
                SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw, true);
                FlatStyle = FlatStyle.Flat; FlatAppearance.BorderSize = 0; Cursor = Cursors.Hand; Font = Th.Ui(Th.Body, FontStyle.Bold);
            }
            protected override void OnClick(EventArgs e) { owner.Select(i); base.OnClick(e); }
            // arrow keys move between the options, like a radio group
            protected override bool IsInputKey(Keys k) => k is Keys.Left or Keys.Right || base.IsInputKey(k);
            protected override void OnKeyDown(KeyEventArgs e)
            {
                if (e.KeyCode is Keys.Left or Keys.Right)
                {
                    int j = Math.Clamp(i + (e.KeyCode == Keys.Right ? 1 : -1), 0, owner.segs.Length - 1);
                    owner.Select(j); owner.segs[j].Focus(); e.Handled = true;
                }
                base.OnKeyDown(e);
            }
            protected override void OnGotFocus(EventArgs e) { Invalidate(); base.OnGotFocus(e); }
            protected override void OnLostFocus(EventArgs e) { Invalidate(); base.OnLostFocus(e); }
            protected override void OnMouseEnter(EventArgs e) { hover = true; Invalidate(); base.OnMouseEnter(e); }
            protected override void OnMouseLeave(EventArgs e) { hover = false; Invalidate(); base.OnMouseLeave(e); }
            protected override void OnPaint(PaintEventArgs e)
            {
                var g = e.Graphics;
                g.Clear(Th.Surface);
                g.SmoothingMode = SmoothingMode.AntiAlias;
                bool on = owner.Index == i;
                if (on || hover)
                {
                    using var path = Round(new RectangleF(0, 0, Width - 1, Height - 1), S(8));
                    using var b = on ? (Brush)new LinearGradientBrush(new RectangleF(0, 0, Width, Height), Th.Accent2, Th.Accent, LinearGradientMode.Horizontal)
                                     : new SolidBrush(Th.AccentSoft);
                    g.FillPath(b, path);
                }
                Color fg = on ? Th.OnAccent : owner.Enabled ? Th.Text : Th.Muted;
                TextRenderer.DrawText(g, Text, Font, ClientRectangle, fg, TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter);
                if (Focused && ShowFocusCues) FocusRing(g, new RectangleF(0, 0, Width - 1, Height - 1), S(8), on ? Th.OnAccent : Th.Accent);
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
        Font? numFont, lblFont, lblBold, icon;   // made on the first paint, kept
        public int Current { get => current; set { current = value; Invalidate(); } }
        public Stepper() { SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw, true); }
        protected override void OnPaint(PaintEventArgs e)
        {
            var g = e.Graphics;
            g.Clear(Parent?.BackColor ?? Th.Surface);
            g.SmoothingMode = SmoothingMode.AntiAlias;
            string[] names = Rev ? new[] { L("讀取 PST", "Read PST"), L("寫入 NSF", "Write NSF"), L("核對結果", "Verify") }
                                 : new[] { L("匯出郵件", "Export mail"), L("寫入 PST", "Write PST"), L("核對結果", "Verify") };
            int n = names.Length, d = S(28), top = S(4);
            float slot = (Width - 1f) / n;
            var numFont = this.numFont ??= Th.Ui(Th.Body, FontStyle.Bold);
            var lblFont = this.lblFont ??= Th.Ui(Th.Body);
            var lblBold = this.lblBold ??= Th.Ui(Th.Body, FontStyle.Bold);
            var icon = this.icon ??= Ico.Font(10f);
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
                if (done) TextRenderer.DrawText(g, Ico.CheckMark, icon, cr, Th.OnAccent, TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter);
                else TextRenderer.DrawText(g, (i + 1).ToString(), numFont, cr, active ? Th.OnAccent : Th.Muted, TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter);
                var lr = new Rectangle((int)(cx - slot / 2), top + d + S(8), (int)slot, S(22));
                TextRenderer.DrawText(g, names[i], active ? lblBold : lblFont, lr, active || done ? Th.Text : Th.Muted, TextFormatFlags.HorizontalCenter);
            }
        }
    }

    /// <summary>
    /// The app logo (same design as app.ico): violet-to-indigo rounded square with a soft top highlight and a white glyph,
    /// optionally with the teal "transfer" badge in the lower right corner.
    /// </summary>
    static void DrawLogo(Graphics g, RectangleF r, string glyph, bool badge, int dpi)
    {
        g.SmoothingMode = SmoothingMode.AntiAlias;
        using var path = Round(r, r.Width / 3.6f);
        using (var lg = new LinearGradientBrush(r, Color.FromArgb(139, 92, 246), Color.FromArgb(79, 70, 229), LinearGradientMode.ForwardDiagonal)) g.FillPath(lg, path);
        var top = new RectangleF(r.X, r.Y, r.Width, r.Height * 0.55f);
        using (var hl = new LinearGradientBrush(top, Color.FromArgb(46, Color.White), Color.FromArgb(0, Color.White), LinearGradientMode.Vertical))
        {
            var clip = g.Clip; g.SetClip(path); g.FillRectangle(hl, top); g.Clip = clip;
        }
        float shift = badge ? r.Width * 0.06f : 0;
        var gr = new RectangleF(r.X - shift, r.Y - shift, r.Width, r.Height);
        using (var f = Ico.Font(r.Height * 0.38f * 72f / dpi))
            TextRenderer.DrawText(g, glyph, f, Rectangle.Round(gr), Color.White, TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter | TextFormatFlags.NoPadding);
        if (!badge) return;
        float d = r.Width * 0.44f, cx = r.Right - d * 0.62f, cy = r.Bottom - d * 0.62f;
        var ring = new RectangleF(cx - d / 2, cy - d / 2, d, d);
        using (var white = new SolidBrush(Color.White)) g.FillEllipse(white, ring);
        var dot = RectangleF.Inflate(ring, -d * 0.11f, -d * 0.11f);
        using (var teal = new LinearGradientBrush(dot, Color.FromArgb(45, 212, 191), Color.FromArgb(13, 148, 136), LinearGradientMode.ForwardDiagonal)) g.FillEllipse(teal, dot);
        using var af = Ico.Font(d * 0.42f * 72f / dpi);
        TextRenderer.DrawText(g, Ico.Forward, af, Rectangle.Round(dot), Color.White, TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter | TextFormatFlags.NoPadding);
    }

    /// <summary>The logo as a control (header) or a plain glyph tile (chosen file).</summary>
    sealed class IconBadge : Control
    {
        readonly string glyph;
        readonly bool badge;
        public IconBadge(string glyph, bool badge = false) { this.glyph = glyph; this.badge = badge; SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw | ControlStyles.SupportsTransparentBackColor, true); }
        protected override void OnPaint(PaintEventArgs e)
        {
            e.Graphics.Clear(Parent?.BackColor ?? Th.Surface);
            DrawLogo(e.Graphics, new RectangleF(0, 0, Width - 1, Height - 1), glyph, badge, DeviceDpi);
        }
    }

    /// <summary>"Nsf → Pst" word mark ("Pst → Nsf" in the other direction): the arrow in the accent colour.</summary>
    sealed class Wordmark : Control
    {
        readonly Font font = Th.Num(15f);
        readonly Font arrow;
        public Wordmark()
        {
            SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw, true);
            arrow = Ico.Font(12f);
            var a = TextRenderer.MeasureText("Nsf", font, Size.Empty, TextFormatFlags.NoPadding);
            var b = TextRenderer.MeasureText(Ico.Forward, arrow, Size.Empty, TextFormatFlags.NoPadding);
            var c = TextRenderer.MeasureText("Pst", font, Size.Empty, TextFormatFlags.NoPadding);
            Size = new Size(a.Width + b.Width + c.Width + S(14), Math.Max(a.Height, c.Height) + S(2));
            AccessibleName = "Nsf2Pst";
        }
        protected override void OnPaint(PaintEventArgs e)
        {
            var g = e.Graphics;
            g.Clear(Parent?.BackColor ?? Th.Surface);
            string first = Rev ? "Pst" : "Nsf", second = Rev ? "Nsf" : "Pst";   // the arrow follows the conversion direction
            var a = TextRenderer.MeasureText(g, first, font, Size.Empty, TextFormatFlags.NoPadding);
            var b = TextRenderer.MeasureText(g, Ico.Forward, arrow, Size.Empty, TextFormatFlags.NoPadding);
            TextRenderer.DrawText(g, first, font, new Point(0, 0), Th.Text, TextFormatFlags.NoPadding);
            int ax = a.Width + S(6), ay = (a.Height - b.Height) / 2 + S(2);
            TextRenderer.DrawText(g, Ico.Forward, arrow, new Point(ax, ay), Th.Accent, TextFormatFlags.NoPadding);
            TextRenderer.DrawText(g, second, font, new Point(ax + b.Width + S(7), 0), Th.Text, TextFormatFlags.NoPadding);
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
        // The calls go to Explorer. They run on a thread of their own, so a busy or hanging Explorer can only delay the
        // taskbar button, never the window; only the latest value is sent.
        static readonly object gate = new();
        static (IntPtr h, int done, int total, int state)? pending;
        static readonly AutoResetEvent signal = new(false);
        static Thread? worker;
        public static void Set(IntPtr h, int done, int total, int state = Normal)
        {
            lock (gate)
            {
                pending = (h, done, total, state);
                if (worker == null)
                {
                    worker = new Thread(Loop) { IsBackground = true, Name = "taskbar" };
                    worker.SetApartmentState(ApartmentState.STA);
                    worker.Start();
                }
            }
            signal.Set();
        }
        static void Loop()
        {
            ITaskbarList3? list = null;
            while (true)
            {
                signal.WaitOne();
                (IntPtr h, int done, int total, int state)? next;
                lock (gate) { next = pending; pending = null; }
                if (next is not { } v) continue;
                try
                {
                    if (list == null) { list = (ITaskbarList3)new TaskbarInstance(); list.HrInit(); }
                    if (v.total <= 0) list.SetProgressState(v.h, NoProgress);
                    else { list.SetProgressState(v.h, v.state); list.SetProgressValue(v.h, (ulong)v.done, (ulong)v.total); }
                    if (v.total <= 0 || v.done == v.total || v.done == 0) Trace($"taskbar {v.done}/{v.total} state {v.state}");
                }
                catch (Exception ex) { list = null; Trace("taskbar failed: " + ex.Message); }   // e.g. Explorer restarted: made again for the next value
            }
        }
        public static void Clear(IntPtr h) => Set(h, 0, 0);
    }

    static class Native
    {
        [DllImport("dwmapi.dll")] static extern int DwmSetWindowAttribute(IntPtr hwnd, int attr, ref int value, int size);
        [DllImport("uxtheme.dll", CharSet = CharSet.Unicode)] static extern int SetWindowTheme(IntPtr hwnd, string? app, string? idList);
        [DllImport("user32.dll")] static extern bool FlashWindowEx(ref FLASHWINFO info);
        [DllImport("user32.dll")] static extern IntPtr MonitorFromPoint(Point pt, uint flags);
        [DllImport("shcore.dll")] static extern int GetDpiForMonitor(IntPtr monitor, int type, out uint dpiX, out uint dpiY);

        /// <summary>Effective DPI of the screen that contains the point (96 if unknown).</summary>
        public static int MonitorDpi(Point p)
        {
            try { return GetDpiForMonitor(MonitorFromPoint(p, 2), 0, out uint x, out _) == 0 && x > 0 ? (int)x : 96; }   // MONITOR_DEFAULTTONEAREST, MDT_EFFECTIVE_DPI
            catch { return 96; }
        }
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

        public static void DarkScrollbars(IntPtr hwnd) { try { SetWindowTheme(hwnd, "DarkMode_Explorer", null); } catch { } }

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
