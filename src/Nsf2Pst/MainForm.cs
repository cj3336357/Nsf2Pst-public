// Main window: pick the NSF and the PST, run export (Notes) then import (PST) as child processes of this exe,
// show progress and the verification summary (numbers only, written to <work>\summary.txt).
using System.Diagnostics;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;

namespace Nsf2Pst;

internal sealed class MainForm : Form
{
    readonly TextBox txtNsf = new(), txtPst = new(), txtServer = new(), txtUser = new(), txtLog = new(), txtId = new(), txtDir = new(), txtPwd = new() { UseSystemPasswordChar = true };
    readonly CheckBox chkAnon = new() { Text = "摘要隱藏 folder 名稱（方便對外提供）", AutoSize = true };
    readonly Button btnNsf = new() { Text = "瀏覽..." }, btnPst = new() { Text = "另存為..." }, btnId = new() { Text = "瀏覽..." }, btnDir = new() { Text = "瀏覽..." };
    readonly Button btnStart = new() { Text = "開始轉換" }, btnStop = new() { Text = "停止", Enabled = false };
    readonly Button btnOpenOut = new() { Text = "開啟輸出資料夾", Enabled = false }, btnOpenLog = new() { Text = "開啟逐封記錄", Enabled = false };
    readonly ProgressBar progress = new() { Minimum = 0, Maximum = 1000 };
    readonly Label lblStatus = new() { Text = "請選擇 NSF 檔案及輸出 PST。", AutoSize = true };
    Process? child;
    bool stopRequested;
    string workDir = "";

    // colours
    static readonly Color Accent = Color.FromArgb(0, 103, 192), AccentDark = Color.FromArgb(0, 84, 158), Header = Color.FromArgb(24, 46, 79);
    static readonly Color Back = Color.FromArgb(243, 245, 248), CardBorder = Color.FromArgb(222, 226, 232), Muted = Color.FromArgb(110, 118, 129);
    static readonly Color Ok = Color.FromArgb(16, 124, 16), Bad = Color.FromArgb(196, 43, 28), Warn = Color.FromArgb(202, 120, 0);
    readonly Label chipNotes = Chip(), chipOutlook = Chip();
    readonly Dictionary<string, Label> tiles = new();
    static float k = 1;                       // DPI scale
    static int S(int px) => (int)Math.Round(px * k);

    public MainForm()
    {
        Text = "Nsf2Pst - Notes 郵箱 (NSF) 轉 Outlook PST";
        Font = new Font("Microsoft JhengHei UI", 9.75f);
        AutoScaleMode = AutoScaleMode.None;
        k = DeviceDpi / 96f;
        BackColor = Back;
        var work = Screen.FromPoint(Cursor.Position).WorkingArea;
        MinimumSize = new Size(Math.Min(S(820), work.Width), Math.Min(S(700), work.Height));
        Size = new Size(Math.Min(S(1000), work.Width), Math.Min(S(900), work.Height));
        StartPosition = FormStartPosition.CenterScreen;

        // ---- header
        var header = new TableLayoutPanel { Dock = DockStyle.Top, AutoSize = true, ColumnCount = 2, BackColor = Header, Padding = new Padding(S(20), S(12), S(20), S(14)) };
        header.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        header.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        var titles = new FlowLayoutPanel { FlowDirection = FlowDirection.TopDown, AutoSize = true, WrapContents = false, Margin = new Padding(0) };
        titles.Controls.Add(new Label { Text = "Nsf2Pst", ForeColor = Color.White, Font = new Font("Segoe UI Semibold", 18f), AutoSize = true, Margin = new Padding(0) });
        titles.Controls.Add(new Label { Text = "Notes 郵箱 (NSF) → Outlook PST", ForeColor = Color.FromArgb(200, 214, 232), AutoSize = true, Margin = new Padding(S(2), 0, 0, 0) });
        titles.Controls.Add(new Label { Text = "郵件 · 資料夾 · 已讀/未讀 · 聯絡人 · 日曆 · 加密郵件", ForeColor = Color.FromArgb(150, 170, 198), Font = new Font(Font.FontFamily, 8.75f), AutoSize = true, Margin = new Padding(S(2), S(2), 0, 0) });
        var chips = new FlowLayoutPanel { AutoSize = true, WrapContents = false, Anchor = AnchorStyles.Right, Margin = new Padding(0) };
        chips.Controls.Add(chipNotes); chips.Controls.Add(chipOutlook);
        header.Controls.Add(titles, 0, 0); header.Controls.Add(chips, 1, 0);

        // ---- body
        var root = new TableLayoutPanel { Dock = DockStyle.Fill, Padding = new Padding(S(18), S(14), S(18), S(16)), ColumnCount = 1, BackColor = Back };
        root.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));

        var files = Grid();
        AddRow(files, 0, "NSF 檔案", txtNsf, btnNsf);
        AddRow(files, 1, "輸出 PST", txtPst, btnPst);
        root.Controls.Add(Card("轉換檔案", files));

        var opts = Grid();
        AddRow(opts, 0, "Domino server", txtServer, Hint("NSF 喺 server 上先填"));
        AddRow(opts, 1, "Domino 通訊錄", txtDir, btnDir);
        AddRow(opts, 2, "已讀/未讀用戶", txtUser, Hint("例：CN=名/O=機構"));
        AddRow(opts, 3, "主人 Notes ID", txtId, btnId);
        AddRow(opts, 4, "ID 密碼", txtPwd, Hint("加密郵件先需要"));
        chkAnon.ForeColor = Color.FromArgb(50, 56, 64); chkAnon.Margin = new Padding(S(3), S(8), S(3), 0);
        opts.Controls.Add(chkAnon, 1, 5);
        root.Controls.Add(Card("進階選項（全部選填）", opts));

        var buttons = new FlowLayoutPanel { AutoSize = true, Dock = DockStyle.Fill, WrapContents = false, Margin = new Padding(0, S(4), 0, S(8)) };
        Style(btnStart, primary: true); Style(btnStop); Style(btnOpenOut); Style(btnOpenLog);
        btnStart.Text = "▶  開始轉換"; btnStop.Text = "■  停止";
        foreach (var b in new[] { btnStart, btnStop, btnOpenOut, btnOpenLog }) buttons.Controls.Add(b);
        root.Controls.Add(buttons);

        var tileRow = new TableLayoutPanel { Dock = DockStyle.Fill, AutoSize = true, ColumnCount = 5, Margin = new Padding(0, 0, 0, S(8)) };
        foreach (var (key, label) in new[] { ("mail", "郵件"), ("contact", "聯絡人"), ("cal", "日曆項目"), ("unread", "未讀"), ("fail", "失敗") })
        {
            tileRow.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 20));
            var value = new Label { Text = "–", Font = new Font("Segoe UI Semibold", 18f), ForeColor = Header, AutoSize = true, Dock = DockStyle.Top };
            var cap = new Label { Text = label, ForeColor = Muted, AutoSize = true, Dock = DockStyle.Top };
            var tile = new BorderPanel { Dock = DockStyle.Fill, BackColor = Color.White, Padding = new Padding(S(14), S(8), S(8), S(8)), Height = S(72), Margin = new Padding(key == "mail" ? 0 : S(5), 0, key == "fail" ? 0 : S(5), 0) };
            tile.Controls.Add(value); tile.Controls.Add(cap);
            tiles[key] = value;
            tileRow.Controls.Add(tile);
        }
        root.Controls.Add(tileRow);

        progress.Dock = DockStyle.Fill; progress.Height = S(10); progress.Style = ProgressBarStyle.Continuous; progress.Margin = new Padding(0, 0, 0, S(6));
        root.Controls.Add(progress);
        lblStatus.Font = new Font(Font.FontFamily, 10.5f, FontStyle.Bold); lblStatus.Margin = new Padding(0, 0, 0, S(8));
        lblStatus.TextChanged += (_, _) => lblStatus.ForeColor =
            lblStatus.Text.StartsWith("完成 ✓") ? Ok :
            Regex.IsMatch(lblStatus.Text, "失敗|出錯|缺少") ? Bad :
            lblStatus.Text.Contains("對唔上") || lblStatus.Text.Contains("已停止") ? Warn : Color.FromArgb(40, 46, 54);
        root.Controls.Add(lblStatus);

        txtLog.Multiline = true; txtLog.ReadOnly = true; txtLog.ScrollBars = ScrollBars.Both; txtLog.WordWrap = false;
        txtLog.Font = new Font("Consolas", 9.5f); txtLog.Dock = DockStyle.Fill; txtLog.BorderStyle = BorderStyle.None;
        txtLog.BackColor = Color.FromArgb(30, 32, 36); txtLog.ForeColor = Color.FromArgb(214, 220, 228);
        var logHost = new Panel { Dock = DockStyle.Fill, BackColor = txtLog.BackColor, Padding = new Padding(S(10), S(8), S(4), S(4)), Margin = new Padding(0) };
        logHost.Controls.Add(txtLog);
        root.Controls.Add(logHost);

        for (int r = 0; r < root.Controls.Count - 1; r++) root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        Controls.Add(root);
        Controls.Add(header);

        // names = UI Automation ids (used by automated tests)
        txtNsf.Name = "txtNsf"; txtPst.Name = "txtPst"; txtServer.Name = "txtServer"; txtUser.Name = "txtUser"; txtLog.Name = "txtLog"; txtId.Name = "txtId"; txtPwd.Name = "txtPwd"; txtDir.Name = "txtDir";
        chkAnon.Name = "chkAnon"; btnStart.Name = "btnStart"; lblStatus.Name = "lblStatus";
        txtNsf.PlaceholderText = "揀 .nsf 檔案（建議用副本）";
        txtPst.PlaceholderText = "新 PST 檔案位置";
        txtUser.PlaceholderText = "留空 = 自動用 mailbox 主人";
        txtId.PlaceholderText = "留空 = 用而家登入嘅 ID（解唔到其他人嘅加密郵件）";
        txtServer.PlaceholderText = "留空 = 本機檔案";
        txtDir.PlaceholderText = "選填：server!!names.nsf 或者 names.nsf 副本，用嚟將 Notes 名轉成 email 地址";
        btnDir.Click += (_, _) => { using var d = new OpenFileDialog { Filter = "Domino 通訊錄 (*.nsf)|*.nsf|所有檔案|*.*" }; if (d.ShowDialog(this) == DialogResult.OK) txtDir.Text = d.FileName; };
        btnId.Click += (_, _) => { using var d = new OpenFileDialog { Filter = "Notes ID (*.id)|*.id|所有檔案|*.*" }; if (d.ShowDialog(this) == DialogResult.OK) txtId.Text = d.FileName; };
        btnNsf.Click += (_, _) => PickNsf();
        btnPst.Click += (_, _) => PickPst();
        btnStart.Click += async (_, _) => await RunAsync();
        btnStop.Click += (_, _) => StopChild();
        btnOpenOut.Click += (_, _) => { if (File.Exists(txtPst.Text)) Process.Start("explorer.exe", $"/select,\"{txtPst.Text}\""); };
        btnOpenLog.Click += (_, _) => OpenFile(Path.Combine(workDir, "export_log.csv"));
        FormClosing += (_, e) => { if (child != null && MessageBox.Show(this, "轉換仲進行緊，確定要停止並關閉？", Text, MessageBoxButtons.YesNo) == DialogResult.No) e.Cancel = true; else StopChild(); };
        Shown += (_, _) => CheckPrerequisites();
    }

    /// <summary>White panel with a 1px border.</summary>
    sealed class BorderPanel : Panel
    {
        public BorderPanel() { DoubleBuffered = true; ResizeRedraw = true; }
        protected override void OnPaint(PaintEventArgs e)
        {
            base.OnPaint(e);
            using var pen = new Pen(CardBorder);
            e.Graphics.DrawRectangle(pen, 0, 0, Width - 1, Height - 1);
        }
    }

    static Control Card(string title, Control content)
    {
        var card = new BorderPanel { Dock = DockStyle.Fill, AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, BackColor = Color.White, Padding = new Padding(S(16), S(10), S(16), S(12)), Margin = new Padding(0, 0, 0, S(12)) };
        content.Dock = DockStyle.Top;
        var head = new Label { Text = title, Font = new Font("Microsoft JhengHei UI", 10.5f, FontStyle.Bold), ForeColor = Header, AutoSize = true, Dock = DockStyle.Top, Padding = new Padding(0, 0, 0, S(6)) };
        card.Controls.Add(content); card.Controls.Add(head);
        return card;
    }

    static TableLayoutPanel Grid()
    {
        var g = new TableLayoutPanel { ColumnCount = 3, AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, BackColor = Color.White };
        g.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        g.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        g.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        return g;
    }

    static Label Hint(string text) => new() { Text = text, AutoSize = true, ForeColor = Muted, Anchor = AnchorStyles.Left, Margin = new Padding(S(8), 0, 0, 0) };

    static Label Chip() => new() { AutoSize = true, ForeColor = Color.White, BackColor = Color.FromArgb(60, 80, 110), Padding = new Padding(S(10), S(4), S(10), S(4)), Margin = new Padding(S(6), 0, 0, 0), Text = "檢查中 ..." };

    static void Style(Button b, bool primary = false)
    {
        b.FlatStyle = FlatStyle.Flat;
        b.AutoSize = true; b.Height = S(36);
        b.Padding = new Padding(S(primary ? 18 : 12), S(4), S(primary ? 18 : 12), S(4));
        b.Margin = new Padding(0, 0, S(8), 0);
        b.Cursor = Cursors.Hand;
        b.FlatAppearance.BorderColor = primary ? Accent : CardBorder;
        b.FlatAppearance.MouseOverBackColor = primary ? AccentDark : Color.FromArgb(232, 237, 244);
        b.BackColor = primary ? Accent : Color.White;
        b.ForeColor = primary ? Color.White : Color.FromArgb(40, 46, 54);
        if (primary) b.Font = new Font("Microsoft JhengHei UI", 10.5f, FontStyle.Bold);
        b.EnabledChanged += (_, _) => b.BackColor = !b.Enabled ? Color.FromArgb(225, 228, 233) : primary ? Accent : Color.White;
    }

    static void AddRow(TableLayoutPanel g, int row, string label, TextBox box, Control right)
    {
        g.Controls.Add(new Label { Text = label, AutoSize = true, ForeColor = Color.FromArgb(50, 56, 64), Anchor = AnchorStyles.Left, Margin = new Padding(0, 0, S(8), 0) }, 0, row);
        box.Dock = DockStyle.Fill; box.Margin = new Padding(0, S(5), 0, S(5)); box.BorderStyle = BorderStyle.FixedSingle;
        g.Controls.Add(box, 1, row);
        if (right is Button b) { Style(b); b.MinimumSize = new Size(S(120), 0); b.Anchor = AnchorStyles.Left; b.Margin = new Padding(S(8), S(2), 0, S(2)); }
        g.Controls.Add(right, 2, row);
    }

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
        tiles["fail"].ForeColor = tiles["fail"].Text is "0" or "–" ? Header : Bad;
    }

    void CheckPrerequisites()
    {
        bool notes = Microsoft.Win32.Registry.GetValue(@"HKEY_CLASSES_ROOT\CLSID\{29131539-2EED-1069-BF5D-00DD011186B7}\InprocServer32", "", null) != null;
        bool outlook = new[] { @"%ProgramFiles%\Microsoft Office\root\Office16\OLMAPI32.DLL", @"%ProgramFiles%\Microsoft Office\Office16\OLMAPI32.DLL" }
            .Any(p => File.Exists(Environment.ExpandEnvironmentVariables(p)));
        SetChip(chipNotes, "Notes", notes); SetChip(chipOutlook, "Outlook 64-bit", outlook);
        Log($"環境檢查：Notes client {(notes ? "✓" : "✗ 搵唔到（需要 64-bit HCL Notes）")}，Outlook 64-bit {(outlook ? "✓" : "✗ 搵唔到（需要 64-bit Outlook）")}");
        if (!notes || !outlook) lblStatus.Text = "缺少必要軟件，未能轉換。";
    }

    static void SetChip(Label chip, string name, bool ok)
    {
        chip.Text = (ok ? "✓ " : "✗ ") + name;
        chip.BackColor = ok ? Color.FromArgb(30, 120, 70) : Color.FromArgb(170, 40, 30);
    }

    void PickNsf()
    {
        using var d = new OpenFileDialog { Filter = "Notes 資料庫 (*.nsf)|*.nsf|所有檔案|*.*", Title = "選擇 NSF（建議用副本）" };
        if (d.ShowDialog(this) != DialogResult.OK) return;
        txtNsf.Text = d.FileName;
        if (txtPst.Text.Length == 0) txtPst.Text = Path.Combine(Path.GetDirectoryName(d.FileName)!, Path.GetFileNameWithoutExtension(d.FileName) + ".pst");
    }

    void PickPst()
    {
        using var d = new SaveFileDialog { Filter = "Outlook 資料檔 (*.pst)|*.pst", Title = "輸出 PST", OverwritePrompt = false, FileName = Path.GetFileName(txtPst.Text) };
        if (d.ShowDialog(this) == DialogResult.OK) txtPst.Text = d.FileName;
    }

    async Task RunAsync()
    {
        string nsf = txtNsf.Text.Trim(), pst = txtPst.Text.Trim(), server = txtServer.Text.Trim(), user = txtUser.Text.Trim(), id = txtId.Text.Trim();
        if (id.Length > 0 && !File.Exists(id)) { MessageBox.Show(this, "搵唔到 Notes ID 檔案。", Text); return; }
        if (server.Length == 0 && !File.Exists(nsf)) { MessageBox.Show(this, "搵唔到 NSF 檔案。", Text); return; }
        if (nsf.Length == 0 || pst.Length == 0) { MessageBox.Show(this, "請填寫 NSF 同輸出 PST。", Text); return; }
        pst = Path.GetFullPath(pst);
        if (!pst.EndsWith(".pst", StringComparison.OrdinalIgnoreCase)) pst += ".pst";
        if (File.Exists(pst)) { MessageBox.Show(this, "輸出 PST 已經存在，請用新檔名（再寫入同一個 PST 會令信件重複）。", Text); return; }
        txtPst.Text = pst;
        workDir = Path.Combine(Path.GetDirectoryName(pst)!, Path.GetFileNameWithoutExtension(pst) + "_work");
        Directory.CreateDirectory(workDir);

        SetRunning(true);
        txtLog.Clear();
        foreach (var t in tiles.Values) { t.Text = "–"; t.ForeColor = Header; }
        var t0 = DateTime.Now;
        try
        {
            Log($"開始：{nsf} → {pst}");
            var exportArgs = new List<string> { "export", nsf, workDir };
            if (server.Length > 0) exportArgs.AddRange(new[] { "--server", server });
            if (user.Length > 0) exportArgs.AddRange(new[] { "--user", user });
            if (id.Length > 0) exportArgs.AddRange(new[] { "--id", id });
            if (txtDir.Text.Trim().Length > 0) exportArgs.AddRange(new[] { "--directory", txtDir.Text.Trim() });
            lblStatus.Text = "第 1 步：由 Notes 匯出信件 ...";
            int rc1 = await RunChild(exportArgs, 0, 700, password: id.Length > 0 ? txtPwd.Text : null);
            if (stopRequested) { lblStatus.Text = "已停止。"; return; }
            if (rc1 != 0 && rc1 != 1) { lblStatus.Text = "匯出失敗，請睇記錄。"; return; }   // 1 = some mails failed, still continue

            lblStatus.Text = "第 2 步：寫入 PST ...";
            int rc2 = await RunChild(new List<string> { "import", pst, Path.Combine(workDir, "manifest.jsonl") }, 700, 950);
            if (stopRequested) { lblStatus.Text = "已停止。"; return; }
            if (rc2 != 0 && rc2 != 1) { lblStatus.Text = "寫入 PST 失敗，請睇記錄。"; return; }

            lblStatus.Text = "第 3 步：核對 ...";
            var counts = new List<string>();
            await RunChild(new List<string> { "counts", pst }, 950, 1000, counts);
            var summary = BuildSummary(nsf, pst, counts, (DateTime.Now - t0).TotalSeconds, chkAnon.Checked, out bool allOk);
            File.WriteAllLines(Path.Combine(workDir, "summary.txt"), summary, new UTF8Encoding(true));
            Log("");
            foreach (var l in summary) Log(l);
            UpdateTiles(summary);
            Log("");
            Log($"摘要檔（只有數字，可以對外提供）：{Path.Combine(workDir, "summary.txt")}");
            progress.Value = progress.Maximum;
            lblStatus.Text = allOk ? "完成 ✓ 所有 folder 數量核對正確。" : "完成，但有 folder 數量對唔上，請睇摘要。";
            btnOpenOut.Enabled = btnOpenLog.Enabled = true;
        }
        catch (Exception ex)
        {
            Log("錯誤：" + ex.Message);
            lblStatus.Text = "出錯，請睇記錄。";
        }
        finally { SetRunning(false); }
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
                        progress.Value = Math.Min(progress.Maximum, from + (int)((long)(to - from) * done / total));
                        lblStatus.Text = $"{lblStatus.Text.Split('（')[0]}（{done:N0} / {total:N0}）";
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
        btnStart.Enabled = btnNsf.Enabled = btnPst.Enabled = btnId.Enabled = btnDir.Enabled = !running;
        txtNsf.ReadOnly = txtPst.ReadOnly = txtServer.ReadOnly = txtUser.ReadOnly = txtId.ReadOnly = txtPwd.ReadOnly = txtDir.ReadOnly = running;
        btnStop.Enabled = running;
        if (running) { progress.Value = 0; btnOpenOut.Enabled = btnOpenLog.Enabled = false; }
    }

    void Log(string s) { txtLog.AppendText(s + Environment.NewLine); }

    static void OpenFile(string path) { if (File.Exists(path)) Process.Start(new ProcessStartInfo(path) { UseShellExecute = true }); }

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
        foreach (var f in expected.Keys.Union(actual.Keys).OrderBy(x => x, StringComparer.Ordinal))
        {
            n++;
            var e = expected.GetValueOrDefault(f); var a = actual.GetValueOrDefault(f);
            bool ok = e == a; if (!ok) bad++;
            string name = anonymize && f != "(default folders)" ? $"folder{n:D3}" : f;
            outp.Add($"{name,-40}{e.total,10}{a.total,10}{e.unread,10}{a.unread,10}  {(ok ? "OK" : "MISMATCH")}");
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
}
