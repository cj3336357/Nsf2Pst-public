// "import" / "list" / "counts" mode: writes EML files into a PST through Extended MAPI (64-bit Outlook).
//   Nsf2Pst import <pst> <manifest.jsonl>   one JSON object per line: {"eml","folder","subject","read","received","sent"}
//   Nsf2Pst list <pst>                       folders and messages, for checking
//   Nsf2Pst counts <pst>                     "folder TAB items TAB unread", numbers only
// Progress lines for the UI: "PROGRESS <done> <total>".
using System.Diagnostics;
using System.Text;
using System.Text.Json;
using PstLite;

namespace Nsf2Pst;

internal static class PstTool
{
    public static int Run(string[] args)
    {
        if (args.Length < 2) { Console.Error.WriteLine("usage: import <pst> <manifest.jsonl> | list <pst> | counts <pst>"); return 2; }
        string cmd = args[0].ToLowerInvariant(), pst = args[1];
        try
        {
            if (cmd == "pst-export" && !File.Exists(pst)) { Console.WriteLine("fatal: PST not found: " + pst); return 3; }
            if (cmd != "import" && File.Exists(pst) && CheckPst(pst) is string problem) { Console.WriteLine("fatal: " + problem); return 3; }
            // Outlook's PST provider writes into every PST it opens (it refuses read-only files), so PST -> NSF reads a copy:
            // the user's PST stays byte for byte unchanged
            string? copy = null;
            if (cmd == "pst-export")
            {
                string work = Path.GetFullPath(args.Length > 2 ? args[2] : Path.ChangeExtension(pst, null) + "_work");
                Directory.CreateDirectory(work);
                if (args.Length < 3) args = new[] { args[0], args[1], work };   // the work folder of the original, not of the copy
                copy = Path.Combine(work, "source_copy.pst");
                try { File.Copy(pst, copy, true); }
                catch (IOException ex) { Console.WriteLine("fatal: the PST could not be copied to the work folder (free disk space?): " + ex.Message); return 3; }
                Console.WriteLine("reading a copy of the PST (the original is not changed)");
            }
            try { return RunOn(copy ?? pst, cmd, args); }
            finally { if (copy != null) try { File.Delete(copy); } catch { } }
        }
        catch (Exception ex)
        {
            Console.WriteLine("fatal: " + ex.Message);
            return 3;
        }
    }

    static int RunOn(string pst, string cmd, string[] args)
    {
        try
        {
            using var store = new PstStore(pst);
            store.ReadOnlyNames = cmd is "pst-export" or "counts" or "list";   // reading the user's PST: add nothing to it
            if (cmd == "list") { store.List(Console.Out); return 0; }
            if (cmd == "counts") { store.Counts(Console.Out); return 0; }
            if (cmd == "pst-export") return ExportPst(store, args.Length > 2 ? args[2] : Path.ChangeExtension(pst, null) + "_work");
            if (cmd == "pst-make-test") { store.MakeTestItems(); Console.WriteLine("test items added"); return 0; }   // test helper
            if (cmd == "pst-make-uat") { store.MakeUatItems(); Console.WriteLine("UAT items added"); return 0; }     // test helper
            if (args.Length < 3) { Console.Error.WriteLine("missing manifest"); return 2; }

            string[] lines = File.ReadAllLines(args[2], Encoding.UTF8).Where(l => !string.IsNullOrWhiteSpace(l)).ToArray();
            int ok = 0, failed = 0;
            var sw = Stopwatch.StartNew();
            for (int i = 0; i < lines.Length; i++)
            {
                if (i % 10 == 0) { Console.WriteLine($"PROGRESS {i} {lines.Length}"); Console.Out.Flush(); }
                string eml = "?";
                try
                {
                    using var doc = JsonDocument.Parse(lines[i]);
                    var r = doc.RootElement;
                    string type = r.TryGetProperty("type", out var ty) ? ty.GetString() ?? "mail" : "mail";
                    if (type == "contact")
                    {
                        eml = "contact";
                        var fields = r.GetProperty("fields").EnumerateObject().ToDictionary(x => x.Name, x => x.Value.GetString() ?? "");
                        store.ImportContact(r.GetProperty("folder").GetString() ?? "@Contacts", fields);
                        ok++;
                        continue;
                    }
                    if (type == "appointment")
                    {
                        eml = "appointment";
                        store.ImportAppointment(r.GetProperty("folder").GetString() ?? "@Calendar",
                            r.GetProperty("subject").GetString() ?? "",
                            r.GetProperty("start").GetDateTimeOffset(), r.GetProperty("end").GetDateTimeOffset(),
                            r.TryGetProperty("allDay", out var ad) && ad.GetBoolean(),
                            r.TryGetProperty("location", out var lo) ? lo.GetString() : null,
                            r.TryGetProperty("body", out var bo) ? bo.GetString() : null,
                            r.TryGetProperty("reminder", out var rm) && rm.ValueKind == JsonValueKind.Number ? rm.GetInt32() : null,
                            r.TryGetProperty("tentative", out var te) && te.GetBoolean(), Attachments(r));
                        ok++;
                        continue;
                    }
                    eml = r.GetProperty("eml").GetString()!;
                    string folder = r.TryGetProperty("folder", out var f) ? f.GetString() ?? "" : "";
                    string? subject = r.TryGetProperty("subject", out var s) ? s.GetString() : null;
                    bool? read = r.TryGetProperty("read", out var rd) ? rd.GetBoolean() : null;
                    DateTimeOffset? received = r.TryGetProperty("received", out var rc) ? rc.GetDateTimeOffset() : null;
                    DateTimeOffset? sent = r.TryGetProperty("sent", out var st) ? st.GetDateTimeOffset() : null;
                    store.Import(eml, folder, subject, read, received, sent);
                    ok++;
                }
                catch (Exception ex)
                {
                    failed++;
                    Console.WriteLine($"FAILED line {i + 1} ({Path.GetFileName(eml)}): {ex.Message}");
                }
            }
            Console.WriteLine($"PROGRESS {lines.Length} {lines.Length}");
            Console.WriteLine($"imported {ok}, failed {failed}, {sw.Elapsed.TotalSeconds:F1}s");
            return failed == 0 ? 0 : 1;
        }
        catch (Exception ex)
        {
            Console.WriteLine("fatal: " + ex.Message);
            return 3;
        }
    }

    /// <summary>PST -> NSF, PST side: EML files + manifest.jsonl + export_log.csv (same columns as the NSF export log).</summary>
    static int ExportPst(PstStore store, string outDir)
    {
        outDir = Path.GetFullPath(outDir);
        string emlDir = Path.Combine(outDir, "eml");
        Directory.CreateDirectory(emlDir);
        var folders = store.AllFolders();
        int total = folders.Sum(f => f.Count);
        Console.WriteLine($"{folders.Count} folder(s), {total} item(s)");
        if (store.SkippedFolders.Count > 0) Console.WriteLine("skipped hidden Outlook folders: " + string.Join(", ", store.SkippedFolders));
        using var manifest = new StreamWriter(Path.Combine(outDir, "manifest.jsonl"), false, new UTF8Encoding(false));
        using var log = new StreamWriter(Path.Combine(outDir, "export_log.csv"), false, new UTF8Encoding(true));
        log.WriteLine("NoteID,UNID,Form,Status,BodyRebuilt,Unread,Folders,Subject");
        int seq = 0, done = 0, mails = 0, contacts = 0, groups = 0, appointments = 0, repeating = 0, tasks = 0, exFixed = 0, failed = 0;
        var sw = Stopwatch.StartNew();
        void Line(Dictionary<string, object?> d) =>
            manifest.WriteLine(JsonSerializer.Serialize(d.Where(x => x.Value != null).ToDictionary(x => x.Key, x => x.Value)));
        void Progress() { if (++done % 5 == 0) { Console.WriteLine($"PROGRESS {done} {total}"); Console.Out.Flush(); } }
        foreach (var f in folders)
        {
            void Fail(string form, string subject, Exception ex)
            {
                failed++; done++;
                log.WriteLine(Csv((++seq).ToString("X"), "", form, "FAILED: " + ex.Message, "", "", f.Target, subject));
            }
            // a folder that cannot be read (e.g. a damaged contents table) is reported and the others are still converted
            try { switch (f.Kind)
            {
                case "contacts":
                    store.ExportContacts(f, fields =>
                    {
                        Line(new() { ["type"] = "contact", ["folder"] = "@Contacts", ["fields"] = fields });
                        log.WriteLine(Csv((++seq).ToString("X"), "", "IPM.Contact", "contact", "", "", "@Contacts", fields.GetValueOrDefault("displayName") ?? ""));
                        contacts++; Progress();
                    }, (name, members) =>
                    {
                        Line(new() { ["type"] = "group", ["folder"] = "@Contacts", ["subject"] = name, ["members"] = members });
                        log.WriteLine(Csv((++seq).ToString("X"), "", "IPM.DistList", $"group ({members.Count} member(s))", "", "", "@Contacts", name));
                        groups++; Progress();
                    }, (subject, ex) => Fail("IPM.Contact", subject, ex));
                    break;
                case "calendar":
                    store.ExportAppointments(f, Path.Combine(outDir, "att"), item =>
                    {
                        Line(item);
                        bool rep = item.ContainsKey("rrule");
                        if (rep) repeating++;
                        log.WriteLine(Csv((++seq).ToString("X"), "", "IPM.Appointment", rep ? "appointment (repeating)" : "appointment", "", "", "@Calendar", item["subject"] as string ?? ""));
                        appointments++; Progress();
                    }, (subject, ex) => Fail("IPM.Appointment", subject, ex));
                    break;
                case "tasks":
                    store.ExportTasks(f, item =>
                    {
                        Line(item);
                        log.WriteLine(Csv((++seq).ToString("X"), "", "IPM.Task", "task", "", "", "@Tasks", item["subject"] as string ?? ""));
                        tasks++; Progress();
                    }, (subject, ex) => Fail("IPM.Task", subject, ex));
                    break;
                default:   // mail, sticky notes, journal: EML files
                    if (!f.Target.StartsWith('@')) Line(new() { ["type"] = "folder", ["folder"] = f.Target });   // also empty folders
                    store.ExportFolder(f, emlDir, ref seq,
                        m =>
                        {
                            Line(new()
                            {
                                ["eml"] = m.Eml, ["folder"] = m.Target, ["subject"] = m.Subject, ["read"] = m.Read, ["draft"] = m.Draft,
                                ["received"] = m.Received, ["sent"] = m.Sent, ["class"] = m.MessageClass,
                            });
                            log.WriteLine(Csv(seq.ToString("X"), "", m.MessageClass, m.Note == null ? "exported" : "exported " + m.Note, "", m.Read ? "" : "yes", m.Target, m.Subject));
                            mails++; exFixed += m.ExFixed; Progress();
                        },
                        (cls, subject, ex) => Fail(cls, subject, ex));
                    break;
            } }
            catch (Exception ex)
            {
                failed++;
                log.WriteLine(Csv((++seq).ToString("X"), "", "(folder)", $"FAILED: folder not readable ({f.Count} item(s)): " + ex.Message, "", "", f.Target, f.Path));
                Console.WriteLine($"FAILED folder {f.Path}: {ex.Message}");
            }
        }
        Console.WriteLine($"PROGRESS {total} {total}");
        Console.WriteLine($"exported from the PST: {mails} mail(s), {contacts} contact(s), {groups} group(s), {appointments} appointment(s) ({repeating} repeating), {tasks} task(s); " +
                          $"{exFixed} Exchange address(es) replaced by SMTP; failed {failed}, {sw.Elapsed.TotalSeconds:F1}s");
        File.WriteAllText(Path.Combine(outDir, "address_stats.txt"), $"{exFixed}\t0\t0");
        return failed == 0 ? 0 : 1;
    }

    /// <summary>Manifest "attachments": [{ "path": file, "name": attachment name }] (calendar entries).</summary>
    internal static List<(string path, string name)> Attachments(JsonElement r) =>
        r.TryGetProperty("attachments", out var a) && a.ValueKind == JsonValueKind.Array
            ? a.EnumerateArray().Select(x => (x.GetProperty("path").GetString() ?? "", x.GetProperty("name").GetString() ?? "")).Where(x => File.Exists(x.Item1)).ToList()
            : new List<(string, string)>();

    /// <summary>Clear reasons instead of MAPI error codes: the file is in use (Outlook keeps the PSTs of its profile
    /// locked) or is not a PST at all. Null when the file looks usable.</summary>
    static string? CheckPst(string pst)
    {
        var head = new byte[4];
        try
        {
            using var fs = new FileStream(pst, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            if (fs.Read(head, 0, 4) < 4) return "not an Outlook data file (PST): " + pst;
        }
        catch (IOException) { return "the PST is in use by another program (close Outlook, or use a copy of the PST): " + pst; }
        catch (UnauthorizedAccessException) { return "no permission to read the PST: " + pst; }
        return head[0] == '!' && head[1] == 'B' && head[2] == 'D' && head[3] == 'N' ? null : "not an Outlook data file (PST): " + pst;
    }

    static string Csv(params string[] f) => string.Join(",", f.Select(x => "\"" + (x ?? "").Replace("\"", "\"\"") + "\""));
}
