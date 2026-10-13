// PST -> NSF, Notes side.
//   Nsf2Pst nsf-import <new.nsf> <manifest.jsonl> [--server S] [--user "CN=..."] [--template mail12.ntf] [--title T]
//   Nsf2Pst nsf-counts <db.nsf> [--server S] [--user "CN=..."]      "folder TAB items TAB unread", numbers only
// A new mail file is created from the Notes mail template; every EML from the PST export becomes a native MIME mail
// (Notes' own MIME parser, C API MIMEStreamItemize), is filed like in the PST and gets the read / unread mark of
// the mailbox owner. Uses the Notes client of this machine (Lotus.NotesSession COM + nnotes.dll C API, 64-bit).
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.Json;
using NsfExport;

namespace Nsf2Pst;

internal static class NsfImportTool
{
    // PST default folders -> Notes: folders are filled with PutInFolder, Sent and Drafts are views (selected by items)
    static readonly Dictionary<string, string> NotesFolder = new(StringComparer.OrdinalIgnoreCase)
    { ["@Inbox"] = "($Inbox)", ["@Junk"] = "($JunkMail)", ["@Deleted"] = "Deleted Items" };
    static readonly Dictionary<string, string> CountViews = new(StringComparer.OrdinalIgnoreCase)
    { ["($Inbox)"] = "@Inbox", ["($Sent)"] = "@Sent", ["($Drafts)"] = "@Drafts", ["($JunkMail)"] = "@Junk" };

    public static int Run(string[] args)
    {
        if (args.Length < 2) { Console.Error.WriteLine("usage: nsf-import <new.nsf> <manifest.jsonl> [--server S] [--user CN] [--template mail12.ntf] [--title T] | nsf-counts <db.nsf>"); return 2; }
        string Opt(string name, string def) { int i = Array.IndexOf(args, name); return i >= 0 && i + 1 < args.Length ? args[i + 1] : def; }
        if (args[0].ToLowerInvariant() == "nsf-mark-unread") return MarkUnread(args, Opt);
        dynamic session = Activator.CreateInstance(Type.GetTypeFromProgID("Lotus.NotesSession", true)!)!;
        session.Initialize(Environment.GetEnvironmentVariable("NOTES_PASSWORD") ?? "");
        try
        {
            Func<string, string, string> opt = Opt;   // a method group cannot be passed along with a dynamic argument
            return args[0].ToLowerInvariant() == "nsf-counts" ? Counts(session, args[1], opt) : Import(session, args, opt);
        }
        catch (Exception ex)
        {
            Console.WriteLine("fatal: " + ex.Message);
            return 3;
        }
    }

    /// <summary>
    /// nsf-mark-unread &lt;db.nsf&gt; &lt;unread_ids.txt&gt;: writes the unread marks of the import. This runs in its own process
    /// with the C API only: a process that uses the Notes COM classes cannot shut Notes down (NotesTerm deadlocks there),
    /// so the unread table it writes is never flushed to disk. Here NotesTerm closes the database cleanly.
    /// File format: "user=CN=..." on the first line, then one note ID (hex) per line.
    /// </summary>
    static int MarkUnread(string[] args, Func<string, string, string> Opt)
    {
        if (args.Length < 3) { Console.Error.WriteLine("usage: nsf-mark-unread <db.nsf> <unread_ids.txt> [--server S]"); return 2; }
        string server = Opt("--server", "");
        string[] lines = File.ReadAllLines(args[2], Encoding.UTF8);
        string user = lines.FirstOrDefault(l => l.StartsWith("user="))?[5..] ?? "";
        var ids = lines.Where(l => l.Length > 0 && !l.StartsWith("user=")).Select(l => Convert.ToUInt32(l.Trim(), 16)).ToList();
        try
        {
            NotesApi.InitCApi();
            string netPath = NotesApi.NetPath(server, server.Length == 0 ? Path.GetFullPath(args[1]) : args[1]);
            NotesApi.MarkUnread(netPath, user, ids);
            Console.WriteLine($"unread marks: {ids.Count} unread for {user}");
            return 0;
        }
        catch (Exception ex)
        {
            Console.WriteLine("fatal: " + ex.Message);
            return 3;
        }
        finally { NotesApi.Term(); }
    }

    static int Import(dynamic session, string[] args, Func<string, string, string> Opt)
    {
        if (args.Length < 3) { Console.Error.WriteLine("missing manifest"); return 2; }
        string server = Opt("--server", ""), nsf = args[1];
        if (server.Length == 0)
        {
            nsf = Path.GetFullPath(nsf);
            if (File.Exists(nsf)) { Console.WriteLine("fatal: the NSF already exists (a new file is always created): " + nsf); return 3; }
            if (Encoding.UTF8.GetByteCount(nsf) > NotesApi.MaxPathBytes)
            { Console.WriteLine($"fatal: the NSF path is too long for Notes (more than about {NotesApi.MaxPathBytes} bytes; use a shorter folder or file name): " + nsf); return 3; }
        }

        // ---- new mail file from the mail template -------------------------------------------------------------
        string dataDir = (string)session.GetEnvironmentString("Directory", true);
        string template = Opt("--template", "");
        if (template.Length == 0)
            template = Directory.GetFiles(dataDir, "mail*.ntf")
                .Select(Path.GetFileName).Where(n => n != null && System.Text.RegularExpressions.Regex.IsMatch(n, @"^mail\d+\.ntf$", System.Text.RegularExpressions.RegexOptions.IgnoreCase))
                .OrderByDescending(n => int.Parse(System.Text.RegularExpressions.Regex.Match(n!, @"\d+").Value)).FirstOrDefault() ?? "";
        if (template.Length == 0) { Console.WriteLine("fatal: no Notes mail template (mailNN.ntf) found in " + dataDir); return 3; }
        Console.WriteLine($"template: {template}");
        dynamic tmpl = session.GetDatabase("", template, false);
        if (!(bool)tmpl.IsOpen) tmpl.Open();
        session.ConvertMime = false;   // keep the imported mail as native MIME (saving must not turn it into rich text)
        dynamic db = tmpl.CreateFromTemplate(server, nsf, true);
        if (db == null) throw new InvalidOperationException("could not create " + nsf);
        db.Title = Opt("--title", Path.GetFileNameWithoutExtension(nsf));

        // read / unread marks belong to a user name: the mailbox owner, by default the Notes user running this
        string user = Opt("--user", "");
        if (user.Length == 0) user = (string)session.UserName;
        try { dynamic prof = db.GetProfileDocument("CalendarProfile"); prof.ReplaceItemValue("Owner", user); prof.Save(true, false); } catch { }
        string filePath = (string)db.FilePath;
        string netPath = NotesApi.NetPath(server, filePath);
        Console.WriteLine($"created {filePath} | owner: {user}");

        // ---- mails ------------------------------------------------------------------------------------------------
        string outDir = Path.GetDirectoryName(Path.GetFullPath(args[2]))!;
        string[] lines = File.ReadAllLines(args[2], Encoding.UTF8).Where(l => !string.IsNullOrWhiteSpace(l)).ToArray();
        using var log = new StreamWriter(Path.Combine(outDir, "import_log.csv"), false, new UTF8Encoding(true));
        log.WriteLine("Line,NoteID,Folder,Status,Subject");
        var unread = new List<uint>();
        int ok = 0, failed = 0, contacts = 0, groups = 0, appointments = 0, tasks = 0;
        var sw = Stopwatch.StartNew();
        using (var cdb = new NotesApi.Db(netPath))
        {
            for (int i = 0; i < lines.Length; i++)
            {
                if (i % 5 == 0) { Console.WriteLine($"PROGRESS {i} {lines.Length}"); Console.Out.Flush(); }
                string folder = "", subject = "";
                try
                {
                    using var j = JsonDocument.Parse(lines[i]);
                    var r = j.RootElement;
                    folder = r.TryGetProperty("folder", out var f) ? f.GetString() ?? "" : "";
                    subject = r.TryGetProperty("subject", out var s) ? s.GetString() ?? "" : "";
                    string type = r.TryGetProperty("type", out var ty) ? ty.GetString() ?? "mail" : "mail";
                    if (type == "contact")
                    {
                        var fields = r.GetProperty("fields").EnumerateObject().ToDictionary(x => x.Name, x => x.Value.GetString() ?? "");
                        subject = fields.GetValueOrDefault("displayName") ?? "";
                        string cid = ImportContact(db, fields);
                        contacts++; ok++;
                        log.WriteLine(Csv((i + 1).ToString(), cid, folder, "contact", subject));
                        continue;
                    }
                    if (type == "group")
                    {
                        var members = r.TryGetProperty("members", out var mb) ? mb.EnumerateArray().Select(x => x.GetString() ?? "").Where(x => x.Length > 0).ToArray() : Array.Empty<string>();
                        string gid = ImportGroup(db, subject, members);
                        groups++; ok++;
                        log.WriteLine(Csv((i + 1).ToString(), gid, folder, $"group ({members.Length} member(s))", subject));
                        continue;
                    }
                    if (type == "appointment")
                    {
                        // the Notes calendar API stores the entry like the Notes client does (repeat rules, alarm, busy time)
                        string uid = Guid.NewGuid().ToString("N") + "-nsf2pst";
                        CreateCalendar(cdb, EventICalendar(r, uid));
                        appointments++; ok++;
                        string status = type;
                        var atts = PstTool.Attachments(r);
                        if (atts.Count > 0)
                        {
                            // the entry exists either way; a failure here only loses its attachments, and is reported
                            try { AttachToEntry(session, db, uid, atts); status += $" ({atts.Count} attachment(s))"; }
                            catch (Exception ex) { failed++; status += " - attachments FAILED: " + ex.Message; Console.WriteLine($"FAILED line {i + 1} (attachments): {ex.Message}"); }
                        }
                        log.WriteLine(Csv((i + 1).ToString(), "", folder, status, subject));
                        continue;
                    }
                    if (type == "task")
                    {
                        ImportTask(db, r, user);
                        tasks++; ok++;
                        log.WriteLine(Csv((i + 1).ToString(), "", folder, type, subject));
                        continue;
                    }
                    string eml = r.GetProperty("eml").GetString()!;
                    bool read = !r.TryGetProperty("read", out var rd) || rd.GetBoolean();
                    bool draft = r.TryGetProperty("draft", out var dr) && dr.GetBoolean();
                    DateTimeOffset? received = r.TryGetProperty("received", out var rc) ? rc.GetDateTimeOffset() : null;
                    DateTimeOffset? sent = r.TryGetProperty("sent", out var st) ? st.GetDateTimeOffset() : null;

                    uint id = NotesApi.CreateMimeNote(cdb, eml);
                    dynamic doc = db.GetDocumentByID(id.ToString("X"));
                    if (subject.Length > 0) doc.ReplaceItemValue("Subject", subject);   // Unicode subject as Outlook had it
                    if (draft)
                    {
                        // ($Drafts) shows memos that were never posted
                        if ((bool)doc.HasItem("PostedDate")) doc.RemoveItem("PostedDate");
                        if ((bool)doc.HasItem("DeliveredDate")) doc.RemoveItem("DeliveredDate");
                    }
                    else
                    {
                        DateTimeOffset posted = sent ?? received ?? DateTimeOffset.Now;
                        if (!(bool)doc.HasItem("PostedDate")) doc.ReplaceItemValue("PostedDate", posted.LocalDateTime);
                        // ($Sent) shows posted memos without DeliveredDate; everything else is a received mail
                        if (folder.Equals("@Sent", StringComparison.OrdinalIgnoreCase)) { if ((bool)doc.HasItem("DeliveredDate")) doc.RemoveItem("DeliveredDate"); }
                        else doc.ReplaceItemValue("DeliveredDate", (received ?? posted).LocalDateTime);
                    }
                    doc.Save(true, false);
                    string? target = NotesFolder.TryGetValue(folder, out var nf) ? nf : folder.StartsWith('@') ? null : folder;
                    if (target != null) doc.PutInFolder(target, true);
                    if (!read) unread.Add(id);
                    ok++;
                    log.WriteLine(Csv((i + 1).ToString(), id.ToString("X"), folder, "imported", subject));
                }
                catch (Exception ex)
                {
                    failed++;
                    log.WriteLine(Csv((i + 1).ToString(), "", folder, "FAILED: " + ex.Message, subject));
                    Console.WriteLine($"FAILED line {i + 1}: {ex.Message}");
                }
            }
        }
        // unread marks are written by the next step (nsf-mark-unread, a separate C-API-only process; see there)
        string idFile = Path.Combine(outDir, "unread_ids.txt");
        File.WriteAllLines(idFile, new[] { "user=" + user }.Concat(unread.Select(x => x.ToString("X"))), new UTF8Encoding(false));
        Console.WriteLine($"PROGRESS {lines.Length} {lines.Length}");
        Console.WriteLine($"imported {ok - contacts - groups - appointments - tasks} mail(s), {contacts} contact(s), {groups} group(s), {appointments} appointment(s), {tasks} to do(s) into {filePath}; " +
                          $"{unread.Count} to mark unread for {user}, failed {failed}, {sw.Elapsed.TotalSeconds:F1}s");
        Console.WriteLine($"next: Nsf2Pst.exe nsf-mark-unread \"{filePath}\" \"{idFile}\"");
        return failed == 0 ? 0 : 1;
    }

    // ---- contacts ------------------------------------------------------------------------------------------------------
    // the same item names ExportTool reads (ContactMap), so a contact survives NSF -> PST -> NSF
    static readonly (string key, string item)[] ContactItems =
    {
        ("givenName", "FirstName"), ("surname", "LastName"), ("middleName", "MiddleInitial"), ("company", "CompanyName"),
        ("jobTitle", "JobTitle"), ("department", "Department"), ("businessPhone", "OfficePhoneNumber"), ("homePhone", "PhoneNumber"),
        ("mobilePhone", "CellPhoneNumber"), ("businessFax", "OfficeFAXPhoneNumber"), ("pager", "PhoneNumber_6"), ("webPage", "WebSite"),
        ("street", "OfficeStreetAddress"), ("city", "OfficeCity"), ("state", "OfficeState"), ("postalCode", "OfficeZIP"),
        ("country", "OfficeCountry"), ("notes", "Comment"),
    };

    /// <summary>Outlook contact fields -> a Notes contact (Person document in the mail file). Returns the note ID.</summary>
    static string ImportContact(dynamic db, Dictionary<string, string> f)
    {
        dynamic doc = db.CreateDocument();
        doc.ReplaceItemValue("Form", "Person");
        doc.ReplaceItemValue("Type", "Person");
        foreach (var (key, item) in ContactItems)
            if (f.TryGetValue(key, out var v) && v.Length > 0) doc.ReplaceItemValue(item, v);
        string display = f.GetValueOrDefault("displayName") ?? ((f.GetValueOrDefault("givenName") ?? "") + " " + (f.GetValueOrDefault("surname") ?? "")).Trim();
        if (display.Length > 0) { doc.ReplaceItemValue("FullName", display); doc.ReplaceItemValue("FullNameInput", display); }
        if (f.TryGetValue("email", out var mail))
        {
            doc.ReplaceItemValue("InternetAddress", mail);
            doc.ReplaceItemValue("MailAddress", mail);
        }
        if (f.TryGetValue("email2", out var mail2)) doc.ReplaceItemValue("nts_emailInternet1", mail2);
        doc.Save(true, false);
        return (string)doc.NoteID;
    }

    /// <summary>Outlook contact group -> a Notes personal group (Form "Group", as the Notes contacts view creates it);
    /// members are kept as Internet addresses ("Name &lt;address&gt;").</summary>
    static string ImportGroup(dynamic db, string name, string[] members)
    {
        dynamic doc = db.CreateDocument();
        doc.ReplaceItemValue("Form", "Group");
        doc.ReplaceItemValue("Type", "Group");
        doc.ReplaceItemValue("GroupType", "0");   // multi-purpose
        doc.ReplaceItemValue("ListName", name);
        doc.ReplaceItemValue("ListDescription", "Outlook contact group");
        if (members.Length > 0) doc.ReplaceItemValue("Members", members);
        doc.Save(true, false);
        return (string)doc.NoteID;
    }

    // ---- calendar and to do (iCalendar for CalCreateEntry) ---------------------------------------------------------------
    static string Esc(string s) => s.Replace("\\", "\\\\").Replace(";", "\\;").Replace(",", "\\,").Replace("\r\n", "\\n").Replace("\n", "\\n");
    static string Utc(DateTimeOffset d) => d.UtcDateTime.ToString("yyyyMMdd'T'HHmmss'Z'");
    static string? Text(JsonElement r, string name) => r.TryGetProperty(name, out var v) && v.ValueKind == JsonValueKind.String && v.GetString()!.Length > 0 ? v.GetString() : null;

    static StringBuilder Begin(string component, string uid)
    {
        var sb = new StringBuilder();
        sb.Append("BEGIN:VCALENDAR\r\nVERSION:2.0\r\nPRODID:-//Nsf2Pst//PST to NSF//EN\r\n");
        sb.Append($"BEGIN:{component}\r\nUID:{uid}\r\nDTSTAMP:{Utc(DateTimeOffset.UtcNow)}\r\n");
        return sb;
    }

    /// <summary>CalCreateEntry takes no attachments: they are added to the new entry's rich text Body afterwards. The entry
    /// is found by the iCalendar UID Notes keeps in $ICAL_UID (the main document, not the instance children).</summary>
    static void AttachToEntry(dynamic session, dynamic db, string uid, List<(string path, string name)> atts)
    {
        // Search needs a cutoff date (it cannot take an empty value through COM): the entry was created just now
        // "Today" keyword: independent of the system's date format. A cutoff of today or yesterday found nothing in tests,
        // one far in the past works.
        dynamic since = session.CreateDateTime("Today");
        since.AdjustYear(-30);
        dynamic dc = db.Search($"$ICAL_UID = \"{uid}\" & !@IsResponseDoc", since, 0);
        dynamic? doc = dc.Count > 0 ? dc.GetFirstDocument() : null;
        if (doc == null) throw new InvalidOperationException("calendar entry created, but not found to add its attachments");
        dynamic body = doc.GetFirstItem("Body");
        if (body == null || (int)body.Type != 1) body = doc.CreateRichTextItem("Body");
        string tmp = Path.Combine(Path.GetTempPath(), "nsf2pst_att_" + Guid.NewGuid().ToString("N"));
        try
        {
            foreach (var (path, name) in atts)
            {
                // the attachment takes the file's name: copy it under its real name first
                string dir = Path.Combine(tmp, Guid.NewGuid().ToString("N")[..8]);
                Directory.CreateDirectory(dir);
                string safe = string.Concat(name.Select(c => Path.GetInvalidFileNameChars().Contains(c) ? '_' : c));
                if (safe.Length > 150) safe = safe[..(150 - Path.GetExtension(safe).Length)] + Path.GetExtension(safe);   // file name limit
                string named = Path.Combine(dir, safe);
                File.Copy(path, named);
                body.EmbedObject(1454, "", named);   // EMBED_ATTACHMENT
            }
            doc.Save(true, false);
        }
        finally { try { Directory.Delete(tmp, true); } catch { } }
    }

    /// <summary>Appointment manifest line -> VEVENT. No ORGANIZER / ATTENDEE lines: creating the entry must never send
    /// invitations; the attendees are kept as text in the description (as the NSF -> PST direction does).</summary>
    static string EventICalendar(JsonElement r, string uid)
    {
        var sb = Begin("VEVENT", uid);
        sb.Append($"SUMMARY:{Esc(Text(r, "subject") ?? "")}\r\n");
        DateTimeOffset start = r.GetProperty("start").GetDateTimeOffset(), end = r.TryGetProperty("end", out var e) ? e.GetDateTimeOffset() : start;
        bool allDay = r.TryGetProperty("allDay", out var ad) && ad.GetBoolean();
        string? rrule = Text(r, "rrule");
        if (allDay)
        {
            DateTime s = start.LocalDateTime.Date, en = end.LocalDateTime.Date;
            if (en <= s) en = s.AddDays(1);
            sb.Append($"DTSTART;VALUE=DATE:{s:yyyyMMdd}\r\nDTEND;VALUE=DATE:{en:yyyyMMdd}\r\n");
        }
        else
            // UTC (a time without a zone is read as UTC by Notes, which moved the first occurrence by the UTC offset)
            sb.Append($"DTSTART:{Utc(start)}\r\nDTEND:{Utc(end < start ? start : end)}\r\n");
        if (rrule != null)
        {
            // the pattern holds local times (UNTIL, deleted occurrences): written in UTC like DTSTART
            static string ToUtc(string local) => Utc(new DateTimeOffset(DateTime.ParseExact(local, "yyyyMMdd'T'HHmmss", System.Globalization.CultureInfo.InvariantCulture)));
            if (!allDay) rrule = System.Text.RegularExpressions.Regex.Replace(rrule, @"UNTIL=(\d{8}T\d{6})", m => "UNTIL=" + ToUtc(m.Groups[1].Value));
            sb.Append($"RRULE:{rrule}\r\n");
            if (r.TryGetProperty("exdates", out var ex) && ex.ValueKind == JsonValueKind.Array)
                foreach (var x in ex.EnumerateArray()) sb.Append(allDay ? $"EXDATE;VALUE=DATE:{x.GetString()![..8]}\r\n" : $"EXDATE:{ToUtc(x.GetString()!)}\r\n");
        }
        if (Text(r, "location") is string loc) sb.Append($"LOCATION:{Esc(loc)}\r\n");
        if (Text(r, "body") is string body) sb.Append($"DESCRIPTION:{Esc(body)}\r\n");
        if (r.TryGetProperty("tentative", out var tt) && tt.GetBoolean()) sb.Append("STATUS:TENTATIVE\r\n");
        if (r.TryGetProperty("private", out var pv) && pv.GetBoolean()) sb.Append("CLASS:PRIVATE\r\n");
        if (allDay) sb.Append("TRANSP:TRANSPARENT\r\n");
        if (r.TryGetProperty("reminder", out var rm) && rm.ValueKind == JsonValueKind.Number)
            sb.Append($"BEGIN:VALARM\r\nACTION:DISPLAY\r\nDESCRIPTION:Reminder\r\nTRIGGER:-PT{rm.GetInt32()}M\r\nEND:VALARM\r\n");
        sb.Append("END:VEVENT\r\nEND:VCALENDAR\r\n");
        return sb.ToString();
    }

    /// <summary>Task manifest line -> a personal Notes to do (Form "Task"). The Notes calendar API (CalCreateEntry) takes
    /// VEVENT only, so to dos are written with the items the mail template's To Do views select on.</summary>
    static void ImportTask(dynamic db, JsonElement r, string owner)
    {
        dynamic doc = db.CreateDocument();
        doc.ReplaceItemValue("Form", "Task");
        // the owner is the chair of a personal to do; without Chair the To Do form shows "Participation is delegated for ..."
        foreach (string item in new[] { "Chair", "From", "Principal" })
        {
            dynamic it = doc.ReplaceItemValue(item, owner);
            it.IsNames = true;
        }
        doc.ReplaceItemValue("TaskType", "1");   // personal to do
        doc.ReplaceItemValue("ExcludeFromView", "D");   // never posted, so without this the Drafts view would show it
        doc.ReplaceItemValue("Subject", Text(r, "subject") ?? "");
        DateTime? Date(string name) => r.TryGetProperty(name, out var v) && v.ValueKind == JsonValueKind.String ? v.GetDateTimeOffset().LocalDateTime : null;
        if (Date("start") is DateTime s) { doc.ReplaceItemValue("StartDateTime", s); doc.ReplaceItemValue("StartDate", s.Date); }
        if (Date("due") is DateTime d) { doc.ReplaceItemValue("DueDateTime", d); doc.ReplaceItemValue("DueDate", d.Date); doc.ReplaceItemValue("CalendarDateTime", d); }
        bool complete = r.TryGetProperty("complete", out var c) && c.GetBoolean();
        // Outlook PR_IMPORTANCE 0 low / 2 high -> Notes Importance "3" low / "1" high (normal is left empty, as Notes does)
        if (r.TryGetProperty("importance", out var im) && im.ValueKind == JsonValueKind.Number)
            doc.ReplaceItemValue("Importance", im.GetInt32() == 2 ? "1" : im.GetInt32() == 0 ? "3" : "2");
        if (complete)
        {
            doc.ReplaceItemValue("DueState", 9);   // the "Completed" view selects DueState = 9
            // the To Do form shows "Completed on <CompletedDateTime>"; without the item it shows a formula error
            DateTime doneOn = Date("completedOn") ?? Date("due") ?? Date("start") ?? DateTime.Now;
            doc.ReplaceItemValue("CompletedDateTime", doneOn);
        }
        var body = new List<string>();
        if (!complete && r.TryGetProperty("percent", out var p) && p.ValueKind == JsonValueKind.Number && p.GetInt32() > 0)
            body.Add($"進度 / Progress: {p.GetInt32()}%");
        if (Text(r, "body") is string b) { if (body.Count > 0) body.Add(""); body.Add(b); }
        if (body.Count > 0)
        {
            // rich text with real paragraphs: the To Do form does not show line breaks of a plain text Body
            dynamic rt = doc.CreateRichTextItem("Body");
            string[] paras = string.Join("\n", body).Replace("\r\n", "\n").Split('\n');
            for (int k = 0; k < paras.Length; k++) { if (k > 0) rt.AddNewLine(1); if (paras[k].Length > 0) rt.AppendText(paras[k]); }
        }
        doc.Save(true, false);
    }
    static void CreateCalendar(NotesApi.Db db, string ical)
    {
        if (Environment.GetEnvironmentVariable("NSF2PST_DEBUG_ICAL") is string dir && dir.Length > 0)
            File.WriteAllText(Path.Combine(dir, Guid.NewGuid().ToString("N")[..8] + ".ics"), ical, new UTF8Encoding(false));
        NotesApi.CreateCalendarEntry(db, ical, lmbcs: true);   // CalCreateEntry reads LMBCS (UTF-8 bytes come out as mojibake)
    }

    /// <summary>Items and unread items per folder (default folders as @Inbox / @Sent / @Drafts / @Junk), numbers only.</summary>
    static int Counts(dynamic session, string nsf, Func<string, string, string> Opt)
    {
        string server = Opt("--server", "");
        dynamic db = NotesApi.OpenDatabase(session, server, server.Length == 0 ? Path.GetFullPath(nsf) : nsf);
        string user = Opt("--user", "");
        if (user.Length == 0)
        {
            try { dynamic prof = db.GetProfileDocument("CalendarProfile"); object v = prof.GetItemValue("Owner"); user = v is object[] a && a.Length > 0 ? a[0] as string ?? "" : ""; } catch { }
            if (user.Length == 0) user = (string)session.UserName;
        }
        var unread = NotesApi.UnreadNoteIds(NotesApi.NetPath(server, (string)db.FilePath), user);
        var shown = new HashSet<uint>();   // documents counted in a folder or in the Sent / Drafts views
        void Count(dynamic view, string label)
        {
            int n = 0, u = 0;
            dynamic d = view.GetFirstDocument();
            while (d != null)
            {
                uint id = Convert.ToUInt32((string)d.NoteID, 16);
                n++; shown.Add(id);
                if (unread.Contains(id)) u++;
                d = view.GetNextDocument(d);
            }
            if (n > 0) Console.WriteLine($"{label}\t{n}\t{u}");
        }
        foreach (var kv in CountViews)
        {
            dynamic v = db.GetView(kv.Key);
            if (v != null) Count(v, kv.Value);
        }
        foreach (dynamic v in (object[])db.Views)
        {
            string name = ((string)v.Name).Trim('\\');
            if (!(bool)v.IsFolder || name.StartsWith("($") || name is "(Rules)" or "(Group Calendars)" or "(Stationery)" or "(Mail Threads)") continue;
            Count(v, name == "Deleted Items" ? "@Deleted" : name);
        }

        // contacts, calendar entries, to dos, and mail that is in no folder at all
        int contacts = 0, groups = 0, calendar = 0, todos = 0, unfiled = 0, unfiledUnread = 0;
        var mailForms = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { "Memo", "Reply", "Reply With History", "Notice" };
        dynamic all = db.AllDocuments;
        dynamic doc = all.GetFirstDocument();
        while (doc != null)
        {
            object fv = doc.GetItemValue("Form");
            string form = fv is object[] a && a.Length > 0 ? a[0] as string ?? "" : "";
            if (form.Equals("Person", StringComparison.OrdinalIgnoreCase)) contacts++;
            else if (form.Equals("Group", StringComparison.OrdinalIgnoreCase)) groups++;
            else if (form.Equals("Appointment", StringComparison.OrdinalIgnoreCase) && !(bool)doc.IsResponse) calendar++;
            else if (form.Equals("Task", StringComparison.OrdinalIgnoreCase) && !(bool)doc.IsResponse) todos++;
            else if (mailForms.Contains(form))
            {
                uint id = Convert.ToUInt32((string)doc.NoteID, 16);
                if (!shown.Contains(id)) { unfiled++; if (unread.Contains(id)) unfiledUnread++; }
            }
            doc = all.GetNextDocument(doc);
        }
        if (contacts + groups > 0) Console.WriteLine($"@Contacts\t{contacts + groups}\t0");
        if (calendar > 0) Console.WriteLine($"@Calendar\t{calendar}\t0");
        if (todos > 0) Console.WriteLine($"@Tasks\t{todos}\t0");
        if (unfiled > 0) Console.WriteLine($"@Unfiled\t{unfiled}\t{unfiledUnread}");
        return 0;
    }

    static string Csv(params string[] f) => string.Join(",", f.Select(x => "\"" + (x ?? "").Replace("\"", "\"\"") + "\""));
}
