// "export" / "mark-unread" mode: exports the mail of a Notes database as EML files plus a manifest for the PST writer.
//   Nsf2Pst export <db> <outdir> [--server S] [--user "CN=..."]
//   Nsf2Pst mark-unread <db> <noteids.txt> [--server S] [--user "CN=..."]   (test helper)
// Output: <outdir>\eml\<UNID>.eml, <outdir>\manifest.jsonl, <outdir>\export_log.csv
// Progress lines for the UI: "PROGRESS <done> <total>".
// Uses the Notes client of this machine (Lotus.NotesSession COM + nnotes.dll C API, 64-bit).
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using MimeKit;
using NsfExport;

namespace Nsf2Pst;

internal static class ExportTool
{
    public static int Run(string[] args)
    {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);   // Notes writes ISO-2022-JP/KR, GB2312, Big5 ... parts
        if (args.Length < 3) { Console.Error.WriteLine("usage: export <db> <outdir> [--server S] [--user CN] [--id file.id] [--directory server!!names.nsf] | mark-unread <db> <ids.txt>"); return 2; }
        string Opt(string name, string def) { int i = Array.IndexOf(args, name); return i >= 0 && i + 1 < args.Length ? args[i + 1] : def; }

        // optional: run as another Notes ID (the mailbox owner's) so that encrypted mail can be decrypted;
        // the password comes from the environment, never from the command line
        string idFile = Opt("--id", "");
        string password = Environment.GetEnvironmentVariable("NOTES_PASSWORD") ?? "";
        NotesApi.RestoreNotesIni();
        try
        {
            if (idFile.Length > 0)
            {
                password = Environment.GetEnvironmentVariable("NOTES_ID_PASSWORD") ?? "";
                Console.WriteLine("switched to Notes ID: " + NotesApi.SwitchId(Path.GetFullPath(idFile), password));
            }
            return Export(args, Opt, password);
        }
        finally { NotesApi.RestoreNotesIni(); }
    }

    static int Export(string[] args, Func<string, string, string> Opt, string password)
    {
        dynamic session = Activator.CreateInstance(Type.GetTypeFromProgID("Lotus.NotesSession", true)!)!;
        session.Initialize(password);
        string server = Opt("--server", ""), dbPath = args[1];
        dynamic db = session.GetDatabase(server, dbPath, false);
        if (!db.IsOpen) db.Open();
        string netPath = NotesApi.NetPath(server, (string)db.FilePath);

        // read/unread marks are stored per user name: default to the mailbox owner (calendar profile), not to the ID we run as
        string user = Opt("--user", "");
        if (user.Length == 0)
        {
            try { dynamic prof = db.GetProfileDocument("CalendarProfile"); user = First(prof, "Owner") as string ?? ""; } catch { }
            if (user.Length == 0) user = (string)session.UserName;
        }

        if (args[0] == "mark-unread")
        {
            var ids = File.ReadAllLines(args[2]).Where(l => l.Trim().Length > 0).Select(l => Convert.ToUInt32(l.Trim(), 16)).ToList();
            NotesApi.MarkUnread(netPath, user, ids);
            Console.WriteLine($"marked {ids.Count} note(s) unread for {user}");
            return 0;
        }

        string outDir = Path.GetFullPath(args[2]);
        string emlDir = Path.Combine(outDir, "eml");
        Directory.CreateDirectory(emlDir);
        Console.WriteLine($"Notes user: {session.UserName} | unread marks of: {user}");

        // ---- folder mapping -----------------------------------------------------------------------------------
        // Notes system folders/views -> default-folder tokens. Everything else that is a folder is a user folder,
        // including names that start with "(" such as "(Old)"; only the template's own hidden folders are skipped.
        var builtin = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            ["($Inbox)"] = "@Inbox", ["($Sent)"] = "@Sent", ["($Drafts)"] = "@Drafts", ["($SoftDeletions)"] = "@Deleted",
            ["($JunkMail)"] = "Junk E-mail",
            ["Deleted Items"] = "@Deleted",   // the folder PST -> NSF puts Outlook's deleted mail into: back to Outlook's own folder
        };
        var systemFolders = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        { "(Rules)", "(Group Calendars)", "(Drafts)", "(Sent)", "(Stationery)", "(Mail Threads)" };
        bool IsSystem(string n) => systemFolders.Contains(n) || n.StartsWith("($") || n.StartsWith("(RunRules", StringComparison.OrdinalIgnoreCase);

        var noteFolders = new Dictionary<uint, List<string>>();   // noteid -> target folder paths
        var skippedFolders = new List<string>();
        void Collect(dynamic view, string target)
        {
            dynamic d = view.GetFirstDocument();
            while (d != null)
            {
                uint id = Convert.ToUInt32((string)d.NoteID, 16);
                if (!noteFolders.TryGetValue(id, out var list)) noteFolders[id] = list = new List<string>();
                if (!list.Contains(target)) list.Add(target);
                d = view.GetNextDocument(d);
            }
        }
        Console.WriteLine("reading folders ...");
        foreach (var kv in builtin)
        {
            dynamic v = db.GetView(kv.Key);
            if (v != null) Collect(v, kv.Value);
        }
        foreach (dynamic v in (object[])db.Views)
        {
            string name = ((string)v.Name).Trim('\\');
            if (!(bool)v.IsFolder || builtin.ContainsKey(name)) continue;
            if (IsSystem(name)) { skippedFolders.Add(name); continue; }
            Collect(v, name);
        }

        // ---- unread marks ---------------------------------------------------------------------------------------
        var unread = NotesApi.UnreadNoteIds(netPath, user);

        // ---- Notes name -> SMTP address table (learned from the whole mail file, optional Domino directory) --
        Console.WriteLine("learning addresses ...");
        var book = new AddressBook(session, Opt("--directory", ""));
        {
            dynamic learnAll = db.AllDocuments;
            dynamic ld = learnAll.GetFirstDocument();
            while (ld != null) { try { book.Learn(ld); } catch { } ld = learnAll.GetNextDocument(ld); }
        }

        // documents Notes itself shows in the calendar: decides which invitations get a tentative appointment
        var inCalendar = new HashSet<uint>();
        try
        {
            dynamic cv = db.GetView("($Calendar)");
            if (cv != null)
            {
                dynamic ec = cv.AllEntries;
                dynamic e = ec.GetFirstEntry();
                while (e != null)
                {
                    if ((bool)e.IsDocument) inCalendar.Add(Convert.ToUInt32((string)e.NoteID, 16));
                    e = ec.GetNextEntry(e);
                }
            }
        }
        catch { }

        // ---- export ---------------------------------------------------------------------------------------------
        // mail-like documents; "Notice" = meeting invitations / replies / cancellations sent as mail (calendar entries stay out)
        var mailForms = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        { "", "Memo", "Reply", "Reply With History", "Return Receipt", "NonDelivery Report", "Delivery Report", "Delivery Failure", "Notice" };
        using var manifest = new StreamWriter(Path.Combine(outDir, "manifest.jsonl"), false, new UTF8Encoding(false));
        using var log = new StreamWriter(Path.Combine(outDir, "export_log.csv"), false, new UTF8Encoding(true));
        log.WriteLine("NoteID,UNID,Form,Status,BodyRebuilt,Unread,Folders,Subject");
        int exported = 0, skipped = 0, failed = 0, rebuilt = 0, done = 0;
        int contacts = 0, invitations = 0, appointments = 0, appointmentItems = 0, encryptedOk = 0, encryptedFailed = 0;
        var skippedForms = new Dictionary<string, int>();

        dynamic all = db.AllDocuments;
        int total = (int)all.Count;
        Console.WriteLine($"exporting {total} document(s) ...");
        dynamic doc = all.GetFirstDocument();
        while (doc != null)
        {
            dynamic next = all.GetNextDocument(doc);
            uint id = Convert.ToUInt32((string)doc.NoteID, 16);
            string unid = doc.UniversalID;
            string form = First(doc, "Form") as string ?? "";
            string subject = Regex.Replace(First(doc, "Subject") as string ?? "", @"\s*[\r\n]+\s*", " ");   // Notes subjects can contain line breaks
            if (form.Equals("Person", StringComparison.OrdinalIgnoreCase))
            {
                try
                {
                    Dictionary<string, string> fields = ContactFields(doc);
                    manifest.WriteLine(JsonSerializer.Serialize(new Dictionary<string, object> { ["type"] = "contact", ["folder"] = "@Contacts", ["fields"] = fields }));
                    contacts++;
                    log.WriteLine(Csv(id.ToString("X"), unid, form, "contact", "", "", "@Contacts", fields.GetValueOrDefault("displayName") ?? ""));
                }
                catch (Exception ex) { failed++; log.WriteLine(Csv(id.ToString("X"), unid, form, "FAILED: " + ex.Message, "", "", "@Contacts", "")); }
            }
            else if (form.Equals("Appointment", StringComparison.OrdinalIgnoreCase))
            {
                try
                {
                    List<Dictionary<string, object>> items = Appointments(doc, subject);
                    foreach (var a in items) manifest.WriteLine(JsonSerializer.Serialize(a));
                    appointments++; appointmentItems += items.Count;
                    log.WriteLine(Csv(id.ToString("X"), unid, form, "appointment", "", "", $"@Calendar x{items.Count}", subject));
                }
                catch (Exception ex) { failed++; log.WriteLine(Csv(id.ToString("X"), unid, form, "FAILED: " + ex.Message, "", "", "@Calendar", subject)); }
            }
            else if (!mailForms.Contains(form))
            {
                skipped++; skippedForms[form] = skippedForms.GetValueOrDefault(form) + 1;
                log.WriteLine(Csv(id.ToString("X"), unid, form, "skipped (not mail)", "", "", "", subject));
            }
            else if (IsEncrypted(doc) && doc.GetFirstItem("Body") == null)
            {
                // encrypted for someone else's key: Notes hides the sealed items from this ID
                failed++; encryptedFailed++;
                log.WriteLine(Csv(id.ToString("X"), unid, form, "FAILED: encrypted - needs the mailbox owner's Notes ID", "", "", "", subject));
            }
            else
            {
                bool encrypted = IsEncrypted(doc);
                if (encrypted) encryptedOk++;
                var folders = noteFolders.TryGetValue(id, out var fl) ? fl : new List<string> { "Unfiled" };
                bool isUnread = unread.Contains(id);
                string eml = Path.Combine(emlDir, unid + ".eml");
                try
                {
                    NotesApi.ExportEml(netPath, id, eml);
                    var attachments = NotesAttachments(session, doc);
                    bool bodyRebuilt = FixEml(eml, subject, doc, attachments, book);
                    if (encrypted) { RebuildEncrypted(eml, doc, attachments); bodyRebuilt = true; }
                    if (bodyRebuilt) rebuilt++;
                    DateTime? posted = First(doc, "PostedDate") as DateTime?;
                    DateTime? delivered = First(doc, "DeliveredDate") as DateTime? ?? posted;
                    foreach (string f in folders)
                    {
                        var line = new Dictionary<string, object?>
                        {
                            ["eml"] = eml, ["folder"] = f, ["subject"] = subject, ["read"] = !isUnread,
                            ["received"] = delivered is DateTime r ? new DateTimeOffset(r) : null,
                            ["sent"] = posted is DateTime p ? new DateTimeOffset(p) : null,
                        };
                        manifest.WriteLine(JsonSerializer.Serialize(line.Where(x => x.Value != null).ToDictionary(x => x.Key, x => x.Value)));
                    }
                    exported++;
                    string status = encrypted ? "exported (encrypted)" : "exported";
                    // unanswered invitations / updates: Notes shows them in the calendar, Outlook adds a tentative appointment
                    string noticeType = Convert.ToString((object?)First(doc, "NoticeType")) ?? "";
                    if (form.Equals("Notice", StringComparison.OrdinalIgnoreCase) && noticeType is "I" or "U" && inCalendar.Contains(id) && Dates(doc, "StartDateTime").Count > 0)
                    {
                        try
                        {
                            List<Dictionary<string, object>> inv = Appointments(doc, subject);
                            foreach (var a in inv) { a["tentative"] = true; manifest.WriteLine(JsonSerializer.Serialize(a)); }
                            invitations++; status += " + calendar (tentative)";
                        }
                        catch (Exception ex) { status += " (calendar copy failed: " + ex.Message + ")"; }
                    }
                    log.WriteLine(Csv(id.ToString("X"), unid, form, status, bodyRebuilt ? "yes" : "", isUnread ? "yes" : "", string.Join(" | ", folders), subject));
                }
                catch (Exception ex)
                {
                    failed++;
                    log.WriteLine(Csv(id.ToString("X"), unid, form, "FAILED: " + ex.Message, "", "", string.Join(" | ", folders), subject));
                }
            }
            done++;
            if (done % 5 == 0 || done == total) { Console.WriteLine($"PROGRESS {done} {total}"); Console.Out.Flush(); }
            doc = next;
        }

        Console.WriteLine($"exported {exported} mail(s), {contacts} contact(s), {appointments} calendar entr(ies) -> {appointmentItems} appointment(s), {invitations} unanswered invitation(s) -> tentative; body rebuilt {rebuilt}, skipped {skipped}, failed {failed}");
        Console.WriteLine($"unread marks of {user}: {unread.Count} | encrypted mails: {encryptedOk} decrypted, {encryptedFailed} not decryptable with this ID");
        Console.WriteLine($"addresses: {book.Fixed} Notes name(s) replaced by SMTP address, {book.Unresolved} not resolvable ({book.UnresolvedNames.Count} distinct; table of {book.Learned} learned from this mail file)");
        File.WriteAllText(Path.Combine(outDir, "address_stats.txt"), $"{book.Fixed}\t{book.Unresolved}\t{book.UnresolvedNames.Count}");
        File.WriteAllLines(Path.Combine(outDir, "unresolved_names.txt"), book.UnresolvedNames.OrderBy(x => x), new UTF8Encoding(true));
        if (skippedForms.Count > 0) Console.WriteLine("skipped forms: " + string.Join(", ", skippedForms.Select(k => $"{(k.Key == "" ? "(none)" : k.Key)}={k.Value}")));
        if (skippedFolders.Count > 0) Console.WriteLine("skipped system folders: " + string.Join(", ", skippedFolders));
        return failed == 0 ? 0 : 1;
    }

    static object? First(dynamic doc, string item)
    {
        object v = doc.GetItemValue(item);
        return v is object[] a && a.Length > 0 ? a[0] : null;
    }

    static bool IsEncrypted(dynamic doc)
    {
        try { if ((bool)doc.IsEncrypted) return true; } catch { }
        return (bool)doc.HasItem("$Seal") || (bool)doc.HasItem("$SealData") || (First(doc, "Encrypt") as string) == "1";
    }

    static string Text(dynamic doc, string item)
    {
        object v = doc.GetItemValue(item);
        if (v is object[] a) return string.Join(", ", a.Select(x => x?.ToString() ?? "").Where(x => x.Length > 0)).Trim();
        return "";
    }

    static string BodyText(dynamic doc)
    {
        dynamic body = doc.GetFirstItem("Body");
        if (body == null) return "";
        try { return ((int)body.Type == 1 ? (string?)body.GetFormattedText(false, 0) : (string?)body.Text) ?? ""; } catch { return ""; }
    }

    /// <summary>"CN=Alex Chan/OU=Sales/O=ExampleCorp" -> "Alex Chan"; other values unchanged.</summary>
    static string CommonName(string n)
    {
        if (n.StartsWith("CN=", StringComparison.OrdinalIgnoreCase)) n = n[3..];
        int slash = n.IndexOf('/');
        return slash > 0 && !n.Contains('@') ? n[..slash] : n;
    }

    // Notes contact (Person) items -> PST contact fields. Items not mapped here are kept in the contact's notes.
    static readonly (string key, string[] items)[] ContactMap =
    {
        ("givenName", new[] { "FirstName" }), ("surname", new[] { "LastName" }), ("middleName", new[] { "MiddleInitial" }),
        ("company", new[] { "CompanyName" }), ("jobTitle", new[] { "JobTitle" }), ("department", new[] { "Department" }),
        ("businessPhone", new[] { "OfficePhoneNumber" }), ("homePhone", new[] { "PhoneNumber" }), ("mobilePhone", new[] { "CellPhoneNumber" }),
        ("businessFax", new[] { "OfficeFAXPhoneNumber" }), ("pager", new[] { "PhoneNumber_6" }), ("webPage", new[] { "WebSite" }),
        ("street", new[] { "OfficeStreetAddress" }), ("city", new[] { "OfficeCity" }), ("state", new[] { "OfficeState" }),
        ("postalCode", new[] { "OfficeZIP" }), ("country", new[] { "OfficeCountry" }),
    };
    static readonly HashSet<string> ContactTechnical = new(StringComparer.OrdinalIgnoreCase)
    {
        "Form", "Type", "MailSystem", "NameDisplayPref", "Logo", "ExcludeFromView", "PROTECTFROMARCHIVE", "InternetDomain",
        "NameHelperInput", "FullNameInput", "FullName", "MailAddress", "InternetAddress", "nts_email1", "nts_emailInternet1",
        "Work0Email", "Comment", "Categories", "MailDomain", "ShortName", "DocumentAccess", "Owner", "LocalAdmin",
    };

    static Dictionary<string, string> ContactFields(dynamic doc)
    {
        var f = new Dictionary<string, string>();
        var used = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var (key, items) in ContactMap)
            foreach (string it in items)
            {
                string v = Text(doc, it);
                if (v.Length > 0) { f[key] = v; used.Add(it); break; }
            }
        string display = Text(doc, "FullNameInput");
        if (display.Length == 0)
        {
            object fn = doc.GetItemValue("FullName");
            display = fn is object[] a && a.Length > 0 ? CommonName(a[0]?.ToString() ?? "") : "";
        }
        if (display.Length > 0) f["displayName"] = display;

        // e-mail: first values that look like Internet addresses, in this order
        var mails = new[] { "InternetAddress", "nts_emailInternet1", "Work0Email", "nts_email1", "MailAddress" }
            .Select(i => (string)Text(doc, i)).Where(v => v.Contains('@') && !v.Contains("CN=", StringComparison.OrdinalIgnoreCase))
            .SelectMany(v => v.Split(',', StringSplitOptions.TrimEntries | StringSplitOptions.RemoveEmptyEntries))
            .Distinct(StringComparer.OrdinalIgnoreCase).ToList();
        if (mails.Count > 0) f["email"] = mails[0];
        if (mails.Count > 1) f["email2"] = mails[1];

        var notes = new List<string>();
        string comment = Text(doc, "Comment");
        if (comment.Length > 0) notes.Add(comment);
        foreach (dynamic item in (object[])doc.Items)
        {
            string name = item.Name;
            if (name.StartsWith('$') || used.Contains(name) || ContactTechnical.Contains(name)) continue;
            string v = Text(doc, name);
            if (v.Length > 0) notes.Add($"{name}: {v}");
        }
        if (notes.Count > 0) f["notes"] = string.Join("\r\n", notes);
        return f;
    }

    /// <summary>Notes calendar entry -> one appointment per instance (repeating entries keep one start/end pair per
    /// occurrence in StartDateTime / EndDateTime, so they become separate single appointments).</summary>
    static List<Dictionary<string, object>> Appointments(dynamic doc, string subject)
    {
        object? typeObj = First(doc, "AppointmentType"); string type = Convert.ToString(typeObj) ?? "0";   // 0 appt, 1 anniversary, 2 all day, 3 meeting, 4 reminder
        bool allDay = type is "1" or "2";
        List<DateTime> starts = Dates(doc, "StartDateTime");
        if (starts.Count == 0) starts = Dates(doc, "CalendarDateTime");
        if (starts.Count == 0) starts = Dates(doc, "StartDate");
        List<DateTime> ends = Dates(doc, "EndDateTime");
        if (starts.Count == 0) throw new InvalidOperationException("calendar entry without a start date");
        TimeSpan firstLength = ends.Count > 0 ? ends[0] - starts[0] : TimeSpan.FromHours(1);
        if (type == "4") firstLength = TimeSpan.Zero;   // reminder: a point in time

        int? reminder = null;
        object? alarm = First(doc, "$Alarm"), offset = First(doc, "$AlarmOffset");
        if (Convert.ToString(alarm) == "1" && offset is double off) reminder = (int)Math.Max(0, -off);

        string location = new[] { Text(doc, "Location"), Text(doc, "Room") }.FirstOrDefault(x => x.Length > 0) ?? "";
        var bodyParts = new List<string>();
        string chair = Text(doc, "Chair");
        if (chair.Length > 0) bodyParts.Add("主持人 / Chair: " + CommonName(chair));
        foreach (var (label, item) in new[] { ("必須出席 / Required", "RequiredAttendees"), ("可選出席 / Optional", "OptionalAttendees") })
        {
            object v = doc.GetItemValue(item);
            if (v is object[] a && a.Length > 0)
            {
                string list = string.Join("; ", a.Select(x => CommonName(x?.ToString() ?? "")).Where(x => x.Length > 0));
                if (list.Length > 0) bodyParts.Add($"{label}: {list}");
            }
        }
        string body = BodyText(doc);
        if (body.Trim().Length > 0) { if (bodyParts.Count > 0) bodyParts.Add(""); bodyParts.Add(body.Trim()); }

        var result = new List<Dictionary<string, object>>();
        for (int i = 0; i < starts.Count; i++)
        {
            DateTime s = starts[i];
            DateTime e = i < ends.Count ? ends[i] : s + firstLength;
            if (allDay) { s = s.Date; e = (e < s ? s : e).Date.AddDays(1); }
            var a = new Dictionary<string, object>
            {
                ["type"] = "appointment", ["folder"] = "@Calendar", ["subject"] = subject,
                ["start"] = new DateTimeOffset(s), ["end"] = new DateTimeOffset(e < s ? s : e), ["allDay"] = allDay,
            };
            if (location.Length > 0) a["location"] = location;
            if (bodyParts.Count > 0) a["body"] = string.Join("\r\n", bodyParts);
            if (reminder.HasValue) a["reminder"] = reminder.Value;
            result.Add(a);
        }
        return result;
    }

    static List<DateTime> Dates(dynamic doc, string item)
    {
        object v = doc.GetItemValue(item);
        return v is object[] a ? a.OfType<DateTime>().ToList() : new List<DateTime>();
    }

    /// <summary>Unicode attachment names and sizes as Notes stores them (@AttachmentNames / @AttachmentLengths).</summary>
    static List<(string name, long size)> NotesAttachments(dynamic session, dynamic doc)
    {
        var result = new List<(string, long)>();
        if (!(bool)doc.HasEmbedded) return result;
        object[] names = session.Evaluate("@AttachmentNames", doc);
        object[] sizes = session.Evaluate("@AttachmentLengths", doc);
        for (int i = 0; i < names.Length; i++)
        {
            string n = names[i] as string ?? "";
            if (n.Length > 0) result.Add((n, i < sizes.Length ? Convert.ToInt64(sizes[i]) : -1));
        }
        return result;
    }

    static string Csv(params string[] f) => string.Join(",", f.Select(x => "\"" + (x ?? "").Replace("\"", "\"\"") + "\""));

    // Puts the Notes Unicode subject into the EML and, when Notes' rich text -> MIME conversion lost characters
    // (it may pick ISO-2022-JP for Traditional Chinese), rebuilds the text parts from the Notes Unicode text.
    static bool FixEml(string path, string notesSubject, dynamic doc, List<(string name, long size)> notesAttachments, AddressBook book)
    {
        MimeMessage msg;
        using (var ms = new MemoryStream(File.ReadAllBytes(path))) msg = MimeMessage.Load(ms);
        msg.Headers.RemoveAll("X-Notes-Item");
        if (notesSubject.Length > 0) msg.Subject = notesSubject;
        book.Fix(msg);

        // drop Notes' 1x1 spacer image (ecblank.gif) and the <img> tags that reference it
        var spacers = msg.BodyParts.OfType<MimePart>()
            .Where(p => string.Equals(p.FileName, "ecblank.gif", StringComparison.OrdinalIgnoreCase)).ToList();
        foreach (var sp in spacers)
        {
            foreach (var html in msg.BodyParts.OfType<TextPart>().Where(t => t.IsHtml && !t.IsAttachment))
                if (!string.IsNullOrEmpty(sp.ContentId))
                    html.SetText("utf-8",   // re-encoding the decoded text as UTF-8 is lossless
                        Regex.Replace(html.Text, "<img[^>]*cid:" + Regex.Escape(sp.ContentId) + "[^>]*>", "", RegexOptions.IgnoreCase));
            RemovePart(msg.Body, sp);
        }

        // attachment names: Notes may write them in a charset that lacks some characters ("?????.pdf" for 季度報告.pdf).
        // Take the Unicode name Notes stores instead: match each MIME attachment to a Notes attachment by decoded size
        // (unique size first, then the remaining ones in order), and re-encode every name as UTF-8.
        var parts = msg.BodyParts.OfType<MimePart>().Where(p => !string.IsNullOrEmpty(p.FileName)).ToList();   // FileName is non-empty below
        var unused = new List<(string name, long size)>(notesAttachments);
        var fixedName = new Dictionary<MimePart, string>();
        foreach (var part in parts)
        {
            if (unused.Any(a => a.name == part.FileName)) { unused.Remove(unused.First(a => a.name == part.FileName)); fixedName[part] = part.FileName!; }
        }
        foreach (var part in parts.Where(p => !fixedName.ContainsKey(p)))
        {
            long size = DecodedSize(part);
            var same = unused.Where(a => a.size == size).ToList();
            var pick = same.Count >= 1 ? same[0] : unused.FirstOrDefault();
            if (pick.name == null) continue;
            unused.Remove(pick);
            if (part.FileName!.Contains('?') || part.FileName.Contains('�')) fixedName[part] = pick.name;
        }
        foreach (var part in parts)
        {
            string name = fixedName.TryGetValue(part, out var fn) ? fn : part.FileName!;
            part.FileName = name;
            foreach (var prm in new[] { part.ContentDisposition?.Parameters, part.ContentType.Parameters })
                foreach (var p in prm?.Where(x => x.Name is "filename" or "name") ?? Enumerable.Empty<Parameter>())
                {
                    p.Value = name;
                    p.Encoding = Encoding.UTF8;
                    p.EncodingMethod = ParameterEncodingMethod.Rfc2047;   // Outlook reads RFC 2047 names reliably
                }
        }

        bool rebuilt = false;
        dynamic body = doc.GetFirstItem("Body");
        if (body != null && (int)body.Type == 1)   // RICHTEXT: converted by Notes, may be lossy
        {
            string notesText = ((string?)body.GetFormattedText(false, 0) ?? "").Replace("\r\n", "\n");
            int notesQ = notesText.Count(c => c == '?');
            foreach (var tp in msg.BodyParts.OfType<TextPart>().Where(t => !t.IsAttachment).ToList())
            {
                string text = tp.IsHtml ? System.Net.WebUtility.HtmlDecode(Regex.Replace(tp.Text, "<[^>]+>", " ")) : tp.Text;
                bool lossy = text.Contains('�') || text.Count(c => c == '?') > notesQ;
                if (!lossy) continue;
                if (tp.IsHtml)
                    tp.SetText("utf-8", "<html><head><meta charset=\"utf-8\"></head><body><div style=\"font-family:Microsoft JhengHei,sans-serif\">"
                        + System.Net.WebUtility.HtmlEncode(notesText).Replace("\n", "<br>\n") + "</div></body></html>");
                else
                    tp.SetText("utf-8", notesText.Replace("\n", "\r\n"));
                rebuilt = true;
            }
        }
        using var fs = File.Create(path);
        msg.WriteTo(fs);
        return rebuilt;
    }

    /// <summary>MIMEEMLExport keeps encrypted items sealed ("Portions of this MIME document are encrypted ...").
    /// The Notes session decrypts them with the current ID, so the body text and attachments are rebuilt from there.</summary>
    static void RebuildEncrypted(string path, dynamic doc, List<(string name, long size)> notesAttachments)
    {
        MimeMessage msg;
        using (var ms = new MemoryStream(File.ReadAllBytes(path))) msg = MimeMessage.Load(ms);
        string text = BodyText(doc).Replace("\r\n", "\n");
        var mixed = new Multipart("mixed")
        {
            new TextPart("html") { Text = "<html><head><meta charset=\"utf-8\"></head><body><div style=\"font-family:Microsoft JhengHei,sans-serif\">"
                + System.Net.WebUtility.HtmlEncode(text).Replace("\n", "<br>\n") + "</div></body></html>" },
        };
        ((TextPart)mixed[0]).SetText("utf-8", ((TextPart)mixed[0]).Text);
        string tmp = Path.Combine(Path.GetTempPath(), "nsf2pst_" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(tmp);
        try
        {
            int n = 0;
            foreach (var (name, _) in notesAttachments)
            {
                dynamic att = doc.GetAttachment(name);
                if (att == null) continue;
                string file = Path.Combine(tmp, (n++).ToString());
                att.ExtractFile(file);
                var part = new MimePart(MimeTypes.GetMimeType(name))
                {
                    Content = new MimeContent(new MemoryStream(File.ReadAllBytes(file))),
                    ContentDisposition = new ContentDisposition(ContentDisposition.Attachment),
                    ContentTransferEncoding = ContentEncoding.Base64, FileName = name,
                };
                foreach (var p in part.ContentDisposition.Parameters.Concat(part.ContentType.Parameters).Where(x => x.Name is "filename" or "name"))
                { p.Encoding = Encoding.UTF8; p.EncodingMethod = ParameterEncodingMethod.Rfc2047; }
                mixed.Add(part);
            }
        }
        finally { try { Directory.Delete(tmp, true); } catch { } }
        msg.Body = mixed.Count == 1 ? mixed[0] : mixed;
        using var fs = File.Create(path);
        msg.WriteTo(fs);
    }

    static long DecodedSize(MimePart part)
    {
        if (part.Content == null) return -1;
        using var ms = new MemoryStream();
        part.Content.DecodeTo(ms);
        return ms.Length;
    }

    static bool RemovePart(MimeEntity? node, MimeEntity target)
    {
        if (node is not Multipart mp) return false;
        if (mp.Remove(target)) return true;
        return mp.Any(child => RemovePart(child, target));
    }
}
