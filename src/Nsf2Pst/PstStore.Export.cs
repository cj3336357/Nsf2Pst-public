// PST -> NSF, PST side: walks the folders of a PST.
//  - mail, sticky-note and journal folders: every message is written as an EML file with Outlook's own MAPI -> MIME
//    converter (IConverterSession.MAPIToMIMEStm); Exchange (EX) addresses are replaced by the SMTP addresses the PST keeps.
//  - contacts, calendar and tasks: read from their MAPI properties into manifest lines (Outlook converts them only to
//    plain text). Repeating appointments are read from PidLidAppointmentRecur (MS-OXOCAL) into an iCalendar RRULE.
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Text;
using MimeKit;
using static PstLite.Consts;
using static PstLite.Props;

namespace PstLite;

internal sealed partial class PstStore
{
    const uint CCSF_INCLUDE_BCC = 0x0020;
    const uint STGM_CREATE_WRITE = 0x00001000 | 0x00000001;
    static readonly uint PR_CONTAINER_CLASS = Tags.Tag(0x3613, Tags.PT_UNICODE);
    static readonly uint PR_COMMENT_W = Tags.Tag(0x3004, Tags.PT_UNICODE);
    static readonly uint PR_INTERNET_CPID = Tags.Tag(0x3FDE, Tags.PT_LONG);
    const uint PR_ATTR_HIDDEN = 0x10F4000B;
    static readonly string[] HiddenClasses =
        { "IPF.Contact.RecipientCache", "IPF.Contact.GalContacts", "IPF.Contact.Company", "IPF.Contact.MOC", "IPF.Configuration" };
    /// <summary>Folders AllFolders left out (hidden Outlook folders), for the log.</summary>
    public List<string> SkippedFolders { get; } = new();
    static readonly Guid PSETID_Task = new("00062003-0000-0000-C000-000000000046");
    const uint PR_SENSITIVITY = 0x00360003, PR_IMPORTANCE = 0x00170003;

    /// <summary>Comment Nsf2Pst puts on the PST folder that holds the mail that was in no Notes folder: PST -> NSF leaves
    /// that mail outside any folder again instead of creating a folder "Unfiled".</summary>
    public const string UnfiledMarker = "Nsf2Pst: mail that was in no Notes folder";

    /// <summary>A PST folder and where its items go in Notes: @Inbox, @Sent, @Drafts, @Deleted, @Outbox, @Junk, @Unfiled,
    /// @Contacts, @Calendar, @Tasks or the folder path. Kind = mail, notes, journal, contacts, calendar or tasks.</summary>
    public sealed record PstFolder(string Path, string Target, string Kind, byte[] EntryId, int Count);

    /// <summary>One message written as EML. ExFixed = Exchange addresses replaced by SMTP addresses.</summary>
    public sealed record PstMessage(string Eml, string Target, string Subject, bool Read, DateTimeOffset? Received, DateTimeOffset? Sent,
                                    bool Draft, string MessageClass, int ExFixed, string? Note = null);

    byte[]? InboxEntryId()
    {
        IntPtr cls = Marshal.StringToHGlobalAnsi("IPM");
        try
        {
            if (store!.GetReceiveFolder(cls, 0, out uint cb, out IntPtr p, out IntPtr explicitClass) < 0) return null;
            var eid = new byte[cb]; Marshal.Copy(p, eid, 0, (int)cb);
            Native.MAPIFreeBuffer(p); if (explicitClass != IntPtr.Zero) Native.MAPIFreeBuffer(explicitClass);
            return eid;
        }
        finally { Marshal.FreeHGlobal(cls); }
    }

    /// <summary>Every folder under the top of the PST (depth first), with its Notes target.</summary>
    public List<PstFolder> AllFolders()
    {
        var special = new Dictionary<string, string>();
        void Mark(byte[]? eid, string target) { if (eid is { Length: > 0 }) special[Convert.ToHexString(eid)] = target; }
        Mark(InboxEntryId(), "@Inbox");
        Mark(StoreEntryId(0x35E4), "@Sent"); Mark(StoreEntryId(0x35E3), "@Deleted"); Mark(StoreEntryId(0x35E2), "@Outbox");
        // drafts, calendar, contacts and tasks ids are kept on the root folder (and on the Inbox)
        Check(store!.OpenEntry(0, IntPtr.Zero, IntPtr.Zero, MAPI_BEST_ACCESS, out _, out object rootObj), "OpenEntry(root)");
        var root = (IMAPIFolder)rootObj;
        try
        {
            foreach (var (id, target) in new[] { (0x36D7u, "@Drafts"), (0x36D0u, "@Calendar"), (0x36D1u, "@Contacts"), (0x36D4u, "@Tasks") })
            {
                uint tag = Tags.Tag(id, Tags.PT_BINARY);
                var rp = Get(t => { int hr = root.GetProps(t, 0, out uint c, out IntPtr p); return (hr, c, p); }, tag);
                Mark(rp.GetValueOrDefault(tag) as byte[], target);
            }
        }
        finally { Marshal.ReleaseComObject(root); }

        var list = new List<PstFolder>();
        Walk(ipmRoot, "", list, special);
        return list;
    }

    void Walk(IMAPIFolder parent, string parentPath, List<PstFolder> list, Dictionary<string, string> special)
    {
        Check(parent.GetHierarchyTable(MAPI_UNICODE, out IMAPITable hier), "GetHierarchyTable");
        var subs = QueryAll(hier, Tags.PR_ENTRYID, Tags.PR_DISPLAY_NAME_W, PR_CONTAINER_CLASS, Tags.PR_CONTENT_COUNT, PR_COMMENT_W, PR_ATTR_HIDDEN);
        Marshal.ReleaseComObject(hier);
        foreach (var s in subs)
        {
            byte[] eid = (byte[])s[Tags.PR_ENTRYID]!;
            string name = s.GetValueOrDefault(Tags.PR_DISPLAY_NAME_W) as string ?? "?";
            string path = parentPath.Length == 0 ? name : parentPath + "\\" + name;
            string cls = s.GetValueOrDefault(PR_CONTAINER_CLASS) as string ?? "";
            // Outlook's own hidden folders (Recipient Cache, GAL Contacts, Companies, PersonMetadata, settings ...) hold no
            // user data: their "contacts" are collected addresses and would flood the Notes contacts
            if (s.GetValueOrDefault(PR_ATTR_HIDDEN) is true || HiddenClasses.Any(h => cls.StartsWith(h, StringComparison.OrdinalIgnoreCase)))
            {
                SkippedFolders.Add(path);
                continue;
            }
            int count = s.GetValueOrDefault(Tags.PR_CONTENT_COUNT) is uint c ? (int)c : 0;
            string kind = cls.StartsWith("IPF.Appointment", StringComparison.OrdinalIgnoreCase) ? "calendar"
                : cls.StartsWith("IPF.Contact", StringComparison.OrdinalIgnoreCase) ? "contacts"
                : cls.StartsWith("IPF.Task", StringComparison.OrdinalIgnoreCase) ? "tasks"
                : cls.StartsWith("IPF.StickyNote", StringComparison.OrdinalIgnoreCase) ? "notes"
                : cls.StartsWith("IPF.Journal", StringComparison.OrdinalIgnoreCase) ? "journal"
                : "mail";
            string target = special.TryGetValue(Convert.ToHexString(eid), out var t) ? t
                : (s.GetValueOrDefault(PR_COMMENT_W) as string) == UnfiledMarker ? "@Unfiled"
                : parentPath.Length == 0 && name is "Junk E-mail" or "Junk Email" or "垃圾電子郵件" or "垃圾郵件" ? "@Junk"
                : kind switch { "calendar" => "@Calendar", "contacts" => "@Contacts", "tasks" => "@Tasks", _ => path };
            list.Add(new PstFolder(path, target, kind, eid, count));
            try
            {
                IMAPIFolder child = OpenFolder(eid);
                try { Walk(child, path, list, special); }
                finally { Marshal.ReleaseComObject(child); }
            }
            catch (Exception ex) { Console.WriteLine($"FAILED subfolders of {path} not readable: {ex.Message}"); }   // keep going with the rest
        }
    }

    /// <summary>Rows of a folder's contents table plus each opened message (released after the callback).</summary>
    void ForEachMessage(PstFolder f, uint[] columns, Action<Dictionary<uint, object?>, IMessage> onMessage, Action<Dictionary<uint, object?>, Exception> onError)
    {
        IMAPIFolder folder = OpenFolder(f.EntryId);
        try
        {
            Check(folder.GetContentsTable(MAPI_UNICODE, out IMAPITable contents), "GetContentsTable");
            var rows = QueryAll(contents, new[] { Tags.PR_ENTRYID }.Concat(columns).ToArray());
            Marshal.ReleaseComObject(contents);
            foreach (var r in rows)
            {
                try
                {
                    byte[] eid = (byte[])r[Tags.PR_ENTRYID]!;
                    object obj = WithBinary(eid, (cb, p) =>
                    {
                        Check(folder.OpenEntry(cb, p, IntPtr.Zero, MAPI_BEST_ACCESS, out _, out object o), "OpenEntry(message)");
                        return o;
                    });
                    var msg = (IMessage)obj;
                    try { onMessage(r, msg); }
                    finally { Marshal.ReleaseComObject(msg); }
                }
                catch (Exception ex) { onError(r, ex); }
            }
        }
        finally { Marshal.ReleaseComObject(folder); }
    }

    // named properties this PST has never used come back as property id 0 (read-only lookup): they are not requested at
    // all, one invalid tag would fail the whole GetProps call
    Dictionary<uint, object?> MsgProps(IMessage msg, params uint[] tags) =>
        Get(t => { int hr = msg.GetProps(t, MAPI_UNICODE, out uint c, out IntPtr p); return (hr, c, p); }, tags.Where(x => (x >> 16) != 0).Distinct().ToArray());

    /// <summary>Writes every message of a mail / sticky-note / journal folder as EML into emlDir.</summary>
    public void ExportFolder(PstFolder f, string emlDir, ref int seq, Action<PstMessage> onMessage, Action<string, string, Exception> onError)
    {
        int n = seq;
        ForEachMessage(f, new[] { Tags.PR_SUBJECT_W, Tags.PR_MESSAGE_FLAGS, Tags.PR_MESSAGE_DELIVERY_TIME, Tags.PR_CLIENT_SUBMIT_TIME, Tags.PR_MESSAGE_CLASS_W },
            (r, msg) =>
            {
                string subject = r.GetValueOrDefault(Tags.PR_SUBJECT_W) as string ?? "";
                string cls = r.GetValueOrDefault(Tags.PR_MESSAGE_CLASS_W) as string ?? "IPM.Note";
                string eml = Path.Combine(emlDir, (++n).ToString("D6") + ".eml");
                // UTF-8 for the MIME output: without PR_INTERNET_CPID the converter falls back to the ANSI code page (GB2312 on a
                // Simplified Chinese system), which loses Traditional Chinese characters. Changed in memory only, never saved.
                // the message's own code page first: the HTML body (PR_HTML) is stored in it, the fallback below decodes with it
                int originalCpid = MsgProps(msg, PR_INTERNET_CPID).GetValueOrDefault(PR_INTERNET_CPID) is uint oc ? (int)oc : 0;
                using (var cp = new Block(new List<(uint, object)> { (PR_INTERNET_CPID, 65001) }))
                {
                    msg.SetProps(cp.Count, cp.Ptr, out IntPtr prob);
                    if (prob != IntPtr.Zero) Native.MAPIFreeBuffer(prob);
                }
                Check(Native.SHCreateStreamOnFileEx(eml, STGM_CREATE_WRITE, 0x80, true, IntPtr.Zero, out IStream stm), "create " + eml);
                int hr;
                try
                {
                    hr = Converter().MAPIToMIMEStm(msg, stm, CCSF_SMTP | CCSF_INCLUDE_BCC);
                    if (hr >= 0) stm.Commit(0);
                }
                finally { Marshal.ReleaseComObject(stm); }
                string? note = null;
                int exFixed = 0;
                if (hr < 0)
                {
                    // Outlook's converter refuses some classes (meeting requests without the calendar components, reports):
                    // build the MIME from the message properties instead of losing the message
                    int missing = WriteMimeFromProps(msg, cls, eml, originalCpid);
                    note = $"from MAPI properties (converter 0x{hr:X8})" + (missing > 0 ? $", {missing} attachment(s) not readable" : "");
                }
                else
                {
                    exFixed = FixExchangeAddresses(msg, eml);
                    FixEmbeddedMessages(msg, eml);
                }

                uint flags = r.GetValueOrDefault(Tags.PR_MESSAGE_FLAGS) is uint fl ? fl : 0;
                var received = r.GetValueOrDefault(Tags.PR_MESSAGE_DELIVERY_TIME) as DateTimeOffset?;
                var sent = r.GetValueOrDefault(Tags.PR_CLIENT_SUBMIT_TIME) as DateTimeOffset?;
                // only Drafts / Outbox: an unsent message filed in another folder would otherwise show in Notes' Drafts view as well
                bool draft = f.Kind == "mail" && f.Target is "@Drafts" or "@Outbox";
                bool read = (flags & MSGFLAG_READ) != 0 || f.Kind != "mail";
                onMessage(new PstMessage(eml, f.Target is "@Outbox" ? "@Drafts" : f.Target, subject, read, received, sent, draft, cls, exFixed, note));
            },
            (r, ex) => onError(r.GetValueOrDefault(Tags.PR_MESSAGE_CLASS_W) as string ?? "", r.GetValueOrDefault(Tags.PR_SUBJECT_W) as string ?? "", ex));
        seq = n;
    }

    // ---- Exchange addresses ------------------------------------------------------------------------------------------
    // Mail from an Exchange account stores sender / recipients as X.500 addresses (address type EX). Without an Exchange
    // address book the MIME converter cannot turn them into Internet addresses and writes only the display names
    // ("From: "Alex Chan""). The PST also keeps the SMTP address next to each of them (PR_SENDER_SMTP_ADDRESS,
    // PR_SMTP_ADDRESS on every recipient ...), so for such mail From / Sender / To / Cc / Bcc are rebuilt from MAPI.
    int FixExchangeAddresses(IMessage msg, string eml)
    {
        static MailboxAddress? Box(string? name, string? type, string? addr, string? smtp) =>
            SafeMailbox(name, string.Equals(type, "EX", StringComparison.OrdinalIgnoreCase) ? smtp : addr ?? smtp);
        var p = MsgProps(msg, 0x0C1E001F, 0x0C1F001F, 0x0C1A001F, 0x5D01001F, 0x0064001F, 0x0065001F, 0x0042001F, 0x5D02001F);
        string? S(uint tag) => p.GetValueOrDefault(tag) as string;
        bool anyEx = string.Equals(S(0x0C1E001F), "EX", StringComparison.OrdinalIgnoreCase) || string.Equals(S(0x0064001F), "EX", StringComparison.OrdinalIgnoreCase);
        var recipients = new List<(uint type, MailboxAddress? box, string? name)>();
        if (msg.GetRecipientTable(MAPI_UNICODE, out IMAPITable table) >= 0)
        {
            try
            {
                foreach (var r in QueryAll(table, 0x0C150003, 0x3001001F, 0x3002001F, 0x3003001F, 0x39FE001F))
                {
                    string? type = r.GetValueOrDefault(0x3002001Fu) as string;
                    anyEx |= string.Equals(type, "EX", StringComparison.OrdinalIgnoreCase);
                    recipients.Add((r.GetValueOrDefault(0x0C150003u) is uint t ? t : 1,
                        Box(r.GetValueOrDefault(0x3001001Fu) as string, type, r.GetValueOrDefault(0x3003001Fu) as string, r.GetValueOrDefault(0x39FE001Fu) as string),
                        r.GetValueOrDefault(0x3001001Fu) as string));
                }
            }
            finally { Marshal.ReleaseComObject(table); }
        }
        if (!anyEx) return 0;

        MimeMessage mime;
        mime = MimeMessage.Load(eml);
        int fixedCount = 0;
        var from = Box(S(0x0042001F), S(0x0064001F), S(0x0065001F), S(0x5D02001F)) ?? Box(S(0x0C1A001F), S(0x0C1E001F), S(0x0C1F001F), S(0x5D01001F));
        var sender = Box(S(0x0C1A001F), S(0x0C1E001F), S(0x0C1F001F), S(0x5D01001F));
        if (from != null) { mime.From.Clear(); mime.From.Add(from); fixedCount++; }
        mime.Sender = sender != null && from != null && !string.Equals(sender.Address, from.Address, StringComparison.OrdinalIgnoreCase) ? sender : null;
        foreach (var (type, list) in new[] { (1u, mime.To), (2u, mime.Cc), (3u, mime.Bcc) })
        {
            var mine = recipients.Where(r => r.type == type).ToList();
            if (!mine.Any(r => r.box != null)) continue;   // nothing usable: keep what the converter wrote
            // recipients without an SMTP address (e.g. Exchange distribution lists) keep the converter's entry or their name
            var converter = list.ToList();
            var rebuilt = new List<InternetAddress>();
            foreach (var (_, box, name) in mine)
            {
                if (box != null) { rebuilt.Add(box); fixedCount++; continue; }
                var same = converter.FirstOrDefault(a => string.Equals(a.Name, name, StringComparison.OrdinalIgnoreCase));
                if (same != null) rebuilt.Add(same);
                else if (!string.IsNullOrWhiteSpace(name)) rebuilt.Add(new GroupAddress(name));   // "Name: ;" keeps the name visible
            }
            list.Clear(); list.AddRange(rebuilt);
        }
        using (var fs = File.Create(eml)) mime.WriteTo(fs);
        return fixedCount;
    }
    /// <summary>A mailbox from a MAPI name and address, or null. The address may come as a whole header value
    /// ("\"Name\" &lt;a@b&gt;"): then the part in angle brackets is used. Never throws on malformed input.</summary>
    internal static MailboxAddress? SafeMailbox(string? name, string? address)
    {
        string a = (address ?? "").Trim();
        int lt = a.LastIndexOf('<'), gt = a.LastIndexOf('>');
        if (lt >= 0 && gt > lt) a = a[(lt + 1)..gt].Trim();
        a = a.Trim('\'', '"');
        if (!a.Contains('@')) return null;
        try { return new MailboxAddress(string.IsNullOrWhiteSpace(name) ? a : name.Trim(), a); }
        catch (ParseException) { return null; }
    }

    // ---- contacts, calendar, tasks -----------------------------------------------------------------------------------
    static string? Str(Dictionary<uint, object?> p, uint tag) => p.GetValueOrDefault(tag) is string s && s.Trim().Length > 0 ? s.Trim() : null;

    /// <summary>Outlook contact -> the same field names the NSF -> PST direction uses (ExportTool.ContactFields).</summary>
    public void ExportContacts(PstFolder f, Action<Dictionary<string, string>> onContact, Action<string, List<string>> onGroup, Action<string, Exception> onError)
    {
        uint fileUnder = Named(PSETID_Address, 0x8005, Tags.PT_UNICODE);
        uint email1 = Named(PSETID_Address, 0x8083, Tags.PT_UNICODE), email1Type = Named(PSETID_Address, 0x8082, Tags.PT_UNICODE);
        uint email1Orig = Named(PSETID_Address, 0x8084, Tags.PT_UNICODE), email2Orig = Named(PSETID_Address, 0x8094, Tags.PT_UNICODE);   // original display name: the SMTP address for EX entries
        uint email2 = Named(PSETID_Address, 0x8093, Tags.PT_UNICODE), email2Type = Named(PSETID_Address, 0x8092, Tags.PT_UNICODE);
        uint street = Named(PSETID_Address, 0x8045, Tags.PT_UNICODE), city = Named(PSETID_Address, 0x8046, Tags.PT_UNICODE);
        uint state = Named(PSETID_Address, 0x8047, Tags.PT_UNICODE), zip = Named(PSETID_Address, 0x8048, Tags.PT_UNICODE), country = Named(PSETID_Address, 0x8049, Tags.PT_UNICODE);
        var map = new (string key, uint tag)[]
        {
            ("givenName", 0x3A06001F), ("surname", 0x3A11001F), ("middleName", 0x3A44001F), ("company", 0x3A16001F), ("jobTitle", 0x3A17001F),
            ("department", 0x3A18001F), ("businessPhone", 0x3A08001F), ("homePhone", 0x3A09001F), ("mobilePhone", 0x3A1C001F),
            ("businessFax", 0x3A24001F), ("pager", 0x3A21001F), ("webPage", 0x3A51001F), ("notes", 0x1000001F),
            ("street", 0x3A29001F), ("city", 0x3A27001F), ("state", 0x3A28001F), ("postalCode", 0x3A2A001F), ("country", 0x3A26001F),
        };
        ForEachMessage(f, new[] { Tags.PR_SUBJECT_W, Tags.PR_MESSAGE_CLASS_W },
            (r, msg) =>
            {
                string cls = r.GetValueOrDefault(Tags.PR_MESSAGE_CLASS_W) as string ?? "";
                if (cls.StartsWith("IPM.DistList", StringComparison.OrdinalIgnoreCase))
                {
                    // contact group: members from PidLidDistributionListOneOffMembers (one-off entry ids with name, type, address)
                    uint name = Named(PSETID_Address, 0x8053, Tags.PT_UNICODE), oneOff = Named(PSETID_Address, 0x8054, Tags.PT_MV_BINARY);
                    var g = MsgProps(msg, name, oneOff);
                    var members = (g.GetValueOrDefault(oneOff) as byte[][] ?? Array.Empty<byte[]>()).Select(OneOffMember).Where(m => m != null).Select(m => m!).ToList();
                    onGroup(Str(g, name) ?? Str(r, Tags.PR_SUBJECT_W) ?? "", members);
                    return;
                }
                if (!cls.StartsWith("IPM.Contact", StringComparison.OrdinalIgnoreCase))
                    throw new InvalidOperationException("not a contact (" + cls + ")");
                var tags = map.Select(m => m.tag).Concat(new[] { Tags.PR_DISPLAY_NAME_W, fileUnder, email1, email1Type, email2, email2Type, email1Orig, email2Orig, street, city, state, zip, country }).ToArray();
                var p = MsgProps(msg, tags);
                var fields = new Dictionary<string, string>();
                foreach (var (key, tag) in map) if (Str(p, tag) is string v) fields[key] = v;
                // business address: the named properties Outlook's form shows win over the standard ones
                foreach (var (key, tag) in new[] { ("street", street), ("city", city), ("state", state), ("postalCode", zip), ("country", country) })
                    if (Str(p, tag) is string v) fields[key] = v;
                string? display = Str(p, Tags.PR_DISPLAY_NAME_W) ?? Str(p, fileUnder) ?? Str(r, Tags.PR_SUBJECT_W);
                if (display != null) fields["displayName"] = display;
                // an Exchange (EX) entry keeps an X.500 address; its SMTP address is usually in the original display name
                string? Mail(uint addr, uint type, uint orig) => !string.Equals(Str(p, type), "EX", StringComparison.OrdinalIgnoreCase) ? Str(p, addr)
                    : Str(p, orig) is string o && o.Contains('@') && !o.Contains(' ') ? o : null;
                if (Mail(email1, email1Type, email1Orig) is string e1) fields["email"] = e1;
                if (Mail(email2, email2Type, email2Orig) is string e2) fields[fields.ContainsKey("email") ? "email2" : "email"] = e2;
                onContact(fields);
            },
            (r, ex) => onError(r.GetValueOrDefault(Tags.PR_SUBJECT_W) as string ?? "", ex));
    }

    /// <summary>One-off entry id (MS-OXCDATA 2.2.5.1) -> "Name &lt;address&gt;" (or the address / name alone).</summary>
    static string? OneOffMember(byte[] eid)
    {
        if (eid.Length < 24) return null;
        ushort flags = BitConverter.ToUInt16(eid, 22);
        bool unicode = (flags & 0x8000) != 0;
        var parts = new List<string>();
        int pos = 24;
        for (int i = 0; i < 3 && pos < eid.Length; i++)
        {
            int start = pos;
            if (unicode) { while (pos + 1 < eid.Length && (eid[pos] != 0 || eid[pos + 1] != 0)) pos += 2; parts.Add(Encoding.Unicode.GetString(eid, start, pos - start)); pos += 2; }
            else { while (pos < eid.Length && eid[pos] != 0) pos++; parts.Add(Encoding.Default.GetString(eid, start, pos - start)); pos++; }
        }
        string display = parts.ElementAtOrDefault(0) ?? "", addr = parts.ElementAtOrDefault(2) ?? "";
        if (addr.Length == 0) return display.Length > 0 ? display : null;
        return display.Length > 0 && display != addr ? $"{display} <{addr}>" : addr;
    }

    /// <summary>Outlook appointment -> manifest values: subject, start, end, allDay, location, body, reminder, tentative,
    /// rrule / exdates for repeating appointments (null when the pattern is not supported: then it stays a single entry).</summary>
    public void ExportAppointments(PstFolder f, string attDir, Action<Dictionary<string, object?>> onItem, Action<string, Exception> onError)
    {
        int seq = 0;
        uint start = Named(PSETID_Appointment, 0x820D, Tags.PT_SYSTIME), end = Named(PSETID_Appointment, 0x820E, Tags.PT_SYSTIME);
        uint allDay = Named(PSETID_Appointment, 0x8215, Tags.PT_BOOLEAN), location = Named(PSETID_Appointment, 0x8208, Tags.PT_UNICODE);
        uint busy = Named(PSETID_Appointment, 0x8205, Tags.PT_LONG), recurring = Named(PSETID_Appointment, 0x8223, Tags.PT_BOOLEAN);
        uint recur = Named(PSETID_Appointment, 0x8216, Tags.PT_BINARY);
        uint toAtt = Named(PSETID_Appointment, 0x823B, Tags.PT_UNICODE), ccAtt = Named(PSETID_Appointment, 0x823C, Tags.PT_UNICODE);
        uint remSet = Named(PSETID_Common, 0x8503, Tags.PT_BOOLEAN), remDelta = Named(PSETID_Common, 0x8501, Tags.PT_LONG);
        ForEachMessage(f, new[] { Tags.PR_SUBJECT_W, Tags.PR_MESSAGE_CLASS_W },
            (r, msg) =>
            {
                var p = MsgProps(msg, start, end, allDay, location, busy, recurring, recur, toAtt, ccAtt, remSet, remDelta, 0x1000001F,
                              0x00600040, 0x00610040, 0x0042001F, PR_SENSITIVITY);
                object? V(uint t) => p.TryGetValue(t, out var o) ? o : null;
                var s = V(start) as DateTimeOffset? ?? V(0x00600040) as DateTimeOffset?;
                var e = V(end) as DateTimeOffset? ?? V(0x00610040) as DateTimeOffset? ?? s;
                if (s == null) throw new InvalidOperationException("appointment without a start time");
                var item = new Dictionary<string, object?>
                {
                    ["type"] = "appointment", ["folder"] = f.Target, ["subject"] = Str(r, Tags.PR_SUBJECT_W) ?? "",
                    ["start"] = s, ["end"] = e, ["allDay"] = p.GetValueOrDefault(allDay) is true,
                    ["location"] = Str(p, location),
                    ["tentative"] = p.GetValueOrDefault(busy) is uint b && b == 1,
                    ["reminder"] = p.GetValueOrDefault(remSet) is true && p.GetValueOrDefault(remDelta) is uint d ? (int)d : null,
                    ["private"] = p.GetValueOrDefault(PR_SENSITIVITY) is uint sens && sens >= 2 ? true : null,   // private / confidential
                };
                // attendees are kept as text in the description (no ATTENDEE lines: creating the entry must not send invitations)
                var bodyParts = new List<string>();
                if (Str(p, 0x0042001F) is string org) bodyParts.Add("主持人 / Organizer: " + org);
                if (Str(p, toAtt) is string ta) bodyParts.Add("必須出席 / Required: " + ta);
                if (Str(p, ccAtt) is string ca) bodyParts.Add("可選出席 / Optional: " + ca);
                if (Str(p, 0x1000001F) is string body) { if (bodyParts.Count > 0) bodyParts.Add(""); bodyParts.Add(body); }
                if (bodyParts.Count > 0) item["body"] = string.Join("\r\n", bodyParts);
                if (p.GetValueOrDefault(recurring) is true && p.GetValueOrDefault(recur) is byte[] blob)
                {
                    var rec = Recurrence.Parse(blob);
                    if (rec != null)
                    {
                        item["rrule"] = rec.Value.rrule; item["exdates"] = rec.Value.exdates; item["recurStartOffset"] = rec.Value.startOffset;
                        // modified occurrences: separate entries at their new time (local, as in the PST)
                        if (rec.Value.exceptions.Count > 0)
                            item["modified"] = rec.Value.exceptions.Select(m => new Dictionary<string, object?>
                            {
                                ["start"] = new DateTimeOffset(DateTime.SpecifyKind(m.Start, DateTimeKind.Local)),
                                ["end"] = new DateTimeOffset(DateTime.SpecifyKind(m.End < m.Start ? m.Start : m.End, DateTimeKind.Local)),
                                ["subject"] = m.Subject, ["location"] = m.Location,
                            }).ToList();
                    }
                    else item["recurrenceNote"] = "unsupported pattern";
                }
                var atts = SaveAttachments(msg, Path.Combine(attDir, "cal" + (++seq) + "_" + Guid.NewGuid().ToString("N")[..6]), out int skippedAtt);
                if (atts.Count > 0) item["attachments"] = atts;
                if (skippedAtt > 0) item["attachmentsSkipped"] = skippedAtt;
                onItem(item);
            },
            (r, ex) => onError(r.GetValueOrDefault(Tags.PR_SUBJECT_W) as string ?? "", ex));
    }

    /// <summary>Outlook task -> manifest values: subject, body, start, due, complete, percent.</summary>
    public void ExportTasks(PstFolder f, Action<Dictionary<string, object?>> onItem, Action<string, Exception> onError)
    {
        uint startT = Named(PSETID_Task, 0x8104, Tags.PT_SYSTIME), due = Named(PSETID_Task, 0x8105, Tags.PT_SYSTIME);
        uint complete = Named(PSETID_Task, 0x811C, Tags.PT_BOOLEAN), percent = Named(PSETID_Task, 0x8102, 0x0005);
        ForEachMessage(f, new[] { Tags.PR_SUBJECT_W, Tags.PR_MESSAGE_CLASS_W },
            (r, msg) =>
            {
                uint doneOn = Named(PSETID_Task, 0x810F, Tags.PT_SYSTIME);   // PidLidTaskDateCompleted
                var p = MsgProps(msg, startT, due, complete, percent, 0x1000001F, PR_IMPORTANCE, doneOn);
                onItem(new Dictionary<string, object?>
                {
                    ["type"] = "task", ["folder"] = f.Target, ["subject"] = Str(r, Tags.PR_SUBJECT_W) ?? "",
                    ["start"] = p.GetValueOrDefault(startT) as DateTimeOffset?, ["due"] = p.GetValueOrDefault(due) as DateTimeOffset?,
                    ["complete"] = p.GetValueOrDefault(complete) is true,
                    ["completedOn"] = p.GetValueOrDefault(doneOn) as DateTimeOffset?,
                    ["percent"] = p.GetValueOrDefault(percent) is double d ? (int)Math.Round(d * 100) : null,
                    ["body"] = Str(p, 0x1000001F),
                    ["importance"] = p.GetValueOrDefault(PR_IMPORTANCE) is uint imp && imp != 1 ? (int)imp : null,   // 0 low, 2 high
                });
            },
            (r, ex) => onError(r.GetValueOrDefault(Tags.PR_SUBJECT_W) as string ?? "", ex));
    }
}

/// <summary>PidLidAppointmentRecur (MS-OXOCAL 2.2.1.44) -> iCalendar RRULE + EXDATE (deleted occurrences).
/// Modified single occurrences (exceptions with other times or subjects) are not carried over: they keep the pattern.</summary>
internal static class Recurrence
{
    static readonly DateTime Epoch = new(1601, 1, 1);
    static readonly string[] Days = { "SU", "MO", "TU", "WE", "TH", "FR", "SA" };

    /// <summary>A modified occurrence: its new local start / end, and the subject / location if they were changed.</summary>
    public sealed record Modified(DateTime Start, DateTime End, string? Subject, string? Location);

    public static (string rrule, List<string> exdates, int startOffset, List<Modified> exceptions)? Parse(byte[] b)
    {
        try
        {
            int pos = 0;
            ushort U16() { ushort v = BitConverter.ToUInt16(b, pos); pos += 2; return v; }
            uint U32() { uint v = BitConverter.ToUInt32(b, pos); pos += 4; return v; }
            U16(); U16();                                   // reader / writer version
            ushort freq = U16(), patternType = U16(), calendarType = U16();
            U32();                                          // FirstDateTime
            uint period = U32();
            U32();                                          // SlidingFlag
            uint weekDays = 0, nth = 0, dayOfMonth = 0;
            switch (patternType)
            {
                case 0x0: break;                            // day
                case 0x1: weekDays = U32(); break;          // week
                case 0x2: case 0x4: dayOfMonth = U32(); break;   // month (0x4 = month end)
                case 0x3: weekDays = U32(); nth = U32(); break;  // month nth
                default: return null;                       // Hijri / Hebrew patterns
            }
            if (calendarType != 0 && calendarType != 1) return null;   // only the Gregorian calendar
            uint endType = U32(), count = U32();
            uint firstDow = U32();
            uint deleted = U32();
            var deletedDates = new List<uint>();
            for (int i = 0; i < deleted; i++) deletedDates.Add(U32());
            uint modified = U32();
            pos += 4 * (int)modified;
            uint startDate = U32(), endDate = U32();
            // AppointmentRecurrencePattern
            U32(); uint writer2 = U32();                    // reader / writer version 2
            uint startOffset = U32();                       // minutes after midnight
            U32();                                          // EndTimeOffset
            var exceptions = ReadExceptions(b, ref pos, writer2);

            var rule = new StringBuilder();
            string byDay = string.Join(",", Enumerable.Range(0, 7).Where(i => (weekDays & (1u << i)) != 0).Select(i => Days[i]));
            bool yearly = freq == 0x200D;
            DateTime first = Epoch.AddMinutes(startDate);
            switch (patternType)
            {
                case 0x0: rule.Append($"FREQ=DAILY;INTERVAL={Math.Max(1, period / 1440)}"); break;
                case 0x1: rule.Append($"FREQ=WEEKLY;INTERVAL={Math.Max(1, period)};BYDAY={byDay}"); break;
                case 0x2:
                    rule.Append(yearly ? $"FREQ=YEARLY;INTERVAL={Math.Max(1, period / 12)};BYMONTH={first.Month}" : $"FREQ=MONTHLY;INTERVAL={Math.Max(1, period)}");
                    // Outlook moves day 29-31 to the last day of a shorter month; BYMONTHDAY alone would skip those months
                    rule.Append(dayOfMonth > 28 ? $";BYMONTHDAY={string.Join(",", Enumerable.Range(28, (int)dayOfMonth - 27))};BYSETPOS=-1" : $";BYMONTHDAY={dayOfMonth}");
                    break;
                case 0x4:
                    rule.Append(yearly ? $"FREQ=YEARLY;INTERVAL={Math.Max(1, period / 12)};BYMONTH={first.Month}" : $"FREQ=MONTHLY;INTERVAL={Math.Max(1, period)}");
                    rule.Append(";BYMONTHDAY=-1"); break;
                case 0x3:
                    rule.Append(yearly ? $"FREQ=YEARLY;INTERVAL={Math.Max(1, period / 12)};BYMONTH={first.Month}" : $"FREQ=MONTHLY;INTERVAL={Math.Max(1, period)}");
                    rule.Append($";BYDAY={byDay};BYSETPOS={(nth == 5 ? -1 : (int)nth)}"); break;
            }
            if (endType == 0x2021) rule.Append($";UNTIL={Epoch.AddMinutes(endDate):yyyyMMdd}T235959");
            else if (endType == 0x2022) rule.Append($";COUNT={count}");
            if (firstDow <= 6 && patternType == 0x1) rule.Append($";WKST={Days[firstDow]}");   // matters for multi-week patterns
            // deleted occurrences (this list also holds the original dates of modified ones, which are added back
            // as separate entries at their new time): local date + start time of the series
            var ex = deletedDates.Select(d => Epoch.AddMinutes(d).AddMinutes(startOffset).ToString("yyyyMMdd'T'HHmmss")).ToList();
            return (rule.ToString(), ex, (int)startOffset, exceptions);
        }
        catch { return null; }
    }

    /// <summary>ExceptionInfo (and ExtendedException for Unicode subject / location) of the modified occurrences.</summary>
    static List<Modified> ReadExceptions(byte[] b, ref int pos, uint writer2)
    {
        var list = new List<Modified>();
        try
        {
            int p = pos;
            ushort U16() { ushort v = BitConverter.ToUInt16(b, p); p += 2; return v; }
            uint U32() { uint v = BitConverter.ToUInt32(b, p); p += 4; return v; }
            int n = U16();
            var flags = new List<ushort>();
            for (int i = 0; i < n; i++)
            {
                DateTime s = Epoch.AddMinutes(U32()), e = Epoch.AddMinutes(U32());
                U32();                                       // OriginalStartDate
                ushort f = U16(); flags.Add(f);
                string? subject = null, location = null;
                if ((f & 0x0001) != 0) { U16(); int len = U16(); subject = Encoding.Default.GetString(b, p, len); p += len; }
                if ((f & 0x0002) != 0) U32();                // MeetingType
                if ((f & 0x0004) != 0) U32();                // ReminderDelta
                if ((f & 0x0008) != 0) U32();                // ReminderSet
                if ((f & 0x0010) != 0) { U16(); int len = U16(); location = Encoding.Default.GetString(b, p, len); p += len; }
                if ((f & 0x0020) != 0) U32();                // BusyStatus
                if ((f & 0x0040) != 0) U32();                // Attachment
                if ((f & 0x0080) != 0) U32();                // SubType
                if ((f & 0x0200) != 0) U32();                // AppointmentColor
                list.Add(new Modified(s, e, subject, location));
            }
            p += 4 + (int)BitConverter.ToUInt32(b, p);       // ReservedBlock1
            if (writer2 >= 0x3009)
            {
                // ExtendedException: the Unicode versions of a changed subject / location
                for (int i = 0; i < n; i++)
                {
                    { uint ch = U32(); p += (int)ch; }        // ChangeHighlightSize + ChangeHighlight
                    p += 4 + (int)BitConverter.ToUInt32(b, p); // ReservedBlockEE1
                    if ((flags[i] & 0x0011) != 0)
                    {
                        p += 12;                              // StartDateTime, EndDateTime, OriginalStartDate
                        string? subject = list[i].Subject, location = list[i].Location;
                        if ((flags[i] & 0x0001) != 0) { int len = U16(); subject = Encoding.Unicode.GetString(b, p, 2 * len); p += 2 * len; }
                        if ((flags[i] & 0x0010) != 0) { int len = U16(); location = Encoding.Unicode.GetString(b, p, 2 * len); p += 2 * len; }
                        p += 4 + (int)BitConverter.ToUInt32(b, p);   // ReservedBlockEE2
                        list[i] = list[i] with { Subject = subject, Location = location };
                    }
                }
            }
            pos = p;
        }
        catch { /* keep what was read: the ANSI subject / location, or the times only */ }
        return list;
    }
}
