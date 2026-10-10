// Contacts and appointments for the PST (built from properties, not from MIME).
// Outlook keeps most contact / calendar fields in named properties (MS-OXPROPS): PSETID_Address, PSETID_Appointment,
// PSETID_Common. Their property IDs are store specific and are resolved once with GetIDsFromNames(MAPI_CREATE).
using System.Runtime.InteropServices;
using static PstLite.Consts;
using static PstLite.Props;

namespace PstLite;

internal sealed partial class PstStore
{
    static readonly Guid PSETID_Appointment = new("00062002-0000-0000-C000-000000000046");
    static readonly Guid PSETID_Common = new("00062008-0000-0000-C000-000000000046");
    static readonly Guid PSETID_Address = new("00062004-0000-0000-C000-000000000046");
    readonly Dictionary<(Guid, int), uint> namedIds = new();

    const uint PR_CONTAINER_CLASS_W = 0x3613001F, PR_ICON_INDEX = 0x10800003, PR_BODY_W = 0x1000001F;
    const uint PR_START_DATE = 0x00600040, PR_END_DATE = 0x00610040;

    /// <summary>Tag of a named property (numeric id in a property set), created in this store if needed.</summary>
    uint Named(Guid set, int lid, uint type)
    {
        if (!namedIds.TryGetValue((set, lid), out uint id))
        {
            // MAPINAMEID (x64) = { LPGUID lpguid; ULONG ulKind; [pad]; union { LONG lID; LPWSTR lpwstrName; } } -> 24 bytes
            IntPtr guid = Marshal.AllocHGlobal(16), nameId = Marshal.AllocHGlobal(24), list = Marshal.AllocHGlobal(8);
            try
            {
                Marshal.Copy(set.ToByteArray(), 0, guid, 16);
                Marshal.WriteIntPtr(nameId, 0, guid);
                Marshal.WriteInt32(nameId, 8, 0);          // MNID_ID
                Marshal.WriteInt32(nameId, 12, 0);
                Marshal.WriteInt64(nameId, 16, 0);
                Marshal.WriteInt32(nameId, 16, lid);
                Marshal.WriteIntPtr(list, nameId);
                Check(store!.GetIDsFromNames(1, list, 0x2 /* MAPI_CREATE */, out IntPtr tags), "GetIDsFromNames");
                id = unchecked((uint)Marshal.ReadInt32(tags, 4)) >> 16;
                Native.MAPIFreeBuffer(tags);
                if (id == 0) throw new InvalidOperationException($"named property {set}/{lid:X} not available");
                namedIds[(set, lid)] = id;
            }
            finally { Marshal.FreeHGlobal(list); Marshal.FreeHGlobal(nameId); Marshal.FreeHGlobal(guid); }
        }
        return (id << 16) | type;
    }

    /// <summary>Default Calendar / Contacts folder of the PST: open it if the store already has one, otherwise create it
    /// and register it (PR_IPM_APPOINTMENT_ENTRYID / PR_IPM_CONTACT_ENTRYID on the root folder and the Inbox).</summary>
    IMAPIFolder SpecialFolder(uint propId, string containerClass, string chineseName, string englishName)
    {
        uint tag = Tags.Tag(propId, Tags.PT_BINARY);
        IMAPIFolder inbox = DefaultFolder("@inbox");
        try
        {
            var have = Get(t => { int hr = inbox.GetProps(t, 0, out uint c, out IntPtr p); return (hr, c, p); }, tag, Tags.PR_DISPLAY_NAME_W);
            if (have.GetValueOrDefault(tag) is byte[] existing) return OpenFolder(existing);

            // name the folder in the PST's language (the Inbox name tells which one Outlook used)
            string inboxName = have.GetValueOrDefault(Tags.PR_DISPLAY_NAME_W) as string ?? "";
            string name = inboxName.Any(ch => ch >= 0x2E80) ? chineseName : englishName;
            Check(ipmRoot.CreateFolder(FOLDER_GENERIC, name, null, IntPtr.Zero, OPEN_IF_EXISTS | MAPI_UNICODE, out IMAPIFolder folder), "CreateFolder(" + name + ")");
            using (var cls = new Block(new List<(uint, object)> { (PR_CONTAINER_CLASS_W, containerClass) }))
            {
                Check(folder.SetProps(cls.Count, cls.Ptr, out IntPtr prob), "SetProps(container class)");
                if (prob != IntPtr.Zero) Native.MAPIFreeBuffer(prob);
            }
            folder.SaveChanges(0);
            var fe = Get(t => { int hr = folder.GetProps(t, 0, out uint c, out IntPtr p); return (hr, c, p); }, Tags.PR_ENTRYID);
            byte[] eid = (byte[])fe[Tags.PR_ENTRYID]!;

            Check(store!.OpenEntry(0, IntPtr.Zero, IntPtr.Zero, MAPI_MODIFY | MAPI_BEST_ACCESS, out _, out object rootObj), "OpenEntry(root)");
            var root = (IMAPIFolder)rootObj;
            try
            {
                foreach (var target in new[] { root, inbox })
                {
                    using var b = new Block(new List<(uint, object)> { (tag, eid) });
                    Check(target.SetProps(b.Count, b.Ptr, out IntPtr prob), "SetProps(special folder id)");
                    if (prob != IntPtr.Zero) Native.MAPIFreeBuffer(prob);
                    target.SaveChanges(KEEP_OPEN_READWRITE);
                }
            }
            finally { Marshal.ReleaseComObject(root); }
            return folder;
        }
        finally { Marshal.ReleaseComObject(inbox); }
    }

    void SaveNew(string folderPath, List<(uint, object)> props)
    {
        IMAPIFolder folder = Folder(folderPath);
        Check(folder.CreateMessage(IntPtr.Zero, 0, out IMessage msg), "CreateMessage");
        try
        {
            using var block = new Block(props);
            Check(msg.SetProps(block.Count, block.Ptr, out IntPtr problems), "SetProps");
            if (problems != IntPtr.Zero) Native.MAPIFreeBuffer(problems);
            Check(msg.SaveChanges(0), "SaveChanges");
        }
        finally { Marshal.ReleaseComObject(msg); }
    }

    static void AddText(List<(uint, object)> props, uint tag, string? value)
    {
        if (!string.IsNullOrWhiteSpace(value)) props.Add((tag, value.Trim()));
    }

    /// <summary>Creates an Outlook contact (IPM.Contact). Field names: see the switch below.</summary>
    public void ImportContact(string folderPath, IReadOnlyDictionary<string, string> f)
    {
        string V(string k) => f.TryGetValue(k, out var v) ? v : "";
        string display = V("displayName");
        if (display.Length == 0) display = (V("surname") + V("givenName")).Trim();
        if (display.Length == 0) display = V("email");
        var p = new List<(uint, object)>
        {
            (Tags.PR_MESSAGE_CLASS_W, "IPM.Contact"),
            (Tags.PR_SUBJECT_W, display), (Tags.PR_SUBJECT_PREFIX_W, ""), (Tags.PR_NORMALIZED_SUBJECT_W, display), (PR_CONVERSATION_TOPIC_W, display),
            (Tags.PR_DISPLAY_NAME_W, display),
            (Tags.PR_MESSAGE_FLAGS, (int)MSGFLAG_READ),
            (PR_ICON_INDEX, 0x200),
            (Named(PSETID_Address, 0x8005, Tags.PT_UNICODE), display),                 // FileUnder
        };
        AddText(p, 0x3A06001F, V("givenName"));
        AddText(p, 0x3A11001F, V("surname"));
        AddText(p, 0x3A44001F, V("middleName"));
        AddText(p, 0x3A16001F, V("company"));
        AddText(p, 0x3A17001F, V("jobTitle"));
        AddText(p, 0x3A18001F, V("department"));
        AddText(p, 0x3A08001F, V("businessPhone"));
        AddText(p, 0x3A09001F, V("homePhone"));
        AddText(p, 0x3A1C001F, V("mobilePhone"));
        AddText(p, 0x3A24001F, V("businessFax"));
        AddText(p, 0x3A21001F, V("pager"));
        AddText(p, 0x3A51001F, V("webPage"));
        AddText(p, PR_BODY_W, V("notes"));
        // business address: standard properties + the named ones Outlook shows on the contact form
        AddText(p, 0x3A29001F, V("street")); AddText(p, 0x3A27001F, V("city")); AddText(p, 0x3A28001F, V("state"));
        AddText(p, 0x3A2A001F, V("postalCode")); AddText(p, 0x3A26001F, V("country"));
        AddText(p, Named(PSETID_Address, 0x8045, Tags.PT_UNICODE), V("street"));
        AddText(p, Named(PSETID_Address, 0x8046, Tags.PT_UNICODE), V("city"));
        AddText(p, Named(PSETID_Address, 0x8047, Tags.PT_UNICODE), V("state"));
        AddText(p, Named(PSETID_Address, 0x8048, Tags.PT_UNICODE), V("postalCode"));
        AddText(p, Named(PSETID_Address, 0x8049, Tags.PT_UNICODE), V("country"));
        // e-mail 1 / 2 (Email1: 0x8080..0x8084, Email2: 0x8090..0x8094)
        var emailSlots = new List<int>();
        foreach (var (key, slot, lidBase) in new[] { ("email", 0, 0x8080), ("email2", 1, 0x8090) })
        {
            string addr = V(key).Trim();
            if (addr.Length == 0) continue;
            p.Add((Named(PSETID_Address, lidBase + 0, Tags.PT_UNICODE), $"{display} ({addr})"));   // display name
            p.Add((Named(PSETID_Address, lidBase + 2, Tags.PT_UNICODE), "SMTP"));                  // address type
            p.Add((Named(PSETID_Address, lidBase + 3, Tags.PT_UNICODE), addr));                    // address
            p.Add((Named(PSETID_Address, lidBase + 4, Tags.PT_UNICODE), addr));                    // original display name
            emailSlots.Add(slot);
        }
        if (emailSlots.Count > 0)
        {
            p.Add((Named(PSETID_Address, 0x8028, Tags.PT_MV_LONG), emailSlots.ToArray()));                 // AddressBookProviderEmailList
            p.Add((Named(PSETID_Address, 0x8029, Tags.PT_LONG), emailSlots.Sum(s => 1 << s)));            // AddressBookProviderArrayType
        }
        SaveNew(folderPath, p);
    }

    static readonly uint PR_CONVERSATION_TOPIC_W = Tags.Tag(0x0070, Tags.PT_UNICODE);

    /// <summary>Creates a single (non-recurring) Outlook appointment (IPM.Appointment).</summary>
    public void ImportAppointment(string folderPath, string subject, DateTimeOffset start, DateTimeOffset end, bool allDay,
        string? location, string? body, int? reminderMinutes, bool tentative = false)
    {
        if (end < start) end = start;
        int minutes = (int)Math.Round((end - start).TotalMinutes);
        var p = new List<(uint, object)>
        {
            (Tags.PR_MESSAGE_CLASS_W, "IPM.Appointment"),
            (Tags.PR_SUBJECT_W, subject), (Tags.PR_SUBJECT_PREFIX_W, ""), (Tags.PR_NORMALIZED_SUBJECT_W, subject), (PR_CONVERSATION_TOPIC_W, subject),
            (Tags.PR_MESSAGE_FLAGS, (int)MSGFLAG_READ),
            (PR_ICON_INDEX, 0x400),
            (PR_START_DATE, start), (PR_END_DATE, end),
            (Named(PSETID_Appointment, 0x820D, Tags.PT_SYSTIME), start),   // AppointmentStartWhole
            (Named(PSETID_Appointment, 0x820E, Tags.PT_SYSTIME), end),     // AppointmentEndWhole
            (Named(PSETID_Appointment, 0x8213, Tags.PT_LONG), minutes),     // AppointmentDuration
            (Named(PSETID_Appointment, 0x8205, Tags.PT_LONG), tentative ? 1 : allDay ? 0 : 2),   // BusyStatus: tentative / free / busy
            (Named(PSETID_Appointment, 0x8215, Tags.PT_BOOLEAN), allDay),  // AppointmentSubType (all day)
            (Named(PSETID_Appointment, 0x8217, Tags.PT_LONG), 0),           // AppointmentStateFlags: plain appointment
            (Named(PSETID_Appointment, 0x8223, Tags.PT_BOOLEAN), false),   // Recurring
            (Named(PSETID_Appointment, 0x8235, Tags.PT_SYSTIME), start),   // ClipStart
            (Named(PSETID_Appointment, 0x8236, Tags.PT_SYSTIME), end),     // ClipEnd
            (Named(PSETID_Common, 0x8516, Tags.PT_SYSTIME), start),        // CommonStart
            (Named(PSETID_Common, 0x8517, Tags.PT_SYSTIME), end),          // CommonEnd
            (Named(PSETID_Common, 0x8503, Tags.PT_BOOLEAN), reminderMinutes.HasValue),   // ReminderSet
        };
        AddText(p, Named(PSETID_Appointment, 0x8208, Tags.PT_UNICODE), location);        // Location
        AddText(p, PR_BODY_W, body);
        if (reminderMinutes is int r)
        {
            p.Add((Named(PSETID_Common, 0x8501, Tags.PT_LONG), r));                          // ReminderDelta
            p.Add((Named(PSETID_Common, 0x8502, Tags.PT_SYSTIME), start));                   // ReminderTime
            p.Add((Named(PSETID_Common, 0x8560, Tags.PT_SYSTIME), start.AddMinutes(-r)));   // ReminderSignalTime
        }
        SaveNew(folderPath, p);
    }
}
