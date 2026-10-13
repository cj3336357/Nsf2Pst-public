// Test helper (Nsf2Pst pst-make-test <pst>): adds items to a PST that the NSF -> PST direction never creates, so that
// PST -> NSF can be tested with them: a repeating appointment (PidLidAppointmentRecur), a mail with Exchange (EX)
// addresses and a body outside GB2312, a task, a sticky note and a journal entry. Fictional data only.
using System.Runtime.InteropServices;
using System.Text;
using static PstLite.Consts;
using static PstLite.Props;

namespace PstLite;

internal sealed partial class PstStore
{
    public void MakeTestItems()
    {
        var local = TimeSpan.FromHours(8);
        DateTimeOffset At(int y, int m, int d, int h, int min) => new(y, m, d, h, min, 0, local);

        // ---- repeating appointment: every Monday and Wednesday, 10 occurrences from 2024-04-01 09:00, 2024-04-03 deleted
        DateTime day0 = new(2024, 4, 1), deleted = new(2024, 4, 3), lastDay = new(2024, 5, 1);
        static uint Minutes(DateTime d) => (uint)(d - new DateTime(1601, 1, 1)).TotalMinutes;
        var blob = new List<byte>();
        void W16(int v) => blob.AddRange(BitConverter.GetBytes((ushort)v));
        void W32(long v) => blob.AddRange(BitConverter.GetBytes((uint)v));
        W16(0x3004); W16(0x3004); W16(0x200B); W16(0x0001); W16(0);   // versions, weekly, week pattern, Gregorian
        W32(0); W32(1); W32(0);                                          // FirstDateTime, every 1 week, SlidingFlag
        W32(0x0A);                                                       // Monday | Wednesday
        W32(0x2022); W32(10); W32(0);                                    // end after 10 occurrences, FirstDOW Sunday
        W32(1); W32(Minutes(deleted));                                   // one deleted occurrence
        W32(0);                                                          // no modified occurrences
        W32(Minutes(day0)); W32(Minutes(lastDay));                       // StartDate, EndDate
        W32(0x3006); W32(0x3009); W32(540); W32(600);                    // AppointmentRecurrencePattern: 09:00 - 10:00
        W16(0); W32(0); W32(0);                                          // no exceptions, reserved blocks
        DateTimeOffset s = At(2024, 4, 1, 9, 0), e = At(2024, 4, 1, 10, 0);
        var appt = new List<(uint, object)>
        {
            (Tags.PR_MESSAGE_CLASS_W, "IPM.Appointment"), (Tags.PR_SUBJECT_W, "測試重複會議 (每週一、三)"),
            (Tags.PR_MESSAGE_FLAGS, (int)MSGFLAG_READ), (PR_START_DATE, s), (PR_END_DATE, e),
            (Named(PSETID_Appointment, 0x820D, Tags.PT_SYSTIME), s), (Named(PSETID_Appointment, 0x820E, Tags.PT_SYSTIME), e),
            (Named(PSETID_Appointment, 0x8213, Tags.PT_LONG), 60), (Named(PSETID_Appointment, 0x8205, Tags.PT_LONG), 2),
            (Named(PSETID_Appointment, 0x8215, Tags.PT_BOOLEAN), false), (Named(PSETID_Appointment, 0x8223, Tags.PT_BOOLEAN), true),
            (Named(PSETID_Appointment, 0x8231, Tags.PT_LONG), 2),                       // RecurrenceType: weekly
            (Named(PSETID_Appointment, 0x8216, Tags.PT_BINARY), blob.ToArray()),       // AppointmentRecur
            (Named(PSETID_Appointment, 0x8208, Tags.PT_UNICODE), "會議室 B"),
            (Named(PSETID_Common, 0x8503, Tags.PT_BOOLEAN), true), (Named(PSETID_Common, 0x8501, Tags.PT_LONG), 15),
            (PR_BODY_W, "重複會議測試：粵語嘅㗎啲。"),
        };
        SaveNew("@Calendar", appt);

        // ---- mail with Exchange addresses and a body that GB2312 cannot hold
        const string senderDn = "/O=EXAMPLECORP/OU=EXCHANGE ADMINISTRATIVE GROUP (FYDIBOHF23SPDLT)/CN=RECIPIENTS/CN=ALEX.CHAN";
        const string rcptDn = "/O=EXAMPLECORP/OU=EXCHANGE ADMINISTRATIVE GROUP (FYDIBOHF23SPDLT)/CN=RECIPIENTS/CN=MARY.WONG";
        IMAPIFolder inbox = Folder("@Inbox");
        Check(inbox.CreateMessage(IntPtr.Zero, 0, out IMessage msg), "CreateMessage");
        try
        {
            DateTimeOffset when = At(2024, 4, 2, 11, 0);
            var p = new List<(uint, object)>
            {
                (Tags.PR_MESSAGE_CLASS_W, "IPM.Note"), (Tags.PR_SUBJECT_W, "Exchange 地址測試 嘅㗎"),
                (PR_BODY_W, "呢封信嘅寄件人同收件人係 Exchange 地址。粵語：嘅㗎啲冇咗佢。"),
                (Tags.PR_MESSAGE_FLAGS, (int)MSGFLAG_READ), (Tags.PR_MESSAGE_DELIVERY_TIME, when), (Tags.PR_CLIENT_SUBMIT_TIME, when),
                (0x0C1E001F, "EX"), (0x0C1F001F, senderDn), (0x0C1A001F, "Alex Chan"), (0x5D01001F, "alex.chan@example.com"),
                (0x0064001F, "EX"), (0x0065001F, senderDn), (0x0042001F, "Alex Chan"), (0x5D02001F, "alex.chan@example.com"),
            };
            using (var b = new Block(p)) { Check(msg.SetProps(b.Count, b.Ptr, out IntPtr prob), "SetProps"); if (prob != IntPtr.Zero) Native.MAPIFreeBuffer(prob); }
            AddRecipient(msg, 1, "Mary Wong", "EX", rcptDn, "mary.wong@example.com");
            AddRecipient(msg, 2, "Ken Lee", "SMTP", "ken.lee@example.net", "ken.lee@example.net");
            Check(msg.SaveChanges(0), "SaveChanges");
        }
        finally { Marshal.ReleaseComObject(msg); }

        // ---- task
        IMAPIFolder tasks = SpecialFolder(0x36D4, "IPF.Task", "工作", "Tasks");
        folders["@Tasks"] = tasks;
        SaveNew("@Tasks", new List<(uint, object)>
        {
            (Tags.PR_MESSAGE_CLASS_W, "IPM.Task"), (Tags.PR_SUBJECT_W, "測試工作：準備季度報告"), (Tags.PR_MESSAGE_FLAGS, (int)MSGFLAG_READ),
            (Named(PSETID_Task, 0x8104, Tags.PT_SYSTIME), At(2024, 4, 8, 0, 0)), (Named(PSETID_Task, 0x8105, Tags.PT_SYSTIME), At(2024, 4, 12, 0, 0)),
            (Named(PSETID_Task, 0x8102, 0x0005), 0.5), (Named(PSETID_Task, 0x8101, Tags.PT_LONG), 1), (Named(PSETID_Task, 0x811C, Tags.PT_BOOLEAN), false),
            (PR_BODY_W, "工作內容：嘅㗎。"),
        });

        // ---- sticky note and journal entry, each in its own folder
        foreach (var (name, cls, msgClass, subject) in new[] { ("記事", "IPF.StickyNote", "IPM.StickyNote", "記事測試 嘅㗎"), ("日誌", "IPF.Journal", "IPM.Activity", "日誌測試：電話") })
        {
            Check(ipmRoot.CreateFolder(FOLDER_GENERIC, name, null, IntPtr.Zero, OPEN_IF_EXISTS | MAPI_UNICODE, out IMAPIFolder f), "CreateFolder(" + name + ")");
            using (var c = new Block(new List<(uint, object)> { (PR_CONTAINER_CLASS_W, cls) })) { f.SetProps(c.Count, c.Ptr, out IntPtr prob); if (prob != IntPtr.Zero) Native.MAPIFreeBuffer(prob); }
            f.SaveChanges(0);
            folders[name] = f;
            SaveNew(name, new List<(uint, object)>
            {
                (Tags.PR_MESSAGE_CLASS_W, msgClass), (Tags.PR_SUBJECT_W, subject), (PR_BODY_W, subject + "：內容。"),
                (Tags.PR_MESSAGE_FLAGS, (int)MSGFLAG_READ), (Tags.PR_MESSAGE_DELIVERY_TIME, At(2024, 4, 3, 15, 0)),
            });
        }
    }

    /// <summary>Adds one recipient with ModifyRecipients (ADRLIST and property arrays in MAPI memory, as MAPI expects).</summary>
    static void AddRecipient(IMessage msg, int type, string name, string addrType, string address, string smtp)
    {
        var values = new (uint tag, object v)[]
        {
            (0x0C150003, type), (0x3001001F, name), (0x3002001F, addrType), (0x3003001F, address), (0x39FE001F, smtp),
        };
        Check(Native.MAPIAllocateBuffer((uint)(24 * values.Length), out IntPtr props), "MAPIAllocateBuffer");
        for (int i = 0; i < values.Length; i++)
        {
            IntPtr p = props + i * 24;
            Marshal.WriteInt32(p, unchecked((int)values[i].tag)); Marshal.WriteInt32(p, 4, 0);
            if (values[i].v is int n) Marshal.WriteInt32(p, 8, n);
            else
            {
                byte[] s = Encoding.Unicode.GetBytes((string)values[i].v + "\0");
                Check(Native.MAPIAllocateMore((uint)s.Length, props, out IntPtr sp), "MAPIAllocateMore");
                Marshal.Copy(s, 0, sp, s.Length);
                Marshal.WriteIntPtr(p, 8, sp);
            }
        }
        Check(Native.MAPIAllocateBuffer(8 + 16, out IntPtr list), "MAPIAllocateBuffer");
        Marshal.WriteInt32(list, 0, 1); Marshal.WriteInt32(list, 4, 0);
        Marshal.WriteInt32(list, 8, 0); Marshal.WriteInt32(list, 12, values.Length); Marshal.WriteIntPtr(list, 16, props);
        try { Check(msg.ModifyRecipients(0x2 /* MODRECIP_ADD */, list), "ModifyRecipients"); }
        finally { Native.MAPIFreeBuffer(Marshal.ReadIntPtr(list, 16)); Native.MAPIFreeBuffer(list); }
    }
}
