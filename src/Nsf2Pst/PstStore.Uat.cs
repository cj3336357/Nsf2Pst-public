// UAT helper (Nsf2Pst pst-make-uat <pst>): adds the item types that only exist as MAPI properties, for testing
// PST -> NSF: repeating appointments with every supported pattern (daily, monthly nth weekday, yearly, month end
// without an end date, weekly with deleted occurrences), a private appointment, a multi-day all-day event, a meeting
// with attendees, completed / undated tasks, a distribution list, a meeting request and a delivery report.
// Mail with HTML, attachments and so on comes from EML files through the normal "import" command. Fictional data only.
using System.Runtime.InteropServices;
using static PstLite.Consts;
using static PstLite.Props;

namespace PstLite;

internal sealed partial class PstStore
{
    /// <summary>PidLidAppointmentRecur (MS-OXOCAL 2.2.1.44) for a pattern, with optional deleted occurrences.</summary>
    static byte[] RecurBlob(ushort freq, ushort patternType, uint period, uint[] specific, uint endType, uint count,
                            DateTime first, DateTime last, int startMin, int endMin, DateTime[] deleted)
    {
        static uint Min(DateTime d) => (uint)(d - new DateTime(1601, 1, 1)).TotalMinutes;
        var b = new List<byte>();
        void W16(int v) => b.AddRange(BitConverter.GetBytes((ushort)v));
        void W32(long v) => b.AddRange(BitConverter.GetBytes((uint)v));
        W16(0x3004); W16(0x3004); W16(freq); W16(patternType); W16(0);
        W32(0); W32(period); W32(0);
        foreach (uint s in specific) W32(s);
        W32(endType); W32(count); W32(0);
        W32(deleted.Length); foreach (var d in deleted) W32(Min(d.Date));
        W32(0);
        W32(Min(first.Date)); W32(endType == 0x2023 ? 0x5AE980DF : Min(last.Date));
        W32(0x3006); W32(0x3009); W32(startMin); W32(endMin);
        W16(0); W32(0); W32(0);
        return b.ToArray();
    }

    /// <summary>One-off entry id (MS-OXCDATA 2.2.5.1), Unicode strings.</summary>
    static byte[] OneOff(string name, string address)
    {
        var b = new List<byte> { 0, 0, 0, 0 };
        b.AddRange(new Guid("A41F2B81-A3BE-1910-9D6E-00DD010F5402").ToByteArray());
        b.AddRange(BitConverter.GetBytes((ushort)0)); b.AddRange(BitConverter.GetBytes((ushort)0x8000));
        foreach (string s in new[] { name, "SMTP", address }) b.AddRange(System.Text.Encoding.Unicode.GetBytes(s + "\0"));
        return b.ToArray();
    }

    /// <summary>Weekly-on-Wednesday pattern with one modified occurrence (ExceptionInfo + ExtendedException, MS-OXOCAL 2.2.1.44.2-4).</summary>
    static byte[] RecurBlobWithException(DateTime first, uint count, int startMin, int endMin, DateTime originalDay, DateTime newStart, string newSubject)
    {
        static uint Min(DateTime d) => (uint)(d - new DateTime(1601, 1, 1)).TotalMinutes;
        var b = new List<byte>();
        void W16(int v) => b.AddRange(BitConverter.GetBytes((ushort)v));
        void W32(long v) => b.AddRange(BitConverter.GetBytes((uint)v));
        DateTime last = first.AddDays(7 * (count - 1));
        W16(0x3004); W16(0x3004); W16(0x200B); W16(0x1); W16(0);
        W32(0); W32(1); W32(0); W32(0x08);                               // every week, Wednesday
        W32(0x2022); W32(count); W32(0);
        W32(1); W32(Min(originalDay));                                    // deleted (= modified) instance: its original date
        W32(1); W32(Min(newStart.Date));                                  // modified instance date
        W32(Min(first)); W32(Min(last));
        W32(0x3006); W32(0x3009); W32(startMin); W32(endMin);
        W16(1);                                                           // one ExceptionInfo
        W32(Min(newStart)); W32(Min(newStart.AddMinutes(endMin - startMin))); W32(Min(originalDay) + startMin);
        W16(0x0001);                                                      // ARO_SUBJECT
        byte[] ansi = System.Text.Encoding.ASCII.GetBytes("moved");
        W16(ansi.Length + 1); W16(ansi.Length); b.AddRange(ansi);
        W32(0);                                                           // ReservedBlock1Size
        W32(4); W32(0);                                                   // ExtendedException: ChangeHighlight
        W32(0);                                                           // ReservedBlockEE1Size
        W32(Min(newStart)); W32(Min(newStart.AddMinutes(endMin - startMin))); W32(Min(originalDay) + startMin);
        W16(newSubject.Length); b.AddRange(System.Text.Encoding.Unicode.GetBytes(newSubject));
        W32(0);                                                           // ReservedBlockEE2Size
        W32(0);                                                           // ReservedBlock2Size
        return b.ToArray();
    }

    /// <summary>
    /// Weekly-on-Wednesday pattern with several moved occurrences, each with its own OverrideFlags. writer2 0x3009 adds the
    /// ExtendedException (Unicode subjects); an older writer (0x3008) has only the ANSI subject, in the system code page.
    /// </summary>
    static byte[] RecurBlobExceptions(DateTime first, uint count, int startMin, int endMin, uint writer2,
                                      params (DateTime originalDay, DateTime newStart, ushort flags, string subject)[] ex)
    {
        static uint Min(DateTime d) => (uint)(d - new DateTime(1601, 1, 1)).TotalMinutes;
        var ansiEnc = Recurrence.Ansi;
        var b = new List<byte>();
        void W16(int v) => b.AddRange(BitConverter.GetBytes((ushort)v));
        void W32(long v) => b.AddRange(BitConverter.GetBytes((uint)v));
        DateTime last = first.AddDays(7 * (count - 1));
        W16(0x3004); W16(0x3004); W16(0x200B); W16(0x1); W16(0);
        W32(0); W32(1); W32(0); W32(0x08);                               // every week, Wednesday
        W32(0x2022); W32(count); W32(0);
        W32(ex.Length); foreach (var e in ex) W32(Min(e.originalDay));   // deleted (= modified) instances
        W32(ex.Length); foreach (var e in ex) W32(Min(e.newStart.Date)); // modified instance dates
        W32(Min(first)); W32(Min(last));
        W32(0x3006); W32(writer2); W32(startMin); W32(endMin);
        W16(ex.Length);
        foreach (var e in ex)
        {
            W32(Min(e.newStart)); W32(Min(e.newStart.AddMinutes(endMin - startMin))); W32(Min(e.originalDay) + startMin);
            W16(e.flags);
            if ((e.flags & 0x0001) != 0) { byte[] a = ansiEnc.GetBytes(e.subject); W16(a.Length + 1); W16(a.Length); b.AddRange(a); }
            if ((e.flags & 0x0100) != 0) W32(3);                          // AppointmentColor
            // 0x0200 (changed body) has no field in ExceptionInfo
        }
        W32(0);                                                           // ReservedBlock1Size
        if (writer2 >= 0x3009)
            foreach (var e in ex)
            {
                W32(4); W32(0);                                           // ChangeHighlight
                W32(0);                                                   // ReservedBlockEE1Size
                if ((e.flags & 0x0011) != 0)
                {
                    W32(Min(e.newStart)); W32(Min(e.newStart.AddMinutes(endMin - startMin))); W32(Min(e.originalDay) + startMin);
                    if ((e.flags & 0x0001) != 0) { W16(e.subject.Length); b.AddRange(System.Text.Encoding.Unicode.GetBytes(e.subject)); }
                    W32(0);                                               // ReservedBlockEE2Size
                }
            }
        W32(0);                                                           // ReservedBlock2Size
        return b.ToArray();
    }

    public void MakeUatItems()
    {
        var local = TimeSpan.FromHours(8);
        DateTimeOffset At(int y, int m, int d, int h, int min) => new(y, m, d, h, min, 0, local);
        uint A(int lid, uint type) => Named(PSETID_Appointment, lid, type);

        List<(uint, object)> Appt(string subject, DateTimeOffset s, DateTimeOffset e, bool allDay, string? location, string body)
        {
            var p = new List<(uint, object)>
            {
                (Tags.PR_MESSAGE_CLASS_W, "IPM.Appointment"), (Tags.PR_SUBJECT_W, subject), (Tags.PR_MESSAGE_FLAGS, (int)MSGFLAG_READ),
                (PR_START_DATE, s), (PR_END_DATE, e), (A(0x820D, Tags.PT_SYSTIME), s), (A(0x820E, Tags.PT_SYSTIME), e),
                (A(0x8205, Tags.PT_LONG), allDay ? 0 : 2), (A(0x8215, Tags.PT_BOOLEAN), allDay), (PR_BODY_W, body),
            };
            if (location != null) p.Add((A(0x8208, Tags.PT_UNICODE), location));
            return p;
        }
        void Repeating(string subject, DateTimeOffset s, DateTimeOffset e, int recurType, byte[] blob, string body)
        {
            var p = Appt(subject, s, e, false, "UAT 會議室", body);
            p.Add((A(0x8223, Tags.PT_BOOLEAN), true)); p.Add((A(0x8231, Tags.PT_LONG), recurType)); p.Add((A(0x8216, Tags.PT_BINARY), blob));
            SaveNew("@Calendar", p);
        }

        // R1 every 2 days, 5 times, 09:30 - 10:00 from 2024-05-06
        Repeating("UAT R1 每兩日 (5 次)", At(2024, 5, 6, 9, 30), At(2024, 5, 6, 10, 0), 1,
            RecurBlob(0x200A, 0x0, 2 * 1440, Array.Empty<uint>(), 0x2022, 5, new(2024, 5, 6), new(2024, 5, 14), 570, 600, Array.Empty<DateTime>()),
            "每兩日一次，共 5 次：5/6 5/8 5/10 5/12 5/14");
        // R2 monthly on the 2nd Tuesday until 2024-08-31, 14:00 - 15:00 from 2024-05-14
        Repeating("UAT R2 每月第二個星期二", At(2024, 5, 14, 14, 0), At(2024, 5, 14, 15, 0), 3,
            RecurBlob(0x200C, 0x3, 1, new uint[] { 0x04, 2 }, 0x2021, 4, new(2024, 5, 14), new(2024, 8, 31), 840, 900, Array.Empty<DateTime>()),
            "5/14 6/11 7/9 8/13");
        // R3 yearly on 20 May, 3 times, 12:00 - 13:00
        Repeating("UAT R3 每年 5 月 20 日 (3 次)", At(2024, 5, 20, 12, 0), At(2024, 5, 20, 13, 0), 4,
            RecurBlob(0x200D, 0x2, 12, new uint[] { 20 }, 0x2022, 3, new(2024, 5, 20), new(2026, 5, 20), 720, 780, Array.Empty<DateTime>()),
            "2024 / 2025 / 2026 年 5 月 20 日");
        // R4 last day of every month, no end date, 17:00 - 17:30 from 2024-05-31
        Repeating("UAT R4 每月最後一日 (無結束日)", At(2024, 5, 31, 17, 0), At(2024, 5, 31, 17, 30), 3,
            RecurBlob(0x200C, 0x4, 1, new uint[] { 31 }, 0x2023, 10, new(2024, 5, 31), new(2024, 5, 31), 1020, 1050, Array.Empty<DateTime>()),
            "無結束日期的重複");
        // R5 every Friday until 2024-06-28, 5/17 and 6/7 deleted
        Repeating("UAT R5 每週五 (刪除兩次)", At(2024, 5, 3, 16, 0), At(2024, 5, 3, 17, 0), 2,
            RecurBlob(0x200B, 0x1, 1, new uint[] { 0x20 }, 0x2021, 9, new(2024, 5, 3), new(2024, 6, 28), 960, 1020, new DateTime[] { new(2024, 5, 17), new(2024, 6, 7) }),
            "5/3 至 6/28 每週五，5/17 及 6/7 取消，共 7 次");

        // R6 every Monday 07:30 (before 08:00 local = the previous day in UTC), 4 times from 2024-06-03
        Repeating("UAT R6 每週一 07:30", At(2024, 6, 3, 7, 30), At(2024, 6, 3, 8, 0), 2,
            RecurBlob(0x200B, 0x1, 1, new uint[] { 0x02 }, 0x2022, 4, new(2024, 6, 3), new(2024, 6, 24), 450, 480, Array.Empty<DateTime>()),
            "應該係 6/3 6/10 6/17 6/24（全部星期一）");
        // R7 monthly on day 31, 5 times from 2024-01-31 (Outlook: the last day of shorter months)
        Repeating("UAT R7 每月 31 號", At(2024, 1, 31, 10, 0), At(2024, 1, 31, 11, 0), 3,
            RecurBlob(0x200C, 0x2, 1, new uint[] { 31 }, 0x2022, 5, new(2024, 1, 31), new(2024, 5, 31), 600, 660, Array.Empty<DateTime>()),
            "應該係 1/31 2/29 3/31 4/30 5/31");
        // R8 all-day, every day until 2024-07-05 (end by date), from 2024-07-01
        {
            var p = Appt("UAT R8 全日每日至 7/5", At(2024, 7, 1, 0, 0), At(2024, 7, 2, 0, 0), true, null, "應該係 7/1 至 7/5 共 5 日");
            p.Add((A(0x8223, Tags.PT_BOOLEAN), true)); p.Add((A(0x8231, Tags.PT_LONG), 1));
            p.Add((A(0x8216, Tags.PT_BINARY), RecurBlob(0x200A, 0x0, 1440, Array.Empty<uint>(), 0x2021, 5, new(2024, 7, 1), new(2024, 7, 5), 0, 1440, Array.Empty<DateTime>())));
            SaveNew("@Calendar", p);
        }
        // R9 every Wednesday 4 times from 2024-06-05 15:00, the 6/12 occurrence moved to Thursday 6/13 16:00 with its own subject
        Repeating("UAT R9 每週三 (一次改咗)", At(2024, 6, 5, 15, 0), At(2024, 6, 5, 16, 0), 2,
            RecurBlobWithException(new(2024, 6, 5), 4, 900, 960, new DateTime(2024, 6, 12), new DateTime(2024, 6, 13, 16, 0, 0), "UAT R9 改期：星期四"),
            "應該係 6/5 6/13(改期) 6/19 6/26");
        // R10 two moved occurrences; the first also has a changed body (flag 0x0200, which has no field in ExceptionInfo)
        Repeating("UAT R10 每週三 (兩次改咗)", At(2024, 9, 4, 10, 0), At(2024, 9, 4, 11, 0), 2,
            RecurBlobExceptions(new(2024, 9, 4), 4, 600, 660, 0x3009,
                (new DateTime(2024, 9, 11), new DateTime(2024, 9, 12, 14, 0, 0), 0x0201, "UAT R10 改期一：星期四"),
                (new DateTime(2024, 9, 18), new DateTime(2024, 9, 20, 9, 0, 0), 0x0001, "UAT R10 改期二：星期五")),
            "應該係 9/4 9/12 14:00(改期一) 9/20 09:00(改期二) 9/25");
        // R11 an older Outlook (writer 0x3008): the changed subject exists only as ANSI text in the system code page
        Repeating("UAT R11 舊版 Outlook", At(2024, 10, 2, 10, 0), At(2024, 10, 2, 11, 0), 2,
            RecurBlobExceptions(new(2024, 10, 2), 3, 600, 660, 0x3008,
                (new DateTime(2024, 10, 9), new DateTime(2024, 10, 10, 15, 0, 0), 0x0001, "R11 改期中文")),
            "應該係 10/2 10/10 15:00(改期中文) 10/16");

        // private appointment, multi-day all-day event
        var priv = Appt("UAT 私人約會：睇醫生", At(2024, 5, 7, 11, 0), At(2024, 5, 7, 12, 0), false, "診所", "私人");
        priv.Add((0x00360003, 2));   // PR_SENSITIVITY = private
        SaveNew("@Calendar", priv);
        SaveNew("@Calendar", Appt("UAT 三日全日活動：展覽", At(2024, 5, 21, 0, 0), At(2024, 5, 24, 0, 0), true, "會展", "5/21 - 5/23 全日"));

        // meeting with attendees (organizer view)
        var meet = Appt("UAT 會議：季度檢討 嘅㗎", At(2024, 5, 9, 15, 0), At(2024, 5, 9, 16, 30), false, "15 樓會議室", "議程：一、二、三");
        meet.Add((A(0x8217, Tags.PT_LONG), 1));   // AppointmentStateFlags: meeting
        meet.Add((A(0x823B, Tags.PT_UNICODE), "Mary Wong; Ken Lee"));
        meet.Add((A(0x823C, Tags.PT_UNICODE), "陳大文"));
        meet.Add((0x0042001F, "Alex Chan"));
        meet.Add((Named(PSETID_Common, 0x8503, Tags.PT_BOOLEAN), true)); meet.Add((Named(PSETID_Common, 0x8501, Tags.PT_LONG), 30));
        // two attachments on the meeting: an agenda with a Chinese name and a 600 KB binary
        string tmp = Path.Combine(Path.GetTempPath(), "nsf2pst_uat_" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(tmp);
        try
        {
            string agenda = Path.Combine(tmp, "a.txt"), data = Path.Combine(tmp, "b.bin");
            File.WriteAllText(agenda, "季度檢討議程：一、業績；二、預算；三、其他。", new System.Text.UTF8Encoding(true));
            var bytes = new byte[600 * 1024]; new Random(11).NextBytes(bytes); File.WriteAllBytes(data, bytes);
            SaveNew("@Calendar", meet, new[] { (agenda, "季度檢討議程 (第二季).txt"), (data, "簡報資料.bin") });
        }
        finally { Directory.Delete(tmp, true); }

        // tasks: completed, undated with high importance
        IMAPIFolder tasks = SpecialFolder(0x36D4, "IPF.Task", "工作", "Tasks");
        folders["@Tasks"] = tasks;
        SaveNew("@Tasks", new List<(uint, object)>
        {
            (Tags.PR_MESSAGE_CLASS_W, "IPM.Task"), (Tags.PR_SUBJECT_W, "UAT 已完成工作"), (Tags.PR_MESSAGE_FLAGS, (int)MSGFLAG_READ),
            (Named(PSETID_Task, 0x8104, Tags.PT_SYSTIME), At(2024, 5, 1, 0, 0)), (Named(PSETID_Task, 0x8105, Tags.PT_SYSTIME), At(2024, 5, 3, 0, 0)),
            (Named(PSETID_Task, 0x8102, 0x0005), 1.0), (Named(PSETID_Task, 0x8101, Tags.PT_LONG), 2), (Named(PSETID_Task, 0x811C, Tags.PT_BOOLEAN), true),
            (Named(PSETID_Task, 0x810F, Tags.PT_SYSTIME), At(2024, 5, 2, 0, 0)), (PR_BODY_W, "第一行\r\n第二行\r\n\r\n第四行"),
        });
        SaveNew("@Tasks", new List<(uint, object)>
        {
            (Tags.PR_MESSAGE_CLASS_W, "IPM.Task"), (Tags.PR_SUBJECT_W, "UAT 無日期工作 (高重要性)"), (Tags.PR_MESSAGE_FLAGS, (int)MSGFLAG_READ),
            (0x00170003, 2), (Named(PSETID_Task, 0x8101, Tags.PT_LONG), 0), (Named(PSETID_Task, 0x811C, Tags.PT_BOOLEAN), false),
        });

        // distribution list in the Contacts folder (not converted: must be reported, not fail the run)
        Folder("@Contacts");
        SaveNew("@Contacts", new List<(uint, object)>
        {
            (Tags.PR_MESSAGE_CLASS_W, "IPM.DistList"), (Tags.PR_SUBJECT_W, "UAT 群組：專案組"), (Tags.PR_MESSAGE_FLAGS, (int)MSGFLAG_READ),
            (0x3001001F, "UAT 群組：專案組"), (Named(PSETID_Address, 0x8053, Tags.PT_UNICODE), "UAT 群組：專案組"),
            (Named(PSETID_Address, 0x8054, Tags.PT_MV_BINARY), new[] { OneOff("陳大文", "david.chan@example.com"), OneOff("Mary Wong", "mary.wong@example.com"), OneOff("Ken Lee", "ken.lee@example.net") }),
        });

        // ---- folder structure: subfolders of default folders, and user folders named like Notes system folders
        DateTimeOffset when0 = At(2024, 5, 10, 9, 0);
        List<(uint, object)> Mail(string subject) => new()
        {
            (Tags.PR_MESSAGE_CLASS_W, "IPM.Note"), (Tags.PR_SUBJECT_W, subject), (PR_BODY_W, "folder test: " + subject),
            (Tags.PR_MESSAGE_FLAGS, (int)MSGFLAG_READ), (Tags.PR_MESSAGE_DELIVERY_TIME, when0), (Tags.PR_CLIENT_SUBMIT_TIME, when0),
            (0x0C1E001F, "SMTP"), (0x0C1F001F, "folders@example.com"), (0x0C1A001F, "Folder Test"),
            (0x0064001F, "SMTP"), (0x0065001F, "folders@example.com"), (0x0042001F, "Folder Test"),
        };
        IMAPIFolder Sub(IMAPIFolder parent, string name)
        {
            Check(parent.CreateFolder(FOLDER_GENERIC, name, null, IntPtr.Zero, OPEN_IF_EXISTS | MAPI_UNICODE, out IMAPIFolder f), "CreateFolder(" + name + ")");
            return f;
        }
        void Put(IMAPIFolder f, string subject)
        {
            Check(f.CreateMessage(IntPtr.Zero, 0, out IMessage m), "CreateMessage");
            try { using var b = new Block(Mail(subject)); Check(m.SetProps(b.Count, b.Ptr, out IntPtr pr), "SetProps"); if (pr != IntPtr.Zero) Native.MAPIFreeBuffer(pr); Check(m.SaveChanges(0), "SaveChanges"); }
            finally { Marshal.ReleaseComObject(m); }
        }
        IMAPIFolder inboxF = Folder("@Inbox");
        IMAPIFolder projA = Sub(inboxF, "專案A");
        Put(projA, "UAT F1 收件匣\\專案A");
        Put(Sub(projA, "子項 (2024)"), "UAT F2 收件匣\\專案A\\子項");
        Put(Sub(Folder("@Sent"), "歸檔"), "UAT F3 寄件備份\\歸檔");
        Put(Sub(Folder("@Deleted"), "舊資料"), "UAT F4 刪除的郵件\\舊資料");
        foreach (string name in new[] { "Inbox", "Deleted Items", "Drafts", "Sent", "Trash", "Calendar", "Junk", "All Documents", "Follow Up" })
            Put(Sub(ipmRoot, name), "UAT N " + name);
        // review findings: folders Notes treats specially, an unsent mail in a normal folder, an Exchange list without SMTP,
        // a hidden Outlook contact folder, a contact with an Exchange address, a meeting request whose HTML is Big5
        Put(Sub(ipmRoot, "(舊郵件)"), "UAT Y1 括號 folder");
        Put(Sub(ipmRoot, "A | B"), "UAT Y2 直線 folder");
        {
            IMAPIFolder y = Sub(ipmRoot, "審查測試");
            Check(y.CreateMessage(IntPtr.Zero, 0, out IMessage un), "CreateMessage");   // a new message is "unsent" by default
            try
            {
                using var b = new Block(new List<(uint, object)> { (Tags.PR_MESSAGE_CLASS_W, "IPM.Note"), (Tags.PR_SUBJECT_W, "UAT Y3 未寄出但喺普通 folder"), (PR_BODY_W, "unsent"), (Tags.PR_MESSAGE_DELIVERY_TIME, when0) });
                Check(un.SetProps(b.Count, b.Ptr, out IntPtr pr), "SetProps"); if (pr != IntPtr.Zero) Native.MAPIFreeBuffer(pr); Check(un.SaveChanges(0), "SaveChanges");
            }
            finally { Marshal.ReleaseComObject(un); }
            Check(y.CreateMessage(IntPtr.Zero, 0, out IMessage dl), "CreateMessage");
            try
            {
                const string listDn = "/O=EXAMPLECORP/OU=EXCHANGE ADMINISTRATIVE GROUP (FYDIBOHF23SPDLT)/CN=RECIPIENTS/CN=ALL-STAFF";
                using (var b = new Block(Mail("UAT Y4 Exchange 群組冇 SMTP"))) { Check(dl.SetProps(b.Count, b.Ptr, out IntPtr pr), "SetProps"); if (pr != IntPtr.Zero) Native.MAPIFreeBuffer(pr); }
                AddRecipient(dl, 1, "Mary Wong", "EX", "/O=EXAMPLECORP/OU=EXCHANGE ADMINISTRATIVE GROUP (FYDIBOHF23SPDLT)/CN=RECIPIENTS/CN=MARY.WONG", "mary.wong@example.com");
                AddRecipient(dl, 1, "全體員工 All Staff", "EX", listDn, "");
                Check(dl.SaveChanges(0), "SaveChanges");
            }
            finally { Marshal.ReleaseComObject(dl); }
        }
        {
            IMAPIFolder hidden = Sub(ipmRoot, "Recipient Cache");
            using (var c = new Block(new List<(uint, object)> { (PR_CONTAINER_CLASS_W, "IPF.Contact.RecipientCache"), (0x10F4000B, true) }))
            { hidden.SetProps(c.Count, c.Ptr, out IntPtr pr); if (pr != IntPtr.Zero) Native.MAPIFreeBuffer(pr); }
            hidden.SaveChanges(0);
            folders["Recipient Cache"] = hidden;
            ImportContact("Recipient Cache", new Dictionary<string, string> { ["displayName"] = "UAT 隱藏快取聯絡人", ["email"] = "cache@example.com" });
            SaveNew("@Contacts", new List<(uint, object)>
            {
                (Tags.PR_MESSAGE_CLASS_W, "IPM.Contact"), (Tags.PR_SUBJECT_W, "UAT Exchange 聯絡人"), (Tags.PR_DISPLAY_NAME_W, "UAT Exchange 聯絡人"),
                (Named(PSETID_Address, 0x8082, Tags.PT_UNICODE), "EX"), (Named(PSETID_Address, 0x8083, Tags.PT_UNICODE), "/O=EXAMPLECORP/CN=RECIPIENTS/CN=EX.CONTACT"),
                (Named(PSETID_Address, 0x8084, Tags.PT_UNICODE), "ex.contact@example.com"), (Named(PSETID_Address, 0x8080, Tags.PT_UNICODE), "UAT Exchange 聯絡人"),
            });
        }
        {
            var big5 = System.Text.CodePagesEncodingProvider.Instance.GetEncoding(950)!;
            SaveNew("@Inbox", new List<(uint, object)>
            {
                (Tags.PR_MESSAGE_CLASS_W, "IPM.Schedule.Meeting.Request"), (Tags.PR_SUBJECT_W, "UAT Y5 Big5 HTML 會議邀請"), (Tags.PR_MESSAGE_FLAGS, (int)MSGFLAG_READ),
                (Tags.PR_MESSAGE_DELIVERY_TIME, when0), (Tags.PR_CLIENT_SUBMIT_TIME, when0), (0x3FDE0003, 950),
                (0x10130102, big5.GetBytes("<html><body><p>繁體 Big5 內文：香港會議室</p></body></html>")),
                (0x0C1E001F, "SMTP"), (0x0C1F001F, "big5@example.com.tw"), (0x0C1A001F, "Big5 Sender"),
            });
        }

        // unusual folder names, an empty folder, a mail without dates, a 5 MB HTML body
        Put(Sub(ipmRoot, "斜線/名稱 A/B"), "UAT X1 斜線 folder");
        Put(Sub(ipmRoot, "尾空格與點. "), "UAT X2 尾空格 folder");
        Put(Sub(ipmRoot, string.Concat(Enumerable.Repeat("很長的資料夾名稱", 15))), "UAT X3 長 folder 名");
        Sub(Sub(ipmRoot, "空資料夾"), "空子資料夾");
        {
            IMAPIFolder x = Sub(ipmRoot, "特殊郵件");
            Check(x.CreateMessage(IntPtr.Zero, 0, out IMessage nd), "CreateMessage");
            try
            {
                using var b = new Block(new List<(uint, object)> { (Tags.PR_MESSAGE_CLASS_W, "IPM.Note"), (Tags.PR_SUBJECT_W, "UAT X4 冇日期"), (PR_BODY_W, "no dates"), (Tags.PR_MESSAGE_FLAGS, 0) });
                Check(nd.SetProps(b.Count, b.Ptr, out IntPtr pr), "SetProps"); if (pr != IntPtr.Zero) Native.MAPIFreeBuffer(pr); Check(nd.SaveChanges(0), "SaveChanges");
            }
            finally { Marshal.ReleaseComObject(nd); }
            Check(x.CreateMessage(IntPtr.Zero, 0, out IMessage big), "CreateMessage");
            try
            {
                var sb = new System.Text.StringBuilder("<html><body>");
                for (int i = 0; i < 60000; i++) sb.Append("<p>第 ").Append(i).Append(" 段：大型 HTML 內文 嘅㗎 測試。</p>");
                sb.Append("</body></html>");
                var p = Mail("UAT X5 5MB HTML 內文"); p.Add((0x10130102, System.Text.Encoding.UTF8.GetBytes(sb.ToString())));   // PR_HTML
                using var b = new Block(p);
                Check(big.SetProps(b.Count, b.Ptr, out IntPtr pr), "SetProps"); if (pr != IntPtr.Zero) Native.MAPIFreeBuffer(pr); Check(big.SaveChanges(0), "SaveChanges");
            }
            finally { Marshal.ReleaseComObject(big); }
        }

        // ---- a mail whose attachment is another Outlook message (ATTACH_EMBEDDED_MSG) with its own file attachment
        Check(inboxF.CreateMessage(IntPtr.Zero, 0, out IMessage outer), "CreateMessage");
        try
        {
            using (var b = new Block(Mail("UAT E1 附件係一封 Outlook 郵件"))) { Check(outer.SetProps(b.Count, b.Ptr, out IntPtr pr), "SetProps"); if (pr != IntPtr.Zero) Native.MAPIFreeBuffer(pr); }
            Check(outer.CreateAttach(IntPtr.Zero, 0, out _, out IntPtr pa), "CreateAttach");
            var att = (IMAPIProp)Marshal.GetObjectForIUnknown(pa); Marshal.Release(pa);
            try
            {
                using (var b = new Block(new List<(uint, object)> { (0x37050003, 5), (0x3001001F, "內層郵件：報價"), (0x370B0003, -1) }))   // ATTACH_EMBEDDED_MSG
                { Check(att.SetProps(b.Count, b.Ptr, out IntPtr pr), "SetProps(attach)"); if (pr != IntPtr.Zero) Native.MAPIFreeBuffer(pr); }
                IntPtr piid = Marshal.AllocHGlobal(16);
                try
                {
                    Marshal.Copy(new Guid("00020307-0000-0000-C000-000000000046").ToByteArray(), 0, piid, 16);   // IID_IMessage
                    Check(att.OpenProperty(0x3701000D, piid, 0, 0x2 | 0x1, out IntPtr pm), "OpenProperty(PR_ATTACH_DATA_OBJ)");
                    var inner = (IMessage)Marshal.GetObjectForIUnknown(pm); Marshal.Release(pm);
                    try
                    {
                        using (var b = new Block(Mail("內層郵件：報價 嘅㗎𨋢")))
                        { Check(inner.SetProps(b.Count, b.Ptr, out IntPtr pr), "SetProps(inner)"); if (pr != IntPtr.Zero) Native.MAPIFreeBuffer(pr); }
                        string innerFile = Path.GetTempFileName();
                        try { File.WriteAllText(innerFile, "內層附件內容 inner attachment", new System.Text.UTF8Encoding(true)); AddAttachment(inner, innerFile, "報價單 嘅㗎𨋢.txt"); }
                        finally { File.Delete(innerFile); }
                        Check(inner.SaveChanges(0), "SaveChanges(inner)");
                    }
                    finally { Marshal.ReleaseComObject(inner); }
                }
                finally { Marshal.FreeHGlobal(piid); }
                Check(att.SaveChanges(0), "SaveChanges(attach)");
            }
            finally { Marshal.ReleaseComObject(att); }
            Check(outer.SaveChanges(0), "SaveChanges");
        }
        finally { Marshal.ReleaseComObject(outer); }

        // meeting request and delivery report in the Inbox
        var when = At(2024, 5, 8, 10, 0);
        SaveNew("@Inbox", new List<(uint, object)>
        {
            (Tags.PR_MESSAGE_CLASS_W, "IPM.Schedule.Meeting.Request"), (Tags.PR_SUBJECT_W, "UAT 會議邀請：培訓"), (Tags.PR_MESSAGE_FLAGS, 0),   // received, unread (a new message is "unsent" otherwise)
            (Tags.PR_MESSAGE_DELIVERY_TIME, when), (Tags.PR_CLIENT_SUBMIT_TIME, when), (PR_BODY_W, "邀請你參加培訓。"),
            (0x0C1E001F, "SMTP"), (0x0C1F001F, "trainer@example.com"), (0x0C1A001F, "Trainer"),
            (0x0064001F, "SMTP"), (0x0065001F, "trainer@example.com"), (0x0042001F, "Trainer"),
            (A(0x820D, Tags.PT_SYSTIME), At(2024, 5, 15, 10, 0)), (A(0x820E, Tags.PT_SYSTIME), At(2024, 5, 15, 12, 0)),
            (A(0x8208, Tags.PT_UNICODE), "培訓室"), (A(0x8217, Tags.PT_LONG), 3),
        });
        SaveNew("@Inbox", new List<(uint, object)>
        {
            (Tags.PR_MESSAGE_CLASS_W, "REPORT.IPM.Note.NDR"), (Tags.PR_SUBJECT_W, "未能傳遞：UAT 測試"), (Tags.PR_MESSAGE_FLAGS, 0),
            (Tags.PR_MESSAGE_DELIVERY_TIME, when), (Tags.PR_CLIENT_SUBMIT_TIME, when),
            (PR_BODY_W, "Delivery has failed to these recipients: nobody@example.invalid"),
            (0x0C1E001F, "SMTP"), (0x0C1F001F, "postmaster@example.com"), (0x0C1A001F, "Mail Delivery System"),
        });
    }
}
