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
        SaveNew("@Calendar", meet);

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
