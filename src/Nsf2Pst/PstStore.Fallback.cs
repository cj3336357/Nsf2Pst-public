// PST -> NSF: MIME built from MAPI properties, for messages Outlook's converter refuses (MAPIToMIMEStm fails, e.g. on
// meeting requests of a PST without the Outlook calendar components, or on delivery reports). Keeps subject, sender,
// recipients, dates, Message-ID, the HTML / plain text body, appointment details and the attachments stored by value.
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Text;
using MimeKit;
using static PstLite.Consts;
using static PstLite.Props;

namespace PstLite;

internal sealed partial class PstStore
{
    const uint PR_HTML = 0x10130102, PR_INTERNET_MESSAGE_ID = 0x1035001F;
    const uint PR_ATTACH_METHOD = 0x37050003, PR_ATTACH_LONG_FILENAME = 0x3707001F, PR_ATTACH_FILENAME = 0x3704001F;
    const uint PR_ATTACH_DATA_BIN = 0x37010102, PR_ATTACH_MIME_TAG = 0x370E001F, PR_ATTACH_CONTENT_ID = 0x3712001F, PR_ATTACH_NUM = 0x0E210003;
    static readonly Guid IID_IStream = new("0000000C-0000-0000-C000-000000000046");

    /// <summary>Writes the message as EML from its properties. Returns the number of attachments that could not be read.</summary>
    int WriteMimeFromProps(IMessage msg, string cls, string eml)
    {
        var p = MsgProps(msg, Tags.PR_SUBJECT_W, PR_BODY_W, PR_HTML, PR_INTERNET_MESSAGE_ID, Tags.PR_CLIENT_SUBMIT_TIME, Tags.PR_MESSAGE_DELIVERY_TIME,
                         0x0042001F, 0x0064001F, 0x0065001F, 0x5D02001F, 0x0C1A001F, 0x0C1E001F, 0x0C1F001F, 0x5D01001F);
        string? S(uint t) => p.GetValueOrDefault(t) as string;
        static MailboxAddress? Box(string? name, string? type, string? addr, string? smtp) =>
            SafeMailbox(name, string.Equals(type, "EX", StringComparison.OrdinalIgnoreCase) ? smtp : addr ?? smtp)
            ?? (name is { Length: > 0 } ? new MailboxAddress(name, "unknown@unknown.invalid") : null);
        var mime = new MimeMessage { Subject = S(Tags.PR_SUBJECT_W) ?? "" };
        if ((Box(S(0x0042001F), S(0x0064001F), S(0x0065001F), S(0x5D02001F)) ?? Box(S(0x0C1A001F), S(0x0C1E001F), S(0x0C1F001F), S(0x5D01001F))) is MailboxAddress from)
            mime.From.Add(from);
        if (msg.GetRecipientTable(MAPI_UNICODE, out IMAPITable table) >= 0)
        {
            try
            {
                foreach (var r in QueryAll(table, 0x0C150003, 0x3001001F, 0x3002001F, 0x3003001F, 0x39FE001F))
                {
                    var box = Box(r.GetValueOrDefault(0x3001001Fu) as string, r.GetValueOrDefault(0x3002001Fu) as string,
                                  r.GetValueOrDefault(0x3003001Fu) as string, r.GetValueOrDefault(0x39FE001Fu) as string);
                    if (box == null) continue;
                    var list = (r.GetValueOrDefault(0x0C150003u) is uint t ? t : 1u) switch { 2 => mime.Cc, 3 => mime.Bcc, _ => mime.To };
                    list.Add(box);
                }
            }
            finally { Marshal.ReleaseComObject(table); }
        }
        if ((p.GetValueOrDefault(Tags.PR_CLIENT_SUBMIT_TIME) ?? p.GetValueOrDefault(Tags.PR_MESSAGE_DELIVERY_TIME)) is DateTimeOffset d) mime.Date = d;
        // without a Message-ID MimeKit would make one with the computer name: use a neutral one instead
        mime.MessageId = S(PR_INTERNET_MESSAGE_ID) is string mid && mid.Trim('<', '>').Contains('@') ? mid.Trim('<', '>') : Guid.NewGuid().ToString("N") + "@nsf2pst.invalid";
        mime.Headers.Add("X-Nsf2Pst-Converted", "from MAPI properties (" + cls + ")");

        var builder = new BodyBuilder();
        string text = S(PR_BODY_W) ?? "";
        // meeting requests / responses: the time and place are in named properties, not in the body
        if (cls.StartsWith("IPM.Schedule.Meeting", StringComparison.OrdinalIgnoreCase))
        {
            var a = MsgProps(msg, Named(PSETID_Appointment, 0x820D, Tags.PT_SYSTIME), Named(PSETID_Appointment, 0x820E, Tags.PT_SYSTIME), Named(PSETID_Appointment, 0x8208, Tags.PT_UNICODE));
            var lines = new List<string>();
            if (a.Values.OfType<DateTimeOffset>().ToList() is { Count: > 0 } times)
                lines.Add("時間 / When: " + string.Join(" - ", times.Select(t => t.LocalDateTime.ToString("yyyy-MM-dd HH:mm"))));
            if (a.Values.OfType<string>().FirstOrDefault() is string where) lines.Add("地點 / Where: " + where);
            if (lines.Count > 0) text = string.Join("\r\n", lines) + "\r\n\r\n" + text;
        }
        builder.TextBody = text;
        if (p.GetValueOrDefault(PR_HTML) is byte[] html && html.Length > 0) builder.HtmlBody = Encoding.UTF8.GetString(html);

        int missing = 0;
        if (msg.GetAttachmentTable(MAPI_UNICODE, out IMAPITable at) >= 0)
        {
            List<Dictionary<uint, object?>> rows;
            try { rows = QueryAll(at, PR_ATTACH_NUM, PR_ATTACH_METHOD); }
            finally { Marshal.ReleaseComObject(at); }
            foreach (var r in rows)
            {
                if (r.GetValueOrDefault(PR_ATTACH_NUM) is not uint num) continue;
                try
                {
                    if (!(r.GetValueOrDefault(PR_ATTACH_METHOD) is uint m && m == 1)) { missing++; continue; }   // only ATTACH_BY_VALUE
                    Check(msg.OpenAttach(num, IntPtr.Zero, 0, out IntPtr pa), "OpenAttach");
                    var att = (IMAPIProp)Marshal.GetObjectForIUnknown(pa);
                    Marshal.Release(pa);
                    try
                    {
                        var ap = Get(t => { int hr = att.GetProps(t, MAPI_UNICODE, out uint c, out IntPtr x); return (hr, c, x); },
                                     PR_ATTACH_LONG_FILENAME, PR_ATTACH_FILENAME, PR_ATTACH_MIME_TAG, PR_ATTACH_CONTENT_ID);
                        string name = ap.GetValueOrDefault(PR_ATTACH_LONG_FILENAME) as string ?? ap.GetValueOrDefault(PR_ATTACH_FILENAME) as string ?? $"attachment{num + 1}";
                        byte[] data = ReadBinaryStream(att, PR_ATTACH_DATA_BIN);
                        var type = ContentType.TryParse(ap.GetValueOrDefault(PR_ATTACH_MIME_TAG) as string ?? "", out var ct) ? ct : new ContentType("application", "octet-stream");
                        string cid = (ap.GetValueOrDefault(PR_ATTACH_CONTENT_ID) as string ?? "").Trim('<', '>');
                        // an attachment the HTML refers to (cid:) stays an inline part of multipart/related
                        bool inline = cid.Length > 0 && builder.HtmlBody != null && builder.HtmlBody.Contains("cid:" + cid, StringComparison.OrdinalIgnoreCase);
                        var part = (inline ? builder.LinkedResources : builder.Attachments).Add(name, data, type);
                        if (cid.Length > 0) part.ContentId = cid;
                    }
                    finally { Marshal.ReleaseComObject(att); }
                }
                catch { missing++; }
            }
        }
        mime.Body = builder.ToMessageBody();
        using (var fs = File.Create(eml)) mime.WriteTo(FormatOptions.Default, fs);
        return missing;
    }

    /// <summary>Reads a binary property through IStream (works for large values that GetProps does not return).</summary>
    static byte[] ReadBinaryStream(IMAPIProp obj, uint tag)
    {
        byte[] iid = IID_IStream.ToByteArray();
        IntPtr piid = Marshal.AllocHGlobal(16);
        try
        {
            Marshal.Copy(iid, 0, piid, 16);
            Check(obj.OpenProperty(tag, piid, 0, 0, out IntPtr ps), "OpenProperty(stream)");
            var stm = (IStream)Marshal.GetObjectForIUnknown(ps);
            Marshal.Release(ps);
            try
            {
                using var ms = new MemoryStream();
                var buf = new byte[65536];
                IntPtr pRead = Marshal.AllocHGlobal(4);
                try
                {
                    while (true)
                    {
                        stm.Read(buf, buf.Length, pRead);
                        int n = Marshal.ReadInt32(pRead);
                        if (n <= 0) break;
                        ms.Write(buf, 0, n);
                    }
                }
                finally { Marshal.FreeHGlobal(pRead); }
                return ms.ToArray();
            }
            finally { Marshal.ReleaseComObject(stm); }
        }
        finally { Marshal.FreeHGlobal(piid); }
    }
}
