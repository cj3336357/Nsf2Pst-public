// Notes name -> SMTP address resolution for the From / To / Cc ... headers of exported mail.
// Internal Notes mail carries Notes names ("Alex_Chan/Sales/ExampleCorp"), which Outlook cannot reply to.
// Sources, in order: pairs learned from the mail file itself (SendTo <-> InetSendTo, From <-> INetFrom, contacts ...),
// then an optional Domino directory (names.nsf) given with --directory.
using MimeKit;

namespace Nsf2Pst;

internal sealed class AddressBook
{
    readonly Dictionary<string, string> map = new();          // normalised Notes name -> SMTP
    readonly Dictionary<string, string?> dirCache = new();
    readonly dynamic? dirView;
    public int Learned => map.Count;
    public int Fixed, Unresolved;
    public readonly HashSet<string> UnresolvedNames = new(StringComparer.OrdinalIgnoreCase);

    public AddressBook(dynamic session, string directory)
    {
        if (directory.Length == 0) return;
        // "server!!names.nsf" or a local file / copy of the directory
        int bang = directory.IndexOf("!!", StringComparison.Ordinal);
        string server = bang >= 0 ? directory[..bang] : "", file = bang >= 0 ? directory[(bang + 2)..] : directory;
        dynamic dir = session.GetDatabase(server, file, false);
        if (!(bool)dir.IsOpen) dir.Open();
        dirView = dir.GetView("($Users)") ?? throw new InvalidOperationException("directory has no ($Users) view: " + directory);
    }

    /// <summary>"CN=Alex Chan/OU=Sales/O=ExampleCorp", "Alex_Chan/Sales/ExampleCorp@ExampleCorp" -> "alex chan/sales/examplecorp".</summary>
    static string Key(string n)
    {
        n = n.Trim().Trim('"');
        n = System.Text.RegularExpressions.Regex.Replace(n, @"(?i)\b(CN|OU|O|C)=", "");
        int pct = n.IndexOf('%'); if (pct >= 0) n = n[..pct];
        int at = n.LastIndexOf('@'); if (at >= 0 && !n[(at + 1)..].Contains('.')) n = n[..at];   // Notes domain, not an Internet domain
        return n.Replace('_', ' ').Trim().ToLowerInvariant();
    }

    /// <summary>Readable form for display names and directory lookups: "Alex Chan/Sales/ExampleCorp".</summary>
    static string Abbreviated(string n)
    {
        n = System.Text.RegularExpressions.Regex.Replace(n.Trim().Trim('"'), @"(?i)\b(CN|OU|O|C)=", "");
        int pct = n.IndexOf('%'); if (pct >= 0) n = n[..pct];
        int at = n.LastIndexOf('@'); if (at >= 0 && !n[(at + 1)..].Contains('.')) n = n[..at];
        return n.Replace('_', ' ').Trim();
    }

    static bool IsSmtp(string a) => System.Text.RegularExpressions.Regex.IsMatch(a.Trim(), @"^[^@\s<>/]+@[^@\s<>/]+\.[^@\s<>/]+$");

    /// <summary>A header address that is a Notes name rather than an Internet address.</summary>
    public static bool IsNotesName(string a) =>
        a.Length > 0 && (!a.Contains('@') || a.Contains('%') || a[..a.LastIndexOf('@')].Contains('/') || !a[(a.LastIndexOf('@') + 1)..].Contains('.'));

    static string[] Values(dynamic doc, string item)
    {
        object v = doc.GetItemValue(item);
        return v is object[] a ? a.Select(x => x?.ToString() ?? "").ToArray() : Array.Empty<string>();
    }

    void Pair(string[] names, string[] addrs)
    {
        for (int i = 0; i < Math.Min(names.Length, addrs.Length); i++)
        {
            string n = names[i], a = addrs[i].Trim();
            if (n.Length == 0 || !IsNotesName(n) || !IsSmtp(a)) continue;
            map.TryAdd(Key(n), a);
        }
    }

    /// <summary>Learns Notes name / Internet address pairs from one document of the mail file.</summary>
    public void Learn(dynamic doc)
    {
        Pair(Values(doc, "SendTo"), Values(doc, "InetSendTo"));
        Pair(Values(doc, "CopyTo"), Values(doc, "InetCopyTo"));
        Pair(Values(doc, "BlindCopyTo"), Values(doc, "InetBlindCopyTo"));
        Pair(Values(doc, "From"), Values(doc, "INetFrom"));
        Pair(Values(doc, "Principal"), Values(doc, "$INetPrincipal"));
        string[] form = Values(doc, "Form");
        if (form.FirstOrDefault() == "Person")
        {
            string[] mail = Values(doc, "InternetAddress");
            if (mail.Length > 0) foreach (string fn in Values(doc, "FullName")) Pair(new[] { fn }, new[] { mail[0] });
        }
    }

    public string? Resolve(string notesName)
    {
        string k = Key(notesName);
        if (map.TryGetValue(k, out var smtp)) return smtp;
        if (dirView == null) return null;
        if (dirCache.TryGetValue(k, out smtp)) return smtp;
        smtp = null;
        try
        {
            dynamic? d = dirView.GetDocumentByKey(Abbreviated(notesName), true);
            if (d != null)
            {
                string[] ia = Values(d, "InternetAddress"); string a = ia.FirstOrDefault() ?? "";
                if (IsSmtp(a)) smtp = a.Trim();
            }
        }
        catch (Exception ex) { if (Environment.GetEnvironmentVariable("NSF2PST_DEBUG") == "1") Console.Error.WriteLine("directory lookup: " + ex.GetType().Name + ": " + ex.Message); }
        dirCache[k] = smtp;
        return smtp;
    }

    /// <summary>Replaces Notes names in the address headers with SMTP addresses (display name = Notes common name).</summary>
    public void Fix(MimeMessage msg)
    {
        foreach (var list in new[] { msg.From, msg.Sender is null ? null : new InternetAddressList { msg.Sender }, msg.ReplyTo, msg.To, msg.Cc, msg.Bcc })
        {
            if (list == null) continue;
            foreach (var mb in list.Mailboxes)
            {
                if (!string.IsNullOrEmpty(mb.Name) && mb.Name.Contains('/') && !mb.Name.Contains('@'))   // "Alex Chan/Sales/ExampleCorp" -> "Alex Chan"
                    mb.Name = Abbreviated(mb.Name).Split('/')[0];
                if (!IsNotesName(mb.Address)) continue;
                string abbr = Abbreviated(mb.Address);
                string common = abbr.Split('/')[0];
                string? smtp = Resolve(mb.Address);
                if (smtp != null)
                {
                    if (string.IsNullOrWhiteSpace(mb.Name) || mb.Name.Contains('/')) mb.Name = common;
                    mb.Address = smtp;
                    Fixed++;
                }
                else
                {
                    if (string.IsNullOrWhiteSpace(mb.Name)) mb.Name = common;
                    Unresolved++; UnresolvedNames.Add(abbr);
                }
            }
        }
    }
}
