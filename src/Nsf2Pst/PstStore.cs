// Opens (or creates) a Unicode PST through a temporary MAPI profile, creates folders and imports EML files.
// The temporary profile is deleted on Dispose, so the user's own Outlook profiles are never touched.
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using static PstLite.Consts;
using static PstLite.Props;

namespace PstLite;

internal sealed partial class PstStore : IDisposable
{
    readonly string profileName = "PstLite_" + Guid.NewGuid().ToString("N");
    IProfAdmin? profAdmin;
    IMAPISession? session;
    IMsgStore? store;
    IMAPIFolder ipmRoot = null!;
    readonly Dictionary<string, IMAPIFolder> folders = new(StringComparer.OrdinalIgnoreCase);
    IConverterSession? converter;
    bool profileCreated, mapiInitialized;

    public PstStore(string pstPath)
    {
        try { Open(Path.GetFullPath(pstPath)); }
        catch { Dispose(); throw; }
    }

    void Open(string pstPath)
    {
        Check(Native.MAPIInitialize(IntPtr.Zero), "MAPIInitialize");
        mapiInitialized = true;
        Check(Native.MAPIAdminProfiles(0, out profAdmin), "MAPIAdminProfiles");
        Check(profAdmin.CreateProfile(profileName, IntPtr.Zero, IntPtr.Zero, 0), "CreateProfile");
        profileCreated = true;

        Check(profAdmin.AdminServices(profileName, IntPtr.Zero, IntPtr.Zero, 0, out IMsgServiceAdmin svcAdmin), "AdminServices");
        try
        {
            Check(svcAdmin.CreateMsgService("MSUPST MS", "PstLite", IntPtr.Zero, 0), "CreateMsgService(MSUPST MS)");
            Check(svcAdmin.GetMsgServiceTable(0, out IMAPITable svcTable), "GetMsgServiceTable");
            var rows = QueryAll(svcTable, Tags.PR_SERVICE_UID, Tags.PR_SERVICE_NAME_A);
            Marshal.ReleaseComObject(svcTable);
            if (Environment.GetEnvironmentVariable("PSTLITE_DEBUG") == "1")
                foreach (var r in rows)
                    Console.Error.WriteLine("svc row: " + string.Join(", ", r.Select(kv => $"{kv.Key:X8}={(kv.Value is byte[] b ? Convert.ToHexString(b) : kv.Value)}")));
            // the temporary profile holds only the PST service we just added (Outlook returns PR_SERVICE_NAME only as PT_ERROR here)
            var row = rows.Single(r => r.ContainsKey(Tags.PR_SERVICE_UID) && r[Tags.PR_SERVICE_UID] is byte[]);
            byte[] uid = (byte[])row[Tags.PR_SERVICE_UID]!;

            IntPtr uidPtr = Marshal.AllocHGlobal(uid.Length);
            Marshal.Copy(uid, 0, uidPtr, uid.Length);
            using var cfg = new Block(new List<(uint, object)>
            {
                (Tags.PR_PST_PATH_W, pstPath),
                (Tags.PR_DISPLAY_NAME_W, Path.GetFileNameWithoutExtension(pstPath)),
            });
            try { Check(svcAdmin.ConfigureMsgService(uidPtr, IntPtr.Zero, 0, cfg.Count, cfg.Ptr), "ConfigureMsgService"); }
            finally { Marshal.FreeHGlobal(uidPtr); }
        }
        finally { Marshal.ReleaseComObject(svcAdmin); }

        Check(Native.MAPILogonEx(IntPtr.Zero, profileName, null,
            MAPI_EXTENDED | MAPI_NEW_SESSION | MAPI_EXPLICIT_PROFILE | MAPI_NO_MAIL, out session), "MAPILogonEx");

        Check(session.GetMsgStoresTable(0, out IMAPITable stores), "GetMsgStoresTable");
        var storeRows = QueryAll(stores, Tags.PR_ENTRYID, Tags.PR_DISPLAY_NAME_W);
        Marshal.ReleaseComObject(stores);
        byte[] storeEid = (byte[])storeRows.First()[Tags.PR_ENTRYID]!;
        store = WithBinary(storeEid, (cb, eid) =>
        {
            Check(session.OpenMsgStore(IntPtr.Zero, cb, eid, IntPtr.Zero, MDB_WRITE | MDB_NO_DIALOG | MDB_NO_MAIL, out IMsgStore s), "OpenMsgStore");
            return s;
        });

        var sp = Get(t => { int hr = store.GetProps(t, 0, out uint c, out IntPtr p); return (hr, c, p); }, Tags.PR_IPM_SUBTREE_ENTRYID);
        byte[] ipmEid = (byte[])sp[Tags.PR_IPM_SUBTREE_ENTRYID]!;
        ipmRoot = OpenFolder(ipmEid);
    }

    static T WithBinary<T>(byte[] data, Func<uint, IntPtr, T> use)
    {
        IntPtr p = Marshal.AllocHGlobal(data.Length);
        try { Marshal.Copy(data, 0, p, data.Length); return use((uint)data.Length, p); }
        finally { Marshal.FreeHGlobal(p); }
    }

    IMAPIFolder OpenFolder(byte[] eid) => WithBinary(eid, (cb, p) =>
    {
        Check(store.OpenEntry(cb, p, IntPtr.Zero, MAPI_MODIFY | MAPI_BEST_ACCESS, out _, out object obj), "OpenEntry(folder)");
        return (IMAPIFolder)obj;
    });

    /// <summary>Returns the folder at a backslash-separated path under the PST's top folder, creating missing levels.</summary>
    public IMAPIFolder Folder(string path)
    {
        path = path.Trim('\\');
        if (path.Length == 0) return ipmRoot;
        if (folders.TryGetValue(path, out var f)) return f;
        int cut = path.LastIndexOf('\\');
        if (cut < 0 && path.StartsWith('@')) return folders[path] = DefaultFolder(path);
        IMAPIFolder parent = cut < 0 ? ipmRoot : Folder(path[..cut]);
        string name = cut < 0 ? path : path[(cut + 1)..];
        Check(parent.CreateFolder(FOLDER_GENERIC, name, null, IntPtr.Zero, OPEN_IF_EXISTS | MAPI_UNICODE, out IMAPIFolder child), $"CreateFolder({path})");
        if (path == "Unfiled")   // mail that was in no Notes folder: marked so that PST -> NSF leaves it unfiled again
        {
            using var c = new Block(new List<(uint, object)> { (Tags.Tag(0x3004, Tags.PT_UNICODE), UnfiledMarker) });
            if (child.SetProps(c.Count, c.Ptr, out IntPtr prob) >= 0) child.SaveChanges(KEEP_OPEN_READWRITE);
            if (prob != IntPtr.Zero) Native.MAPIFreeBuffer(prob);
        }
        folders[path] = child;
        return child;
    }

    /// <summary>@Inbox, @Sent, @Deleted, @Outbox, @Drafts -> the PST's own default folders (names follow the Outlook language).</summary>
    IMAPIFolder DefaultFolder(string token)
    {
        byte[]? eid = null;
        switch (token.ToLowerInvariant())
        {
            case "@inbox":
                IntPtr cls = Marshal.StringToHGlobalAnsi("IPM");
                try
                {
                    Check(store!.GetReceiveFolder(cls, 0, out uint cb, out IntPtr p, out IntPtr explicitClass), "GetReceiveFolder");
                    eid = new byte[cb]; Marshal.Copy(p, eid, 0, (int)cb);
                    Native.MAPIFreeBuffer(p); if (explicitClass != IntPtr.Zero) Native.MAPIFreeBuffer(explicitClass);
                }
                finally { Marshal.FreeHGlobal(cls); }
                break;
            case "@sent": eid = StoreEntryId(0x35E4); break;
            case "@deleted": eid = StoreEntryId(0x35E3); break;
            case "@outbox": eid = StoreEntryId(0x35E2); break;
            case "@calendar": return SpecialFolder(0x36D0, "IPF.Appointment", "行事曆", "Calendar");
            case "@contacts": return SpecialFolder(0x36D1, "IPF.Contact", "連絡人", "Contacts");
            case "@drafts": return SpecialFolder(0x36D7, "IPF.Note", "草稿", "Drafts");   // PR_IPM_DRAFTS_ENTRYID (root / Inbox); created when the PST has none
        }
        if (eid == null) throw new InvalidOperationException("default folder not found: " + token);
        return OpenFolder(eid);
    }

    byte[]? StoreEntryId(uint propId)
    {
        uint tag = Tags.Tag(propId, Tags.PT_BINARY);
        var sp = Get(t => { int hr = store!.GetProps(t, 0, out uint c, out IntPtr p); return (hr, c, p); }, tag);
        return sp.GetValueOrDefault(tag) as byte[];
    }

    IConverterSession Converter()
    {
        if (converter != null) return converter;
        Guid clsid = CLSID_IConverterSession, iid = IID_IConverterSession;
        // Click-to-Run Office registers OUTLMIME's classes only in its virtual registry, so neither we nor the converter
        // (which creates CLSID_IMimeMessage etc. internally) can CoCreate them from an outside process. Register their
        // class factories for this process only (CoRegisterClassObject) - no registry change, gone when we exit.
        if (Native.CoCreateInstance(ref clsid, IntPtr.Zero, 1, ref iid, out IntPtr probe) >= 0) Marshal.Release(probe);
        else RegisterOutlMimeClasses();
        Check(Native.CoCreateInstance(ref clsid, IntPtr.Zero, 1, ref iid, out IntPtr p), "CoCreateInstance(IConverterSession)");
        converter = (IConverterSession)Marshal.GetObjectForIUnknown(p);
        Marshal.Release(p);
        return converter;
    }

    [UnmanagedFunctionPointer(CallingConvention.StdCall)]
    delegate int DllGetClassObjectFn(ref Guid clsid, ref Guid iid, out IntPtr ppv);

    // classes served by OUTLMIME.DLL (from the Click-to-Run registry: HKLM\SOFTWARE\Microsoft\Office\ClickToRun\REGISTRY\...\CLSID)
    static readonly string[] OutlMimeClsids =
    {
        "4E3A7680-B77A-11D0-9DA5-00C04FD65685", // CLSID_IConverterSession
        "9EADBD10-447B-4240-A9DD-73FE7C53A981", // CLSID_IMimeInternational
        "9EADBD14-447B-4240-A9DD-73FE7C53A981", // CLSID_IMimeSecurity
        "9EADBD1A-447B-4240-A9DD-73FE7C53A981", // CLSID_IMimeMessage
        "9EADBD20-447B-4240-A9DD-73FE7C53A981", // CLSID_IRangeList
    };
    readonly List<uint> classCookies = new();

    void RegisterOutlMimeClasses()
    {
        string? path = null;
        foreach (string dir in new[] { @"%ProgramFiles%\Microsoft Office\root\Office16", @"%ProgramFiles%\Microsoft Office\Office16" })
        {
            string candidate = Path.Combine(Environment.ExpandEnvironmentVariables(dir), "OUTLMIME.DLL");
            if (File.Exists(candidate)) { path = candidate; break; }
        }
        if (path == null) throw new FileNotFoundException("OUTLMIME.DLL (Outlook MIME converter) not found");
        IntPtr lib = NativeLibrary.Load(path);
        var getClassObject = Marshal.GetDelegateForFunctionPointer<DllGetClassObjectFn>(NativeLibrary.GetExport(lib, "DllGetClassObject"));
        foreach (string s in OutlMimeClsids)
        {
            Guid clsid = new(s), iidCf = IID_IClassFactory;
            Check(getClassObject(ref clsid, ref iidCf, out IntPtr factory), $"OUTLMIME DllGetClassObject({s})");
            try
            {
                Check(Native.CoRegisterClassObject(ref clsid, factory, 1 /* CLSCTX_INPROC_SERVER */, 1 /* REGCLS_MULTIPLEUSE */, out uint cookie),
                    $"CoRegisterClassObject({s})");
                classCookies.Add(cookie);
            }
            finally { Marshal.Release(factory); }
        }
    }

    /// <summary>Imports one EML/MHT file as a message. subject (Unicode, e.g. from Notes) overrides the MIME subject.</summary>
    public void Import(string emlPath, string folderPath, string? subject, bool? read, DateTimeOffset? received, DateTimeOffset? sent)
    {
        IMAPIFolder folder = Folder(folderPath);
        Check(folder.CreateMessage(IntPtr.Zero, 0, out IMessage msg), "CreateMessage");
        try
        {
            Check(Native.SHCreateStreamOnFileEx(Path.GetFullPath(emlPath), STGM_READ, 0, false, IntPtr.Zero, out IStream stm), "open " + emlPath);
            try
            {
                int chr = Converter().MIMEToMAPI(stm, msg, IntPtr.Zero, CCSF_SMTP);
                if (Environment.GetEnvironmentVariable("PSTLITE_DEBUG") == "1")
                {
                    msg.GetPropList(MAPI_UNICODE, out IntPtr list);
                    int n = list == IntPtr.Zero ? -1 : Marshal.ReadInt32(list);
                    var tags = new List<string>();
                    for (int i = 0; i < Math.Max(n, 0); i++) tags.Add(((uint)Marshal.ReadInt32(list, 4 + 4 * i)).ToString("X8"));
                    if (list != IntPtr.Zero) Native.MAPIFreeBuffer(list);
                    uint prBody = Tags.Tag(0x1000, Tags.PT_UNICODE);
                    var subj = Get(t => { int hr = msg.GetProps(t, MAPI_UNICODE, out uint c, out IntPtr pp); return (hr, c, pp); }, Tags.PR_SUBJECT_W, prBody);
                    string body = (subj.GetValueOrDefault(prBody) as string ?? "").Replace("\r", "\\r").Replace("\n", "\\n");
                    Console.Error.WriteLine($"MIMEToMAPI hr=0x{chr:X8} props={n} subject=[{subj.GetValueOrDefault(Tags.PR_SUBJECT_W)}] body=[{body[..Math.Min(200, body.Length)]}]");
                }
                Check(chr, "MIMEToMAPI");
            }
            finally { Marshal.ReleaseComObject(stm); }

            var set = new List<(uint, object)>();
            if (!string.IsNullOrEmpty(subject))
            {
                // drop the MIME-derived prefix/normalized parts so the store derives them from the new subject
                IntPtr del = TagArray(Tags.PR_SUBJECT_PREFIX_W, Tags.PR_NORMALIZED_SUBJECT_W);
                try { msg.DeleteProps(del, out IntPtr prob); if (prob != IntPtr.Zero) Native.MAPIFreeBuffer(prob); }
                finally { Marshal.FreeHGlobal(del); }
                set.Add((Tags.PR_SUBJECT_W, subject));
            }
            var cur = Get(t => { int hr = msg.GetProps(t, 0, out uint c, out IntPtr p); return (hr, c, p); }, Tags.PR_MESSAGE_FLAGS);
            uint flags = cur.TryGetValue(Tags.PR_MESSAGE_FLAGS, out var fv) && fv is uint u ? u : 0;
            flags &= ~MSGFLAG_UNSENT;
            if (read == true) flags |= MSGFLAG_READ;
            if (read == false) flags &= ~MSGFLAG_READ;
            set.Add((Tags.PR_MESSAGE_FLAGS, (int)flags));
            if (received.HasValue) set.Add((Tags.PR_MESSAGE_DELIVERY_TIME, received.Value));
            if (sent.HasValue) set.Add((Tags.PR_CLIENT_SUBMIT_TIME, sent.Value));

            using (var block = new Block(set))
            {
                Check(msg.SetProps(block.Count, block.Ptr, out IntPtr problems), "SetProps");
                if (problems != IntPtr.Zero) Native.MAPIFreeBuffer(problems);
            }
            Check(msg.SaveChanges(0), "SaveChanges");
        }
        finally { Marshal.ReleaseComObject(msg); }
    }

    /// <summary>Prints every folder and message (subject, read flag, delivery time) for verification.</summary>
    public void List(TextWriter w) => ListFolder(w, ipmRoot, "", 0);

    /// <summary>Prints "path TAB items TAB unread" per folder - numbers only, safe to share.</summary>
    public void Counts(TextWriter w) => CountFolder(w, ipmRoot, "");

    void CountFolder(TextWriter w, IMAPIFolder folder, string path)
    {
        Check(folder.GetContentsTable(MAPI_UNICODE, out IMAPITable contents), "GetContentsTable");
        var msgs = QueryAll(contents, Tags.PR_MESSAGE_FLAGS);
        Marshal.ReleaseComObject(contents);
        int unreadCount = msgs.Count(m => m.TryGetValue(Tags.PR_MESSAGE_FLAGS, out var f) && f is uint x && (x & MSGFLAG_READ) == 0);
        if (path.Length > 0) w.WriteLine($"{path}\t{msgs.Count}\t{unreadCount}");
        Check(folder.GetHierarchyTable(MAPI_UNICODE, out IMAPITable hier), "GetHierarchyTable");
        var subs = QueryAll(hier, Tags.PR_ENTRYID, Tags.PR_DISPLAY_NAME_W);
        Marshal.ReleaseComObject(hier);
        foreach (var s in subs)
        {
            IMAPIFolder child = OpenFolder((byte[])s[Tags.PR_ENTRYID]!);
            try { CountFolder(w, child, path.Length == 0 ? (string)s[Tags.PR_DISPLAY_NAME_W]! : path + "\\" + s[Tags.PR_DISPLAY_NAME_W]); }
            finally { Marshal.ReleaseComObject(child); }
        }
    }

    void ListFolder(TextWriter w, IMAPIFolder folder, string path, int depth)
    {
        Check(folder.GetContentsTable(MAPI_UNICODE, out IMAPITable contents), "GetContentsTable");
        uint prBody = Tags.Tag(0x1000, Tags.PT_UNICODE), prHasAttach = Tags.Tag(0x0E1B, 0x000B);
        var msgs = QueryAll(contents, Tags.PR_SUBJECT_W, Tags.PR_MESSAGE_FLAGS, Tags.PR_MESSAGE_DELIVERY_TIME, Tags.PR_SENDER_NAME_W, prBody, prHasAttach);
        Marshal.ReleaseComObject(contents);
        bool details = Environment.GetEnvironmentVariable("PSTLITE_DETAILS") == "1";
        w.WriteLine($"[{(path.Length == 0 ? "(top)" : path)}]  {msgs.Count} item(s)");
        foreach (var m in msgs)
        {
            uint fl = m.TryGetValue(Tags.PR_MESSAGE_FLAGS, out var f) && f is uint x ? x : 0;
            string when = m.TryGetValue(Tags.PR_MESSAGE_DELIVERY_TIME, out var t) && t is DateTimeOffset d ? d.ToString("yyyy-MM-dd HH:mm") : "-";
            w.WriteLine($"    {((fl & MSGFLAG_READ) != 0 ? "read  " : "UNREAD")}  {when}  {m.GetValueOrDefault(Tags.PR_SUBJECT_W)}");
            if (details)
            {
                string body = (m.GetValueOrDefault(prBody) as string ?? "").Replace("\r", " ").Replace("\n", " ").Trim();
                bool att = (fl & 0x10) != 0;   // MSGFLAG_HASATTACH
                w.WriteLine($"            from: {m.GetValueOrDefault(Tags.PR_SENDER_NAME_W)} | attach: {(att ? "yes" : "no")} | body: {body[..Math.Min(70, body.Length)]}");
            }
        }
        Check(folder.GetHierarchyTable(MAPI_UNICODE, out IMAPITable hier), "GetHierarchyTable");
        var subs = QueryAll(hier, Tags.PR_ENTRYID, Tags.PR_DISPLAY_NAME_W);
        Marshal.ReleaseComObject(hier);
        foreach (var s in subs)
        {
            IMAPIFolder child = OpenFolder((byte[])s[Tags.PR_ENTRYID]!);
            try { ListFolder(w, child, path.Length == 0 ? (string)s[Tags.PR_DISPLAY_NAME_W]! : path + "\\" + s[Tags.PR_DISPLAY_NAME_W], depth + 1); }
            finally { Marshal.ReleaseComObject(child); }
        }
    }

    public void Dispose()
    {
        if (converter != null) Marshal.ReleaseComObject(converter);
        foreach (uint cookie in classCookies) Native.CoRevokeClassObject(cookie);
        classCookies.Clear();
        foreach (var f in folders.Values) Marshal.ReleaseComObject(f);
        folders.Clear();
        if (ipmRoot != null) Marshal.ReleaseComObject(ipmRoot);
        if (store != null) { uint fl = 0; store.StoreLogoff(ref fl); Marshal.ReleaseComObject(store); }
        if (session != null) { session.Logoff(IntPtr.Zero, 0, 0); Marshal.ReleaseComObject(session); }
        if (profileCreated && profAdmin != null) profAdmin.DeleteProfile(profileName, 0);
        if (profAdmin != null) Marshal.ReleaseComObject(profAdmin);
        if (mapiInitialized) Native.MAPIUninitialize();
        converter = null; store = null; session = null; profAdmin = null;
        profileCreated = mapiInitialized = false;
    }
}
