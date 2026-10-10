// Notes C API calls (nnotes.dll, 64-bit). Handles are 32-bit DHANDLE/DBHANDLE values, NOTEID is a DWORD.
// Strings passed to the C API are LMBCS (translated from UTF-8 with OSTranslate).
using System.Runtime.InteropServices;
using System.Text;
using Microsoft.Win32;

namespace NsfExport;

internal static class NotesApi
{
    const string Dll = "nnotes.dll";
    const ushort OS_TRANSLATE_LMBCS_TO_UTF8 = 22, OS_TRANSLATE_UTF8_TO_LMBCS = 24;

    [DllImport(Dll)] static extern ushort OSTranslate(ushort mode, byte[] inStr, ushort len, byte[] outStr, ushort outLen);
    [DllImport(Dll)] static extern ushort OSLoadString(IntPtr hModule, ushort errorCode, byte[] buf, ushort bufLen);
    [DllImport(Dll)] static extern ushort OSPathNetConstruct(byte[]? port, byte[] server, byte[] file, byte[] retPath);
    [DllImport(Dll)] static extern ushort NSFDbOpen(byte[] path, out uint hDb);
    [DllImport(Dll)] static extern ushort NSFDbClose(uint hDb);
    [DllImport(Dll)] static extern ushort NSFDbGetUnreadNoteTable(uint hDb, byte[] user, ushort len, bool create, out uint hTable);
    [DllImport(Dll)] static extern ushort NSFDbUpdateUnread(uint hDb, uint hTable);
    [DllImport(Dll)] static extern ushort NSFDbSetUnreadNoteTable(uint hDb, byte[] user, ushort len, bool flush, uint hOriginal, uint hUnread);
    [DllImport(Dll)] static extern bool IDScan(uint hTable, bool first, out uint id);
    [DllImport(Dll)] static extern ushort IDInsert(uint hTable, uint id, out bool inserted);
    [DllImport(Dll)] static extern ushort IDTableCopy(uint hTable, out uint hCopy);
    [DllImport(Dll)] static extern ushort IDDelete(uint hTable, uint id, out bool deleted);
    [DllImport(Dll)] static extern ushort IDDestroyTable(uint hTable);
    [DllImport(Dll)] static extern ushort MIMEEMLExport(byte[] dbNetPath, uint noteId, byte[] fileName);
    [DllImport(Dll)] static extern ushort NotesInitExtended(int argc, [MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.LPStr)] string[] argv);
    [DllImport(Dll)] static extern ushort SECKFMSwitchToIDFile(byte[] idFile, byte[] password, byte[] userName, ushort maxLen, uint flags, IntPtr reserved);

    /// <summary>notes.ini of this Notes install (next to nnotes.dll).</summary>
    public static string? NotesIni() => ProgramDirs().Select(d => Path.Combine(d, "notes.ini")).FirstOrDefault(File.Exists);

    /// <summary>Switches this process to another Notes ID (e.g. the mailbox owner's, to read encrypted mail).
    /// SECKFMSwitchToIDFile may write KeyFileName into notes.ini, so notes.ini is backed up first and restored by
    /// RestoreNotesIni() at the end of the run (and at the next start if a run was killed).</summary>
    public static string SwitchId(string idFile, string password)
    {
        string? ini = NotesIni();
        if (ini == null) throw new FileNotFoundException("notes.ini not found");
        RestoreNotesIni();                       // a backup left by an interrupted run
        File.Copy(ini, ini + ".nsf2pst.bak", true);
        Check(NotesInitExtended(2, new[] { Path.GetDirectoryName(ini)!, "=" + ini }), "NotesInitExtended");
        var name = new byte[256];
        Check(SECKFMSwitchToIDFile(Lmbcs(idFile), Lmbcs(password), name, (ushort)(name.Length - 1), 0, IntPtr.Zero), "SECKFMSwitchToIDFile");
        int len = Array.IndexOf(name, (byte)0); if (len < 0) len = 0;
        var u8 = new byte[len * 3 + 4];
        ushort n = OSTranslate(OS_TRANSLATE_LMBCS_TO_UTF8, name, (ushort)len, u8, (ushort)u8.Length);
        return Encoding.UTF8.GetString(u8, 0, n);
    }

    /// <summary>Puts notes.ini back from the backup taken by SwitchId (no-op if there is none).</summary>
    public static void RestoreNotesIni()
    {
        string? ini = NotesIni();
        if (ini == null || !File.Exists(ini + ".nsf2pst.bak")) return;
        File.Copy(ini + ".nsf2pst.bak", ini, true);
        File.Delete(ini + ".nsf2pst.bak");
    }

    /// <summary>Folder of nnotes.dll: the Notes program folder registered for the Notes COM classes.</summary>
    public static IEnumerable<string> ProgramDirs()
    {
        string? com = Registry.GetValue(@"HKEY_CLASSES_ROOT\CLSID\{29131539-2EED-1069-BF5D-00DD011186B7}\InprocServer32", "", null) as string;
        if (com != null && Path.GetDirectoryName(com) is string d) yield return d;
        yield return @"C:\Lotus\Notes";
        yield return @"C:\Program Files\HCL\Notes";
    }

    public static byte[] Lmbcs(string s)
    {
        byte[] u8 = Encoding.UTF8.GetBytes(s);
        byte[] outBuf = new byte[u8.Length * 3 + 2];
        ushort n = OSTranslate(OS_TRANSLATE_UTF8_TO_LMBCS, u8, (ushort)u8.Length, outBuf, (ushort)(outBuf.Length - 1));
        Array.Resize(ref outBuf, n + 1);   // NUL terminated
        outBuf[n] = 0;
        return outBuf;
    }

    static string Error(ushort status)
    {
        var buf = new byte[256];
        OSLoadString(IntPtr.Zero, (ushort)(status & 0x3FFF), buf, (ushort)buf.Length);
        int len = Array.IndexOf(buf, (byte)0); if (len < 0) len = buf.Length;
        var u8 = new byte[1024];
        ushort n = OSTranslate(OS_TRANSLATE_LMBCS_TO_UTF8, buf, (ushort)len, u8, (ushort)u8.Length);
        return Encoding.UTF8.GetString(u8, 0, n);
    }

    static void Check(ushort status, string what)
    {
        if ((status & 0x3FFF) != 0) throw new InvalidOperationException($"{what}: Notes error {status & 0x3FFF} {Error(status)}");
    }

    public static string NetPath(string server, string file)
    {
        var ret = new byte[1024];
        Check(OSPathNetConstruct(null, Lmbcs(server), Lmbcs(file), ret), "OSPathNetConstruct");
        int len = Array.IndexOf(ret, (byte)0);
        var u8 = new byte[len * 3 + 4];
        ushort n = OSTranslate(OS_TRANSLATE_LMBCS_TO_UTF8, ret, (ushort)len, u8, (ushort)u8.Length);
        return Encoding.UTF8.GetString(u8, 0, n);
    }

    /// <summary>Note IDs that are unread for the given user (canonical name). Empty set if the user has no unread table.</summary>
    public static HashSet<uint> UnreadNoteIds(string netPath, string user)
    {
        var result = new HashSet<uint>();
        Check(NSFDbOpen(Lmbcs(netPath), out uint hDb), "NSFDbOpen");
        try
        {
            byte[] name = Lmbcs(user);
            ushort status = NSFDbGetUnreadNoteTable(hDb, name, (ushort)(name.Length - 1), false, out uint hTable);
            if ((status & 0x3FFF) != 0 || hTable == 0) return result;   // no unread table for this user
            try
            {
                Check(NSFDbUpdateUnread(hDb, hTable), "NSFDbUpdateUnread");
                bool first = true;
                while (IDScan(hTable, first, out uint id)) { first = false; result.Add(id); }
            }
            finally { IDDestroyTable(hTable); }
        }
        finally { NSFDbClose(hDb); }
        return result;
    }

    /// <summary>Test helper: makes exactly the given note IDs unread for the user (all other notes become read).</summary>
    public static void MarkUnread(string netPath, string user, IEnumerable<uint> ids)
    {
        var wanted = new HashSet<uint>(ids);
        Check(NSFDbOpen(Lmbcs(netPath), out uint hDb), "NSFDbOpen");
        try
        {
            byte[] name = Lmbcs(user);
            ushort len = (ushort)(name.Length - 1);
            Check(NSFDbGetUnreadNoteTable(hDb, name, len, true, out uint hOrig), "NSFDbGetUnreadNoteTable");
            Check(NSFDbUpdateUnread(hDb, hOrig), "NSFDbUpdateUnread");
            Check(IDTableCopy(hOrig, out uint hNew), "IDTableCopy");
            var current = new List<uint>();
            bool first = true;
            while (IDScan(hNew, first, out uint cur)) { first = false; current.Add(cur); }
            foreach (uint id in current.Where(c => !wanted.Contains(c))) Check(IDDelete(hNew, id, out _), "IDDelete");
            foreach (uint id in wanted) Check(IDInsert(hNew, id, out _), "IDInsert");
            Check(NSFDbSetUnreadNoteTable(hDb, name, len, true, hOrig, hNew), "NSFDbSetUnreadNoteTable");
            IDDestroyTable(hNew); IDDestroyTable(hOrig);
        }
        finally { NSFDbClose(hDb); }
    }

    /// <summary>Exports one note as MIME (.eml) with Notes' own rich text -> MIME conversion.</summary>
    public static void ExportEml(string netPath, uint noteId, string file)
    {
        Check(MIMEEMLExport(Lmbcs(netPath), noteId, Lmbcs(file)), "MIMEEMLExport");
    }
}
