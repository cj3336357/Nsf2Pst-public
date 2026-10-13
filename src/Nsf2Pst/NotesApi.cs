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

    /// <summary>Opens a database through the Notes COM classes with a clear message instead of a null reference when the
    /// file is missing, is not a Notes database, is damaged or cannot be opened.</summary>
    /// <summary>Notes keeps database paths in 256-byte buffers (MAXPATH): longer paths fail with "database object is invalid".</summary>
    public const int MaxPathBytes = 250;

    public static dynamic OpenDatabase(dynamic session, string server, string path)
    {
        if (server.Length == 0)
        {
            if (!File.Exists(path)) throw new FileNotFoundException("NSF file not found: " + path);
            if (Encoding.UTF8.GetByteCount(path) > MaxPathBytes)
                throw new PathTooLongException($"the NSF path is too long for Notes (more than about {MaxPathBytes} bytes; use a shorter folder or file name): " + path);
            var head = new byte[2];
            using (var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)) fs.ReadExactly(head);
            if (head[0] != 0x1A || head[1] != 0x00) throw new InvalidDataException("not a Notes database (NSF): " + path);
        }
        dynamic? db = null;
        try
        {
            db = session.GetDatabase(server, path, false);
            if (db != null && !(bool)db.IsOpen) db.Open();
        }
        catch (Exception ex) { throw new IOException("cannot open the NSF (damaged, or no access with this Notes ID): " + path + " - " + ex.Message, ex); }
        if (db == null || !(bool)db.IsOpen) throw new IOException("cannot open the NSF (damaged, or no access with this Notes ID): " + path);
        return db;
    }

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

    /// <summary>UTF-16 -> LMBCS, NUL terminated. OSTranslate takes 16-bit lengths, so the text is translated in pieces
    /// (LMBCS has no shift state, the pieces can be joined); a piece never splits a surrogate pair.</summary>
    public static byte[] Lmbcs(string s)
    {
        const int piece = 8000;   // chars: at most 32,000 UTF-8 bytes, and the LMBCS output fits the 65,535-byte buffer
        var result = new List<byte>(s.Length * 2 + 1);
        for (int i = 0; i < s.Length;)
        {
            int len = Math.Min(piece, s.Length - i);
            if (i + len < s.Length && char.IsHighSurrogate(s[i + len - 1])) len--;
            byte[] u8 = Encoding.UTF8.GetBytes(s, i, len);
            byte[] outBuf = new byte[Math.Min(65535, u8.Length * 3 + 2)];
            ushort n = OSTranslate(OS_TRANSLATE_UTF8_TO_LMBCS, u8, (ushort)u8.Length, outBuf, (ushort)(outBuf.Length - 1));
            result.AddRange(outBuf.AsSpan(0, n).ToArray());
            i += len;
        }
        result.Add(0);
        return result.ToArray();
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

    // ---- PST -> NSF: create mail notes from MIME (the reverse of MIMEEMLExport) -------------------------------------
    [DllImport(Dll)] static extern void NotesTerm();
    /// <summary>Shuts the Notes runtime down so that every database is closed cleanly (call last, after all COM objects are released).</summary>
    public static void Term() { try { NotesTerm(); } catch { } }

    /// <summary>Starts the Notes runtime for a process that uses only the C API (no Notes COM classes).</summary>
    public static void InitCApi()
    {
        string? ini = NotesIni() ?? throw new FileNotFoundException("notes.ini not found");
        Check(NotesInitExtended(2, new[] { Path.GetDirectoryName(ini)!, "=" + ini }), "NotesInitExtended");
    }
    [DllImport(Dll)] static extern ushort NSFNoteCreate(uint hDb, out uint hNote);
    [DllImport(Dll)] static extern ushort NSFNoteUpdate(uint hNote, ushort flags);
    [DllImport(Dll)] static extern ushort NSFNoteClose(uint hNote);
    [DllImport(Dll)] static extern void NSFNoteGetInfo(uint hNote, ushort type, out uint value);
    [DllImport(Dll)] static extern ushort NSFItemSetText(uint hNote, byte[] name, byte[] text, ushort textLen);
    [DllImport(Dll)] static extern ushort MIMEStreamOpen(uint hNote, byte[] itemName, ushort itemNameLen, uint openFlags, out IntPtr hStream);
    [DllImport(Dll)] static extern int MIMEStreamWrite(byte[] data, ushort len, IntPtr hStream);
    [DllImport(Dll)] static extern ushort MIMEStreamItemize(uint hNote, byte[] itemName, ushort itemNameLen, uint flags, IntPtr hStream);
    [DllImport(Dll)] static extern void MIMEStreamClose(IntPtr hStream);
    const uint MIME_STREAM_OPEN_WRITE = 0x2, MIME_STREAM_ITEMIZE_FULL = 0x2 | 0x4;   // ITEMIZE_HEADERS | ITEMIZE_BODY (0x1 = NO_DELIMITER)
    const ushort NOTE_ID = 1;

    [DllImport(Dll)] static extern ushort CalCreateEntry(uint hDb, byte[] icalendar, uint flags, IntPtr hRetUid, IntPtr ctx);

    /// <summary>Creates a calendar entry (VEVENT) or to do (VTODO) from iCalendar text with the Notes calendar API, the same
    /// way the Notes client stores it (repeating entries, alarms ...). The text is passed as UTF-8, or LMBCS when lmbcs is set.</summary>
    public static void CreateCalendarEntry(Db db, string icalendar, bool lmbcs = false)
    {
        byte[] data = lmbcs ? Lmbcs(icalendar) : Encoding.UTF8.GetBytes(icalendar + "\0");
        Check(CalCreateEntry(db.Handle, data, 0, IntPtr.Zero, IntPtr.Zero), "CalCreateEntry");
    }

    /// <summary>An open database handle for creating notes (C API).</summary>
    public sealed class Db : IDisposable
    {
        public uint Handle { get; private set; }
        public Db(string netPath) { Check(NSFDbOpen(Lmbcs(netPath), out uint h), "NSFDbOpen"); Handle = h; }
        public void Dispose() { if (Handle != 0) NSFDbClose(Handle); Handle = 0; }
    }

    /// <summary>Creates a mail note from an EML file: Notes parses the MIME into native MIME items (Body, From, SendTo,
    /// Subject, PostedDate ...); Form is set to Memo. Returns the note ID.</summary>
    public static uint CreateMimeNote(Db db, string emlFile)
    {
        Check(NSFNoteCreate(db.Handle, out uint hNote), "NSFNoteCreate");
        try
        {
            byte[] body = Lmbcs("Body");
            ushort bodyLen = (ushort)(body.Length - 1);
            Check(MIMEStreamOpen(hNote, body, bodyLen, MIME_STREAM_OPEN_WRITE, out IntPtr stream), "MIMEStreamOpen");
            try
            {
                // streamed in 32 KB pieces: a mail with a very large attachment is never held in memory as a whole
                const int chunk = 32 * 1024;
                byte[] part = new byte[chunk];
                using (var fs = new FileStream(emlFile, FileMode.Open, FileAccess.Read, FileShare.Read, chunk))
                {
                    int n;
                    while ((n = fs.Read(part, 0, chunk)) > 0)
                    {
                        byte[] piece = n == chunk ? part : part[..n];
                        int rc = MIMEStreamWrite(piece, (ushort)n, stream);
                        if (rc != 0) throw new InvalidOperationException($"MIMEStreamWrite: {rc}");
                    }
                }
                Check(MIMEStreamItemize(hNote, body, bodyLen, MIME_STREAM_ITEMIZE_FULL, stream), "MIMEStreamItemize");
            }
            finally { MIMEStreamClose(stream); }
            byte[] form = Lmbcs("Form"), memo = Lmbcs("Memo");
            Check(NSFItemSetText(hNote, form, memo, (ushort)(memo.Length - 1)), "NSFItemSetText(Form)");
            Check(NSFNoteUpdate(hNote, 0), "NSFNoteUpdate");
            NSFNoteGetInfo(hNote, NOTE_ID, out uint id);
            return id;
        }
        finally { NSFNoteClose(hNote); }
    }
}
