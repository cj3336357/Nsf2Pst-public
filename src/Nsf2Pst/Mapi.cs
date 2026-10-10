// Minimal Extended MAPI interop for writing PST files (x64 Outlook / OLMAPI32.DLL).
// Interfaces are declared in the vtable order of the Windows SDK / Outlook MAPI headers (mapidefs.h, mapix.h, mimeole.h).
// Every method returns the raw HRESULT ([PreserveSig]); complex parameters are passed as IntPtr and marshalled by hand.
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;

namespace PstLite;

internal static class Native
{
    const string Dll = "OLMAPI32.DLL";

    [DllImport(Dll)] public static extern int MAPIInitialize(IntPtr lpMapiInit);
    [DllImport(Dll)] public static extern void MAPIUninitialize();
    [DllImport(Dll)] public static extern int MAPIAdminProfiles(uint ulFlags, out IProfAdmin lppProfAdmin);
    [DllImport(Dll, CharSet = CharSet.Ansi)]
    public static extern int MAPILogonEx(IntPtr ulUIParam, string lpszProfileName, string? lpszPassword, uint flFlags, out IMAPISession lppSession);
    [DllImport(Dll)] public static extern int MAPIFreeBuffer(IntPtr lpBuffer);
    [DllImport(Dll)] public static extern int MAPIAllocateBuffer(uint cbSize, out IntPtr lppBuffer);
    [DllImport(Dll)] public static extern void FreeProws(IntPtr lpRows);
    [DllImport(Dll, EntryPoint = "DllGetClassObject")]
    public static extern int OlmapiDllGetClassObject(ref Guid rclsid, ref Guid riid, out IClassFactory ppv);

    [DllImport("shlwapi.dll", CharSet = CharSet.Unicode, PreserveSig = true)]
    public static extern int SHCreateStreamOnFileEx(string pszFile, uint grfMode, uint dwAttributes, bool fCreate, IntPtr pstmTemplate, out IStream ppstm);

    [DllImport("ole32.dll")]
    public static extern int CoInitializeEx(IntPtr reserved, uint coInit);

    [DllImport("ole32.dll")]
    public static extern int CoRegisterClassObject(ref Guid rclsid, IntPtr pUnk, uint dwClsContext, uint flags, out uint lpdwRegister);

    [DllImport("ole32.dll")]
    public static extern int CoRevokeClassObject(uint dwRegister);

    [DllImport("ole32.dll")]
    public static extern int CoCreateInstance(ref Guid rclsid, IntPtr pUnkOuter, uint dwClsContext, ref Guid riid, out IntPtr ppv);
}

internal static class Tags
{
    // property types
    public const uint PT_LONG = 0x0003, PT_UNICODE = 0x001F, PT_SYSTIME = 0x0040, PT_BINARY = 0x0102, PT_STRING8 = 0x001E;
    public const uint PT_BOOLEAN = 0x000B, PT_MV_LONG = 0x1003;
    public static uint Tag(uint id, uint type) => (id << 16) | type;

    public static readonly uint PR_ENTRYID = Tag(0x0FFF, PT_BINARY);
    public static readonly uint PR_DISPLAY_NAME_W = Tag(0x3001, PT_UNICODE);
    public static readonly uint PR_SERVICE_UID = Tag(0x3D0C, PT_BINARY);
    public static readonly uint PR_SERVICE_NAME_A = Tag(0x3D00, PT_STRING8);
    public static readonly uint PR_PST_PATH_W = Tag(0x6700, PT_UNICODE);
    public static readonly uint PR_IPM_SUBTREE_ENTRYID = Tag(0x35E0, PT_BINARY);
    public static readonly uint PR_SUBJECT_W = Tag(0x0037, PT_UNICODE);
    public static readonly uint PR_SUBJECT_PREFIX_W = Tag(0x003D, PT_UNICODE);
    public static readonly uint PR_NORMALIZED_SUBJECT_W = Tag(0x0E1D, PT_UNICODE);
    public static readonly uint PR_MESSAGE_FLAGS = Tag(0x0E07, PT_LONG);
    public static readonly uint PR_MESSAGE_DELIVERY_TIME = Tag(0x0E06, PT_SYSTIME);
    public static readonly uint PR_CLIENT_SUBMIT_TIME = Tag(0x0039, PT_SYSTIME);
    public static readonly uint PR_CONTENT_COUNT = Tag(0x3602, PT_LONG);
    public static readonly uint PR_MESSAGE_CLASS_W = Tag(0x001A, PT_UNICODE);
    public static readonly uint PR_SENDER_NAME_W = Tag(0x0C1A, PT_UNICODE);
}

internal static class Consts
{
    public const uint MAPI_UNICODE = 0x80000000, MAPI_MODIFY = 0x1, MAPI_BEST_ACCESS = 0x10;
    public const uint MAPI_EXTENDED = 0x20, MAPI_NEW_SESSION = 0x2, MAPI_NO_MAIL = 0x8000, MAPI_EXPLICIT_PROFILE = 0x10;
    public const uint MDB_NO_DIALOG = 0x1, MDB_WRITE = 0x4, MDB_NO_MAIL = 0x80;
    public const uint FOLDER_GENERIC = 1, OPEN_IF_EXISTS = 0x1;
    public const uint MSGFLAG_READ = 0x1, MSGFLAG_UNSENT = 0x8;
    public const uint CCSF_SMTP = 0x0002;
    public const uint STGM_READ = 0;
    public const uint KEEP_OPEN_READWRITE = 0x4;
    public const int MAPI_E_NOT_FOUND = unchecked((int)0x8004010F);
    public static Guid IID_IMAPIFolder = new("0002030C-0000-0000-C000-000000000046");
    public static Guid CLSID_IConverterSession = new("4E3A7680-B77A-11D0-9DA5-00C04FD65685");
    public static Guid IID_IConverterSession = new("4B401570-B77B-11D0-9DA5-00C04FD65685");
    public static Guid IID_IClassFactory = new("00000001-0000-0000-C000-000000000046");
}

[ComImport, Guid("00000001-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
internal interface IClassFactory
{
    [PreserveSig] int CreateInstance(IntPtr pUnkOuter, ref Guid riid, out IntPtr ppvObject);
    [PreserveSig] int LockServer(bool fLock);
}

[ComImport, Guid("0002031C-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
internal interface IProfAdmin
{
    [PreserveSig] int GetLastError(int hr, uint flags, out IntPtr err);
    [PreserveSig] int GetProfileTable(uint flags, out IMAPITable table);
    [PreserveSig] int CreateProfile([MarshalAs(UnmanagedType.LPStr)] string name, IntPtr pw, IntPtr uiParam, uint flags);
    [PreserveSig] int DeleteProfile([MarshalAs(UnmanagedType.LPStr)] string name, uint flags);
    [PreserveSig] int ChangeProfilePassword(IntPtr a, IntPtr b, IntPtr c, uint d);
    [PreserveSig] int CopyProfile(IntPtr a, IntPtr b, IntPtr c, IntPtr d, uint e);
    [PreserveSig] int RenameProfile(IntPtr a, IntPtr b, IntPtr c, IntPtr d, uint e);
    [PreserveSig] int SetDefaultProfile(IntPtr a, uint b);
    [PreserveSig] int AdminServices([MarshalAs(UnmanagedType.LPStr)] string name, IntPtr pw, IntPtr uiParam, uint flags, out IMsgServiceAdmin admin);
}

[ComImport, Guid("0002031D-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
internal interface IMsgServiceAdmin
{
    [PreserveSig] int GetLastError(int hr, uint flags, out IntPtr err);
    [PreserveSig] int GetMsgServiceTable(uint flags, out IMAPITable table);
    [PreserveSig] int CreateMsgService([MarshalAs(UnmanagedType.LPStr)] string service, [MarshalAs(UnmanagedType.LPStr)] string display, IntPtr uiParam, uint flags);
    [PreserveSig] int DeleteMsgService(IntPtr uid);
    [PreserveSig] int CopyMsgService(IntPtr a, IntPtr b, IntPtr c, IntPtr d, IntPtr e, IntPtr f, uint g);
    [PreserveSig] int RenameMsgService(IntPtr a, uint b, IntPtr c);
    [PreserveSig] int ConfigureMsgService(IntPtr uid, IntPtr uiParam, uint flags, uint cValues, IntPtr props);
    [PreserveSig] int OpenProfileSection(IntPtr a, IntPtr b, uint c, out IntPtr d);
    [PreserveSig] int MsgServiceTransportOrder(uint a, IntPtr b, uint c);
    [PreserveSig] int AdminProviders(IntPtr a, uint b, out IntPtr c);
    [PreserveSig] int SetPrimaryIdentity(IntPtr a, uint b);
    [PreserveSig] int GetProviderTable(uint a, out IMAPITable b);
}

[ComImport, Guid("00020300-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
internal interface IMAPISession
{
    [PreserveSig] int GetLastError(int hr, uint flags, out IntPtr err);
    [PreserveSig] int GetMsgStoresTable(uint flags, out IMAPITable table);
    [PreserveSig] int OpenMsgStore(IntPtr uiParam, uint cbEntryID, IntPtr entryID, IntPtr iid, uint flags, out IMsgStore store);
    [PreserveSig] int OpenAddressBook(IntPtr a, IntPtr b, uint c, out IntPtr d);
    [PreserveSig] int OpenProfileSection(IntPtr a, IntPtr b, uint c, out IntPtr d);
    [PreserveSig] int GetStatusTable(uint a, out IntPtr b);
    [PreserveSig] int OpenEntry(uint cb, IntPtr eid, IntPtr iid, uint flags, out uint type, [MarshalAs(UnmanagedType.IUnknown)] out object obj);
    [PreserveSig] int CompareEntryIDs(uint a, IntPtr b, uint c, IntPtr d, uint e, out uint f);
    [PreserveSig] int Advise(uint a, IntPtr b, uint c, IntPtr d, out uint e);
    [PreserveSig] int Unadvise(uint a);
    [PreserveSig] int MessageOptions(IntPtr a, uint b, IntPtr c, IntPtr d);
    [PreserveSig] int QueryDefaultMessageOpt(IntPtr a, IntPtr b, uint c, out uint d, out IntPtr e);
    [PreserveSig] int EnumAdrTypes(uint a, out uint b, out IntPtr c);
    [PreserveSig] int QueryIdentity(out uint a, out IntPtr b);
    [PreserveSig] int Logoff(IntPtr uiParam, uint flags, uint reserved);
}

[ComImport, Guid("00020301-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
internal interface IMAPITable
{
    [PreserveSig] int GetLastError(int hr, uint flags, out IntPtr err);
    [PreserveSig] int Advise(uint a, IntPtr b, out uint c);
    [PreserveSig] int Unadvise(uint a);
    [PreserveSig] int GetStatus(out uint a, out uint b);
    [PreserveSig] int SetColumns(IntPtr propTagArray, uint flags);
    [PreserveSig] int QueryColumns(uint flags, out IntPtr cols);
    [PreserveSig] int GetRowCount(uint flags, out uint count);
    [PreserveSig] int SeekRow(uint bookmark, int rowCount, out int sought);
    [PreserveSig] int SeekRowApprox(uint num, uint den);
    [PreserveSig] int QueryPosition(out uint a, out uint b, out uint c);
    [PreserveSig] int FindRow(IntPtr a, uint b, uint c);
    [PreserveSig] int Restrict(IntPtr a, uint b);
    [PreserveSig] int CreateBookmark(out uint a);
    [PreserveSig] int FreeBookmark(uint a);
    [PreserveSig] int SortTable(IntPtr a, uint b);
    [PreserveSig] int QuerySortOrder(out IntPtr a);
    [PreserveSig] int QueryRows(int rowCount, uint flags, out IntPtr rows);
}

/// <summary>IMAPIProp methods, repeated at the top of every derived interface (COM interop needs the full vtable).</summary>
[ComImport, Guid("00020303-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
internal interface IMAPIProp
{
    [PreserveSig] int GetLastError(int hr, uint flags, out IntPtr err);
    [PreserveSig] int SaveChanges(uint flags);
    [PreserveSig] int GetProps(IntPtr tags, uint flags, out uint count, out IntPtr props);
    [PreserveSig] int GetPropList(uint flags, out IntPtr tags);
    [PreserveSig] int OpenProperty(uint tag, IntPtr iid, uint a, uint b, out IntPtr obj);
    [PreserveSig] int SetProps(uint count, IntPtr props, out IntPtr problems);
    [PreserveSig] int DeleteProps(IntPtr tags, out IntPtr problems);
}

[ComImport, Guid("00020306-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
internal interface IMsgStore
{
    [PreserveSig] int GetLastError(int hr, uint flags, out IntPtr err);
    [PreserveSig] int SaveChanges(uint flags);
    [PreserveSig] int GetProps(IntPtr tags, uint flags, out uint count, out IntPtr props);
    [PreserveSig] int GetPropList(uint flags, out IntPtr tags);
    [PreserveSig] int OpenProperty(uint tag, IntPtr iid, uint a, uint b, out IntPtr obj);
    [PreserveSig] int SetProps(uint count, IntPtr props, out IntPtr problems);
    [PreserveSig] int DeleteProps(IntPtr tags, out IntPtr problems);
    [PreserveSig] int CopyTo(uint a, IntPtr b, IntPtr c, IntPtr d, IntPtr e, IntPtr f, IntPtr g, uint h, out IntPtr i);
    [PreserveSig] int CopyProps(IntPtr a, IntPtr b, IntPtr c, IntPtr d, IntPtr e, IntPtr f, uint g, out IntPtr h);
    [PreserveSig] int GetNamesFromIDs(IntPtr a, IntPtr b, uint c, out uint d, out IntPtr e);
    [PreserveSig] int GetIDsFromNames(uint a, IntPtr b, uint c, out IntPtr d);
    // IMsgStore
    [PreserveSig] int Advise(uint a, IntPtr b, uint c, IntPtr d, out uint e);
    [PreserveSig] int Unadvise(uint a);
    [PreserveSig] int CompareEntryIDs(uint a, IntPtr b, uint c, IntPtr d, uint e, out uint f);
    [PreserveSig] int OpenEntry(uint cb, IntPtr eid, IntPtr iid, uint flags, out uint type, [MarshalAs(UnmanagedType.IUnknown)] out object obj);
    [PreserveSig] int SetReceiveFolder(IntPtr a, uint b, uint c, IntPtr d);
    [PreserveSig] int GetReceiveFolder(IntPtr a, uint b, out uint c, out IntPtr d, out IntPtr e);
    [PreserveSig] int GetReceiveFolderTable(uint a, out IMAPITable b);
    [PreserveSig] int StoreLogoff(ref uint flags);
}

[ComImport, Guid("0002030C-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
internal interface IMAPIFolder
{
    [PreserveSig] int GetLastError(int hr, uint flags, out IntPtr err);
    [PreserveSig] int SaveChanges(uint flags);
    [PreserveSig] int GetProps(IntPtr tags, uint flags, out uint count, out IntPtr props);
    [PreserveSig] int GetPropList(uint flags, out IntPtr tags);
    [PreserveSig] int OpenProperty(uint tag, IntPtr iid, uint a, uint b, out IntPtr obj);
    [PreserveSig] int SetProps(uint count, IntPtr props, out IntPtr problems);
    [PreserveSig] int DeleteProps(IntPtr tags, out IntPtr problems);
    [PreserveSig] int CopyTo(uint a, IntPtr b, IntPtr c, IntPtr d, IntPtr e, IntPtr f, IntPtr g, uint h, out IntPtr i);
    [PreserveSig] int CopyProps(IntPtr a, IntPtr b, IntPtr c, IntPtr d, IntPtr e, IntPtr f, uint g, out IntPtr h);
    [PreserveSig] int GetNamesFromIDs(IntPtr a, IntPtr b, uint c, out uint d, out IntPtr e);
    [PreserveSig] int GetIDsFromNames(uint a, IntPtr b, uint c, out IntPtr d);
    // IMAPIContainer
    [PreserveSig] int GetContentsTable(uint flags, out IMAPITable table);
    [PreserveSig] int GetHierarchyTable(uint flags, out IMAPITable table);
    [PreserveSig] int OpenEntry(uint cb, IntPtr eid, IntPtr iid, uint flags, out uint type, [MarshalAs(UnmanagedType.IUnknown)] out object obj);
    [PreserveSig] int SetSearchCriteria(IntPtr a, IntPtr b, uint c);
    [PreserveSig] int GetSearchCriteria(uint a, out IntPtr b, out IntPtr c, out uint d);
    // IMAPIFolder
    [PreserveSig] int CreateMessage(IntPtr iid, uint flags, out IMessage msg);
    [PreserveSig] int CopyMessages(IntPtr a, IntPtr b, IntPtr c, IntPtr d, IntPtr e, uint f);
    [PreserveSig] int DeleteMessages(IntPtr a, IntPtr b, IntPtr c, uint d);
    [PreserveSig] int CreateFolder(uint type, [MarshalAs(UnmanagedType.LPWStr)] string name, [MarshalAs(UnmanagedType.LPWStr)] string? comment, IntPtr iid, uint flags, out IMAPIFolder folder);
}

[ComImport, Guid("00020307-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
internal interface IMessage
{
    [PreserveSig] int GetLastError(int hr, uint flags, out IntPtr err);
    [PreserveSig] int SaveChanges(uint flags);
    [PreserveSig] int GetProps(IntPtr tags, uint flags, out uint count, out IntPtr props);
    [PreserveSig] int GetPropList(uint flags, out IntPtr tags);
    [PreserveSig] int OpenProperty(uint tag, IntPtr iid, uint a, uint b, out IntPtr obj);
    [PreserveSig] int SetProps(uint count, IntPtr props, out IntPtr problems);
    [PreserveSig] int DeleteProps(IntPtr tags, out IntPtr problems);
}

/// <summary>Outlook MIME converter (mimeole.h). Only MIMEToMAPI is used.</summary>
[ComImport, Guid("4B401570-B77B-11D0-9DA5-00C04FD65685"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
internal interface IConverterSession
{
    [PreserveSig] int SetAdrBook(IntPtr pab);
    [PreserveSig] int SetEncoding(int et);
    [PreserveSig] int PlaceHolder1();
    [PreserveSig] int MIMEToMAPI(IStream stm, IMessage msg, IntPtr srcSrv, uint flags);
}
