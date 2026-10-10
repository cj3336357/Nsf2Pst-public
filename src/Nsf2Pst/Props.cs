// Hand marshalling of MAPI property structures (x64 layout).
//   SPropValue  = { ULONG ulPropTag; ULONG dwAlignPad; union Value (16 bytes) }   -> 24 bytes, value at +8
//   SBinary     = { ULONG cb; LPBYTE lpb }                                          -> cb at +8, lpb at +16
//   SRow        = { ULONG ulAdrEntryPad; ULONG cValues; LPSPropValue lpProps }      -> 16 bytes
//   SRowSet     = { ULONG cRows; SRow aRow[] }                                      -> rows start at +8
//   SPropTagArray = { ULONG cValues; ULONG aulPropTag[] }
using System.Runtime.InteropServices;

namespace PstLite;

internal static class Props
{
    const int PropSize = 24, ValueOffset = 8;

    public static uint Type(uint tag) => tag & 0xFFFF;

    /// <summary>Builds an SPropTagArray in HGlobal memory (caller frees with Marshal.FreeHGlobal).</summary>
    public static IntPtr TagArray(params uint[] tags)
    {
        IntPtr p = Marshal.AllocHGlobal(4 + 4 * tags.Length);
        Marshal.WriteInt32(p, tags.Length);
        for (int i = 0; i < tags.Length; i++) Marshal.WriteInt32(p, 4 + 4 * i, unchecked((int)tags[i]));
        return p;
    }

    /// <summary>An SPropValue array written to HGlobal memory, with all string/binary buffers it references.</summary>
    internal sealed class Block : IDisposable
    {
        readonly List<IntPtr> allocs = new();
        public IntPtr Ptr { get; }
        public uint Count { get; }

        public Block(IList<(uint tag, object value)> values)
        {
            Count = (uint)values.Count;
            Ptr = Alloc(PropSize * values.Count);
            for (int i = 0; i < values.Count; i++)
            {
                IntPtr p = Ptr + i * PropSize;
                var (tag, value) = values[i];
                Marshal.WriteInt32(p, unchecked((int)tag));
                Marshal.WriteInt32(p, 4, 0);
                switch (Type(tag))
                {
                    case Tags.PT_LONG: Marshal.WriteInt32(p, ValueOffset, Convert.ToInt32(value)); break;
                    case Tags.PT_BOOLEAN: Marshal.WriteInt16(p, ValueOffset, (short)((bool)value ? 1 : 0)); break;
                    case Tags.PT_MV_LONG:
                        int[] ints = (int[])value; IntPtr ip = Alloc(4 * ints.Length); Marshal.Copy(ints, 0, ip, ints.Length);
                        Marshal.WriteInt32(p, ValueOffset, ints.Length); Marshal.WriteIntPtr(p, ValueOffset + 8, ip); break;
                    case Tags.PT_UNICODE:
                        IntPtr s = Marshal.StringToHGlobalUni((string)value); allocs.Add(s);
                        Marshal.WriteIntPtr(p, ValueOffset, s); break;
                    case Tags.PT_STRING8:
                        IntPtr a = Marshal.StringToHGlobalAnsi((string)value); allocs.Add(a);
                        Marshal.WriteIntPtr(p, ValueOffset, a); break;
                    case Tags.PT_SYSTIME:
                        Marshal.WriteInt64(p, ValueOffset, ((DateTimeOffset)value).UtcDateTime.ToFileTimeUtc()); break;
                    case Tags.PT_BINARY:
                        byte[] b = (byte[])value; IntPtr bp = Alloc(b.Length); Marshal.Copy(b, 0, bp, b.Length);
                        Marshal.WriteInt32(p, ValueOffset, b.Length); Marshal.WriteIntPtr(p, ValueOffset + 8, bp); break;
                    default: throw new NotSupportedException($"property type 0x{Type(tag):X4}");
                }
            }
        }

        IntPtr Alloc(int size) { IntPtr p = Marshal.AllocHGlobal(Math.Max(size, 1)); allocs.Add(p); return p; }

        public void Dispose() { foreach (var p in allocs) Marshal.FreeHGlobal(p); allocs.Clear(); }
    }

    /// <summary>Reads one SPropValue (from MAPI-owned memory) into a managed value; null for errors / unsupported types.</summary>
    public static (uint tag, object? value) Read(IntPtr prop)
    {
        uint tag = unchecked((uint)Marshal.ReadInt32(prop));
        IntPtr v = prop + ValueOffset;
        object? value = Type(tag) switch
        {
            Tags.PT_LONG => (object)unchecked((uint)Marshal.ReadInt32(v)),
            Tags.PT_UNICODE => Marshal.PtrToStringUni(Marshal.ReadIntPtr(v)),
            Tags.PT_STRING8 => Marshal.PtrToStringAnsi(Marshal.ReadIntPtr(v)),
            Tags.PT_SYSTIME => DateTimeOffset.FromFileTime(Marshal.ReadInt64(v)),
            Tags.PT_BINARY => ReadBinary(v),
            _ => null,   // PT_ERROR and anything we do not use
        };
        return (tag, value);
    }

    static byte[] ReadBinary(IntPtr v)
    {
        int cb = Marshal.ReadInt32(v);
        var bytes = new byte[cb];
        if (cb > 0) Marshal.Copy(Marshal.ReadIntPtr(v + 8), bytes, 0, cb);
        return bytes;
    }

    /// <summary>Reads all rows of an SRowSet into dictionaries keyed by property tag, then frees it.</summary>
    public static List<Dictionary<uint, object?>> ReadRows(IntPtr rowSet)
    {
        var result = new List<Dictionary<uint, object?>>();
        if (rowSet == IntPtr.Zero) return result;
        try
        {
            int cRows = Marshal.ReadInt32(rowSet);
            for (int r = 0; r < cRows; r++)
            {
                IntPtr row = rowSet + 8 + r * 16;
                int cValues = Marshal.ReadInt32(row, 4);
                IntPtr props = Marshal.ReadIntPtr(row, 8);
                var d = new Dictionary<uint, object?>();
                for (int i = 0; i < cValues; i++)
                {
                    var (tag, value) = Read(props + i * PropSize);
                    d[tag] = value;
                }
                result.Add(d);
            }
        }
        finally { Native.FreeProws(rowSet); }
        return result;
    }

    /// <summary>Reads the given properties of an object (GetProps) into a dictionary keyed by the requested tag.</summary>
    public static Dictionary<uint, object?> Get(Func<IntPtr, (int hr, uint count, IntPtr props)> getProps, params uint[] tags)
    {
        IntPtr tagArray = TagArray(tags);
        try
        {
            var (hr, count, props) = getProps(tagArray);
            var d = new Dictionary<uint, object?>();
            if (hr < 0 || props == IntPtr.Zero) return d;
            try
            {
                for (int i = 0; i < count; i++)
                {
                    var (tag, value) = Read(props + i * PropSize);
                    if (Type(tag) == Type(tags[i])) d[tags[i]] = value;   // PT_ERROR entries are skipped
                }
            }
            finally { Native.MAPIFreeBuffer(props); }
            return d;
        }
        finally { Marshal.FreeHGlobal(tagArray); }
    }

    /// <summary>Queries all rows of a MAPI table with the given columns.</summary>
    public static List<Dictionary<uint, object?>> QueryAll(IMAPITable table, params uint[] columns)
    {
        IntPtr cols = TagArray(columns);
        try
        {
            Check(table.SetColumns(cols, 0), "SetColumns");
            var all = new List<Dictionary<uint, object?>>();
            while (true)
            {
                Check(table.QueryRows(200, 0, out IntPtr rows), "QueryRows");
                var batch = ReadRows(rows);
                if (batch.Count == 0) break;
                all.AddRange(batch);
            }
            return all;
        }
        finally { Marshal.FreeHGlobal(cols); }
    }

    public static void Check(int hr, string what)
    {
        if (hr < 0) throw new MapiException(what, hr);
    }
}

internal sealed class MapiException(string what, int hr) : Exception($"{what} failed: 0x{hr:X8}")
{
    public int HResult2 { get; } = hr;
}
