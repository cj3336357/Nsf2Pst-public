// One DllImport resolver for the whole assembly (it can only be set once):
//   nnotes.dll   -> HCL Notes program folder
//   OLMAPI32.DLL -> Click-to-Run / MSI Office16 folder
using System.Runtime.InteropServices;

namespace Nsf2Pst;

internal static class Resolver
{
    public static void Register() => NativeLibrary.SetDllImportResolver(typeof(Resolver).Assembly, Resolve);

    static IntPtr Resolve(string name, System.Reflection.Assembly asm, DllImportSearchPath? path)
    {
        IEnumerable<string> dirs;
        if (name.Equals("nnotes.dll", StringComparison.OrdinalIgnoreCase)) dirs = NsfExport.NotesApi.ProgramDirs();
        else if (name.Equals("OLMAPI32.DLL", StringComparison.OrdinalIgnoreCase))
            dirs = new[] { @"%ProgramFiles%\Microsoft Office\root\Office16", @"%ProgramFiles%\Microsoft Office\Office16" }
                .Select(Environment.ExpandEnvironmentVariables);
        else return IntPtr.Zero;

        foreach (string dir in dirs)
        {
            string p = Path.Combine(dir, name);
            if (File.Exists(p) && NativeLibrary.TryLoad(p, out IntPtr h)) return h;
        }
        return IntPtr.Zero;
    }
}
