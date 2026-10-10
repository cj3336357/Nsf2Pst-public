// Nsf2Pst - Notes mail file (NSF) -> Outlook PST.
//   Nsf2Pst.exe                 opens the window
//   Nsf2Pst.exe export ...      Notes side   (run by the window as a child process)
//   Nsf2Pst.exe import|list|counts ...  PST side (run by the window as a separate child process)
// The Notes and Outlook parts always run in different processes so their DLLs never meet.
using System.Runtime.InteropServices;
using System.Text;

namespace Nsf2Pst;

internal static class Program
{
    [DllImport("ole32.dll")] static extern int CoInitializeEx(IntPtr reserved, uint coInit);
    [DllImport("kernel32.dll")] static extern bool AttachConsole(int pid);

    [STAThread]
    static int Main(string[] args)
    {
        if (args.Length == 0)
        {
            ApplicationConfiguration.Initialize();
            Application.Run(new MainForm());
            return 0;
        }

        AttachConsole(-1);   // when started from a command prompt, print there (no-op when output is redirected)
        // a child of the window has no console (output is a pipe read as UTF-8); setting the console code page then throws
        try { Console.OutputEncoding = Encoding.UTF8; } catch (IOException) { }
        if (Console.IsOutputRedirected) Console.SetOut(new StreamWriter(Console.OpenStandardOutput(), new UTF8Encoding(false)) { AutoFlush = true });
        CoInitializeEx(IntPtr.Zero, 0x2);   // STA: Notes COM, and the Outlook MIME converter needs COM on this thread
        Resolver.Register();
        try
        {
            return args[0].ToLowerInvariant() switch
            {
                "export" or "mark-unread" => ExportTool.Run(args),
                "import" or "list" or "counts" => PstTool.Run(args),
                _ => Usage(),
            };
        }
        catch (Exception ex)
        {
            Console.WriteLine("fatal: " + ex.Message);
            return 3;
        }
    }

    static int Usage()
    {
        Console.WriteLine("Nsf2Pst.exe (no arguments) opens the window; command line: export | import | list | counts");
        return 2;
    }
}
