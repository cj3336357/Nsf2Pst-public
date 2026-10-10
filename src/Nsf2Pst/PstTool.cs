// "import" / "list" / "counts" mode: writes EML files into a PST through Extended MAPI (64-bit Outlook).
//   Nsf2Pst import <pst> <manifest.jsonl>   one JSON object per line: {"eml","folder","subject","read","received","sent"}
//   Nsf2Pst list <pst>                       folders and messages, for checking
//   Nsf2Pst counts <pst>                     "folder TAB items TAB unread", numbers only
// Progress lines for the UI: "PROGRESS <done> <total>".
using System.Diagnostics;
using System.Text;
using System.Text.Json;
using PstLite;

namespace Nsf2Pst;

internal static class PstTool
{
    public static int Run(string[] args)
    {
        if (args.Length < 2) { Console.Error.WriteLine("usage: import <pst> <manifest.jsonl> | list <pst> | counts <pst>"); return 2; }
        string cmd = args[0].ToLowerInvariant(), pst = args[1];
        try
        {
            using var store = new PstStore(pst);
            if (cmd == "list") { store.List(Console.Out); return 0; }
            if (cmd == "counts") { store.Counts(Console.Out); return 0; }
            if (args.Length < 3) { Console.Error.WriteLine("missing manifest"); return 2; }

            string[] lines = File.ReadAllLines(args[2], Encoding.UTF8).Where(l => !string.IsNullOrWhiteSpace(l)).ToArray();
            int ok = 0, failed = 0;
            var sw = Stopwatch.StartNew();
            for (int i = 0; i < lines.Length; i++)
            {
                if (i % 10 == 0) { Console.WriteLine($"PROGRESS {i} {lines.Length}"); Console.Out.Flush(); }
                string eml = "?";
                try
                {
                    using var doc = JsonDocument.Parse(lines[i]);
                    var r = doc.RootElement;
                    string type = r.TryGetProperty("type", out var ty) ? ty.GetString() ?? "mail" : "mail";
                    if (type == "contact")
                    {
                        eml = "contact";
                        var fields = r.GetProperty("fields").EnumerateObject().ToDictionary(x => x.Name, x => x.Value.GetString() ?? "");
                        store.ImportContact(r.GetProperty("folder").GetString() ?? "@Contacts", fields);
                        ok++;
                        continue;
                    }
                    if (type == "appointment")
                    {
                        eml = "appointment";
                        store.ImportAppointment(r.GetProperty("folder").GetString() ?? "@Calendar",
                            r.GetProperty("subject").GetString() ?? "",
                            r.GetProperty("start").GetDateTimeOffset(), r.GetProperty("end").GetDateTimeOffset(),
                            r.TryGetProperty("allDay", out var ad) && ad.GetBoolean(),
                            r.TryGetProperty("location", out var lo) ? lo.GetString() : null,
                            r.TryGetProperty("body", out var bo) ? bo.GetString() : null,
                            r.TryGetProperty("reminder", out var rm) && rm.ValueKind == JsonValueKind.Number ? rm.GetInt32() : null,
                            r.TryGetProperty("tentative", out var te) && te.GetBoolean());
                        ok++;
                        continue;
                    }
                    eml = r.GetProperty("eml").GetString()!;
                    string folder = r.TryGetProperty("folder", out var f) ? f.GetString() ?? "" : "";
                    string? subject = r.TryGetProperty("subject", out var s) ? s.GetString() : null;
                    bool? read = r.TryGetProperty("read", out var rd) ? rd.GetBoolean() : null;
                    DateTimeOffset? received = r.TryGetProperty("received", out var rc) ? rc.GetDateTimeOffset() : null;
                    DateTimeOffset? sent = r.TryGetProperty("sent", out var st) ? st.GetDateTimeOffset() : null;
                    store.Import(eml, folder, subject, read, received, sent);
                    ok++;
                }
                catch (Exception ex)
                {
                    failed++;
                    Console.WriteLine($"FAILED line {i + 1} ({Path.GetFileName(eml)}): {ex.Message}");
                }
            }
            Console.WriteLine($"PROGRESS {lines.Length} {lines.Length}");
            Console.WriteLine($"imported {ok}, failed {failed}, {sw.Elapsed.TotalSeconds:F1}s");
            return failed == 0 ? 0 : 1;
        }
        catch (Exception ex)
        {
            Console.WriteLine("fatal: " + ex.Message);
            return 3;
        }
    }
}
