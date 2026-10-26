// The conversion steps run as child processes. If the window is killed or crashes, a child would keep running on its own
// (holding the Notes ID, the NSF or the PST). Every child is put into a job object that ends with this process.
using System.Diagnostics;
using System.Runtime.InteropServices;

namespace Nsf2Pst;

internal static class ChildJob
{
    [DllImport("kernel32.dll", CharSet = CharSet.Unicode)] static extern IntPtr CreateJobObject(IntPtr attributes, string? name);
    [DllImport("kernel32.dll")] static extern bool SetInformationJobObject(IntPtr job, int infoClass, ref ExtendedLimitInformation info, int length);
    [DllImport("kernel32.dll")] static extern bool AssignProcessToJobObject(IntPtr job, IntPtr process);

    [StructLayout(LayoutKind.Sequential)]
    struct ExtendedLimitInformation
    {
        public long PerProcessUserTimeLimit, PerJobUserTimeLimit;
        public uint LimitFlags;
        public UIntPtr MinimumWorkingSetSize, MaximumWorkingSetSize;
        public uint ActiveProcessLimit;
        public UIntPtr Affinity;
        public uint PriorityClass, SchedulingClass;
        public ulong ReadOperationCount, WriteOperationCount, OtherOperationCount, ReadTransferCount, WriteTransferCount, OtherTransferCount;
        public UIntPtr ProcessMemoryLimit, JobMemoryLimit, PeakProcessMemoryUsed, PeakJobMemoryUsed;
    }

    const int JobObjectExtendedLimitInformation = 9;
    const uint JOB_OBJECT_LIMIT_KILL_ON_JOB_CLOSE = 0x2000;

    // the handle is never closed: Windows closes it when this process ends, and that ends the children
    static readonly Lazy<IntPtr> Job = new(() =>
    {
        IntPtr job = CreateJobObject(IntPtr.Zero, null);
        if (job == IntPtr.Zero) return IntPtr.Zero;
        var info = new ExtendedLimitInformation { LimitFlags = JOB_OBJECT_LIMIT_KILL_ON_JOB_CLOSE };
        return SetInformationJobObject(job, JobObjectExtendedLimitInformation, ref info, Marshal.SizeOf<ExtendedLimitInformation>()) ? job : IntPtr.Zero;
    });

    /// <summary>Ties a started child to this process; best effort (a failure only loses the clean-up).</summary>
    public static void Add(Process p)
    {
        try { if (Job.Value != IntPtr.Zero) AssignProcessToJobObject(Job.Value, p.Handle); } catch { }
    }
}
