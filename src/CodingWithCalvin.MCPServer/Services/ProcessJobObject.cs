using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using Microsoft.Win32.SafeHandles;

namespace CodingWithCalvin.MCPServer.Services;

/// <summary>
/// Wraps a Windows job object configured with <c>JOB_OBJECT_LIMIT_KILL_ON_JOB_CLOSE</c>.
/// Any process assigned to the job is terminated by the operating system as soon as the
/// job handle closes, which happens when this instance is disposed or when the owning
/// process (devenv.exe) exits for any reason.
/// </summary>
/// <remarks>
/// This is the backstop for orphaned server processes. Cooperative shutdown handles the
/// normal case, but if Visual Studio crashes or is killed from Task Manager it never gets
/// the chance to run, and the server process would otherwise survive and keep holding its
/// HTTP port against the next Visual Studio session.
/// </remarks>
internal sealed class ProcessJobObject : IDisposable
{
    private SafeJobHandle? _handle;

    private ProcessJobObject(SafeJobHandle handle)
    {
        _handle = handle;
    }

    /// <summary>
    /// Creates a kill-on-close job object, or returns <see langword="null"/> if the
    /// operating system refuses. Callers treat a null result as "backstop unavailable"
    /// and continue without it.
    /// </summary>
    public static ProcessJobObject? Create()
    {
        var handle = NativeMethods.CreateJobObject(IntPtr.Zero, null);

        if (handle.IsInvalid)
        {
            handle.Dispose();
            return null;
        }

        var extendedLimits = new NativeMethods.JOBOBJECT_EXTENDED_LIMIT_INFORMATION();
        extendedLimits.BasicLimitInformation.LimitFlags = NativeMethods.JOB_OBJECT_LIMIT_KILL_ON_JOB_CLOSE;

        var length = Marshal.SizeOf(typeof(NativeMethods.JOBOBJECT_EXTENDED_LIMIT_INFORMATION));
        var buffer = Marshal.AllocHGlobal(length);

        try
        {
            Marshal.StructureToPtr(extendedLimits, buffer, false);

            var configured = NativeMethods.SetInformationJobObject(
                handle,
                NativeMethods.JobObjectExtendedLimitInformation,
                buffer,
                (uint)length);

            if (!configured)
            {
                handle.Dispose();
                return null;
            }
        }
        finally
        {
            Marshal.FreeHGlobal(buffer);
        }

        return new ProcessJobObject(handle);
    }

    /// <summary>
    /// Assigns <paramref name="process"/> to the job. Returns <see langword="false"/> if the
    /// assignment fails, in which case the caller simply loses the kill-on-close guarantee.
    /// </summary>
    public bool TryAssign(Process process)
    {
        if (process == null)
        {
            throw new ArgumentNullException(nameof(process));
        }

        var handle = _handle;

        if (handle == null || handle.IsInvalid || handle.IsClosed)
        {
            return false;
        }

        try
        {
            return NativeMethods.AssignProcessToJobObject(handle, process.Handle);
        }
        catch (InvalidOperationException)
        {
            // The process exited before we could assign it.
            return false;
        }
    }

    public void Dispose()
    {
        var handle = _handle;
        _handle = null;
        handle?.Dispose();
    }

    private sealed class SafeJobHandle : SafeHandleZeroOrMinusOneIsInvalid
    {
        public SafeJobHandle() : base(true)
        {
        }

        protected override bool ReleaseHandle() => NativeMethods.CloseHandle(handle);
    }

    private static class NativeMethods
    {
        public const int JobObjectExtendedLimitInformation = 9;
        public const uint JOB_OBJECT_LIMIT_KILL_ON_JOB_CLOSE = 0x2000;

        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        public static extern SafeJobHandle CreateJobObject(IntPtr lpJobAttributes, string? lpName);

        [DllImport("kernel32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool SetInformationJobObject(
            SafeJobHandle hJob,
            int jobObjectInfoClass,
            IntPtr lpJobObjectInfo,
            uint cbJobObjectInfoLength);

        [DllImport("kernel32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool AssignProcessToJobObject(SafeJobHandle hJob, IntPtr hProcess);

        [DllImport("kernel32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool CloseHandle(IntPtr hObject);

        // Fields on these interop structures are populated by the marshaller rather than by
        // managed code, so the compiler cannot see them being assigned.
#pragma warning disable CS0649
        [StructLayout(LayoutKind.Sequential)]
        public struct JOBOBJECT_BASIC_LIMIT_INFORMATION
        {
            public long PerProcessUserTimeLimit;
            public long PerJobUserTimeLimit;
            public uint LimitFlags;
            public UIntPtr MinimumWorkingSetSize;
            public UIntPtr MaximumWorkingSetSize;
            public uint ActiveProcessLimit;
            public UIntPtr Affinity;
            public uint PriorityClass;
            public uint SchedulingClass;
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct IO_COUNTERS
        {
            public ulong ReadOperationCount;
            public ulong WriteOperationCount;
            public ulong OtherOperationCount;
            public ulong ReadTransferCount;
            public ulong WriteTransferCount;
            public ulong OtherTransferCount;
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct JOBOBJECT_EXTENDED_LIMIT_INFORMATION
        {
            public JOBOBJECT_BASIC_LIMIT_INFORMATION BasicLimitInformation;
            public IO_COUNTERS IoInfo;
            public UIntPtr ProcessMemoryLimit;
            public UIntPtr JobMemoryLimit;
            public UIntPtr PeakProcessMemoryUsed;
            public UIntPtr PeakJobMemoryUsed;
        }
#pragma warning restore CS0649
    }
}
