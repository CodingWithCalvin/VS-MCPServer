using System;
using System.Diagnostics;
using CodingWithCalvin.MCPServer.Services;
using Xunit;

namespace CodingWithCalvin.MCPServer.Tests;

/// <summary>
/// Covers the kill-on-close backstop that stops the MCP server process from outliving Visual
/// Studio when devenv.exe terminates without running its normal shutdown path.
/// </summary>
public class ProcessJobObjectTests
{
    private const int ExitWaitMs = 5000;

    [Fact]
    public void Create_ReturnsJobObject()
    {
        using var job = ProcessJobObject.Create();

        Assert.NotNull(job);
    }

    [Fact]
    public void Dispose_TerminatesAssignedProcess()
    {
        var process = StartLongRunningProcess();

        try
        {
            var job = ProcessJobObject.Create();
            Assert.NotNull(job);

            Assert.True(job!.TryAssign(process), "Failed to assign the process to the job object.");
            Assert.False(process.HasExited, "The child process exited before the job was closed.");

            job.Dispose();

            Assert.True(
                process.WaitForExit(ExitWaitMs),
                "Closing the job object did not terminate the assigned process.");
        }
        finally
        {
            KillIfRunning(process);
            process.Dispose();
        }
    }

    [Fact]
    public void TryAssign_ReturnsFalse_AfterDispose()
    {
        var job = ProcessJobObject.Create();
        Assert.NotNull(job);
        job!.Dispose();

        var process = StartLongRunningProcess();

        try
        {
            Assert.False(job.TryAssign(process));
        }
        finally
        {
            KillIfRunning(process);
            process.Dispose();
        }
    }

    [Fact]
    public void Dispose_IsIdempotent()
    {
        var job = ProcessJobObject.Create();
        Assert.NotNull(job);

        job!.Dispose();
        job.Dispose();
    }

    [Fact]
    public void TryAssign_Throws_ForNullProcess()
    {
        using var job = ProcessJobObject.Create();

        Assert.Throws<ArgumentNullException>(() => job!.TryAssign(null!));
    }

    /// <summary>
    /// Starts a process that runs for long enough that any exit observed during a test is
    /// attributable to the job object rather than to natural termination.
    /// </summary>
    private static Process StartLongRunningProcess()
    {
        var startInfo = new ProcessStartInfo("ping.exe", "-n 120 127.0.0.1")
        {
            UseShellExecute = false,
            CreateNoWindow = true,
            RedirectStandardOutput = true,
        };

        var process = Process.Start(startInfo);
        Assert.NotNull(process);

        return process!;
    }

    private static void KillIfRunning(Process process)
    {
        try
        {
            if (!process.HasExited)
            {
                process.Kill();
                process.WaitForExit(ExitWaitMs);
            }
        }
        catch (InvalidOperationException)
        {
            // Already gone.
        }
    }
}
