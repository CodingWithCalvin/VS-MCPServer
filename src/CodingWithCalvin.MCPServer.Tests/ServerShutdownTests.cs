using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using CodingWithCalvin.MCPServer.Services;
using CodingWithCalvin.MCPServer.Shared.Models;
using Xunit;

namespace CodingWithCalvin.MCPServer.Tests;

/// <summary>
/// Regression tests for issue #97: package disposal blocked the Visual Studio UI thread waiting
/// on <c>StopAsync</c>, whose continuations were posted straight back to that same blocked
/// thread. devenv.exe never finished exiting and stayed resident in Task Manager.
/// </summary>
public class ServerShutdownTests
{
    private static readonly TimeSpan CompletionTimeout = TimeSpan.FromSeconds(5);
    private static readonly TimeSpan ThreadJoinTimeout = TimeSpan.FromSeconds(30);

    [Fact]
    public void ServerProcessManager_StopAsync_CompletesWhenCallerBlocksItsSynchronizationContext()
    {
        var completed = RunWithBlockedSynchronizationContext(() =>
        {
            var manager = new ServerProcessManager(new StubRpcServer());
            return manager.StopAsync();
        });

        Assert.True(
            completed,
            "ServerProcessManager.StopAsync did not complete. A continuation is being posted back to "
                + "the caller's synchronization context, which is the deadlock from issue #97. "
                + "Every await in the shutdown path needs ConfigureAwait(false).");
    }

    /// <summary>
    /// Mirrors the shape of package disposal itself: a synchronous, blocking wait on the shutdown
    /// path from a thread that cannot pump its own message queue.
    /// </summary>
    [Fact]
    public void BlockingOnStopAsync_Returns_WhenCallerBlocksItsSynchronizationContext()
    {
        var completed = RunWithBlockedSynchronizationContext(() =>
        {
            var manager = new ServerProcessManager(new StubRpcServer());
            manager.StopAsync().GetAwaiter().GetResult();
            return Task.CompletedTask;
        });

        Assert.True(completed, "Blocking on the shutdown path deadlocked the calling thread.");
    }

    /// <summary>
    /// Runs <paramref name="operation"/> on a thread whose synchronization context silently drops
    /// everything posted to it — the observable behaviour of a UI thread blocked inside
    /// <c>Dispose</c>. Any continuation that tries to resume on the caller's context will never
    /// run, so the returned task never completes and this reports <see langword="false"/>.
    /// </summary>
    private static bool RunWithBlockedSynchronizationContext(Func<Task> operation)
    {
        var completed = false;

        var thread = new Thread(() =>
        {
            SynchronizationContext.SetSynchronizationContext(new BlockedSynchronizationContext());
            completed = operation().Wait(CompletionTimeout);
        })
        {
            IsBackground = true,
        };

        thread.Start();
        thread.Join(ThreadJoinTimeout);

        return completed;
    }

    private sealed class BlockedSynchronizationContext : SynchronizationContext
    {
        public override void Post(SendOrPostCallback d, object? state)
        {
            // Deliberately dropped: the thread that would pump this is blocked.
        }

        public override void Send(SendOrPostCallback d, object? state)
            => throw new NotSupportedException("The blocked thread cannot run work synchronously.");
    }

    /// <summary>
    /// Minimal <see cref="IRpcServer"/> whose async members complete on the thread pool, which is
    /// what makes the captured-context bug observable.
    /// </summary>
    private sealed class StubRpcServer : IRpcServer
    {
        public string PipeName => string.Empty;

        public bool IsListening => true;

        public bool IsConnected => false;

        public Task StartAsync(string pipeName) => Task.CompletedTask;

        public async Task StopAsync() => await Task.Delay(25).ConfigureAwait(false);

        public Task<List<ToolInfo>> GetAvailableToolsAsync() => Task.FromResult(new List<ToolInfo>());

        public async Task RequestShutdownAsync() => await Task.Delay(25).ConfigureAwait(false);

        public void Dispose()
        {
        }
    }
}
