using System;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using CodingWithCalvin.Otel4Vsix;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Debugger.Interop;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;

namespace CodingWithCalvin.MCPServer.Services;

internal sealed class DebuggerPerformanceCounterService : IDebugEventCallback2, IVsDebuggerEvents, IDisposable
{
    private static readonly TimeSpan QueryTimeout = TimeSpan.FromSeconds(5);

    private readonly IServiceProvider _serviceProvider;
    private readonly object _sync = new();
    private IVsDebugger? _debugger;
    private IDebugProgram150? _lastBreakProgram;
    private uint _modeEventsCookie;
    private bool _debugEventsRegistered;
    private bool _modeEventsRegistered;
    private bool _disposed;
    private int _version;

    public DebuggerPerformanceCounterService(IServiceProvider serviceProvider)
    {
        _serviceProvider = serviceProvider;
    }

    public void Initialize()
    {
        ThreadHelper.ThrowIfNotOnUIThread();
        if (_debugger != null)
        {
            return;
        }

        _debugger = _serviceProvider.GetService(typeof(SVsShellDebugger)) as IVsDebugger
            ?? throw new InvalidOperationException("Could not get Visual Studio debugger service");

        try
        {
            ErrorHandler.ThrowOnFailure(_debugger.AdviseDebugEventCallback(this));
            _debugEventsRegistered = true;
            ErrorHandler.ThrowOnFailure(_debugger.AdviseDebuggerEvents(this, out _modeEventsCookie));
            _modeEventsRegistered = true;
        }
        catch
        {
            Dispose();
            throw;
        }
    }

    public async Task<long?> GetElapsedMillisecondsAsync()
    {
        await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();

        try
        {
            CompletionRoutine callback;
            int version;

            lock (_sync)
            {
                if (_disposed || _lastBreakProgram == null)
                {
                    return null;
                }

                callback = new CompletionRoutine();
                version = _version;
                if (_lastBreakProgram.QueryPerformanceCounters(callback) != VSConstants.S_OK)
                {
                    return null;
                }
            }

            var completed = await Task.WhenAny(callback.Results, Task.Delay(QueryTimeout)).ConfigureAwait(false);
            if (completed != callback.Results)
            {
                return null;
            }

            var results = await callback.Results.ConfigureAwait(false);
            lock (_sync)
            {
                return !_disposed && version == _version
                    ? ExtractElapsedMilliseconds(results)
                    : null;
            }
        }
        catch (Exception ex)
        {
            VsixTelemetry.TrackException(ex);
            return null;
        }
    }

    internal static long? ExtractElapsedMilliseconds(QUERY_PROCESS_PERFORMANCE_COUNTERS_RESULT[]? results)
    {
        if (results == null || results.Length == 0 || results[0].ErrorCode != VSConstants.S_OK)
        {
            return null;
        }

        // PerfTipService converts SystemTime to UInt32 before updating its control.
        return unchecked((uint)results[0].Counters.SystemTime);
    }

    public int Event(
        IDebugEngine2 pEngine,
        IDebugProcess2 pProcess,
        IDebugProgram2 pProgram,
        IDebugThread2 pThread,
        IDebugEvent2 pEvent,
        ref Guid riidEvent,
        uint dwAttrib)
    {
        var retainProgram = false;
        try
        {
            if (riidEvent == typeof(IDebugBreakpointEvent2).GUID ||
                riidEvent == typeof(IDebugStepCompleteEvent2).GUID)
            {
                if (pProgram is IDebugProgram150 program)
                {
                    retainProgram = TryRetainProgram(program);
                }
                else
                {
                    ClearProgram();
                }
            }

            return VSConstants.S_OK;
        }
        finally
        {
            ReleaseComObject(pEngine);
            ReleaseComObject(pProcess);
            ReleaseComObject(pThread);
            ReleaseComObject(pEvent);
            if (!retainProgram)
            {
                ReleaseComObject(pProgram);
            }
        }
    }

    public int OnModeChange(DBGMODE newmode)
    {
        if (newmode != DBGMODE.DBGMODE_Break)
        {
            ClearProgram();
        }

        return VSConstants.S_OK;
    }

    private bool TryRetainProgram(IDebugProgram150 program)
    {
        IDebugProgram150? previous;
        lock (_sync)
        {
            if (_disposed)
            {
                return false;
            }

            if (ReferenceEquals(_lastBreakProgram, program))
            {
                return true;
            }

            previous = _lastBreakProgram;
            _lastBreakProgram = program;
            _version++;
        }

        ReleaseComObject(previous);
        return true;
    }

    private void ClearProgram()
    {
        IDebugProgram150? previous;
        lock (_sync)
        {
            previous = _lastBreakProgram;
            _lastBreakProgram = null;
            if (previous == null)
            {
                return;
            }

            _version++;
        }

        ReleaseComObject(previous);
    }

    private static void ReleaseComObject(object? value)
    {
        if (value != null && Marshal.IsComObject(value))
        {
            Marshal.ReleaseComObject(value);
        }
    }

    public void Dispose()
    {
        ThreadHelper.ThrowIfNotOnUIThread();
        if (_disposed)
        {
            return;
        }

        _disposed = true;
        ClearProgram();

        if (_modeEventsRegistered)
        {
            try
            {
                _debugger?.UnadviseDebuggerEvents(_modeEventsCookie);
            }
            catch (Exception ex)
            {
                VsixTelemetry.TrackException(ex);
            }
        }

        if (_debugEventsRegistered)
        {
            try
            {
                _debugger?.UnadviseDebugEventCallback(this);
            }
            catch (Exception ex)
            {
                VsixTelemetry.TrackException(ex);
            }
        }

        _modeEventsRegistered = false;
        _debugEventsRegistered = false;
        _modeEventsCookie = 0;
        _debugger = null;
    }

    private sealed class CompletionRoutine : IDebugQueryPerformanceCountersCompletionRoutine150
    {
        private readonly TaskCompletionSource<QUERY_PROCESS_PERFORMANCE_COUNTERS_RESULT[]> _completion =
            new(TaskCreationOptions.RunContinuationsAsynchronously);

        public Task<QUERY_PROCESS_PERFORMANCE_COUNTERS_RESULT[]> Results => _completion.Task;

        public int QueryPerformanceCountersOnComplete(QUERY_PROCESS_PERFORMANCE_COUNTERS_RESULT[] pResults)
        {
            _completion.TrySetResult(pResults ?? Array.Empty<QUERY_PROCESS_PERFORMANCE_COUNTERS_RESULT>());
            return VSConstants.S_OK;
        }
    }
}
