using System;
using System.ComponentModel.Composition.Primitives;
using System.Linq;
using System.Reflection;
using CodingWithCalvin.MCPServer.Shared.Models;
using CodingWithCalvin.Otel4Vsix;
using Microsoft.VisualStudio.ComponentModelHost;

namespace CodingWithCalvin.MCPServer.Services;

/// <summary>
/// Reflection bridge to the Test Explorer extensibility services.
/// </summary>
/// <remarks>
/// <para>
/// The types used here live in <c>Microsoft.VisualStudio.TestWindow.Interfaces.dll</c>, which
/// ships only inside the Visual Studio installation. It cannot be referenced at compile time:
/// nuget.org carries nothing newer than 11.0.61030 (2012), it is absent from the
/// Microsoft.VisualStudio.SDK metapackage, and <c>$(DevEnvDir)</c> is undefined under
/// <c>dotnet build</c> — which is how this repo builds both locally and in CI — so a HintPath
/// would break the build outright.
/// </para>
/// <para>
/// The surface needed is small and stable: one bool, one struct of four ints, and one event.
/// It is byte-identical between VS 2022 17.14 and VS 2026 18.0, the range the VSIX manifest
/// targets. Both services are MEF parts (<c>TestExplorerStatsService</c> exports
/// <c>ITestExplorerStatsService</c>, <c>OperationBroker</c> exports <c>IOperationState</c>), so
/// they are resolved by contract name and read reflectively.
/// </para>
/// <para>
/// Every failure path degrades to "unavailable" rather than throwing. Test Explorer may never
/// have been opened, and on a future Visual Studio the shape could change; neither should take
/// an MCP tool call down with it.
/// </para>
/// </remarks>
internal sealed class TestExplorerInterop
{
    private const string StatsServiceContract =
        "Microsoft.VisualStudio.TestWindow.Extensibility.ITestExplorerStatsService";
    private const string OperationStateContract =
        "Microsoft.VisualStudio.TestWindow.Extensibility.IOperationState";

    private const string DiscoveryPrefix = "Discovery";
    private const string ExecutionPrefix = "TestExecution";

    private readonly object _gate = new object();
    private readonly Func<IComponentModel?> _componentModelAccessor;

    private bool _statsResolved;
    private object? _statsService;
    private PropertyInfo? _isRunningProperty;
    private PropertyInfo? _statsProperty;
    private PropertyInfo? _passedProperty;
    private PropertyInfo? _failedProperty;
    private PropertyInfo? _skippedProperty;
    private PropertyInfo? _notRunProperty;

    private bool _trackingAttempted;
    private string? _lastOperation;
    private string? _lastExecutionOperation;

    internal TestExplorerInterop(Func<IComponentModel?> componentModelAccessor)
    {
        _componentModelAccessor = componentModelAccessor;
    }

    /// <summary>
    /// Reads the current Test Explorer counts. Returns <see cref="TestStats.Available"/> false
    /// when Test Explorer has not been initialized, so that an uninitialized window is not
    /// reported as a solution with zero tests.
    /// </summary>
    internal TestStats GetStats()
    {
        try
        {
            lock (_gate)
            {
                if (!TryResolveStatsService())
                {
                    return new TestStats();
                }

                if (_isRunningProperty!.GetValue(_statsService) is not true)
                {
                    return new TestStats();
                }

                var stats = _statsProperty!.GetValue(_statsService);
                if (stats == null)
                {
                    return new TestStats();
                }

                return new TestStats
                {
                    Available = true,
                    Passed = ReadCount(_passedProperty, stats),
                    Failed = ReadCount(_failedProperty, stats),
                    Skipped = ReadCount(_skippedProperty, stats),
                    NotRun = ReadCount(_notRunProperty, stats)
                };
            }
        }
        catch (Exception ex)
        {
            VsixTelemetry.TrackException(ex);
            return new TestStats();
        }
    }

    /// <summary>
    /// Reports the tracked run state alongside the current counts. State is only observed from
    /// the point <see cref="EnsureTracking"/> first succeeds, so a run started outside this
    /// extension before then reads as "NoRunObserved".
    /// </summary>
    internal TestRunStatus GetRunStatus()
    {
        EnsureTracking();

        string? lastOperation;
        string? lastExecution;

        lock (_gate)
        {
            lastOperation = _lastOperation;
            lastExecution = _lastExecutionOperation;
        }

        var state = DeriveState(lastOperation, lastExecution);

        return new TestRunStatus
        {
            State = state,
            IsRunning = state is "Discovering" or "Running" or "Canceling",
            LastOperation = lastOperation,
            Stats = GetStats()
        };
    }

    /// <summary>
    /// Subscribes to Test Explorer operation state changes. Safe to call repeatedly; the
    /// subscription is attempted once and the outcome cached either way.
    /// </summary>
    internal void EnsureTracking()
    {
        lock (_gate)
        {
            if (_trackingAttempted)
            {
                return;
            }

            _trackingAttempted = true;

            try
            {
                var operationState = ResolveExport(OperationStateContract);
                if (operationState == null)
                {
                    return;
                }

                // IOperationState.StateChanged is implemented explicitly on OperationBroker, so
                // the event has to be reached through the interface rather than the class.
                var contract = FindInterface(operationState, OperationStateContract);
                var stateChanged = contract?.GetEvent("StateChanged");
                if (stateChanged?.EventHandlerType == null)
                {
                    return;
                }

                var callback = typeof(TestExplorerInterop).GetMethod(
                    nameof(OnOperationStateChanged),
                    BindingFlags.Instance | BindingFlags.NonPublic);
                if (callback == null)
                {
                    return;
                }

                // The event is EventHandler<OperationStateChangedEventArgs>; the callback takes
                // the EventArgs base. Delegate creation allows that contravariance.
                var handler = Delegate.CreateDelegate(
                    stateChanged.EventHandlerType,
                    this,
                    callback,
                    throwOnBindFailure: false);
                if (handler == null)
                {
                    return;
                }

                stateChanged.AddEventHandler(operationState, handler);
            }
            catch (Exception ex)
            {
                VsixTelemetry.TrackException(ex);
            }
        }
    }

    private void OnOperationStateChanged(object sender, EventArgs args)
    {
        try
        {
            var operation = args.GetType().GetProperty("State")?.GetValue(args)?.ToString();
            if (string.IsNullOrEmpty(operation))
            {
                return;
            }

            lock (_gate)
            {
                _lastOperation = operation;

                if (operation!.StartsWith(ExecutionPrefix, StringComparison.Ordinal))
                {
                    _lastExecutionOperation = operation;
                }
            }
        }
        catch (Exception ex)
        {
            VsixTelemetry.TrackException(ex);
        }
    }

    /// <summary>
    /// Collapses the raw Test Explorer operation names into the small state set the tools
    /// report. Discovery is only surfaced while it is in flight, so that a discovery pass
    /// triggered after a run does not erase the outcome of that run.
    /// </summary>
    internal static string DeriveState(string? lastOperation, string? lastExecutionOperation)
    {
        if (lastOperation != null
            && lastOperation.StartsWith(DiscoveryPrefix, StringComparison.Ordinal)
            && IsInFlight(lastOperation))
        {
            return "Discovering";
        }

        return lastExecutionOperation switch
        {
            null => "NoRunObserved",
            "TestExecutionFinished" => "Completed",
            "TestExecutionCancelAndFinished" => "Canceled",
            "TestExecutionCanceling" => "Canceling",
            _ => "Running"
        };
    }

    private static bool IsInFlight(string operation) =>
        !operation.EndsWith("Finished", StringComparison.Ordinal)
        && !operation.EndsWith("Canceled", StringComparison.Ordinal);

    private static int ReadCount(PropertyInfo? property, object stats) =>
        property?.GetValue(stats) is int value ? value : 0;

    private bool TryResolveStatsService()
    {
        if (_statsResolved)
        {
            return _statsService != null;
        }

        _statsResolved = true;

        var service = ResolveExport(StatsServiceContract);
        if (service == null)
        {
            return false;
        }

        var contract = FindInterface(service, StatsServiceContract);
        if (contract == null)
        {
            return false;
        }

        _isRunningProperty = contract.GetProperty("IsTestExplorerStatsServiceRunning");
        _statsProperty = contract.GetProperty("TestExplorerStats");
        if (_isRunningProperty == null || _statsProperty == null)
        {
            return false;
        }

        var statsType = _statsProperty.PropertyType;
        _passedProperty = statsType.GetProperty("PassedTestCount");
        _failedProperty = statsType.GetProperty("FailedTestCount");
        _skippedProperty = statsType.GetProperty("SkippedTestCount");
        _notRunProperty = statsType.GetProperty("NotRunTestCount");

        _statsService = service;
        return true;
    }

    private object? ResolveExport(string contractName)
    {
        var componentModel = _componentModelAccessor();
        if (componentModel == null)
        {
            return null;
        }

        var definition = new ImportDefinition(
            d => d.ContractName == contractName,
            contractName,
            ImportCardinality.ZeroOrMore,
            isRecomposable: false,
            isPrerequisite: false);

        return componentModel.DefaultExportProvider
            .GetExports(definition)
            .FirstOrDefault()
            ?.Value;
    }

    private static Type? FindInterface(object instance, string fullName) =>
        instance.GetType()
            .GetInterfaces()
            .FirstOrDefault(i => string.Equals(i.FullName, fullName, StringComparison.Ordinal));
}
