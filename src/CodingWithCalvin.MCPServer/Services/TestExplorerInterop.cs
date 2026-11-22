using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel.Composition.Primitives;
using System.Linq;
using System.Reflection;
using System.Threading;
using System.Threading.Tasks;
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
/// The TestWindow types ship only inside Visual Studio. The statistics, operation state, and
/// test services are resolved by MEF contract name and accessed reflectively so the VSIX can
/// build without a Visual Studio installation path.
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
    private const string TestsServiceContract =
        "Microsoft.VisualStudio.TestWindow.Extensibility.ITestsService";
    private const string VsTestServiceContract =
        "Microsoft.VisualStudio.TestWindow.Extensibility.IVsTestService";
    private const string TestQueryTypeName =
        "Microsoft.VisualStudio.TestWindow.Extensibility.TestQuery";
    private const string TestPropertyTypeName =
        "Microsoft.VisualStudio.TestWindow.Extensibility.TestPropertyType";
    private const string FilterMatchKindTypeName =
        "Microsoft.VisualStudio.TestWindow.Extensibility.FilterMatchKind";

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

    internal async Task<TestExplorerTestListResult> GetTestsAsync(
        string? filter,
        string? state,
        int maxResults)
    {
        try
        {
            var tests = await ReadTestsAsync().ConfigureAwait(false);
            var filtered = tests
                .Select(ReadTestInfo)
                .Where(test => MatchesTestFilter(test, filter, state))
                .ToList();
            var limit = NormalizeResultLimit(maxResults);

            return new TestExplorerTestListResult
            {
                Available = true,
                TotalCount = filtered.Count,
                Truncated = filtered.Count > limit,
                Tests = filtered.Take(limit).ToList()
            };
        }
        catch (Exception ex)
        {
            VsixTelemetry.TrackException(ex);
            return new TestExplorerTestListResult
            {
                Message = GetErrorMessage(ex)
            };
        }
    }

    internal async Task<TestExplorerTestResultsResult> GetResultsAsync(
        IReadOnlyCollection<string>? testIds,
        string? filter,
        bool failedOnly,
        int maxResults)
    {
        try
        {
            var ids = ParseTestIds(testIds);
            if (testIds != null && ids == null)
            {
                return new TestExplorerTestResultsResult
                {
                    Message = "One or more test IDs are not valid GUIDs."
                };
            }

            var tests = await ReadTestsAsync().ConfigureAwait(false);
            var results = tests
                .Select(test => ReadTestResults(test, ids, filter, failedOnly))
                .SelectMany(testResults => testResults)
                .OrderByDescending(result => result.ReceivedOrder)
                .ToList();
            var limit = NormalizeResultLimit(maxResults);

            return new TestExplorerTestResultsResult
            {
                Available = true,
                TotalCount = results.Count,
                Truncated = results.Count > limit,
                Results = results.Take(limit).Select(result => result.Info).ToList()
            };
        }
        catch (Exception ex)
        {
            VsixTelemetry.TrackException(ex);
            return new TestExplorerTestResultsResult
            {
                Message = GetErrorMessage(ex)
            };
        }
    }

    internal async Task<TestExplorerSelectionResult> RunTestsAsync(
        IReadOnlyCollection<string> testIds,
        bool debug)
    {
        if (testIds == null || testIds.Count == 0)
        {
            return new TestExplorerSelectionResult { Message = "At least one test ID is required." };
        }

        var ids = ParseTestIds(testIds);
        if (ids == null || ids.Count == 0)
        {
            return new TestExplorerSelectionResult { Message = "One or more test IDs are not valid GUIDs." };
        }

        try
        {
            var service = ResolveVsTestService(out var contract);
            if (service == null || contract == null)
            {
                return new TestExplorerSelectionResult
                {
                    SelectedCount = ids.Count,
                    Message = "The Test Explorer test service is unavailable. Open Test Explorer and wait for test discovery."
                };
            }

            var query = CreateRecordIdQuery(contract, ids);
            var methodName = debug ? "DebugTestsAsync" : "RunTestsAsync";
            var method = contract.GetMethods().FirstOrDefault(candidate =>
            {
                var parameters = candidate.GetParameters();
                return candidate.Name == methodName
                    && parameters.Length == 2
                    && parameters[0].ParameterType == query.GetType()
                    && parameters[1].ParameterType == typeof(CancellationToken);
            });
            if (method == null)
            {
                return new TestExplorerSelectionResult
                {
                    SelectedCount = ids.Count,
                    Message = $"The installed Test Explorer does not expose {methodName}."
                };
            }

            var started = await InvokeTaskResultAsync(method, service, query, CancellationToken.None).ConfigureAwait(false) is true;
            return new TestExplorerSelectionResult
            {
                Started = started,
                SelectedCount = ids.Count,
                Message = started ? null : $"{methodName} returned false; Test Explorer did not start the selected tests."
            };
        }
        catch (Exception ex)
        {
            VsixTelemetry.TrackException(ex);
            return new TestExplorerSelectionResult { SelectedCount = ids.Count, Message = GetErrorMessage(ex) };
        }
    }

    internal static List<TestExplorerTestInfo> FindTestsForTarget(
        IEnumerable<TestExplorerTestInfo> tests,
        string target)
    {
        var candidates = tests.ToList();
        if (Guid.TryParse(target, out var id))
        {
            return candidates.Where(test => Guid.TryParse(test.Id, out var testId) && testId == id).ToList();
        }

        var exact = candidates.Where(test =>
            string.Equals(test.FullyQualifiedName, target, StringComparison.OrdinalIgnoreCase)
            || string.Equals(test.DisplayName, target, StringComparison.OrdinalIgnoreCase)).ToList();
        if (exact.Count > 0)
        {
            return exact;
        }

        var classMatches = candidates.Where(test =>
            test.FullyQualifiedName.StartsWith(target + ".", StringComparison.OrdinalIgnoreCase)
            || test.FullyQualifiedName.IndexOf("." + target + ".", StringComparison.OrdinalIgnoreCase) >= 0)
            .ToList();
        if (classMatches.Count > 0)
        {
            return classMatches;
        }

        return candidates.Where(test =>
            test.FullyQualifiedName.EndsWith("." + target, StringComparison.OrdinalIgnoreCase)
            || string.Equals(test.DisplayName, target, StringComparison.OrdinalIgnoreCase)).ToList();
    }

    private async Task<List<object>> ReadTestsAsync()
    {
        var service = ResolveTestsService(out var contract);
        if (service == null || contract == null)
        {
            throw new InvalidOperationException(
                "The Test Explorer test service is unavailable. Open Test Explorer and wait for test discovery.");
        }

        if (contract.GetProperty("IsTestServiceRunning")?.GetValue(service) is not true)
        {
            throw new InvalidOperationException(
                "Test Explorer has not completed initialization. Open Test Explorer and wait for test discovery.");
        }

        var method = contract.GetMethods().FirstOrDefault(candidate =>
            candidate.Name == "GetTestsAsEnumerableAsync" && candidate.GetParameters().Length == 0);
        if (method == null)
        {
            throw new InvalidOperationException("The installed Test Explorer does not expose test enumeration.");
        }

        var value = await InvokeTaskResultAsync(method, service).ConfigureAwait(false);
        return value is IEnumerable tests ? tests.Cast<object>().ToList() : new List<object>();
    }

    private object? ResolveVsTestService(out Type? contract)
    {
        var service = ResolveExport(VsTestServiceContract);
        if (service == null)
        {
            // VS 18's legacy wrapper imports the public service into this private field.
            var legacyService = ResolveExport(TestsServiceContract);
            service = legacyService?.GetType()
                .GetField("vsTestService", BindingFlags.Instance | BindingFlags.NonPublic)
                ?.GetValue(legacyService);
        }

        contract = service == null ? null : FindInterface(service, VsTestServiceContract);
        return contract == null ? null : service;
    }

    private static object CreateRecordIdQuery(Type contract, IEnumerable<Guid> ids)
    {
        var assembly = contract.Assembly;
        var queryType = assembly.GetType(TestQueryTypeName, throwOnError: true);
        var propertyType = assembly.GetType(TestPropertyTypeName, throwOnError: true);
        var matchKindType = assembly.GetType(FilterMatchKindTypeName, throwOnError: true);
        var property = Enum.Parse(propertyType, "TestCaseRecordId");
        var matchKind = Enum.Parse(matchKindType, "ExactMatch");
        var values = ids.Select(id => id.ToString()).ToArray();

        return Activator.CreateInstance(
            queryType,
            BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic,
            binder: null,
            args: new[] { property, values, matchKind },
            culture: null)
            ?? throw new InvalidOperationException("Could not create the Test Explorer query.");
    }

    private object? ResolveTestsService(out Type? contract)
    {
        var service = ResolveExport(TestsServiceContract);
        contract = service == null ? null : FindInterface(service, TestsServiceContract);
        return contract == null ? null : service;
    }

    private static async Task<object?> InvokeTaskResultAsync(
        MethodInfo method,
        object instance,
        params object?[] arguments)
    {
        if (method.Invoke(instance, arguments) is not Task task)
        {
            throw new InvalidOperationException($"{method.Name} did not return a Task.");
        }

        await task.ConfigureAwait(false);
        return task.GetType().GetProperty("Result")?.GetValue(task);
    }

    private static TestExplorerTestInfo ReadTestInfo(object test)
    {
        return new TestExplorerTestInfo
        {
            Id = ReadProperty(test, "Id")?.ToString() ?? string.Empty,
            FullyQualifiedName = ReadProperty(test, "FullyQualifiedName")?.ToString() ?? string.Empty,
            DisplayName = ReadProperty(test, "DisplayName")?.ToString() ?? string.Empty,
            State = ReadProperty(test, "State")?.ToString() ?? string.Empty,
            Source = ReadProperty(test, "Source")?.ToString() ?? string.Empty,
            FilePath = ReadProperty(test, "FilePath")?.ToString(),
            LineNumber = ReadInt(test, "LineNumber"),
            IsRunning = ReadProperty(test, "IsCurrentlyRunning") is true,
            IsStale = ReadProperty(test, "Stale") is true
        };
    }

    private static IEnumerable<(long ReceivedOrder, TestExplorerTestResultInfo Info)> ReadTestResults(
        object test,
        HashSet<Guid>? ids,
        string? filter,
        bool failedOnly)
    {
        var testInfo = ReadTestInfo(test);
        if (ids != null
            && (!Guid.TryParse(testInfo.Id, out var id) || !ids.Contains(id)))
        {
            return Enumerable.Empty<(long, TestExplorerTestResultInfo)>();
        }

        if (!MatchesTestFilter(testInfo, filter, null)
            || ReadProperty(test, "Results") is not IEnumerable results)
        {
            return Enumerable.Empty<(long, TestExplorerTestResultInfo)>();
        }

        return results.Cast<object>()
            .Select(result =>
            {
                var state = ReadProperty(result, "State")?.ToString() ?? string.Empty;
                return new
                {
                    State = state,
                    ReceivedOrder = ReadLong(result, "ReceivedOrder"),
                    Info = new TestExplorerTestResultInfo
                    {
                        TestId = testInfo.Id,
                        FullyQualifiedName = testInfo.FullyQualifiedName,
                        DisplayName = ReadProperty(result, "DisplayName")?.ToString() ?? testInfo.DisplayName,
                        State = state,
                        RunId = ReadInt(result, "RunId"),
                        DurationMilliseconds = ReadDurationMilliseconds(result),
                        ErrorMessage = ReadProperty(result, "ErrorMessage")?.ToString(),
                        ErrorStackTrace = ReadProperty(result, "ErrorStackTrace")?.ToString(),
                        StandardOutput = ReadProperty(result, "StandardOutput")?.ToString(),
                        StandardError = ReadProperty(result, "StandardError")?.ToString(),
                        Attachments = ReadProperty(result, "AttachmentsString")?.ToString()
                    }
                };
            })
            .Where(result => !failedOnly || string.Equals(result.State, "Failed", StringComparison.OrdinalIgnoreCase))
            .Select(result => (result.ReceivedOrder, result.Info));
    }

    private static bool MatchesTestFilter(TestExplorerTestInfo test, string? filter, string? state)
    {
        var matchesFilter = string.IsNullOrWhiteSpace(filter)
            || test.FullyQualifiedName.IndexOf(filter, StringComparison.OrdinalIgnoreCase) >= 0
            || test.DisplayName.IndexOf(filter, StringComparison.OrdinalIgnoreCase) >= 0
            || test.Source.IndexOf(filter, StringComparison.OrdinalIgnoreCase) >= 0;
        var matchesState = string.IsNullOrWhiteSpace(state)
            || string.Equals(test.State, state, StringComparison.OrdinalIgnoreCase);
        return matchesFilter && matchesState;
    }

    private static HashSet<Guid>? ParseTestIds(IReadOnlyCollection<string>? testIds)
    {
        if (testIds == null)
        {
            return null;
        }

        var ids = new HashSet<Guid>();
        foreach (var testId in testIds)
        {
            if (!Guid.TryParse(testId, out var id))
            {
                return null;
            }

            ids.Add(id);
        }

        return ids;
    }

    private static int NormalizeResultLimit(int maxResults) => Math.Min(Math.Max(maxResults, 1), 1000);

    private static object? ReadProperty(object instance, string name)
    {
        var property = instance.GetType().GetInterfaces()
            .Select(type => type.GetProperty(name))
            .FirstOrDefault(candidate => candidate != null)
            ?? instance.GetType().GetProperty(name, BindingFlags.Instance | BindingFlags.Public);
        return property?.GetValue(instance);
    }

    private static int ReadInt(object instance, string name) =>
        Convert.ToInt32(ReadProperty(instance, name) ?? 0);

    private static long ReadLong(object instance, string name) =>
        Convert.ToInt64(ReadProperty(instance, name) ?? 0L);

    private static long ReadDurationMilliseconds(object instance) =>
        ReadProperty(instance, "Duration") is TimeSpan duration ? (long)duration.TotalMilliseconds : 0;

    private static string GetErrorMessage(Exception exception) =>
        exception is TargetInvocationException { InnerException: not null } invocation
            ? invocation.InnerException.Message
            : exception.Message;

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
    internal async Task<TestRunStatus> GetRunStatusAsync()
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
        try
        {
            if ((await ReadTestsAsync().ConfigureAwait(false)).Any(test =>
                ReadProperty(test, "IsCurrentlyRunning") is true))
            {
                state = "Running";
            }
        }
        catch (Exception ex)
        {
            VsixTelemetry.TrackException(ex);
        }

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
