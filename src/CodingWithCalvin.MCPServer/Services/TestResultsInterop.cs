using System;
using System.Collections;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.ComponentModel.Composition.Primitives;
using System.Linq;
using System.Reflection;
using System.Threading.Tasks;
using CodingWithCalvin.MCPServer.Shared.Models;
using CodingWithCalvin.Otel4Vsix;
using Microsoft.VisualStudio.ComponentModelHost;

namespace CodingWithCalvin.MCPServer.Services;

/// <summary>
/// Reflection bridge to the individual test results Test Explorer holds: which tests failed,
/// and their messages, stack traces, files and lines.
/// </summary>
/// <remarks>
/// <para>
/// The public Test Explorer surface (<see cref="TestExplorerInterop"/>) reports counts only.
/// The per-test data lives behind <c>ITestsService</c>, <c>ITest</c> and <c>IResult</c> in
/// <c>Microsoft.VisualStudio.TestWindow.Interfaces.dll</c>. Those interfaces are
/// <b>internal</b>, so they cannot be referenced even if the assembly could be (see
/// <see cref="TestExplorerInterop"/> for why it cannot). <c>ITestsService</c> is still a MEF
/// export under its own contract name (implemented by <c>TestsServiceWrapper</c>), so it is
/// resolved by that name and read reflectively, always through the interface types because the
/// implementations may implement members explicitly.
/// </para>
/// <para>
/// The shapes were checked on Visual Studio 2026 18.7
/// (<c>Microsoft.VisualStudio.TestWindow.Interfaces.dll</c> 18.700.26.51602). Being internal,
/// they carry no compatibility promise and a Visual Studio update can rename or remove them.
/// Every failure therefore degrades to <see cref="TestResultsReport.Available"/> false with a
/// reason; nothing here throws.
/// </para>
/// </remarks>
internal sealed class TestResultsInterop
{
    internal const string TestsServiceContract =
        "Microsoft.VisualStudio.TestWindow.Extensibility.ITestsService";
    internal const string TestContract = "Microsoft.VisualStudio.TestWindow.Extensibility.ITest";
    internal const string ResultContract = "Microsoft.VisualStudio.TestWindow.Extensibility.IResult";

    /// <summary>
    /// Longest error message, stack trace or output string returned per result. An MSTest stack
    /// can run to many kilobytes and the MCP client has a token limit.
    /// </summary>
    internal const int MaxTextLength = 4000;

    internal const string ServiceNotFoundMessage =
        "Test Explorer's test service was not found in this Visual Studio. test_results reads an "
        + "internal Test Explorer API (checked on Visual Studio 18.7) that this version may not have. "
        + "test_stats still reports counts. This does not mean there are no failed tests.";

    internal const string NotRunningMessage =
        "Test Explorer has not been initialized in this Visual Studio session, so no results are "
        + "available. Open Test Explorer (toolwindow_show with 'TestExplorer') or start a run first. "
        + "This does not mean there are no failed tests.";

    private static readonly Dictionary<string, string[]?> Outcomes =
        new Dictionary<string, string[]?>(StringComparer.OrdinalIgnoreCase)
        {
            ["Failed"] = new[] { "Failed" },
            ["NotRun"] = new[] { "NotRun", "None" },
            ["Skipped"] = new[] { "Skipped" },
            ["Passed"] = new[] { "Passed" },
            ["All"] = null
        };

    private readonly Func<IComponentModel?> _componentModelAccessor;
    private readonly ConcurrentDictionary<Type, TestReader?> _testReaders = new();
    private readonly ConcurrentDictionary<Type, ResultReader?> _resultReaders = new();

    internal TestResultsInterop(Func<IComponentModel?> componentModelAccessor)
    {
        _componentModelAccessor = componentModelAccessor;
    }

    /// <summary>
    /// The outcome names <see cref="TestResultsQuery.Outcome"/> accepts.
    /// </summary>
    internal static IReadOnlyCollection<string> OutcomeNames => Outcomes.Keys;

    internal static bool IsValidOutcome(string? outcome) =>
        outcome != null && Outcomes.ContainsKey(outcome);

    /// <summary>
    /// Reads the tests that match <paramref name="query"/>. <see cref="TestResultsReport.State"/>
    /// is left for the caller, which owns the run-state tracking.
    /// </summary>
    internal async Task<TestResultsReport> GetResultsAsync(TestResultsQuery query)
    {
        if (!Outcomes.TryGetValue(query.Outcome ?? string.Empty, out var states))
        {
            return Unavailable(
                $"Invalid outcome '{query.Outcome}'. Valid values are: {string.Join(", ", Outcomes.Keys)}.");
        }

        try
        {
            var service = ResolveExport(TestsServiceContract);
            var contract = service == null ? null : FindInterface(service.GetType(), TestsServiceContract);
            if (service == null || contract == null)
            {
                return Unavailable(ServiceNotFoundMessage);
            }

            if (contract.GetProperty("IsTestServiceRunning")?.GetValue(service) is not true)
            {
                return Unavailable(NotRunningMessage);
            }

            var getTests = contract.GetMethod("GetTestsAsEnumerableAsync", Type.EmptyTypes);
            if (getTests == null || getTests.Invoke(service, null) is not Task task)
            {
                return Unavailable(ServiceNotFoundMessage);
            }

            await task;

            if (task.GetType().GetProperty("Result")?.GetValue(task) is not IEnumerable tests)
            {
                return Unavailable(ServiceNotFoundMessage);
            }

            return Collect(tests, states, query);
        }
        catch (Exception ex)
        {
            var inner = ex is TargetInvocationException { InnerException: { } cause } ? cause : ex;
            VsixTelemetry.TrackException(inner);
            return Unavailable(
                $"Reading Test Explorer results failed ({inner.GetType().Name}: {inner.Message}). "
                + "The internal Test Explorer API may have changed in this Visual Studio version. "
                + "This does not mean there are no failed tests.");
        }
    }

    private TestResultsReport Collect(IEnumerable tests, string[]? states, TestResultsQuery query)
    {
        var matched = new List<(object Test, TestReader Reader, string Name, string DisplayName)>();

        foreach (var test in tests)
        {
            if (test == null)
            {
                continue;
            }

            var reader = GetTestReader(test.GetType());
            if (reader == null)
            {
                return Unavailable(ServiceNotFoundMessage);
            }

            if (states != null && !states.Contains(reader.ReadState(test), StringComparer.Ordinal))
            {
                continue;
            }

            var name = reader.ReadString(reader.FullyQualifiedName, test) ?? string.Empty;
            if (!string.IsNullOrEmpty(query.Filter)
                && name.IndexOf(query.Filter, StringComparison.OrdinalIgnoreCase) < 0)
            {
                continue;
            }

            matched.Add((test, reader, name, reader.ReadString(reader.DisplayName, test) ?? name));
        }

        var limit = Math.Max(0, query.MaxResults);
        var report = new TestResultsReport
        {
            Available = true,
            TotalMatched = matched.Count,
            Truncated = matched.Count > limit
        };

        // Data rows can be separate tests sharing one name, so the display name, which carries
        // the row's arguments, keeps their order stable.
        foreach (var (test, reader, name, displayName) in matched
            .OrderBy(m => m.Name, StringComparer.OrdinalIgnoreCase)
            .ThenBy(m => m.DisplayName, StringComparer.OrdinalIgnoreCase)
            .Take(limit))
        {
            report.Tests.Add(ToEntry(test, reader, name, displayName, query.IncludeOutput));
        }

        return report;
    }

    private TestResultEntry ToEntry(
        object test,
        TestReader reader,
        string name,
        string displayName,
        bool includeOutput)
    {
        var entry = new TestResultEntry
        {
            FullyQualifiedName = name,
            DisplayName = displayName,
            Outcome = reader.ReadState(test),
            DurationMs = ReadMilliseconds(reader.Duration, test),
            Stale = reader.Stale?.GetValue(test) is true,
            FilePath = NullIfEmpty(reader.ReadString(reader.FilePath, test)),
            LineNumber = reader.LineNumber?.GetValue(test) is int line ? line : 0,
            Source = NullIfEmpty(reader.ReadString(reader.Source, test))
        };

        if (reader.Results?.GetValue(test) is IEnumerable results)
        {
            foreach (var result in results)
            {
                if (result != null && GetResultReader(result.GetType()) is { } resultReader)
                {
                    entry.Results.Add(ToDetail(result, resultReader, includeOutput));
                }
            }
        }

        return entry;
    }

    private static TestResultDetail ToDetail(object result, ResultReader reader, bool includeOutput)
    {
        var truncated = false;
        var detail = new TestResultDetail
        {
            DisplayName = reader.Read(reader.DisplayName, result) ?? string.Empty,
            Outcome = reader.State?.GetValue(result)?.ToString() ?? string.Empty,
            DurationMs = ReadMilliseconds(reader.Duration, result),
            ErrorMessage = Cap(reader.Read(reader.ErrorMessage, result), ref truncated),
            ErrorStackTrace = Cap(reader.Read(reader.ErrorStackTrace, result), ref truncated)
        };

        if (includeOutput)
        {
            detail.StandardOutput = Cap(reader.Read(reader.StandardOutput, result), ref truncated);
            detail.StandardError = Cap(reader.Read(reader.StandardError, result), ref truncated);
        }

        detail.Truncated = truncated;
        return detail;
    }

    internal static string? Cap(string? text, ref bool truncated)
    {
        if (string.IsNullOrEmpty(text))
        {
            return null;
        }

        if (text!.Length <= MaxTextLength)
        {
            return text;
        }

        truncated = true;
        return text.Substring(0, MaxTextLength);
    }

    private static long ReadMilliseconds(PropertyInfo? property, object instance) =>
        property?.GetValue(instance) is TimeSpan duration ? (long)duration.TotalMilliseconds : 0;

    private static string? NullIfEmpty(string? value) => string.IsNullOrEmpty(value) ? null : value;

    private static TestResultsReport Unavailable(string message) =>
        new TestResultsReport { Available = false, Message = message };

    private TestReader? GetTestReader(Type type) =>
        _testReaders.GetOrAdd(type, t => FindInterface(t, TestContract) is { } contract
            ? TestReader.Create(contract)
            : null);

    private ResultReader? GetResultReader(Type type) =>
        _resultReaders.GetOrAdd(type, t => FindInterface(t, ResultContract) is { } contract
            ? new ResultReader(contract)
            : null);

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

    private static Type? FindInterface(Type type, string fullName) =>
        string.Equals(type.FullName, fullName, StringComparison.Ordinal) && type.IsInterface
            ? type
            : type.GetInterfaces()
                .FirstOrDefault(i => string.Equals(i.FullName, fullName, StringComparison.Ordinal));

    /// <summary>
    /// Property accessors for <c>ITest</c>. <see cref="FullyQualifiedName"/> and
    /// <see cref="State"/> are required; the rest are read when present.
    /// </summary>
    private sealed class TestReader
    {
        private TestReader(Type contract, PropertyInfo fullyQualifiedName, PropertyInfo state)
        {
            FullyQualifiedName = fullyQualifiedName;
            State = state;
            DisplayName = contract.GetProperty("DisplayName");
            FilePath = contract.GetProperty("FilePath");
            Source = contract.GetProperty("Source");
            Duration = contract.GetProperty("Duration");
            LineNumber = contract.GetProperty("LineNumber");
            Stale = contract.GetProperty("Stale");
            Results = contract.GetProperty("Results");
        }

        internal PropertyInfo FullyQualifiedName { get; }
        internal PropertyInfo State { get; }
        internal PropertyInfo? DisplayName { get; }
        internal PropertyInfo? FilePath { get; }
        internal PropertyInfo? Source { get; }
        internal PropertyInfo? Duration { get; }
        internal PropertyInfo? LineNumber { get; }
        internal PropertyInfo? Stale { get; }
        internal PropertyInfo? Results { get; }

        internal static TestReader? Create(Type contract)
        {
            var name = contract.GetProperty("FullyQualifiedName");
            var state = contract.GetProperty("State");
            return name == null || state == null ? null : new TestReader(contract, name, state);
        }

        internal string ReadState(object test) => State.GetValue(test)?.ToString() ?? string.Empty;

        internal string? ReadString(PropertyInfo? property, object test) =>
            property?.GetValue(test) as string;
    }

    /// <summary>
    /// Property accessors for <c>IResult</c>, each read when present.
    /// </summary>
    private sealed class ResultReader
    {
        internal ResultReader(Type contract)
        {
            DisplayName = contract.GetProperty("DisplayName");
            State = contract.GetProperty("State");
            Duration = contract.GetProperty("Duration");
            ErrorMessage = contract.GetProperty("ErrorMessage");
            ErrorStackTrace = contract.GetProperty("ErrorStackTrace");
            StandardOutput = contract.GetProperty("StandardOutput");
            StandardError = contract.GetProperty("StandardError");
        }

        internal PropertyInfo? DisplayName { get; }
        internal PropertyInfo? State { get; }
        internal PropertyInfo? Duration { get; }
        internal PropertyInfo? ErrorMessage { get; }
        internal PropertyInfo? ErrorStackTrace { get; }
        internal PropertyInfo? StandardOutput { get; }
        internal PropertyInfo? StandardError { get; }

        internal string? Read(PropertyInfo? property, object result) =>
            property?.GetValue(result) as string;
    }
}
