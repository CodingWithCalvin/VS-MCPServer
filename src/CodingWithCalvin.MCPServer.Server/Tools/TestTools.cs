using System;
using System.ComponentModel;
using System.Linq;
using System.Text.Json;
using System.Threading.Tasks;
using CodingWithCalvin.MCPServer.Shared.Models;
using ModelContextProtocol.Server;

namespace CodingWithCalvin.MCPServer.Server.Tools;

[McpServerToolType]
public class TestTools
{
    private const string StatsUnavailableMessage =
        "Test Explorer has not been initialized in this Visual Studio session, so no counts are "
        + "available. Open Test Explorer (window_show with 'TestExplorer') or start a run first. "
        + "Note this is not the same as the solution having no tests.";

    private static readonly string[] Outcomes = { "failed", "notrun", "skipped", "passed", "all" };

    private readonly RpcClient _rpcClient;
    private readonly JsonSerializerOptions _jsonOptions;

    public TestTools(RpcClient rpcClient)
    {
        _rpcClient = rpcClient;
        _jsonOptions = new JsonSerializerOptions { WriteIndented = true };
    }

    [McpServerTool(Name = "test_run_all", Destructive = false)]
    [Description("Run every test in the solution via Test Explorer. The run starts asynchronously and this returns immediately; poll test_status to observe completion. Requires a solution with discovered tests.")]
    public async Task<string> RunAllTestsAsync()
    {
        var started = await _rpcClient.RunAllTestsAsync();
        return started
            ? "Test run started. Poll test_status for progress."
            : "Failed to start the test run. Test Explorer may still be discovering tests, or the solution has none.";
    }

    [McpServerTool(Name = "test_debug_all", Destructive = false)]
    [Description("Debug every test in the solution via Test Explorer, stopping on any breakpoints that are set. The run starts asynchronously and this returns immediately; poll test_status to observe completion.")]
    public async Task<string> DebugAllTestsAsync()
    {
        var started = await _rpcClient.DebugAllTestsAsync();
        return started
            ? "Test debug run started. Poll test_status for progress."
            : "Failed to start the test debug run. Test Explorer may still be discovering tests, or the solution has none.";
    }

    [McpServerTool(Name = "test_run", Destructive = false)]
    [Description("Run the tests in a single test class or test method. Accepts a simple name ('CalculatorTests', 'Add_ReturnsSum') or a fully qualified one ('MyApp.Tests.CalculatorTests.Add_ReturnsSum'); a fully qualified name is matched first and avoids ambiguity. The run starts asynchronously; poll test_status to observe completion.")]
    public async Task<string> RunTestAsync(
        [Description("Test class or test method name to run. Fully qualified names are preferred when several types share a simple name.")] string target)
    {
        var result = await _rpcClient.RunTestsInContextAsync(target, debug: false);
        return JsonSerializer.Serialize(result, _jsonOptions);
    }

    [McpServerTool(Name = "test_debug", Destructive = false)]
    [Description("Debug the tests in a single test class or test method, stopping on any breakpoints that are set. Accepts a simple or fully qualified name; a fully qualified name is matched first and avoids ambiguity. The run starts asynchronously; poll test_status to observe completion.")]
    public async Task<string> DebugTestAsync(
        [Description("Test class or test method name to debug. Fully qualified names are preferred when several types share a simple name.")] string target)
    {
        var result = await _rpcClient.RunTestsInContextAsync(target, debug: true);
        return JsonSerializer.Serialize(result, _jsonOptions);
    }

    [McpServerTool(Name = "test_cancel", Destructive = false, Idempotent = true)]
    [Description("Cancel the test run that is currently in progress.")]
    public async Task<string> CancelTestRunAsync()
    {
        var cancelled = await _rpcClient.CancelTestRunAsync();
        return cancelled
            ? "Test run cancelled"
            : "No test run is currently in progress";
    }

    [McpServerTool(Name = "test_status", ReadOnly = true)]
    [Description("Get the Test Explorer run state plus current counts. State is one of 'NoRunObserved', 'Discovering', 'Running', 'Canceling', 'Completed', or 'Canceled'. Use this to poll for completion after test_run_all, test_run, or their debug equivalents. State is only tracked for runs started after this extension first contacted Test Explorer.")]
    public async Task<string> GetTestStatusAsync()
    {
        var status = await _rpcClient.GetTestRunStatusAsync();
        return JsonSerializer.Serialize(status, _jsonOptions);
    }

    [McpServerTool(Name = "test_stats", ReadOnly = true)]
    [Description("Get the passed, failed, skipped and not-run test counts from Test Explorer. Returns counts only; use test_results for the individual tests, with failure messages, stack traces, files and lines.")]
    public async Task<string> GetTestStatsAsync()
    {
        var stats = await _rpcClient.GetTestStatsAsync();
        if (!stats.Available)
        {
            return StatsUnavailableMessage;
        }

        return JsonSerializer.Serialize(stats, _jsonOptions);
    }

    [McpServerTool(Name = "test_results", ReadOnly = true)]
    [Description("Get the individual tests from Test Explorer with their outcome, file and line, and per result the error message and stack trace. Defaults to failed tests only. Data rows of a data-driven test are usually separate entries that share FullyQualifiedName, with the row's arguments in DisplayName; Results can also hold one result per row. LineNumber is where the test method starts; the failing line is in ErrorStackTrace. State is the run state as in test_status, so results read while a run is still going can be told apart. Stale is true when a result is out of date: the code has changed since it ran, or it was restored from an earlier Visual Studio session. Error messages, stack traces and output are cut to 4,000 characters each, flagged by Truncated on the result. TotalMatched counts every test that passed the filters; the top-level Truncated is true when that is more than maxResults. When Available is false, Message says why, and that does not mean there are no failed tests.")]
    public async Task<string> GetTestResultsAsync(
        [Description("Which tests to return: failed, notrun, skipped, passed, or all. Defaults to failed. Case-insensitive.")]
        string outcome = "failed",
        [Description("Substring of the fully qualified test name, such as a class or namespace. Case-insensitive.")]
        string? filter = null,
        [Description("Maximum number of tests to return. Defaults to 50.")]
        int maxResults = 50,
        [Description("Include each result's standard output and standard error. Defaults to false.")]
        bool includeOutput = false)
    {
        if (!Outcomes.Contains(outcome, StringComparer.OrdinalIgnoreCase))
        {
            return $"Invalid outcome '{outcome}'. Valid values are: {string.Join(", ", Outcomes)}.";
        }

        var report = await _rpcClient.GetTestResultsAsync(new TestResultsQuery
        {
            Outcome = outcome,
            Filter = string.IsNullOrWhiteSpace(filter) ? null : filter,
            MaxResults = maxResults,
            IncludeOutput = includeOutput
        });

        return JsonSerializer.Serialize(report, _jsonOptions);
    }
}
