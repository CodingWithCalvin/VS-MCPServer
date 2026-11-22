using System.ComponentModel;
using System.Collections.Generic;
using System.Text.Json;
using System.Threading.Tasks;
using ModelContextProtocol.Server;

namespace CodingWithCalvin.MCPServer.Server.Tools;

[McpServerToolType]
public class TestTools
{
    private const string StatsUnavailableMessage =
        "Test Explorer has not been initialized in this Visual Studio session, so no counts are "
        + "available. Open Test Explorer (window_show with 'TestExplorer') or start a run first. "
        + "Note this is not the same as the solution having no tests.";

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
    [Description("Run tests already discovered by Test Explorer matching a class or test name. This does not move the editor caret. Use test_list and test_run_selected when a name is ambiguous. Poll test_status to observe completion.")]
    public async Task<string> RunTestAsync(
        [Description("Test class or test method name to run. Fully qualified names are preferred when several types share a simple name.")] string target)
    {
        var result = await _rpcClient.RunTestsInContextAsync(target, debug: false);
        return JsonSerializer.Serialize(result, _jsonOptions);
    }

    [McpServerTool(Name = "test_debug", Destructive = false)]
    [Description("Debug tests already discovered by Test Explorer matching a class or test name. This does not move the editor caret. Use test_list and test_debug_selected when a name is ambiguous. Poll test_status to observe completion.")]
    public async Task<string> DebugTestAsync(
        [Description("Test class or test method name to debug. Fully qualified names are preferred when several types share a simple name.")] string target)
    {
        var result = await _rpcClient.RunTestsInContextAsync(target, debug: true);
        return JsonSerializer.Serialize(result, _jsonOptions);
    }

    [McpServerTool(Name = "test_list", ReadOnly = true)]
    [Description("List tests currently discovered by Test Explorer. Filter by fully qualified name, display name, or source; optionally filter by state. Each item includes its Test Explorer ID for test_run_selected or test_debug_selected. Results are capped by maxResults.")]
    public async Task<string> GetTestExplorerTestsAsync(
        [Description("Optional case-insensitive substring filter over test name and source.")] string? filter = null,
        [Description("Optional exact state filter, such as Passed, Failed, Skipped, NotRun, or None.")] string? state = null,
        [Description("Maximum number of tests to return (1-1000). Defaults to 200.")] int maxResults = 200)
    {
        var result = await _rpcClient.GetTestExplorerTestsAsync(filter, state, maxResults);
        return JsonSerializer.Serialize(result, _jsonOptions);
    }

    [McpServerTool(Name = "test_run_selected", Destructive = false)]
    [Description("Run the specified Test Explorer tests by IDs returned from test_list. The run starts asynchronously; poll test_status for progress.")]
    public async Task<string> RunSelectedTestsAsync(
        [Description("One or more Test Explorer test IDs returned by test_list.")] List<string> testIds)
    {
        var result = await _rpcClient.RunSelectedTestsAsync(testIds, debug: false);
        return JsonSerializer.Serialize(result, _jsonOptions);
    }

    [McpServerTool(Name = "test_debug_selected", Destructive = false)]
    [Description("Debug the specified Test Explorer tests by IDs returned from test_list. The debug run starts asynchronously; poll test_status for progress.")]
    public async Task<string> DebugSelectedTestsAsync(
        [Description("One or more Test Explorer test IDs returned by test_list.")] List<string> testIds)
    {
        var result = await _rpcClient.RunSelectedTestsAsync(testIds, debug: true);
        return JsonSerializer.Serialize(result, _jsonOptions);
    }

    [McpServerTool(Name = "test_results", ReadOnly = true)]
    [Description("Get structured Test Explorer results, including failed test messages, stack traces, standard output/error, duration, and attachments. By default returns failed results only; use testIds or filter to narrow the results.")]
    public async Task<string> GetTestExplorerResultsAsync(
        [Description("Optional Test Explorer IDs returned by test_list.")] List<string>? testIds = null,
        [Description("Optional case-insensitive substring filter over test name or display name.")] string? filter = null,
        [Description("Return only failed test results. Defaults to true.")] bool failedOnly = true,
        [Description("Maximum number of results to return (1-1000). Defaults to 100.")] int maxResults = 100)
    {
        var result = await _rpcClient.GetTestExplorerResultsAsync(testIds, filter, failedOnly, maxResults);
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
    [Description("Get the aggregate passed, failed, skipped and not-run counts from Test Explorer. Use test_list for individual states and test_results for failure details.")]
    public async Task<string> GetTestStatsAsync()
    {
        var stats = await _rpcClient.GetTestStatsAsync();
        if (!stats.Available)
        {
            return StatsUnavailableMessage;
        }

        return JsonSerializer.Serialize(stats, _jsonOptions);
    }
}
