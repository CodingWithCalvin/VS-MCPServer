using System.ComponentModel;
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
    [Description("Get the passed, failed, skipped and not-run test counts from Test Explorer. Returns counts only - Visual Studio exposes no public API for the names of the individual failed or not-run tests.")]
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
