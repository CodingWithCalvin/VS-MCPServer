using System.ComponentModel;
using System.Text.Json;
using System.Threading.Tasks;
using ModelContextProtocol.Server;

namespace CodingWithCalvin.MCPServer.Server.Tools;

[McpServerToolType]
public class CoverageTools
{
    private readonly RpcClient _rpcClient;
    private readonly JsonSerializerOptions _jsonOptions;

    public CoverageTools(RpcClient rpcClient)
    {
        _rpcClient = rpcClient;
        _jsonOptions = new JsonSerializerOptions { WriteIndented = true };
    }

    [McpServerTool(Name = "coverage_analyze", Destructive = false)]
    [Description("Run all tests with code coverage collection enabled. The run starts asynchronously and this returns immediately; poll test_status until it reports 'Completed', then call coverage_report. Requires an edition of Visual Studio with code coverage: Enterprise only through VS 2022, all editions from VS 2026.")]
    public async Task<string> AnalyzeCoverageAsync()
    {
        var result = await _rpcClient.AnalyzeCodeCoverageAsync();

        if (!result.Supported)
        {
            return result.Message ?? "Code coverage is not supported in this edition of Visual Studio.";
        }

        return result.Started
            ? "Coverage run started. Poll test_status for completion, then call coverage_report."
            : result.Message ?? "Failed to start the coverage run.";
    }

    [McpServerTool(Name = "coverage_report", ReadOnly = true)]
    [Description("Read code coverage results as a module, class and method tree with covered/uncovered line and block counts. Defaults to the newest .coverage file under the solution's TestResults folder. Unlike coverage_analyze, this works on every edition of Visual Studio, so an existing .coverage file produced elsewhere can be read on VS 2022 Community.")]
    public async Task<string> GetCoverageReportAsync(
        [Description("Detail level: 'summary' for totals per module, 'class' (default) for per-class breakdown, or 'method' for the fully expanded tree.")] string? detail = null,
        [Description("Only include modules or classes whose name contains this text, for example 'OrderService'. Case-insensitive.")] string? filter = null,
        [Description("Explicit path to a .coverage file. Omit to use the newest one under the solution's TestResults folder.")] string? coverageFile = null)
    {
        var result = await _rpcClient.GetCoverageReportAsync(coverageFile, detail, filter);

        if (!result.Available)
        {
            return result.Message ?? "No coverage results are available.";
        }

        return JsonSerializer.Serialize(result, _jsonOptions);
    }

    [McpServerTool(Name = "coverage_show", Destructive = false, Idempotent = true)]
    [Description("Open the Visual Studio Code Coverage Results window so the user can see the results and the editor coverage coloring.")]
    public async Task<string> ShowCoverageResultsAsync()
    {
        var shown = await _rpcClient.ShowCoverageResultsAsync();
        return shown
            ? "Code Coverage Results window shown"
            : "Could not open the Code Coverage Results window. This edition of Visual Studio may not support code coverage.";
    }
}
