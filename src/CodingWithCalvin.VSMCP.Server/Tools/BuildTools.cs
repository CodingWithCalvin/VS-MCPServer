using System.ComponentModel;
using System.Text.Json;
using System.Threading.Tasks;
using ModelContextProtocol.Server;

namespace CodingWithCalvin.VSMCP.Server.Tools;

[McpServerToolType]
public class BuildTools
{
    private readonly RpcClient _rpcClient;

    public BuildTools(RpcClient rpcClient)
    {
        _rpcClient = rpcClient;
    }

    [McpServerTool]
    [Description("Build the entire solution")]
    public async Task<string> build_solution()
    {
        var success = await _rpcClient.BuildSolutionAsync();
        return success ? "Build started" : "Failed to start build (is a solution open?)";
    }

    [McpServerTool]
    [Description("Build a specific project")]
    public async Task<string> build_project(
        [Description("The name of the project to build")] string projectName)
    {
        var success = await _rpcClient.BuildProjectAsync(projectName);
        return success ? $"Build started for project: {projectName}" : $"Failed to build project: {projectName}";
    }

    [McpServerTool]
    [Description("Clean the entire solution (remove build outputs)")]
    public async Task<string> clean_solution()
    {
        var success = await _rpcClient.CleanSolutionAsync();
        return success ? "Clean started" : "Failed to start clean (is a solution open?)";
    }

    [McpServerTool]
    [Description("Cancel the current build operation")]
    public async Task<string> build_cancel()
    {
        await _rpcClient.CancelBuildAsync();
        return "Build cancel requested";
    }

    [McpServerTool]
    [Description("Get the current build status")]
    public async Task<string> build_status()
    {
        var status = await _rpcClient.GetBuildStatusAsync();
        return JsonSerializer.Serialize(status, new JsonSerializerOptions { WriteIndented = true });
    }
}
