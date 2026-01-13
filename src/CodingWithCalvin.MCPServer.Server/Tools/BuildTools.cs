using System.ComponentModel;
using System.Text.Json;
using System.Threading.Tasks;
using ModelContextProtocol.Server;

namespace CodingWithCalvin.MCPServer.Server.Tools;

[McpServerToolType]
public class BuildTools
{
    private readonly RpcClient _rpcClient;
    private readonly JsonSerializerOptions _jsonOptions;

    public BuildTools(RpcClient rpcClient)
    {
        _rpcClient = rpcClient;
        _jsonOptions = new JsonSerializerOptions { WriteIndented = true };
    }

    [McpServerTool(Name = "build_solution", Destructive = false)]
    [Description("Build the entire solution. The build runs asynchronously; use build_status to check progress. Returns immediately after starting the build.")]
    public async Task<string> BuildSolutionAsync()
    {
        var success = await _rpcClient.BuildSolutionAsync();
        return success ? "Build started" : "Failed to start build (is a solution open?)";
    }

    [McpServerTool(Name = "build_project", Destructive = false)]
    [Description("Build a specific project. The build runs asynchronously; use build_status to check progress. IMPORTANT: Requires the full path to the .csproj file, not just the project name. Use project_list first to get the correct path.")]
    public async Task<string> BuildProjectAsync(
        [Description("The full absolute path to the project file (.csproj). Get this from project_list. Supports forward slashes (/) or backslashes (\\).")] string projectName)
    {
        var success = await _rpcClient.BuildProjectAsync(projectName);
        return success ? $"Build started for project: {projectName}" : $"Failed to build project: {projectName}";
    }

    [McpServerTool(Name = "clean_solution", Destructive = true, Idempotent = true)]
    [Description("Clean the entire solution by removing all build outputs (bin/obj folders). The clean runs asynchronously; use build_status to check progress.")]
    public async Task<string> CleanSolutionAsync()
    {
        var success = await _rpcClient.CleanSolutionAsync();
        return success ? "Clean started" : "Failed to start clean (is a solution open?)";
    }

    [McpServerTool(Name = "build_cancel", Destructive = false, Idempotent = true)]
    [Description("Cancel the current build or clean operation if one is in progress.")]
    public async Task<string> CancelBuildAsync()
    {
        var cancelled = await _rpcClient.CancelBuildAsync();
        return cancelled ? "Build cancelled" : "No build is currently in progress";
    }

    [McpServerTool(Name = "build_status", ReadOnly = true)]
    [Description("Get the current build status. Returns State ('NoBuildPerformed', 'InProgress', or 'Done') and FailedProjects count. Use this to poll for build completion after starting a build.")]
    public async Task<string> GetBuildStatusAsync()
    {
        var status = await _rpcClient.GetBuildStatusAsync();
        return JsonSerializer.Serialize(status, _jsonOptions);
    }
}
