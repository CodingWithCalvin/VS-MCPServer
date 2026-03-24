using System.ComponentModel;
using System.Text.Json;
using System.Threading.Tasks;
using ModelContextProtocol.Server;

namespace CodingWithCalvin.MCPServer.Server.Tools;

[McpServerToolType]
public class SolutionTools
{
    private readonly RpcClient _rpcClient;
    private readonly JsonSerializerOptions _jsonOptions;

    public SolutionTools(RpcClient rpcClient)
    {
        _rpcClient = rpcClient;
        _jsonOptions = new JsonSerializerOptions { WriteIndented = true };
    }

    [McpServerTool(Name = "solution_info", ReadOnly = true)]
    [Description("Get information about the currently open solution in Visual Studio. Returns the solution name, full path, and whether it's open. Use this to verify a solution is loaded before performing other operations.")]
    public async Task<string> GetSolutionInfoAsync()
    {
        var info = await _rpcClient.GetSolutionInfoAsync();
        if (info == null)
        {
            return "No solution is currently open";
        }

        return JsonSerializer.Serialize(info, _jsonOptions);
    }

    [McpServerTool(Name = "solution_open", Destructive = true, Idempotent = true)]
    [Description("Open a solution file in Visual Studio. This will close any currently open solution.")]
    public async Task<string> OpenSolutionAsync(
        [Description("The full absolute path to the solution file (.sln or .slnx). Supports forward slashes (/) or backslashes (\\).")] string path)
    {
        var success = await _rpcClient.OpenSolutionAsync(path);
        return success ? $"Opened solution: {path}" : $"Failed to open solution: {path}";
    }

    [McpServerTool(Name = "solution_close", Destructive = true, Idempotent = true)]
    [Description("Close the currently open solution in Visual Studio")]
    public async Task<string> CloseSolutionAsync(
        [Description("Whether to save changes before closing")] bool save = true)
    {
        await _rpcClient.CloseSolutionAsync(save);
        return "Solution closed";
    }

    [McpServerTool(Name = "project_list", ReadOnly = true)]
    [Description("Get a list of all projects in the current solution. Returns each project's Name, full Path (.csproj), and Kind (GUID). Use the Path value when calling build_project.")]
    public async Task<string> GetProjectListAsync()
    {
        var projects = await _rpcClient.GetProjectsAsync();
        if (projects.Count == 0)
        {
            return "No projects found (is a solution open?)";
        }

        return JsonSerializer.Serialize(projects, _jsonOptions);
    }

    [McpServerTool(Name = "startup_project_get", ReadOnly = true)]
    [Description("Get the current startup project name. Returns the project that will be launched when debugging starts.")]
    public async Task<string> GetStartupProjectAsync()
    {
        var startupProject = await _rpcClient.GetStartupProjectAsync();
        return startupProject ?? "No startup project is set";
    }

    [McpServerTool(Name = "startup_project_set", Destructive = false)]
    [Description("Set the startup project for debugging. Use project_list to get available project names.")]
    public async Task<string> SetStartupProjectAsync(
        [Description("The display name of the project to set as the startup project (e.g., 'MyProject'). Use project_list to see available project names.")] string name)
    {
        var success = await _rpcClient.SetStartupProjectAsync(name);
        return success ? $"Startup project set to: {name}" : $"Failed to set startup project: {name}";
    }

    [McpServerTool(Name = "project_info", ReadOnly = true)]
    [Description("Get detailed information about a specific project by its display name.")]
    public async Task<string> GetProjectInfoAsync(
        [Description("The display name of the project (e.g., 'MyProject'), not the full path. Use project_list to see available project names.")] string name)
    {
        var projects = await _rpcClient.GetProjectsAsync();
        var project = projects.Find(p => p.Name == name);

        if (project == null)
        {
            return $"Project not found: {name}";
        }

        return JsonSerializer.Serialize(project, _jsonOptions);
    }
}
