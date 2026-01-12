using System.ComponentModel;
using System.Text.Json;
using System.Threading.Tasks;
using ModelContextProtocol.Server;

namespace CodingWithCalvin.VSMCP.Server.Tools;

[McpServerToolType]
public class SolutionTools
{
    private readonly RpcClient _rpcClient;

    public SolutionTools(RpcClient rpcClient)
    {
        _rpcClient = rpcClient;
    }

    [McpServerTool]
    [Description("Get information about the currently open solution in Visual Studio")]
    public async Task<string> solution_info()
    {
        var info = await _rpcClient.GetSolutionInfoAsync();
        if (info == null)
        {
            return "No solution is currently open";
        }

        return JsonSerializer.Serialize(info, new JsonSerializerOptions { WriteIndented = true });
    }

    [McpServerTool]
    [Description("Open a solution file in Visual Studio")]
    public async Task<string> solution_open(
        [Description("The full path to the solution file (.sln or .slnx)")] string path)
    {
        var success = await _rpcClient.OpenSolutionAsync(path);
        return success ? $"Opened solution: {path}" : $"Failed to open solution: {path}";
    }

    [McpServerTool]
    [Description("Close the currently open solution in Visual Studio")]
    public async Task<string> solution_close(
        [Description("Whether to save changes before closing")] bool save = true)
    {
        await _rpcClient.CloseSolutionAsync(save);
        return "Solution closed";
    }

    [McpServerTool]
    [Description("Get a list of all projects in the current solution")]
    public async Task<string> project_list()
    {
        var projects = await _rpcClient.GetProjectsAsync();
        if (projects.Count == 0)
        {
            return "No projects found (is a solution open?)";
        }

        return JsonSerializer.Serialize(projects, new JsonSerializerOptions { WriteIndented = true });
    }

    [McpServerTool]
    [Description("Get detailed information about a specific project")]
    public async Task<string> project_info(
        [Description("The name of the project")] string name)
    {
        var projects = await _rpcClient.GetProjectsAsync();
        var project = projects.Find(p => p.Name == name);

        if (project == null)
        {
            return $"Project not found: {name}";
        }

        return JsonSerializer.Serialize(project, new JsonSerializerOptions { WriteIndented = true });
    }
}
