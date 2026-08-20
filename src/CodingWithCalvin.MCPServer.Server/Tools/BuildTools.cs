using System.ComponentModel;
using System.Text.Json;
using System.Threading.Tasks;
using ModelContextProtocol.Server;

namespace CodingWithCalvin.MCPServer.Server.Tools;

[McpServerToolType]
public class BuildTools
{
    private const string DebugSessionActiveMessage =
        "A debug session is currently active. Stop debugging first using debugger_stop before building.";
    private const string ConfigurationDebugSessionActiveMessage =
        "A debug session is currently active. Stop debugging first using debugger_stop before changing the build configuration.";

    private readonly RpcClient _rpcClient;
    private readonly JsonSerializerOptions _jsonOptions;

    public BuildTools(RpcClient rpcClient)
    {
        _rpcClient = rpcClient;
        _jsonOptions = new JsonSerializerOptions { WriteIndented = true };
    }

    private async Task<bool> IsDebuggingActiveAsync()
    {
        var status = await _rpcClient.GetDebuggerStatusAsync();
        return status.Mode != "Design";
    }

    [McpServerTool(Name = "build_solution", Destructive = false)]
    [Description("Build the entire solution. The build runs asynchronously; use build_status to check progress. Returns immediately after starting the build. If a debug session is active, the build cannot proceed — use debugger_stop first.")]
    public async Task<string> BuildSolutionAsync()
    {
        if (await IsDebuggingActiveAsync())
        {
            return DebugSessionActiveMessage;
        }

        var success = await _rpcClient.BuildSolutionAsync();
        return success ? "Build started" : "Failed to start build (is a solution open?)";
    }

    [McpServerTool(Name = "build_project", Destructive = false)]
    [Description("Build a specific project. The build runs asynchronously; use build_status to check progress. IMPORTANT: Requires the full path to the .csproj file, not just the project name. Use project_list first to get the correct path. If a debug session is active, the build cannot proceed — use debugger_stop first.")]
    public async Task<string> BuildProjectAsync(
        [Description("The full absolute path to the project file (.csproj). Get this from project_list. Supports forward slashes (/) or backslashes (\\).")] string projectName)
    {
        if (await IsDebuggingActiveAsync())
        {
            return DebugSessionActiveMessage;
        }

        var success = await _rpcClient.BuildProjectAsync(projectName);
        return success ? $"Build started for project: {projectName}" : $"Failed to build project: {projectName}";
    }

    [McpServerTool(Name = "clean_solution", Destructive = true, Idempotent = true)]
    [Description("Clean the entire solution by removing all build outputs (bin/obj folders). The clean runs asynchronously; use build_status to check progress. If a debug session is active, the clean cannot proceed — use debugger_stop first.")]
    public async Task<string> CleanSolutionAsync()
    {
        if (await IsDebuggingActiveAsync())
        {
            return DebugSessionActiveMessage;
        }

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

    [McpServerTool(Name = "build_configuration_get", ReadOnly = true)]
    [Description("Get the active solution build configuration and platform, plus every available configuration/platform pair. Use this before build_configuration_set to discover exact values such as 'Release' and 'x64'.")]
    public async Task<string> GetBuildConfigurationAsync()
    {
        var configuration = await _rpcClient.GetBuildConfigurationAsync();
        return JsonSerializer.Serialize(configuration, _jsonOptions);
    }

    [McpServerTool(Name = "build_configuration_set", Destructive = false, Idempotent = true)]
    [Description("Activate an existing solution build configuration and platform in Visual Studio, such as 'Release' and 'x64'. Call build_configuration_get first because only listed pairs are accepted. Cannot change configuration while a build or debug session is active.")]
    public async Task<string> SetBuildConfigurationAsync(
        [Description("Solution configuration name, such as 'Debug' or 'Release'.")] string configuration,
        [Description("Solution platform/architecture, such as 'Any CPU', 'x64', or 'ARM64'.")] string platform)
    {
        if (string.IsNullOrWhiteSpace(configuration) || string.IsNullOrWhiteSpace(platform))
        {
            return "Configuration and platform are required. Call build_configuration_get for available values.";
        }

        if (await IsDebuggingActiveAsync())
        {
            return ConfigurationDebugSessionActiveMessage;
        }

        var success = await _rpcClient.SetBuildConfigurationAsync(configuration.Trim(), platform.Trim());
        if (!success)
        {
            return $"Failed to activate build configuration '{configuration}|{platform}'. A build may be running or the pair may not exist. Call build_configuration_get for available values.";
        }

        var activeConfiguration = await _rpcClient.GetBuildConfigurationAsync();
        return $"Active build configuration: {activeConfiguration.ActiveConfiguration}|{activeConfiguration.ActivePlatform}";
    }
}
