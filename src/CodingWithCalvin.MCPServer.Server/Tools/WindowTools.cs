using System.ComponentModel;
using System.Text.Json;
using System.Threading.Tasks;
using ModelContextProtocol.Server;

namespace CodingWithCalvin.MCPServer.Server.Tools;

[McpServerToolType]
public class WindowTools
{
    private static readonly string[] SupportedToolWindows =
    [
        "SolutionExplorer",
        "ErrorList",
        "Output",
        "TeamExplorer",
        "Terminal",
        "TaskList",
        "Properties",
        "Toolbox",
        "FindResults",
        "Bookmarks",
    ];

    private readonly RpcClient _rpcClient;
    private readonly JsonSerializerOptions _jsonOptions;

    public WindowTools(RpcClient rpcClient)
    {
        _rpcClient = rpcClient;
        _jsonOptions = new JsonSerializerOptions { WriteIndented = true };
    }

    [McpServerTool(Name = "window_list", ReadOnly = true)]
    [Description("List all open windows in Visual Studio. Returns each window's caption, kind (Document or Tool), visibility, and GUID.")]
    public async Task<string> GetWindowsAsync()
    {
        var windows = await _rpcClient.GetWindowsAsync();

        if (windows.Count == 0)
        {
            return "No windows found";
        }

        return JsonSerializer.Serialize(windows, _jsonOptions);
    }

    [McpServerTool(Name = "window_activate", Destructive = false, Idempotent = true)]
    [Description("Activate (bring to front and focus) a specific window by its caption. Use window_list to find available window captions.")]
    public async Task<string> ActivateWindowAsync(
        [Description("The caption/title of the window to activate. Case-insensitive.")]
        string caption)
    {
        var success = await _rpcClient.ActivateWindowAsync(caption);
        return success
            ? $"Activated window: {caption}"
            : $"Window not found: {caption}";
    }

    [McpServerTool(Name = "toolwindow_show", Destructive = false, Idempotent = true)]
    [Description("Show a tool window by well-known name. Supported names: SolutionExplorer, ErrorList, Output, TeamExplorer, Terminal, TaskList, Properties, Toolbox, FindResults, Bookmarks.")]
    public async Task<string> ShowToolWindowAsync(
        [Description("Well-known tool window name (e.g., \"SolutionExplorer\", \"ErrorList\", \"Output\"). Case-insensitive.")]
        string name)
    {
        var success = await _rpcClient.ShowToolWindowAsync(name);

        if (success)
        {
            return $"Shown tool window: {name}";
        }

        var supported = string.Join(", ", SupportedToolWindows);
        return $"Unknown tool window: {name}. Supported names: {supported}";
    }

    [McpServerTool(Name = "toolwindow_hide", Destructive = false, Idempotent = true)]
    [Description("Hide (close) a tool window by its caption. Use window_list to find available window captions.")]
    public async Task<string> HideToolWindowAsync(
        [Description("The caption/title of the tool window to hide. Case-insensitive.")]
        string caption)
    {
        var success = await _rpcClient.HideToolWindowAsync(caption);
        return success
            ? $"Hidden tool window: {caption}"
            : $"Tool window not found: {caption}";
    }
}
