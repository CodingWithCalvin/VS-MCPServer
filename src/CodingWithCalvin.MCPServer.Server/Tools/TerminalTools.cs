using System.ComponentModel;
using System.Text.Json;
using System.Threading.Tasks;
using ModelContextProtocol.Server;

namespace CodingWithCalvin.MCPServer.Server.Tools;

[McpServerToolType]
public class TerminalTools
{
    private readonly RpcClient _rpcClient;
    private readonly JsonSerializerOptions _jsonOptions;

    public TerminalTools(RpcClient rpcClient)
    {
        _rpcClient = rpcClient;
        _jsonOptions = new JsonSerializerOptions { WriteIndented = true };
    }

    [McpServerTool(Name = "terminal_run", Destructive = true)]
    [Description("Run a command in a new Visual Studio integrated terminal. The command is launched inside the Visual Studio developer environment, so msbuild, vstest.console and dotnet-coverage are on PATH. IMPORTANT: output is NOT returned - the Visual Studio terminal is a raw PTY with no exit code or command boundaries, so this tool reports only whether the terminal was opened. To read results, redirect the command's output to a file and then read it with document_read.")]
    public async Task<string> RunInTerminalAsync(
        [Description("Command line to run, for example 'dotnet-coverage collect -f cobertura -o coverage.xml dotnet test'.")] string command,
        [Description("Working directory for the command. Defaults to the directory containing the open solution.")] string? workingDirectory = null,
        [Description("Caption for the terminal tab. Defaults to a Visual Studio generated name.")] string? name = null)
    {
        if (string.IsNullOrWhiteSpace(command))
        {
            return "A command is required.";
        }

        var result = await _rpcClient.CreateTerminalAsync(name, workingDirectory, command);
        return JsonSerializer.Serialize(result, _jsonOptions);
    }

    [McpServerTool(Name = "terminal_create", Destructive = false)]
    [Description("Open a new empty Visual Studio integrated terminal using the user's default terminal profile, without running anything in it. Use terminal_run instead to launch a command.")]
    public async Task<string> CreateTerminalAsync(
        [Description("Caption for the terminal tab. Defaults to a Visual Studio generated name.")] string? name = null,
        [Description("Working directory for the terminal. Defaults to the directory containing the open solution.")] string? workingDirectory = null)
    {
        var result = await _rpcClient.CreateTerminalAsync(name, workingDirectory, command: null);
        return JsonSerializer.Serialize(result, _jsonOptions);
    }

    [McpServerTool(Name = "terminal_list", ReadOnly = true)]
    [Description("List the identifiers of every integrated terminal currently open in Visual Studio. Use these identifiers with terminal_show and terminal_close.")]
    public async Task<string> ListTerminalsAsync()
    {
        var result = await _rpcClient.GetTerminalsAsync();
        return JsonSerializer.Serialize(result, _jsonOptions);
    }

    [McpServerTool(Name = "terminal_show", Destructive = false, Idempotent = true)]
    [Description("Bring an integrated terminal into view. Get the identifier from terminal_run, terminal_create or terminal_list.")]
    public async Task<string> ShowTerminalAsync(
        [Description("Terminal identifier (a GUID) returned by terminal_run, terminal_create or terminal_list.")] string terminalId)
    {
        var shown = await _rpcClient.ShowTerminalAsync(terminalId);
        return shown
            ? $"Terminal {terminalId} shown"
            : $"Could not show terminal {terminalId}. Call terminal_list to see the open terminals.";
    }

    [McpServerTool(Name = "terminal_close", Destructive = true)]
    [Description("Close a single integrated terminal, ending any process running in it. Get the identifier from terminal_list.")]
    public async Task<string> CloseTerminalAsync(
        [Description("Terminal identifier (a GUID) returned by terminal_run, terminal_create or terminal_list.")] string terminalId)
    {
        var closed = await _rpcClient.CloseTerminalAsync(terminalId);
        return closed
            ? $"Terminal {terminalId} closed"
            : $"Could not close terminal {terminalId}. Call terminal_list to see the open terminals.";
    }

    [McpServerTool(Name = "terminal_close_all", Destructive = true)]
    [Description("Close every integrated terminal, ending any processes running in them.")]
    public async Task<string> CloseAllTerminalsAsync()
    {
        var closed = await _rpcClient.CloseAllTerminalsAsync();
        return closed
            ? "All integrated terminals closed"
            : "Could not close the terminals. The Visual Studio terminal component may be unavailable.";
    }
}
