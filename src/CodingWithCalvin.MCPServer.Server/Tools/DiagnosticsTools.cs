using System.ComponentModel;
using System.Text.Json;
using System.Threading.Tasks;
using ModelContextProtocol.Server;

namespace CodingWithCalvin.MCPServer.Server.Tools;

[McpServerToolType]
public class DiagnosticsTools
{
    private readonly RpcClient _rpcClient;
    private readonly JsonSerializerOptions _jsonOptions;

    public DiagnosticsTools(RpcClient rpcClient)
    {
        _rpcClient = rpcClient;
        _jsonOptions = new JsonSerializerOptions { WriteIndented = true };
    }

    [McpServerTool(Name = "errors_list", ReadOnly = true)]
    [Description("Get errors, warnings, and messages from the Error List. Returns diagnostics with file, line, description, and severity. Filter by severity to focus on specific issues.")]
    public async Task<string> GetErrorListAsync(
        [Description("Filter by severity: \"Error\", \"Warning\", \"Message\", or null for all. Case-insensitive.")]
        string? severity = null,
        [Description("Maximum number of items to return. Defaults to 100.")]
        int maxResults = 100)
    {
        var result = await _rpcClient.GetErrorListAsync(severity, maxResults);

        // Always return the JSON result (includes debug info if TotalCount is 0)
        return JsonSerializer.Serialize(result, _jsonOptions);
    }

    [McpServerTool(Name = "output_read", ReadOnly = true)]
    [Description("Read content from an Output window pane. Specify pane by GUID or well-known name (\"Build\", \"Debug\", \"General\"). Note: Some panes may not support reading due to VS API limitations.")]
    public async Task<string> ReadOutputPaneAsync(
        [Description("Output pane identifier: GUID string or well-known name (\"Build\", \"Debug\", \"General\").")]
        string paneIdentifier)
    {
        var result = await _rpcClient.ReadOutputPaneAsync(paneIdentifier);

        if (string.IsNullOrEmpty(result.Content))
        {
            return $"Output pane '{paneIdentifier}' is empty or does not support reading";
        }

        return JsonSerializer.Serialize(result, _jsonOptions);
    }

    [McpServerTool(Name = "output_write", Destructive = false, Idempotent = false)]
    [Description("Write a message to an Output window pane. Custom panes are auto-created. System panes (Build, Debug) must already exist. Message is appended to existing content.")]
    public async Task<string> WriteOutputPaneAsync(
        [Description("Output pane identifier: GUID string or name. Custom GUIDs/names will create new panes if needed.")]
        string paneIdentifier,
        [Description("Message to write. Appended to existing content.")]
        string message,
        [Description("Whether to activate (bring to front) the Output window. Defaults to false.")]
        bool activate = false)
    {
        var success = await _rpcClient.WriteOutputPaneAsync(paneIdentifier, message, activate);
        return success
            ? $"Message written to output pane: {paneIdentifier}"
            : $"Failed to write to output pane: {paneIdentifier}";
    }

    [McpServerTool(Name = "output_list_panes", ReadOnly = true)]
    [Description("List available Output window panes. Returns well-known panes (Build, Debug, General) with their names and GUIDs.")]
    public async Task<string> GetOutputPanesAsync()
    {
        var panes = await _rpcClient.GetOutputPanesAsync();

        if (panes.Count == 0)
        {
            return "No output panes available";
        }

        return JsonSerializer.Serialize(panes, _jsonOptions);
    }
}
