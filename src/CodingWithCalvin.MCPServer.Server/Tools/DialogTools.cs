using System.ComponentModel;
using System.Text.Json;
using System.Threading.Tasks;
using ModelContextProtocol.Server;

namespace CodingWithCalvin.MCPServer.Server.Tools;

[McpServerToolType]
public class DialogTools
{
    private readonly RpcClient _rpcClient;
    private readonly JsonSerializerOptions _jsonOptions;

    public DialogTools(RpcClient rpcClient)
    {
        _rpcClient = rpcClient;
        _jsonOptions = new JsonSerializerOptions { WriteIndented = true };
    }

    [McpServerTool(Name = "dialog_list", ReadOnly = true)]
    [Description("List the modal dialogs Visual Studio is waiting on, such as message boxes and prompts raised by the debugger, Hot Reload or a build. A modal dialog blocks Visual Studio until someone answers it, so if another tool call hangs or times out, call this to check for one. Works while another tool call is still blocked. Returns each dialog's id, title, message and buttons; answer it with dialog_respond.")]
    public async Task<string> ListDialogsAsync()
    {
        var result = await _rpcClient.GetDialogsAsync();

        if (result.Dialogs.Count == 0)
        {
            return result.Message ?? "No dialogs are open. Visual Studio is not waiting on a dialog.";
        }

        return JsonSerializer.Serialize(result.Dialogs, _jsonOptions);
    }

    [McpServerTool(Name = "dialog_respond", Destructive = true)]
    [Description("Answer a modal dialog by clicking one of its buttons, exactly as a user would. Read the dialog with dialog_list first: the answer takes effect immediately and can discard changes, stop debugging or start a rebuild. Returns whether the dialog closed and the dialogs still open, because an answer can raise a follow-up prompt.")]
    public async Task<string> RespondToDialogAsync(
        [Description("Name of the button to click, as listed by dialog_list, for example \"Yes\", \"No\" or \"Cancel\". Case-insensitive.")] string button,
        [Description("Id of the dialog to answer, from dialog_list. Only needed when more than one dialog is open.")] string? dialogId = null)
    {
        if (string.IsNullOrWhiteSpace(button))
        {
            return "A button name is required. Call dialog_list to see the dialog's buttons.";
        }

        var result = await _rpcClient.RespondToDialogAsync(button, dialogId);
        return JsonSerializer.Serialize(result, _jsonOptions);
    }
}
