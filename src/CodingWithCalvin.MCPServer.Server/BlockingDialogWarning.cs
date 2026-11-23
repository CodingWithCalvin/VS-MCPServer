using System;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Threading.Tasks;
using ModelContextProtocol.Protocol;

namespace CodingWithCalvin.MCPServer.Server;

/// <summary>
/// Appends the fixed blocking-dialog warning to the output of an ordinary tool call.
/// </summary>
/// <remarks>
/// <para>
/// A modal dialog holds the Visual Studio UI thread, so a tool call can return while Visual
/// Studio is stuck waiting on a dialog nobody has noticed. The warning is only added at the point
/// where the tool was already going to return: the tool's own work, waiting and output are left
/// untouched, and no dialog is answered or closed.
/// </para>
/// <para>
/// <c>dialog_list</c> and <c>dialog_respond</c> are exempt. They exist to find and answer dialogs,
/// so warning them about the dialog they are handling would only add noise.
/// </para>
/// </remarks>
internal static class BlockingDialogWarning
{
    internal const string Message =
        "Operation may not proceed - Visual Studio is waiting on a dialog. Please close the dialog first.";

    private static readonly JsonSerializerOptions JsonOptions = new() { WriteIndented = true };

    private static readonly string[] ExemptTools = { "dialog_list", "dialog_respond" };

    internal static bool IsExempt(string? toolName) =>
        toolName != null && Array.IndexOf(ExemptTools, toolName) >= 0;

    /// <summary>
    /// Returns <paramref name="result"/> unchanged unless Visual Studio is waiting on a dialog at
    /// the moment the tool returns; otherwise wraps the original output together with the warning.
    /// </summary>
    internal static async Task<CallToolResult> ApplyAsync(CallToolResult result, RpcClient rpcClient)
    {
        if (result.Content is not { Count: > 0 })
        {
            return result;
        }

        bool blocked;

        try
        {
            blocked = await rpcClient.HasBlockingDialogsAsync().ConfigureAwait(false);
        }
        catch (Exception)
        {
            // The check is advisory: losing the connection must not change what the tool returned.
            return result;
        }

        if (!blocked)
        {
            return result;
        }

        foreach (var block in result.Content)
        {
            if (block is TextContentBlock text)
            {
                text.Text = Wrap(text.Text ?? string.Empty);
                break;
            }
        }

        return result;
    }

    /// <summary>
    /// Wraps the original text, keeping output that is already a JSON object or array as JSON
    /// instead of escaping it into a string.
    /// </summary>
    private static string Wrap(string text)
    {
        var wrapper = new JsonObject
        {
            ["result"] = TryParseJson(text) ?? JsonValue.Create(text),
            ["warning"] = Message,
        };

        return wrapper.ToJsonString(JsonOptions);
    }

    private static JsonNode? TryParseJson(string text)
    {
        var first = text.TrimStart();
        if (first.Length == 0 || (first[0] != '{' && first[0] != '['))
        {
            return null;
        }

        try
        {
            return JsonNode.Parse(text);
        }
        catch (JsonException)
        {
            return null;
        }
    }
}
