using System.ComponentModel;
using System.Text.Json;
using System.Threading.Tasks;
using ModelContextProtocol.Server;

namespace CodingWithCalvin.VSMCP.Server.Tools;

[McpServerToolType]
public class DocumentTools
{
    private readonly RpcClient _rpcClient;

    public DocumentTools(RpcClient rpcClient)
    {
        _rpcClient = rpcClient;
    }

    [McpServerTool]
    [Description("Get a list of all open documents in Visual Studio")]
    public async Task<string> document_list()
    {
        var documents = await _rpcClient.GetOpenDocumentsAsync();
        if (documents.Count == 0)
        {
            return "No documents are currently open";
        }

        return JsonSerializer.Serialize(documents, new JsonSerializerOptions { WriteIndented = true });
    }

    [McpServerTool]
    [Description("Get information about the currently active document")]
    public async Task<string> document_active()
    {
        var doc = await _rpcClient.GetActiveDocumentAsync();
        if (doc == null)
        {
            return "No document is currently active";
        }

        return JsonSerializer.Serialize(doc, new JsonSerializerOptions { WriteIndented = true });
    }

    [McpServerTool]
    [Description("Open a file in Visual Studio")]
    public async Task<string> document_open(
        [Description("The full path to the file to open")] string path)
    {
        var success = await _rpcClient.OpenDocumentAsync(path);
        return success ? $"Opened: {path}" : $"Failed to open: {path}";
    }

    [McpServerTool]
    [Description("Close a document in Visual Studio")]
    public async Task<string> document_close(
        [Description("The full path to the document to close")] string path,
        [Description("Whether to save changes before closing")] bool save = true)
    {
        var success = await _rpcClient.CloseDocumentAsync(path, save);
        return success ? $"Closed: {path}" : $"Document not found or failed to close: {path}";
    }

    [McpServerTool]
    [Description("Read the contents of a document. If the document is open in VS, reads from the editor buffer; otherwise reads from disk.")]
    public async Task<string> document_read(
        [Description("The full path to the document")] string path)
    {
        var content = await _rpcClient.ReadDocumentAsync(path);
        return content ?? $"Could not read document: {path}";
    }

    [McpServerTool]
    [Description("Write content to an open document in Visual Studio")]
    public async Task<string> document_write(
        [Description("The full path to the document")] string path,
        [Description("The new content for the document")] string content)
    {
        var success = await _rpcClient.WriteDocumentAsync(path, content);
        return success ? $"Updated: {path}" : $"Failed to update (is the document open?): {path}";
    }

    [McpServerTool]
    [Description("Get the current text selection in the active document")]
    public async Task<string> selection_get()
    {
        var selection = await _rpcClient.GetSelectionAsync();
        if (selection == null)
        {
            return "No active document or selection";
        }

        return JsonSerializer.Serialize(selection, new JsonSerializerOptions { WriteIndented = true });
    }

    [McpServerTool]
    [Description("Set the text selection in a document")]
    public async Task<string> selection_set(
        [Description("The full path to the document")] string path,
        [Description("Starting line number (1-based)")] int startLine,
        [Description("Starting column number (1-based)")] int startColumn,
        [Description("Ending line number (1-based)")] int endLine,
        [Description("Ending column number (1-based)")] int endColumn)
    {
        var success = await _rpcClient.SetSelectionAsync(path, startLine, startColumn, endLine, endColumn);
        return success ? "Selection set" : "Failed to set selection (is the document open?)";
    }

    [McpServerTool]
    [Description("Insert text at the current cursor position in the active document")]
    public async Task<string> editor_insert(
        [Description("The text to insert")] string text)
    {
        var success = await _rpcClient.InsertTextAsync(text);
        return success ? "Text inserted" : "Failed to insert text (no active document?)";
    }

    [McpServerTool]
    [Description("Find and replace text in the active document")]
    public async Task<string> editor_replace(
        [Description("The text to find")] string oldText,
        [Description("The replacement text")] string newText)
    {
        var success = await _rpcClient.ReplaceTextAsync(oldText, newText);
        return success ? "Text replaced" : "Text not found or no active document";
    }

    [McpServerTool]
    [Description("Navigate to a specific line in the active document")]
    public async Task<string> editor_goto_line(
        [Description("The line number to navigate to (1-based)")] int line)
    {
        var success = await _rpcClient.GoToLineAsync(line);
        return success ? $"Navigated to line {line}" : "Failed to navigate (no active document?)";
    }

    [McpServerTool]
    [Description("Search for text in the active document")]
    public async Task<string> editor_find(
        [Description("The text to search for")] string searchText,
        [Description("Whether to match case")] bool matchCase = false,
        [Description("Whether to match whole words only")] bool wholeWord = false)
    {
        var results = await _rpcClient.FindAsync(searchText, matchCase, wholeWord);
        if (results.Count == 0)
        {
            return "No matches found";
        }

        return JsonSerializer.Serialize(results, new JsonSerializerOptions { WriteIndented = true });
    }
}
