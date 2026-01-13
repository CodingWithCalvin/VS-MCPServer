using System.ComponentModel;
using System.Text.Json;
using System.Threading.Tasks;
using ModelContextProtocol.Server;

namespace CodingWithCalvin.MCPServer.Server.Tools;

[McpServerToolType]
public class DocumentTools
{
    private readonly RpcClient _rpcClient;
    private readonly JsonSerializerOptions _jsonOptions;

    public DocumentTools(RpcClient rpcClient)
    {
        _rpcClient = rpcClient;
        _jsonOptions = new JsonSerializerOptions { WriteIndented = true };
    }

    [McpServerTool(Name = "document_list", ReadOnly = true)]
    [Description("Get a list of all open documents in Visual Studio. Returns each document's Name, full Path, and IsSaved status. Use the Path value for other document operations.")]
    public async Task<string> GetDocumentListAsync()
    {
        var documents = await _rpcClient.GetOpenDocumentsAsync();
        if (documents.Count == 0)
        {
            return "No documents are currently open";
        }

        return JsonSerializer.Serialize(documents, _jsonOptions);
    }

    [McpServerTool(Name = "document_active", ReadOnly = true)]
    [Description("Get information about the currently active (focused) document in the editor. Returns the document's Name, full Path, and IsSaved status.")]
    public async Task<string> GetActiveDocumentAsync()
    {
        var doc = await _rpcClient.GetActiveDocumentAsync();
        if (doc == null)
        {
            return "No document is currently active";
        }

        return JsonSerializer.Serialize(doc, _jsonOptions);
    }

    [McpServerTool(Name = "document_open", Destructive = false, Idempotent = true)]
    [Description("Open a file in the Visual Studio editor. The file will become the active document.")]
    public async Task<string> OpenDocumentAsync(
        [Description("The full absolute path to the file. Supports forward slashes (/) or backslashes (\\).")] string path)
    {
        var success = await _rpcClient.OpenDocumentAsync(path);
        return success ? $"Opened: {path}" : $"Failed to open: {path}";
    }

    [McpServerTool(Name = "document_close", Destructive = true, Idempotent = true)]
    [Description("Close an open document in Visual Studio.")]
    public async Task<string> CloseDocumentAsync(
        [Description("The full absolute path to the document. Get this from document_list or document_active. Supports forward slashes (/) or backslashes (\\).")] string path,
        [Description("Whether to save unsaved changes before closing. Defaults to true.")] bool save = true)
    {
        var success = await _rpcClient.CloseDocumentAsync(path, save);
        return success ? $"Closed: {path}" : $"Document not found or failed to close: {path}";
    }

    [McpServerTool(Name = "document_read", ReadOnly = true)]
    [Description("Read the contents of a document. If the document is open in VS, reads the current editor buffer (including unsaved changes); otherwise reads from disk.")]
    public async Task<string> ReadDocumentAsync(
        [Description("The full absolute path to the document. Supports forward slashes (/) or backslashes (\\).")] string path)
    {
        var content = await _rpcClient.ReadDocumentAsync(path);
        return content ?? $"Could not read document: {path}";
    }

    [McpServerTool(Name = "document_write", Destructive = true)]
    [Description("Replace the entire contents of an open document. The document must already be open in VS. Changes are made to the editor buffer but not automatically saved.")]
    public async Task<string> WriteDocumentAsync(
        [Description("The full absolute path to the document. Must be open in VS. Get the path from document_list. Supports forward slashes (/) or backslashes (\\).")] string path,
        [Description("The new content to replace the entire document contents with.")] string content)
    {
        var success = await _rpcClient.WriteDocumentAsync(path, content);
        return success ? $"Updated: {path}" : $"Failed to update (is the document open?): {path}";
    }

    [McpServerTool(Name = "selection_get", ReadOnly = true)]
    [Description("Get the current text selection in the active document. Returns the selected text, start/end line and column positions (1-based), and the document path.")]
    public async Task<string> GetSelectionAsync()
    {
        var selection = await _rpcClient.GetSelectionAsync();
        if (selection == null)
        {
            return "No active document or selection";
        }

        return JsonSerializer.Serialize(selection, _jsonOptions);
    }

    [McpServerTool(Name = "selection_set", Destructive = false, Idempotent = true)]
    [Description("Set the text selection in an open document. The document must already be open in VS. Use the same start and end positions to place the cursor without selecting text.")]
    public async Task<string> SetSelectionAsync(
        [Description("The full absolute path to the document. Must be open in VS. Supports forward slashes (/) or backslashes (\\).")] string path,
        [Description("Starting line number (1-based, first line is 1).")] int startLine,
        [Description("Starting column number (1-based, first column is 1).")] int startColumn,
        [Description("Ending line number (1-based). Use same as startLine to place cursor on single line.")] int endLine,
        [Description("Ending column number (1-based). Use same as startColumn to place cursor without selection.")] int endColumn)
    {
        var success = await _rpcClient.SetSelectionAsync(path, startLine, startColumn, endLine, endColumn);
        return success ? "Selection set" : "Failed to set selection (is the document open?)";
    }

    [McpServerTool(Name = "editor_insert", Destructive = false)]
    [Description("Insert text at the current cursor position (or replace current selection) in the active document. Use selection_set first to position the cursor.")]
    public async Task<string> InsertTextAsync(
        [Description("The text to insert. Can include newlines for multi-line inserts.")] string text)
    {
        var success = await _rpcClient.InsertTextAsync(text);
        return success ? "Text inserted" : "Failed to insert text (no active document?)";
    }

    [McpServerTool(Name = "editor_replace", Destructive = true, Idempotent = true)]
    [Description("Find and replace ALL occurrences of text in the active document. Performs a case-sensitive search and replaces every match.")]
    public async Task<string> ReplaceTextAsync(
        [Description("The exact text to find (case-sensitive).")] string oldText,
        [Description("The replacement text. Use empty string to delete matches.")] string newText)
    {
        var success = await _rpcClient.ReplaceTextAsync(oldText, newText);
        return success ? "Text replaced" : "Text not found or no active document";
    }

    [McpServerTool(Name = "editor_goto_line", Destructive = false, Idempotent = true)]
    [Description("Navigate to a specific line in the active document and place the cursor at the beginning of that line.")]
    public async Task<string> GoToLineAsync(
        [Description("The line number to navigate to (1-based, first line is 1).")] int line)
    {
        var success = await _rpcClient.GoToLineAsync(line);
        return success ? $"Navigated to line {line}" : "Failed to navigate (no active document?)";
    }

    [McpServerTool(Name = "editor_find", ReadOnly = true)]
    [Description("Search for all occurrences of text in the active document. Returns a list of matches with line numbers, column positions, and the matching line content.")]
    public async Task<string> FindTextAsync(
        [Description("The text to search for.")] string searchText,
        [Description("Whether to match case exactly. Defaults to false (case-insensitive).")] bool matchCase = false,
        [Description("Whether to match whole words only (not partial matches). Defaults to false.")] bool wholeWord = false)
    {
        var results = await _rpcClient.FindAsync(searchText, matchCase, wholeWord);
        if (results.Count == 0)
        {
            return "No matches found";
        }

        return JsonSerializer.Serialize(results, _jsonOptions);
    }
}
