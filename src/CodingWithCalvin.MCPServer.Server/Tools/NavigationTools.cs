using System.ComponentModel;
using System.Text.Json;
using System.Threading.Tasks;
using ModelContextProtocol.Server;

namespace CodingWithCalvin.MCPServer.Server.Tools;

[McpServerToolType]
public class NavigationTools
{
    private readonly RpcClient _rpcClient;
    private readonly JsonSerializerOptions _jsonOptions;

    public NavigationTools(RpcClient rpcClient)
    {
        _rpcClient = rpcClient;
        _jsonOptions = new JsonSerializerOptions { WriteIndented = true };
    }

    [McpServerTool(Name = "symbol_document", ReadOnly = true)]
    [Description("Get all symbols (classes, methods, properties, etc.) defined in a file. Returns a hierarchical list of symbols with their names, kinds, and locations. The file must be part of a project in the open solution.")]
    public async Task<string> GetDocumentSymbolsAsync(
        [Description("The full absolute path to the source file. Must be a file in a project within the open solution. Supports forward slashes (/) or backslashes (\\).")] string path)
    {
        var symbols = await _rpcClient.GetDocumentSymbolsAsync(path);
        if (symbols.Count == 0)
        {
            return "No symbols found. The file may not be part of the solution or may not have a code model (only works with C#/VB files in projects).";
        }

        return JsonSerializer.Serialize(symbols, _jsonOptions);
    }

    [McpServerTool(Name = "symbol_workspace", ReadOnly = true)]
    [Description("Search for symbols (classes, methods, properties, etc.) across the entire solution. Returns symbols matching the query with their locations. Useful for finding types or members by name.")]
    public async Task<string> SearchWorkspaceSymbolsAsync(
        [Description("The search query to match against symbol names. Case-insensitive. Partial matches are supported.")] string query,
        [Description("Maximum number of results to return. Defaults to 100. Use lower values for faster results on large solutions.")] int maxResults = 100)
    {
        var result = await _rpcClient.SearchWorkspaceSymbolsAsync(query, maxResults);
        if (result.Symbols.Count == 0)
        {
            return $"No symbols matching '{query}' found in the solution.";
        }

        return JsonSerializer.Serialize(result, _jsonOptions);
    }

    [McpServerTool(Name = "goto_definition", ReadOnly = true)]
    [Description("Navigate to the definition of a symbol at a specific position in a file. Opens the file containing the definition and returns its location. Uses Visual Studio's 'Go To Definition' feature.")]
    public async Task<string> GoToDefinitionAsync(
        [Description("The full absolute path to the source file containing the symbol reference. Supports forward slashes (/) or backslashes (\\).")] string path,
        [Description("The line number (1-based) where the symbol reference is located.")] int line,
        [Description("The column number (1-based) where the symbol reference is located. Position the cursor within or at the start of the symbol name.")] int column)
    {
        var result = await _rpcClient.GoToDefinitionAsync(path, line, column);
        if (!result.Found)
        {
            return "Definition not found. The cursor may not be on a navigable symbol, or the definition may be in external/compiled code.";
        }

        return JsonSerializer.Serialize(result, _jsonOptions);
    }

    [McpServerTool(Name = "find_references", ReadOnly = true)]
    [Description("Find all references to a symbol at a specific position in a file. Returns a list of locations where the symbol is used throughout the solution. Uses text-based search with word boundary matching.")]
    public async Task<string> FindReferencesAsync(
        [Description("The full absolute path to the source file containing the symbol. Supports forward slashes (/) or backslashes (\\).")] string path,
        [Description("The line number (1-based) where the symbol is located.")] int line,
        [Description("The column number (1-based) where the symbol is located. Position within or at the start of the symbol name.")] int column,
        [Description("Maximum number of references to return. Defaults to 100. Use lower values for faster results.")] int maxResults = 100)
    {
        var result = await _rpcClient.FindReferencesAsync(path, line, column, maxResults);
        if (!result.Found)
        {
            return "No references found. The cursor may not be on a valid identifier.";
        }

        return JsonSerializer.Serialize(result, _jsonOptions);
    }
}
