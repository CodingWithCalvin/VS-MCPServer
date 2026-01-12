using System.Collections.Generic;
using System.Threading.Tasks;
using CodingWithCalvin.VSMCP.Shared.Models;

namespace CodingWithCalvin.VSMCP.Shared;

/// <summary>
/// RPC interface for Visual Studio operations.
/// Implemented by VS extension, called by MCP server process.
/// </summary>
public interface IVisualStudioRpc
{
    // Solution operations
    Task<SolutionInfo?> GetSolutionInfoAsync();
    Task<bool> OpenSolutionAsync(string path);
    Task CloseSolutionAsync(bool saveFirst);
    Task<List<ProjectInfo>> GetProjectsAsync();

    // Document operations
    Task<List<DocumentInfo>> GetOpenDocumentsAsync();
    Task<DocumentInfo?> GetActiveDocumentAsync();
    Task<bool> OpenDocumentAsync(string path);
    Task<bool> CloseDocumentAsync(string path, bool save);
    Task<string?> ReadDocumentAsync(string path);
    Task<bool> WriteDocumentAsync(string path, string content);
    Task<SelectionInfo?> GetSelectionAsync();
    Task<bool> SetSelectionAsync(string path, int startLine, int startColumn, int endLine, int endColumn);

    // Editor operations
    Task<bool> InsertTextAsync(string text);
    Task<bool> ReplaceTextAsync(string oldText, string newText);
    Task<bool> GoToLineAsync(int line);
    Task<List<FindResult>> FindAsync(string searchText, bool matchCase, bool wholeWord);

    // Build operations
    Task<bool> BuildSolutionAsync();
    Task<bool> BuildProjectAsync(string projectName);
    Task<bool> CleanSolutionAsync();
    Task CancelBuildAsync();
    Task<BuildStatus> GetBuildStatusAsync();
}
