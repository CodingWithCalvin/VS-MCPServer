using System.Collections.Generic;
using System.Threading.Tasks;
using CodingWithCalvin.MCPServer.Shared.Models;

namespace CodingWithCalvin.MCPServer.Services;

public interface IVisualStudioService
{
    Task<SolutionInfo?> GetSolutionInfoAsync();
    Task<bool> OpenSolutionAsync(string path);
    Task CloseSolutionAsync(bool saveFirst = true);
    Task<List<ProjectInfo>> GetProjectsAsync();

    Task<List<DocumentInfo>> GetOpenDocumentsAsync();
    Task<DocumentInfo?> GetActiveDocumentAsync();
    Task<bool> OpenDocumentAsync(string path);
    Task<bool> CloseDocumentAsync(string path, bool save = true);
    Task<string?> ReadDocumentAsync(string path);
    Task<bool> WriteDocumentAsync(string path, string content);
    Task<SelectionInfo?> GetSelectionAsync();
    Task<bool> SetSelectionAsync(string path, int startLine, int startColumn, int endLine, int endColumn);

    Task<bool> InsertTextAsync(string text);
    Task<bool> ReplaceTextAsync(string oldText, string newText);
    Task<bool> GoToLineAsync(int line);
    Task<List<FindResult>> FindAsync(string searchText, bool matchCase = false, bool wholeWord = false);

    Task<bool> BuildSolutionAsync();
    Task<bool> BuildProjectAsync(string projectName);
    Task<bool> CleanSolutionAsync();
    Task<bool> CancelBuildAsync();
    Task<BuildStatus> GetBuildStatusAsync();
}
