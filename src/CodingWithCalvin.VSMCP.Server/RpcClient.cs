using System;
using System.Collections.Generic;
using System.IO.Pipes;
using System.Threading.Tasks;
using CodingWithCalvin.VSMCP.Shared;
using CodingWithCalvin.VSMCP.Shared.Models;
using StreamJsonRpc;

namespace CodingWithCalvin.VSMCP.Server;

public class RpcClient : IVisualStudioRpc, IDisposable
{
    private NamedPipeClientStream? _pipeClient;
    private JsonRpc? _jsonRpc;
    private IVisualStudioRpc? _proxy;
    private bool _disposed;

    public bool IsConnected => _pipeClient?.IsConnected ?? false;

    public async Task ConnectAsync(string pipeName, int timeoutMs = 10000)
    {
        _pipeClient = new NamedPipeClientStream(
            ".",
            pipeName,
            PipeDirection.InOut,
            PipeOptions.Asynchronous);

        await _pipeClient.ConnectAsync(timeoutMs);

        _jsonRpc = JsonRpc.Attach(_pipeClient);
        _proxy = _jsonRpc.Attach<IVisualStudioRpc>();
    }

    public void Dispose()
    {
        if (_disposed)
        {
            return;
        }

        _disposed = true;
        _jsonRpc?.Dispose();
        _pipeClient?.Dispose();
    }

    private IVisualStudioRpc Proxy => _proxy ?? throw new InvalidOperationException("Not connected to Visual Studio");

    #region IVisualStudioRpc Implementation

    public Task<SolutionInfo?> GetSolutionInfoAsync() => Proxy.GetSolutionInfoAsync();
    public Task<bool> OpenSolutionAsync(string path) => Proxy.OpenSolutionAsync(path);
    public Task CloseSolutionAsync(bool saveFirst) => Proxy.CloseSolutionAsync(saveFirst);
    public Task<List<ProjectInfo>> GetProjectsAsync() => Proxy.GetProjectsAsync();
    public Task<List<DocumentInfo>> GetOpenDocumentsAsync() => Proxy.GetOpenDocumentsAsync();
    public Task<DocumentInfo?> GetActiveDocumentAsync() => Proxy.GetActiveDocumentAsync();
    public Task<bool> OpenDocumentAsync(string path) => Proxy.OpenDocumentAsync(path);
    public Task<bool> CloseDocumentAsync(string path, bool save) => Proxy.CloseDocumentAsync(path, save);
    public Task<string?> ReadDocumentAsync(string path) => Proxy.ReadDocumentAsync(path);
    public Task<bool> WriteDocumentAsync(string path, string content) => Proxy.WriteDocumentAsync(path, content);
    public Task<SelectionInfo?> GetSelectionAsync() => Proxy.GetSelectionAsync();
    public Task<bool> SetSelectionAsync(string path, int startLine, int startColumn, int endLine, int endColumn)
        => Proxy.SetSelectionAsync(path, startLine, startColumn, endLine, endColumn);
    public Task<bool> InsertTextAsync(string text) => Proxy.InsertTextAsync(text);
    public Task<bool> ReplaceTextAsync(string oldText, string newText) => Proxy.ReplaceTextAsync(oldText, newText);
    public Task<bool> GoToLineAsync(int line) => Proxy.GoToLineAsync(line);
    public Task<List<FindResult>> FindAsync(string searchText, bool matchCase, bool wholeWord)
        => Proxy.FindAsync(searchText, matchCase, wholeWord);
    public Task<bool> BuildSolutionAsync() => Proxy.BuildSolutionAsync();
    public Task<bool> BuildProjectAsync(string projectName) => Proxy.BuildProjectAsync(projectName);
    public Task<bool> CleanSolutionAsync() => Proxy.CleanSolutionAsync();
    public Task CancelBuildAsync() => Proxy.CancelBuildAsync();
    public Task<BuildStatus> GetBuildStatusAsync() => Proxy.GetBuildStatusAsync();

    #endregion
}
