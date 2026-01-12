using System;
using System.Collections.Generic;
using System.IO.Pipes;
using System.Threading;
using System.Threading.Tasks;
using CodingWithCalvin.VSMCP.Shared;
using CodingWithCalvin.VSMCP.Shared.Models;
using StreamJsonRpc;

namespace CodingWithCalvin.VSMCP.Services;

public class RpcServer : IVisualStudioRpc, IDisposable
{
    private readonly VisualStudioService _vsService;
    private NamedPipeServerStream? _pipeServer;
    private JsonRpc? _jsonRpc;
    private CancellationTokenSource? _cts;
    private Task? _listenerTask;
    private bool _disposed;

    public string PipeName { get; private set; } = string.Empty;
    public bool IsListening { get; private set; }

    public RpcServer(VisualStudioService vsService)
    {
        _vsService = vsService;
    }

    public async Task StartAsync(string pipeName)
    {
        if (IsListening)
        {
            return;
        }

        PipeName = pipeName;
        _cts = new CancellationTokenSource();
        IsListening = true;

        _listenerTask = Task.Run(() => ListenAsync(_cts.Token));
        await Task.CompletedTask;
    }

    private async Task ListenAsync(CancellationToken cancellationToken)
    {
        while (!cancellationToken.IsCancellationRequested)
        {
            try
            {
                _pipeServer = new NamedPipeServerStream(
                    PipeName,
                    PipeDirection.InOut,
                    1,
                    PipeTransmissionMode.Byte,
                    PipeOptions.Asynchronous);

                await _pipeServer.WaitForConnectionAsync(cancellationToken);

                _jsonRpc = JsonRpc.Attach(_pipeServer, this);
                await _jsonRpc.Completion;
            }
            catch (OperationCanceledException)
            {
                break;
            }
            catch (Exception)
            {
                // Connection lost, restart listening
                await Task.Delay(100, cancellationToken);
            }
            finally
            {
                _jsonRpc?.Dispose();
                _jsonRpc = null;
                _pipeServer?.Dispose();
                _pipeServer = null;
            }
        }
    }

    public async Task StopAsync()
    {
        if (!IsListening)
        {
            return;
        }

        IsListening = false;
        _cts?.Cancel();

        if (_listenerTask != null)
        {
            try
            {
                await _listenerTask;
            }
            catch (OperationCanceledException)
            {
                // Expected
            }
        }

        _cts?.Dispose();
        _cts = null;
    }

    public void Dispose()
    {
        if (_disposed)
        {
            return;
        }

        _disposed = true;
        StopAsync().GetAwaiter().GetResult();
    }

    #region IVisualStudioRpc Implementation

    public Task<SolutionInfo?> GetSolutionInfoAsync() => _vsService.GetSolutionInfoAsync();
    public Task<bool> OpenSolutionAsync(string path) => _vsService.OpenSolutionAsync(path);
    public Task CloseSolutionAsync(bool saveFirst) => _vsService.CloseSolutionAsync(saveFirst);
    public Task<List<ProjectInfo>> GetProjectsAsync() => _vsService.GetProjectsAsync();
    public Task<List<DocumentInfo>> GetOpenDocumentsAsync() => _vsService.GetOpenDocumentsAsync();
    public Task<DocumentInfo?> GetActiveDocumentAsync() => _vsService.GetActiveDocumentAsync();
    public Task<bool> OpenDocumentAsync(string path) => _vsService.OpenDocumentAsync(path);
    public Task<bool> CloseDocumentAsync(string path, bool save) => _vsService.CloseDocumentAsync(path, save);
    public Task<string?> ReadDocumentAsync(string path) => _vsService.ReadDocumentAsync(path);
    public Task<bool> WriteDocumentAsync(string path, string content) => _vsService.WriteDocumentAsync(path, content);
    public Task<SelectionInfo?> GetSelectionAsync() => _vsService.GetSelectionAsync();
    public Task<bool> SetSelectionAsync(string path, int startLine, int startColumn, int endLine, int endColumn)
        => _vsService.SetSelectionAsync(path, startLine, startColumn, endLine, endColumn);
    public Task<bool> InsertTextAsync(string text) => _vsService.InsertTextAsync(text);
    public Task<bool> ReplaceTextAsync(string oldText, string newText) => _vsService.ReplaceTextAsync(oldText, newText);
    public Task<bool> GoToLineAsync(int line) => _vsService.GoToLineAsync(line);
    public Task<List<FindResult>> FindAsync(string searchText, bool matchCase, bool wholeWord)
        => _vsService.FindAsync(searchText, matchCase, wholeWord);
    public Task<bool> BuildSolutionAsync() => _vsService.BuildSolutionAsync();
    public Task<bool> BuildProjectAsync(string projectName) => _vsService.BuildProjectAsync(projectName);
    public Task<bool> CleanSolutionAsync() => _vsService.CleanSolutionAsync();
    public Task CancelBuildAsync() => _vsService.CancelBuildAsync();
    public Task<BuildStatus> GetBuildStatusAsync() => _vsService.GetBuildStatusAsync();

    #endregion
}
