using System;
using System.Collections.Generic;
using System.ComponentModel.Composition;
using System.IO.Pipes;
using System.Threading;
using System.Threading.Tasks;
using CodingWithCalvin.MCPServer.Shared;
using CodingWithCalvin.MCPServer.Shared.Models;
using StreamJsonRpc;

namespace CodingWithCalvin.MCPServer.Services;

[Export(typeof(IRpcServer))]
[PartCreationPolicy(CreationPolicy.Shared)]
public class RpcServer : IRpcServer, IVisualStudioRpc
{
    private readonly IVisualStudioService _vsService;
    private NamedPipeServerStream? _pipeServer;
    private JsonRpc? _jsonRpc;
    private IServerRpc? _serverProxy;
    private CancellationTokenSource? _cts;
    private Task? _listenerTask;
    private bool _disposed;

    public string PipeName { get; private set; } = string.Empty;
    public bool IsListening { get; private set; }
    public bool IsConnected => _serverProxy != null;

    [ImportingConstructor]
    public RpcServer(IVisualStudioService vsService)
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
                _serverProxy = _jsonRpc.Attach<IServerRpc>();
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
                _serverProxy = null;
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

        // Dispose JsonRpc to break out of the Completion await
        _jsonRpc?.Dispose();
        _pipeServer?.Dispose();

        if (_listenerTask != null)
        {
            try
            {
                // Use a timeout to prevent hanging forever
                var timeoutTask = Task.Delay(2000);
                var completedTask = await Task.WhenAny(_listenerTask, timeoutTask);
                if (completedTask == timeoutTask)
                {
                    // Listener didn't stop in time, just continue
                }
            }
            catch (OperationCanceledException)
            {
                // Expected
            }
            catch
            {
                // Ignore other exceptions during shutdown
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

    public async Task<List<ToolInfo>> GetAvailableToolsAsync()
    {
        if (_serverProxy == null)
        {
            return new List<ToolInfo>();
        }

        return await _serverProxy.GetAvailableToolsAsync();
    }

    public async Task RequestShutdownAsync()
    {
        if (_serverProxy != null)
        {
            try
            {
                await _serverProxy.ShutdownAsync();
            }
            catch
            {
            }
        }
    }

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
    public Task<bool> CancelBuildAsync() => _vsService.CancelBuildAsync();
    public Task<BuildStatus> GetBuildStatusAsync() => _vsService.GetBuildStatusAsync();
}
