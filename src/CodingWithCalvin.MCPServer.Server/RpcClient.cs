using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO.Pipes;
using System.Linq;
using System.Reflection;
using System.Threading;
using System.Threading.Tasks;
using CodingWithCalvin.MCPServer.Shared;
using CodingWithCalvin.MCPServer.Shared.Models;
using ModelContextProtocol.Server;
using StreamJsonRpc;

namespace CodingWithCalvin.MCPServer.Server;

public class RpcClient : IVisualStudioRpc, IServerRpc, IDisposable
{
    private readonly CancellationTokenSource _shutdownCts;
    private NamedPipeClientStream? _pipeClient;
    private JsonRpc? _jsonRpc;
    private IVisualStudioRpc? _proxy;
    private bool _disposed;
    private List<ToolInfo>? _cachedTools;

    public bool IsConnected => _pipeClient?.IsConnected ?? false;

    public RpcClient(CancellationTokenSource shutdownCts)
    {
        _shutdownCts = shutdownCts;
    }

    public async Task ConnectAsync(string pipeName, int timeoutMs = 10000)
    {
        _pipeClient = new NamedPipeClientStream(
            ".",
            pipeName,
            PipeDirection.InOut,
            PipeOptions.Asynchronous);

        await _pipeClient.ConnectAsync(timeoutMs);

        _jsonRpc = JsonRpc.Attach(_pipeClient, this);
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

    public Task<List<ToolInfo>> GetAvailableToolsAsync()
    {
        if (_cachedTools != null)
        {
            return Task.FromResult(_cachedTools);
        }

        var tools = new List<ToolInfo>();
        var toolTypes = new[] { typeof(Tools.SolutionTools), typeof(Tools.DocumentTools), typeof(Tools.BuildTools), typeof(Tools.NavigationTools) };

        foreach (var toolType in toolTypes)
        {
            var category = toolType.Name.Replace("Tools", "");

            foreach (var method in toolType.GetMethods(BindingFlags.Public | BindingFlags.Instance))
            {
                var toolAttr = method.GetCustomAttribute<McpServerToolAttribute>();
                if (toolAttr == null)
                {
                    continue;
                }

                var descAttr = method.GetCustomAttribute<DescriptionAttribute>();

                tools.Add(new ToolInfo
                {
                    Name = toolAttr.Name ?? method.Name,
                    Description = descAttr?.Description ?? string.Empty,
                    Category = category
                });
            }
        }

        _cachedTools = tools;
        return Task.FromResult(tools);
    }

    public Task ShutdownAsync()
    {
        Console.Error.WriteLine("Shutdown requested via RPC");
        _shutdownCts.Cancel();
        return Task.CompletedTask;
    }

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
    public Task<bool> CancelBuildAsync() => Proxy.CancelBuildAsync();
    public Task<BuildStatus> GetBuildStatusAsync() => Proxy.GetBuildStatusAsync();

    public Task<List<SymbolInfo>> GetDocumentSymbolsAsync(string path) => Proxy.GetDocumentSymbolsAsync(path);
    public Task<WorkspaceSymbolResult> SearchWorkspaceSymbolsAsync(string query, int maxResults = 100)
        => Proxy.SearchWorkspaceSymbolsAsync(query, maxResults);
    public Task<DefinitionResult> GoToDefinitionAsync(string path, int line, int column)
        => Proxy.GoToDefinitionAsync(path, line, column);
    public Task<ReferencesResult> FindReferencesAsync(string path, int line, int column, int maxResults = 100)
        => Proxy.FindReferencesAsync(path, line, column, maxResults);
}
