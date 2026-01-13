using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using CodingWithCalvin.MCPServer.Shared.Models;

namespace CodingWithCalvin.MCPServer.Services;

public interface IRpcServer : IDisposable
{
    string PipeName { get; }
    bool IsListening { get; }
    bool IsConnected { get; }

    Task StartAsync(string pipeName);
    Task StopAsync();
    Task<List<ToolInfo>> GetAvailableToolsAsync();
    Task RequestShutdownAsync();
}
