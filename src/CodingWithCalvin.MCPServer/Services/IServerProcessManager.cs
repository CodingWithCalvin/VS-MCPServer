using System.Threading.Tasks;

namespace CodingWithCalvin.MCPServer.Services;

public struct ServerStartSettings
{
    public string BindingAddress;
    public int Port;
    public string ServerName;
    public string LogLevel;
    public int LogRetentionDays;
    public Microsoft.VisualStudio.Shell.Interop.IVsOutputWindowPane? OutputPane;
}

public interface IServerProcessManager
{
    bool IsRunning { get; }
    string? LogFilePath { get; }

    Task StartAsync(ServerStartSettings settings);
    Task StopAsync();
}
