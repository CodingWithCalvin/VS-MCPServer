using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Threading.Tasks;
using CodingWithCalvin.VSMCP.Options;
using Microsoft.VisualStudio.Shell;

namespace CodingWithCalvin.VSMCP.Services;

public class ServerProcessManager
{
    private readonly AsyncPackage _package;
    private readonly RpcServer _rpcServer;
    private Process? _serverProcess;
    private string _pipeName = string.Empty;

    public bool IsRunning => _serverProcess != null && !_serverProcess.HasExited;

    public ServerProcessManager(AsyncPackage package, RpcServer rpcServer)
    {
        _package = package;
        _rpcServer = rpcServer;
    }

    public async Task StartAsync()
    {
        if (IsRunning)
        {
            return;
        }

        var options = await GeneralOptions.GetLiveInstanceAsync();

        // Generate unique pipe name for this VS instance
        _pipeName = $"vsmcp-{Process.GetCurrentProcess().Id}";

        // Start the RPC server first
        await _rpcServer.StartAsync(_pipeName);

        // Find the server executable
        var extensionDir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
        var serverExe = Path.Combine(extensionDir!, "Server", "CodingWithCalvin.VSMCP.Server.exe");

        if (!File.Exists(serverExe))
        {
            throw new FileNotFoundException("MCP Server executable not found", serverExe);
        }

        // Start the server process
        var startInfo = new ProcessStartInfo
        {
            FileName = serverExe,
            Arguments = $"--pipe \"{_pipeName}\" --port {options.HttpPort} --name \"{options.ServerName}\"",
            UseShellExecute = false,
            CreateNoWindow = true,
            RedirectStandardOutput = true,
            RedirectStandardError = true
        };

        _serverProcess = Process.Start(startInfo);

        if (_serverProcess == null)
        {
            throw new InvalidOperationException("Failed to start MCP Server process");
        }

        _serverProcess.EnableRaisingEvents = true;
        _serverProcess.Exited += OnProcessExited;

        // Give the server a moment to start
        await Task.Delay(500);

        if (_serverProcess.HasExited)
        {
            var error = await _serverProcess.StandardError.ReadToEndAsync();
            throw new InvalidOperationException($"MCP Server process exited immediately: {error}");
        }
    }

    public async Task StopAsync()
    {
        if (_serverProcess != null && !_serverProcess.HasExited)
        {
            try
            {
                _serverProcess.Kill();
                await Task.Run(() => _serverProcess.WaitForExit(5000));
            }
            catch
            {
                // Process may have already exited
            }
        }

        _serverProcess?.Dispose();
        _serverProcess = null;

        await _rpcServer.StopAsync();
    }

    private void OnProcessExited(object? sender, EventArgs e)
    {
        // Server process exited unexpectedly
        _serverProcess?.Dispose();
        _serverProcess = null;
    }
}
