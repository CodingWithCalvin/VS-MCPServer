using System;
using System.ComponentModel.Composition;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Threading.Tasks;
using CodingWithCalvin.MCPServer.Dialogs;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.VisualStudio.Threading;

namespace CodingWithCalvin.MCPServer.Services;

[Export(typeof(IServerProcessManager))]
[PartCreationPolicy(CreationPolicy.Shared)]
public class ServerProcessManager : IServerProcessManager
{
    /// <summary>How long to wait for the server to acknowledge the RPC shutdown request.</summary>
    private const int RpcShutdownTimeoutMs = 1000;

    /// <summary>How long to wait for the server process to exit on its own afterwards.</summary>
    private const int GracefulExitTimeoutMs = 1500;

    /// <summary>How long to wait for the process to die after being killed.</summary>
    private const int ForcedExitTimeoutMs = 500;

    private readonly IRpcServer _rpcServer;
    private Process? _serverProcess;
    private ProcessJobObject? _jobObject;
    private string _pipeName = string.Empty;
    private StreamWriter? _logFileWriter;
    private string? _logFilePath;
    private IVsOutputWindowPane? _outputPane;

    public bool IsRunning => _serverProcess != null && !_serverProcess.HasExited;
    public string? LogFilePath => _logFilePath;

    [ImportingConstructor]
    public ServerProcessManager(IRpcServer rpcServer)
    {
        _rpcServer = rpcServer;
    }

    public async Task StartAsync(ServerStartSettings settings)
    {
        if (IsRunning)
        {
            return;
        }

        // Initialize logging (file + output pane from UI thread)
        InitializeLogging(settings);

        // Generate unique pipe name for this VS instance
        _pipeName = $"vsmcp-{Process.GetCurrentProcess().Id}";

        // Start the RPC server first
        await _rpcServer.StartAsync(_pipeName).ConfigureAwait(false);

        // Find the server executable
        var extensionDir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
        var serverExe = Path.Combine(extensionDir!, "Server", "CodingWithCalvin.MCPServer.Server.exe");

        if (!File.Exists(serverExe))
        {
            throw new FileNotFoundException("MCP Server executable not found", serverExe);
        }

        // Start the server process
        var arguments = $"--pipe \"{_pipeName}\" --host \"{settings.BindingAddress}\" --port {settings.Port} --name \"{settings.ServerName}\" --log-level {settings.LogLevel}";

        var startInfo = new ProcessStartInfo
        {
            FileName = serverExe,
            Arguments = arguments,
            UseShellExecute = false,
            CreateNoWindow = true,
            RedirectStandardOutput = true,
            RedirectStandardError = true
        };

        var process = Process.Start(startInfo);

        if (process == null)
        {
            throw new InvalidOperationException("Failed to start MCP Server process");
        }

        // Store in field and capture local reference to avoid race conditions
        _serverProcess = process;

        process.EnableRaisingEvents = true;
        process.Exited += OnProcessExited;

        // Tie the server's lifetime to this process. Cooperative shutdown below is the
        // normal path; the job object is what saves us when devenv.exe dies without it.
        AssignToJobObject(process);

        // Start reading output streams (server logs go to stderr by convention)
        _ = ReadOutputAsync(process.StandardOutput);
        _ = ReadOutputAsync(process.StandardError);

        Log($"Server started (PID: {process.Id})");
        Log($"Binding: http://{settings.BindingAddress}:{settings.Port}");
        Log($"Log file: {_logFilePath}");

        // Give the server a moment to start
        await Task.Delay(500).ConfigureAwait(false);

        // Check if process exited
        if (process.HasExited)
        {
            throw new InvalidOperationException($"MCP Server process exited immediately with code {process.ExitCode}");
        }
    }

    /// <remarks>
    /// Every await in this method — and everything it calls — must use
    /// <c>ConfigureAwait(false)</c>. Package disposal blocks a thread waiting on this, so a
    /// continuation that needs the caller's synchronization context back would deadlock and
    /// leave devenv.exe running forever (issue #97). The total wall time is also bounded, so
    /// an unresponsive server cannot stall Visual Studio's shutdown.
    /// </remarks>
    public async Task StopAsync()
    {
        // Capture reference to avoid race conditions during shutdown
        var process = _serverProcess;
        _serverProcess = null;

        if (process != null)
        {
            try
            {
                if (!process.HasExited)
                {
                    Log("Stopping server...");

                    // Unsubscribe from Exited event to prevent duplicate logging
                    process.Exited -= OnProcessExited;

                    await RequestGracefulShutdownAsync(process).ConfigureAwait(false);
                }

                if (!process.HasExited)
                {
                    // Force kill if graceful shutdown timed out
                    Log("Graceful shutdown timed out, forcing termination...");
                    process.Kill();
                    await Task.Run(() => process.WaitForExit(ForcedExitTimeoutMs)).ConfigureAwait(false);
                }

                Log($"Server stopped (Code: {process.ExitCode})");
            }
            catch (Exception ex)
            {
                // Never allow a shutdown failure to propagate into package disposal.
                Log($"Error stopping server: {ex.Message}");
            }
            finally
            {
                process.Dispose();
            }
        }

        // Closing the job terminates the server process if it somehow outlived the above.
        _jobObject?.Dispose();
        _jobObject = null;

        await _rpcServer.StopAsync().ConfigureAwait(false);

        // Close log file
        _logFileWriter?.Dispose();
        _logFileWriter = null;
    }

    /// <summary>
    /// Asks the server to shut down over RPC and waits briefly for it to exit. Both steps are
    /// individually bounded because a half-open named pipe can leave an RPC call pending
    /// indefinitely.
    /// </summary>
    private async Task RequestGracefulShutdownAsync(Process process)
    {
        // RequestShutdownAsync swallows its own errors, so the abandoned task on timeout is
        // harmless and cannot surface as an unobserved exception.
        var shutdownRequest = _rpcServer.RequestShutdownAsync();
        await Task.WhenAny(shutdownRequest, Task.Delay(RpcShutdownTimeoutMs)).ConfigureAwait(false);

        await Task.Run(() => process.WaitForExit(GracefulExitTimeoutMs)).ConfigureAwait(false);
    }

    private void AssignToJobObject(Process process)
    {
        try
        {
            _jobObject ??= ProcessJobObject.Create();

            if (_jobObject == null || !_jobObject.TryAssign(process))
            {
                Log("Warning: could not assign the server to a job object; it may outlive Visual Studio if devenv.exe terminates abnormally.");
            }
        }
        catch (Exception ex)
        {
            Log($"Warning: job object setup failed ({ex.Message}); the server may outlive Visual Studio if devenv.exe terminates abnormally.");
        }
    }

    private void InitializeLogging(ServerStartSettings settings)
    {
        // Create log file in temp directory (daily rotation)
        try
        {
            var logDir = SettingsDialogPage.LogFolderPath;
            Directory.CreateDirectory(logDir);

            var date = DateTime.Now.ToString("yyyy-MM-dd");
            _logFilePath = Path.Combine(logDir, $"server_{date}.log");
            _logFileWriter = new StreamWriter(_logFilePath, append: true) { AutoFlush = true };

            // Clean up old log files (fire and forget to not block startup)
            var retentionDays = settings.LogRetentionDays;
            Task.Run(() => CleanupOldLogFiles(logDir, retentionDays));
        }
        catch
        {
            // File logging will be unavailable, but continue anyway
            _logFilePath = null;
            _logFileWriter = null;
        }

        // Use output pane passed from UI thread
        _outputPane = settings.OutputPane;

        Log($"=== MCP Server Log Started {DateTime.Now:yyyy-MM-dd HH:mm:ss} ===");
    }

    private void CleanupOldLogFiles(string logDir, int retentionDays)
    {
        if (retentionDays <= 0)
        {
            return; // Keep all logs
        }

        try
        {
            var cutoffDate = DateTime.Now.AddDays(-retentionDays);
            var logFiles = Directory.GetFiles(logDir, "server_*.log");

            foreach (var file in logFiles)
            {
                var fileInfo = new FileInfo(file);
                if (fileInfo.LastWriteTime < cutoffDate)
                {
                    try
                    {
                        fileInfo.Delete();
                    }
                    catch
                    {
                        // Ignore individual file deletion errors
                    }
                }
            }
        }
        catch
        {
            // Ignore cleanup errors
        }
    }

    private async Task ReadOutputAsync(StreamReader reader)
    {
        try
        {
            while (!reader.EndOfStream)
            {
                var line = await reader.ReadLineAsync().ConfigureAwait(false);
                if (!string.IsNullOrEmpty(line))
                {
                    Log($"[SERVER] {line}");
                }
            }
        }
        catch
        {
            // Stream closed, ignore
        }
    }

    private void Log(string message)
    {
        var timestamp = DateTime.Now.ToString("HH:mm:ss.fff");
        var logLine = $"[{timestamp}] {message}";

        // Write to file
        try
        {
            _logFileWriter?.WriteLine(logLine);
        }
        catch
        {
            // Ignore file write errors
        }

        // Write to output pane (OutputStringThreadSafe is thread-safe, no main thread needed)
        try
        {
            _outputPane?.OutputStringThreadSafe(logLine + Environment.NewLine);
        }
        catch
        {
            // Ignore output pane errors
        }
    }

    private void OnProcessExited(object? sender, EventArgs e)
    {
        // Use sender to get exit code - don't modify _serverProcess here to avoid race conditions
        var exitCode = (sender as Process)?.ExitCode;
        Log($"Server process exited (Code: {exitCode})");
    }
}
