using System;
using System.CommandLine;
using System.Threading;
using System.Threading.Tasks;
using CodingWithCalvin.MCPServer.Server;
using CodingWithCalvin.MCPServer.Server.Tools;
using Microsoft.AspNetCore.Builder;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using ModelContextProtocol.AspNetCore;
using ModelContextProtocol.Protocol;
using ModelContextProtocol.Server;

var pipeOption = new Option<string>(
    name: "--pipe",
    description: "Named pipe name for connecting to Visual Studio")
{
    IsRequired = true
};

var hostOption = new Option<string>(
    name: "--host",
    getDefaultValue: () => "localhost",
    description: "Host address to bind the HTTP server to (e.g., localhost, 0.0.0.0, *)");

var portOption = new Option<int>(
    name: "--port",
    getDefaultValue: () => 5050,
    description: "HTTP port for the MCP server");

var nameOption = new Option<string>(
    name: "--name",
    getDefaultValue: () => "Visual Studio MCP",
    description: "Server name displayed to MCP clients");

var logLevelOption = new Option<string>(
    name: "--log-level",
    getDefaultValue: () => "Information",
    description: "Minimum log level (Error, Warning, Information, Debug)");

var rootCommand = new RootCommand("Visual Studio MCP Server")
{
    pipeOption,
    hostOption,
    portOption,
    nameOption,
    logLevelOption
};

rootCommand.SetHandler(async (string pipeName, string host, int port, string serverName, string logLevel) =>
{
    await RunServerAsync(pipeName, host, port, serverName, logLevel);
}, pipeOption, hostOption, portOption, nameOption, logLevelOption);

return await rootCommand.InvokeAsync(args);

static async Task RunServerAsync(string pipeName, string host, int port, string serverName, string logLevel)
{
    // Parse log level
    var msLogLevel = logLevel switch
    {
        "Error" => LogLevel.Error,
        "Warning" => LogLevel.Warning,
        "Debug" => LogLevel.Debug,
        _ => LogLevel.Information
    };

    // Create shutdown token for graceful shutdown
    using var shutdownCts = new CancellationTokenSource();

    // Connect to Visual Studio via named pipe
    var rpcClient = new RpcClient(shutdownCts);
    await rpcClient.ConnectAsync(pipeName);

    Console.Error.WriteLine($"Connected to Visual Studio via pipe: {pipeName}");

    // Build the web application
    var builder = WebApplication.CreateBuilder();

    // Configure logging
    builder.Logging.SetMinimumLevel(msLogLevel);
    builder.Logging.AddFilter("Microsoft.AspNetCore", msLogLevel);
    builder.Logging.AddFilter("ModelContextProtocol", msLogLevel);

    builder.Services.AddSingleton(rpcClient);

    builder.Services.AddMcpServer(options =>
    {
        options.ServerInfo = new Implementation
        {
            Name = serverName,
            Version = "1.0.0"
        };
    })
    .WithHttpTransport()
    .WithTools<SolutionTools>()
    .WithTools<DocumentTools>()
    .WithTools<BuildTools>();

    var app = builder.Build();

    app.MapMcp();

    var bindingUrl = $"http://{host}:{port}";
    app.Urls.Add(bindingUrl);

    // Register shutdown token to stop the application
    var lifetime = app.Services.GetRequiredService<IHostApplicationLifetime>();
    shutdownCts.Token.Register(() => lifetime.StopApplication());

    Console.Error.WriteLine($"MCP Server listening on {bindingUrl} (LogLevel: {logLevel})");

    await app.RunAsync();

    Console.Error.WriteLine("Server shutdown complete");
}
