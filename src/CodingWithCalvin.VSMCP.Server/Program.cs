using System;
using System.CommandLine;
using System.Threading.Tasks;
using CodingWithCalvin.VSMCP.Server;
using CodingWithCalvin.VSMCP.Server.Tools;
using Microsoft.AspNetCore.Builder;
using Microsoft.Extensions.DependencyInjection;
using ModelContextProtocol.AspNetCore;
using ModelContextProtocol.Protocol;
using ModelContextProtocol.Server;

var pipeOption = new Option<string>(
    name: "--pipe",
    description: "Named pipe name for connecting to Visual Studio")
{
    IsRequired = true
};

var portOption = new Option<int>(
    name: "--port",
    getDefaultValue: () => 5050,
    description: "HTTP port for the MCP server");

var nameOption = new Option<string>(
    name: "--name",
    getDefaultValue: () => "Visual Studio MCP",
    description: "Server name displayed to MCP clients");

var rootCommand = new RootCommand("Visual Studio MCP Server")
{
    pipeOption,
    portOption,
    nameOption
};

rootCommand.SetHandler(async (string pipeName, int port, string serverName) =>
{
    await RunServerAsync(pipeName, port, serverName);
}, pipeOption, portOption, nameOption);

return await rootCommand.InvokeAsync(args);

static async Task RunServerAsync(string pipeName, int port, string serverName)
{
    // Connect to Visual Studio via named pipe
    var rpcClient = new RpcClient();
    await rpcClient.ConnectAsync(pipeName);

    Console.Error.WriteLine($"Connected to Visual Studio via pipe: {pipeName}");

    // Build the web application
    var builder = WebApplication.CreateBuilder();

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

    Console.Error.WriteLine($"MCP Server listening on http://localhost:{port}");

    await app.RunAsync($"http://localhost:{port}");
}
