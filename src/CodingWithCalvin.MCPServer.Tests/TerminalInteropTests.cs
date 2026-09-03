using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading;
using System.Threading.Tasks;
using CodingWithCalvin.MCPServer.Services;
using Microsoft.ServiceHub.Framework;
using Microsoft.VisualStudio.Terminal;
using Xunit;

namespace CodingWithCalvin.MCPServer.Tests;

/// <summary>
/// Covers the reflection bridge to the Visual Studio integrated terminal.
/// </summary>
/// <remarks>
/// <para>
/// <c>Microsoft.VisualStudio.Terminal.dll</c> cannot be referenced at build time, so
/// <see cref="TerminalInterop"/> loads it by strong name and reaches <c>ITerminalService</c>,
/// <c>ProfileConfig</c> and <c>TerminalServiceDescriptors</c> reflectively. These tests supply
/// the test assembly as the resolver and let the whole path run against stand-ins declared in
/// the genuine namespace, exercising type resolution, descriptor retrieval, the generic
/// <c>GetProxyAsync&lt;T&gt;</c> invocation with its ValueTask unwrapping, overload selection
/// and profile construction.
/// </para>
/// <para>
/// All tests live in one class deliberately: several manipulate the VSAPPIDDIR environment
/// variable, which is process-global, and xUnit runs tests within a class sequentially.
/// </para>
/// </remarks>
public class TerminalInteropTests
{
    private static readonly Func<Assembly?> StandInAssembly =
        () => typeof(ITerminalService).Assembly;

    [Fact]
    public async Task CreateTerminal_WithCommand_LaunchesThroughCommandProcessor()
    {
        var service = new StubTerminalService();
        var interop = Build(service);

        var outcome = await interop.CreateTerminalAsync(
            "MCP",
            @"C:\src\app",
            "dotnet-coverage collect dotnet test",
            CancellationToken.None);

        Assert.True(outcome.Created);
        Assert.Equal(service.NextTerminalId, outcome.TerminalId);

        var profile = Assert.IsType<ProfileConfig>(service.LastProfile);
        Assert.EndsWith("cmd.exe", profile.Location, StringComparison.OrdinalIgnoreCase);
        Assert.StartsWith("/k", profile.Arguments);
        Assert.Contains("dotnet-coverage collect dotnet test", profile.Arguments);
        Assert.True(profile.CreatePTY);
    }

    [Fact]
    public async Task CreateTerminal_SelectsOverloadCarryingWorkingDirectory()
    {
        var service = new StubTerminalService();
        var interop = Build(service);

        await interop.CreateTerminalAsync("MCP", @"C:\src\app", "dir", CancellationToken.None);

        // The three-argument overload would silently drop the working directory.
        Assert.Equal(@"C:\src\app", service.LastWorkingDirectory);
        Assert.Equal("MCP", service.LastName);
    }

    [Fact]
    public async Task CreateTerminal_WithoutCommand_LeavesProfileToVisualStudio()
    {
        var service = new StubTerminalService();
        var interop = Build(service);

        var outcome = await interop.CreateTerminalAsync("MCP", null, null, CancellationToken.None);

        Assert.True(outcome.Created);
        Assert.Null(service.LastProfile);
        Assert.False(outcome.DeveloperEnvironment);
    }

    [Fact]
    public async Task CreateTerminal_ChainsDeveloperCommandScript_WhenAvailable()
    {
        using var vsDir = new FakeVsInstall();
        var service = new StubTerminalService();
        var interop = Build(service);

        var outcome = await interop.CreateTerminalAsync(null, null, "msbuild", CancellationToken.None);

        var profile = Assert.IsType<ProfileConfig>(service.LastProfile);
        Assert.True(outcome.DeveloperEnvironment);
        Assert.Contains("VsDevCmd.bat", profile.Arguments);
        Assert.Contains("&& msbuild", profile.Arguments);
    }

    [Fact]
    public async Task CreateTerminal_ReportsUnavailable_WhenTerminalAssemblyMissing()
    {
        var interop = new TerminalInterop(() => new StubServiceBroker(new StubTerminalService()), () => null);

        var outcome = await interop.CreateTerminalAsync(null, null, "dir", CancellationToken.None);

        Assert.False(outcome.Created);
        Assert.Equal(TerminalInterop.UnavailableMessage, outcome.Message);
    }

    [Fact]
    public async Task CreateTerminal_ReportsUnavailable_WhenBrokerMissing()
    {
        var interop = new TerminalInterop(() => null, StandInAssembly);

        var outcome = await interop.CreateTerminalAsync(null, null, "dir", CancellationToken.None);

        Assert.False(outcome.Created);
        Assert.Equal(TerminalInterop.UnavailableMessage, outcome.Message);
    }

    [Fact]
    public async Task GetTerminalIds_ReturnsOpenTerminals()
    {
        var service = new StubTerminalService();
        service.OpenTerminals.Add(Guid.NewGuid());
        service.OpenTerminals.Add(Guid.NewGuid());

        var ids = await Build(service).GetTerminalIdsAsync(CancellationToken.None);

        Assert.NotNull(ids);
        Assert.Equal(service.OpenTerminals, ids!);
    }

    [Fact]
    public async Task GetTerminalIds_ReturnsNull_WhenUnavailable()
    {
        var interop = new TerminalInterop(() => null, StandInAssembly);

        // Null is distinct from an empty list, which would mean "no terminals are open".
        Assert.Null(await interop.GetTerminalIdsAsync(CancellationToken.None));
    }

    [Fact]
    public async Task ShowTerminal_ForwardsIdentifier()
    {
        var service = new StubTerminalService();
        var id = Guid.NewGuid();

        Assert.True(await Build(service).ShowTerminalAsync(id, CancellationToken.None));
        Assert.Equal(id, service.LastShown);
    }

    [Fact]
    public async Task CloseTerminal_ForwardsIdentifier()
    {
        var service = new StubTerminalService();
        var id = Guid.NewGuid();

        Assert.True(await Build(service).CloseTerminalAsync(id, CancellationToken.None));
        Assert.Equal(id, service.LastClosed);
    }

    [Fact]
    public async Task CloseAllTerminals_InvokesService()
    {
        var service = new StubTerminalService();

        Assert.True(await Build(service).CloseAllTerminalsAsync(CancellationToken.None));
        Assert.Equal(1, service.CloseAllCount);
    }

    [Fact]
    public void TryFindDeveloperCommandScript_ReturnsNull_WhenIdeDirectoryUnknown()
    {
        var previous = Environment.GetEnvironmentVariable("VSAPPIDDIR");
        try
        {
            Environment.SetEnvironmentVariable("VSAPPIDDIR", null);
            Assert.Null(TerminalInterop.TryFindDeveloperCommandScript());
        }
        finally
        {
            Environment.SetEnvironmentVariable("VSAPPIDDIR", previous);
        }
    }

    [Fact]
    public void TryFindDeveloperCommandScript_ReturnsNull_WhenScriptAbsent()
    {
        var previous = Environment.GetEnvironmentVariable("VSAPPIDDIR");
        try
        {
            Environment.SetEnvironmentVariable("VSAPPIDDIR", Path.GetTempPath());
            Assert.Null(TerminalInterop.TryFindDeveloperCommandScript());
        }
        finally
        {
            Environment.SetEnvironmentVariable("VSAPPIDDIR", previous);
        }
    }

    [Fact]
    public void TryFindDeveloperCommandScript_LocatesScriptAlongsideIdeDirectory()
    {
        using var vsInstall = new FakeVsInstall();

        var script = TerminalInterop.TryFindDeveloperCommandScript();

        Assert.NotNull(script);
        Assert.True(File.Exists(script));
        Assert.EndsWith(@"Tools\VsDevCmd.bat", script, StringComparison.OrdinalIgnoreCase);
    }

    private static TerminalInterop Build(ITerminalService service) =>
        new(() => new StubServiceBroker(service), StandInAssembly);

    /// <summary>
    /// Lays out Common7\IDE and Common7\Tools\VsDevCmd.bat in a temp directory and points
    /// VSAPPIDDIR at the IDE folder, mirroring how devenv sets it.
    /// </summary>
    private sealed class FakeVsInstall : IDisposable
    {
        private readonly string _root;
        private readonly string? _previous;

        internal FakeVsInstall()
        {
            _root = Path.Combine(Path.GetTempPath(), "vsmcp-" + Guid.NewGuid().ToString("N"));
            var ide = Path.Combine(_root, "Common7", "IDE");
            var tools = Path.Combine(_root, "Common7", "Tools");

            Directory.CreateDirectory(ide);
            Directory.CreateDirectory(tools);
            File.WriteAllText(Path.Combine(tools, "VsDevCmd.bat"), "@echo off");

            _previous = Environment.GetEnvironmentVariable("VSAPPIDDIR");
            Environment.SetEnvironmentVariable("VSAPPIDDIR", ide);
        }

        public void Dispose()
        {
            Environment.SetEnvironmentVariable("VSAPPIDDIR", _previous);

            try
            {
                Directory.Delete(_root, recursive: true);
            }
            catch (IOException)
            {
                // A leftover temp directory is not worth failing a test over.
            }
        }
    }

    private sealed class StubTerminalService : ITerminalService
    {
        internal Guid NextTerminalId { get; } = Guid.NewGuid();

        internal List<Guid> OpenTerminals { get; } = new();

        internal string? LastName { get; private set; }

        internal ProfileConfig? LastProfile { get; private set; }

        internal string? LastWorkingDirectory { get; private set; }

        internal Guid? LastShown { get; private set; }

        internal Guid? LastClosed { get; private set; }

        internal int CloseAllCount { get; private set; }

        public Task<Guid> CreateTerminalAsync(CancellationToken cancellationToken, string? name, ProfileConfig? profile)
            => throw new InvalidOperationException(
                "The overload without a working directory should not be selected.");

        public Task<Guid> CreateTerminalAsync(
            CancellationToken cancellationToken,
            string? name,
            ProfileConfig? profile,
            string? workingDirectory)
        {
            LastName = name;
            LastProfile = profile;
            LastWorkingDirectory = workingDirectory;
            return Task.FromResult(NextTerminalId);
        }

        public Task<IEnumerable<Guid>> GetTerminalGuidsAsync(CancellationToken cancellationToken)
            => Task.FromResult<IEnumerable<Guid>>(OpenTerminals.ToList());

        public Task ShowAsync(Guid terminalGuid, CancellationToken cancellationToken)
        {
            LastShown = terminalGuid;
            return Task.CompletedTask;
        }

        public Task CloseAsync(Guid terminalGuid, CancellationToken cancellationToken)
        {
            LastClosed = terminalGuid;
            return Task.CompletedTask;
        }

        public Task CloseAllInstancesAsync(CancellationToken cancellationToken)
        {
            CloseAllCount++;
            return Task.CompletedTask;
        }
    }

    private sealed class StubServiceBroker : IServiceBroker
    {
        private readonly object _service;

        internal StubServiceBroker(object service)
        {
            _service = service;
        }

#pragma warning disable CS0067 // Required by the interface; nothing under test raises it.
        public event EventHandler<BrokeredServicesChangedEventArgs>? AvailabilityChanged;
#pragma warning restore CS0067

        public ValueTask<T?> GetProxyAsync<T>(
            ServiceRpcDescriptor serviceDescriptor,
            ServiceActivationOptions options = default,
            CancellationToken cancellationToken = default)
            where T : class
        {
            Assert.NotNull(serviceDescriptor);
            return new ValueTask<T?>((T?)_service);
        }

        public ValueTask<System.IO.Pipelines.IDuplexPipe?> GetPipeAsync(
            ServiceMoniker serviceMoniker,
            ServiceActivationOptions options = default,
            CancellationToken cancellationToken = default)
            => throw new NotSupportedException();
    }
}
