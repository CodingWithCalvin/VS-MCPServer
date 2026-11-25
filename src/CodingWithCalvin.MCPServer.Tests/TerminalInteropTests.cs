using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
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
/// <see cref="TerminalInterop"/> loads it by strong name and reaches its types reflectively.
/// These tests supply filtered views of the test assembly to exercise both the VS 2022 broker
/// path and the VS 2026 direct-service path against stand-ins declared in the genuine namespace,
/// including type resolution, service acquisition, overload selection, and the input stream used
/// to type commands into an open terminal.
/// </para>
/// </remarks>
public class TerminalInteropTests
{
    private const string DeveloperShellArguments =
        "-NoLogo -NoExit -Command \"Import-Module 'C:\\VS\\Tools\\Microsoft.VisualStudio.DevShell.dll'; Enter-VsDevShell\"";

    private static readonly Func<Assembly?> BrokerStandInAssembly =
        () => new FilteredAssembly(
            typeof(ITerminalService).Assembly,
            typeof(SVsTerminalService).FullName!,
            typeof(IVsTerminalService).FullName!);

    private static readonly Func<Assembly?> DirectStandInAssembly =
        () => new FilteredAssembly(
            typeof(ITerminalService).Assembly,
            typeof(TerminalServiceDescriptors).FullName!);

    [Fact]
    public async Task RunInTerminal_WithoutTerminalId_OpensDefaultProfileTerminalAndTypesCommand()
    {
        var service = new StubTerminalService();
        var interop = Build(service);

        var outcome = await interop.RunInTerminalAsync(
            null,
            "MCP",
            @"C:\src\app",
            "dotnet-coverage collect dotnet test",
            CancellationToken.None);

        Assert.True(outcome.Created);
        Assert.False(outcome.Reused);
        Assert.Equal(service.NextTerminalId, outcome.TerminalId);
        Assert.True(outcome.DeveloperEnvironment);

        // The user's default profile is used as-is: no synthetic cmd.exe /k profile.
        Assert.Same(service.DefaultProfile, service.LastProfile);
        Assert.DoesNotContain("cmd.exe", service.DefaultProfile.Location, StringComparison.OrdinalIgnoreCase);

        var input = await service.WaitForInputAsync(text => text.Contains("dotnet-coverage collect dotnet test\r"));
        Assert.Equal($"cd \"C:\\src\\app\"\rdotnet-coverage collect dotnet test\r", input);
        Assert.Equal(service.NextTerminalId, Assert.Single(service.AttachedTerminals));
        Assert.Equal(0, service.DisposeCount);
    }

    [Fact]
    public async Task CreateTerminal_AttachesAnInputStreamWithoutTypingAnything()
    {
        var service = new StubTerminalService();
        var interop = Build(service);

        var outcome = await interop.CreateTerminalAsync("MCP", null, CancellationToken.None);

        Assert.True(outcome.Created);
        Assert.Equal(service.NextTerminalId, Assert.Single(service.AttachedTerminals));

        // Give the background reader a chance to observe nothing, then confirm the stream is idle.
        await Task.Delay(100);
        Assert.Equal(string.Empty, service.AttachedInput);
    }

    [Fact]
    public async Task CreateTerminal_ReportsDeveloperEnvironment_FromDefaultProfile()
    {
        var service = new StubTerminalService
        {
            DefaultProfile = new ProfileConfig("Windows PowerShell", "powershell.exe")
        };
        var interop = Build(service);

        var outcome = await interop.CreateTerminalAsync(null, null, CancellationToken.None);

        Assert.True(outcome.Created);
        Assert.False(outcome.DeveloperEnvironment);
    }

    [Fact]
    public async Task RunInTerminal_WithTerminalId_TypesIntoExistingTerminal()
    {
        var service = new StubTerminalService();
        var vsService = new StubVsTerminalService(service);
        var interop = new TerminalInterop(
            () => throw new InvalidOperationException("The broker path should not be used."),
            DirectStandInAssembly,
            _ => vsService);

        var created = await interop.CreateTerminalAsync("MCP", null, CancellationToken.None);
        var first = await interop.RunInTerminalAsync(
            created.TerminalId.ToString(), null, null, "dotnet build", CancellationToken.None);
        var second = await interop.RunInTerminalAsync(
            created.TerminalId.ToString(), null, @"C:\src\app", "dotnet test", CancellationToken.None);

        Assert.True(created.Created);
        Assert.True(first.Reused);
        Assert.False(first.Created);
        Assert.True(second.Reused);
        Assert.Equal(created.TerminalId, first.TerminalId);
        Assert.Equal(created.TerminalId, second.TerminalId);
        Assert.Equal(1, vsService.CreateCount);
        Assert.Single(service.AttachedTerminals);

        var input = await service.WaitForInputAsync(text => text.Contains("dotnet test\r"));
        Assert.Equal("dotnet build\rcd \"C:\\src\\app\"\rdotnet test\r", input);
        Assert.Equal(created.TerminalId, service.LastShown);
    }

    [Fact]
    public async Task RunInTerminal_WithoutWait_DoesNotReturnOutput()
    {
        var service = new StubTerminalService();
        var interop = Build(service);

        var outcome = await interop.RunInTerminalAsync(
            null, "MCP", null, "echo hello", CancellationToken.None);

        Assert.True(outcome.Created);
        Assert.Null(outcome.Output);
    }

    [Fact]
    public async Task RunInTerminal_ReturnsRawOutputCapturedDuringTheWait()
    {
        var service = new StubTerminalService();
        var interop = Build(service);

        var outcome = await interop.RunInTerminalAsync(
            null, "MCP", null, "echo hello", CancellationToken.None, waitMilliseconds: 500);

        Assert.True(outcome.Created);
        Assert.NotNull(outcome.Output);
        Assert.Contains("echo hello", outcome.Output);
    }

    [Fact]
    public async Task RunInTerminal_ReuseReturnsOnlyTheNewCommandOutput()
    {
        var service = new StubTerminalService();
        var interop = Build(service);

        var created = await interop.CreateTerminalAsync("MCP", null, CancellationToken.None);
        var first = await interop.RunInTerminalAsync(
            created.TerminalId.ToString(), null, null, "first-command", CancellationToken.None, 400);
        var second = await interop.RunInTerminalAsync(
            created.TerminalId.ToString(), null, null, "second-command", CancellationToken.None, 400);

        Assert.NotNull(first.Output);
        Assert.Contains("first-command", first.Output);
        Assert.NotNull(second.Output);
        Assert.Contains("second-command", second.Output);
        Assert.DoesNotContain("first-command", second.Output);
    }

    [Fact]
    public async Task RunInTerminal_WithExistingTerminal_AttachesWhenStreamIsNotOpen()
    {
        var service = new StubTerminalService();
        var existing = Guid.NewGuid();
        service.OpenTerminals.Add(existing);

        var interop = Build(service);
        var outcome = await interop.RunInTerminalAsync(
            existing.ToString(), null, null, "dir", CancellationToken.None);

        Assert.True(outcome.Reused);
        Assert.Equal(existing, Assert.Single(service.AttachedTerminals));
        Assert.Empty(service.CreatedTerminals);

        var input = await service.WaitForInputAsync(text => text.Contains("dir\r"));
        Assert.Equal("dir\r", input);
    }

    [Fact]
    public async Task RunInTerminal_WithUnknownTerminal_ReportsFailureWithoutCreatingOne()
    {
        var service = new StubTerminalService();
        var interop = Build(service);

        var outcome = await interop.RunInTerminalAsync(
            Guid.NewGuid().ToString(), null, null, "dir", CancellationToken.None);

        Assert.False(outcome.Created);
        Assert.False(outcome.Reused);
        Assert.Contains("does not accept external input", outcome.Message);
        Assert.Empty(service.CreatedTerminals);
    }

    [Fact]
    public async Task RunInTerminal_WithInvalidTerminalId_ReportsFailure()
    {
        var service = new StubTerminalService();
        var interop = Build(service);

        var outcome = await interop.RunInTerminalAsync("not-a-guid", null, null, "dir", CancellationToken.None);

        Assert.False(outcome.Created);
        Assert.Contains("not a valid terminal identifier", outcome.Message);
    }

    [Fact]
    public async Task CreateTerminal_UsesVsTerminalService_WhenDescriptorIsUnavailable()
    {
        var service = new StubTerminalService();
        var vsService = new StubVsTerminalService(service);
        var brokerAccessed = false;
        var interop = new TerminalInterop(
            () =>
            {
                brokerAccessed = true;
                return null;
            },
            DirectStandInAssembly,
            serviceType =>
            {
                Assert.Equal(typeof(SVsTerminalService), serviceType);
                return vsService;
            });

        var outcome = await interop.CreateTerminalAsync("MCP", @"C:\src\app", CancellationToken.None);

        Assert.True(outcome.Created);
        Assert.Equal(service.NextTerminalId, outcome.TerminalId);
        Assert.NotNull(service.LastWindowOptions);
        Assert.Equal("MCP", service.LastWindowOptions!.Name);
        Assert.Equal(@"C:\src\app", service.LastWindowOptions.WorkingDirectory);
        Assert.Same(service.DefaultProfile, service.LastWindowOptions.Profile);
        Assert.Equal(1, vsService.CreateCount);
        Assert.Single(service.AttachedTerminals);

        Assert.True(await interop.CloseTerminalAsync(outcome.TerminalId, default));
        Assert.False(brokerAccessed);
        Assert.Equal(0, service.DisposeCount);

        interop.Dispose();
        Assert.Equal(0, service.DisposeCount);
    }

    [Fact]
    public async Task TerminalOperations_ReuseVsTerminalService()
    {
        var service = new StubTerminalService();
        var vsService = new StubVsTerminalService(service);
        var interop = new TerminalInterop(
            () => throw new InvalidOperationException("The broker path should not be used."),
            DirectStandInAssembly,
            _ => vsService);

        var outcome = await interop.RunInTerminalAsync(null, "MCP", @"C:\src\app", "dir", CancellationToken.None);
        var ids = await interop.GetTerminalIdsAsync(CancellationToken.None);
        var shown = await interop.ShowTerminalAsync(outcome.TerminalId, CancellationToken.None);
        var closed = await interop.CloseTerminalAsync(outcome.TerminalId, CancellationToken.None);

        Assert.True(outcome.Created);
        Assert.Equal(outcome.TerminalId, Assert.Single(ids!));
        Assert.True(shown);
        Assert.True(closed);
        Assert.Equal(1, vsService.CreateCount);
        Assert.Equal(0, service.DisposeCount);
    }

    [Fact]
    public async Task RunInTerminal_RetriesAfterAssemblyBecomesAvailable()
    {
        var service = new StubTerminalService();
        var vsService = new StubVsTerminalService(service);
        var assemblyAvailable = false;
        var interop = new TerminalInterop(
            () => null,
            () => assemblyAvailable ? DirectStandInAssembly() : null,
            _ => vsService);

        var first = await interop.RunInTerminalAsync(null, null, null, "dir", CancellationToken.None);
        assemblyAvailable = true;
        var second = await interop.RunInTerminalAsync(null, null, null, "dir", CancellationToken.None);

        Assert.False(first.Created);
        Assert.True(second.Created);
    }

    [Fact]
    public async Task RunInTerminal_RetriesAfterVsServiceBecomesAvailable()
    {
        var service = new StubTerminalService();
        var vsService = new StubVsTerminalService(service);
        var serviceAvailable = false;
        var interop = new TerminalInterop(
            () => null,
            DirectStandInAssembly,
            _ => serviceAvailable ? vsService : null);

        var first = await interop.RunInTerminalAsync(null, null, null, "dir", CancellationToken.None);
        serviceAvailable = true;
        var second = await interop.RunInTerminalAsync(null, null, null, "dir", CancellationToken.None);

        Assert.False(first.Created);
        Assert.True(second.Created);
        Assert.Equal(1, vsService.CreateCount);
    }

    [Fact]
    public async Task CreateTerminal_ReportsUnavailable_WhenTerminalAssemblyMissing()
    {
        var interop = new TerminalInterop(() => new StubServiceBroker(new StubTerminalService()), () => null);

        var outcome = await interop.CreateTerminalAsync(null, null, CancellationToken.None);

        Assert.False(outcome.Created);
        Assert.Equal(TerminalInterop.UnavailableMessage, outcome.Message);
    }

    [Fact]
    public async Task CreateTerminal_ReportsUnavailable_WhenBrokerMissing()
    {
        var interop = new TerminalInterop(() => null, BrokerStandInAssembly);

        var outcome = await interop.CreateTerminalAsync(null, null, CancellationToken.None);

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
        var interop = new TerminalInterop(() => null, BrokerStandInAssembly);

        // Null is distinct from an empty list, which would mean "no terminals are open".
        Assert.Null(await interop.GetTerminalIdsAsync(CancellationToken.None));
    }

    [Fact]
    public async Task GetTerminalIds_DropsSessionsForClosedTerminals()
    {
        var service = new StubTerminalService();
        var interop = Build(service);

        var outcome = await interop.CreateTerminalAsync(null, null, CancellationToken.None);
        service.OpenTerminals.Clear();

        await interop.GetTerminalIdsAsync(CancellationToken.None);

        // The stale session was pruned, so the same terminal must be attached again.
        service.OpenTerminals.Add(outcome.TerminalId);
        await interop.RunInTerminalAsync(outcome.TerminalId.ToString(), null, null, "dir", CancellationToken.None);

        Assert.Equal(2, service.AttachedTerminals.Count);
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

    private static TerminalInterop Build(ITerminalService service) =>
        new(() => new StubServiceBroker(service), BrokerStandInAssembly);

    private sealed class FilteredAssembly : Assembly
    {
        private readonly Assembly _assembly;
        private readonly HashSet<string> _hiddenTypes;

        internal FilteredAssembly(Assembly assembly, params string[] hiddenTypes)
        {
            _assembly = assembly;
            _hiddenTypes = new HashSet<string>(hiddenTypes, StringComparer.Ordinal);
        }

        public override Type? GetType(string name) =>
            _hiddenTypes.Contains(name) ? null : _assembly.GetType(name);
    }

    private sealed class StubVsTerminalService : IVsTerminalService
    {
        private readonly ITerminalService _service;

        internal StubVsTerminalService(ITerminalService service)
        {
            _service = service;
        }

        internal int CreateCount { get; private set; }

        public ITerminalService CreateTerminalService()
        {
            CreateCount++;
            return _service;
        }
    }

    private sealed class StubTerminalService : ITerminalService, IDisposable
    {
        private readonly object _inputGate = new();
        private readonly StringBuilder _attachedInput = new();

        internal Guid NextTerminalId { get; } = Guid.NewGuid();

        internal List<Guid> OpenTerminals { get; } = new();

        internal List<Guid> CreatedTerminals { get; } = new();

        internal ProfileConfig DefaultProfile { get; set; } =
            new("Developer PowerShell", "pwsh.exe", DeveloperShellArguments);

        internal string? LastName { get; private set; }

        internal ProfileConfig? LastProfile { get; private set; }

        internal string? LastWorkingDirectory { get; private set; }

        internal TerminalWindowOptions? LastWindowOptions { get; private set; }

        internal List<Guid> AttachedTerminals { get; } = new();

        internal string AttachedInput
        {
            get
            {
                lock (_inputGate)
                {
                    return _attachedInput.ToString();
                }
            }
        }

        internal Guid? LastShown { get; private set; }

        internal Guid? LastClosed { get; private set; }

        internal int CloseAllCount { get; private set; }

        internal int DisposeCount { get; private set; }

        public Task<Guid> CreateTerminalWindowAsync(
            CancellationToken cancellationToken,
            TerminalWindowOptions options)
        {
            LastWindowOptions = options;
            LastName = options.Name;
            LastProfile = options.Profile as ProfileConfig;
            LastWorkingDirectory = options.WorkingDirectory;

            CreatedTerminals.Add(NextTerminalId);
            OpenTerminals.Add(NextTerminalId);
            return Task.FromResult(NextTerminalId);
        }

        public Task<Guid> CreateTerminalAsync(
            CancellationToken cancellationToken,
            string? name,
            ITerminalProfile? profile)
        {
            LastName = name;
            LastProfile = profile as ProfileConfig;
            LastWorkingDirectory = null;

            CreatedTerminals.Add(NextTerminalId);
            OpenTerminals.Add(NextTerminalId);
            return Task.FromResult(NextTerminalId);
        }

        public Task<ITerminalProfile> GetDefaultProfileAsync(CancellationToken cancellationToken)
            => Task.FromResult<ITerminalProfile>(DefaultProfile);

        public Task AttachTerminalStreamAsync(
            Guid terminalGuid,
            Stream stream,
            CancellationToken cancellationToken)
        {
            if (!OpenTerminals.Contains(terminalGuid))
            {
                throw new InvalidOperationException($"Terminal {terminalGuid} is not open.");
            }

            lock (_inputGate)
            {
                AttachedTerminals.Add(terminalGuid);
            }

            _ = Task.Run(async () =>
            {
                var buffer = new byte[1024];
                try
                {
                    while (true)
                    {
                        var read = await stream.ReadAsync(buffer, 0, buffer.Length).ConfigureAwait(false);
                        if (read == 0)
                        {
                            break;
                        }

                        var text = Encoding.UTF8.GetString(buffer, 0, read);
                        lock (_inputGate)
                        {
                            _attachedInput.Append(text);
                        }

                        // Act like a PTY that echoes its input so tests can assert captured output.
                        await stream.WriteAsync(buffer, 0, read).ConfigureAwait(false);
                        await stream.FlushAsync().ConfigureAwait(false);
                    }
                }
                catch (Exception)
                {
                    // The production side disposes the stream when the terminal session ends.
                }
            });

            return Task.CompletedTask;
        }

        internal async Task<string> WaitForInputAsync(
            Func<string, bool> predicate,
            int timeoutMilliseconds = 3000)
        {
            var deadline = DateTime.UtcNow.AddMilliseconds(timeoutMilliseconds);
            while (DateTime.UtcNow < deadline)
            {
                var text = AttachedInput;
                if (predicate(text))
                {
                    return text;
                }

                await Task.Delay(20);
            }

            return AttachedInput;
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

        public void Dispose() => DisposeCount++;
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
