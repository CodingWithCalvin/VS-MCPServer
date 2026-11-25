using System;
using System.Collections;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using CodingWithCalvin.Otel4Vsix;
using Microsoft.ServiceHub.Framework;
using Nerdbank.Streams;

namespace CodingWithCalvin.MCPServer.Services;

/// <summary>
/// Reflection bridge to the Visual Studio integrated terminal.
/// </summary>
/// <remarks>
/// <para>
/// The brokered service plumbing itself is ordinary compile-time code:
/// <c>IBrokeredServiceContainer</c>, <c>IServiceBroker</c>, <c>ServiceRpcDescriptor</c> and
/// <c>ServiceActivationOptions</c> all come from Microsoft.ServiceHub.Framework, which is a real
/// NuGet package the VSIX already references.
/// </para>
/// <para>
/// What cannot be referenced is <c>Microsoft.VisualStudio.Terminal.dll</c>, which supplies
/// <c>ITerminalService</c>, <c>ProfileConfig</c> and the version-specific service entry points.
/// It ships only inside the Visual Studio installation, is absent from NuGet and from the
/// Microsoft.VisualStudio.SDK metapackage, and <c>$(DevEnvDir)</c> is undefined under
/// <c>dotnet build</c> - the same constraint documented on
/// <see cref="TestExplorerInterop"/>. These types are therefore reached reflectively.
/// </para>
/// <para>
/// VS 2022 exposes the service through <c>TerminalServiceDescriptors</c> and the brokered
/// service container. VS 2026 removed that descriptor and instead exposes
/// <c>SVsTerminalService</c> and <c>IVsTerminalService.CreateTerminalService()</c>.
/// The assembly is loaded by strong name at version 17.0.0.0; both versions ship a binding
/// redirect and codeBase entry that resolve it to the installed assembly.
/// </para>
/// <para>
/// Terminals are opened with the user's default terminal profile instead of a synthetic
/// cmd.exe /k profile. In VS 2026 that profile is Developer PowerShell, which imports the
/// Visual Studio developer environment, so msbuild, vstest.console and dotnet-coverage are on
/// PATH without chaining VsDevCmd.bat. A duplex stream is then attached with
/// <c>AttachTerminalStreamAsync</c>; the terminal reads that stream as user input and mirrors
/// its PTY output back into it. Keeping the stream open lets commands be typed into a terminal
/// that is already running.
/// </para>
/// <para>
/// Every failure path degrades to a message rather than throwing. The terminal component may be
/// absent, and on a future Visual Studio the shape could change; neither should take an MCP tool
/// call down with it.
/// </para>
/// </remarks>
internal sealed class TerminalInterop : IDisposable
{
    private const string TerminalAssemblyName =
        "Microsoft.VisualStudio.Terminal, Version=17.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a";
    private const string SimpleAssemblyName = "Microsoft.VisualStudio.Terminal";
    private const string DescriptorsTypeName = "Microsoft.VisualStudio.Terminal.TerminalServiceDescriptors";
    private const string VsServiceTypeName = "Microsoft.VisualStudio.Terminal.SVsTerminalService";
    private const string VsServiceInterfaceTypeName = "Microsoft.VisualStudio.Terminal.IVsTerminalService";
    private const string ServiceTypeName = "Microsoft.VisualStudio.Terminal.ITerminalService";
    private const string ProfileTypeName = "Microsoft.VisualStudio.Terminal.ITerminalProfile";
    private const string TerminalWindowOptionsTypeName = "Microsoft.VisualStudio.Terminal.TerminalWindowOptions";

    internal const string UnavailableMessage =
        "The Visual Studio integrated terminal component is unavailable in this instance.";

    private readonly Func<IServiceBroker?> _serviceBrokerAccessor;
    private readonly Func<Assembly?> _assemblyResolver;
    private readonly Func<Type, object?>? _serviceAccessor;

    private Assembly? _terminalAssembly;
    private Type? _serviceType;
    private object? _directService;
    private object? _brokeredService;

    /// <summary>
    /// Input streams attached to open terminals, keyed by terminal identifier. A terminal only
    /// accepts typed commands while its stream is attached.
    /// </summary>
    private readonly ConcurrentDictionary<Guid, TerminalSession> _sessions = new();

    internal TerminalInterop(
        Func<IServiceBroker?> serviceBrokerAccessor,
        Func<Type, object?> serviceAccessor)
        : this(serviceBrokerAccessor, TryLoadTerminalAssembly, serviceAccessor)
    {
    }

    /// <summary>
    /// Overload taking explicit resolvers, so that tests can supply stand-in terminal types and
    /// services without a Visual Studio installation.
    /// </summary>
    internal TerminalInterop(
        Func<IServiceBroker?> serviceBrokerAccessor,
        Func<Assembly?> assemblyResolver,
        Func<Type, object?>? serviceAccessor = null)
    {
        _serviceBrokerAccessor = serviceBrokerAccessor;
        _assemblyResolver = assemblyResolver;
        _serviceAccessor = serviceAccessor;
    }

    /// <summary>
    /// Opens a terminal with the user's default profile and attaches an input stream so that
    /// later commands can be typed into it.
    /// </summary>
    internal async Task<TerminalCreateOutcome> CreateTerminalAsync(
        string? name,
        string? workingDirectory,
        CancellationToken cancellationToken)
    {
        var proxy = await AcquireProxyAsync(cancellationToken).ConfigureAwait(false);
        if (proxy == null)
        {
            return TerminalCreateOutcome.Failure(UnavailableMessage);
        }

        using (proxy)
        {
            try
            {
                var terminal = await OpenTerminalAsync(proxy, name, workingDirectory, cancellationToken)
                    .ConfigureAwait(false);

                return terminal.Session == null
                    ? TerminalCreateOutcome.Failure(terminal.Message ?? UnavailableMessage)
                    : TerminalCreateOutcome.Success(terminal.TerminalId, terminal.DeveloperEnvironment);
            }
            catch (Exception ex)
            {
                VsixTelemetry.TrackException(ex);
                return TerminalCreateOutcome.Failure(ex.Message);
            }
        }
    }

    /// <summary>
    /// Types <paramref name="command"/> into an existing terminal, or opens a new terminal and
    /// types it there when <paramref name="terminalId"/> is not supplied.
    /// </summary>
    internal async Task<TerminalCreateOutcome> RunInTerminalAsync(
        string? terminalId,
        string? name,
        string? workingDirectory,
        string command,
        CancellationToken cancellationToken,
        int waitMilliseconds = 0)
    {
        if (string.IsNullOrWhiteSpace(command))
        {
            return TerminalCreateOutcome.Failure("A command is required.");
        }

        var proxy = await AcquireProxyAsync(cancellationToken).ConfigureAwait(false);
        if (proxy == null)
        {
            return TerminalCreateOutcome.Failure(UnavailableMessage);
        }

        using (proxy)
        {
            try
            {
                if (!string.IsNullOrWhiteSpace(terminalId))
                {
                    if (!Guid.TryParse(terminalId, out var existingId))
                    {
                        return TerminalCreateOutcome.Failure(
                            $"'{terminalId}' is not a valid terminal identifier.");
                    }

                    var session = await AttachSessionAsync(proxy, existingId, cancellationToken)
                        .ConfigureAwait(false);
                    if (session == null)
                    {
                        return TerminalCreateOutcome.Failure(
                            $"Terminal {terminalId} is not open, or it does not accept external input. "
                            + "Use terminal_create or terminal_run to open a terminal and reuse the identifier it returns.");
                    }

                    await ShowCoreAsync(proxy, existingId, cancellationToken).ConfigureAwait(false);

                    if (!string.IsNullOrWhiteSpace(workingDirectory))
                    {
                        await session.SendAsync(ChangeDirectory(workingDirectory!), cancellationToken)
                            .ConfigureAwait(false);
                    }

                    var baseline = session.CapturedLength;
                    await session.SendAsync(command + "\r", cancellationToken).ConfigureAwait(false);
                    return TerminalCreateOutcome.Submitted(
                        existingId,
                        await ReadOutputAsync(session, baseline, waitMilliseconds).ConfigureAwait(false));
                }

                var terminal = await OpenTerminalAsync(proxy, name, workingDirectory, cancellationToken)
                    .ConfigureAwait(false);
                if (terminal.Session == null)
                {
                    return TerminalCreateOutcome.Failure(terminal.Message ?? UnavailableMessage);
                }

                var createdBaseline = terminal.Session.CapturedLength;
                await terminal.Session.SendAsync(command + "\r", cancellationToken).ConfigureAwait(false);
                return TerminalCreateOutcome.Success(
                    terminal.TerminalId,
                    terminal.DeveloperEnvironment,
                    await ReadOutputAsync(terminal.Session, createdBaseline, waitMilliseconds)
                        .ConfigureAwait(false));
            }
            catch (Exception ex)
            {
                VsixTelemetry.TrackException(ex);
                return TerminalCreateOutcome.Failure(ex.Message);
            }
        }
    }

    /// <summary>
    /// Waits for the terminal to print and returns whatever it captured after
    /// <paramref name="baseline"/>. A PTY has no completion signal, so the wait window is the
    /// only bound on what comes back.
    /// </summary>
    private static async Task<string?> ReadOutputAsync(
        TerminalSession session,
        long baseline,
        int waitMilliseconds)
    {
        if (waitMilliseconds <= 0)
        {
            return null;
        }

        await Task.Delay(waitMilliseconds).ConfigureAwait(false);
        var output = session.ReadCapturedSince(baseline);
        return output.Length == 0 ? null : output;
    }

    /// <summary>
    /// Creates a terminal in the default profile, wires up its input stream, and applies the
    /// working directory where the creation API cannot do so itself.
    /// </summary>
    private async Task<OpenedTerminal> OpenTerminalAsync(
        TerminalProxy proxy,
        string? name,
        string? workingDirectory,
        CancellationToken cancellationToken)
    {
        var profile = await GetDefaultProfileAsync(proxy.Service, cancellationToken).ConfigureAwait(false);
        var developerEnvironment = IsDeveloperProfile(profile);

        var terminalId = await CreateWindowAsync(proxy, name, workingDirectory, profile, cancellationToken)
            .ConfigureAwait(false);
        if (terminalId == null)
        {
            return OpenedTerminal.Failure("Visual Studio did not return a terminal identifier.");
        }

        var session = await AttachSessionAsync(proxy, terminalId.Value, cancellationToken).ConfigureAwait(false);
        if (session == null)
        {
            // Do not leave a window behind that no caller can ever reach.
            await CloseCoreAsync(proxy, terminalId.Value, cancellationToken).ConfigureAwait(false);
            return OpenedTerminal.Failure("Visual Studio did not accept an input stream for the new terminal.");
        }

        // VS 2022's window overload has no working-directory argument, so change directory
        // through the shell itself. VS 2026 stores it on the window options instead.
        if (!proxy.UseTerminalWindowOptions && !string.IsNullOrWhiteSpace(workingDirectory))
        {
            await session.SendAsync(ChangeDirectory(workingDirectory!), cancellationToken).ConfigureAwait(false);
        }

        return OpenedTerminal.Success(terminalId.Value, session, developerEnvironment);
    }

    private async Task<Guid?> CreateWindowAsync(
        TerminalProxy proxy,
        string? name,
        string? workingDirectory,
        object? profile,
        CancellationToken cancellationToken)
    {
        MethodInfo? method;
        object?[] arguments;

        if (proxy.UseTerminalWindowOptions)
        {
            var optionsType = _terminalAssembly!.GetType(TerminalWindowOptionsTypeName);
            if (optionsType == null)
            {
                return null;
            }

            method = _serviceType!.GetMethod(
                "CreateTerminalWindowAsync",
                new[] { typeof(CancellationToken), optionsType });
            var options = Activator.CreateInstance(optionsType);
            if (method == null || options == null)
            {
                return null;
            }

            SetProperty(options, "Name", name);
            SetProperty(options, "Profile", profile);
            SetProperty(options, "WorkingDirectory", workingDirectory);
            arguments = new object?[] { cancellationToken, options };
        }
        else
        {
            // Use the direct window overload rather than the renderer-based one: only a window
            // owns the PTY in a way that accepts an attached input stream for later commands.
            var profileType = _terminalAssembly!.GetType(ProfileTypeName);
            method = _serviceType!.GetMethods().FirstOrDefault(m =>
                m.Name == "CreateTerminalAsync"
                && m.GetParameters() is { Length: 3 } parameters
                && parameters[0].ParameterType == typeof(CancellationToken)
                && parameters[1].ParameterType == typeof(string)
                && profileType != null
                && parameters[2].ParameterType == profileType);
            arguments = new object?[] { cancellationToken, name, profile };
        }

        if (method == null)
        {
            return null;
        }

        var result = await InvokeAsync(method, proxy.Service, arguments).ConfigureAwait(false);
        return result is Guid id ? id : null;
    }

    private async Task<object?> GetDefaultProfileAsync(object service, CancellationToken cancellationToken)
    {
        try
        {
            var method = _serviceType!.GetMethod("GetDefaultProfileAsync", new[] { typeof(CancellationToken) });
            return method == null
                ? null
                : await InvokeAsync(method, service, new object?[] { cancellationToken }).ConfigureAwait(false);
        }
        catch (Exception ex)
        {
            // A terminal can still be opened without a profile: Visual Studio then applies its
            // own default. Losing the profile only makes the developer-environment hint unknown.
            VsixTelemetry.TrackException(ex);
            return null;
        }
    }

    /// <summary>
    /// Attaches a duplex stream to a terminal. The terminal reads what is written to the stream
    /// as user input and copies its PTY output back into it, so the host half must stay drained
    /// or the terminal's output pipeline stalls.
    /// </summary>
    private async Task<TerminalSession?> AttachSessionAsync(
        TerminalProxy proxy,
        Guid terminalId,
        CancellationToken cancellationToken)
    {
        if (_sessions.TryGetValue(terminalId, out var existing))
        {
            return existing;
        }

        var method = _serviceType!.GetMethod(
            "AttachTerminalStreamAsync",
            new[] { typeof(Guid), typeof(Stream), typeof(CancellationToken) });
        if (method == null)
        {
            return null;
        }

        var (terminalHalf, hostHalf) = FullDuplexStream.CreatePair();
        try
        {
            await InvokeAsync(method, proxy.Service, new object?[] { terminalId, terminalHalf, cancellationToken })
                .ConfigureAwait(false);
        }
        catch (Exception ex)
        {
            VsixTelemetry.TrackException(ex);
            terminalHalf.Dispose();
            hostHalf.Dispose();
            return null;
        }

        var session = new TerminalSession(terminalHalf, hostHalf);
        _sessions[terminalId] = session;
        session.Start();
        return session;
    }

    private async Task ShowCoreAsync(TerminalProxy proxy, Guid terminalId, CancellationToken cancellationToken)
    {
        try
        {
            var method = _serviceType!.GetMethod("ShowAsync", new[] { typeof(Guid), typeof(CancellationToken) });
            if (method != null)
            {
                await InvokeAsync(method, proxy.Service, new object?[] { terminalId, cancellationToken })
                    .ConfigureAwait(false);
            }
        }
        catch (Exception ex)
        {
            // Failing to raise the window must not stop the command from being submitted.
            VsixTelemetry.TrackException(ex);
        }
    }

    private async Task CloseCoreAsync(TerminalProxy proxy, Guid terminalId, CancellationToken cancellationToken)
    {
        var method = _serviceType!.GetMethod("CloseAsync", new[] { typeof(Guid), typeof(CancellationToken) });
        if (method != null)
        {
            await InvokeAsync(method, proxy.Service, new object?[] { terminalId, cancellationToken })
                .ConfigureAwait(false);
        }
    }

    private static bool IsDeveloperProfile(object? profile)
    {
        var arguments = profile?.GetType().GetProperty("Arguments")?.GetValue(profile) as string;
        return arguments != null
            && (arguments.IndexOf("Enter-VsDevShell", StringComparison.OrdinalIgnoreCase) >= 0
                || arguments.IndexOf("VsDevCmd.bat", StringComparison.OrdinalIgnoreCase) >= 0);
    }

    /// <summary>Queues a directory change that both PowerShell and cmd.exe understand.</summary>
    private static string ChangeDirectory(string workingDirectory) => $"cd \"{workingDirectory}\"\r";

    internal async Task<IReadOnlyList<Guid>?> GetTerminalIdsAsync(CancellationToken cancellationToken)
    {
        var proxy = await AcquireProxyAsync(cancellationToken).ConfigureAwait(false);
        if (proxy == null)
        {
            return null;
        }

        using (proxy)
        {
            var method = _serviceType!.GetMethod("GetTerminalGuidsAsync", new[] { typeof(CancellationToken) });
            if (method == null)
            {
                return null;
            }

            var result = await InvokeAsync(method, proxy.Service, new object?[] { cancellationToken })
                .ConfigureAwait(false);

            if (result is not IEnumerable ids)
            {
                return null;
            }

            var terminalIds = ids.Cast<Guid>().ToList();
            PruneSessions(terminalIds);
            return terminalIds;
        }
    }

    internal Task<bool> ShowTerminalAsync(Guid terminalId, CancellationToken cancellationToken) =>
        InvokeByIdAsync("ShowAsync", terminalId, cancellationToken);

    internal async Task<bool> CloseTerminalAsync(Guid terminalId, CancellationToken cancellationToken)
    {
        var closed = await InvokeByIdAsync("CloseAsync", terminalId, cancellationToken).ConfigureAwait(false);
        if (closed)
        {
            RemoveSession(terminalId);
        }

        return closed;
    }

    internal async Task<bool> CloseAllTerminalsAsync(CancellationToken cancellationToken)
    {
        var proxy = await AcquireProxyAsync(cancellationToken).ConfigureAwait(false);
        if (proxy == null)
        {
            return false;
        }

        using (proxy)
        {
            var method = _serviceType!.GetMethod("CloseAllInstancesAsync", new[] { typeof(CancellationToken) });
            if (method == null)
            {
                return false;
            }

            await InvokeAsync(method, proxy.Service, new object?[] { cancellationToken }).ConfigureAwait(false);

            foreach (var terminalId in _sessions.Keys)
            {
                RemoveSession(terminalId);
            }

            return true;
        }
    }

    private async Task<bool> InvokeByIdAsync(
        string methodName,
        Guid terminalId,
        CancellationToken cancellationToken)
    {
        var proxy = await AcquireProxyAsync(cancellationToken).ConfigureAwait(false);
        if (proxy == null)
        {
            return false;
        }

        using (proxy)
        {
            var method = _serviceType!.GetMethod(methodName, new[] { typeof(Guid), typeof(CancellationToken) });
            if (method == null)
            {
                return false;
            }

            await InvokeAsync(method, proxy.Service, new object?[] { terminalId, cancellationToken })
                .ConfigureAwait(false);

            return true;
        }
    }

    /// <summary>
    /// Drops sessions whose terminal was closed outside these tools, so a stale stream is not
    /// reused for a later command.
    /// </summary>
    private void PruneSessions(IReadOnlyCollection<Guid> liveTerminalIds)
    {
        foreach (var terminalId in _sessions.Keys)
        {
            if (!liveTerminalIds.Contains(terminalId))
            {
                RemoveSession(terminalId);
            }
        }
    }

    private void RemoveSession(Guid terminalId)
    {
        if (_sessions.TryRemove(terminalId, out var session))
        {
            session.Dispose();
        }
    }

    private static void SetProperty(object target, string name, object? value) =>
        target.GetType().GetProperty(name)?.SetValue(target, value);

    /// <summary>
    /// Invokes a reflected async method and unwraps its result. The declared return type is used
    /// rather than the runtime type, because an async method returning a non-generic Task is
    /// backed by Task&lt;VoidTaskResult&gt; and would otherwise yield a meaningless value.
    /// </summary>
    private static async Task<object?> InvokeAsync(MethodInfo method, object target, object?[] arguments)
    {
        var pending = method.Invoke(target, arguments);
        if (pending is not Task task)
        {
            return null;
        }

        await task.ConfigureAwait(false);

        var returnType = method.ReturnType;
        if (!returnType.IsGenericType || returnType.GetGenericTypeDefinition() != typeof(Task<>))
        {
            return null;
        }

        return returnType.GetProperty("Result")?.GetValue(task);
    }

    /// <summary>
    /// Acquires the terminal service through the VS 2026 service provider when available, or
    /// through the VS 2022 brokered-service descriptor otherwise. Both paths retain one service
    /// instance because terminal IDs are stored per instance. Only the brokered proxy is owned
    /// by this object; directly acquired VS services remain owned by VS.
    /// </summary>
    private async Task<TerminalProxy?> AcquireProxyAsync(CancellationToken cancellationToken)
    {
        try
        {
            if (!TryResolveTypes())
            {
                return null;
            }

            var vsServiceType = _terminalAssembly!.GetType(VsServiceTypeName);
            var vsServiceInterfaceType = _terminalAssembly.GetType(VsServiceInterfaceTypeName);
            if (vsServiceType != null && vsServiceInterfaceType != null)
            {
                if (_directService == null)
                {
                    var vsService = _serviceAccessor?.Invoke(vsServiceType);
                    if (vsService == null)
                    {
                        return null;
                    }

                    var createService = vsServiceInterfaceType.GetMethod("CreateTerminalService", Type.EmptyTypes);
                    var directService = createService?.Invoke(vsService, null);
                    if (directService == null || !_serviceType!.IsInstanceOfType(directService))
                    {
                        return null;
                    }

                    // Each factory call returns a service with an independent terminal list.
                    // Reuse one instance so create/list/show/close all observe the same terminals.
                    _directService = directService;
                }

                return new TerminalProxy(
                    _directService,
                    ownsService: false,
                    useTerminalWindowOptions: true);
            }

            if (_brokeredService != null)
            {
                return new TerminalProxy(_brokeredService, ownsService: false);
            }

            var descriptors = _terminalAssembly.GetType(DescriptorsTypeName);
            var descriptor = descriptors
                ?.GetProperty("TerminalServiceDescriptor", BindingFlags.Public | BindingFlags.Static)
                ?.GetValue(null) as ServiceRpcDescriptor;
            if (descriptor == null)
            {
                return null;
            }

            var broker = _serviceBrokerAccessor();
            if (broker == null)
            {
                return null;
            }

            var getProxy = typeof(IServiceBroker)
                .GetMethod(nameof(IServiceBroker.GetProxyAsync))
                ?.MakeGenericMethod(_serviceType!);
            if (getProxy == null)
            {
                return null;
            }

            var pending = getProxy.Invoke(
                broker,
                new object?[] { descriptor, default(ServiceActivationOptions), cancellationToken });

            // GetProxyAsync returns ValueTask<T>, which has to be turned into a Task before it
            // can be awaited through reflection.
            var asTask = pending?.GetType().GetMethod("AsTask", Type.EmptyTypes);
            if (asTask?.Invoke(pending, null) is not Task task)
            {
                return null;
            }

            await task.ConfigureAwait(false);

            var service = task.GetType().GetProperty("Result")?.GetValue(task);
            if (service == null)
            {
                return null;
            }

            _brokeredService = service;
            return new TerminalProxy(service, ownsService: false);
        }
        catch (Exception ex)
        {
            VsixTelemetry.TrackException(ex);
            return null;
        }
    }

    private bool TryResolveTypes()
    {
        if (_terminalAssembly != null && _serviceType != null)
        {
            return true;
        }

        var assembly = _assemblyResolver();
        if (assembly == null)
        {
            return false;
        }

        var serviceType = assembly.GetType(ServiceTypeName);
        if (serviceType == null)
        {
            return false;
        }

        _terminalAssembly = assembly;
        _serviceType = serviceType;
        return true;
    }

    private static Assembly? TryLoadTerminalAssembly()
    {
        try
        {
            // Prefer an already-loaded copy: Visual Studio loads this itself once the terminal
            // window has been used, and reusing it avoids any dependence on binding policy.
            var loaded = AppDomain.CurrentDomain.GetAssemblies()
                .FirstOrDefault(a => string.Equals(
                    a.GetName().Name,
                    SimpleAssemblyName,
                    StringComparison.OrdinalIgnoreCase));

            return loaded ?? Assembly.Load(TerminalAssemblyName);
        }
        catch (Exception ex)
        {
            VsixTelemetry.TrackException(ex);
            return null;
        }
    }

    public void Dispose()
    {
        foreach (var terminalId in _sessions.Keys)
        {
            RemoveSession(terminalId);
        }

        (_brokeredService as IDisposable)?.Dispose();
        _brokeredService = null;
    }

    /// <summary>Outcome of opening a terminal, including the stream used to type into it.</summary>
    private sealed class OpenedTerminal
    {
        internal Guid TerminalId { get; private set; }

        internal TerminalSession? Session { get; private set; }

        internal bool DeveloperEnvironment { get; private set; }

        internal string? Message { get; private set; }

        internal static OpenedTerminal Success(Guid terminalId, TerminalSession session, bool developerEnvironment) =>
            new()
            {
                TerminalId = terminalId,
                Session = session,
                DeveloperEnvironment = developerEnvironment
            };

        internal static OpenedTerminal Failure(string message) => new() { Message = message };
    }

    /// <summary>
    /// Duplex stream pair attached to one terminal. The terminal half is given to Visual Studio,
    /// which reads it as user input and writes PTY output into it. Commands are written to the
    /// host half, and the PTY output read from it is kept in a bounded buffer so callers can read
    /// back what the terminal printed; draining it also stops the terminal blocking on a full
    /// pipe when nothing else reads its output.
    /// </summary>
    private sealed class TerminalSession : IDisposable
    {
        private const int DrainBufferSize = 4096;

        /// <summary>Cap on retained output, keeping the most recent text.</summary>
        private const int MaxCapturedCharacters = 256 * 1024;

        private readonly Stream _terminalHalf;
        private readonly Stream _hostHalf;
        private readonly CancellationTokenSource _drainCancellation = new();
        private readonly SemaphoreSlim _writeGate = new(1, 1);
        private readonly Decoder _decoder = Encoding.UTF8.GetDecoder();
        private readonly object _outputGate = new();
        private readonly StringBuilder _captured = new();
        private long _capturedTotal;

        internal TerminalSession(Stream terminalHalf, Stream hostHalf)
        {
            _terminalHalf = terminalHalf;
            _hostHalf = hostHalf;
        }

        internal void Start() => _ = Task.Run(DrainAsync);

        /// <summary>
        /// Number of characters captured so far, used as the start marker for a command so that
        /// output from earlier commands is not returned again.
        /// </summary>
        internal long CapturedLength
        {
            get
            {
                lock (_outputGate)
                {
                    return _capturedTotal;
                }
            }
        }

        /// <summary>
        /// Raw text captured since <paramref name="offset"/>, clipped to what is still retained.
        /// </summary>
        internal string ReadCapturedSince(long offset)
        {
            lock (_outputGate)
            {
                var firstRetained = _capturedTotal - _captured.Length;
                if (offset < firstRetained)
                {
                    offset = firstRetained;
                }

                if (offset >= _capturedTotal)
                {
                    return string.Empty;
                }

                var start = (int)(offset - firstRetained);
                return _captured.ToString(start, _captured.Length - start);
            }
        }

        private void AppendCaptured(char[] characters, int count)
        {
            lock (_outputGate)
            {
                _captured.Append(characters, 0, count);
                _capturedTotal += count;

                var overflow = _captured.Length - MaxCapturedCharacters;
                if (overflow > 0)
                {
                    _captured.Remove(0, overflow);
                }
            }
        }

        internal async Task SendAsync(string text, CancellationToken cancellationToken)
        {
            var bytes = Encoding.UTF8.GetBytes(text);

            await _writeGate.WaitAsync(cancellationToken).ConfigureAwait(false);
            try
            {
                await _hostHalf.WriteAsync(bytes, 0, bytes.Length, cancellationToken).ConfigureAwait(false);
                await _hostHalf.FlushAsync(cancellationToken).ConfigureAwait(false);
            }
            finally
            {
                _writeGate.Release();
            }
        }

        private async Task DrainAsync()
        {
            var buffer = new byte[DrainBufferSize];
            var characters = new char[DrainBufferSize + 1];
            try
            {
                while (!_drainCancellation.IsCancellationRequested)
                {
                    var read = await _hostHalf
                        .ReadAsync(buffer, 0, buffer.Length, _drainCancellation.Token)
                        .ConfigureAwait(false);
                    if (read == 0)
                    {
                        break;
                    }

                    var count = _decoder.GetChars(buffer, 0, read, characters, 0);
                    if (count > 0)
                    {
                        AppendCaptured(characters, count);
                    }
                }
            }
            catch (OperationCanceledException)
            {
            }
            catch (Exception)
            {
                // The stream is torn down when the terminal or the extension shuts down.
            }
        }

        public void Dispose()
        {
            try
            {
                _drainCancellation.Cancel();
            }
            catch (ObjectDisposedException)
            {
            }

            _drainCancellation.Dispose();
            _hostHalf.Dispose();
            _writeGate.Dispose();
            _terminalHalf.Dispose();
        }
    }
    /// <summary>Scopes the lifetime of an owned brokered proxy or a VS-owned service.</summary>
    private sealed class TerminalProxy : IDisposable
    {
        private readonly bool _ownsService;

        internal TerminalProxy(
            object service,
            bool ownsService = true,
            bool useTerminalWindowOptions = false)
        {
            Service = service;
            _ownsService = ownsService;
            UseTerminalWindowOptions = useTerminalWindowOptions;
        }

        internal object Service { get; }

        internal bool UseTerminalWindowOptions { get; }

        public void Dispose()
        {
            if (_ownsService)
            {
                (Service as IDisposable)?.Dispose();
            }
        }
    }
}

/// <summary>Result of a terminal operation, before it is shaped for the wire.</summary>
internal sealed class TerminalCreateOutcome
{
    private TerminalCreateOutcome()
    {
    }

    internal bool Created { get; private set; }

    internal bool Reused { get; private set; }

    internal Guid TerminalId { get; private set; }

    internal bool DeveloperEnvironment { get; private set; }

    internal string? Output { get; private set; }

    internal string? Message { get; private set; }

    internal static TerminalCreateOutcome Success(Guid id, bool developerEnvironment, string? output = null) => new()
    {
        Created = true,
        TerminalId = id,
        DeveloperEnvironment = developerEnvironment,
        Output = output
    };

    internal static TerminalCreateOutcome Submitted(Guid id, string? output = null) => new()
    {
        Reused = true,
        TerminalId = id,
        Output = output
    };

    internal static TerminalCreateOutcome Failure(string message) => new() { Message = message };
}

