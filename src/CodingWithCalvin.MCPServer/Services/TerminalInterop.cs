using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading;
using System.Threading.Tasks;
using CodingWithCalvin.Otel4Vsix;
using Microsoft.ServiceHub.Framework;

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
/// <c>ITerminalService</c>, <c>ProfileConfig</c> and <c>TerminalServiceDescriptors</c>. It ships
/// only inside the Visual Studio installation, is absent from NuGet and from the
/// Microsoft.VisualStudio.SDK metapackage, and <c>$(DevEnvDir)</c> is undefined under
/// <c>dotnet build</c> - the same constraint documented on
/// <see cref="TestExplorerInterop"/>. Those three types are therefore reached reflectively.
/// </para>
/// <para>
/// The assembly is loaded by strong name at version 17.0.0.0. Visual Studio ships a binding
/// redirect covering 0.0.0.0 through the installed version, plus a codeBase entry pointing at
/// CommonExtensions\Microsoft\Terminal, in both VS 2022 and VS 2026 - so one version number
/// resolves correctly on both.
/// </para>
/// <para>
/// Every failure path degrades to a message rather than throwing. The terminal component may be
/// absent, and on a future Visual Studio the shape could change; neither should take an MCP tool
/// call down with it.
/// </para>
/// </remarks>
internal sealed class TerminalInterop
{
    private const string TerminalAssemblyName =
        "Microsoft.VisualStudio.Terminal, Version=17.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a";
    private const string SimpleAssemblyName = "Microsoft.VisualStudio.Terminal";
    private const string DescriptorsTypeName = "Microsoft.VisualStudio.Terminal.TerminalServiceDescriptors";
    private const string ServiceTypeName = "Microsoft.VisualStudio.Terminal.ITerminalService";
    private const string ProfileConfigTypeName = "Microsoft.VisualStudio.Terminal.ProfileConfig";

    internal const string UnavailableMessage =
        "The Visual Studio integrated terminal component is unavailable in this instance.";

    private readonly Func<IServiceBroker?> _serviceBrokerAccessor;
    private readonly Func<Assembly?> _assemblyResolver;

    private bool _typesResolved;
    private Type? _serviceType;
    private Type? _profileConfigType;
    private ServiceRpcDescriptor? _descriptor;

    internal TerminalInterop(Func<IServiceBroker?> serviceBrokerAccessor)
        : this(serviceBrokerAccessor, TryLoadTerminalAssembly)
    {
    }

    /// <summary>
    /// Overload taking an explicit assembly resolver, so that tests can supply stand-in terminal
    /// types without a Visual Studio installation.
    /// </summary>
    internal TerminalInterop(Func<IServiceBroker?> serviceBrokerAccessor, Func<Assembly?> assemblyResolver)
    {
        _serviceBrokerAccessor = serviceBrokerAccessor;
        _assemblyResolver = assemblyResolver;
    }

    /// <summary>
    /// Opens a terminal, optionally launching a command in it. When a command is supplied it is
    /// run through cmd.exe with /k so the window stays open and its output remains readable.
    /// </summary>
    internal async Task<TerminalCreateOutcome> CreateTerminalAsync(
        string? name,
        string? workingDirectory,
        string? command,
        CancellationToken cancellationToken)
    {
        var proxy = await AcquireProxyAsync(cancellationToken).ConfigureAwait(false);
        if (proxy == null)
        {
            return TerminalCreateOutcome.Unavailable(UnavailableMessage);
        }

        using (proxy)
        {
            var developerEnvironment = false;
            object? profile = null;

            if (!string.IsNullOrWhiteSpace(command))
            {
                profile = BuildCommandProfile(command!, out developerEnvironment);
                if (profile == null)
                {
                    return TerminalCreateOutcome.Unavailable(UnavailableMessage);
                }
            }

            // Four CreateTerminalAsync overloads exist; this selects the ProfileConfig one that
            // also takes a working directory. Passing a null profile leaves Visual Studio to use
            // the user's default terminal profile.
            var method = _serviceType!.GetMethods()
                .FirstOrDefault(m => m.Name == "CreateTerminalAsync" && MatchesCreateSignature(m));
            if (method == null)
            {
                return TerminalCreateOutcome.Unavailable(UnavailableMessage);
            }

            var result = await InvokeAsync(
                method,
                proxy.Service,
                new[] { cancellationToken, name, profile, workingDirectory }).ConfigureAwait(false);

            return result is Guid id
                ? TerminalCreateOutcome.Success(id, developerEnvironment)
                : TerminalCreateOutcome.Unavailable("Visual Studio did not return a terminal identifier.");
        }
    }

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

            return result is IEnumerable ids
                ? ids.Cast<Guid>().ToList()
                : null;
        }
    }

    internal Task<bool> ShowTerminalAsync(Guid terminalId, CancellationToken cancellationToken) =>
        InvokeByIdAsync("ShowAsync", terminalId, cancellationToken);

    internal Task<bool> CloseTerminalAsync(Guid terminalId, CancellationToken cancellationToken) =>
        InvokeByIdAsync("CloseAsync", terminalId, cancellationToken);

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
            return true;
        }
    }

    private async Task<bool> InvokeByIdAsync(string methodName, Guid terminalId, CancellationToken cancellationToken)
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
    /// Builds a ProfileConfig that runs <paramref name="command"/> through cmd.exe. When the
    /// Visual Studio developer command script can be located it is chained in first, so the
    /// command sees the same PATH and environment as the Developer Command Prompt - which is
    /// what tooling like msbuild, vstest.console and dotnet-coverage expects.
    /// </summary>
    private object? BuildCommandProfile(string command, out bool developerEnvironment)
    {
        developerEnvironment = false;

        var profile = Activator.CreateInstance(_profileConfigType!);
        if (profile == null)
        {
            return null;
        }

        var comSpec = Environment.GetEnvironmentVariable("COMSPEC");
        if (string.IsNullOrWhiteSpace(comSpec))
        {
            comSpec = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.System),
                "cmd.exe");
        }

        var devCmd = TryFindDeveloperCommandScript();
        string arguments;
        if (devCmd != null)
        {
            developerEnvironment = true;
            // The outer pair of quotes is what cmd.exe /k requires when the command line itself
            // contains quoted paths.
            arguments = $"/k \"\"{devCmd}\" && {command}\"";
        }
        else
        {
            arguments = $"/k {command}";
        }

        SetProperty(profile, "Location", comSpec);
        SetProperty(profile, "Arguments", arguments);
        SetProperty(profile, "CreatePTY", true);

        return profile;
    }

    /// <summary>
    /// Locates VsDevCmd.bat for the running instance. VSAPPIDDIR points at Common7\IDE inside
    /// devenv, and the script sits alongside in Common7\Tools.
    /// </summary>
    internal static string? TryFindDeveloperCommandScript()
    {
        try
        {
            var ideDirectory = Environment.GetEnvironmentVariable("VSAPPIDDIR");
            if (string.IsNullOrWhiteSpace(ideDirectory))
            {
                return null;
            }

            var script = Path.GetFullPath(
                Path.Combine(ideDirectory, "..", "Tools", "VsDevCmd.bat"));

            return File.Exists(script) ? script : null;
        }
        catch (Exception ex)
        {
            VsixTelemetry.TrackException(ex);
            return null;
        }
    }

    private static bool MatchesCreateSignature(MethodInfo method)
    {
        var parameters = method.GetParameters();
        return parameters.Length == 4
            && parameters[0].ParameterType == typeof(CancellationToken)
            && parameters[1].ParameterType == typeof(string)
            && parameters[2].ParameterType.FullName == ProfileConfigTypeName
            && parameters[3].ParameterType == typeof(string);
    }

    private static void SetProperty(object target, string name, object value) =>
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
    /// Acquires a brokered-service proxy to the terminal. The proxy is created per operation and
    /// disposed with it; it is an RPC channel, so disposing it does not affect terminals that
    /// were opened through it.
    /// </summary>
    private async Task<TerminalProxy?> AcquireProxyAsync(CancellationToken cancellationToken)
    {
        try
        {
            if (!TryResolveTypes())
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
                new object?[] { _descriptor, default(ServiceActivationOptions), cancellationToken });

            // GetProxyAsync returns ValueTask<T>, which has to be turned into a Task before it
            // can be awaited through reflection.
            var asTask = pending?.GetType().GetMethod("AsTask", Type.EmptyTypes);
            if (asTask?.Invoke(pending, null) is not Task task)
            {
                return null;
            }

            await task.ConfigureAwait(false);

            var service = task.GetType().GetProperty("Result")?.GetValue(task);
            return service == null ? null : new TerminalProxy(service);
        }
        catch (Exception ex)
        {
            VsixTelemetry.TrackException(ex);
            return null;
        }
    }

    private bool TryResolveTypes()
    {
        if (_typesResolved)
        {
            return _serviceType != null && _profileConfigType != null && _descriptor != null;
        }

        _typesResolved = true;

        var assembly = _assemblyResolver();
        if (assembly == null)
        {
            return false;
        }

        _serviceType = assembly.GetType(ServiceTypeName);
        _profileConfigType = assembly.GetType(ProfileConfigTypeName);

        var descriptors = assembly.GetType(DescriptorsTypeName);
        _descriptor = descriptors
            ?.GetProperty("TerminalServiceDescriptor", BindingFlags.Public | BindingFlags.Static)
            ?.GetValue(null) as ServiceRpcDescriptor;

        return _serviceType != null && _profileConfigType != null && _descriptor != null;
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

    /// <summary>Owns the lifetime of a brokered-service proxy.</summary>
    private sealed class TerminalProxy : IDisposable
    {
        internal TerminalProxy(object service)
        {
            Service = service;
        }

        internal object Service { get; }

        public void Dispose() => (Service as IDisposable)?.Dispose();
    }
}

/// <summary>Result of a terminal creation attempt, before it is shaped for the wire.</summary>
internal sealed class TerminalCreateOutcome
{
    private TerminalCreateOutcome()
    {
    }

    internal bool Created { get; private set; }

    internal Guid TerminalId { get; private set; }

    internal bool DeveloperEnvironment { get; private set; }

    internal string? Message { get; private set; }

    internal static TerminalCreateOutcome Success(Guid id, bool developerEnvironment) => new()
    {
        Created = true,
        TerminalId = id,
        DeveloperEnvironment = developerEnvironment
    };

    internal static TerminalCreateOutcome Unavailable(string message) => new() { Message = message };
}
