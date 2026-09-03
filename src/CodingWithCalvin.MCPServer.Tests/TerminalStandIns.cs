using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.ServiceHub.Framework;

// Stand-ins for the Visual Studio integrated terminal types, declared in the real namespace so
// that TerminalInterop's reflective type lookups resolve against them.
//
// Microsoft.VisualStudio.Terminal.dll ships only inside the Visual Studio installation and is
// absent from NuGet and from the Microsoft.VisualStudio.SDK metapackage, so the production code
// loads it by strong name and reads these types reflectively. Supplying the test assembly as the
// assembly resolver lets that whole path run without Visual Studio.
//
// Shapes are transcribed from the shipped assembly and are identical in VS 2022 17.14 and
// VS 2026 18.0. Only the members the interop touches are reproduced - including both
// CreateTerminalAsync overloads that differ solely by the trailing working-directory parameter,
// so the overload selection is genuinely exercised.
namespace Microsoft.VisualStudio.Terminal;

public class ProfileConfig
{
    public string? Id { get; set; }
    public string? DisplayName { get; set; }
    public string? Location { get; set; }
    public string? Arguments { get; set; }
    public bool IsDefault { get; set; }
    public bool CreatePTY { get; set; }
}

public interface ITerminalService
{
    Task<Guid> CreateTerminalAsync(CancellationToken cancellationToken, string? name, ProfileConfig? profile);

    Task<Guid> CreateTerminalAsync(
        CancellationToken cancellationToken,
        string? name,
        ProfileConfig? profile,
        string? workingDirectory);

    Task<IEnumerable<Guid>> GetTerminalGuidsAsync(CancellationToken cancellationToken);

    Task ShowAsync(Guid terminalGuid, CancellationToken cancellationToken);

    Task CloseAsync(Guid terminalGuid, CancellationToken cancellationToken);

    Task CloseAllInstancesAsync(CancellationToken cancellationToken);
}

public static class TerminalServiceDescriptors
{
    public const string TerminalMoniker = "Microsoft.VisualStudio.Terminal.TerminalService";

    public static ServiceRpcDescriptor TerminalServiceDescriptor { get; } =
        new ServiceJsonRpcDescriptor(
            new ServiceMoniker(TerminalMoniker),
            ServiceJsonRpcDescriptor.Formatters.UTF8,
            ServiceJsonRpcDescriptor.MessageDelimiters.HttpLikeHeaders);
}
