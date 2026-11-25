using System;
using System.Collections.Generic;
using System.IO;
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
// ITerminalService and ITerminalProfile have the required members in both VS 2022 17.14 and
// VS 2026 18.0. TerminalServiceDescriptors represents the VS 2022 broker path, while
// SVsTerminalService and IVsTerminalService represent the VS 2026 direct-service path. Tests
// hide the other version's types to exercise each path independently.
namespace Microsoft.VisualStudio.Terminal;

public interface ITerminalProfile
{
    string? DisplayName { get; }
    string? Location { get; }
    string? Arguments { get; }
}

public class ProfileConfig : ITerminalProfile
{
    public ProfileConfig(string displayName, string location, string? arguments = null, bool isDefault = false)
    {
        DisplayName = displayName;
        Location = location;
        Arguments = arguments;
        IsDefault = isDefault;
    }

    public string? Id { get; set; }
    public string? DisplayName { get; set; }
    public string? Location { get; set; }
    public string? Arguments { get; set; }
    public bool IsDefault { get; set; }
    public bool CreatePTY => true;
}

public class TerminalWindowOptions
{
    public string? Name { get; set; }
    public ITerminalProfile? Profile { get; set; }
    public string? WorkingDirectory { get; set; }
}

public interface ITerminalService
{
    Task<Guid> CreateTerminalWindowAsync(CancellationToken cancellationToken, TerminalWindowOptions options);

    Task<Guid> CreateTerminalAsync(
        CancellationToken cancellationToken,
        string? name,
        ITerminalProfile? profile);

    Task<ITerminalProfile> GetDefaultProfileAsync(CancellationToken cancellationToken);

    Task AttachTerminalStreamAsync(Guid terminalGuid, Stream stream, CancellationToken cancellationToken);

    Task<IEnumerable<Guid>> GetTerminalGuidsAsync(CancellationToken cancellationToken);

    Task ShowAsync(Guid terminalGuid, CancellationToken cancellationToken);

    Task CloseAsync(Guid terminalGuid, CancellationToken cancellationToken);

    Task CloseAllInstancesAsync(CancellationToken cancellationToken);
}

public sealed class SVsTerminalService
{
}

public interface IVsTerminalService
{
    ITerminalService CreateTerminalService();
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
