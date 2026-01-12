using System;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using CodingWithCalvin.VSMCP.Commands;
using CodingWithCalvin.VSMCP.Options;
using CodingWithCalvin.VSMCP.Services;
using Community.VisualStudio.Toolkit;
using Microsoft.VisualStudio.Shell;

namespace CodingWithCalvin.VSMCP;

[PackageRegistration(UseManagedResourcesOnly = true, AllowsBackgroundLoading = true)]
[InstalledProductRegistration(VsixInfo.DisplayName, VsixInfo.Description, VsixInfo.Version)]
[ProvideMenuResource("Menus.ctmenu", 1)]
[Guid(VSCommandTableVsct.guidVSMCPPackageString)]
[ProvideOptionPage(typeof(GeneralOptionsPage), "VSMCP", "General", 0, 0, true)]
[ProvideAutoLoad(Microsoft.VisualStudio.Shell.Interop.UIContextGuids80.NoSolution, PackageAutoLoadFlags.BackgroundLoad)]
public sealed class VSMCPPackage : ToolkitPackage
{
    public static ServerProcessManager? ServerManager { get; private set; }
    public static RpcServer? RpcServer { get; private set; }
    public static VisualStudioService? VsService { get; private set; }

    protected override async Task InitializeAsync(CancellationToken cancellationToken, IProgress<ServiceProgressData> progress)
    {
        await JoinableTaskFactory.SwitchToMainThreadAsync(cancellationToken);

        // Initialize services
        VsService = new VisualStudioService(this);
        RpcServer = new RpcServer(VsService);
        ServerManager = new ServerProcessManager(this, RpcServer);

        // Register commands
        await ServerCommands.InitializeAsync(this);

        // Auto-start server if configured
        var options = await GeneralOptions.GetLiveInstanceAsync();
        if (options.AutoStartServer)
        {
            await ServerManager.StartAsync();
        }
    }

    protected override void Dispose(bool disposing)
    {
        if (disposing)
        {
            ServerManager?.StopAsync().FireAndForget();
            RpcServer?.Dispose();
        }

        base.Dispose(disposing);
    }
}
