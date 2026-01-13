using System;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using CodingWithCalvin.Otel4Vsix;
using CodingWithCalvin.MCPServer.Commands;
using CodingWithCalvin.MCPServer.Dialogs;
using CodingWithCalvin.MCPServer.Services;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.ComponentModelHost;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;

namespace CodingWithCalvin.MCPServer;

[PackageRegistration(UseManagedResourcesOnly = true, AllowsBackgroundLoading = true)]
[InstalledProductRegistration(VsixInfo.DisplayName, VsixInfo.Description, VsixInfo.Version)]
[ProvideOptionPage(
    typeof(SettingsDialogPage),
    "MCP Server",
    "General",
    101,
    111,
    true,
    new string[0],
    ProvidesLocalizedCategoryName = false
)]
[ProvideAutoLoad(VSConstants.UICONTEXT.ShellInitialized_string, PackageAutoLoadFlags.BackgroundLoad)]
[ProvideAutoLoad(VSConstants.UICONTEXT.NoSolution_string, PackageAutoLoadFlags.BackgroundLoad)]
[ProvideAutoLoad(VSConstants.UICONTEXT.SolutionExists_string, PackageAutoLoadFlags.BackgroundLoad)]
[ProvideAutoLoad(VSConstants.UICONTEXT.EmptySolution_string, PackageAutoLoadFlags.BackgroundLoad)]
[ProvideMenuResource("Menus.ctmenu", 1)]
[Guid(VSCommandTableVsct.guidMCPServerPackageString)]
public sealed class MCPServerPackage : AsyncPackage
{
    public static MCPServerPackage? Instance { get; private set; }
    public static IServerProcessManager? ServerManager { get; private set; }
    public static IRpcServer? RpcServer { get; private set; }
    public static IVisualStudioService? VsService { get; private set; }
    public static IOutputPaneService? OutputPaneService { get; private set; }
    public static SettingsDialogPage? Settings { get; private set; }

    private IComponentModel? _componentModel;

    protected override async Task InitializeAsync(CancellationToken cancellationToken, IProgress<ServiceProgressData> progress)
    {
        await base.InitializeAsync(cancellationToken, progress);
        await JoinableTaskFactory.SwitchToMainThreadAsync(cancellationToken);

        Instance = this;
        Settings = (SettingsDialogPage)GetDialogPage(typeof(SettingsDialogPage));

        // Get MEF component model
        _componentModel = await GetServiceAsync(typeof(SComponentModel)) as IComponentModel;

        // Initialize telemetry
        var builder = VsixTelemetry.Configure()
            .WithServiceName(VsixInfo.DisplayName)
            .WithServiceVersion(VsixInfo.Version)
            .WithVisualStudioAttributes(this)
            .WithEnvironmentAttributes();

#if !DEBUG
        builder
            .WithOtlpHttp("https://api.honeycomb.io")
            .WithHeader("x-honeycomb-team", HoneycombConfig.ApiKey);
#endif

        builder.Initialize();

        await ServerCommands.InitializeAsync(this);

        // Auto-start server if configured
        if (Settings.AutoStart)
        {
            InitializeServices();
            if (ServerManager != null)
            {
                // Capture settings on UI thread (including output pane), then start on background
                var startSettings = new Services.ServerStartSettings
                {
                    BindingAddress = Settings.BindingAddress,
                    Port = Settings.Port,
                    ServerName = Settings.ServerName,
                    LogLevel = Settings.LogLevel.ToString(),
                    LogRetentionDays = Settings.LogRetentionDays,
                    OutputPane = OutputPaneService?.GetPane()
                };
                _ = Task.Run(async () => await ServerManager.StartAsync(startSettings));
            }
        }
    }

    public void InitializeServices()
    {
        if (VsService == null && _componentModel != null)
        {
            VsService = _componentModel.GetService<IVisualStudioService>();
            RpcServer = _componentModel.GetService<IRpcServer>();
            ServerManager = _componentModel.GetService<IServerProcessManager>();
            OutputPaneService = _componentModel.GetService<IOutputPaneService>();
        }
    }

    protected override void Dispose(bool disposing)
    {
        if (disposing)
        {
            ServerManager?.StopAsync().GetAwaiter().GetResult();
            RpcServer?.Dispose();
            VsixTelemetry.Shutdown();
            Instance = null;
        }

        base.Dispose(disposing);
    }
}
