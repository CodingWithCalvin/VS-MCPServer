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

    /// <summary>
    /// Total time package disposal will spend shutting the server down before giving up and
    /// letting Visual Studio finish exiting.
    /// </summary>
    private static readonly TimeSpan ShutdownTimeout = TimeSpan.FromSeconds(5);

    /// <remarks>
    /// Visual Studio calls this on the UI thread. Shutdown work is therefore pushed onto the
    /// thread pool and waited on with a timeout: <see cref="Task.Run(Func{Task})"/> starts with
    /// no synchronization context, so no continuation can need the UI thread back, and the
    /// timeout bounds the damage if one ever does. Blocking the UI thread directly on
    /// <c>StopAsync</c> deadlocked and left devenv.exe resident after the main window closed
    /// (issue #97).
    /// </remarks>
    protected override void Dispose(bool disposing)
    {
        if (disposing)
        {
            var serverManager = ServerManager;

            if (serverManager != null)
            {
                try
                {
                    var stopTask = Task.Run(() => serverManager.StopAsync());

                    // VSTHRD002: Dispose cannot be async, so a blocking wait is unavoidable. It is
                    // safe here because Task.Run starts the work without a synchronization context
                    // and StopAsync uses ConfigureAwait(false) throughout, so no continuation can
                    // require this thread. The timeout guarantees VS exits regardless.
#pragma warning disable VSTHRD002
                    if (!stopTask.Wait(ShutdownTimeout))
#pragma warning restore VSTHRD002
                    {
                        // The job object assigned at start-up still guarantees the server
                        // process dies when devenv.exe does, so exiting is safe here.
                        System.Diagnostics.Debug.WriteLine("MCPServer: server shutdown timed out during package disposal.");
                    }
                }
                catch (Exception ex)
                {
                    // A failure to stop the server must never prevent Visual Studio from exiting.
                    System.Diagnostics.Debug.WriteLine($"MCPServer: error stopping server during package disposal: {ex}");
                }
            }

            try
            {
                RpcServer?.Dispose();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"MCPServer: error disposing RPC server: {ex}");
            }

            VsixTelemetry.Shutdown();

            ServerManager = null;
            RpcServer = null;
            VsService = null;
            OutputPaneService = null;
            Settings = null;
            Instance = null;
        }

        base.Dispose(disposing);
    }
}
