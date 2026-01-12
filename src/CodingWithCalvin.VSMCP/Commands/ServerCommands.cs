using System;
using System.ComponentModel.Design;
using System.Windows;
using CodingWithCalvin.VSMCP.Options;
using Community.VisualStudio.Toolkit;
using Microsoft.VisualStudio.Shell;
using Task = System.Threading.Tasks.Task;

namespace CodingWithCalvin.VSMCP.Commands;

internal sealed class ServerCommands
{
    public static async Task InitializeAsync(AsyncPackage package)
    {
        await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();

        var commandService = await VS.Services.GetCommandServiceAsync();
        if (commandService == null)
        {
            return;
        }

        // Start Server command
        var startCommandId = new CommandID(VSCommandTableVsct.guidVSMCPPackageCmdSet.Guid, VSCommandTableVsct.guidVSMCPPackageCmdSet.cmdidStartServer);
        var startCommand = new OleMenuCommand(OnStartServer, startCommandId);
        startCommand.BeforeQueryStatus += OnBeforeQueryStatusStart;
        commandService.AddCommand(startCommand);

        // Stop Server command
        var stopCommandId = new CommandID(VSCommandTableVsct.guidVSMCPPackageCmdSet.Guid, VSCommandTableVsct.guidVSMCPPackageCmdSet.cmdidStopServer);
        var stopCommand = new OleMenuCommand(OnStopServer, stopCommandId);
        stopCommand.BeforeQueryStatus += OnBeforeQueryStatusStop;
        commandService.AddCommand(stopCommand);

        // Restart Server command
        var restartCommandId = new CommandID(VSCommandTableVsct.guidVSMCPPackageCmdSet.Guid, VSCommandTableVsct.guidVSMCPPackageCmdSet.cmdidRestartServer);
        var restartCommand = new OleMenuCommand(OnRestartServer, restartCommandId);
        restartCommand.BeforeQueryStatus += OnBeforeQueryStatusStop;
        commandService.AddCommand(restartCommand);

        // Copy Server URL command
        var copyUrlCommandId = new CommandID(VSCommandTableVsct.guidVSMCPPackageCmdSet.Guid, VSCommandTableVsct.guidVSMCPPackageCmdSet.cmdidCopyServerUrl);
        var copyUrlCommand = new OleMenuCommand(OnCopyServerUrl, copyUrlCommandId);
        commandService.AddCommand(copyUrlCommand);
    }

    private static void OnBeforeQueryStatusStart(object sender, EventArgs e)
    {
        if (sender is OleMenuCommand command)
        {
            command.Enabled = VSMCPPackage.ServerManager != null && !VSMCPPackage.ServerManager.IsRunning;
        }
    }

    private static void OnBeforeQueryStatusStop(object sender, EventArgs e)
    {
        if (sender is OleMenuCommand command)
        {
            command.Enabled = VSMCPPackage.ServerManager != null && VSMCPPackage.ServerManager.IsRunning;
        }
    }

    private static void OnStartServer(object sender, EventArgs e)
    {
        ThreadHelper.JoinableTaskFactory.RunAsync(async () =>
        {
            if (VSMCPPackage.ServerManager != null)
            {
                await VSMCPPackage.ServerManager.StartAsync();
                await VS.StatusBar.ShowMessageAsync("VSMCP Server started");
            }
        }).FireAndForget();
    }

    private static void OnStopServer(object sender, EventArgs e)
    {
        ThreadHelper.JoinableTaskFactory.RunAsync(async () =>
        {
            if (VSMCPPackage.ServerManager != null)
            {
                await VSMCPPackage.ServerManager.StopAsync();
                await VS.StatusBar.ShowMessageAsync("VSMCP Server stopped");
            }
        }).FireAndForget();
    }

    private static void OnRestartServer(object sender, EventArgs e)
    {
        ThreadHelper.JoinableTaskFactory.RunAsync(async () =>
        {
            if (VSMCPPackage.ServerManager != null)
            {
                await VSMCPPackage.ServerManager.StopAsync();
                await VSMCPPackage.ServerManager.StartAsync();
                await VS.StatusBar.ShowMessageAsync("VSMCP Server restarted");
            }
        }).FireAndForget();
    }

    private static void OnCopyServerUrl(object sender, EventArgs e)
    {
        ThreadHelper.JoinableTaskFactory.RunAsync(async () =>
        {
            var options = await GeneralOptions.GetLiveInstanceAsync();
            var url = $"http://localhost:{options.HttpPort}";
            Clipboard.SetText(url);
            await VS.StatusBar.ShowMessageAsync($"Copied: {url}");
        }).FireAndForget();
    }
}
