using System;
using System.ComponentModel.Composition;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;

namespace CodingWithCalvin.MCPServer.Services;

[Export(typeof(IOutputPaneService))]
[PartCreationPolicy(CreationPolicy.Shared)]
public class OutputPaneService : IOutputPaneService
{
    private static readonly Guid OutputPaneGuid = new("A1B2C3D4-E5F6-4A5B-8C9D-0E1F2A3B4C5D");
    private const string OutputPaneTitle = "MCP Server";

    private IVsOutputWindowPane? _outputPane;
    private bool _initialized;

    public IVsOutputWindowPane? GetPane()
    {
        ThreadHelper.ThrowIfNotOnUIThread();

        if (_initialized)
        {
            return _outputPane;
        }

        _initialized = true;

        try
        {
            var outputWindow = Package.GetGlobalService(typeof(SVsOutputWindow)) as IVsOutputWindow;
            if (outputWindow != null)
            {
                var paneGuid = OutputPaneGuid;
                outputWindow.CreatePane(ref paneGuid, OutputPaneTitle, 1, 1);
                outputWindow.GetPane(ref paneGuid, out _outputPane);
            }
        }
        catch
        {
            // Output pane unavailable
            _outputPane = null;
        }

        return _outputPane;
    }
}
