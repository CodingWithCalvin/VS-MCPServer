using Microsoft.VisualStudio.Shell.Interop;

namespace CodingWithCalvin.MCPServer.Services;

public interface IOutputPaneService
{
    /// <summary>
    /// Gets or creates the MCP Server output pane. Must be called from UI thread.
    /// </summary>
    IVsOutputWindowPane? GetPane();
}
