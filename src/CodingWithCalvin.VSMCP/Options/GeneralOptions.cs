using System.ComponentModel;
using Community.VisualStudio.Toolkit;

namespace CodingWithCalvin.VSMCP.Options;

public class GeneralOptions : BaseOptionModel<GeneralOptions>
{
    [Category("General")]
    [DisplayName("Auto-start server")]
    [Description("Automatically start the MCP server when Visual Studio launches")]
    [DefaultValue(false)]
    public bool AutoStartServer { get; set; } = false;

    [Category("General")]
    [DisplayName("Server name")]
    [Description("The name of this MCP server as it appears to clients")]
    [DefaultValue("Visual Studio MCP")]
    public string ServerName { get; set; } = "Visual Studio MCP";

    [Category("Server")]
    [DisplayName("HTTP port")]
    [Description("The port number for the HTTP/SSE MCP server (default: 5050)")]
    [DefaultValue(5050)]
    public int HttpPort { get; set; } = 5050;
}
