using System;
using System.ComponentModel;
using System.Runtime.InteropServices;
using Microsoft.VisualStudio.Shell;

namespace CodingWithCalvin.VSMCP.Options;

[ComVisible(true)]
[Guid("c3d4e5f6-a7b8-9012-cdef-234567890123")]
public class GeneralOptionsPage : DialogPage
{
    private GeneralOptions _options = null!;

    protected override void OnActivate(CancelEventArgs e)
    {
        base.OnActivate(e);
        _options = GeneralOptions.Instance;
    }

    [Category("General")]
    [DisplayName("Auto-start server")]
    [Description("Automatically start the MCP server when Visual Studio launches")]
    public bool AutoStartServer
    {
        get => _options?.AutoStartServer ?? false;
        set
        {
            if (_options != null)
            {
                _options.AutoStartServer = value;
            }
        }
    }

    [Category("General")]
    [DisplayName("Server name")]
    [Description("The name of this MCP server as it appears to clients")]
    public string ServerName
    {
        get => _options?.ServerName ?? "Visual Studio MCP";
        set
        {
            if (_options != null)
            {
                _options.ServerName = value;
            }
        }
    }

    [Category("Server")]
    [DisplayName("HTTP port")]
    [Description("The port number for the HTTP/SSE MCP server (default: 5050)")]
    public int HttpPort
    {
        get => _options?.HttpPort ?? 5050;
        set
        {
            if (_options != null)
            {
                _options.HttpPort = value;
            }
        }
    }

    public override void SaveSettingsToStorage()
    {
        base.SaveSettingsToStorage();
        _options?.Save();
    }
}
