using System;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing.Design;
using System.IO;
using System.Windows.Forms;
using System.Windows.Forms.Design;
using Microsoft.VisualStudio.Shell;

namespace CodingWithCalvin.MCPServer.Dialogs;

public enum LogLevel
{
    Error,
    Warning,
    Information,
    Debug
}

public class LogFolderEditor : UITypeEditor
{
    public override UITypeEditorEditStyle GetEditStyle(ITypeDescriptorContext? context) =>
        UITypeEditorEditStyle.Modal;

    public override object? EditValue(ITypeDescriptorContext? context, IServiceProvider? provider, object? value)
    {
        var path = value as string;
        if (!string.IsNullOrEmpty(path) && Directory.Exists(path))
        {
            Process.Start("explorer.exe", path);
        }
        else
        {
            MessageBox.Show($"Log folder does not exist yet.\n\nPath: {path}\n\nStart the server to create it.",
                "Log Folder", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        return value;
    }
}

public class SettingsDialogPage : DialogPage
{
    private const string DefaultBindingAddress = "localhost";
    private const int DefaultPort = 5050;
    private const string DefaultServerName = "Visual Studio MCP";
    private const LogLevel DefaultLogLevel = LogLevel.Information;
    private const int DefaultLogRetentionDays = 7;

    public static string LogFolderPath => Path.Combine(Path.GetTempPath(), "MCPServer");

    [Category("Server")]
    [DisplayName("Binding Address")]
    [Description("The address the MCP server binds to. Use 'localhost' for local only, '0.0.0.0' or '*' for all interfaces, or a specific IP. Requires server restart to take effect.")]
    public string BindingAddress { get; set; } = DefaultBindingAddress;

    [Category("Server")]
    [DisplayName("Port")]
    [Description("The HTTP port the MCP server listens on. Requires server restart to take effect.")]
    public int Port { get; set; } = DefaultPort;

    [Category("Server")]
    [DisplayName("Server Name")]
    [Description("The name reported by the MCP server to connected clients.")]
    public string ServerName { get; set; } = DefaultServerName;

    [Category("Logging")]
    [DisplayName("Log Level")]
    [Description("The minimum log level for server output. Requires server restart to take effect.")]
    [TypeConverter(typeof(EnumConverter))]
    public LogLevel LogLevel { get; set; } = DefaultLogLevel;

    [Category("Logging")]
    [DisplayName("Log Retention (Days)")]
    [Description("Number of days to keep log files. Log files older than this will be automatically deleted. Set to 0 to keep all logs.")]
    public int LogRetentionDays { get; set; } = DefaultLogRetentionDays;

    [Category("Logging")]
    [DisplayName("Log Folder")]
    [Description("Click the '...' button to open the log folder in Explorer.")]
    [Editor(typeof(LogFolderEditor), typeof(UITypeEditor))]
    public string LogFolder => LogFolderPath;

    [Category("Startup")]
    [DisplayName("Auto-start Server")]
    [Description("Automatically start the MCP server when Visual Studio launches.")]
    public bool AutoStart { get; set; } = false;

    public override void LoadSettingsFromStorage()
    {
        base.LoadSettingsFromStorage();

        // Ensure defaults if not set
        if (string.IsNullOrEmpty(BindingAddress))
        {
            BindingAddress = DefaultBindingAddress;
        }

        if (Port <= 0)
        {
            Port = DefaultPort;
        }

        if (string.IsNullOrEmpty(ServerName))
        {
            ServerName = DefaultServerName;
        }

        if (LogRetentionDays < 0)
        {
            LogRetentionDays = DefaultLogRetentionDays;
        }
    }
}
