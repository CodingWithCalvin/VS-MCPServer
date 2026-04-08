namespace CodingWithCalvin.MCPServer.Shared.Models;

public class WindowInfo
{
    public string Caption { get; set; } = string.Empty;
    public string Kind { get; set; } = string.Empty;  // "Document" or "Tool"
    public bool IsVisible { get; set; }
    public string ObjectKind { get; set; } = string.Empty;  // Window GUID
}
