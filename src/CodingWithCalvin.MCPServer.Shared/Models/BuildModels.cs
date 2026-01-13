namespace CodingWithCalvin.MCPServer.Shared.Models;

public class BuildStatus
{
    public string State { get; set; } = string.Empty;
    public int FailedProjects { get; set; }
}
