namespace CodingWithCalvin.VSMCP.Shared.Models;

public class SolutionInfo
{
    public string Name { get; set; } = string.Empty;
    public string Path { get; set; } = string.Empty;
    public bool IsOpen { get; set; }
}

public class ProjectInfo
{
    public string Name { get; set; } = string.Empty;
    public string Path { get; set; } = string.Empty;
    public string Kind { get; set; } = string.Empty;
}
