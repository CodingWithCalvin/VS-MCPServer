using System.Collections.Generic;

namespace CodingWithCalvin.MCPServer.Shared.Models;

public class BuildConfiguration
{
    public string Configuration { get; set; } = string.Empty;
    public string Platform { get; set; } = string.Empty;
}

public class BuildConfigurationInfo
{
    public string ActiveConfiguration { get; set; } = string.Empty;
    public string ActivePlatform { get; set; } = string.Empty;
    public List<BuildConfiguration> AvailableConfigurations { get; set; } = new();
}

public class BuildStatus
{
    public string State { get; set; } = string.Empty;
    public int FailedProjects { get; set; }
}
