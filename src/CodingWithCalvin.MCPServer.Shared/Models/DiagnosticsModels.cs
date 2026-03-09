using System.Collections.Generic;

namespace CodingWithCalvin.MCPServer.Shared.Models;

public class ErrorListResult
{
    public List<ErrorItemInfo> Items { get; set; } = new();
    public int TotalCount { get; set; }
    public int ErrorCount { get; set; }
    public int WarningCount { get; set; }
    public int MessageCount { get; set; }
    public bool Truncated { get; set; }
}

public class ErrorItemInfo
{
    public string Severity { get; set; } = string.Empty;  // "Error", "Warning", "Message"
    public string Description { get; set; } = string.Empty;
    public string ErrorCode { get; set; } = string.Empty;  // e.g., "CS0103"
    public string Project { get; set; } = string.Empty;
    public string FilePath { get; set; } = string.Empty;
    public int Line { get; set; }
    public int Column { get; set; }
}

public class OutputPaneInfo
{
    public string Name { get; set; } = string.Empty;
    public string Guid { get; set; } = string.Empty;
}

public class OutputReadResult
{
    public string Content { get; set; } = string.Empty;
    public string PaneName { get; set; } = string.Empty;
    public int LinesRead { get; set; }
}
