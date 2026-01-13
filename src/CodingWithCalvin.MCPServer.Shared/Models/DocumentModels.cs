namespace CodingWithCalvin.MCPServer.Shared.Models;

public class DocumentInfo
{
    public string Name { get; set; } = string.Empty;
    public string Path { get; set; } = string.Empty;
    public bool IsSaved { get; set; }
}

public class SelectionInfo
{
    public string Text { get; set; } = string.Empty;
    public int StartLine { get; set; }
    public int StartColumn { get; set; }
    public int EndLine { get; set; }
    public int EndColumn { get; set; }
    public string DocumentPath { get; set; } = string.Empty;
}

public class FindResult
{
    public int Line { get; set; }
    public int Column { get; set; }
    public string Text { get; set; } = string.Empty;
    public string DocumentPath { get; set; } = string.Empty;
}
