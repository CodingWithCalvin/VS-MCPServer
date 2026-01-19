using System.Collections.Generic;

namespace CodingWithCalvin.MCPServer.Shared.Models;

public enum SymbolKind
{
    Unknown,
    Namespace,
    Class,
    Struct,
    Interface,
    Enum,
    Delegate,
    Function,
    Property,
    Field,
    Event,
    Variable,
    Parameter,
    EnumMember,
    Constant
}

public class SymbolInfo
{
    public string Name { get; set; } = string.Empty;
    public string FullName { get; set; } = string.Empty;
    public SymbolKind Kind { get; set; }
    public string FilePath { get; set; } = string.Empty;
    public int StartLine { get; set; }
    public int StartColumn { get; set; }
    public int EndLine { get; set; }
    public int EndColumn { get; set; }
    public string ContainerName { get; set; } = string.Empty;
    public List<SymbolInfo> Children { get; set; } = new();
}

public class LocationInfo
{
    public string FilePath { get; set; } = string.Empty;
    public int Line { get; set; }
    public int Column { get; set; }
    public int EndLine { get; set; }
    public int EndColumn { get; set; }
    public string Preview { get; set; } = string.Empty;
}

public class WorkspaceSymbolResult
{
    public List<SymbolInfo> Symbols { get; set; } = new();
    public int TotalCount { get; set; }
    public bool Truncated { get; set; }
}

public class DefinitionResult
{
    public bool Found { get; set; }
    public List<LocationInfo> Definitions { get; set; } = new();
    public string SymbolName { get; set; } = string.Empty;
    public SymbolKind SymbolKind { get; set; }
}

public class ReferencesResult
{
    public bool Found { get; set; }
    public List<LocationInfo> References { get; set; } = new();
    public string SymbolName { get; set; } = string.Empty;
    public int TotalCount { get; set; }
    public bool Truncated { get; set; }
}
