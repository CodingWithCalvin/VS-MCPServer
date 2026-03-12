namespace CodingWithCalvin.MCPServer.Shared.Models;

public class DebuggerStatus
{
    public string Mode { get; set; } = string.Empty;
    public bool IsDebugging { get; set; }
    public string LastBreakReason { get; set; } = string.Empty;
    public string CurrentProcessName { get; set; } = string.Empty;
    public string CurrentFile { get; set; } = string.Empty;
    public int CurrentLine { get; set; }
    public string CurrentFunction { get; set; } = string.Empty;
}

public class BreakpointInfo
{
    public string File { get; set; } = string.Empty;
    public int Line { get; set; }
    public int Column { get; set; }
    public string FunctionName { get; set; } = string.Empty;
    public string Condition { get; set; } = string.Empty;
    public bool Enabled { get; set; }
    public int CurrentHits { get; set; }
}

public class LocalVariableInfo
{
    public string Name { get; set; } = string.Empty;
    public string Value { get; set; } = string.Empty;
    public string Type { get; set; } = string.Empty;
    public bool IsValidValue { get; set; }
}

public class CallStackFrameInfo
{
    public int Depth { get; set; }
    public string FunctionName { get; set; } = string.Empty;
    public string FileName { get; set; } = string.Empty;
    public int LineNumber { get; set; }
    public string Module { get; set; } = string.Empty;
    public string Language { get; set; } = string.Empty;
    public string ReturnType { get; set; } = string.Empty;
}
