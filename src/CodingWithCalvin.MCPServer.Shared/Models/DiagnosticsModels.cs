using System.Collections.Generic;

namespace CodingWithCalvin.MCPServer.Shared.Models;

/// <summary>
/// Which documents an Error List query covers, mirroring the scope drop-down in the Error List window.
/// </summary>
public enum ErrorListScope
{
    Solution,
    OpenDocuments,
    CurrentProject,
    CurrentDocument
}

/// <summary>
/// Which producer an Error List entry must come from, mirroring the Build + IntelliSense drop-down.
/// </summary>
public enum ErrorListSource
{
    All,
    Build,
    IntelliSense
}

/// <summary>
/// Which suppression state an Error List entry must have, mirroring the Suppression State column.
/// </summary>
public enum ErrorListSuppressionState
{
    All,
    Active,
    Suppressed,
    NotApplicable
}

/// <summary>
/// Filters applied to the entries the Error List window is currently showing. Every filter that is
/// set must match for an entry to be returned.
/// </summary>
public class ErrorListQuery
{
    /// <summary>"Error", "Warning", or "Message"; null for all.</summary>
    public string? Severity { get; set; }
    public ErrorListScope Scope { get; set; } = ErrorListScope.Solution;
    public ErrorListSource Source { get; set; } = ErrorListSource.All;
    public ErrorListSuppressionState SuppressionState { get; set; } = ErrorListSuppressionState.All;

    /// <summary>Error codes to keep (e.g. "CS0103"); empty for all.</summary>
    public List<string> Codes { get; set; } = new();

    /// <summary>Tool names to keep, as shown in the Tool column; empty for all.</summary>
    public List<string> Tools { get; set; } = new();

    /// <summary>Categories to keep, as shown in the Category column; empty for all.</summary>
    public List<string> Categories { get; set; } = new();
    public string? Project { get; set; }

    /// <summary>File name, folder, or partial path, matched on whole path segments.</summary>
    public string? Path { get; set; }

    /// <summary>Text to find in the description, code, project, file path, tool, or category.</summary>
    public string? Search { get; set; }
    public int MaxResults { get; set; } = 100;
}

public class ErrorListResult
{
    public List<ErrorItemInfo> Items { get; set; } = new();
    public int TotalCount { get; set; }
    public int ErrorCount { get; set; }
    public int WarningCount { get; set; }
    public int MessageCount { get; set; }

    /// <summary>
    /// Entries that passed every filter, including any cut off by the result limit. The other counts
    /// cover the whole Error List.
    /// </summary>
    public int MatchedCount { get; set; }
    public bool Truncated { get; set; }

    /// <summary>
    /// Caveats about the results, such as categories the Error List window itself is hiding.
    /// </summary>
    public List<string> Notes { get; set; } = new();
}

public class ErrorItemInfo
{
    public string Severity { get; set; } = string.Empty;  // "Error", "Warning", "Message"
    public string Description { get; set; } = string.Empty;
    public string ErrorCode { get; set; } = string.Empty;  // e.g., "CS0103"
    public string Project { get; set; } = string.Empty;  // comma-separated when the entry belongs to several projects
    public string FilePath { get; set; } = string.Empty;
    public int Line { get; set; }
    public int Column { get; set; }
    public string Source { get; set; } = string.Empty;  // "Build", "IntelliSense", or empty when unknown
    public string Tool { get; set; } = string.Empty;  // e.g., "Compiler"
    public string Category { get; set; } = string.Empty;  // e.g., "Style"
    public string SuppressionState { get; set; } = string.Empty;  // "Active", "Suppressed", "NotApplicable", or empty when unknown
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
