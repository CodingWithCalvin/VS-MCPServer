using System.Collections.Generic;

namespace CodingWithCalvin.MCPServer.Shared.Models;

public class TestExplorerTestListResult
{
	public bool Available { get; set; }
	public string? Message { get; set; }
	public int TotalCount { get; set; }
	public bool Truncated { get; set; }
	public List<TestExplorerTestInfo> Tests { get; set; } = new();
}

public class TestExplorerTestInfo
{
	public string Id { get; set; } = string.Empty;
	public string FullyQualifiedName { get; set; } = string.Empty;
	public string DisplayName { get; set; } = string.Empty;
	public string State { get; set; } = string.Empty;
	public string Source { get; set; } = string.Empty;
	public string? FilePath { get; set; }
	public int LineNumber { get; set; }
	public bool IsRunning { get; set; }
	public bool IsStale { get; set; }
}

public class TestExplorerTestResultInfo
{
	public string TestId { get; set; } = string.Empty;
	public string FullyQualifiedName { get; set; } = string.Empty;
	public string DisplayName { get; set; } = string.Empty;
	public string State { get; set; } = string.Empty;
	public int RunId { get; set; }
	public long DurationMilliseconds { get; set; }
	public string? ErrorMessage { get; set; }
	public string? ErrorStackTrace { get; set; }
	public string? StandardOutput { get; set; }
	public string? StandardError { get; set; }
	public string? Attachments { get; set; }
}

public class TestExplorerTestResultsResult
{
	public bool Available { get; set; }
	public string? Message { get; set; }
	public int TotalCount { get; set; }
	public bool Truncated { get; set; }
	public List<TestExplorerTestResultInfo> Results { get; set; } = new();
}

public class TestExplorerSelectionResult
{
	public bool Started { get; set; }
	public int SelectedCount { get; set; }
	public string? Message { get; set; }
}
