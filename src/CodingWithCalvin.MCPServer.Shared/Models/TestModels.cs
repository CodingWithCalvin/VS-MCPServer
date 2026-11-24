using System.Collections.Generic;

namespace CodingWithCalvin.MCPServer.Shared.Models;

/// <summary>
/// Aggregate test counts as reported by the Test Explorer window.
/// </summary>
public class TestStats
{
    /// <summary>
    /// False when Test Explorer has not been initialized in this Visual Studio session.
    /// The counts are meaningless when this is false, and must not be read as "no tests".
    /// </summary>
    public bool Available { get; set; }

    public int Passed { get; set; }
    public int Failed { get; set; }
    public int Skipped { get; set; }
    public int NotRun { get; set; }
}

/// <summary>
/// Test Explorer run state, plus the counts as of the most recent state change.
/// </summary>
public class TestRunStatus
{
    /// <summary>
    /// One of "NoRunObserved", "Discovering", "Running", "Canceling", "Completed", or "Canceled".
    /// </summary>
    public string State { get; set; } = string.Empty;

    /// <summary>
    /// True while discovery or execution is still in flight.
    /// </summary>
    public bool IsRunning { get; set; }

    /// <summary>
    /// Raw Test Explorer operation name behind <see cref="State"/>, retained for diagnostics.
    /// Null until a state change has been observed.
    /// </summary>
    public string? LastOperation { get; set; }

    public TestStats Stats { get; set; } = new();
}

/// <summary>
/// Outcome of asking Test Explorer to start a run for a specific class or method.
/// </summary>
public class TestTargetResult
{
    public bool Started { get; set; }

    /// <summary>
    /// Fully qualified name of the symbol the caret was placed on, when one was resolved.
    /// </summary>
    public string? ResolvedTarget { get; set; }

    public string? FilePath { get; set; }
    public int Line { get; set; }

    /// <summary>
    /// Populated when <see cref="Started"/> is false, or when the resolved target was
    /// ambiguous and the first match was used.
    /// </summary>
    public string? Message { get; set; }
}

/// <summary>
/// Which tests <c>test_results</c> should return, and how much detail to include.
/// </summary>
public class TestResultsQuery
{
    /// <summary>
    /// "Failed", "NotRun", "Skipped", "Passed", or "All". Case-insensitive.
    /// </summary>
    public string Outcome { get; set; } = "Failed";

    /// <summary>
    /// Case-insensitive substring of the fully qualified test name; null for all.
    /// </summary>
    public string? Filter { get; set; }

    public int MaxResults { get; set; } = 50;

    /// <summary>
    /// When true, each result also carries its standard output and standard error.
    /// </summary>
    public bool IncludeOutput { get; set; }
}

/// <summary>
/// Individual test outcomes read from Test Explorer.
/// </summary>
public class TestResultsReport
{
    /// <summary>
    /// False when Test Explorer's results could not be read. <see cref="Message"/> says why, and
    /// an empty <see cref="Tests"/> list must then not be read as "no failures".
    /// </summary>
    public bool Available { get; set; }

    /// <summary>
    /// Run state, as reported by <c>test_status</c>, so that results read mid-run can be told apart
    /// from a finished run.
    /// </summary>
    public string State { get; set; } = string.Empty;

    /// <summary>
    /// Tests that matched the outcome and name filters, including any cut off by the result limit.
    /// </summary>
    public int TotalMatched { get; set; }

    /// <summary>
    /// True when <see cref="TotalMatched"/> is more than the number of tests returned.
    /// </summary>
    public bool Truncated { get; set; }

    public string? Message { get; set; }

    public List<TestResultEntry> Tests { get; set; } = new();
}

/// <summary>
/// One test as listed in Test Explorer.
/// </summary>
public class TestResultEntry
{
    public string FullyQualifiedName { get; set; } = string.Empty;
    public string DisplayName { get; set; } = string.Empty;

    /// <summary>
    /// Test Explorer state: "Failed", "Passed", "Skipped", "NotRun", "NotFound", or "None".
    /// </summary>
    public string Outcome { get; set; } = string.Empty;

    public long DurationMs { get; set; }

    /// <summary>
    /// True when the outcome is out of date: the code has changed since it ran, or Test Explorer
    /// restored it from an earlier Visual Studio session.
    /// </summary>
    public bool Stale { get; set; }

    public string? FilePath { get; set; }
    public int LineNumber { get; set; }

    /// <summary>
    /// Test assembly the test was discovered in.
    /// </summary>
    public string? Source { get; set; }

    /// <summary>
    /// One entry per result. Test Explorer usually lists each data row of a data-driven test as
    /// its own entry, sharing <see cref="FullyQualifiedName"/>, with the row's arguments in
    /// <see cref="DisplayName"/>; some adapters instead report the rows as several results here.
    /// </summary>
    public List<TestResultDetail> Results { get; set; } = new();
}

/// <summary>
/// One result of a test: a single execution, or a single data row of a data-driven test.
/// </summary>
public class TestResultDetail
{
    public string DisplayName { get; set; } = string.Empty;

    /// <summary>
    /// "Failed", "Passed", "Skipped", "NotRun", "NotFound", or "None".
    /// </summary>
    public string Outcome { get; set; } = string.Empty;

    public long DurationMs { get; set; }
    public string? ErrorMessage { get; set; }
    public string? ErrorStackTrace { get; set; }

    /// <summary>
    /// Only filled when the query asked for output.
    /// </summary>
    public string? StandardOutput { get; set; }

    /// <summary>
    /// Only filled when the query asked for output.
    /// </summary>
    public string? StandardError { get; set; }

    /// <summary>
    /// True when at least one of the text fields was cut to the size limit.
    /// </summary>
    public bool Truncated { get; set; }
}
