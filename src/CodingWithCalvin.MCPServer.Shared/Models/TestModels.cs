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
