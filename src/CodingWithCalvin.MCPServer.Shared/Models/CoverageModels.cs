using System.Collections.Generic;

namespace CodingWithCalvin.MCPServer.Shared.Models;

/// <summary>
/// Covered and uncovered counts for one node of the coverage tree.
/// </summary>
public class CoverageSummary
{
    public int LinesCovered { get; set; }
    public int LinesPartiallyCovered { get; set; }
    public int LinesNotCovered { get; set; }
    public int BlocksCovered { get; set; }
    public int BlocksNotCovered { get; set; }

    /// <summary>
    /// Percentage of lines fully covered. Partially covered lines count against the total, which
    /// matches how Visual Studio reports line coverage.
    /// </summary>
    public double LineCoveragePercent { get; set; }

    public double BlockCoveragePercent { get; set; }
}

public class CoverageMethod
{
    public string Name { get; set; } = string.Empty;
    public CoverageSummary Summary { get; set; } = new();
}

public class CoverageClass
{
    public string Name { get; set; } = string.Empty;
    public string Namespace { get; set; } = string.Empty;
    public CoverageSummary Summary { get; set; } = new();

    /// <summary>Populated only when method-level detail is requested.</summary>
    public List<CoverageMethod> Methods { get; set; } = new();
}

public class CoverageModule
{
    public string Name { get; set; } = string.Empty;
    public string? Path { get; set; }
    public string? TargetFramework { get; set; }
    public CoverageSummary Summary { get; set; } = new();

    /// <summary>Populated only when class-level or method-level detail is requested.</summary>
    public List<CoverageClass> Classes { get; set; } = new();
}

/// <summary>
/// A parsed code coverage result set.
/// </summary>
public class CoverageReportResult
{
    /// <summary>
    /// False when no coverage data could be read. Distinct from a report whose counts are all
    /// zero, which would mean coverage ran but exercised nothing.
    /// </summary>
    public bool Available { get; set; }

    /// <summary>Path of the .coverage file the report was read from.</summary>
    public string? CoverageFile { get; set; }

    public CoverageSummary Summary { get; set; } = new();

    public List<CoverageModule> Modules { get; set; } = new();

    public string? Message { get; set; }
}

/// <summary>
/// Outcome of asking Visual Studio to start a coverage run.
/// </summary>
public class CoverageRunResult
{
    public bool Started { get; set; }

    /// <summary>
    /// False when this edition of Visual Studio has no code coverage support at all, as opposed
    /// to the command being momentarily disabled.
    /// </summary>
    public bool Supported { get; set; }

    public string? Message { get; set; }
}
