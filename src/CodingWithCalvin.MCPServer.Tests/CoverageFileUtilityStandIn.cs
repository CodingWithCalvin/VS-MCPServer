using Microsoft.CodeCoverage.IO.Coverage;

// Declared separately from the data types because it sits in the parent namespace, matching the
// shipped assembly where CoverageFileUtility is in Microsoft.CodeCoverage.IO and the data types
// are in Microsoft.CodeCoverage.IO.Coverage.
namespace Microsoft.CodeCoverage.IO;

/// <summary>
/// Stand-in for the real reader. Mirrors the V1 shape the production code targets: a
/// parameterless constructor and a ReadCoverageFile(string) overload.
/// </summary>
public class CoverageFileUtility
{
    /// <summary>
    /// Data the next read returns. Static because CoverageInterop constructs the utility itself
    /// via Activator, so a test has no other way to seed it. Tests within a class run
    /// sequentially, so this is safe here.
    /// </summary>
    public static CoverageData? Fixture { get; set; }

    public static string? LastPathRead { get; private set; }

    public CoverageData ReadCoverageFile(string path)
    {
        LastPathRead = path;
        return Fixture ?? new CoverageData();
    }

    // Present so the overload selection in CoverageInterop is exercised against a real ambiguity,
    // as it is on the shipped type.
    public CoverageData ReadCoverageFile(string path, bool readModules, bool readSkippedMessages)
        => throw new System.InvalidOperationException(
            "The single-path ReadCoverageFile overload should be selected.");
}
