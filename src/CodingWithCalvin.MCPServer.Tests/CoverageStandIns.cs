using System.Collections.Generic;

// Stand-ins for the Visual Studio code coverage reader, declared in the real namespaces so that
// CoverageInterop's reflective lookups resolve against them.
//
// Microsoft.CodeCoverage.IO.dll ships only inside the Visual Studio installation, so the
// production code loads it from disk and reads it reflectively. Supplying the test assembly as
// the resolver lets that path run without Visual Studio.
//
// Shapes are transcribed from the shipped assembly: counts are uint and are inherited from
// CoverageStatistics by both ModuleWrapper's functions and the functions themselves, which is
// why no arithmetic is needed below the class level.
namespace Microsoft.CodeCoverage.IO.Coverage;

public class CoverageStatistics
{
    public uint BlocksCovered { get; set; }
    public uint BlocksNotCovered { get; set; }
    public uint LinesCovered { get; set; }
    public uint LinesPartiallyCovered { get; set; }
    public uint LinesNotCovered { get; set; }
}

public class Function : CoverageStatistics
{
    public string Name { get; set; } = string.Empty;
    public string NamespaceName { get; set; } = string.Empty;
    public string TypeName { get; set; } = string.Empty;
}

public class ModuleWrapper
{
    public string Name { get; set; } = string.Empty;
    public string Path { get; set; } = string.Empty;
    public string TargetFramework { get; set; } = string.Empty;
    public List<Function> Functions { get; set; } = new();
}

public class CoverageData
{
    public IList<ModuleWrapper> Modules { get; set; } = new List<ModuleWrapper>();
}
