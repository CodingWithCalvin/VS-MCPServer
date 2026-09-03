using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using CodingWithCalvin.MCPServer.Services;
using CodingWithCalvin.MCPServer.Shared.Models;
using Microsoft.CodeCoverage.IO;
using Microsoft.CodeCoverage.IO.Coverage;
using Xunit;

namespace CodingWithCalvin.MCPServer.Tests;

/// <summary>
/// Covers the coverage-report reader and the roll-up arithmetic built on top of it.
/// </summary>
/// <remarks>
/// <para>
/// <c>Microsoft.CodeCoverage.IO.dll</c> cannot be referenced at build time, so
/// <see cref="CoverageInterop"/> loads it from the Visual Studio installation and reads it
/// reflectively. These tests supply the test assembly as the resolver, against stand-ins
/// declared in the genuine namespaces.
/// </para>
/// <para>
/// The arithmetic gets the most attention deliberately. Module and function counts come straight
/// from the reader, but the class level is derived here by grouping functions on their declaring
/// type - the one place a wrong number would be produced silently rather than failing.
/// </para>
/// <para>
/// All tests live in one class because the stand-in reader is seeded through static state and
/// xUnit runs tests within a class sequentially.
/// </para>
/// </remarks>
public class CoverageInteropTests : IDisposable
{
    private static readonly Func<Assembly?> StandInAssembly = () => typeof(CoverageFileUtility).Assembly;

    private readonly string _coverageFile;

    public CoverageInteropTests()
    {
        // CoverageInterop checks the file exists before reading, so a real one is needed even
        // though the stand-in reader ignores its contents.
        _coverageFile = Path.Combine(Path.GetTempPath(), $"vsmcp-{Guid.NewGuid():N}.coverage");
        File.WriteAllBytes(_coverageFile, Array.Empty<byte>());
    }

    public void Dispose()
    {
        CoverageFileUtility.Fixture = null;

        try
        {
            File.Delete(_coverageFile);
        }
        catch (IOException)
        {
            // A leftover temp file is not worth failing a test over.
        }
    }

    [Fact]
    public void ReadReport_GroupsFunctionsByDeclaringType()
    {
        CoverageFileUtility.Fixture = Data(
            Module("MyApp.Core.dll",
                Function("Add", "MyApp.Core", "Calculator", covered: 6, notCovered: 2),
                Function("Subtract", "MyApp.Core", "Calculator", covered: 4, notCovered: 0),
                Function("Parse", "MyApp.Core", "Parser", covered: 1, notCovered: 9)));

        var report = Read(CoverageDetail.Class);

        var module = Assert.Single(report.Modules);
        Assert.Equal(2, module.Classes.Count);

        var calculator = module.Classes.Single(c => c.Name == "Calculator");
        Assert.Equal("MyApp.Core", calculator.Namespace);
        Assert.Equal(10, calculator.Summary.LinesCovered);
        Assert.Equal(2, calculator.Summary.LinesNotCovered);
        Assert.Equal(83.33, calculator.Summary.LineCoveragePercent);

        var parser = module.Classes.Single(c => c.Name == "Parser");
        Assert.Equal(10, parser.Summary.LinesCovered + parser.Summary.LinesNotCovered);
        Assert.Equal(10d, parser.Summary.LineCoveragePercent);
    }

    [Fact]
    public void ReadReport_RollsClassesUpIntoModuleAndOverallTotals()
    {
        CoverageFileUtility.Fixture = Data(
            Module("A.dll", Function("M1", "N", "C1", covered: 3, notCovered: 1)),
            Module("B.dll", Function("M2", "N", "C2", covered: 1, notCovered: 5)));

        var report = Read(CoverageDetail.Class);

        var moduleA = report.Modules.Single(m => m.Name == "A.dll");
        Assert.Equal(3, moduleA.Summary.LinesCovered);
        Assert.Equal(1, moduleA.Summary.LinesNotCovered);

        // 3 + 1 covered against 1 + 5 uncovered across both modules.
        Assert.Equal(4, report.Summary.LinesCovered);
        Assert.Equal(6, report.Summary.LinesNotCovered);
        Assert.Equal(40d, report.Summary.LineCoveragePercent);
    }

    [Fact]
    public void ReadReport_CountsPartiallyCoveredLinesAgainstTheTotal()
    {
        CoverageFileUtility.Fixture = Data(
            Module("A.dll", Function("M", "N", "C", covered: 8, notCovered: 1, partiallyCovered: 1)));

        var summary = Read(CoverageDetail.Class).Summary;

        // A partially covered line is not a covered line: 8 of 10, not 9 of 10 or 8 of 9.
        Assert.Equal(80d, summary.LineCoveragePercent);
    }

    [Fact]
    public void ReadReport_SummaryDetail_OmitsClasses()
    {
        CoverageFileUtility.Fixture = Data(
            Module("A.dll", Function("M", "N", "C", covered: 5, notCovered: 5)));

        var module = Assert.Single(Read(CoverageDetail.Summary).Modules);

        Assert.Empty(module.Classes);
        // Totals are still correct even though the breakdown is trimmed.
        Assert.Equal(50d, module.Summary.LineCoveragePercent);
    }

    [Fact]
    public void ReadReport_MethodDetail_IncludesEachMethod()
    {
        CoverageFileUtility.Fixture = Data(
            Module("A.dll",
                Function("Add", "N", "C", covered: 2, notCovered: 0),
                Function("Remove", "N", "C", covered: 0, notCovered: 4)));

        var klass = Assert.Single(Assert.Single(Read(CoverageDetail.Method).Modules).Classes);

        Assert.Equal(new[] { "Add", "Remove" }, klass.Methods.Select(m => m.Name).OrderBy(n => n));
        Assert.Equal(100d, klass.Methods.Single(m => m.Name == "Add").Summary.LineCoveragePercent);
        Assert.Equal(0d, klass.Methods.Single(m => m.Name == "Remove").Summary.LineCoveragePercent);
    }

    [Fact]
    public void ReadReport_ClassDetail_OmitsMethods()
    {
        CoverageFileUtility.Fixture = Data(
            Module("A.dll", Function("M", "N", "C", covered: 1, notCovered: 1)));

        Assert.Empty(Assert.Single(Assert.Single(Read(CoverageDetail.Class).Modules).Classes).Methods);
    }

    [Fact]
    public void ReadReport_FilterSelectsMatchingClassOnly()
    {
        CoverageFileUtility.Fixture = Data(
            Module("A.dll",
                Function("M", "N", "OrderService", covered: 5, notCovered: 0),
                Function("M", "N", "Parser", covered: 0, notCovered: 5)));

        var report = Read(CoverageDetail.Class, filter: "orderservice");

        var klass = Assert.Single(Assert.Single(report.Modules).Classes);
        Assert.Equal("OrderService", klass.Name);
        // Totals reflect only what survived the filter.
        Assert.Equal(100d, report.Summary.LineCoveragePercent);
    }

    [Fact]
    public void ReadReport_FilterMatchingModule_KeepsAllOfItsClasses()
    {
        CoverageFileUtility.Fixture = Data(
            Module("MyApp.Core.dll",
                Function("M", "N", "OrderService", covered: 1, notCovered: 0),
                Function("M", "N", "Parser", covered: 1, notCovered: 0)),
            Module("Other.dll", Function("M", "N", "Ignored", covered: 1, notCovered: 0)));

        var report = Read(CoverageDetail.Class, filter: "MyApp.Core");

        var module = Assert.Single(report.Modules);
        Assert.Equal("MyApp.Core.dll", module.Name);
        Assert.Equal(2, module.Classes.Count);
    }

    [Fact]
    public void ReadReport_FilterMatchingNothing_ExplainsWhyItIsEmpty()
    {
        CoverageFileUtility.Fixture = Data(
            Module("A.dll", Function("M", "N", "C", covered: 1, notCovered: 0)));

        var report = Read(CoverageDetail.Class, filter: "nonexistent");

        Assert.True(report.Available);
        Assert.Empty(report.Modules);
        Assert.Contains("nonexistent", report.Message);
    }

    [Fact]
    public void ReadReport_MissingFile_ReportsNotFound()
    {
        var missing = Path.Combine(Path.GetTempPath(), $"vsmcp-{Guid.NewGuid():N}.coverage");

        var report = new CoverageInterop(StandInAssembly).ReadReport(missing, CoverageDetail.Class, null);

        Assert.False(report.Available);
        Assert.Contains("not found", report.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void ReadReport_ReportsUnavailable_WhenReaderCannotBeLoaded()
    {
        var report = new CoverageInterop(() => null).ReadReport(_coverageFile, CoverageDetail.Class, null);

        Assert.False(report.Available);
        Assert.Equal(CoverageInterop.UnavailableMessage, report.Message);
    }

    [Fact]
    public void ReadReport_PassesTheRequestedPathToTheReader()
    {
        CoverageFileUtility.Fixture = Data(Module("A.dll"));

        Read(CoverageDetail.Class);

        Assert.Equal(_coverageFile, CoverageFileUtility.LastPathRead);
    }

    [Theory]
    [InlineData(0, 0, 0, 0d)]
    [InlineData(1, 0, 0, 100d)]
    [InlineData(0, 3, 0, 0d)]
    [InlineData(1, 2, 0, 33.33)]
    public void ApplyPercentages_HandlesEdgeCases(
        int covered,
        int notCovered,
        int partiallyCovered,
        double expected)
    {
        var summary = new CoverageSummary
        {
            LinesCovered = covered,
            LinesNotCovered = notCovered,
            LinesPartiallyCovered = partiallyCovered
        };

        CoverageInterop.ApplyPercentages(summary);

        // Zero lines must read as 0%, not a divide-by-zero or NaN.
        Assert.Equal(expected, summary.LineCoveragePercent);
    }

    [Fact]
    public void FindNewestCoverageFile_PicksTheMostRecentlyWritten()
    {
        var root = Path.Combine(Path.GetTempPath(), $"vsmcp-{Guid.NewGuid():N}");
        var nested = Path.Combine(root, "run-2", "In");
        Directory.CreateDirectory(nested);

        var older = Path.Combine(root, "older.coverage");
        var newer = Path.Combine(nested, "newer.coverage");
        File.WriteAllText(older, string.Empty);
        File.WriteAllText(newer, string.Empty);
        File.SetLastWriteTimeUtc(older, DateTime.UtcNow.AddHours(-1));

        try
        {
            // Visual Studio writes into a generated subdirectory, so the search has to recurse.
            Assert.Equal(newer, CoverageInterop.FindNewestCoverageFile(root));
        }
        finally
        {
            Directory.Delete(root, recursive: true);
        }
    }

    [Fact]
    public void FindNewestCoverageFile_ReturnsNull_WhenDirectoryAbsent()
    {
        var missing = Path.Combine(Path.GetTempPath(), $"vsmcp-{Guid.NewGuid():N}");

        Assert.Null(CoverageInterop.FindNewestCoverageFile(missing));
        Assert.Null(CoverageInterop.FindNewestCoverageFile(null));
    }

    [Fact]
    public void TryFindCoverageAssemblyPath_ReturnsNull_WhenIdeDirectoryUnknown()
    {
        var previous = Environment.GetEnvironmentVariable("VSAPPIDDIR");
        try
        {
            Environment.SetEnvironmentVariable("VSAPPIDDIR", null);
            Assert.Null(CoverageInterop.TryFindCoverageAssemblyPath());
        }
        finally
        {
            Environment.SetEnvironmentVariable("VSAPPIDDIR", previous);
        }
    }

    private CoverageReportResult Read(CoverageDetail detail, string? filter = null) =>
        new CoverageInterop(StandInAssembly).ReadReport(_coverageFile, detail, filter);

    private static CoverageData Data(params ModuleWrapper[] modules) =>
        new() { Modules = modules.ToList() };

    private static ModuleWrapper Module(string name, params Function[] functions) => new()
    {
        Name = name,
        Path = @"C:\src\bin\" + name,
        TargetFramework = "net8.0",
        Functions = functions.ToList()
    };

    private static Function Function(
        string name,
        string namespaceName,
        string typeName,
        uint covered,
        uint notCovered,
        uint partiallyCovered = 0) => new()
    {
        Name = name,
        NamespaceName = namespaceName,
        TypeName = typeName,
        LinesCovered = covered,
        LinesNotCovered = notCovered,
        LinesPartiallyCovered = partiallyCovered,
        BlocksCovered = covered,
        BlocksNotCovered = notCovered
    };
}
