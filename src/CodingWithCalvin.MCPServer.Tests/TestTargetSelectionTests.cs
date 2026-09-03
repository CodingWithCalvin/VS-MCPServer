using System.Collections.Generic;
using CodingWithCalvin.MCPServer.Services;
using CodingWithCalvin.MCPServer.Shared.Models;
using Xunit;

namespace CodingWithCalvin.MCPServer.Tests;

/// <summary>
/// Covers how a caller-supplied test name is resolved to the symbol the caret gets placed on.
/// Test Explorer's run-in-context commands act on the caret, so picking the wrong symbol here
/// silently runs the wrong tests rather than failing.
/// </summary>
public class TestTargetSelectionTests
{
    [Fact]
    public void SelectTestTarget_PrefersExactFullNameOverSimpleName()
    {
        var candidates = new List<SymbolInfo>
        {
            Symbol("CalculatorTests", "Other.CalculatorTests", SymbolKind.Class),
            Symbol("CalculatorTests", "MyApp.Tests.CalculatorTests", SymbolKind.Class)
        };

        var selected = VisualStudioService.SelectTestTarget(candidates, "MyApp.Tests.CalculatorTests");

        Assert.Equal("MyApp.Tests.CalculatorTests", selected?.FullName);
    }

    [Fact]
    public void SelectTestTarget_MatchesSimpleName()
    {
        var candidates = new List<SymbolInfo>
        {
            Symbol("Helper", "MyApp.Tests.Helper", SymbolKind.Class),
            Symbol("Add_ReturnsSum", "MyApp.Tests.CalculatorTests.Add_ReturnsSum", SymbolKind.Function)
        };

        var selected = VisualStudioService.SelectTestTarget(candidates, "Add_ReturnsSum");

        Assert.Equal("MyApp.Tests.CalculatorTests.Add_ReturnsSum", selected?.FullName);
    }

    [Fact]
    public void SelectTestTarget_FallsBackToTrailingSegmentMatch()
    {
        var candidates = new List<SymbolInfo>
        {
            Symbol("Unrelated", "MyApp.Tests.Unrelated", SymbolKind.Class),
            Symbol("Add_ReturnsSum", "MyApp.Tests.CalculatorTests.Add_ReturnsSum", SymbolKind.Function)
        };

        var selected = VisualStudioService.SelectTestTarget(
            candidates,
            "CalculatorTests.Add_ReturnsSum");

        Assert.Equal("MyApp.Tests.CalculatorTests.Add_ReturnsSum", selected?.FullName);
    }

    [Fact]
    public void SelectTestTarget_PrefersExactSimpleNameOverTrailingSegment()
    {
        var candidates = new List<SymbolInfo>
        {
            Symbol("Nested", "MyApp.Tests.Outer.Nested", SymbolKind.Class),
            Symbol("Nested", "MyApp.Tests.Nested", SymbolKind.Class)
        };

        var selected = VisualStudioService.SelectTestTarget(candidates, "Nested");

        // Both end with ".Nested"; the exact simple-name rule runs first and takes the earlier
        // candidate, so the result is deterministic rather than dependent on suffix ordering.
        Assert.Equal("MyApp.Tests.Outer.Nested", selected?.FullName);
    }

    [Fact]
    public void SelectTestTarget_ReturnsNull_WhenNoCandidates()
    {
        var selected = VisualStudioService.SelectTestTarget(new List<SymbolInfo>(), "Anything");

        Assert.Null(selected);
    }

    [Fact]
    public void SelectTestTarget_FallsBackToFirstCandidate_WhenNothingMatches()
    {
        var candidates = new List<SymbolInfo>
        {
            Symbol("CalculatorTests", "MyApp.Tests.CalculatorTests", SymbolKind.Class)
        };

        // The caller's text already matched during the workspace symbol search, so a candidate
        // that fails the stricter rules here is still a better answer than nothing.
        var selected = VisualStudioService.SelectTestTarget(candidates, "Calculator");

        Assert.Equal("MyApp.Tests.CalculatorTests", selected?.FullName);
    }

    private static SymbolInfo Symbol(string name, string fullName, SymbolKind kind) => new()
    {
        Name = name,
        FullName = fullName,
        Kind = kind,
        FilePath = @"C:\src\Tests.cs",
        StartLine = 10,
        StartColumn = 5
    };
}
