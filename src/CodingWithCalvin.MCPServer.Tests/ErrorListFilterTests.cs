using System;
using System.Collections.Generic;
using CodingWithCalvin.MCPServer.Services;
using CodingWithCalvin.MCPServer.Shared.Models;
using Microsoft.VisualStudio.Shell.TableManager;
using Xunit;

namespace CodingWithCalvin.MCPServer.Tests;

public class ErrorListFilterTests
{
    private static readonly Guid AppProjectGuid = new("6f1c2a4e-0b7d-4c55-9a3e-2d8f6b1e7c90");
    private static readonly Guid TestsProjectGuid = new("a3b9e7d1-5c2f-4e86-8b14-9f0d3c6a2e57");

    [Fact]
    public void Matches_WithDefaultQuery_KeepsEveryEntry()
    {
        var filter = CreateFilter(new ErrorListQuery());

        Assert.True(Match(filter, CreateItem()));
        Assert.True(Match(filter, CreateItem(severity: "Message", source: "")));
    }

    [Theory]
    [InlineData("Error", "error", true)]
    [InlineData("Warning", "Error", false)]
    public void Matches_BySeverity_IsCaseInsensitive(string itemSeverity, string severity, bool expected)
    {
        var filter = CreateFilter(new ErrorListQuery { Severity = severity });

        Assert.Equal(expected, Match(filter, CreateItem(severity: itemSeverity)));
    }

    [Theory]
    [InlineData(ErrorListSource.All, "Build", true)]
    [InlineData(ErrorListSource.All, "", true)]
    [InlineData(ErrorListSource.Build, "Build", true)]
    [InlineData(ErrorListSource.Build, "IntelliSense", false)]
    [InlineData(ErrorListSource.Build, "", false)]
    [InlineData(ErrorListSource.IntelliSense, "IntelliSense", true)]
    [InlineData(ErrorListSource.IntelliSense, "Build", false)]
    [InlineData(ErrorListSource.IntelliSense, "", false)]
    public void Matches_BySource_ExcludesUnknownSourcesUnlessAll(
        ErrorListSource source,
        string itemSource,
        bool expected)
    {
        var filter = CreateFilter(new ErrorListQuery { Source = source });

        Assert.Equal(expected, Match(filter, CreateItem(source: itemSource)));
    }

    [Theory]
    [InlineData("CS0103", true)]
    [InlineData("cs8618", true)]
    [InlineData("CS0168", false)]
    [InlineData("", false)]
    public void Matches_ByCode_KeepsAnyListedCode(string itemCode, bool expected)
    {
        var filter = CreateFilter(new ErrorListQuery { Codes = new List<string> { "CS0103", "CS8618" } });

        Assert.Equal(expected, Match(filter, CreateItem(code: itemCode)));
    }

    [Theory]
    [InlineData("undefined", true)]
    [InlineData("CS01", true)]
    [InlineData("MYAPP", true)]
    [InlineData("program.cs", true)]
    [InlineData("fxcop", true)]
    [InlineData("naming", true)]
    [InlineData("nullable", false)]
    public void Matches_BySearch_LooksAcrossEveryTextColumn(string search, bool expected)
    {
        var filter = CreateFilter(new ErrorListQuery { Search = search });

        Assert.Equal(expected, Match(filter, CreateItem(tool: "FxCop", category: "Naming")));
    }

    [Theory]
    [InlineData(ErrorListSuppressionState.All, "Suppressed", true)]
    [InlineData(ErrorListSuppressionState.All, "", true)]
    [InlineData(ErrorListSuppressionState.Active, "Active", true)]
    [InlineData(ErrorListSuppressionState.Active, "Suppressed", false)]
    [InlineData(ErrorListSuppressionState.Suppressed, "Suppressed", true)]
    [InlineData(ErrorListSuppressionState.NotApplicable, "NotApplicable", true)]
    [InlineData(ErrorListSuppressionState.NotApplicable, "", false)]
    public void Matches_BySuppressionState_ExcludesUnknownStatesUnlessAll(
        ErrorListSuppressionState suppressionState,
        string itemSuppressionState,
        bool expected)
    {
        var filter = CreateFilter(new ErrorListQuery { SuppressionState = suppressionState });

        Assert.Equal(expected, Match(filter, CreateItem(suppressionState: itemSuppressionState)));
    }

    [Theory]
    [InlineData("Compiler", true)]
    [InlineData("fxcop", true)]
    [InlineData("Roslyn", false)]
    [InlineData("", false)]
    public void Matches_ByTool_KeepsAnyListedTool(string itemTool, bool expected)
    {
        var filter = CreateFilter(new ErrorListQuery { Tools = new List<string> { "compiler", "FxCop" } });

        Assert.Equal(expected, Match(filter, CreateItem(tool: itemTool)));
    }

    [Theory]
    [InlineData("Style", true)]
    [InlineData("Compiler", false)]
    [InlineData("", false)]
    public void Matches_ByCategory_KeepsAnyListedCategory(string itemCategory, bool expected)
    {
        var filter = CreateFilter(new ErrorListQuery { Categories = new List<string> { "style" } });

        Assert.Equal(expected, Match(filter, CreateItem(category: itemCategory)));
    }

    [Fact]
    public void Matches_ByProject_KeepsEntriesSharedWithThatProject()
    {
        var filter = CreateFilter(new ErrorListQuery { Project = "App.Android" });
        var item = CreateItem(project: "App.iOS, App.Android");

        Assert.True(filter.Matches(new ErrorListEntry(item, new[] { "App.iOS", "App.Android" }, Array.Empty<Guid>())));
        Assert.False(filter.Matches(new ErrorListEntry(item, new[] { "App.iOS", "App.Windows" }, Array.Empty<Guid>())));
    }

    [Fact]
    public void Matches_ByProject_WithoutAnyProject_MatchesNothing()
    {
        var filter = CreateFilter(new ErrorListQuery { Project = "MyApp" });

        Assert.False(Match(filter, CreateItem(project: "")));
    }

    [Fact]
    public void Matches_CurrentProjectScope_KeepsEntriesSharedWithTheCurrentProject()
    {
        var filter = CreateFilter(
            new ErrorListQuery { Scope = ErrorListScope.CurrentProject },
            new ErrorListScopeContext { CurrentProjectName = "MyApp", CurrentProjectGuid = AppProjectGuid });
        var names = new[] { "MyApp.Tests", "MyApp" };

        Assert.True(filter.Matches(new ErrorListEntry(CreateItem(), names, new[] { TestsProjectGuid, AppProjectGuid })));
        Assert.True(filter.Matches(new ErrorListEntry(CreateItem(), names, Array.Empty<Guid>())));
        Assert.False(filter.Matches(new ErrorListEntry(CreateItem(), names, new[] { TestsProjectGuid })));
    }

    [Fact]
    public void NeedsEntryDetails_IsFalseForADefaultOrSeverityOnlyQuery()
    {
        Assert.False(CreateFilter(new ErrorListQuery()).NeedsEntryDetails);
        Assert.False(CreateFilter(new ErrorListQuery { Severity = "Error" }).NeedsEntryDetails);
    }

    public static TheoryData<ErrorListQuery> DetailQueries => new()
    {
        new ErrorListQuery { Scope = ErrorListScope.CurrentDocument },
        new ErrorListQuery { Source = ErrorListSource.Build },
        new ErrorListQuery { SuppressionState = ErrorListSuppressionState.Active },
        new ErrorListQuery { Codes = new List<string> { "CS0103" } },
        new ErrorListQuery { Tools = new List<string> { "Compiler" } },
        new ErrorListQuery { Categories = new List<string> { "Style" } },
        new ErrorListQuery { Project = "MyApp" },
        new ErrorListQuery { Path = "Program.cs" },
        new ErrorListQuery { Search = "undefined" }
    };

    [Theory]
    [MemberData(nameof(DetailQueries))]
    public void NeedsEntryDetails_IsTrueForAnyOtherFilter(ErrorListQuery query)
    {
        Assert.True(CreateFilter(query).NeedsEntryDetails);
    }

    [Fact]
    public void Matches_WithSeveralFilters_RequiresAllOfThem()
    {
        var filter = CreateFilter(new ErrorListQuery
        {
            Severity = "Error",
            Source = ErrorListSource.Build,
            Codes = new List<string> { "CS0103" }
        });

        Assert.True(Match(filter, CreateItem()));
        Assert.False(Match(filter, CreateItem(severity: "Warning")));
        Assert.False(Match(filter, CreateItem(source: "IntelliSense")));
    }

    [Theory]
    [InlineData(@"C:\src\App\Program.cs", true)]
    [InlineData("c:/src/app/program.cs", true)]
    [InlineData(@"C:\src\App\Other.cs", false)]
    [InlineData("", false)]
    public void Matches_CurrentDocumentScope_ComparesFullPaths(string itemPath, bool expected)
    {
        var filter = CreateFilter(
            new ErrorListQuery { Scope = ErrorListScope.CurrentDocument },
            new ErrorListScopeContext { ActiveDocumentPath = @"C:\src\App\Program.cs" });

        Assert.Equal(expected, Match(filter, CreateItem(path: itemPath)));
    }

    [Fact]
    public void Matches_CurrentDocumentScope_WithoutActiveDocument_MatchesNothing()
    {
        var filter = CreateFilter(new ErrorListQuery { Scope = ErrorListScope.CurrentDocument });

        Assert.False(Match(filter, CreateItem()));
    }

    [Theory]
    [InlineData(@"C:\src\App\Program.cs", true)]
    [InlineData(@"c:\src\app\startup.cs", true)]
    [InlineData(@"C:\src\App\Other.cs", false)]
    [InlineData("", false)]
    public void Matches_OpenDocumentsScope_KeepsOpenFiles(string itemPath, bool expected)
    {
        var filter = CreateFilter(
            new ErrorListQuery { Scope = ErrorListScope.OpenDocuments },
            new ErrorListScopeContext
            {
                OpenDocumentPaths = new[] { @"C:\src\App\Program.cs", "C:/src/App/Startup.cs" }
            });

        Assert.Equal(expected, Match(filter, CreateItem(path: itemPath)));
    }

    [Fact]
    public void Matches_CurrentProjectScope_PrefersProjectGuidOverName()
    {
        var filter = CreateFilter(
            new ErrorListQuery { Scope = ErrorListScope.CurrentProject },
            new ErrorListScopeContext { CurrentProjectName = "MyApp", CurrentProjectGuid = AppProjectGuid });

        Assert.True(Match(filter, CreateItem(project: "Renamed"), AppProjectGuid));
        Assert.False(Match(filter, CreateItem(project: "MyApp"), TestsProjectGuid));
    }

    [Theory]
    [InlineData("MyApp", true)]
    [InlineData("MyApp (net8.0)", true)]
    [InlineData("MyApp.Tests", false)]
    [InlineData("", false)]
    public void Matches_CurrentProjectScope_FallsBackToNameWhenEntryHasNoGuid(string itemProject, bool expected)
    {
        var filter = CreateFilter(
            new ErrorListQuery { Scope = ErrorListScope.CurrentProject },
            new ErrorListScopeContext { CurrentProjectName = "MyApp", CurrentProjectGuid = AppProjectGuid });

        Assert.Equal(expected, Match(filter, CreateItem(project: itemProject)));
    }

    [Fact]
    public void Matches_CurrentProjectScope_WithoutCurrentProject_MatchesNothing()
    {
        var filter = CreateFilter(new ErrorListQuery { Scope = ErrorListScope.CurrentProject });

        Assert.False(Match(filter, CreateItem(), AppProjectGuid));
    }

    [Theory]
    [InlineData("MyApp", "myapp", true)]
    [InlineData("MyApp (net8.0)", "MyApp", true)]
    [InlineData("MyApp (net8.0)", "MyApp (net8.0)", true)]
    [InlineData("MyApp (net8.0)", " MyApp ", true)]
    [InlineData("MyApp.Tests", "MyApp", false)]
    [InlineData("MyApp (net8.0)", "MyAp", false)]
    [InlineData("", "MyApp", false)]
    public void MatchesProjectName_AcceptsTargetFrameworkSuffix(string entryProject, string project, bool expected)
    {
        Assert.Equal(expected, ErrorListFilter.MatchesProjectName(entryProject, project));
    }

    [Theory]
    [InlineData(@"C:\src\App\Program.cs", "Program.cs", true)]
    [InlineData(@"C:\src\App\Program.cs", "program.cs", true)]
    [InlineData(@"C:\src\App\Program.cs", "src/App", true)]
    [InlineData(@"C:\src\App\Program.cs", @"src\App\", true)]
    [InlineData(@"C:\src\App\Program.cs", "App/Program.cs", true)]
    [InlineData(@"C:\src\App\Program.cs", @"C:\src\App\Program.cs", true)]
    [InlineData(@"C:\src\App\Program.cs", "gram.cs", false)]
    [InlineData(@"C:\src\Application\Program.cs", "src/App", false)]
    [InlineData(@"C:\src\App\Program.cs", "/", false)]
    [InlineData("", "Program.cs", false)]
    public void MatchesPath_MatchesWholePathSegments(string entryPath, string path, bool expected)
    {
        Assert.Equal(expected, ErrorListFilter.MatchesPath(entryPath, path));
    }

    [Fact]
    public void DescribeHiddenEntries_WhenWindowShowsEverything_ReturnsNull()
    {
        Assert.Null(ErrorListFilter.DescribeHiddenEntries(true, true, true, true, true));
    }

    [Fact]
    public void DescribeHiddenEntries_NamesASingleHiddenCategory()
    {
        Assert.Equal(
            "The Error List window is hiding warnings, so they are not included. Change the window's filters to include them.",
            ErrorListFilter.DescribeHiddenEntries(true, false, true, true, true));
    }

    [Fact]
    public void DescribeHiddenEntries_ListsEveryHiddenCategory()
    {
        Assert.Equal(
            "The Error List window is hiding messages, build entries and IntelliSense entries, so they are not included. Change the window's filters to include them.",
            ErrorListFilter.DescribeHiddenEntries(true, true, false, false, false));
    }

    [Theory]
    [InlineData(ErrorSource.Build, "Build")]
    [InlineData(ErrorSource.Other, "IntelliSense")]
    [InlineData((ErrorSource)0, "")]
    public void GetErrorSourceName_UsesErrorListLabels(ErrorSource source, string expected)
    {
        Assert.Equal(expected, VisualStudioService.GetErrorSourceName(source));
    }

    [Theory]
    [InlineData(SuppressionState.Active, ErrorListSuppressionState.Active)]
    [InlineData(SuppressionState.Suppressed, ErrorListSuppressionState.Suppressed)]
    [InlineData(SuppressionState.NotApplicable, ErrorListSuppressionState.NotApplicable)]
    public void GetSuppressionStateName_MatchesTheQueryFilterValues(
        SuppressionState state,
        ErrorListSuppressionState expected)
    {
        var filter = CreateFilter(new ErrorListQuery { SuppressionState = expected });
        var item = CreateItem(suppressionState: VisualStudioService.GetSuppressionStateName(state));

        Assert.True(Match(filter, item));
    }

    [Fact]
    public void GetSuppressionStateName_WithUnknownState_ReturnsEmpty()
    {
        Assert.Equal(string.Empty, VisualStudioService.GetSuppressionStateName((SuppressionState)99));
    }

    private static ErrorListFilter CreateFilter(ErrorListQuery query, ErrorListScopeContext? scope = null) =>
        new(query, scope ?? new ErrorListScopeContext());

    /// <summary>Matches an entry belonging to the single project named in <see cref="ErrorItemInfo.Project"/>.</summary>
    private static bool Match(ErrorListFilter filter, ErrorItemInfo item, params Guid[] projectGuids) =>
        filter.Matches(new ErrorListEntry(
            item,
            string.IsNullOrEmpty(item.Project) ? Array.Empty<string>() : new[] { item.Project },
            projectGuids));

    private static ErrorItemInfo CreateItem(
        string severity = "Error",
        string source = "Build",
        string code = "CS0103",
        string project = "MyApp",
        string path = @"C:\src\App\Program.cs",
        string tool = "Compiler",
        string category = "Compiler",
        string suppressionState = "Active") => new()
    {
        Severity = severity,
        Source = source,
        ErrorCode = code,
        Project = project,
        FilePath = path,
        Tool = tool,
        Category = category,
        SuppressionState = suppressionState,
        Description = "The name 'undefined' does not exist in the current context"
    };
}
