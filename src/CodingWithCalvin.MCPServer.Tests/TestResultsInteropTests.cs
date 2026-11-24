using System;
using System.Collections.Generic;
using System.ComponentModel.Composition;
using System.ComponentModel.Composition.Hosting;
using System.ComponentModel.Composition.Primitives;
using System.Linq;
using System.Threading.Tasks;
using CodingWithCalvin.MCPServer.Services;
using CodingWithCalvin.MCPServer.Shared.Models;
using Microsoft.VisualStudio.ComponentModelHost;
using Microsoft.VisualStudio.TestWindow.Extensibility;
using Xunit;

namespace CodingWithCalvin.MCPServer.Tests;

/// <summary>
/// Covers the reflection bridge to Test Explorer's individual test results.
/// </summary>
/// <remarks>
/// <c>ITestsService</c>, <c>ITest</c> and <c>IResult</c> are internal to
/// Microsoft.VisualStudio.TestWindow.Interfaces.dll, so <see cref="TestResultsInterop"/> resolves
/// the service by MEF contract name and reads everything through the interface types. These
/// tests stand up a real MEF container holding a stand-in service declared in the genuine
/// namespace, with members implemented explicitly, so the contract-name lookup, the overload
/// selection and the interface-based property reads are all exercised.
/// </remarks>
public class TestResultsInteropTests
{
    [Fact]
    public async Task GetResults_ReportsUnavailable_WhenComponentModelMissing()
    {
        var report = await new TestResultsInterop(() => null).GetResultsAsync(new TestResultsQuery());

        Assert.False(report.Available);
        Assert.Equal(TestResultsInterop.ServiceNotFoundMessage, report.Message);
    }

    [Fact]
    public async Task GetResults_ReportsUnavailable_WhenServiceNotExported()
    {
        using var container = new CompositionContainer(new TypeCatalog(typeof(UnrelatedPart)));
        var interop = new TestResultsInterop(() => new StubComponentModel(container));

        var report = await interop.GetResultsAsync(new TestResultsQuery());

        Assert.False(report.Available);
        Assert.Equal(TestResultsInterop.ServiceNotFoundMessage, report.Message);
    }

    [Fact]
    public async Task GetResults_ReportsUnavailable_WhenServiceNotRunning()
    {
        using var host = new StubHost(Failed("A.B.C"));
        host.Service.IsRunning = false;

        var report = await host.Interop.GetResultsAsync(new TestResultsQuery());

        // An uninitialized Test Explorer must not read as "no failed tests".
        Assert.False(report.Available);
        Assert.Equal(TestResultsInterop.NotRunningMessage, report.Message);
        Assert.Empty(report.Tests);
    }

    [Fact]
    public async Task GetResults_MapsFailedTest_WithMessageStackFileAndLine()
    {
        var test = new StubTest
        {
            FullyQualifiedName = "Sample.Tests.CalculatorTests.Divide_ReturnsQuotient",
            DisplayName = "Divide_ReturnsQuotient",
            State = TestState.Failed,
            FilePath = @"C:\scratch\Sample.Tests\CalculatorTests.cs",
            LineNumber = 40,
            Source = @"C:\scratch\Sample.Tests\bin\Debug\net10.0\Sample.Tests.dll",
            Duration = TimeSpan.FromMilliseconds(12),
            Stale = true,
            Results =
            {
                new StubResult
                {
                    DisplayName = "Divide_ReturnsQuotient",
                    State = TestState.Failed,
                    Duration = TimeSpan.FromMilliseconds(11),
                    ErrorMessage = "Assert.AreEqual failed. Expected:<2>. Actual:<3>.",
                    ErrorStackTrace = "   at Sample.Tests.CalculatorTests.Divide_ReturnsQuotient() in CalculatorTests.cs:line 44"
                }
            }
        };
        using var host = new StubHost(test, Passed("Sample.Tests.CalculatorTests.Add_ReturnsSum"));

        var report = await host.Interop.GetResultsAsync(new TestResultsQuery());

        Assert.True(report.Available);
        Assert.Equal(1, report.TotalMatched);
        Assert.False(report.Truncated);

        var entry = Assert.Single(report.Tests);
        Assert.Equal("Sample.Tests.CalculatorTests.Divide_ReturnsQuotient", entry.FullyQualifiedName);
        Assert.Equal("Divide_ReturnsQuotient", entry.DisplayName);
        Assert.Equal("Failed", entry.Outcome);
        Assert.Equal(12, entry.DurationMs);
        Assert.True(entry.Stale);
        Assert.Equal(@"C:\scratch\Sample.Tests\CalculatorTests.cs", entry.FilePath);
        Assert.Equal(40, entry.LineNumber);
        Assert.Equal(@"C:\scratch\Sample.Tests\bin\Debug\net10.0\Sample.Tests.dll", entry.Source);

        var result = Assert.Single(entry.Results);
        Assert.Equal("Failed", result.Outcome);
        Assert.Equal(11, result.DurationMs);
        Assert.Equal("Assert.AreEqual failed. Expected:<2>. Actual:<3>.", result.ErrorMessage);
        Assert.Contains("line 44", result.ErrorStackTrace);
        Assert.False(result.Truncated);
    }

    [Theory]
    [InlineData("failed", new[] { "T.Failed" })]
    [InlineData("Passed", new[] { "T.Passed" })]
    [InlineData("SKIPPED", new[] { "T.Skipped" })]
    [InlineData("notrun", new[] { "T.Never", "T.NotRun" })]
    [InlineData("all", new[] { "T.Failed", "T.Missing", "T.Never", "T.NotRun", "T.Passed", "T.Skipped" })]
    public async Task GetResults_FiltersByOutcome(string outcome, string[] expected)
    {
        using var host = new StubHost(
            Failed("T.Failed"),
            Passed("T.Passed"),
            new StubTest { FullyQualifiedName = "T.Skipped", State = TestState.Skipped },
            new StubTest { FullyQualifiedName = "T.NotRun", State = TestState.NotRun },
            new StubTest { FullyQualifiedName = "T.Never", State = TestState.None },
            new StubTest { FullyQualifiedName = "T.Missing", State = TestState.NotFound });

        var report = await host.Interop.GetResultsAsync(new TestResultsQuery { Outcome = outcome });

        Assert.Equal(expected, report.Tests.Select(t => t.FullyQualifiedName));
        Assert.Equal(expected.Length, report.TotalMatched);
    }

    [Fact]
    public async Task GetResults_ReportsInvalidOutcome_AsUnavailable()
    {
        using var host = new StubHost(Failed("A.B.C"));

        var report = await host.Interop.GetResultsAsync(new TestResultsQuery { Outcome = "broken" });

        Assert.False(report.Available);
        Assert.Contains("Invalid outcome 'broken'", report.Message);
    }

    [Fact]
    public async Task GetResults_FiltersByName_CaseInsensitive()
    {
        using var host = new StubHost(
            Failed("App.Tests.CalculatorTests.Divide"),
            Failed("App.Tests.ParserTests.Parse"));

        var report = await host.Interop.GetResultsAsync(new TestResultsQuery { Filter = "calculatortests" });

        var entry = Assert.Single(report.Tests);
        Assert.Equal("App.Tests.CalculatorTests.Divide", entry.FullyQualifiedName);
        Assert.Equal(1, report.TotalMatched);
    }

    [Fact]
    public async Task GetResults_LimitsToMaxResults_AndCountsEveryMatch()
    {
        using var host = new StubHost(Failed("T.C"), Failed("T.A"), Failed("T.B"));

        var report = await host.Interop.GetResultsAsync(new TestResultsQuery { MaxResults = 2 });

        // Sorted by name, so the cut is predictable.
        Assert.Equal(new[] { "T.A", "T.B" }, report.Tests.Select(t => t.FullyQualifiedName));
        Assert.Equal(3, report.TotalMatched);
        Assert.True(report.Truncated);
    }

    [Fact]
    public async Task GetResults_CapsLongText_AndFlagsTheResult()
    {
        var longStack = new string('x', TestResultsInterop.MaxTextLength + 500);
        var test = Failed("T.Long");
        test.Results.Add(new StubResult { State = TestState.Failed, ErrorMessage = "short", ErrorStackTrace = longStack });
        using var host = new StubHost(test);

        var report = await host.Interop.GetResultsAsync(new TestResultsQuery());

        var result = Assert.Single(Assert.Single(report.Tests).Results);
        Assert.Equal("short", result.ErrorMessage);
        Assert.Equal(TestResultsInterop.MaxTextLength, result.ErrorStackTrace!.Length);
        Assert.True(result.Truncated);
    }

    [Fact]
    public async Task GetResults_IncludesOutput_OnlyWhenAsked()
    {
        var test = Failed("T.Output");
        test.Results.Add(new StubResult { State = TestState.Failed, StandardOutput = "out", StandardError = "err" });
        using var host = new StubHost(test);

        var without = await host.Interop.GetResultsAsync(new TestResultsQuery());
        var with = await host.Interop.GetResultsAsync(new TestResultsQuery { IncludeOutput = true });

        var withoutResult = Assert.Single(Assert.Single(without.Tests).Results);
        Assert.Null(withoutResult.StandardOutput);
        Assert.Null(withoutResult.StandardError);

        var withResult = Assert.Single(Assert.Single(with.Tests).Results);
        Assert.Equal("out", withResult.StandardOutput);
        Assert.Equal("err", withResult.StandardError);
    }

    [Fact]
    public async Task GetResults_KeepsEveryResult_ForDataDrivenTest()
    {
        var test = Failed("T.Add_DataRows");
        test.Results.Add(new StubResult { DisplayName = "Add_DataRows (1,1,2)", State = TestState.Passed });
        test.Results.Add(new StubResult { DisplayName = "Add_DataRows (2,2,5)", State = TestState.Failed, ErrorMessage = "Expected:<5>. Actual:<4>." });
        test.Results.Add(new StubResult { DisplayName = "Add_DataRows (3,3,6)", State = TestState.Passed });
        using var host = new StubHost(test);

        var report = await host.Interop.GetResultsAsync(new TestResultsQuery());

        var results = Assert.Single(report.Tests).Results;
        Assert.Equal(new[] { "Passed", "Failed", "Passed" }, results.Select(r => r.Outcome));
        Assert.Equal("Expected:<5>. Actual:<4>.", results[1].ErrorMessage);
        Assert.Equal("Add_DataRows (2,2,5)", results[1].DisplayName);
    }

    [Fact]
    public async Task GetResults_ListsDataRowsAsSeparateTests_InDisplayNameOrder()
    {
        // How VS 18.7 with MSTest 4 reports [DataRow] tests: one test per row, same name.
        using var host = new StubHost(
            new StubTest { FullyQualifiedName = "T.Add", DisplayName = "Add (3,3,6)", State = TestState.Passed },
            new StubTest { FullyQualifiedName = "T.Add", DisplayName = "Add (2,2,5)", State = TestState.Failed },
            new StubTest { FullyQualifiedName = "T.Add", DisplayName = "Add (1,1,2)", State = TestState.Passed });

        var report = await host.Interop.GetResultsAsync(new TestResultsQuery { Outcome = "all" });

        Assert.Equal(
            new[] { "Add (1,1,2)", "Add (2,2,5)", "Add (3,3,6)" },
            report.Tests.Select(t => t.DisplayName));
        Assert.Equal("Failed", report.Tests[1].Outcome);
    }

    [Fact]
    public async Task GetResults_ReportsUnavailable_WhenGetterThrows()
    {
        using var host = new StubHost(new ThrowingTest());

        var report = await host.Interop.GetResultsAsync(new TestResultsQuery());

        Assert.False(report.Available);
        Assert.Contains("InvalidOperationException", report.Message);
        Assert.Contains("does not mean there are no failed tests", report.Message);
    }

    private static StubTest Failed(string name) =>
        new StubTest { FullyQualifiedName = name, DisplayName = name, State = TestState.Failed };

    private static StubTest Passed(string name) =>
        new StubTest { FullyQualifiedName = name, DisplayName = name, State = TestState.Passed };

    /// <summary>
    /// A MEF container holding the stand-in <see cref="ITestsService"/>, exposed through the
    /// <see cref="IComponentModel"/> shape that <see cref="TestResultsInterop"/> consumes.
    /// </summary>
    private sealed class StubHost : IDisposable
    {
        private readonly CompositionContainer _container;

        internal StubHost(params ITest[] tests)
        {
            _container = new CompositionContainer(new TypeCatalog(typeof(StubTestsService)));
            Service = (StubTestsService)_container.GetExportedValue<ITestsService>();
            Service.Tests.AddRange(tests);

            var componentModel = new StubComponentModel(_container);
            Interop = new TestResultsInterop(() => componentModel);
        }

        internal StubTestsService Service { get; }

        internal TestResultsInterop Interop { get; }

        public void Dispose() => _container.Dispose();
    }

    private sealed class StubComponentModel : IComponentModel
    {
        internal StubComponentModel(ExportProvider exportProvider)
        {
            DefaultExportProvider = exportProvider;
        }

        public ExportProvider DefaultExportProvider { get; }

        public ComposablePartCatalog DefaultCatalog => throw new NotSupportedException();

        public ICompositionService DefaultCompositionService => throw new NotSupportedException();

#pragma warning disable CS0618, CS0672 // Obsolete member is part of the interface being stubbed.
        public ComposablePartCatalog GetCatalog(string catalogName) => throw new NotSupportedException();
#pragma warning restore CS0618, CS0672

        public IEnumerable<T> GetExtensions<T>() where T : class => throw new NotSupportedException();

        public T GetService<T>() where T : class => throw new NotSupportedException();
    }

    [Export]
    private sealed class UnrelatedPart
    {
    }

    /// <summary>
    /// Implements every member explicitly, so a lookup through the class instead of the
    /// interface would fail.
    /// </summary>
    [Export(typeof(ITestsService))]
    [PartCreationPolicy(CreationPolicy.Shared)]
    private sealed class StubTestsService : ITestsService
    {
        internal bool IsRunning { get; set; } = true;

        internal List<ITest> Tests { get; } = new();

        bool ITestsService.IsTestServiceRunning => IsRunning;

        Task<IEnumerable<ITest>> ITestsService.GetTestsAsEnumerableAsync() =>
            Task.FromResult<IEnumerable<ITest>>(Tests.ToArray());

        Task<IEnumerable<ITest>> ITestsService.GetTestsAsEnumerableAsync(IEnumerable<Guid> testIds) =>
            throw new InvalidOperationException("The parameterless overload should be used.");
    }

    /// <summary>
    /// Explicit implementation, matching how the real Test Explorer types may implement it.
    /// </summary>
    private sealed class StubTest : ITest
    {
        internal string FullyQualifiedName { get; set; } = string.Empty;
        internal string DisplayName { get; set; } = string.Empty;
        internal string FilePath { get; set; } = string.Empty;
        internal string Source { get; set; } = string.Empty;
        internal TimeSpan Duration { get; set; }
        internal int LineNumber { get; set; }
        internal TestState State { get; set; }
        internal bool Stale { get; set; }
        internal List<IResult> Results { get; } = new();

        string ITest.FullyQualifiedName => FullyQualifiedName;
        string ITest.DisplayName => DisplayName;
        string ITest.FilePath => FilePath;
        string ITest.Source => Source;
        TimeSpan ITest.Duration => Duration;
        int ITest.LineNumber => LineNumber;
        TestState ITest.State => State;
        bool ITest.Stale => Stale;
        IEnumerable<IResult> ITest.Results => Results;
    }

    private sealed class StubResult : IResult
    {
        public string DisplayName { get; set; } = string.Empty;
        public TestState State { get; set; }
        public TimeSpan Duration { get; set; }
        public string ErrorMessage { get; set; } = null!;
        public string ErrorStackTrace { get; set; } = null!;
        public string StandardOutput { get; set; } = null!;
        public string StandardError { get; set; } = null!;
    }

    private sealed class ThrowingTest : ITest
    {
        public string FullyQualifiedName => "T.Throws";
        public string DisplayName => "Throws";
        public string FilePath => string.Empty;
        public string Source => string.Empty;
        public TimeSpan Duration => TimeSpan.Zero;
        public int LineNumber => 0;
        public TestState State => throw new InvalidOperationException("Test Explorer changed shape.");
        public bool Stale => false;
        public IEnumerable<IResult> Results => Array.Empty<IResult>();
    }
}
