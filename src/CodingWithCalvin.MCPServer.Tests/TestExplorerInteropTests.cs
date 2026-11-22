using System;
using System.Collections.Generic;
using System.ComponentModel.Composition;
using System.ComponentModel.Composition.Hosting;
using System.ComponentModel.Composition.Primitives;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using CodingWithCalvin.MCPServer.Services;
using Microsoft.VisualStudio.ComponentModelHost;
using Microsoft.VisualStudio.TestWindow.Extensibility;
using Xunit;

namespace CodingWithCalvin.MCPServer.Tests;

/// <summary>
/// Covers the reflection bridge to Test Explorer.
/// </summary>
/// <remarks>
/// Microsoft.VisualStudio.TestWindow.Interfaces.dll cannot be referenced at build time, so
/// <see cref="TestExplorerInterop"/> resolves its services by MEF contract name and reads them
/// reflectively. These tests stand up a real MEF container holding stand-in parts declared in
/// the genuine namespace and shaped like the real ones, which exercises the contract-name
/// lookup, the interface discovery, the property reads and the explicitly implemented event
/// subscription rather than just the pure helpers.
/// </remarks>
public class TestExplorerInteropTests
{
    [Fact]
    public void GetStats_ReportsUnavailable_WhenComponentModelMissing()
    {
        var interop = new TestExplorerInterop(() => null);

        var stats = interop.GetStats();

        Assert.False(stats.Available);
    }

    [Fact]
    public void GetStats_ReportsUnavailable_WhenTestExplorerNotRunning()
    {
        using var host = new StubHost();
        host.Stats.IsTestExplorerStatsServiceRunning = false;
        host.Stats.Counts = new TestExplorerStats { PassedTestCount = 7 };

        var stats = new TestExplorerInterop(host.Accessor).GetStats();

        // Counts must not leak through as zeros that read like "solution has no tests".
        Assert.False(stats.Available);
        Assert.Equal(0, stats.Passed);
    }

    [Fact]
    public void GetStats_ReadsCounts_WhenTestExplorerRunning()
    {
        using var host = new StubHost();
        host.Stats.IsTestExplorerStatsServiceRunning = true;
        host.Stats.Counts = new TestExplorerStats
        {
            PassedTestCount = 12,
            FailedTestCount = 3,
            SkippedTestCount = 1,
            NotRunTestCount = 4
        };

        var stats = new TestExplorerInterop(host.Accessor).GetStats();

        Assert.True(stats.Available);
        Assert.Equal(12, stats.Passed);
        Assert.Equal(3, stats.Failed);
        Assert.Equal(1, stats.Skipped);
        Assert.Equal(4, stats.NotRun);
    }

    [Fact]
    public async Task GetTestsAsync_MapsAndFiltersDiscoveredTests()
    {
        using var host = new StubHost();
        host.Tests.Tests.Add(Test("CalculatorTests.Add", "CalculatorTests.Add", TestState.Passed));
        host.Tests.Tests.Add(Test("CalculatorTests.Subtract", "CalculatorTests.Subtract", TestState.Failed));

        var result = await new TestExplorerInterop(host.Accessor).GetTestsAsync("Subtract", "Failed", 10);

        Assert.True(result.Available);
        Assert.Equal(1, result.TotalCount);
        Assert.Equal("CalculatorTests.Subtract", Assert.Single(result.Tests).FullyQualifiedName);
    }

    [Fact]
    public async Task GetResultsAsync_ReturnsStructuredFailureDetails()
    {
        using var host = new StubHost();
        var test = Test("CalculatorTests.Add", "CalculatorTests.Add", TestState.Failed);
        test.Results.Add(new StubResult
        {
            State = TestState.Failed,
            ErrorMessage = "Expected 2 but got 3",
            ErrorStackTrace = "at CalculatorTests.Add()",
            StandardOutput = "test output",
            ReceivedOrder = 42
        });
        host.Tests.Tests.Add(test);

        var result = await new TestExplorerInterop(host.Accessor).GetResultsAsync(null, null, true, 10);

        var failure = Assert.Single(result.Results);
        Assert.True(result.Available);
        Assert.Equal("Expected 2 but got 3", failure.ErrorMessage);
        Assert.Equal("at CalculatorTests.Add()", failure.ErrorStackTrace);
        Assert.Equal("test output", failure.StandardOutput);
    }

    [Theory]
    [InlineData(false, "Run")]
    [InlineData(true, "Debug")]
    public async Task RunTestsAsync_InvokesTestExplorerWithSelectedIds(bool debug, string expectedOperation)
    {
        using var host = new StubHost();
        var id = Guid.NewGuid();

        var result = await new TestExplorerInterop(host.Accessor).RunTestsAsync(new[] { id.ToString() }, debug);

        Assert.True(result.Started);
        Assert.Equal(1, result.SelectedCount);
        Assert.Equal(expectedOperation, host.VsTests.LastOperation);
        Assert.Equal(TestPropertyType.TestCaseRecordId, host.VsTests.LastQuery?.PropertyType);
        Assert.Equal(FilterMatchKind.ExactMatch, host.VsTests.LastQuery?.MatchKind);
        Assert.Equal(id.ToString(), Assert.Single(host.VsTests.LastQuery!.Values));
        Assert.Null(host.Tests.LastOperation);
    }

    [Fact]
    public async Task RunTestsAsync_ReportsRejectedRunWithoutLosingSelection()
    {
        using var host = new StubHost();
        host.VsTests.Result = false;
        var id = Guid.NewGuid();

        var result = await new TestExplorerInterop(host.Accessor).RunTestsAsync(new[] { id.ToString() }, debug: false);

        Assert.False(result.Started);
        Assert.Equal(1, result.SelectedCount);
        Assert.Equal("RunTestsAsync returned false; Test Explorer did not start the selected tests.", result.Message);
    }

    [Fact]
    public async Task RunTestsAsync_ReportsInvocationFailureWithoutLosingSelection()
    {
        using var host = new StubHost();
        host.VsTests.Exception = new InvalidOperationException("Selected test run failed.");
        var id = Guid.NewGuid();

        var result = await new TestExplorerInterop(host.Accessor).RunTestsAsync(new[] { id.ToString() }, debug: false);

        Assert.False(result.Started);
        Assert.Equal(1, result.SelectedCount);
        Assert.Equal("Selected test run failed.", result.Message);
    }

    [Fact]
    public async Task GetRunStatus_ReportsNoRunObserved_BeforeAnyStateChange()
    {
        using var host = new StubHost();

        var status = await new TestExplorerInterop(host.Accessor).GetRunStatusAsync();

        Assert.Equal("NoRunObserved", status.State);
        Assert.False(status.IsRunning);
        Assert.Null(status.LastOperation);
    }

    [Fact]
    public async Task GetRunStatus_ReportsRunning_WhenTestNodeIsRunningWithoutOperationEvent()
    {
        using var host = new StubHost();
        var test = Test("CalculatorTests.Add", "CalculatorTests.Add", TestState.NotRun);
        test.IsCurrentlyRunning = true;
        host.Tests.Tests.Add(test);

        var status = await new TestExplorerInterop(host.Accessor).GetRunStatusAsync();

        Assert.Equal("Running", status.State);
        Assert.True(status.IsRunning);
        Assert.Null(status.LastOperation);
    }

    [Fact]
    public async Task GetRunStatus_TracksExecutionThroughToCompletion()
    {
        using var host = new StubHost();
        var interop = new TestExplorerInterop(host.Accessor);
        interop.EnsureTracking();

        host.Operations.Raise(TestOperationStates.TestExecutionStarted);
        var running = await interop.GetRunStatusAsync();

        host.Operations.Raise(TestOperationStates.TestExecutionFinished);
        var finished = await interop.GetRunStatusAsync();

        Assert.Equal("Running", running.State);
        Assert.True(running.IsRunning);

        Assert.Equal("Completed", finished.State);
        Assert.False(finished.IsRunning);
        Assert.Equal("TestExecutionFinished", finished.LastOperation);
    }

    [Fact]
    public async Task GetRunStatus_ReportsCanceled_AfterCancellation()
    {
        using var host = new StubHost();
        var interop = new TestExplorerInterop(host.Accessor);
        interop.EnsureTracking();

        host.Operations.Raise(TestOperationStates.TestExecutionStarted);
        host.Operations.Raise(TestOperationStates.TestExecutionCancelAndFinished);

        Assert.Equal("Canceled", (await interop.GetRunStatusAsync()).State);
    }

    [Fact]
    public async Task GetRunStatus_KeepsRunOutcome_WhenDiscoveryRunsAfterwards()
    {
        using var host = new StubHost();
        var interop = new TestExplorerInterop(host.Accessor);
        interop.EnsureTracking();

        host.Operations.Raise(TestOperationStates.TestExecutionFinished);
        host.Operations.Raise(TestOperationStates.DiscoveryFinished);

        // A discovery pass triggered by editing code must not erase the last run's outcome.
        Assert.Equal("Completed", (await interop.GetRunStatusAsync()).State);
    }

    [Fact]
    public void EnsureTracking_SubscribesOnlyOnce()
    {
        using var host = new StubHost();
        var interop = new TestExplorerInterop(host.Accessor);

        interop.EnsureTracking();
        interop.EnsureTracking();
        interop.EnsureTracking();

        Assert.Equal(1, host.Operations.SubscriberCount);
    }

    [Fact]
    public async Task EnsureTracking_DoesNotThrow_WhenComponentModelMissing()
    {
        var interop = new TestExplorerInterop(() => null);

        interop.EnsureTracking();

        Assert.Equal("NoRunObserved", (await interop.GetRunStatusAsync()).State);
    }

    [Theory]
    [InlineData(null, null, "NoRunObserved")]
    [InlineData("DiscoveryStarted", null, "Discovering")]
    [InlineData("DiscoveryStarting", null, "Discovering")]
    [InlineData("DiscoveryFinished", null, "NoRunObserved")]
    [InlineData("TestExecutionStarted", "TestExecutionStarted", "Running")]
    [InlineData("TestExecutionStarting", "TestExecutionStarting", "Running")]
    [InlineData("TestExecutionCanceling", "TestExecutionCanceling", "Canceling")]
    [InlineData("TestExecutionFinished", "TestExecutionFinished", "Completed")]
    [InlineData("TestExecutionCancelAndFinished", "TestExecutionCancelAndFinished", "Canceled")]
    [InlineData("DiscoveryStarted", "TestExecutionFinished", "Discovering")]
    [InlineData("DiscoveryFinished", "TestExecutionFinished", "Completed")]
    public void DeriveState_CollapsesOperationNames(string? last, string? lastExecution, string expected)
    {
        Assert.Equal(expected, TestExplorerInterop.DeriveState(last, lastExecution));
    }

    /// <summary>
    /// A MEF container holding the stand-in Test Explorer parts, exposed through the
    /// <see cref="IComponentModel"/> shape that <see cref="TestExplorerInterop"/> consumes.
    /// </summary>
    private sealed class StubHost : IDisposable
    {
        private readonly CompositionContainer _container;
        private readonly StubComponentModel _componentModel;

        internal StubHost()
        {
            _container = new CompositionContainer(
                new TypeCatalog(
                    typeof(StubStatsService),
                    typeof(StubOperationState),
                    typeof(StubTestsService),
                    typeof(StubVsTestService)));
            _componentModel = new StubComponentModel(_container);

            Stats = (StubStatsService)_container.GetExportedValue<ITestExplorerStatsService>();
            Operations = (StubOperationState)_container.GetExportedValue<IOperationState>();
            Tests = (StubTestsService)_container.GetExportedValue<ITestsService>();
            VsTests = (StubVsTestService)_container.GetExportedValue<IVsTestService>();
        }

        internal StubStatsService Stats { get; }

        internal StubOperationState Operations { get; }

        internal StubTestsService Tests { get; }

        internal StubVsTestService VsTests { get; }

        internal Func<IComponentModel?> Accessor => () => _componentModel;

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

        public System.Collections.Generic.IEnumerable<T> GetExtensions<T>() where T : class
            => throw new NotSupportedException();

        public T GetService<T>() where T : class => throw new NotSupportedException();
    }

    [Export(typeof(ITestExplorerStatsService))]
    [PartCreationPolicy(CreationPolicy.Shared)]
    private sealed class StubStatsService : ITestExplorerStatsService
    {
        internal TestExplorerStats Counts { get; set; }

        public bool IsTestExplorerStatsServiceRunning { get; set; }

        TestExplorerStats ITestExplorerStatsService.TestExplorerStats => Counts;

        public event EventHandler<TestExplorerStats>? TestExplorerStatsChanged;

        internal void RaiseStatsChanged() => TestExplorerStatsChanged?.Invoke(this, Counts);
    }

    /// <summary>
    /// Implements StateChanged explicitly, matching the real OperationBroker, so the interop's
    /// interface-based event lookup is genuinely exercised.
    /// </summary>
    [Export(typeof(IOperationState))]
    [PartCreationPolicy(CreationPolicy.Shared)]
    private sealed class StubOperationState : IOperationState
    {
        private EventHandler<OperationStateChangedEventArgs>? _stateChanged;

        internal int SubscriberCount { get; private set; }

        event EventHandler<OperationStateChangedEventArgs> IOperationState.StateChanged
        {
            add
            {
                _stateChanged += value;
                SubscriberCount++;
            }
            remove
            {
                _stateChanged -= value;
                SubscriberCount--;
            }
        }

        internal void Raise(TestOperationStates state) =>
            _stateChanged?.Invoke(this, new OperationStateChangedEventArgs { State = state });
    }

    [Export(typeof(IVsTestService))]
    [PartCreationPolicy(CreationPolicy.Shared)]
    private sealed class StubVsTestService : IVsTestService
    {
        internal TestQuery? LastQuery { get; private set; }
        internal string? LastOperation { get; private set; }
        internal bool Result { get; set; } = true;
        internal Exception? Exception { get; set; }

        public bool IsServiceAvailable => true;

        public event EventHandler OnReady
        {
            add { }
            remove { }
        }

        public event EventHandler OnInactive
        {
            add { }
            remove { }
        }

        public Task<IEnumerable<ITestNode>> GetTestsAsync(TestQuery query, CancellationToken cancellationToken) =>
            Task.FromResult<IEnumerable<ITestNode>>(Array.Empty<ITestNode>());

        public Task<bool> RunTestsAsync(TestQuery query, CancellationToken cancellationToken) =>
            ExecuteAsync("Run", query);

        public Task<bool> DebugTestsAsync(TestQuery query, CancellationToken cancellationToken) =>
            ExecuteAsync("Debug", query);

        private Task<bool> ExecuteAsync(string operation, TestQuery query)
        {
            LastOperation = operation;
            LastQuery = query;
            return Exception == null ? Task.FromResult(Result) : Task.FromException<bool>(Exception);
        }
    }

    [Export(typeof(ITestsService))]
    [PartCreationPolicy(CreationPolicy.Shared)]
    private sealed class StubTestsService : ITestsService
    {
        internal List<ITest> Tests { get; } = new List<ITest>();
        internal List<Guid> LastTestIds { get; private set; } = new List<Guid>();
        internal string? LastOperation { get; private set; }

        public bool IsTestServiceRunning => true;

        public event EventHandler TestServiceStarted
        {
            add { }
            remove { }
        }

        public Task<IEnumerable<ITest>> GetTestsAsEnumerableAsync() =>
            Task.FromResult<IEnumerable<ITest>>(Tests);

        public Task<IEnumerable<ITest>> GetTestsAsEnumerableAsync(IEnumerable<Guid> testIds)
        {
            var ids = new HashSet<Guid>(testIds);
            return Task.FromResult<IEnumerable<ITest>>(Tests.Where(test => ids.Contains(test.Id)).ToList());
        }

        public Task<bool> RunTestsAsync(IEnumerable<Guid> tests)
        {
            LastOperation = "Run";
            LastTestIds = tests.ToList();
            return Task.FromResult(true);
        }

        public Task<bool> DebugTestsAsync(IEnumerable<Guid> tests)
        {
            LastOperation = "Debug";
            LastTestIds = tests.ToList();
            return Task.FromResult(true);
        }
    }

    private static StubTest Test(string fullyQualifiedName, string displayName, TestState state) => new()
    {
        Id = Guid.NewGuid(),
        FullyQualifiedName = fullyQualifiedName,
        DisplayName = displayName,
        State = state,
        FilePath = @"C:\src\CalculatorTests.cs",
        LineNumber = 12,
        Source = "CalculatorTests.dll"
    };

    private sealed class StubTest : ITest
    {
        public Guid Id { get; set; }
        public string FullyQualifiedName { get; set; } = string.Empty;
        public string DisplayName { get; set; } = string.Empty;
        public TestState State { get; set; }
        public bool Stale { get; set; }
        public bool IsCurrentlyRunning { get; set; }
        public string? FilePath { get; set; }
        public int LineNumber { get; set; }
        public string Source { get; set; } = string.Empty;
        public List<IResult> Results { get; } = new List<IResult>();
        IEnumerable<IResult> ITest.Results => Results;
    }

    private sealed class StubResult : IResult
    {
        public string DisplayName { get; set; } = string.Empty;
        public short RunId { get; set; }
        public long ReceivedOrder { get; set; }
        public short Outcome { get; set; }
        public TestState State { get; set; }
        public string? ErrorStackTrace { get; set; }
        public string? ErrorMessage { get; set; }
        public TimeSpan Duration { get; set; }
        public string? StandardOutput { get; set; }
        public string? StandardError { get; set; }
        public string? AttachmentsString { get; set; }
    }
}
