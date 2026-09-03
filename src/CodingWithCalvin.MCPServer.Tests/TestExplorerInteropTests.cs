using System;
using System.ComponentModel.Composition;
using System.ComponentModel.Composition.Hosting;
using System.ComponentModel.Composition.Primitives;
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
    public void GetRunStatus_ReportsNoRunObserved_BeforeAnyStateChange()
    {
        using var host = new StubHost();

        var status = new TestExplorerInterop(host.Accessor).GetRunStatus();

        Assert.Equal("NoRunObserved", status.State);
        Assert.False(status.IsRunning);
        Assert.Null(status.LastOperation);
    }

    [Fact]
    public void GetRunStatus_TracksExecutionThroughToCompletion()
    {
        using var host = new StubHost();
        var interop = new TestExplorerInterop(host.Accessor);
        interop.EnsureTracking();

        host.Operations.Raise(TestOperationStates.TestExecutionStarted);
        var running = interop.GetRunStatus();

        host.Operations.Raise(TestOperationStates.TestExecutionFinished);
        var finished = interop.GetRunStatus();

        Assert.Equal("Running", running.State);
        Assert.True(running.IsRunning);

        Assert.Equal("Completed", finished.State);
        Assert.False(finished.IsRunning);
        Assert.Equal("TestExecutionFinished", finished.LastOperation);
    }

    [Fact]
    public void GetRunStatus_ReportsCanceled_AfterCancellation()
    {
        using var host = new StubHost();
        var interop = new TestExplorerInterop(host.Accessor);
        interop.EnsureTracking();

        host.Operations.Raise(TestOperationStates.TestExecutionStarted);
        host.Operations.Raise(TestOperationStates.TestExecutionCancelAndFinished);

        Assert.Equal("Canceled", interop.GetRunStatus().State);
    }

    [Fact]
    public void GetRunStatus_KeepsRunOutcome_WhenDiscoveryRunsAfterwards()
    {
        using var host = new StubHost();
        var interop = new TestExplorerInterop(host.Accessor);
        interop.EnsureTracking();

        host.Operations.Raise(TestOperationStates.TestExecutionFinished);
        host.Operations.Raise(TestOperationStates.DiscoveryFinished);

        // A discovery pass triggered by editing code must not erase the last run's outcome.
        Assert.Equal("Completed", interop.GetRunStatus().State);
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
    public void EnsureTracking_DoesNotThrow_WhenComponentModelMissing()
    {
        var interop = new TestExplorerInterop(() => null);

        interop.EnsureTracking();

        Assert.Equal("NoRunObserved", interop.GetRunStatus().State);
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
                new TypeCatalog(typeof(StubStatsService), typeof(StubOperationState)));
            _componentModel = new StubComponentModel(_container);

            Stats = (StubStatsService)_container.GetExportedValue<ITestExplorerStatsService>();
            Operations = (StubOperationState)_container.GetExportedValue<IOperationState>();
        }

        internal StubStatsService Stats { get; }

        internal StubOperationState Operations { get; }

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
}
