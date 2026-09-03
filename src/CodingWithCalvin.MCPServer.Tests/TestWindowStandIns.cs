using System;

// Stand-ins for the Test Explorer extensibility types, declared in the real namespace so that
// their full names match the MEF contract names TestExplorerInterop looks up.
//
// Microsoft.VisualStudio.TestWindow.Interfaces.dll ships only inside the Visual Studio
// installation: it is absent from the Microsoft.VisualStudio.SDK metapackage, nuget.org carries
// nothing newer than 11.0.61030 (2012), and $(DevEnvDir) is undefined under `dotnet build`. The
// production code therefore resolves these services by contract name and reads them
// reflectively, and these declarations let the tests drive that path without Visual Studio.
//
// Shapes are transcribed from the shipped assembly and are identical in VS 2022 17.14 and
// VS 2026 18.0. Only the members the interop touches are reproduced.
namespace Microsoft.VisualStudio.TestWindow.Extensibility;

public struct TestExplorerStats
{
    public int PassedTestCount { get; set; }
    public int FailedTestCount { get; set; }
    public int SkippedTestCount { get; set; }
    public int NotRunTestCount { get; set; }
}

public interface ITestExplorerStatsService
{
    TestExplorerStats TestExplorerStats { get; }
    bool IsTestExplorerStatsServiceRunning { get; }
    event EventHandler<TestExplorerStats> TestExplorerStatsChanged;
}

public enum TestOperationStates
{
    None = 0x00000000,
    Discovery = 0x00010000,
    DiscoveryStarting = 0x00010008,
    DiscoveryStarted = 0x00010001,
    DiscoveryFinished = 0x00010004,
    DiscoveryCanceled = 0x00010006,
    TestExecution = 0x00020000,
    TestExecutionStarting = 0x00020008,
    TestExecutionStarted = 0x00020001,
    TestExecutionCanceling = 0x00020003,
    TestExecutionFinished = 0x00020004,
    TestExecutionCancelAndFinished = 0x00020006
}

public class OperationStateChangedEventArgs : EventArgs
{
    public TestOperationStates State { get; set; }
}

public interface IOperationState
{
    event EventHandler<OperationStateChangedEventArgs> StateChanged;
}
