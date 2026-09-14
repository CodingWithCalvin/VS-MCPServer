using CodingWithCalvin.MCPServer.Services;
using CodingWithCalvin.MCPServer.Shared.Models;
using EnvDTE;
using Xunit;

namespace CodingWithCalvin.MCPServer.Tests;

public class DebuggerBreakpointTests
{
    [Fact]
    public void ConditionTypes_MapToDteAndApiNames()
    {
        Assert.Equal(
            dbgBreakpointConditionType.dbgBreakpointConditionTypeWhenTrue,
            VisualStudioService.ToDteConditionType(BreakpointConditionType.WhenTrue));
        Assert.Equal(
            dbgBreakpointConditionType.dbgBreakpointConditionTypeWhenChanged,
            VisualStudioService.ToDteConditionType(BreakpointConditionType.WhenChanged));
        Assert.Equal(
            "whenTrue",
            VisualStudioService.GetConditionTypeName(dbgBreakpointConditionType.dbgBreakpointConditionTypeWhenTrue));
        Assert.Equal(
            "whenChanged",
            VisualStudioService.GetConditionTypeName(dbgBreakpointConditionType.dbgBreakpointConditionTypeWhenChanged));
    }

    [Fact]
    public void HitCountTypes_MapToDteAndApiNames()
    {
        Assert.Equal(dbgHitCountType.dbgHitCountTypeNone, VisualStudioService.ToDteHitCountType(BreakpointHitCountType.None));
        Assert.Equal(dbgHitCountType.dbgHitCountTypeEqual, VisualStudioService.ToDteHitCountType(BreakpointHitCountType.Equal));
        Assert.Equal(dbgHitCountType.dbgHitCountTypeGreaterOrEqual, VisualStudioService.ToDteHitCountType(BreakpointHitCountType.GreaterOrEqual));
        Assert.Equal(dbgHitCountType.dbgHitCountTypeMultiple, VisualStudioService.ToDteHitCountType(BreakpointHitCountType.Multiple));

        Assert.Equal("none", VisualStudioService.GetHitCountTypeName(dbgHitCountType.dbgHitCountTypeNone));
        Assert.Equal("equal", VisualStudioService.GetHitCountTypeName(dbgHitCountType.dbgHitCountTypeEqual));
        Assert.Equal("greaterOrEqual", VisualStudioService.GetHitCountTypeName(dbgHitCountType.dbgHitCountTypeGreaterOrEqual));
        Assert.Equal("multiple", VisualStudioService.GetHitCountTypeName(dbgHitCountType.dbgHitCountTypeMultiple));
    }

    [Theory]
    [InlineData(1, null, BreakpointHitCountType.None, true)]
    [InlineData(0, null, BreakpointHitCountType.None, false)]
    [InlineData(1, 1, BreakpointHitCountType.Equal, true)]
    [InlineData(1, 5, BreakpointHitCountType.GreaterOrEqual, true)]
    [InlineData(1, 2, BreakpointHitCountType.Multiple, true)]
    [InlineData(1, null, BreakpointHitCountType.Equal, false)]
    [InlineData(1, 0, BreakpointHitCountType.Equal, false)]
    [InlineData(1, 1, BreakpointHitCountType.None, false)]
    public void BreakpointConfiguration_ValidatesLineAndHitCount(
        int line,
        int? hitCount,
        BreakpointHitCountType hitCountType,
        bool expected)
    {
        Assert.Equal(
            expected,
            VisualStudioService.IsValidBreakpointConfiguration(line, hitCount, hitCountType));
    }
}
