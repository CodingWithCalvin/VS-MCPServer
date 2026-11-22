using CodingWithCalvin.MCPServer.Services;
using Microsoft.VisualStudio.Debugger.Interop;
using Xunit;

namespace CodingWithCalvin.MCPServer.Tests;

public class DebuggerPerformanceCounterTests
{
    [Fact]
    public void CounterResult_ReturnsPerfTipSystemTime()
    {
        var results = new[]
        {
            new QUERY_PROCESS_PERFORMANCE_COUNTERS_RESULT
            {
                ErrorCode = 0,
                Counters = new PROCESS_PERFORMANCE_COUNTERS { SystemTime = 1234 }
            }
        };

        Assert.Equal(1234L, DebuggerPerformanceCounterService.ExtractElapsedMilliseconds(results));
    }

    [Fact]
    public void CounterResult_MatchesPerfTipUInt32Conversion()
    {
        var results = new[]
        {
            new QUERY_PROCESS_PERFORMANCE_COUNTERS_RESULT
            {
                Counters = new PROCESS_PERFORMANCE_COUNTERS { SystemTime = (ulong)uint.MaxValue + 1 }
            }
        };

        Assert.Equal(0L, DebuggerPerformanceCounterService.ExtractElapsedMilliseconds(results));
    }

    [Fact]
    public void CounterResult_RejectsMissingOrFailedResults()
    {
        Assert.Null(DebuggerPerformanceCounterService.ExtractElapsedMilliseconds(null));
        Assert.Null(DebuggerPerformanceCounterService.ExtractElapsedMilliseconds(
            System.Array.Empty<QUERY_PROCESS_PERFORMANCE_COUNTERS_RESULT>()));
        Assert.Null(DebuggerPerformanceCounterService.ExtractElapsedMilliseconds(new[]
        {
            new QUERY_PROCESS_PERFORMANCE_COUNTERS_RESULT
            {
                ErrorCode = -1,
                Counters = new PROCESS_PERFORMANCE_COUNTERS { SystemTime = 1234 }
            }
        }));
    }
}
