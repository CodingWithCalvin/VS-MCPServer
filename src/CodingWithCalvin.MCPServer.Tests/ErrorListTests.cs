using System;
using System.Collections.Generic;
using System.Linq;
using CodingWithCalvin.MCPServer.Services;
using CodingWithCalvin.MCPServer.Shared.Models;
using Xunit;

namespace CodingWithCalvin.MCPServer.Tests;

public class ErrorListTests
{
    [Fact]
    public void CollectErrorListEntries_CountsEveryEntry_WhenResultsAreCapped()
    {
        var entries = Entries(errors: 250, warnings: 3, messages: 2);

        var result = Collect(entries, severityFilter: null, maxResults: 100);

        Assert.Equal(100, result.Items.Count);
        Assert.Equal(255, result.TotalCount);
        Assert.Equal(250, result.ErrorCount);
        Assert.Equal(3, result.WarningCount);
        Assert.Equal(2, result.MessageCount);
        Assert.True(result.Truncated);
    }

    [Fact]
    public void CollectErrorListEntries_CountsFilteredOutSeverities()
    {
        var entries = Entries(errors: 2, warnings: 4, messages: 1);

        var result = Collect(entries, severityFilter: "error", maxResults: 100);

        Assert.Equal(2, result.Items.Count);
        Assert.All(result.Items, item => Assert.Equal("Error", item.Severity));
        Assert.Equal(7, result.TotalCount);
        Assert.Equal(2, result.ErrorCount);
        Assert.Equal(4, result.WarningCount);
        Assert.Equal(1, result.MessageCount);
        Assert.False(result.Truncated);
    }

    [Fact]
    public void CollectErrorListEntries_TruncatedOnlyCountsMatchingEntries()
    {
        // 3 warnings fit within the cap even though the whole list is larger than it.
        var entries = Entries(errors: 10, warnings: 3, messages: 0);

        var result = Collect(entries, severityFilter: "Warning", maxResults: 5);

        Assert.Equal(3, result.Items.Count);
        Assert.Equal(13, result.TotalCount);
        Assert.False(result.Truncated);
    }

    [Theory]
    [InlineData(5, 5, false)]
    [InlineData(6, 5, true)]
    [InlineData(1, 0, true)]
    public void CollectErrorListEntries_SetsTruncatedAtTheCapBoundary(int errors, int maxResults, bool expected)
    {
        var result = Collect(Entries(errors, warnings: 0, messages: 0), severityFilter: null, maxResults);

        Assert.Equal(Math.Min(errors, maxResults), result.Items.Count);
        Assert.Equal(expected, result.Truncated);
    }

    [Fact]
    public void CollectErrorListEntries_ReturnsNoItems_WhenNothingMatches()
    {
        var entries = Entries(errors: 0, warnings: 2, messages: 1);

        var result = Collect(entries, severityFilter: "Error", maxResults: 100);

        Assert.Empty(result.Items);
        Assert.Equal(3, result.TotalCount);
        Assert.Equal(0, result.ErrorCount);
        Assert.False(result.Truncated);
    }

    [Fact]
    public void CollectErrorListEntries_ReturnsEmptyResult_ForEmptyErrorList()
    {
        var result = Collect(Array.Empty<string>(), severityFilter: null, maxResults: 100);

        Assert.Empty(result.Items);
        Assert.Equal(0, result.TotalCount);
        Assert.False(result.Truncated);
    }

    [Fact]
    public void CollectErrorListEntries_OnlyReadsEntriesThatAreReturned()
    {
        var created = 0;

        var result = VisualStudioService.CollectErrorListEntries(
            Entries(errors: 50, warnings: 0, messages: 0),
            entry => entry,
            (entry, severity) =>
            {
                created++;
                return new ErrorItemInfo { Severity = severity };
            },
            severityFilter: null,
            maxResults: 10);

        Assert.Equal(10, created);
        Assert.Equal(50, result.TotalCount);
    }

    [Fact]
    public void CollectErrorListEntries_SkipsEntriesWhoseSeverityCannotBeRead()
    {
        var entries = new[] { "Error", null, "Warning" };

        var result = VisualStudioService.CollectErrorListEntries(
            entries,
            entry => entry,
            (entry, severity) => new ErrorItemInfo { Severity = severity },
            severityFilter: null,
            maxResults: 100);

        Assert.Equal(2, result.Items.Count);
        Assert.Equal(2, result.TotalCount);
    }

    [Fact]
    public void CollectErrorListEntries_CountsEntriesThatFailToRead()
    {
        // An entry whose details cannot be read is still counted, so it is not silently lost.
        var entries = Entries(errors: 3, warnings: 0, messages: 0);
        var index = 0;

        var result = VisualStudioService.CollectErrorListEntries(
            entries,
            entry => entry,
            (entry, severity) => index++ == 1 ? null : new ErrorItemInfo { Severity = severity },
            severityFilter: null,
            maxResults: 100);

        Assert.Equal(2, result.Items.Count);
        Assert.Equal(3, result.TotalCount);
        Assert.Equal(3, result.ErrorCount);
        Assert.False(result.Truncated);
    }

    private static ErrorListResult Collect(IEnumerable<string> entries, string? severityFilter, int maxResults) =>
        VisualStudioService.CollectErrorListEntries(
            entries,
            entry => entry,
            (entry, severity) => new ErrorItemInfo { Severity = severity },
            severityFilter,
            maxResults);

    // Interleaves severities so capping and filtering cannot pass by relying on order.
    private static List<string> Entries(int errors, int warnings, int messages)
    {
        var groups = new[]
        {
            Enumerable.Repeat("Error", errors).ToList(),
            Enumerable.Repeat("Warning", warnings).ToList(),
            Enumerable.Repeat("Message", messages).ToList()
        };

        var entries = new List<string>();
        for (var i = 0; i < groups.Max(g => g.Count); i++)
        {
            foreach (var group in groups.Where(g => i < g.Count))
            {
                entries.Add(group[i]);
            }
        }

        return entries;
    }
}
