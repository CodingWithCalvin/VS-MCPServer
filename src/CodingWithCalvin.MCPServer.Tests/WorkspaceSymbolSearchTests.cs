using System;
using System.Linq;
using CodingWithCalvin.MCPServer.Services;
using CodingWithCalvin.MCPServer.Shared.Models;
using Xunit;

namespace CodingWithCalvin.MCPServer.Tests;

/// <summary>
/// Covers the matching and stopping rules behind <c>symbol_workspace</c> (issue #115).
/// </summary>
public class WorkspaceSymbolSearchTests
{
    [Theory]
    [InlineData("GetOrders", "Shop.OrderService.GetOrders")]
    [InlineData("getorders", "Shop.OrderService.getorders")]
    [InlineData("Load", "Shop.OrderService.Load")]
    public void Matches_IgnoresCase_OnNameOrFullName(string name, string fullName)
    {
        var search = new WorkspaceSymbolSearch("ORDER", maxResults: 10);

        Assert.True(search.Matches(name, fullName));
    }

    [Fact]
    public void Matches_False_WhenNeitherNameNorFullNameContainsQuery()
    {
        var search = new WorkspaceSymbolSearch("Order", maxResults: 10);

        Assert.False(search.Matches("Load", "Shop.Inventory.Load"));
    }

    [Fact]
    public void IsComplete_OnlyOnceOneMatchPastTheCapIsFound()
    {
        var search = new WorkspaceSymbolSearch("Get", maxResults: 3);

        AddSymbols(search, 3);
        Assert.False(search.IsComplete);

        AddSymbols(search, 1, first: 3);
        Assert.True(search.IsComplete);
    }

    [Fact]
    public void ToResult_ReturnsFirstMaxResults_AndIsTruncated_WhenMoreMatched()
    {
        var search = new WorkspaceSymbolSearch("Get", maxResults: 3);
        AddSymbols(search, 4);

        var result = search.ToResult();

        Assert.Equal(new[] { "Get0", "Get1", "Get2" }, result.Symbols.Select(s => s.Name).ToArray());
        Assert.True(result.Truncated);
    }

    [Fact]
    public void ToResult_IsNotTruncated_WhenExactlyMaxResultsMatched()
    {
        var search = new WorkspaceSymbolSearch("Get", maxResults: 3);
        AddSymbols(search, 3);

        var result = search.ToResult();

        Assert.Equal(3, result.Symbols.Count);
        Assert.False(result.Truncated);
    }

    [Fact]
    public void ToResult_IsEmpty_WhenNothingMatched()
    {
        var result = new WorkspaceSymbolSearch("Get", maxResults: 3).ToResult();

        Assert.Empty(result.Symbols);
        Assert.False(result.Truncated);
    }

    [Fact]
    public void Add_IgnoresSymbols_OnceComplete()
    {
        // A walk that kept going past IsComplete must not change the result.
        var search = new WorkspaceSymbolSearch("Get", maxResults: 2);
        AddSymbols(search, 10);

        var result = search.ToResult();

        Assert.Equal(new[] { "Get0", "Get1" }, result.Symbols.Select(s => s.Name).ToArray());
        Assert.True(result.Truncated);
    }

    [Theory]
    [InlineData(0)]
    [InlineData(-1)]
    public void Constructor_Rejects_MaxResultsBelowOne(int maxResults)
    {
        Assert.Throws<ArgumentOutOfRangeException>(() => new WorkspaceSymbolSearch("Get", maxResults));
    }

    [Fact]
    public void MaxResults_OfIntMaxValue_DoesNotOverflow()
    {
        var search = new WorkspaceSymbolSearch("Get", int.MaxValue);
        AddSymbols(search, 1);

        Assert.False(search.IsComplete);
        Assert.False(search.ToResult().Truncated);
    }

    private static void AddSymbols(WorkspaceSymbolSearch search, int count, int first = 0)
    {
        for (var i = first; i < first + count; i++)
        {
            search.Add(new SymbolInfo { Name = $"Get{i}", FullName = $"Shop.Service.Get{i}", Kind = SymbolKind.Function });
        }
    }
}
