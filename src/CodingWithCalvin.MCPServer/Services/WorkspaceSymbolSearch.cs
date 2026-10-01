using System;
using System.Collections.Generic;
using System.Linq;
using CodingWithCalvin.MCPServer.Shared.Models;

namespace CodingWithCalvin.MCPServer.Services;

/// <summary>
/// Collects the symbols that match a workspace symbol query and decides when the search can stop.
/// </summary>
/// <remarks>
/// <para>
/// The search walks every file's code model on the UI thread, so it stops as soon as the result is
/// settled instead of counting every match. One match past <c>maxResults</c> is enough to know the
/// results were truncated. That is also why there is no total count: reporting one would mean
/// finishing the walk and freezing Visual Studio for as long as that takes.
/// </para>
/// <para>
/// Kept free of Visual Studio types so the matching and stopping rules can be unit tested. The
/// caller walks the code model and checks <see cref="IsComplete"/> before each element.
/// </para>
/// </remarks>
internal sealed class WorkspaceSymbolSearch
{
    private readonly string _query;
    private readonly int _maxResults;
    private readonly int _limit;
    private readonly List<SymbolInfo> _symbols = new();

    public WorkspaceSymbolSearch(string query, int maxResults)
    {
        if (maxResults < 1)
        {
            throw new ArgumentOutOfRangeException(nameof(maxResults), maxResults, "At least one result must be requested.");
        }

        _query = query;
        _maxResults = maxResults;

        // One past the cap, so Truncated is exact. int.MaxValue cannot be exceeded anyway.
        _limit = maxResults == int.MaxValue ? maxResults : maxResults + 1;
    }

    /// <summary>
    /// True once enough matches have been found to settle the result, so the walk can stop.
    /// </summary>
    public bool IsComplete => _symbols.Count >= _limit;

    public bool Matches(string name, string fullName) =>
        name.IndexOf(_query, StringComparison.OrdinalIgnoreCase) >= 0
        || fullName.IndexOf(_query, StringComparison.OrdinalIgnoreCase) >= 0;

    /// <summary>
    /// Records a matching symbol. Ignored once the search is complete.
    /// </summary>
    public void Add(SymbolInfo symbol)
    {
        if (!IsComplete)
        {
            _symbols.Add(symbol);
        }
    }

    public WorkspaceSymbolResult ToResult() => new()
    {
        Symbols = _symbols.Take(_maxResults).ToList(),
        Truncated = _symbols.Count > _maxResults,
    };
}
