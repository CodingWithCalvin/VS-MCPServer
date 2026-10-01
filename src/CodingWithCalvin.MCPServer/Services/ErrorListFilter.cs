using System;
using System.Collections.Generic;
using System.Linq;
using CodingWithCalvin.MCPServer.Shared.Models;

namespace CodingWithCalvin.MCPServer.Services;

/// <summary>
/// Decides whether an Error List entry passes an <see cref="ErrorListQuery"/>.
/// </summary>
/// <remarks>
/// Kept free of Visual Studio types so the matching rules can be unit tested. The caller reads the
/// entries and resolves the scope (active document, open documents, current project) up front.
/// </remarks>
internal sealed class ErrorListFilter
{
    internal const string BuildSource = "Build";
    internal const string IntelliSenseSource = "IntelliSense";

    private readonly ErrorListQuery _query;
    private readonly ErrorListScopeContext _scope;
    private readonly HashSet<string> _openDocumentPaths;

    public ErrorListFilter(ErrorListQuery query, ErrorListScopeContext scope)
    {
        _query = query;
        _scope = scope;
        _openDocumentPaths = new HashSet<string>(
            scope.OpenDocumentPaths.Select(NormalizeSeparators),
            StringComparer.OrdinalIgnoreCase);

        NeedsEntryDetails = query.Scope != ErrorListScope.Solution
            || query.Source != ErrorListSource.All
            || query.SuppressionState != ErrorListSuppressionState.All
            || query.Codes.Count > 0
            || query.Tools.Count > 0
            || query.Categories.Count > 0
            || !string.IsNullOrWhiteSpace(query.Project)
            || !string.IsNullOrWhiteSpace(query.Path)
            || !string.IsNullOrWhiteSpace(query.Search);
    }

    /// <summary>
    /// True when some filter other than severity is set, so an entry has to be read in full to know
    /// whether it matches. Severity alone can be decided before reading anything else.
    /// </summary>
    public bool NeedsEntryDetails { get; }

    public bool MatchesSeverity(string severity) =>
        string.IsNullOrEmpty(_query.Severity)
        || severity.Equals(_query.Severity, StringComparison.OrdinalIgnoreCase);

    public bool Matches(ErrorListEntry entry)
    {
        var item = entry.Item;

        return MatchesSeverity(item.Severity)
            && MatchesSource(item.Source)
            && MatchesSuppressionState(item.SuppressionState)
            && MatchesScope(entry)
            && MatchesAny(item.ErrorCode, _query.Codes)
            && MatchesAny(item.Tool, _query.Tools)
            && MatchesAny(item.Category, _query.Categories)
            && (string.IsNullOrWhiteSpace(_query.Project)
                || entry.ProjectNames.Any(name => MatchesProjectName(name, _query.Project!)))
            && (string.IsNullOrWhiteSpace(_query.Path) || MatchesPath(item.FilePath, _query.Path!))
            && (string.IsNullOrWhiteSpace(_query.Search) || MatchesSearch(item, _query.Search!));
    }

    /// <summary>
    /// Matches a project name exactly, or the "Name (target framework)" form the Error List uses for
    /// multi-targeted projects.
    /// </summary>
    internal static bool MatchesProjectName(string entryProject, string project)
    {
        if (string.IsNullOrEmpty(entryProject))
        {
            return false;
        }

        var name = project.Trim();

        return entryProject.Equals(name, StringComparison.OrdinalIgnoreCase)
            || (entryProject.StartsWith(name + " (", StringComparison.OrdinalIgnoreCase)
                && entryProject.EndsWith(")", StringComparison.Ordinal));
    }

    /// <summary>
    /// Matches a file name, folder, or partial path against whole segments of the entry's path, so
    /// "src/App" matches "C:\src\App\Program.cs" but not "C:\src\Application\Program.cs".
    /// </summary>
    internal static bool MatchesPath(string entryPath, string path)
    {
        if (string.IsNullOrEmpty(entryPath))
        {
            return false;
        }

        var segments = NormalizeSeparators(path.Trim()).Trim('\\');
        if (segments.Length == 0)
        {
            return false;
        }

        return ("\\" + NormalizeSeparators(entryPath) + "\\")
            .IndexOf("\\" + segments + "\\", StringComparison.OrdinalIgnoreCase) >= 0;
    }

    /// <summary>
    /// Describes the categories the Error List window's own filters are hiding, or returns null when
    /// it hides none. Hidden entries never reach this tool, so callers should be told.
    /// </summary>
    internal static string? DescribeHiddenEntries(
        bool errorsShown,
        bool warningsShown,
        bool messagesShown,
        bool buildShown,
        bool intelliSenseShown)
    {
        var hidden = new List<string>();
        if (!errorsShown) hidden.Add("errors");
        if (!warningsShown) hidden.Add("warnings");
        if (!messagesShown) hidden.Add("messages");
        if (!buildShown) hidden.Add("build entries");
        if (!intelliSenseShown) hidden.Add("IntelliSense entries");

        if (hidden.Count == 0)
        {
            return null;
        }

        var list = hidden.Count == 1
            ? hidden[0]
            : string.Join(", ", hidden.Take(hidden.Count - 1)) + " and " + hidden[hidden.Count - 1];

        return $"The Error List window is hiding {list}, so they are not included. Change the window's filters to include them.";
    }

    private bool MatchesSource(string source) =>
        _query.Source switch
        {
            ErrorListSource.Build => source == BuildSource,
            ErrorListSource.IntelliSense => source == IntelliSenseSource,
            _ => true
        };

    private bool MatchesSuppressionState(string suppressionState) =>
        _query.SuppressionState == ErrorListSuppressionState.All
        || suppressionState == _query.SuppressionState.ToString();

    private static bool MatchesAny(string value, IReadOnlyCollection<string> allowed) =>
        allowed.Count == 0
        || allowed.Any(candidate => candidate.Equals(value, StringComparison.OrdinalIgnoreCase));

    private bool MatchesScope(ErrorListEntry entry)
    {
        var item = entry.Item;

        switch (_query.Scope)
        {
            case ErrorListScope.CurrentDocument:
                return _scope.ActiveDocumentPath != null
                    && !string.IsNullOrEmpty(item.FilePath)
                    && NormalizeSeparators(item.FilePath).Equals(
                        NormalizeSeparators(_scope.ActiveDocumentPath),
                        StringComparison.OrdinalIgnoreCase);

            case ErrorListScope.OpenDocuments:
                return !string.IsNullOrEmpty(item.FilePath)
                    && _openDocumentPaths.Contains(NormalizeSeparators(item.FilePath));

            case ErrorListScope.CurrentProject:
                // The Error List identifies the current project by GUID; build entries often carry
                // only a project name, so fall back to that.
                if (_scope.CurrentProjectGuid != Guid.Empty && entry.ProjectGuids.Count > 0)
                {
                    return entry.ProjectGuids.Contains(_scope.CurrentProjectGuid);
                }

                return _scope.CurrentProjectName != null
                    && entry.ProjectNames.Any(name => MatchesProjectName(name, _scope.CurrentProjectName));

            default:
                return true;
        }
    }

    private static bool MatchesSearch(ErrorItemInfo item, string search) =>
        Contains(item.Description, search)
        || Contains(item.ErrorCode, search)
        || Contains(item.Project, search)
        || Contains(item.FilePath, search)
        || Contains(item.Tool, search)
        || Contains(item.Category, search);

    private static bool Contains(string value, string search) =>
        value.IndexOf(search, StringComparison.OrdinalIgnoreCase) >= 0;

    private static string NormalizeSeparators(string path) => path.Replace('/', '\\');
}

/// <summary>
/// An Error List entry read for filtering: the item returned to the caller, plus every project the
/// entry belongs to.
/// </summary>
internal sealed class ErrorListEntry
{
    /// <param name="item">The entry, mapped from the Error List.</param>
    /// <param name="projectNames">
    /// Every project the entry belongs to. Usually one, but an error in a shared file belongs to each
    /// project that compiles it.
    /// </param>
    /// <param name="projectGuids">The GUIDs of those projects, where the entry supplies them.</param>
    public ErrorListEntry(
        ErrorItemInfo item,
        IReadOnlyCollection<string> projectNames,
        IReadOnlyCollection<Guid> projectGuids)
    {
        Item = item;
        ProjectNames = projectNames;
        ProjectGuids = projectGuids;
    }

    public ErrorItemInfo Item { get; }
    public IReadOnlyCollection<string> ProjectNames { get; }
    public IReadOnlyCollection<Guid> ProjectGuids { get; }
}

/// <summary>
/// The parts of the IDE state an <see cref="ErrorListScope"/> depends on, captured on the UI thread.
/// Only the members the requested scope needs are filled in.
/// </summary>
internal sealed class ErrorListScopeContext
{
    public string? ActiveDocumentPath { get; set; }
    public IReadOnlyCollection<string> OpenDocumentPaths { get; set; } = Array.Empty<string>();
    public string? CurrentProjectName { get; set; }
    public Guid CurrentProjectGuid { get; set; }
}
