using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text.Json;
using System.Threading.Tasks;
using CodingWithCalvin.MCPServer.Shared.Models;
using ModelContextProtocol.Server;

namespace CodingWithCalvin.MCPServer.Server.Tools;

[McpServerToolType]
public class DiagnosticsTools
{
    private static readonly string[] Severities = { "Error", "Warning", "Message" };

    private readonly RpcClient _rpcClient;
    private readonly JsonSerializerOptions _jsonOptions;

    public DiagnosticsTools(RpcClient rpcClient)
    {
        _rpcClient = rpcClient;
        _jsonOptions = new JsonSerializerOptions { WriteIndented = true };
    }

    [McpServerTool(Name = "errors_list", ReadOnly = true)]
    [Description("Get errors, warnings, and messages from the Error List. Returns diagnostics with file, line, description, severity, code, project, source (Build or IntelliSense), tool, category, and suppression state. Filters combine: an entry must match every filter given. Only entries the Error List window is currently showing can be returned, so the window's own scope and filters apply first; Notes reports anything the window is hiding. TotalCount, ErrorCount, WarningCount, and MessageCount always cover the whole Error List, regardless of filters or maxResults. MatchedCount is how many entries passed the filters, and Truncated is true when that is more than maxResults. An empty Error List can mean nothing has been built yet; build the solution to populate it.")]
    public async Task<string> GetErrorListAsync(
        [Description("Filter by severity: \"Error\", \"Warning\", \"Message\", or null for all. Case-insensitive.")]
        string? severity = null,
        [Description("Which documents to include: solution, openDocuments, currentProject (the project containing the active document), or currentDocument. Defaults to solution. The Error List's Changed Documents scope is not available.")]
        string scope = "solution",
        [Description("Which producer to include: all, build, or intellisense. Defaults to all.")]
        string source = "all",
        [Description("Which suppression state to include: all, active, suppressed, or notApplicable. Defaults to all. The Error List window hides suppressed entries by default, so suppressed returns them only when the window's Suppression State column filter includes them.")]
        string suppressionState = "all",
        [Description("Error code to match exactly, such as \"CS0103\". Separate several codes with commas. Case-insensitive.")]
        string? code = null,
        [Description("Tool that produced the entry, as shown in the Error List's Tool column, such as \"Compiler\". Separate several with commas. Case-insensitive.")]
        string? tool = null,
        [Description("Category, as shown in the Error List's Category column, such as \"Style\". Separate several with commas. Case-insensitive.")]
        string? category = null,
        [Description("Project name. Also matches multi-targeted entries such as \"MyApp (net8.0)\". Case-insensitive.")]
        string? project = null,
        [Description("File name, folder, or partial path, matched on whole path segments, such as \"Program.cs\" or \"src/Services\". Case-insensitive.")]
        string? path = null,
        [Description("Text to find in the description, code, project, file path, tool, or category. Case-insensitive.")]
        string? search = null,
        [Description("Maximum number of items to return. Defaults to 100.")]
        int maxResults = 100)
    {
        if (!string.IsNullOrEmpty(severity) && !Severities.Contains(severity, StringComparer.OrdinalIgnoreCase))
        {
            return $"Invalid severity '{severity}'. Valid values are: Error, Warning, Message.";
        }

        if (!TryParseScope(scope, out var parsedScope))
        {
            return $"Invalid scope '{scope}'. Valid values are: solution, openDocuments, currentProject, currentDocument.";
        }

        if (!TryParseSource(source, out var parsedSource))
        {
            return $"Invalid source '{source}'. Valid values are: all, build, intellisense.";
        }

        if (!TryParseSuppressionState(suppressionState, out var parsedSuppressionState))
        {
            return $"Invalid suppressionState '{suppressionState}'. Valid values are: all, active, suppressed, notApplicable.";
        }

        var query = new ErrorListQuery
        {
            Severity = string.IsNullOrEmpty(severity) ? null : severity,
            Scope = parsedScope,
            Source = parsedSource,
            SuppressionState = parsedSuppressionState,
            Codes = SplitList(code),
            Tools = SplitList(tool),
            Categories = SplitList(category),
            Project = project,
            Path = path,
            Search = search,
            MaxResults = maxResults
        };

        var result = await _rpcClient.GetErrorListAsync(query);

        // Always return the JSON result (includes debug info if TotalCount is 0)
        return JsonSerializer.Serialize(result, _jsonOptions);
    }

    [McpServerTool(Name = "output_read", ReadOnly = true)]
    [Description("Read content from an Output window pane. Specify pane by GUID or well-known name (\"Build\", \"Debug\", \"General\"). Note: Some panes may not support reading due to VS API limitations.")]
    public async Task<string> ReadOutputPaneAsync(
        [Description("Output pane identifier: GUID string or well-known name (\"Build\", \"Debug\", \"General\").")]
        string paneIdentifier)
    {
        var result = await _rpcClient.ReadOutputPaneAsync(paneIdentifier);

        if (string.IsNullOrEmpty(result.Content))
        {
            return $"Output pane '{paneIdentifier}' is empty or does not support reading";
        }

        return JsonSerializer.Serialize(result, _jsonOptions);
    }

    [McpServerTool(Name = "output_write", Destructive = false, Idempotent = false)]
    [Description("Write a message to an Output window pane. Custom panes are auto-created. System panes (Build, Debug) must already exist. Message is appended to existing content.")]
    public async Task<string> WriteOutputPaneAsync(
        [Description("Output pane identifier: GUID string or name. Custom GUIDs/names will create new panes if needed.")]
        string paneIdentifier,
        [Description("Message to write. Appended to existing content.")]
        string message,
        [Description("Whether to activate (bring to front) the Output window. Defaults to false.")]
        bool activate = false)
    {
        var success = await _rpcClient.WriteOutputPaneAsync(paneIdentifier, message, activate);
        return success
            ? $"Message written to output pane: {paneIdentifier}"
            : $"Failed to write to output pane: {paneIdentifier}";
    }

    [McpServerTool(Name = "output_list_panes", ReadOnly = true)]
    [Description("List available Output window panes. Returns well-known panes (Build, Debug, General) with their names and GUIDs.")]
    public async Task<string> GetOutputPanesAsync()
    {
        var panes = await _rpcClient.GetOutputPanesAsync();

        if (panes.Count == 0)
        {
            return "No output panes available";
        }

        return JsonSerializer.Serialize(panes, _jsonOptions);
    }

    private static bool TryParseScope(string value, out ErrorListScope scope)
    {
        return Enum.TryParse(value, ignoreCase: true, out scope)
            && Enum.IsDefined(typeof(ErrorListScope), scope);
    }

    private static bool TryParseSource(string value, out ErrorListSource source)
    {
        return Enum.TryParse(value, ignoreCase: true, out source)
            && Enum.IsDefined(typeof(ErrorListSource), source);
    }

    private static bool TryParseSuppressionState(string value, out ErrorListSuppressionState suppressionState)
    {
        return Enum.TryParse(value, ignoreCase: true, out suppressionState)
            && Enum.IsDefined(typeof(ErrorListSuppressionState), suppressionState);
    }

    private static List<string> SplitList(string? value)
    {
        return (value ?? string.Empty)
            .Split(',')
            .Select(item => item.Trim())
            .Where(item => item.Length > 0)
            .ToList();
    }
}
