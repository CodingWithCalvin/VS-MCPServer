using System;
using System.Collections.Generic;
using System.ComponentModel.Composition;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using CodingWithCalvin.MCPServer.Shared.Models;
using CodingWithCalvin.Otel4Vsix;
using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio.Shell;

namespace CodingWithCalvin.MCPServer.Services;

[Export(typeof(IVisualStudioService))]
[PartCreationPolicy(CreationPolicy.Shared)]
public class VisualStudioService : IVisualStudioService
{
    private IServiceProvider? _serviceProvider;

    private IServiceProvider ServiceProvider =>
        _serviceProvider ??= MCPServerPackage.Instance as IServiceProvider
            ?? throw new InvalidOperationException("Package not initialized");

    private async Task<DTE2> GetDteAsync()
    {
        await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();
        return ServiceProvider.GetService(typeof(DTE)) as DTE2
            ?? throw new InvalidOperationException("Could not get DTE service");
    }

    private static string NormalizePath(string path)
    {
        return Path.GetFullPath(path.Replace('/', '\\'));
    }

    private static bool PathsEqual(string path1, string path2)
    {
        return NormalizePath(path1).Equals(NormalizePath(path2), StringComparison.OrdinalIgnoreCase);
    }

    public async Task<SolutionInfo?> GetSolutionInfoAsync()
    {
        await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();
        var dte = await GetDteAsync();

        if (dte.Solution == null || string.IsNullOrEmpty(dte.Solution.FullName))
        {
            return null;
        }

        return new SolutionInfo
        {
            Name = Path.GetFileNameWithoutExtension(dte.Solution.FullName),
            Path = dte.Solution.FullName,
            IsOpen = dte.Solution.IsOpen
        };
    }

    public async Task<bool> OpenSolutionAsync(string path)
    {
        using var activity = VsixTelemetry.Tracer.StartActivity("OpenSolution");

        await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();
        var dte = await GetDteAsync();

        try
        {
            dte.Solution.Open(path);
            return true;
        }
        catch (Exception ex)
        {
            activity?.SetStatus(ActivityStatusCode.Error, ex.Message);
            activity?.RecordException(ex);
            return false;
        }
    }

    public async Task CloseSolutionAsync(bool saveFirst = true)
    {
        await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();
        var dte = await GetDteAsync();
        dte.Solution.Close(saveFirst);
    }

    public async Task<List<ProjectInfo>> GetProjectsAsync()
    {
        await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();
        var dte = await GetDteAsync();
        var projects = new List<ProjectInfo>();

        if (dte.Solution == null)
        {
            return projects;
        }

        foreach (EnvDTE.Project project in dte.Solution.Projects)
        {
            try
            {
                projects.Add(new ProjectInfo
                {
                    Name = project.Name,
                    Path = project.FullName,
                    Kind = project.Kind
                });
            }
            catch (Exception ex)
            {
                VsixTelemetry.TrackException(ex);
            }
        }

        return projects;
    }

    public async Task<List<DocumentInfo>> GetOpenDocumentsAsync()
    {
        await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();
        var dte = await GetDteAsync();
        var documents = new List<DocumentInfo>();

        foreach (Document doc in dte.Documents)
        {
            try
            {
                documents.Add(new DocumentInfo
                {
                    Name = doc.Name,
                    Path = doc.FullName,
                    IsSaved = doc.Saved
                });
            }
            catch (Exception ex)
            {
                VsixTelemetry.TrackException(ex);
            }
        }

        return documents;
    }

    public async Task<DocumentInfo?> GetActiveDocumentAsync()
    {
        await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();
        var dte = await GetDteAsync();

        var doc = dte.ActiveDocument;
        if (doc == null)
        {
            return null;
        }

        return new DocumentInfo
        {
            Name = doc.Name,
            Path = doc.FullName,
            IsSaved = doc.Saved
        };
    }

    public async Task<bool> OpenDocumentAsync(string path)
    {
        using var activity = VsixTelemetry.Tracer.StartActivity("OpenDocument");

        await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();
        var dte = await GetDteAsync();

        try
        {
            dte.ItemOperations.OpenFile(path);
            return true;
        }
        catch (Exception ex)
        {
            activity?.SetStatus(ActivityStatusCode.Error, ex.Message);
            activity?.RecordException(ex);
            return false;
        }
    }

    public async Task<bool> CloseDocumentAsync(string path, bool save = true)
    {
        await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();
        var dte = await GetDteAsync();

        foreach (Document doc in dte.Documents)
        {
            try
            {
                if (PathsEqual(doc.FullName, path))
                {
                    doc.Close(save ? vsSaveChanges.vsSaveChangesYes : vsSaveChanges.vsSaveChangesNo);
                    return true;
                }
            }
            catch (Exception ex)
            {
                VsixTelemetry.TrackException(ex);
            }
        }

        return false;
    }

    public async Task<string?> ReadDocumentAsync(string path)
    {
        await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();
        var dte = await GetDteAsync();

        foreach (Document doc in dte.Documents)
        {
            try
            {
                if (PathsEqual(doc.FullName, path))
                {
                    var textDoc = doc.Object("TextDocument") as TextDocument;
                    if (textDoc != null)
                    {
                        var editPoint = textDoc.StartPoint.CreateEditPoint();
                        return editPoint.GetText(textDoc.EndPoint);
                    }
                }
            }
            catch (Exception ex)
            {
                VsixTelemetry.TrackException(ex);
            }
        }

        if (File.Exists(path))
        {
            return await Task.Run(() => File.ReadAllText(path));
        }

        return null;
    }

    public async Task<bool> WriteDocumentAsync(string path, string content)
    {
        await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();
        var dte = await GetDteAsync();

        foreach (Document doc in dte.Documents)
        {
            try
            {
                if (PathsEqual(doc.FullName, path))
                {
                    var textDoc = doc.Object("TextDocument") as TextDocument;
                    if (textDoc != null)
                    {
                        var editPoint = textDoc.StartPoint.CreateEditPoint();
                        editPoint.Delete(textDoc.EndPoint);
                        editPoint.Insert(content);
                        return true;
                    }
                }
            }
            catch (Exception ex)
            {
                VsixTelemetry.TrackException(ex);
            }
        }

        return false;
    }

    public async Task<SelectionInfo?> GetSelectionAsync()
    {
        await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();
        var dte = await GetDteAsync();

        var doc = dte.ActiveDocument;
        if (doc == null)
        {
            return null;
        }

        var textDoc = doc.Object("TextDocument") as TextDocument;
        if (textDoc == null)
        {
            return null;
        }

        var selection = textDoc.Selection;
        return new SelectionInfo
        {
            Text = selection.Text,
            StartLine = selection.TopLine,
            StartColumn = selection.TopPoint.DisplayColumn,
            EndLine = selection.BottomLine,
            EndColumn = selection.BottomPoint.DisplayColumn,
            DocumentPath = doc.FullName
        };
    }

    public async Task<bool> SetSelectionAsync(string path, int startLine, int startColumn, int endLine, int endColumn)
    {
        await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();
        var dte = await GetDteAsync();

        foreach (Document doc in dte.Documents)
        {
            try
            {
                if (PathsEqual(doc.FullName, path))
                {
                    var textDoc = doc.Object("TextDocument") as TextDocument;
                    if (textDoc != null)
                    {
                        textDoc.Selection.MoveToLineAndOffset(startLine, startColumn);
                        textDoc.Selection.MoveToLineAndOffset(endLine, endColumn, true);
                        return true;
                    }
                }
            }
            catch (Exception ex)
            {
                VsixTelemetry.TrackException(ex);
            }
        }

        return false;
    }

    public async Task<bool> InsertTextAsync(string text)
    {
        await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();
        var dte = await GetDteAsync();

        var doc = dte.ActiveDocument;
        if (doc == null)
        {
            return false;
        }

        var textDoc = doc.Object("TextDocument") as TextDocument;
        if (textDoc == null)
        {
            return false;
        }

        textDoc.Selection.Insert(text);
        return true;
    }

    public async Task<bool> ReplaceTextAsync(string oldText, string newText)
    {
        await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();
        var dte = await GetDteAsync();

        var doc = dte.ActiveDocument;
        if (doc == null)
        {
            return false;
        }

        var textDoc = doc.Object("TextDocument") as TextDocument;
        if (textDoc == null)
        {
            return false;
        }

        var editPoint = textDoc.StartPoint.CreateEditPoint();
        var content = editPoint.GetText(textDoc.EndPoint);
        var newContent = content.Replace(oldText, newText);

        if (content != newContent)
        {
            editPoint.Delete(textDoc.EndPoint);
            editPoint.Insert(newContent);
            return true;
        }

        return false;
    }

    public async Task<bool> GoToLineAsync(int line)
    {
        await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();
        var dte = await GetDteAsync();

        var doc = dte.ActiveDocument;
        if (doc == null)
        {
            return false;
        }

        var textDoc = doc.Object("TextDocument") as TextDocument;
        if (textDoc == null)
        {
            return false;
        }

        textDoc.Selection.GotoLine(line);
        return true;
    }

    public async Task<List<FindResult>> FindAsync(string searchText, bool matchCase = false, bool wholeWord = false)
    {
        await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();
        var dte = await GetDteAsync();
        var results = new List<FindResult>();

        var doc = dte.ActiveDocument;
        if (doc == null)
        {
            return results;
        }

        var textDoc = doc.Object("TextDocument") as TextDocument;
        if (textDoc == null)
        {
            return results;
        }

        var editPoint = textDoc.StartPoint.CreateEditPoint();
        var content = editPoint.GetText(textDoc.EndPoint);
        var comparison = matchCase ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;

        var lines = content.Split('\n');
        for (int i = 0; i < lines.Length; i++)
        {
            var line = lines[i];
            var index = 0;
            while ((index = line.IndexOf(searchText, index, comparison)) >= 0)
            {
                results.Add(new FindResult
                {
                    Line = i + 1,
                    Column = index + 1,
                    Text = line.Trim(),
                    DocumentPath = doc.FullName
                });
                index += searchText.Length;
            }
        }

        return results;
    }

    public async Task<bool> BuildSolutionAsync()
    {
        using var activity = VsixTelemetry.Tracer.StartActivity("BuildSolution");

        await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();
        var dte = await GetDteAsync();

        try
        {
            dte.Solution.SolutionBuild.Build(true);
            return true;
        }
        catch (Exception ex)
        {
            activity?.SetStatus(ActivityStatusCode.Error, ex.Message);
            activity?.RecordException(ex);
            return false;
        }
    }

    public async Task<bool> BuildProjectAsync(string projectName)
    {
        using var activity = VsixTelemetry.Tracer.StartActivity("BuildProject");

        await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();
        var dte = await GetDteAsync();

        try
        {
            var config = dte.Solution.SolutionBuild.ActiveConfiguration.Name;
            var normalizedPath = NormalizePath(projectName);
            dte.Solution.SolutionBuild.BuildProject(config, normalizedPath, true);
            return true;
        }
        catch (Exception ex)
        {
            activity?.SetStatus(ActivityStatusCode.Error, ex.Message);
            activity?.RecordException(ex);
            return false;
        }
    }

    public async Task<bool> CleanSolutionAsync()
    {
        using var activity = VsixTelemetry.Tracer.StartActivity("CleanSolution");

        await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();
        var dte = await GetDteAsync();

        try
        {
            dte.Solution.SolutionBuild.Clean(true);
            return true;
        }
        catch (Exception ex)
        {
            activity?.SetStatus(ActivityStatusCode.Error, ex.Message);
            activity?.RecordException(ex);
            return false;
        }
    }

    public async Task<bool> CancelBuildAsync()
    {
        await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();
        var dte = await GetDteAsync();

        if (dte.Solution.SolutionBuild.BuildState != vsBuildState.vsBuildStateInProgress)
        {
            return false;
        }

        dte.ExecuteCommand("Build.Cancel");
        return true;
    }

    public async Task<BuildStatus> GetBuildStatusAsync()
    {
        await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();
        var dte = await GetDteAsync();

        var buildState = dte.Solution.SolutionBuild.BuildState;

        if (buildState == vsBuildState.vsBuildStateNotStarted)
        {
            return new BuildStatus
            {
                State = "NoBuildPerformed",
                FailedProjects = 0
            };
        }

        var lastInfo = dte.Solution.SolutionBuild.LastBuildInfo;

        return new BuildStatus
        {
            State = buildState switch
            {
                vsBuildState.vsBuildStateInProgress => "InProgress",
                vsBuildState.vsBuildStateDone => "Done",
                _ => "Unknown"
            },
            FailedProjects = lastInfo
        };
    }

    public async Task<List<SymbolInfo>> GetDocumentSymbolsAsync(string path)
    {
        await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();
        var dte = await GetDteAsync();
        var symbols = new List<SymbolInfo>();

        if (dte.Solution == null)
        {
            return symbols;
        }

        var normalizedPath = NormalizePath(path);
        var projectItem = dte.Solution.FindProjectItem(normalizedPath);
        if (projectItem == null)
        {
            return symbols;
        }

        var fileCodeModel = projectItem.FileCodeModel;
        if (fileCodeModel == null)
        {
            return symbols;
        }

        ExtractSymbols(fileCodeModel.CodeElements, symbols, normalizedPath, string.Empty);
        return symbols;
    }

    private void ExtractSymbols(CodeElements elements, List<SymbolInfo> symbols, string filePath, string containerName)
    {
        ThreadHelper.ThrowIfNotOnUIThread();

        foreach (CodeElement element in elements)
        {
            try
            {
                var kind = MapElementKind(element.Kind);
                if (kind == SymbolKind.Unknown)
                {
                    if (element.Kind == vsCMElement.vsCMElementImportStmt ||
                        element.Kind == vsCMElement.vsCMElementAttribute ||
                        element.Kind == vsCMElement.vsCMElementParameter)
                    {
                        continue;
                    }
                }

                var startPoint = element.StartPoint;
                var endPoint = element.EndPoint;

                var symbolInfo = new SymbolInfo
                {
                    Name = element.Name,
                    FullName = element.FullName,
                    Kind = kind,
                    FilePath = filePath,
                    StartLine = startPoint.Line,
                    StartColumn = startPoint.LineCharOffset,
                    EndLine = endPoint.Line,
                    EndColumn = endPoint.LineCharOffset,
                    ContainerName = containerName
                };

                var childElements = GetChildElements(element);
                if (childElements != null && childElements.Count > 0)
                {
                    ExtractSymbols(childElements, symbolInfo.Children, filePath, element.Name);
                }

                symbols.Add(symbolInfo);
            }
            catch (Exception ex)
            {
                VsixTelemetry.TrackException(ex);
            }
        }
    }

    private static CodeElements? GetChildElements(CodeElement element)
    {
        ThreadHelper.ThrowIfNotOnUIThread();

        try
        {
            return element.Kind switch
            {
                vsCMElement.vsCMElementNamespace => ((CodeNamespace)element).Members,
                vsCMElement.vsCMElementClass => ((CodeClass)element).Members,
                vsCMElement.vsCMElementStruct => ((CodeStruct)element).Members,
                vsCMElement.vsCMElementInterface => ((CodeInterface)element).Members,
                vsCMElement.vsCMElementEnum => ((CodeEnum)element).Members,
                _ => null
            };
        }
        catch
        {
            return null;
        }
    }

    private static SymbolKind MapElementKind(vsCMElement kind) => kind switch
    {
        vsCMElement.vsCMElementNamespace => SymbolKind.Namespace,
        vsCMElement.vsCMElementClass => SymbolKind.Class,
        vsCMElement.vsCMElementStruct => SymbolKind.Struct,
        vsCMElement.vsCMElementInterface => SymbolKind.Interface,
        vsCMElement.vsCMElementEnum => SymbolKind.Enum,
        vsCMElement.vsCMElementFunction => SymbolKind.Function,
        vsCMElement.vsCMElementProperty => SymbolKind.Property,
        vsCMElement.vsCMElementVariable => SymbolKind.Field,
        vsCMElement.vsCMElementEvent => SymbolKind.Event,
        vsCMElement.vsCMElementDelegate => SymbolKind.Delegate,
        _ => SymbolKind.Unknown
    };

    public async Task<WorkspaceSymbolResult> SearchWorkspaceSymbolsAsync(string query, int maxResults = 100)
    {
        await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();
        var dte = await GetDteAsync();
        var result = new WorkspaceSymbolResult();

        if (dte.Solution == null || string.IsNullOrWhiteSpace(query))
        {
            return result;
        }

        var allSymbols = new List<SymbolInfo>();
        var lowerQuery = query.ToLowerInvariant();

        foreach (EnvDTE.Project project in dte.Solution.Projects)
        {
            try
            {
                CollectProjectSymbols(project.ProjectItems, allSymbols, lowerQuery, maxResults * 2);
                if (allSymbols.Count >= maxResults * 2)
                {
                    break;
                }
            }
            catch (Exception ex)
            {
                VsixTelemetry.TrackException(ex);
            }
        }

        var matchingSymbols = allSymbols
            .Where(s => s.Name.ToLowerInvariant().Contains(lowerQuery) ||
                       s.FullName.ToLowerInvariant().Contains(lowerQuery))
            .Take(maxResults)
            .ToList();

        result.Symbols = matchingSymbols;
        result.TotalCount = allSymbols.Count;
        result.Truncated = allSymbols.Count > maxResults;

        return result;
    }

    private void CollectProjectSymbols(ProjectItems? items, List<SymbolInfo> allSymbols, string query, int limit)
    {
        ThreadHelper.ThrowIfNotOnUIThread();

        if (items == null || allSymbols.Count >= limit)
        {
            return;
        }

        foreach (ProjectItem item in items)
        {
            try
            {
                if (item.FileCodeModel != null)
                {
                    var filePath = item.FileNames[1];
                    CollectCodeElements(item.FileCodeModel.CodeElements, allSymbols, filePath, string.Empty, query, limit);
                }

                if (item.ProjectItems != null && item.ProjectItems.Count > 0)
                {
                    CollectProjectSymbols(item.ProjectItems, allSymbols, query, limit);
                }
            }
            catch (Exception ex)
            {
                VsixTelemetry.TrackException(ex);
            }
        }
    }

    private void CollectCodeElements(CodeElements elements, List<SymbolInfo> allSymbols, string filePath, string containerName, string query, int limit)
    {
        ThreadHelper.ThrowIfNotOnUIThread();

        if (allSymbols.Count >= limit)
        {
            return;
        }

        foreach (CodeElement element in elements)
        {
            try
            {
                var kind = MapElementKind(element.Kind);
                if (kind == SymbolKind.Unknown)
                {
                    continue;
                }

                var lowerName = element.Name.ToLowerInvariant();
                var lowerFullName = element.FullName.ToLowerInvariant();

                if (lowerName.Contains(query) || lowerFullName.Contains(query))
                {
                    var startPoint = element.StartPoint;
                    var endPoint = element.EndPoint;

                    allSymbols.Add(new SymbolInfo
                    {
                        Name = element.Name,
                        FullName = element.FullName,
                        Kind = kind,
                        FilePath = filePath,
                        StartLine = startPoint.Line,
                        StartColumn = startPoint.LineCharOffset,
                        EndLine = endPoint.Line,
                        EndColumn = endPoint.LineCharOffset,
                        ContainerName = containerName
                    });
                }

                var childElements = GetChildElements(element);
                if (childElements != null)
                {
                    CollectCodeElements(childElements, allSymbols, filePath, element.Name, query, limit);
                }
            }
            catch (Exception ex)
            {
                VsixTelemetry.TrackException(ex);
            }
        }
    }

    public async Task<DefinitionResult> GoToDefinitionAsync(string path, int line, int column)
    {
        await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();
        var dte = await GetDteAsync();
        var result = new DefinitionResult();

        try
        {
            var opened = await OpenDocumentAsync(path);
            if (!opened)
            {
                return result;
            }

            var doc = dte.ActiveDocument;
            if (doc == null)
            {
                return result;
            }

            var textDoc = doc.Object("TextDocument") as TextDocument;
            if (textDoc == null)
            {
                return result;
            }

            textDoc.Selection.MoveToLineAndOffset(line, column);

            var originalPath = doc.FullName;
            var originalLine = textDoc.Selection.ActivePoint.Line;

            dte.ExecuteCommand("Edit.GoToDefinition");

            await Task.Delay(100);

            var newDoc = dte.ActiveDocument;
            if (newDoc != null)
            {
                var newTextDoc = newDoc.Object("TextDocument") as TextDocument;
                if (newTextDoc != null)
                {
                    var newPath = newDoc.FullName;
                    var newLine = newTextDoc.Selection.ActivePoint.Line;
                    var newColumn = newTextDoc.Selection.ActivePoint.LineCharOffset;

                    if (!PathsEqual(newPath, originalPath) || newLine != originalLine)
                    {
                        result.Found = true;
                        result.SymbolName = GetWordAtPosition(textDoc, line, column);

                        var editPoint = newTextDoc.StartPoint.CreateEditPoint();
                        editPoint.MoveToLineAndOffset(newLine, 1);
                        var lineText = editPoint.GetLines(newLine, newLine + 1).Trim();

                        result.Definitions.Add(new LocationInfo
                        {
                            FilePath = newPath,
                            Line = newLine,
                            Column = newColumn,
                            EndLine = newLine,
                            EndColumn = newColumn,
                            Preview = lineText
                        });
                    }
                }
            }
        }
        catch (Exception ex)
        {
            VsixTelemetry.TrackException(ex);
        }

        return result;
    }

    private static string GetWordAtPosition(TextDocument textDoc, int line, int column)
    {
        ThreadHelper.ThrowIfNotOnUIThread();

        try
        {
            var editPoint = textDoc.StartPoint.CreateEditPoint();
            editPoint.MoveToLineAndOffset(line, column);

            var startPoint = editPoint.CreateEditPoint();
            startPoint.WordLeft(1);
            var endPoint = editPoint.CreateEditPoint();
            endPoint.WordRight(1);

            return startPoint.GetText(endPoint).Trim();
        }
        catch
        {
            return string.Empty;
        }
    }

    public async Task<ReferencesResult> FindReferencesAsync(string path, int line, int column, int maxResults = 100)
    {
        await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();
        var dte = await GetDteAsync();
        var result = new ReferencesResult();

        try
        {
            var opened = await OpenDocumentAsync(path);
            if (!opened)
            {
                return result;
            }

            var doc = dte.ActiveDocument;
            if (doc == null)
            {
                return result;
            }

            var textDoc = doc.Object("TextDocument") as TextDocument;
            if (textDoc == null)
            {
                return result;
            }

            textDoc.Selection.MoveToLineAndOffset(line, column);
            var symbolName = GetWordAtPosition(textDoc, line, column);

            if (string.IsNullOrWhiteSpace(symbolName))
            {
                return result;
            }

            result.SymbolName = symbolName;

            var references = await FindInSolutionAsync(dte, symbolName, maxResults);
            result.References = references;
            result.TotalCount = references.Count;
            result.Found = references.Count > 0;
            result.Truncated = references.Count >= maxResults;
        }
        catch (Exception ex)
        {
            VsixTelemetry.TrackException(ex);
        }

        return result;
    }

    private async Task<List<LocationInfo>> FindInSolutionAsync(DTE2 dte, string searchText, int maxResults)
    {
        await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();

        var locations = new List<LocationInfo>();

        if (dte.Solution == null)
        {
            return locations;
        }

        foreach (EnvDTE.Project project in dte.Solution.Projects)
        {
            try
            {
                await SearchProjectItemsAsync(project.ProjectItems, searchText, locations, maxResults);
                if (locations.Count >= maxResults)
                {
                    break;
                }
            }
            catch (Exception ex)
            {
                VsixTelemetry.TrackException(ex);
            }
        }

        return locations;
    }

    private async Task SearchProjectItemsAsync(ProjectItems? items, string searchText, List<LocationInfo> locations, int maxResults)
    {
        await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();

        if (items == null || locations.Count >= maxResults)
        {
            return;
        }

        foreach (ProjectItem item in items)
        {
            try
            {
                if (item.FileNames[1] is string filePath &&
                    (filePath.EndsWith(".cs", StringComparison.OrdinalIgnoreCase) ||
                     filePath.EndsWith(".vb", StringComparison.OrdinalIgnoreCase)))
                {
                    var content = await Task.Run(() =>
                    {
                        if (File.Exists(filePath))
                        {
                            return File.ReadAllText(filePath);
                        }
                        return null;
                    });

                    if (content != null)
                    {
                        var lines = content.Split('\n');
                        for (int i = 0; i < lines.Length && locations.Count < maxResults; i++)
                        {
                            var lineText = lines[i];
                            var index = 0;
                            while ((index = lineText.IndexOf(searchText, index, StringComparison.Ordinal)) >= 0 &&
                                   locations.Count < maxResults)
                            {
                                if (IsWordBoundary(lineText, index, searchText.Length))
                                {
                                    locations.Add(new LocationInfo
                                    {
                                        FilePath = filePath,
                                        Line = i + 1,
                                        Column = index + 1,
                                        EndLine = i + 1,
                                        EndColumn = index + 1 + searchText.Length,
                                        Preview = lineText.Trim()
                                    });
                                }
                                index += searchText.Length;
                            }
                        }
                    }
                }

                if (item.ProjectItems != null && item.ProjectItems.Count > 0)
                {
                    await SearchProjectItemsAsync(item.ProjectItems, searchText, locations, maxResults);
                }
            }
            catch (Exception ex)
            {
                VsixTelemetry.TrackException(ex);
            }
        }
    }

    private static bool IsWordBoundary(string text, int start, int length)
    {
        var beforeOk = start == 0 || !char.IsLetterOrDigit(text[start - 1]);
        var afterOk = start + length >= text.Length || !char.IsLetterOrDigit(text[start + length]);
        return beforeOk && afterOk;
    }

    public async Task<DebuggerStatus> GetDebuggerStatusAsync()
    {
        await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();
        var dte = await GetDteAsync();
        var debugger = (Debugger2)dte.Debugger;

        var status = new DebuggerStatus
        {
            Mode = debugger.CurrentMode switch
            {
                dbgDebugMode.dbgDesignMode => "Design",
                dbgDebugMode.dbgBreakMode => "Break",
                dbgDebugMode.dbgRunMode => "Run",
                _ => "Unknown"
            },
            IsDebugging = debugger.CurrentMode != dbgDebugMode.dbgDesignMode,
            LastBreakReason = debugger.LastBreakReason switch
            {
                dbgEventReason.dbgEventReasonNone => "None",
                dbgEventReason.dbgEventReasonBreakpoint => "Breakpoint",
                dbgEventReason.dbgEventReasonStep => "Step",
                dbgEventReason.dbgEventReasonUserBreak => "UserBreak",
                dbgEventReason.dbgEventReasonExceptionThrown => "ExceptionThrown",
                dbgEventReason.dbgEventReasonExceptionNotHandled => "ExceptionNotHandled",
                dbgEventReason.dbgEventReasonStopDebugging => "StopDebugging",
                dbgEventReason.dbgEventReasonGo => "Go",
                dbgEventReason.dbgEventReasonAttachProgram => "AttachProgram",
                dbgEventReason.dbgEventReasonDetachProgram => "DetachProgram",
                dbgEventReason.dbgEventReasonLaunchProgram => "LaunchProgram",
                dbgEventReason.dbgEventReasonEndProgram => "EndProgram",
                dbgEventReason.dbgEventReasonContextSwitch => "ContextSwitch",
                _ => "Unknown"
            }
        };

        try
        {
            if (debugger.DebuggedProcesses?.Count > 0)
            {
                var process = debugger.DebuggedProcesses.Item(1);
                status.CurrentProcessName = process.Name;
            }
        }
        catch (Exception ex)
        {
            VsixTelemetry.TrackException(ex);
        }

        try
        {
            if (debugger.CurrentMode == dbgDebugMode.dbgBreakMode && debugger.CurrentStackFrame != null)
            {
                var frame = (EnvDTE90a.StackFrame2)debugger.CurrentStackFrame;
                status.CurrentFunction = frame.FunctionName;
                status.CurrentFile = frame.FileName;
                status.CurrentLine = (int)frame.LineNumber;
            }
        }
        catch (Exception ex)
        {
            VsixTelemetry.TrackException(ex);
        }

        return status;
    }

    public async Task<bool> DebugLaunchAsync()
    {
        using var activity = VsixTelemetry.Tracer.StartActivity("DebugLaunch");

        await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();
        var dte = await GetDteAsync();

        try
        {
            dte.ExecuteCommand("Debug.Start");
            return true;
        }
        catch (Exception ex)
        {
            activity?.SetStatus(ActivityStatusCode.Error, ex.Message);
            activity?.RecordException(ex);
            return false;
        }
    }

    public async Task<bool> DebugLaunchWithoutDebuggingAsync()
    {
        using var activity = VsixTelemetry.Tracer.StartActivity("DebugLaunchWithoutDebugging");

        await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();
        var dte = await GetDteAsync();

        try
        {
            dte.ExecuteCommand("Debug.StartWithoutDebugging");
            return true;
        }
        catch (Exception ex)
        {
            activity?.SetStatus(ActivityStatusCode.Error, ex.Message);
            activity?.RecordException(ex);
            return false;
        }
    }

    public async Task<bool> DebugContinueAsync()
    {
        await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();
        var dte = await GetDteAsync();
        var debugger = (Debugger2)dte.Debugger;

        if (debugger.CurrentMode != dbgDebugMode.dbgBreakMode)
        {
            return false;
        }

        try
        {
            debugger.Go(false);
            return true;
        }
        catch (Exception ex)
        {
            VsixTelemetry.TrackException(ex);
            return false;
        }
    }

    public async Task<bool> DebugBreakAsync()
    {
        await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();
        var dte = await GetDteAsync();
        var debugger = (Debugger2)dte.Debugger;

        if (debugger.CurrentMode != dbgDebugMode.dbgRunMode)
        {
            return false;
        }

        try
        {
            debugger.Break(false);
            return true;
        }
        catch (Exception ex)
        {
            VsixTelemetry.TrackException(ex);
            return false;
        }
    }

    public async Task<bool> DebugStopAsync()
    {
        await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();
        var dte = await GetDteAsync();
        var debugger = (Debugger2)dte.Debugger;

        if (debugger.CurrentMode == dbgDebugMode.dbgDesignMode)
        {
            return false;
        }

        try
        {
            debugger.Stop(false);
            return true;
        }
        catch (Exception ex)
        {
            VsixTelemetry.TrackException(ex);
            return false;
        }
    }

    public async Task<bool> DebugStepOverAsync()
    {
        await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();
        var dte = await GetDteAsync();
        var debugger = (Debugger2)dte.Debugger;

        if (debugger.CurrentMode != dbgDebugMode.dbgBreakMode)
        {
            return false;
        }

        try
        {
            debugger.StepOver(false);
            return true;
        }
        catch (Exception ex)
        {
            VsixTelemetry.TrackException(ex);
            return false;
        }
    }

    public async Task<bool> DebugStepIntoAsync()
    {
        await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();
        var dte = await GetDteAsync();
        var debugger = (Debugger2)dte.Debugger;

        if (debugger.CurrentMode != dbgDebugMode.dbgBreakMode)
        {
            return false;
        }

        try
        {
            debugger.StepInto(false);
            return true;
        }
        catch (Exception ex)
        {
            VsixTelemetry.TrackException(ex);
            return false;
        }
    }

    public async Task<bool> DebugStepOutAsync()
    {
        await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();
        var dte = await GetDteAsync();
        var debugger = (Debugger2)dte.Debugger;

        if (debugger.CurrentMode != dbgDebugMode.dbgBreakMode)
        {
            return false;
        }

        try
        {
            debugger.StepOut(false);
            return true;
        }
        catch (Exception ex)
        {
            VsixTelemetry.TrackException(ex);
            return false;
        }
    }

    public async Task<bool> DebugAddBreakpointAsync(string file, int line)
    {
        using var activity = VsixTelemetry.Tracer.StartActivity("DebugAddBreakpoint");

        await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();
        var dte = await GetDteAsync();
        var debugger = (Debugger2)dte.Debugger;

        try
        {
            var normalizedPath = NormalizePath(file);
            debugger.Breakpoints.Add(Function: "", File: normalizedPath, Line: line);
            return true;
        }
        catch (Exception ex)
        {
            activity?.SetStatus(ActivityStatusCode.Error, ex.Message);
            activity?.RecordException(ex);
            return false;
        }
    }

    public async Task<bool> DebugRemoveBreakpointAsync(string file, int line)
    {
        using var activity = VsixTelemetry.Tracer.StartActivity("DebugRemoveBreakpoint");

        await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();
        var dte = await GetDteAsync();
        var debugger = (Debugger2)dte.Debugger;

        try
        {
            var breakpoints = debugger.Breakpoints;
            for (int i = breakpoints.Count; i >= 1; i--)
            {
                var bp = breakpoints.Item(i);
                if (bp.File != null && PathsEqual(bp.File, file) && bp.FileLine == line)
                {
                    bp.Delete();
                    return true;
                }
            }

            return false;
        }
        catch (Exception ex)
        {
            activity?.SetStatus(ActivityStatusCode.Error, ex.Message);
            activity?.RecordException(ex);
            return false;
        }
    }

    public async Task<List<BreakpointInfo>> DebugGetBreakpointsAsync()
    {
        await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();
        var dte = await GetDteAsync();
        var debugger = (Debugger2)dte.Debugger;
        var results = new List<BreakpointInfo>();

        try
        {
            foreach (Breakpoint bp in debugger.Breakpoints)
            {
                try
                {
                    results.Add(new BreakpointInfo
                    {
                        File = bp.File ?? string.Empty,
                        Line = bp.FileLine,
                        Column = bp.FileColumn,
                        FunctionName = bp.FunctionName ?? string.Empty,
                        Condition = bp.Condition ?? string.Empty,
                        Enabled = bp.Enabled,
                        CurrentHits = bp.CurrentHits
                    });
                }
                catch (Exception ex)
                {
                    VsixTelemetry.TrackException(ex);
                }
            }
        }
        catch (Exception ex)
        {
            VsixTelemetry.TrackException(ex);
        }

        return results;
    }

    public async Task<List<LocalVariableInfo>> DebugGetLocalsAsync()
    {
        await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();
        var dte = await GetDteAsync();
        var debugger = (Debugger2)dte.Debugger;
        var results = new List<LocalVariableInfo>();

        if (debugger.CurrentMode != dbgDebugMode.dbgBreakMode)
        {
            return results;
        }

        try
        {
            var frame = debugger.CurrentStackFrame;
            if (frame == null)
            {
                return results;
            }

            foreach (Expression local in frame.Locals)
            {
                try
                {
                    results.Add(new LocalVariableInfo
                    {
                        Name = local.Name,
                        Value = local.Value ?? string.Empty,
                        Type = local.Type ?? string.Empty,
                        IsValidValue = local.IsValidValue
                    });
                }
                catch (Exception ex)
                {
                    VsixTelemetry.TrackException(ex);
                }
            }
        }
        catch (Exception ex)
        {
            VsixTelemetry.TrackException(ex);
        }

        return results;
    }

    public async Task<List<CallStackFrameInfo>> DebugGetCallStackAsync()
    {
        await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();
        var dte = await GetDteAsync();
        var debugger = (Debugger2)dte.Debugger;
        var results = new List<CallStackFrameInfo>();

        if (debugger.CurrentMode != dbgDebugMode.dbgBreakMode)
        {
            return results;
        }

        try
        {
            var thread = debugger.CurrentThread;
            if (thread == null)
            {
                return results;
            }

            var depth = 0;
            foreach (EnvDTE.StackFrame frame in thread.StackFrames)
            {
                try
                {
                    var info = new CallStackFrameInfo
                    {
                        Depth = depth,
                        FunctionName = frame.FunctionName,
                        Module = frame.Module ?? string.Empty,
                        Language = frame.Language ?? string.Empty,
                        ReturnType = frame.ReturnType ?? string.Empty
                    };

                    if (frame is EnvDTE90a.StackFrame2 frame2)
                    {
                        info.FileName = frame2.FileName ?? string.Empty;
                        info.LineNumber = (int)frame2.LineNumber;
                    }

                    results.Add(info);
                    depth++;
                }
                catch (Exception ex)
                {
                    VsixTelemetry.TrackException(ex);
                }
            }
        }
        catch (Exception ex)
        {
            VsixTelemetry.TrackException(ex);
        }

        return results;
    }
}
