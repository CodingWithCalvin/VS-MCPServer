using System;
using System.Collections.Generic;
using System.ComponentModel.Composition;
using System.Diagnostics;
using System.IO;
using System.Threading.Tasks;
using CodingWithCalvin.MCPServer.Shared.Models;
using CodingWithCalvin.Otel4Vsix;
using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;

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
}
