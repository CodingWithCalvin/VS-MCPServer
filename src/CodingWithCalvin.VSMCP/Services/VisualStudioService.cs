using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using CodingWithCalvin.VSMCP.Shared.Models;
using Community.VisualStudio.Toolkit;
using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;

namespace CodingWithCalvin.VSMCP.Services;

public class VisualStudioService
{
    private readonly AsyncPackage _package;

    public VisualStudioService(AsyncPackage package)
    {
        _package = package;
    }

    private async Task<DTE2> GetDteAsync()
    {
        await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();
        return await _package.GetServiceAsync(typeof(DTE)) as DTE2
            ?? throw new InvalidOperationException("Could not get DTE service");
    }

    #region Solution Operations

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
        await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();
        var dte = await GetDteAsync();

        try
        {
            dte.Solution.Open(path);
            return true;
        }
        catch
        {
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
            catch
            {
                // Skip projects that throw (e.g., solution folders)
            }
        }

        return projects;
    }

    #endregion

    #region Document Operations

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
            catch
            {
                // Skip documents that throw
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
        await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();
        var dte = await GetDteAsync();

        try
        {
            dte.ItemOperations.OpenFile(path);
            return true;
        }
        catch
        {
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
                if (doc.FullName.Equals(path, StringComparison.OrdinalIgnoreCase))
                {
                    doc.Close(save ? vsSaveChanges.vsSaveChangesYes : vsSaveChanges.vsSaveChangesNo);
                    return true;
                }
            }
            catch
            {
                // Continue checking other documents
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
                if (doc.FullName.Equals(path, StringComparison.OrdinalIgnoreCase))
                {
                    var textDoc = doc.Object("TextDocument") as TextDocument;
                    if (textDoc != null)
                    {
                        var editPoint = textDoc.StartPoint.CreateEditPoint();
                        return editPoint.GetText(textDoc.EndPoint);
                    }
                }
            }
            catch
            {
                // Continue checking other documents
            }
        }

        // Document not open, read from file
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
                if (doc.FullName.Equals(path, StringComparison.OrdinalIgnoreCase))
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
            catch
            {
                // Continue checking other documents
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
                if (doc.FullName.Equals(path, StringComparison.OrdinalIgnoreCase))
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
            catch
            {
                // Continue checking other documents
            }
        }

        return false;
    }

    #endregion

    #region Editor Operations

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

    #endregion

    #region Build Operations

    public async Task<bool> BuildSolutionAsync()
    {
        await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();
        var dte = await GetDteAsync();

        try
        {
            dte.Solution.SolutionBuild.Build(true);
            return true;
        }
        catch
        {
            return false;
        }
    }

    public async Task<bool> BuildProjectAsync(string projectName)
    {
        await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();
        var dte = await GetDteAsync();

        try
        {
            var config = dte.Solution.SolutionBuild.ActiveConfiguration.Name;
            dte.Solution.SolutionBuild.BuildProject(config, projectName, true);
            return true;
        }
        catch
        {
            return false;
        }
    }

    public async Task<bool> CleanSolutionAsync()
    {
        await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();
        var dte = await GetDteAsync();

        try
        {
            dte.Solution.SolutionBuild.Clean(true);
            return true;
        }
        catch
        {
            return false;
        }
    }

    public async Task CancelBuildAsync()
    {
        await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();
        var dte = await GetDteAsync();
        dte.ExecuteCommand("Build.Cancel");
    }

    public async Task<BuildStatus> GetBuildStatusAsync()
    {
        await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();
        var dte = await GetDteAsync();

        var buildState = dte.Solution.SolutionBuild.BuildState;
        var lastInfo = dte.Solution.SolutionBuild.LastBuildInfo;

        return new BuildStatus
        {
            State = buildState switch
            {
                vsBuildState.vsBuildStateNotStarted => "NotStarted",
                vsBuildState.vsBuildStateInProgress => "InProgress",
                vsBuildState.vsBuildStateDone => "Done",
                _ => "Unknown"
            },
            FailedProjects = lastInfo
        };
    }

    #endregion
}
