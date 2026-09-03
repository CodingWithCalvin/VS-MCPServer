using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using CodingWithCalvin.MCPServer.Shared.Models;
using CodingWithCalvin.Otel4Vsix;

namespace CodingWithCalvin.MCPServer.Services;

/// <summary>
/// Reflection bridge to the Visual Studio code coverage file reader.
/// </summary>
/// <remarks>
/// <para>
/// <c>Microsoft.CodeCoverage.IO.dll</c> lives under
/// <c>CommonExtensions\Microsoft\TestWindow\VsTest</c> in the Visual Studio installation and, in
/// contrast to the code coverage feature itself, is present in <b>both</b> VS 2022 and VS 2026
/// regardless of edition. It is not on NuGet and not in the Microsoft.VisualStudio.SDK
/// metapackage, and <c>$(DevEnvDir)</c> is undefined under <c>dotnet build</c>, so it is loaded
/// from disk and read reflectively - the same constraint documented on
/// <see cref="TestExplorerInterop"/> and <see cref="TerminalInterop"/>.
/// </para>
/// <para>
/// Unlike those two, this assembly has no binding redirect or codeBase entry in
/// devenv.exe.config, so it is loaded by path rather than by strong name. The path is derived
/// from VSAPPIDDIR, which devenv sets to its own Common7\IDE directory.
/// </para>
/// <para>
/// The V1 <c>CoverageFileUtility</c> is used deliberately: it has a parameterless constructor,
/// whereas <c>CoverageFileUtilityV2</c> requires an <c>ICoverageFileConfiguration</c> that
/// cannot be implemented without a compile-time reference. Both expose the same underlying data.
/// </para>
/// <para>
/// Per-module and per-function counts come straight from the reader - <c>Function</c> and
/// <c>ModuleWrapper</c> both carry coverage statistics, so nothing is recomputed. Only the class
/// level is derived here, by grouping functions on their declaring type.
/// </para>
/// </remarks>
internal sealed class CoverageInterop
{
    private const string AssemblyFileName = "Microsoft.CodeCoverage.IO.dll";
    private const string SimpleAssemblyName = "Microsoft.CodeCoverage.IO";
    private const string UtilityTypeName = "Microsoft.CodeCoverage.IO.CoverageFileUtility";

    internal const string UnavailableMessage =
        "The Visual Studio code coverage reader could not be loaded, so coverage results cannot "
        + "be parsed in this instance.";

    private readonly Func<Assembly?> _assemblyResolver;

    private bool _resolved;
    private Type? _utilityType;
    private MethodInfo? _readCoverageFile;

    internal CoverageInterop()
        : this(TryLoadCoverageAssembly)
    {
    }

    /// <summary>
    /// Overload taking an explicit assembly resolver, so that tests can supply stand-in coverage
    /// types without a Visual Studio installation.
    /// </summary>
    internal CoverageInterop(Func<Assembly?> assemblyResolver)
    {
        _assemblyResolver = assemblyResolver;
    }

    /// <summary>
    /// Reads a .coverage file and projects it into the module, class and method tree, trimmed to
    /// the requested depth.
    /// </summary>
    internal CoverageReportResult ReadReport(string coverageFile, CoverageDetail detail, string? filter)
    {
        if (string.IsNullOrWhiteSpace(coverageFile) || !File.Exists(coverageFile))
        {
            return new CoverageReportResult
            {
                CoverageFile = coverageFile,
                Message = $"Coverage file not found: {coverageFile}"
            };
        }

        try
        {
            if (!TryResolve())
            {
                return new CoverageReportResult { CoverageFile = coverageFile, Message = UnavailableMessage };
            }

            var utility = Activator.CreateInstance(_utilityType!);
            if (utility == null)
            {
                return new CoverageReportResult { CoverageFile = coverageFile, Message = UnavailableMessage };
            }

            var data = _readCoverageFile!.Invoke(utility, new object?[] { coverageFile });
            var modules = data?.GetType().GetProperty("Modules")?.GetValue(data) as IEnumerable;
            if (modules == null)
            {
                return new CoverageReportResult { CoverageFile = coverageFile, Message = UnavailableMessage };
            }

            var result = new CoverageReportResult { Available = true, CoverageFile = coverageFile };

            foreach (var module in modules)
            {
                var projected = ProjectModule(module, detail, filter);
                if (projected != null)
                {
                    result.Modules.Add(projected);
                }
            }

            result.Summary = Combine(result.Modules.Select(m => m.Summary));

            if (result.Modules.Count == 0 && !string.IsNullOrWhiteSpace(filter))
            {
                result.Message = $"No module or class matched '{filter}'.";
            }

            return result;
        }
        catch (Exception ex)
        {
            VsixTelemetry.TrackException(ex);
            return new CoverageReportResult { CoverageFile = coverageFile, Message = ex.Message };
        }
    }

    private static CoverageModule? ProjectModule(object module, CoverageDetail detail, string? filter)
    {
        var name = GetString(module, "Name");
        var moduleMatches = Matches(name, filter);

        var functions = module.GetType().GetProperty("Functions")?.GetValue(module) as IEnumerable;
        var classes = functions == null
            ? new List<CoverageClass>()
            : GroupIntoClasses(functions, detail, moduleMatches ? null : filter);

        // A module survives when it matches the filter itself, or when one of its classes does.
        if (!moduleMatches && classes.Count == 0)
        {
            return null;
        }

        var projected = new CoverageModule
        {
            Name = name,
            Path = GetString(module, "Path"),
            TargetFramework = GetString(module, "TargetFramework"),
            Summary = Combine(classes.Select(c => c.Summary))
        };

        if (detail != CoverageDetail.Summary)
        {
            projected.Classes = classes;
        }

        return projected;
    }

    private static List<CoverageClass> GroupIntoClasses(
        IEnumerable functions,
        CoverageDetail detail,
        string? filter)
    {
        var grouped = new Dictionary<string, CoverageClass>(StringComparer.Ordinal);

        foreach (var function in functions)
        {
            var typeName = GetString(function, "TypeName");
            var namespaceName = GetString(function, "NamespaceName");
            var key = namespaceName + "." + typeName;

            if (!Matches(key, filter) && !Matches(typeName, filter))
            {
                continue;
            }

            if (!grouped.TryGetValue(key, out var entry))
            {
                entry = new CoverageClass { Name = typeName, Namespace = namespaceName };
                grouped[key] = entry;
            }

            var summary = ReadStatistics(function);

            if (detail == CoverageDetail.Method)
            {
                entry.Methods.Add(new CoverageMethod
                {
                    Name = GetString(function, "Name"),
                    Summary = summary
                });
            }

            Accumulate(entry.Summary, summary);
        }

        foreach (var entry in grouped.Values)
        {
            ApplyPercentages(entry.Summary);
        }

        return grouped.Values
            .OrderBy(c => c.Namespace, StringComparer.Ordinal)
            .ThenBy(c => c.Name, StringComparer.Ordinal)
            .ToList();
    }

    /// <summary>
    /// Reads the counts a coverage node inherits from CoverageStatistics.
    /// </summary>
    private static CoverageSummary ReadStatistics(object node)
    {
        var summary = new CoverageSummary
        {
            LinesCovered = GetInt(node, "LinesCovered"),
            LinesPartiallyCovered = GetInt(node, "LinesPartiallyCovered"),
            LinesNotCovered = GetInt(node, "LinesNotCovered"),
            BlocksCovered = GetInt(node, "BlocksCovered"),
            BlocksNotCovered = GetInt(node, "BlocksNotCovered")
        };

        ApplyPercentages(summary);
        return summary;
    }

    internal static void Accumulate(CoverageSummary target, CoverageSummary addition)
    {
        target.LinesCovered += addition.LinesCovered;
        target.LinesPartiallyCovered += addition.LinesPartiallyCovered;
        target.LinesNotCovered += addition.LinesNotCovered;
        target.BlocksCovered += addition.BlocksCovered;
        target.BlocksNotCovered += addition.BlocksNotCovered;
    }

    internal static CoverageSummary Combine(IEnumerable<CoverageSummary> summaries)
    {
        var combined = new CoverageSummary();

        foreach (var summary in summaries)
        {
            Accumulate(combined, summary);
        }

        ApplyPercentages(combined);
        return combined;
    }

    /// <summary>
    /// Partially covered lines count against the total rather than towards it, matching how
    /// Visual Studio reports line coverage.
    /// </summary>
    internal static void ApplyPercentages(CoverageSummary summary)
    {
        var totalLines = summary.LinesCovered + summary.LinesPartiallyCovered + summary.LinesNotCovered;
        summary.LineCoveragePercent = Percentage(summary.LinesCovered, totalLines);

        var totalBlocks = summary.BlocksCovered + summary.BlocksNotCovered;
        summary.BlockCoveragePercent = Percentage(summary.BlocksCovered, totalBlocks);
    }

    private static double Percentage(int covered, int total) =>
        total == 0 ? 0d : Math.Round(covered * 100d / total, 2);

    private static bool Matches(string value, string? filter) =>
        string.IsNullOrWhiteSpace(filter)
        || value.IndexOf(filter!, StringComparison.OrdinalIgnoreCase) >= 0;

    private static string GetString(object target, string property) =>
        target.GetType().GetProperty(property)?.GetValue(target) as string ?? string.Empty;

    private static int GetInt(object target, string property)
    {
        var value = target.GetType().GetProperty(property)?.GetValue(target);

        return value switch
        {
            uint unsigned => unsigned > int.MaxValue ? int.MaxValue : (int)unsigned,
            int signed => signed,
            _ => 0
        };
    }

    private bool TryResolve()
    {
        if (_resolved)
        {
            return _utilityType != null && _readCoverageFile != null;
        }

        _resolved = true;

        var assembly = _assemblyResolver();
        _utilityType = assembly?.GetType(UtilityTypeName);

        // ReadCoverageFile is overloaded; the single-path overload is the one wanted here.
        _readCoverageFile = _utilityType?.GetMethod("ReadCoverageFile", new[] { typeof(string) });

        return _utilityType != null && _readCoverageFile != null;
    }

    private static Assembly? TryLoadCoverageAssembly()
    {
        try
        {
            var loaded = AppDomain.CurrentDomain.GetAssemblies()
                .FirstOrDefault(a => string.Equals(
                    a.GetName().Name,
                    SimpleAssemblyName,
                    StringComparison.OrdinalIgnoreCase));
            if (loaded != null)
            {
                return loaded;
            }

            var path = TryFindCoverageAssemblyPath();
            return path == null ? null : Assembly.LoadFrom(path);
        }
        catch (Exception ex)
        {
            VsixTelemetry.TrackException(ex);
            return null;
        }
    }

    /// <summary>
    /// Locates the reader inside the running Visual Studio. VSAPPIDDIR points at Common7\IDE.
    /// </summary>
    internal static string? TryFindCoverageAssemblyPath()
    {
        try
        {
            var ideDirectory = Environment.GetEnvironmentVariable("VSAPPIDDIR");
            if (string.IsNullOrWhiteSpace(ideDirectory))
            {
                return null;
            }

            var path = Path.GetFullPath(Path.Combine(
                ideDirectory,
                "CommonExtensions",
                "Microsoft",
                "TestWindow",
                "VsTest",
                AssemblyFileName));

            return File.Exists(path) ? path : null;
        }
        catch (Exception ex)
        {
            VsixTelemetry.TrackException(ex);
            return null;
        }
    }

    /// <summary>
    /// Finds the most recent .coverage file beneath a directory. Visual Studio writes coverage
    /// results under the solution's TestResults folder, but the exact subdirectory carries a
    /// generated name, so the newest file is the only reliable handle on "the run that just
    /// finished".
    /// </summary>
    internal static string? FindNewestCoverageFile(string? searchRoot)
    {
        if (string.IsNullOrWhiteSpace(searchRoot) || !Directory.Exists(searchRoot))
        {
            return null;
        }

        try
        {
            return new DirectoryInfo(searchRoot)
                .GetFiles("*.coverage", SearchOption.AllDirectories)
                .OrderByDescending(f => f.LastWriteTimeUtc)
                .FirstOrDefault()
                ?.FullName;
        }
        catch (Exception ex)
        {
            VsixTelemetry.TrackException(ex);
            return null;
        }
    }
}

/// <summary>How deep a coverage report should be projected.</summary>
internal enum CoverageDetail
{
    Summary,
    Class,
    Method
}
