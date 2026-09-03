using System.Collections.Generic;

namespace CodingWithCalvin.MCPServer.Shared.Models;

/// <summary>
/// Outcome of opening a Visual Studio integrated terminal.
/// </summary>
public class TerminalResult
{
    public bool Created { get; set; }

    /// <summary>
    /// Identifier for the new terminal, used by the show and close tools.
    /// </summary>
    public string? TerminalId { get; set; }

    /// <summary>
    /// Command the terminal was launched with, when one was supplied.
    /// </summary>
    public string? Command { get; set; }

    public string? WorkingDirectory { get; set; }

    /// <summary>
    /// True when the command was launched inside the Visual Studio developer environment, so
    /// that tooling such as msbuild and dotnet-coverage is on PATH.
    /// </summary>
    public bool DeveloperEnvironment { get; set; }

    /// <summary>
    /// Populated when <see cref="Created"/> is false, or to carry a caveat about a terminal
    /// that was created successfully.
    /// </summary>
    public string? Message { get; set; }
}

/// <summary>
/// The integrated terminals currently open in Visual Studio.
/// </summary>
public class TerminalListResult
{
    /// <summary>
    /// False when the terminal service could not be reached at all, which is distinct from
    /// there being no terminals open.
    /// </summary>
    public bool Available { get; set; }

    public List<string> TerminalIds { get; set; } = new();

    public string? Message { get; set; }
}
