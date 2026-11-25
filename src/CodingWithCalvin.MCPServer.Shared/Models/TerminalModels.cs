using System.Collections.Generic;

namespace CodingWithCalvin.MCPServer.Shared.Models;

/// <summary>
/// Outcome of opening a Visual Studio integrated terminal.
/// </summary>
public class TerminalResult
{
    public bool Created { get; set; }

    /// <summary>
    /// True when the command was typed into a terminal that was already open rather than into a
    /// newly created one.
    /// </summary>
    public bool Reused { get; set; }

    /// <summary>
    /// Identifier for the terminal the command was submitted to, used by the show and close tools.
    /// </summary>
    public string? TerminalId { get; set; }

    /// <summary>
    /// Command that was submitted to the terminal, when one was supplied.
    /// </summary>
    public string? Command { get; set; }

    public string? WorkingDirectory { get; set; }

    /// <summary>
    /// Raw text the terminal printed while the command was being submitted, captured for the
    /// requested wait window. It is unprocessed PTY output: it still contains ANSI escape
    /// sequences, the echoed command and the shell prompt, and there is no exit code or
    /// indication that the command finished.
    /// </summary>
    public string? Output { get; set; }

    /// <summary>
    /// True when the terminal's default profile loads the Visual Studio developer environment,
    /// so that tooling such as msbuild and dotnet-coverage is on PATH. Unknown (false) when a
    /// command is submitted to a terminal that was opened by someone else.
    /// </summary>
    public bool DeveloperEnvironment { get; set; }

    /// <summary>
    /// Populated when the command could not be submitted, or to carry a caveat about a
    /// successful call.
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
