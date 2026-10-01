using System.Collections.Generic;

namespace CodingWithCalvin.MCPServer.Shared.Models;

/// <summary>
/// A modal dialog that Visual Studio is waiting on.
/// </summary>
public class DialogInfo
{
    /// <summary>
    /// Window handle of the dialog in hexadecimal, used to pick a dialog when more than one is open.
    /// </summary>
    public string Id { get; set; } = string.Empty;

    public string Title { get; set; } = string.Empty;

    /// <summary>
    /// The dialog's text, usually the question being asked or the error being reported.
    /// </summary>
    public string Message { get; set; } = string.Empty;

    public List<DialogButtonInfo> Buttons { get; set; } = new();

    /// <summary>
    /// Populated when the dialog's contents could not be read, for example because it stopped
    /// responding. The dialog is still listed so the caller knows something is blocking.
    /// </summary>
    public string? Note { get; set; }
}

public class DialogButtonInfo
{
    public string Name { get; set; } = string.Empty;
    public bool IsEnabled { get; set; }
}

/// <summary>
/// The modal dialogs currently open in Visual Studio.
/// </summary>
public class DialogListResult
{
    public List<DialogInfo> Dialogs { get; set; } = new();

    /// <summary>
    /// Populated when the dialogs could not be enumerated at all, which is distinct from there
    /// being none open.
    /// </summary>
    public string? Message { get; set; }
}

/// <summary>
/// Outcome of answering a modal dialog by clicking one of its buttons.
/// </summary>
public class DialogResponseResult
{
    public bool Clicked { get; set; }

    public string? DialogId { get; set; }

    /// <summary>
    /// Name of the button that was clicked, as the dialog displays it.
    /// </summary>
    public string? Button { get; set; }

    /// <summary>
    /// True when the dialog had closed by the time this result was produced.
    /// </summary>
    public bool DialogClosed { get; set; }

    public string? Message { get; set; }

    /// <summary>
    /// Dialogs still open afterwards. Answering one dialog can raise a follow-up prompt.
    /// </summary>
    public List<DialogInfo> OpenDialogs { get; set; } = new();
}
