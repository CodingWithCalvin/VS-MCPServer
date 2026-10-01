using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Automation;
using CodingWithCalvin.MCPServer.Shared.Models;

namespace CodingWithCalvin.MCPServer.Services;

/// <summary>
/// Finds the modal dialogs Visual Studio is waiting on, reads them, and clicks their buttons.
/// </summary>
/// <remarks>
/// <para>
/// A modal dialog holds the Visual Studio UI thread inside its own message loop, and every other
/// tool reaches Visual Studio through that thread. Nothing here may switch to it: these calls
/// arrive on the thread pool while the call that raised the dialog is still waiting, and have to
/// complete on their own. Every await therefore uses <c>ConfigureAwait(false)</c>.
/// </para>
/// <para>
/// Discovery uses Win32 window enumeration, which reads window state without sending messages
/// and so cannot hang on a busy UI thread. Reading a dialog and clicking its buttons use UI
/// Automation, which handles native message boxes, task dialogs and WPF dialogs alike. That does
/// send messages to the dialog, so it runs on a dedicated thread under a timeout; a dialog that
/// stops responding is still listed, just without its contents.
/// </para>
/// <para>
/// A window counts as a blocking dialog when it is visible and enabled and its owner has been
/// disabled, which is how both Win32 and WPF implement modality, or when it is a native dialog
/// with no owner. Windows belonging to other processes are included when a Visual Studio window
/// owns them, because some prompts are raised out of process.
/// </para>
/// </remarks>
internal sealed class DialogAutomation
{
    internal static readonly TimeSpan ReadTimeout = TimeSpan.FromSeconds(5);
    internal static readonly TimeSpan ClickTimeout = TimeSpan.FromSeconds(5);
    internal static readonly TimeSpan CloseTimeout = TimeSpan.FromSeconds(3);

    private static readonly TimeSpan ClosePollInterval = TimeSpan.FromMilliseconds(100);

    private const string NativeDialogClassName = "#32770";
    private const int MaxMessageLength = 4000;
    private const int MaxElementsVisited = 500;
    private const int MaxTreeDepth = 12;

    internal const string NoDialogMessage = "No dialog is open in Visual Studio.";

    /// <summary>
    /// Control types whose descendants are never part of the dialog's message. Their own labels
    /// would otherwise be read as text: a WPF button exposes its caption as a Text child.
    /// </summary>
    private static readonly HashSet<ControlType> SkippedSubtrees = new()
    {
        ControlType.TitleBar,
        ControlType.MenuBar,
        ControlType.ScrollBar,
        ControlType.CheckBox,
        ControlType.RadioButton,
        ControlType.Hyperlink,
        ControlType.ComboBox,
    };

    private readonly int _processId;

    internal DialogAutomation()
        : this((int)NativeMethods.GetCurrentProcessId())
    {
    }

    internal DialogAutomation(int processId)
    {
        _processId = processId;
    }

    internal async Task<DialogListResult> GetDialogsAsync()
    {
        var result = new DialogListResult();

        foreach (var window in FindBlockingDialogs())
        {
            var dialog = await DescribeAsync(window).ConfigureAwait(false);
            if (dialog != null)
            {
                result.Dialogs.Add(dialog);
            }
        }

        return result;
    }

    internal async Task<DialogResponseResult> RespondAsync(string button, string? dialogId)
    {
        var result = new DialogResponseResult { Button = button };

        if (!TrySelectDialog(FindBlockingDialogs(), dialogId, out var window, out var error))
        {
            result.Message = error;
            result.OpenDialogs = (await GetDialogsAsync().ConfigureAwait(false)).Dialogs;
            return result;
        }

        result.DialogId = FormatId(window!.Handle);

        DialogContent? content;

        try
        {
            var title = ReadTitle(window.Handle);
            var read = await RunWithTimeoutAsync(() => ReadContent(window.Handle, title), ReadTimeout)
                .ConfigureAwait(false);

            if (!read.Completed)
            {
                result.Message = "The dialog did not respond in time, so its buttons could not be read.";
                return result;
            }

            content = read.Result;
        }
        catch (ElementNotAvailableException)
        {
            result.DialogClosed = !IsOpen(window.Handle);
            result.Message = result.DialogClosed
                ? "The dialog closed before it could be answered."
                : "The dialog changed while it was being read. Call dialog_list and try again.";
            return result;
        }

        var index = FindButton(content.Buttons, button, out error);
        if (index < 0)
        {
            result.Message = error;
            result.OpenDialogs = (await GetDialogsAsync().ConfigureAwait(false)).Dialogs;
            return result;
        }

        result.Button = content.Buttons[index].Name;

        try
        {
            var element = content.Elements[index];
            var click = await RunWithTimeoutAsync(() => Invoke(element), ClickTimeout).ConfigureAwait(false);

            if (click.Completed && !click.Result)
            {
                result.Message = $"The \"{result.Button}\" button cannot be clicked through UI Automation.";
                return result;
            }

            result.Clicked = true;

            if (!click.Completed)
            {
                // Invoke delivers the click synchronously, so it only fails to return while the
                // button's handler is still running - almost always because it opened another
                // modal dialog of its own.
                result.Message = "The click was delivered but Visual Studio is still handling it. It may have opened another dialog.";
            }
        }
        catch (ElementNotEnabledException)
        {
            result.Message = $"The \"{result.Button}\" button is disabled.";
            return result;
        }
        catch (ElementNotAvailableException)
        {
            // The button, and usually the dialog with it, went away between reading and clicking.
            result.DialogClosed = !IsOpen(window.Handle);
            result.Message = "The button was no longer available to click.";
            result.OpenDialogs = (await GetDialogsAsync().ConfigureAwait(false)).Dialogs;
            return result;
        }

        result.DialogClosed = await WaitForCloseAsync(window.Handle).ConfigureAwait(false);

        if (!result.DialogClosed && result.Message == null)
        {
            result.Message = "The button was clicked but the dialog is still open.";
        }

        result.OpenDialogs = (await GetDialogsAsync().ConfigureAwait(false)).Dialogs;
        return result;
    }

    /// <summary>
    /// The Win32 facts that decide whether a top-level window is a blocking dialog. Captured
    /// separately from the decision so the rules can be tested without real windows.
    /// </summary>
    internal sealed class TopLevelWindow
    {
        public IntPtr Handle { get; set; }
        public int ProcessId { get; set; }
        public string ClassName { get; set; } = string.Empty;
        public bool IsVisible { get; set; }
        public bool IsEnabled { get; set; }

        /// <summary>
        /// False for windows that never take focus, such as tooltips and popups.
        /// </summary>
        public bool CanActivate { get; set; } = true;

        public IntPtr Owner { get; set; }
        public int OwnerProcessId { get; set; }
        public bool IsOwnerEnabled { get; set; }
    }

    internal static bool IsBlockingDialog(TopLevelWindow window, int processId)
    {
        if (!window.IsVisible || !window.IsEnabled || !window.CanActivate)
        {
            return false;
        }

        if (window.Owner == IntPtr.Zero)
        {
            // With no owner there is nothing to disable, but a native dialog box still runs a
            // modal loop. MessageBox called with a null owner is the usual case.
            return window.ProcessId == processId
                && string.Equals(window.ClassName, NativeDialogClassName, StringComparison.Ordinal);
        }

        // A floating tool window also has an owner, but that owner stays enabled.
        return !window.IsOwnerEnabled
            && (window.ProcessId == processId || window.OwnerProcessId == processId);
    }

    internal static bool TrySelectDialog(
        IReadOnlyList<TopLevelWindow> dialogs,
        string? dialogId,
        out TopLevelWindow? selected,
        out string? error)
    {
        selected = null;
        error = null;

        if (dialogs.Count == 0)
        {
            error = NoDialogMessage;
            return false;
        }

        if (string.IsNullOrWhiteSpace(dialogId))
        {
            if (dialogs.Count > 1)
            {
                error = $"{dialogs.Count} dialogs are open, so a dialog id is required. Call dialog_list to see them.";
                return false;
            }

            selected = dialogs[0];
            return true;
        }

        if (TryParseId(dialogId!, out var handle))
        {
            selected = dialogs.FirstOrDefault(dialog => dialog.Handle == handle);
        }

        if (selected == null)
        {
            error = $"No open dialog has the id {dialogId}. Call dialog_list to see the open dialogs.";
            return false;
        }

        return true;
    }

    /// <summary>
    /// Finds the button the caller named, returning its index or -1 with an explanation.
    /// </summary>
    internal static int FindButton(IReadOnlyList<DialogButtonInfo> buttons, string requested, out string? error)
    {
        error = null;
        var key = ToMatchKey(requested);

        var matches = Enumerable.Range(0, buttons.Count)
            .Where(i => string.Equals(ToMatchKey(buttons[i].Name), key, StringComparison.OrdinalIgnoreCase))
            .ToList();

        if (matches.Count == 0)
        {
            error = buttons.Count == 0
                ? "The dialog has no buttons that can be clicked."
                : $"The dialog has no button named \"{requested}\". Its buttons are: {string.Join(", ", buttons.Select(b => b.Name))}.";
            return -1;
        }

        if (matches.Count > 1)
        {
            error = $"The dialog has {matches.Count} buttons named \"{requested}\", so the one to click is ambiguous.";
            return -1;
        }

        var index = matches[0];
        if (!buttons[index].IsEnabled)
        {
            error = $"The \"{buttons[index].Name}\" button is disabled.";
            return -1;
        }

        return index;
    }

    /// <summary>
    /// Reduces a button caption to the form used for matching: no access key marker, no trailing
    /// ellipsis and no surrounding whitespace.
    /// </summary>
    internal static string ToMatchKey(string name)
    {
        var builder = new StringBuilder(name.Length);

        for (var i = 0; i < name.Length; i++)
        {
            if (name[i] == '&')
            {
                // "&&" is a literal ampersand; a single one only marks the access key.
                if (i + 1 < name.Length && name[i + 1] == '&')
                {
                    builder.Append('&');
                    i++;
                }

                continue;
            }

            builder.Append(name[i]);
        }

        var key = builder.ToString().Trim();

        if (key.EndsWith("...", StringComparison.Ordinal))
        {
            key = key.Substring(0, key.Length - 3);
        }
        else if (key.EndsWith("…", StringComparison.Ordinal))
        {
            key = key.Substring(0, key.Length - 1);
        }

        return key.TrimEnd();
    }

    internal static string FormatId(IntPtr handle) =>
        "0x" + handle.ToInt64().ToString("X8", CultureInfo.InvariantCulture);

    internal static bool TryParseId(string id, out IntPtr handle)
    {
        handle = IntPtr.Zero;
        var digits = id.Trim();

        if (digits.StartsWith("0x", StringComparison.OrdinalIgnoreCase))
        {
            digits = digits.Substring(2);
        }

        if (!long.TryParse(digits, NumberStyles.AllowHexSpecifier, CultureInfo.InvariantCulture, out var value)
            || value == 0)
        {
            return false;
        }

        handle = new IntPtr(value);
        return true;
    }

    private List<TopLevelWindow> FindBlockingDialogs()
    {
        var dialogs = new List<TopLevelWindow>();

        NativeMethods.EnumWindowsProc callback = (handle, _) =>
        {
            // Hidden windows are the vast majority and can never be dialogs; skip them before
            // reading anything else.
            if (NativeMethods.IsWindowVisible(handle))
            {
                var window = ReadWindow(handle);
                if (IsBlockingDialog(window, _processId))
                {
                    dialogs.Add(window);
                }
            }

            return true;
        };

        NativeMethods.EnumWindows(callback, IntPtr.Zero);
        GC.KeepAlive(callback);

        return dialogs;
    }

    /// <summary>
    /// Reads a window's state. None of these calls sends a message to the window, so a hung UI
    /// thread cannot stall them.
    /// </summary>
    private static TopLevelWindow ReadWindow(IntPtr handle)
    {
        NativeMethods.GetWindowThreadProcessId(handle, out var processId);

        var className = new StringBuilder(256);
        NativeMethods.GetClassName(handle, className, className.Capacity);

        var extendedStyle = NativeMethods.GetWindowLong(handle, NativeMethods.GWL_EXSTYLE);
        var owner = NativeMethods.GetWindow(handle, NativeMethods.GW_OWNER);

        var window = new TopLevelWindow
        {
            Handle = handle,
            ProcessId = (int)processId,
            ClassName = className.ToString(),
            IsVisible = true,
            IsEnabled = NativeMethods.IsWindowEnabled(handle),
            CanActivate = (extendedStyle & NativeMethods.WS_EX_NOACTIVATE) == 0,
            Owner = owner,
        };

        if (owner != IntPtr.Zero)
        {
            NativeMethods.GetWindowThreadProcessId(owner, out var ownerProcessId);
            window.OwnerProcessId = (int)ownerProcessId;
            window.IsOwnerEnabled = NativeMethods.IsWindowEnabled(owner);
        }

        return window;
    }

    /// <summary>
    /// Describes a dialog, or returns null when it closed before it could be read.
    /// </summary>
    private static async Task<DialogInfo?> DescribeAsync(TopLevelWindow window)
    {
        var dialog = new DialogInfo { Id = FormatId(window.Handle), Title = ReadTitle(window.Handle) };

        try
        {
            var read = await RunWithTimeoutAsync(() => ReadContent(window.Handle, dialog.Title), ReadTimeout)
                .ConfigureAwait(false);

            if (read.Completed)
            {
                dialog.Title = read.Result.Title;
                dialog.Message = read.Result.Message;
                dialog.Buttons = read.Result.Buttons;
                return dialog;
            }

            dialog.Note = "The dialog did not respond in time, so only its title could be read.";
        }
        catch (ElementNotAvailableException)
        {
            // Raised when any element vanishes mid-read, which is not necessarily the dialog.
            if (!IsOpen(window.Handle))
            {
                return null;
            }

            dialog.Note = "The dialog changed while it was being read. Call dialog_list again.";
        }
        catch (Exception ex)
        {
            dialog.Note = $"The dialog's contents could not be read: {ex.Message}";
        }

        return dialog;
    }

    /// <summary>
    /// Reads a window caption from the window manager's own copy. Unlike GetWindowText this
    /// sends no WM_GETTEXT, so it answers even for a dialog whose thread has stopped responding -
    /// where UI Automation can come back with an empty name instead.
    /// </summary>
    private static string ReadTitle(IntPtr handle)
    {
        var buffer = new StringBuilder(512);
        var length = NativeMethods.InternalGetWindowText(handle, buffer, buffer.Capacity);
        return length > 0 ? buffer.ToString() : string.Empty;
    }

    private sealed class DialogContent
    {
        public string Title { get; set; } = string.Empty;
        public string Message { get; set; } = string.Empty;
        public List<DialogButtonInfo> Buttons { get; } = new();

        /// <summary>
        /// The automation element behind each entry in <see cref="Buttons"/>, at the same index.
        /// </summary>
        public List<AutomationElement> Elements { get; } = new();
    }

    /// <param name="title">The caption already read from Win32, preferred over UI Automation's.</param>
    private static DialogContent ReadContent(IntPtr handle, string title)
    {
        var root = AutomationElement.FromHandle(handle);
        var content = new DialogContent
        {
            Title = string.IsNullOrEmpty(title) ? root.Current.Name ?? string.Empty : title,
        };
        var walker = TreeWalker.ControlViewWalker;
        var lines = new List<string>();
        var visited = 0;

        void AddLine(string? text)
        {
            text = text?.Trim();

            // Custom-chrome dialogs draw their caption as ordinary text.
            if (string.IsNullOrEmpty(text) || text == content.Title || (lines.Count > 0 && lines[lines.Count - 1] == text))
            {
                return;
            }

            lines.Add(text!);
        }

        void Collect(AutomationElement? element, int depth)
        {
            for (; element != null && visited < MaxElementsVisited; element = walker.GetNextSibling(element))
            {
                visited++;
                var current = element.Current;
                var controlType = current.ControlType;

                if (controlType == ControlType.Button)
                {
                    var name = current.Name?.Trim();
                    if (!string.IsNullOrEmpty(name))
                    {
                        content.Buttons.Add(new DialogButtonInfo { Name = name!, IsEnabled = current.IsEnabled });
                        content.Elements.Add(element);
                    }

                    continue;
                }

                if (controlType == ControlType.Text)
                {
                    AddLine(current.Name);
                    continue;
                }

                if (controlType == ControlType.Edit || controlType == ControlType.Document)
                {
                    AddLine(ReadOnlyText(element));
                    continue;
                }

                if (!SkippedSubtrees.Contains(controlType) && depth < MaxTreeDepth)
                {
                    Collect(walker.GetFirstChild(element), depth + 1);
                }
            }
        }

        Collect(walker.GetFirstChild(root), 1);

        var message = string.Join("\n", lines);
        content.Message = message.Length > MaxMessageLength
            ? message.Substring(0, MaxMessageLength) + "…"
            : message;

        return content;
    }

    /// <summary>
    /// Some dialogs show their message in a read-only text box so it can be copied. Editable
    /// boxes are input fields, not message text, and are left out.
    /// </summary>
    private static string? ReadOnlyText(AutomationElement element)
    {
        if (element.TryGetCurrentPattern(ValuePattern.Pattern, out var value))
        {
            var current = ((ValuePattern)value).Current;
            return current.IsReadOnly ? current.Value : null;
        }

        if (element.TryGetCurrentPattern(TextPattern.Pattern, out var text))
        {
            return ((TextPattern)text).DocumentRange.GetText(MaxMessageLength);
        }

        return null;
    }

    private static bool Invoke(AutomationElement element)
    {
        if (!element.TryGetCurrentPattern(InvokePattern.Pattern, out var pattern))
        {
            return false;
        }

        ((InvokePattern)pattern).Invoke();
        return true;
    }

    private static bool IsOpen(IntPtr handle) =>
        NativeMethods.IsWindow(handle) && NativeMethods.IsWindowVisible(handle);

    private static async Task<bool> WaitForCloseAsync(IntPtr handle)
    {
        var deadline = DateTime.UtcNow + CloseTimeout;

        while (IsOpen(handle))
        {
            if (DateTime.UtcNow >= deadline)
            {
                return false;
            }

            await Task.Delay(ClosePollInterval).ConfigureAwait(false);
        }

        return true;
    }

    /// <summary>
    /// Runs <paramref name="work"/> on a dedicated background thread and stops waiting after
    /// <paramref name="timeout"/>. A call into a window that never answers then costs one parked
    /// thread rather than the caller, and the thread pool is never starved by it.
    /// </summary>
    private static async Task<(bool Completed, T Result)> RunWithTimeoutAsync<T>(Func<T> work, TimeSpan timeout)
    {
        var completion = new TaskCompletionSource<T>(TaskCreationOptions.RunContinuationsAsynchronously);

        var thread = new Thread(() =>
        {
            try
            {
                completion.SetResult(work());
            }
            catch (Exception ex)
            {
                completion.SetException(ex);
            }
        })
        {
            IsBackground = true,
            Name = "MCPServer dialog automation",
        };

        // UI Automation clients belong in the multithreaded apartment, where calls into the
        // dialog are not funnelled back through a message loop of their own.
        thread.SetApartmentState(ApartmentState.MTA);
        thread.Start();

        var finished = await Task.WhenAny(completion.Task, Task.Delay(timeout)).ConfigureAwait(false);

        return finished == completion.Task
            ? (true, await completion.Task.ConfigureAwait(false))
            : (false, default!);
    }

    private static class NativeMethods
    {
        public const int GWL_EXSTYLE = -20;
        public const int WS_EX_NOACTIVATE = 0x08000000;
        public const uint GW_OWNER = 4;

        public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

        [DllImport("kernel32.dll")]
        public static extern uint GetCurrentProcessId();

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool IsWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool IsWindowVisible(IntPtr hWnd);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool IsWindowEnabled(IntPtr hWnd);

        [DllImport("user32.dll")]
        public static extern IntPtr GetWindow(IntPtr hWnd, uint uCmd);

        [DllImport("user32.dll")]
        public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        public static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

        [DllImport("user32.dll", EntryPoint = "GetWindowLongW")]
        public static extern int GetWindowLong(IntPtr hWnd, int nIndex);

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        public static extern int InternalGetWindowText(IntPtr hWnd, StringBuilder pString, int cchMaxCount);
    }
}
