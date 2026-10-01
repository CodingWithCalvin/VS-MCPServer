using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using CodingWithCalvin.MCPServer.Services;
using CodingWithCalvin.MCPServer.Shared.Models;
using Xunit;
using Forms = System.Windows.Forms;
using Wpf = System.Windows;
using WpfControls = System.Windows.Controls;

namespace CodingWithCalvin.MCPServer.Tests;

/// <summary>
/// Covers finding, reading and answering the modal dialogs that block Visual Studio (issue #112).
/// </summary>
/// <remarks>
/// The rule and matching tests run against captured window facts. The rest raise genuine Win32
/// and WPF modal dialogs on a dedicated UI thread in the test process and drive them through UI
/// Automation from another thread, which is how the tools reach a dialog that is holding the
/// Visual Studio UI thread.
/// </remarks>
public class DialogAutomationTests
{
    private const int VisualStudioProcessId = 100;
    private const int OtherProcessId = 200;

    /// <summary>
    /// Generous because the first WPF window in a fresh process takes several seconds to start
    /// up, and longer on a cold build agent. Only a failing test ever waits this long.
    /// </summary>
    private static readonly TimeSpan AppearTimeout = TimeSpan.FromSeconds(60);
    private static readonly TimeSpan ResultTimeout = TimeSpan.FromSeconds(10);

    [Fact]
    public void IsBlockingDialog_True_WhenOwnerIsDisabled()
    {
        Assert.True(DialogAutomation.IsBlockingDialog(OwnedWindow(), VisualStudioProcessId));
    }

    [Fact]
    public void IsBlockingDialog_False_ForFloatingToolWindow()
    {
        // Owned just like a dialog, but nothing has disabled its owner.
        var window = OwnedWindow();
        window.IsOwnerEnabled = true;

        Assert.False(DialogAutomation.IsBlockingDialog(window, VisualStudioProcessId));
    }

    [Fact]
    public void IsBlockingDialog_False_WhenANestedDialogHasDisabledIt()
    {
        // Only the innermost dialog of a stack can be answered.
        var window = OwnedWindow();
        window.IsEnabled = false;

        Assert.False(DialogAutomation.IsBlockingDialog(window, VisualStudioProcessId));
    }

    [Fact]
    public void IsBlockingDialog_False_WhenHidden()
    {
        var window = OwnedWindow();
        window.IsVisible = false;

        Assert.False(DialogAutomation.IsBlockingDialog(window, VisualStudioProcessId));
    }

    [Fact]
    public void IsBlockingDialog_False_ForWindowThatCannotActivate()
    {
        // A tooltip or popup left over from before the dialog opened.
        var window = OwnedWindow();
        window.CanActivate = false;

        Assert.False(DialogAutomation.IsBlockingDialog(window, VisualStudioProcessId));
    }

    [Fact]
    public void IsBlockingDialog_True_ForUnownedNativeDialog()
    {
        var window = OwnedWindow();
        window.ClassName = "#32770";
        window.Owner = IntPtr.Zero;

        Assert.True(DialogAutomation.IsBlockingDialog(window, VisualStudioProcessId));
    }

    [Fact]
    public void IsBlockingDialog_False_ForUnownedWindow()
    {
        // The Visual Studio main window.
        var window = OwnedWindow();
        window.Owner = IntPtr.Zero;

        Assert.False(DialogAutomation.IsBlockingDialog(window, VisualStudioProcessId));
    }

    [Fact]
    public void IsBlockingDialog_False_ForUnownedNativeDialogInAnotherProcess()
    {
        var window = OwnedWindow();
        window.ClassName = "#32770";
        window.Owner = IntPtr.Zero;
        window.ProcessId = OtherProcessId;

        Assert.False(DialogAutomation.IsBlockingDialog(window, VisualStudioProcessId));
    }

    [Fact]
    public void IsBlockingDialog_True_ForOutOfProcessDialogOwnedByVisualStudio()
    {
        var window = OwnedWindow();
        window.ProcessId = OtherProcessId;

        Assert.True(DialogAutomation.IsBlockingDialog(window, VisualStudioProcessId));
    }

    [Fact]
    public void IsBlockingDialog_False_ForDialogBelongingToAnotherApplication()
    {
        var window = OwnedWindow();
        window.ProcessId = OtherProcessId;
        window.OwnerProcessId = OtherProcessId;

        Assert.False(DialogAutomation.IsBlockingDialog(window, VisualStudioProcessId));
    }

    [Fact]
    public void TrySelectDialog_Fails_WhenNoDialogIsOpen()
    {
        var selected = DialogAutomation.TrySelectDialog(
            new List<DialogAutomation.TopLevelWindow>(),
            dialogId: null,
            out _,
            out var error);

        Assert.False(selected);
        Assert.Equal(DialogAutomation.NoDialogMessage, error);
    }

    [Fact]
    public void TrySelectDialog_PicksTheOnlyDialog_WhenNoIdIsGiven()
    {
        var dialogs = new[] { Dialog(0x1A2B) };

        Assert.True(DialogAutomation.TrySelectDialog(dialogs, dialogId: null, out var selected, out _));
        Assert.Same(dialogs[0], selected);
    }

    [Fact]
    public void TrySelectDialog_RequiresAnId_WhenSeveralDialogsAreOpen()
    {
        var dialogs = new[] { Dialog(0x1A2B), Dialog(0x3C4D) };

        Assert.False(DialogAutomation.TrySelectDialog(dialogs, dialogId: " ", out _, out var error));
        Assert.Contains("2 dialogs are open", error);
    }

    [Theory]
    [InlineData("0x00003C4D")]
    [InlineData("0x3c4d")]
    [InlineData("3C4D")]
    public void TrySelectDialog_FindsDialogById(string dialogId)
    {
        var dialogs = new[] { Dialog(0x1A2B), Dialog(0x3C4D) };

        Assert.True(DialogAutomation.TrySelectDialog(dialogs, dialogId, out var selected, out _));
        Assert.Equal(new IntPtr(0x3C4D), selected!.Handle);
    }

    [Theory]
    [InlineData("0x9999")]
    [InlineData("not-an-id")]
    [InlineData("0x0")]
    public void TrySelectDialog_Fails_ForIdThatMatchesNoOpenDialog(string dialogId)
    {
        var dialogs = new[] { Dialog(0x1A2B) };

        Assert.False(DialogAutomation.TrySelectDialog(dialogs, dialogId, out _, out var error));
        Assert.Contains(dialogId, error);
    }

    [Fact]
    public void FormatId_RoundTripsThroughTryParseId()
    {
        var handle = new IntPtr(0x3C4D);
        var id = DialogAutomation.FormatId(handle);

        Assert.Equal("0x00003C4D", id);
        Assert.True(DialogAutomation.TryParseId(id, out var parsed));
        Assert.Equal(handle, parsed);
    }

    [Theory]
    [InlineData("yes", 0)]
    [InlineData("NO", 1)]
    [InlineData(" Cancel ", 2)]
    [InlineData("&Cancel", 2)]
    public void FindButton_MatchesIgnoringCaseWhitespaceAndAccessKey(string requested, int expected)
    {
        var buttons = Buttons("Yes", "No", "Cancel");

        Assert.Equal(expected, DialogAutomation.FindButton(buttons, requested, out _));
    }

    [Theory]
    [InlineData("Browse...")]
    [InlineData("Browse…")]
    [InlineData("&Browse...")]
    public void FindButton_IgnoresTrailingEllipsis(string caption)
    {
        Assert.Equal(0, DialogAutomation.FindButton(Buttons(caption), "browse", out _));
    }

    [Fact]
    public void FindButton_ListsTheButtons_WhenNoneMatches()
    {
        var index = DialogAutomation.FindButton(Buttons("Yes", "No"), "Maybe", out var error);

        Assert.Equal(-1, index);
        Assert.Contains("Yes, No", error);
    }

    [Fact]
    public void FindButton_RefusesDisabledButton()
    {
        var buttons = Buttons("Apply");
        buttons[0].IsEnabled = false;

        Assert.Equal(-1, DialogAutomation.FindButton(buttons, "Apply", out var error));
        Assert.Contains("disabled", error);
    }

    [Fact]
    public void FindButton_RefusesToGuess_WhenTwoButtonsShareAName()
    {
        Assert.Equal(-1, DialogAutomation.FindButton(Buttons("OK", "&OK"), "OK", out var error));
        Assert.Contains("ambiguous", error);
    }

    [Fact]
    public void FindButton_ExplainsADialogWithNoButtons()
    {
        Assert.Equal(-1, DialogAutomation.FindButton(Buttons(), "OK", out var error));
        Assert.Contains("no buttons", error);
    }

    [Fact]
    public void ToMatchKey_KeepsLiteralAmpersand()
    {
        Assert.Equal("Save & Close", DialogAutomation.ToMatchKey("&Save && Close"));
    }

    [Fact]
    public async Task NoDialogOpen_ListsNothingAndRespondExplainsWhy()
    {
        var automation = new DialogAutomation();

        Assert.Empty((await automation.GetDialogsAsync()).Dialogs);

        var response = await automation.RespondAsync("OK", dialogId: null);

        Assert.False(response.Clicked);
        Assert.Equal(DialogAutomation.NoDialogMessage, response.Message);
    }

    [Fact]
    public async Task MessageBox_IsListedAndAnsweredByButtonName()
    {
        const string Title = "Breakpoint dialog test";
        var automation = new DialogAutomation();

        using var host = new DialogThread(() =>
        {
            using var owner = new Forms.Form();
            _ = owner.Handle;

            return Forms.MessageBox.Show(
                owner,
                "The source code is different from the original version.",
                Title,
                Forms.MessageBoxButtons.YesNoCancel,
                Forms.MessageBoxIcon.Warning);
        });

        var dialog = await WaitForDialogAsync(automation, host, Title);

        Assert.Null(dialog.Note);
        Assert.Equal("The source code is different from the original version.", dialog.Message);
        Assert.Equal(new[] { "Yes", "No", "Cancel" }, dialog.Buttons.Select(b => b.Name).ToArray());
        Assert.All(dialog.Buttons, button => Assert.True(button.IsEnabled));

        var unknown = await automation.RespondAsync("Maybe", dialog.Id);

        Assert.False(unknown.Clicked);
        Assert.Contains("Yes, No, Cancel", unknown.Message);
        Assert.Contains(unknown.OpenDialogs, open => open.Id == dialog.Id);

        var response = await automation.RespondAsync("no", dialog.Id);

        Assert.True(response.Clicked, response.Message);
        Assert.Equal("No", response.Button);
        Assert.True(response.DialogClosed, response.Message);
        Assert.DoesNotContain(response.OpenDialogs, open => open.Id == dialog.Id);
        Assert.Equal(Forms.DialogResult.No, host.WaitForResult());
    }

    [Fact]
    public async Task WpfDialog_IsListedAndAnsweredByButtonName()
    {
        const string Title = "Hot Reload dialog test";
        var automation = new DialogAutomation();

        using var host = new DialogThread(() =>
        {
            var owner = new Wpf.Window
            {
                Left = -10000,
                Top = -10000,
                Width = 100,
                Height = 100,
                ShowInTaskbar = false,
                ShowActivated = false,
            };
            owner.Show();

            var dialog = new Wpf.Window
            {
                Owner = owner,
                Title = Title,
                SizeToContent = Wpf.SizeToContent.WidthAndHeight,
                ShowInTaskbar = false,
                WindowStartupLocation = Wpf.WindowStartupLocation.CenterOwner,
            };

            var edit = new WpfControls.Button { Content = "_Edit" };
            edit.Click += (_, _) => dialog.DialogResult = false;

            var rebuild = new WpfControls.Button { Content = "_Rebuild and Apply Changes" };
            rebuild.Click += (_, _) => dialog.DialogResult = true;

            var panel = new WpfControls.StackPanel();
            panel.Children.Add(new WpfControls.TextBlock { Text = "Hot Reload changes not supported." });
            panel.Children.Add(edit);
            panel.Children.Add(rebuild);
            dialog.Content = panel;

            try
            {
                return dialog.ShowDialog();
            }
            finally
            {
                owner.Close();
            }
        });

        var listed = await WaitForDialogAsync(automation, host, Title);

        // A WPF button also exposes its caption as a Text child; it must not leak into the message.
        Assert.Equal("Hot Reload changes not supported.", listed.Message);
        Assert.Equal(
            new[] { "Edit", "Rebuild and Apply Changes" },
            listed.Buttons.Select(b => b.Name).ToArray());

        var response = await automation.RespondAsync("rebuild and apply changes", dialogId: null);

        Assert.True(response.Clicked, response.Message);
        Assert.True(response.DialogClosed, response.Message);
        Assert.Equal(true, host.WaitForResult());
    }

    /// <summary>
    /// A dialog whose thread is blocked without pumping messages answers nothing sent to it. It
    /// must still be listed, by title, and listing must not hang along with it.
    /// </summary>
    [Fact]
    public async Task UnresponsiveDialog_IsStillListed_WithoutHanging()
    {
        const string Title = "Unresponsive dialog test";
        var automation = new DialogAutomation();
        using var release = new ManualResetEvent(false);

        using var host = new DialogThread(() =>
        {
            using var owner = new Forms.Form();
            using var dialog = new Forms.Form
            {
                Text = Title,
                ShowInTaskbar = false,
                StartPosition = Forms.FormStartPosition.Manual,
                Location = new System.Drawing.Point(-10000, -10000),
            };

            _ = owner.Handle;
            dialog.Show(owner);
            owner.Enabled = false;

            NativeMethods.WaitForSingleObject(release.SafeWaitHandle.DangerousGetHandle(), NativeMethods.Infinite);
            return null;
        });

        try
        {
            var stopwatch = Stopwatch.StartNew();

            var listed = await WaitForDialogAsync(automation, host, Title);

            Assert.Equal(Title, listed.Title);
            Assert.True(
                stopwatch.Elapsed < AppearTimeout + DialogAutomation.ReadTimeout,
                $"Listing an unresponsive dialog took {stopwatch.Elapsed}.");
        }
        finally
        {
            release.Set();
        }
    }

    private static async Task<DialogInfo> WaitForDialogAsync(DialogAutomation automation, DialogThread host, string title)
    {
        var deadline = DateTime.UtcNow + AppearTimeout;

        while (true)
        {
            host.ThrowIfEnded();

            var dialogs = (await automation.GetDialogsAsync()).Dialogs;
            var dialog = dialogs.FirstOrDefault(d => d.Title == title);
            if (dialog != null)
            {
                return dialog;
            }

            Assert.True(
                DateTime.UtcNow < deadline,
                $"No dialog titled \"{title}\" appeared. Listed: [{string.Join(", ", dialogs.Select(d => $"\"{d.Title}\" ({d.Note})"))}]");
            await Task.Delay(100);
        }
    }

    private static DialogAutomation.TopLevelWindow OwnedWindow() => new()
    {
        Handle = new IntPtr(0x1A2B),
        ProcessId = VisualStudioProcessId,
        ClassName = "HwndWrapper[DefaultDomain;;]",
        IsVisible = true,
        IsEnabled = true,
        Owner = new IntPtr(0x5E6F),
        OwnerProcessId = VisualStudioProcessId,
        IsOwnerEnabled = false,
    };

    private static DialogAutomation.TopLevelWindow Dialog(int handle)
    {
        var window = OwnedWindow();
        window.Handle = new IntPtr(handle);
        return window;
    }

    private static List<DialogButtonInfo> Buttons(params string[] names) =>
        names.Select(name => new DialogButtonInfo { Name = name, IsEnabled = true }).ToList();

    /// <summary>
    /// Runs a dialog on its own STA thread, standing in for the Visual Studio UI thread, and
    /// makes sure it is gone afterwards even when an assertion fails before it is answered.
    /// </summary>
    private sealed class DialogThread : IDisposable
    {
        private readonly Thread _thread;
        private readonly TaskCompletionSource<object?> _result =
            new(TaskCreationOptions.RunContinuationsAsynchronously);
        private uint _nativeThreadId;

        public DialogThread(Func<object?> show)
        {
            _thread = new Thread(() =>
            {
                _nativeThreadId = NativeMethods.GetCurrentThreadId();

                try
                {
                    _result.SetResult(show());
                }
                catch (Exception ex)
                {
                    _result.SetException(ex);
                }
            })
            {
                IsBackground = true,
            };

            _thread.SetApartmentState(ApartmentState.STA);
            _thread.Start();
        }

        /// <summary>
        /// Fails the test straight away when the dialog thread finished without ever showing a
        /// dialog, rather than waiting out the timeout with no explanation.
        /// </summary>
        public void ThrowIfEnded()
        {
            if (_result.Task.IsCompleted)
            {
                throw new InvalidOperationException(
                    "The dialog thread ended before its dialog was answered.",
                    _result.Task.Exception?.GetBaseException());
            }
        }

        /// <summary>
        /// Blocks for the dialog's result. Safe because the dialog runs on its own thread and
        /// never needs the caller's.
        /// </summary>
        public object? WaitForResult()
        {
            Assert.True(_result.Task.Wait(ResultTimeout), "The dialog never returned a result.");
            return _result.Task.Result;
        }

        public void Dispose()
        {
            if (!_result.Task.IsCompleted && _nativeThreadId != 0)
            {
                // WM_CLOSE answers a message box with Cancel and closes a WPF window. Disabled
                // windows are owners waiting on a dialog and close with it.
                NativeMethods.EnumThreadWindows(
                    _nativeThreadId,
                    (handle, _) =>
                    {
                        if (NativeMethods.IsWindowEnabled(handle))
                        {
                            NativeMethods.PostMessage(handle, NativeMethods.WM_CLOSE, IntPtr.Zero, IntPtr.Zero);
                        }

                        return true;
                    },
                    IntPtr.Zero);
            }

            _thread.Join(ResultTimeout);
        }
    }

    private static class NativeMethods
    {
        public const uint WM_CLOSE = 0x0010;
        public const uint Infinite = 0xFFFFFFFF;

        public delegate bool EnumThreadWindowsProc(IntPtr hWnd, IntPtr lParam);

        [DllImport("kernel32.dll")]
        public static extern uint GetCurrentThreadId();

        [DllImport("kernel32.dll")]
        public static extern uint WaitForSingleObject(IntPtr hHandle, uint dwMilliseconds);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool EnumThreadWindows(uint dwThreadId, EnumThreadWindowsProc lpfn, IntPtr lParam);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool IsWindowEnabled(IntPtr hWnd);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool PostMessage(IntPtr hWnd, uint msg, IntPtr wParam, IntPtr lParam);
    }
}
