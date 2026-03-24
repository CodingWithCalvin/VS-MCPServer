using System.ComponentModel;
using System.Text.Json;
using System.Threading.Tasks;
using ModelContextProtocol.Server;

namespace CodingWithCalvin.MCPServer.Server.Tools;

[McpServerToolType]
public class DebuggerTools
{
    private readonly RpcClient _rpcClient;
    private readonly JsonSerializerOptions _jsonOptions;

    public DebuggerTools(RpcClient rpcClient)
    {
        _rpcClient = rpcClient;
        _jsonOptions = new JsonSerializerOptions { WriteIndented = true };
    }

    [McpServerTool(Name = "debugger_status", ReadOnly = true)]
    [Description("Get the current debugger state. Returns the mode (Design = not debugging, Run = executing, Break = paused at breakpoint/step), break reason, current source location (file, line, function), and debugged process name. Always succeeds regardless of debugger state.")]
    public async Task<string> GetDebuggerStatusAsync()
    {
        var status = await _rpcClient.GetDebuggerStatusAsync();
        return JsonSerializer.Serialize(status, _jsonOptions);
    }

    [McpServerTool(Name = "debugger_launch", Destructive = false)]
    [Description("Start debugging a project (equivalent to F5). If projectName is specified, launches that specific project without changing the startup project. Otherwise debugs the current startup project. A solution must be open. Use debugger_status to check the resulting state.")]
    public async Task<string> DebugLaunchAsync(
        [Description("Optional: The display name of the project to debug (e.g., 'MyProject'). Launches this project directly without changing the startup project. Use project_list to see available project names.")] string? projectName = null)
    {
        if (projectName != null)
        {
            var success = await _rpcClient.DebugLaunchProjectAsync(projectName, noDebug: false);
            return success
                ? $"Debugging started for project: {projectName}"
                : $"Failed to start debugging for project '{projectName}'. Use project_list to verify the project name.";
        }
        else
        {
            var success = await _rpcClient.DebugLaunchAsync();
            return success ? "Debugging started" : "Failed to start debugging (is a solution open with a startup project configured?)";
        }
    }

    [McpServerTool(Name = "debugger_launch_without_debugging", Destructive = false)]
    [Description("Start a project without the debugger attached (equivalent to Ctrl+F5). If projectName is specified, launches that specific project without changing the startup project. Otherwise runs the current startup project. The application runs normally without breakpoints or stepping. A solution must be open.")]
    public async Task<string> DebugLaunchWithoutDebuggingAsync(
        [Description("Optional: The display name of the project to run (e.g., 'MyProject'). Launches this project directly without changing the startup project. Use project_list to see available project names.")] string? projectName = null)
    {
        if (projectName != null)
        {
            var success = await _rpcClient.DebugLaunchProjectAsync(projectName, noDebug: true);
            return success
                ? $"Started without debugging for project: {projectName}"
                : $"Failed to start without debugging for project '{projectName}'. Use project_list to verify the project name.";
        }
        else
        {
            var success = await _rpcClient.DebugLaunchWithoutDebuggingAsync();
            return success ? "Started without debugging" : "Failed to start without debugging (is a solution open with a startup project configured?)";
        }
    }

    [McpServerTool(Name = "debugger_continue", Destructive = false)]
    [Description("Continue execution after a break (equivalent to F5 while paused). Only works when the debugger is in Break mode (paused at a breakpoint or after stepping). Use debugger_status to verify the debugger is in Break mode first.")]
    public async Task<string> DebugContinueAsync()
    {
        var success = await _rpcClient.DebugContinueAsync();
        return success ? "Execution continued" : "Cannot continue (debugger is not in Break mode)";
    }

    [McpServerTool(Name = "debugger_break", Destructive = false)]
    [Description("Pause execution of the running program (equivalent to Ctrl+Alt+Break). Only works when the debugger is in Run mode. Use debugger_status to verify the debugger is in Run mode first.")]
    public async Task<string> DebugBreakAsync()
    {
        var success = await _rpcClient.DebugBreakAsync();
        return success ? "Execution paused" : "Cannot break (debugger is not in Run mode)";
    }

    [McpServerTool(Name = "debugger_stop", Destructive = true)]
    [Description("Stop the current debugging session (equivalent to Shift+F5). Terminates the debugged process. Only works when a debugging session is active (Run or Break mode).")]
    public async Task<string> DebugStopAsync()
    {
        var success = await _rpcClient.DebugStopAsync();
        return success ? "Debugging stopped" : "Cannot stop (no active debugging session)";
    }

    [McpServerTool(Name = "debugger_step_over", Destructive = false)]
    [Description("Step over the current statement (equivalent to F10). Executes the current line and stops at the next line in the same function. Only works when the debugger is in Break mode.")]
    public async Task<string> DebugStepOverAsync()
    {
        var success = await _rpcClient.DebugStepOverAsync();
        return success ? "Stepped over" : "Cannot step over (debugger is not in Break mode)";
    }

    [McpServerTool(Name = "debugger_step_into", Destructive = false)]
    [Description("Step into the current statement (equivalent to F11). If the current line contains a function call, steps into that function. Only works when the debugger is in Break mode.")]
    public async Task<string> DebugStepIntoAsync()
    {
        var success = await _rpcClient.DebugStepIntoAsync();
        return success ? "Stepped into" : "Cannot step into (debugger is not in Break mode)";
    }

    [McpServerTool(Name = "debugger_step_out", Destructive = false)]
    [Description("Step out of the current function (equivalent to Shift+F11). Continues execution until the current function returns, then breaks at the caller. Only works when the debugger is in Break mode.")]
    public async Task<string> DebugStepOutAsync()
    {
        var success = await _rpcClient.DebugStepOutAsync();
        return success ? "Stepped out" : "Cannot step out (debugger is not in Break mode)";
    }

    [McpServerTool(Name = "debugger_add_breakpoint", Destructive = false)]
    [Description("Add a breakpoint at a specific file and line number. Works in any debugger mode (Design, Run, or Break). The file path must be absolute.")]
    public async Task<string> DebugAddBreakpointAsync(string path, int line)
    {
        var success = await _rpcClient.DebugAddBreakpointAsync(path, line);
        return success ? $"Breakpoint added at {path}:{line}" : $"Failed to add breakpoint at {path}:{line}";
    }

    [McpServerTool(Name = "debugger_remove_breakpoint", Destructive = true)]
    [Description("Remove a breakpoint at a specific file and line number. Returns whether a breakpoint was found and removed.")]
    public async Task<string> DebugRemoveBreakpointAsync(string path, int line)
    {
        var success = await _rpcClient.DebugRemoveBreakpointAsync(path, line);
        return success ? $"Breakpoint removed from {path}:{line}" : $"No breakpoint found at {path}:{line}";
    }

    [McpServerTool(Name = "debugger_list_breakpoints", ReadOnly = true)]
    [Description("List all breakpoints in the current solution. Returns file, line, column, function name, condition, enabled state, and hit count for each breakpoint.")]
    public async Task<string> DebugListBreakpointsAsync()
    {
        var breakpoints = await _rpcClient.DebugGetBreakpointsAsync();
        return JsonSerializer.Serialize(breakpoints, _jsonOptions);
    }

    [McpServerTool(Name = "debugger_get_locals", ReadOnly = true)]
    [Description("Get local variables in the current stack frame. Only works when the debugger is in Break mode. Returns name, value, type, and validity for each local variable.")]
    public async Task<string> DebugGetLocalsAsync()
    {
        var locals = await _rpcClient.DebugGetLocalsAsync();
        return JsonSerializer.Serialize(locals, _jsonOptions);
    }

    [McpServerTool(Name = "debugger_evaluate", ReadOnly = true)]
    [Description("Evaluate an expression in the current debug context (like the Immediate Window). Only works when the debugger is in Break mode. Returns the expression, its value, type, and whether the value is valid.")]
    public async Task<string> DebugEvaluateExpressionAsync(
        [Description("The expression to evaluate (e.g., 'myVariable', 'list.Count', 'x + y', 'myObject.ToString()')")] string expression)
    {
        var result = await _rpcClient.DebugEvaluateExpressionAsync(expression);
        return JsonSerializer.Serialize(result, _jsonOptions);
    }

    [McpServerTool(Name = "debugger_set_variable", Destructive = true)]
    [Description("Set the value of a local variable in the current stack frame. Only works when the debugger is in Break mode. The variable must exist in the current scope.")]
    public async Task<string> DebugSetVariableValueAsync(
        [Description("The name of the local variable to modify (e.g., 'count', 'name'). Use debugger_get_locals to see available variables.")] string variableName,
        [Description("The new value to assign (e.g., '42', '\"hello\"', 'true'). Must be a valid expression for the variable's type.")] string value)
    {
        var success = await _rpcClient.DebugSetVariableValueAsync(variableName, value);
        return success
            ? $"Variable '{variableName}' set to: {value}"
            : $"Failed to set variable '{variableName}'. Ensure the debugger is in Break mode, the variable exists in the current scope, and the value is valid for its type.";
    }

    [McpServerTool(Name = "debugger_get_callstack", ReadOnly = true)]
    [Description("Get the call stack of the current thread. Only works when the debugger is in Break mode. Returns depth, function name, file name, line number, module, language, and return type for each frame.")]
    public async Task<string> DebugGetCallStackAsync()
    {
        var callStack = await _rpcClient.DebugGetCallStackAsync();
        return JsonSerializer.Serialize(callStack, _jsonOptions);
    }
}
