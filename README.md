<p align="center">
  <img src="https://raw.githubusercontent.com/CodingWithCalvin/VS-MCPServer/main/resources/logo.png" alt="VS MCP Server Logo" width="128" height="128">
</p>

<h1 align="center">VS MCP Server</h1>

<p align="center">
  <strong>Let AI assistants like Claude control Visual Studio through the Model Context Protocol!</strong>
</p>

<p align="center">
  <a href="https://github.com/CodingWithCalvin/VS-MCPServer/blob/main/LICENSE">
    <img src="https://img.shields.io/github/license/CodingWithCalvin/VS-MCPServer?style=for-the-badge" alt="License">
  </a>
  <a href="https://github.com/CodingWithCalvin/VS-MCPServer/actions/workflows/build.yml">
    <img src="https://img.shields.io/github/actions/workflow/status/CodingWithCalvin/VS-MCPServer/build.yml?style=for-the-badge" alt="Build Status">
  </a>
</p>

---

## 🤔 What is this?

**VS MCP Server** exposes Visual Studio features through the [Model Context Protocol (MCP)](https://modelcontextprotocol.io/), enabling AI assistants like Claude to interact with your IDE programmatically. Open files, read code, build projects, and more - all through natural conversation!

## ✨ Features

### 📂 Solution Tools

| Tool | Description |
|------|-------------|
| `project_info` | Get detailed project information |
| `project_list` | List all projects in the solution |
| `solution_close` | Close the current solution |
| `solution_info` | Get information about the current solution |
| `solution_open` | Open a solution file |
| `startup_project_get` | Get the current startup project |
| `startup_project_set` | Set the startup project for debugging |

### 📝 Document Tools

| Tool | Description |
|------|-------------|
| `document_active` | Get the active document |
| `document_cleanup` | Run code cleanup on a document |
| `document_close` | Close a document |
| `document_list` | List all open documents |
| `document_open` | Open a file in the editor |
| `document_read` | Read document contents |
| `document_save` | Saves an open document |
| `document_write` | Write to a document |

### ✏️ Editor Tools

| Tool | Description |
|------|-------------|
| `editor_find` | Search within documents |
| `editor_goto_line` | Navigate to a specific line |
| `editor_insert` | Insert text at cursor position |
| `editor_replace` | Find and replace text |
| `selection_get` | Get the current text selection |
| `selection_set` | Set the selection range |

### 🔨 Build Tools

| Tool | Description |
|------|-------------|
| `build_cancel` | Cancel a running build |
| `build_configuration_get` | Get the active and available build configuration/platform pairs |
| `build_configuration_set` | Change the active build configuration and platform |
| `build_project` | Build a specific project |
| `build_solution` | Build the entire solution |
| `build_status` | Get current build status |
| `clean_solution` | Clean the solution |

### 🧭 Navigation Tools

| Tool | Description |
|------|-------------|
| `find_references` | Find all references to a symbol |
| `goto_definition` | Navigate to the definition of a symbol |
| `symbol_document` | Get all symbols defined in a document |
| `symbol_workspace` | Search for symbols across the solution |

### 🐛 Debugger Tools

| Tool | Description |
|------|-------------|
| `debugger_add_breakpoint` | Add a breakpoint at a file and line |
| `debugger_break` | Pause execution (Ctrl+Alt+Break) |
| `debugger_continue` | Continue execution (F5) |
| `debugger_evaluate` | Evaluate an expression in the current debug context |
| `debugger_get_callstack` | Get the call stack |
| `debugger_get_locals` | Get local variables in current frame |
| `debugger_launch` | Start debugging (F5), optionally for a specific project |
| `debugger_launch_without_debugging` | Start without debugger (Ctrl+F5), optionally for a specific project |
| `debugger_list_breakpoints` | List all breakpoints |
| `debugger_remove_breakpoint` | Remove a breakpoint |
| `debugger_set_variable` | Set the value of a local variable |
| `debugger_status` | Get current debugger state |
| `debugger_step_into` | Step into (F11) |
| `debugger_step_out` | Step out (Shift+F11) |
| `debugger_step_over` | Step over (F10) |
| `debugger_stop` | Stop debugging (Shift+F5) |

### 🔍 Diagnostics Tools

| Tool | Description |
|------|-------------|
| `errors_list` | Read build errors, warnings, and messages from the Error List |
| `output_list_panes` | List all available Output window panes |
| `output_read` | Read content from an Output window pane |
| `output_write` | Write a message to an Output window pane |

### 🧪 Test Tools

| Tool | Description |
|------|-------------|
| `test_cancel` | Cancel the test run in progress |
| `test_debug` | Debug the tests in a single class or method |
| `test_debug_all` | Debug every test in the solution |
| `test_run` | Run the tests in a single class or method |
| `test_run_all` | Run every test in the solution |
| `test_stats` | Get passed, failed, skipped, and not-run counts |
| `test_status` | Get the run state plus current counts |

### 📊 Coverage Tools

| Tool | Description |
|------|-------------|
| `coverage_analyze` | Run all tests with code coverage collection |
| `coverage_report` | Read results as a module / class / method tree with line and block counts |
| `coverage_show` | Open the Code Coverage Results window |

> ℹ️ **Running** coverage needs an edition that supports it — Enterprise only through
> VS 2022, all editions from VS 2026. **Reading** an existing `.coverage` file with
> `coverage_report` works on every edition.

### 💻 Terminal Tools

| Tool | Description |
|------|-------------|
| `terminal_close` | Close a single integrated terminal |
| `terminal_close_all` | Close every integrated terminal |
| `terminal_create` | Open an empty terminal using the default profile |
| `terminal_list` | List the open terminal identifiers |
| `terminal_run` | Run a command in a new terminal, inside the VS developer environment |
| `terminal_show` | Bring a terminal into view |

> ⚠️ **Terminal output is not captured.** The Visual Studio terminal is a raw PTY with no
> exit code or command boundaries, so `terminal_run` reports only whether the terminal opened.
> To read results, redirect output to a file and read it back with `document_read`.

### 🪟 Window Tools

| Tool | Description |
|------|-------------|
| `toolwindow_hide` | Hide (close) a tool window by caption |
| `toolwindow_show` | Show a tool window by name (SolutionExplorer, ErrorList, Output, Terminal, etc.) |
| `window_activate` | Activate (focus) a window by caption |
| `window_list` | List all open windows with caption, kind, visibility, and GUID |

## 🛠️ Installation

### Visual Studio Marketplace

1. Open Visual Studio 2022 or 2026
2. Go to **Extensions > Manage Extensions**
3. Search for "MCP Server"
4. Click **Download** and restart Visual Studio

### Manual Installation

Download the latest `.vsix` from the [Releases](https://github.com/CodingWithCalvin/VS-MCPServer/releases) page and double-click to install.

## 🚀 Usage

### ▶️ Starting the Server

1. Open Visual Studio
2. Go to **Tools > MCP Server > Start Server** (or enable auto-start in settings)
3. The MCP server starts on `http://localhost:5050`

### 🤖 Configuring Claude Desktop & Claude Code

Add this to your Claude Desktop or Claude Code MCP settings (preferred HTTP method):

```json
{
  "mcpServers": {
    "visualstudio": {
      "type": "http",
      "url": "http://localhost:5050"
    }
  }
}
```

**Legacy SSE method** (deprecated, but still supported):

```json
{
  "mcpServers": {
    "visualstudio": {
      "type": "sse",
      "url": "http://localhost:5050/sse"
    }
  }
}
```

**Claude Code (CLI) - alternate installation technique:**
```
   claude mcp add --transport http visualstudio http://localhost:5050
```

> ℹ️ **Note:** The HTTP method is the preferred standard. SSE (Server-Sent Events) is a legacy protocol and should only be used for backward compatibility.

### ⚙️ Settings

Configure the extension at **Tools > Options > MCP Server**:

| Setting | Description | Default |
|---------|-------------|---------|
| Auto-start server | Start the MCP server when Visual Studio launches | Off |
| Binding Address | Address the server binds to | `localhost` |
| HTTP Port | Port for the MCP server | `5050` |
| Server Name | Name reported to MCP clients | `Visual Studio MCP` |
| Log Level | Minimum log level for output | `Information` |
| Log Retention | Days to keep log files | `7` |

> ⚠️ **`terminal_run` executes commands on your machine.** Combined with a **Binding
> Address** other than `localhost`, this exposes command execution to your network. Leave the
> binding address at `localhost` unless you specifically need remote access.

## 🏗️ Architecture

```
+------------------+              +----------------------+   named pipes   +------------------+
|  Claude Desktop  |   HTTP/SSE  |  MCPServer.Server    | <-------------> |  VS Extension    |
|  (MCP Client)    | <---------> |  (MCP Server)        |    JSON-RPC     |  (Tool Impl)     |
+------------------+    :5050    +----------------------+                 +------------------+
```

## 🤝 Contributing

Contributions are welcome! Whether it's bug reports, feature requests, or pull requests - all feedback helps make this extension better.

### 🔧 Development Setup

1. Clone the repository
2. Open `src/CodingWithCalvin.MCPServer.slnx` in Visual Studio 2022
3. Ensure you have the "Visual Studio extension development" workload installed
4. Ensure you have .NET 10.0 SDK installed
5. Press F5 to launch the experimental instance

## 📄 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

---

## 👥 Contributors

<!-- readme: contributors -start -->
<a href="https://github.com/CalvinAllen"><img src="https://avatars.githubusercontent.com/u/41448698?v=4&s=64" width="64" height="64" align="left" alt="CalvinAllen"></a> <a href="https://github.com/Arenoros"><img src="https://avatars.githubusercontent.com/u/2578918?v=4&s=64" width="64" height="64" align="left" alt="Arenoros"></a> <a href="https://github.com/Gh61"><img src="https://avatars.githubusercontent.com/u/10837736?v=4&s=64" width="64" height="64" align="left" alt="Gh61"></a> <a href="https://github.com/Vidoy"><img src="https://avatars.githubusercontent.com/u/6883779?v=4&s=64" width="64" height="64" align="left" alt="Vidoy"></a> <a href="https://github.com/fivestar1103"><img src="https://avatars.githubusercontent.com/u/70753360?v=4&s=64" width="64" height="64" align="left" alt="fivestar1103"></a> <a href="https://github.com/gclpixel"><img src="https://avatars.githubusercontent.com/u/9727108?v=4&s=64" width="64" height="64" align="left" alt="gclpixel"></a> <a href="https://github.com/hurricanepkt"><img src="https://avatars.githubusercontent.com/u/161399?v=4&s=64" width="64" height="64" align="left" alt="hurricanepkt"></a> <a href="https://github.com/laviRZ"><img src="https://avatars.githubusercontent.com/u/29277997?v=4&s=64" width="64" height="64" align="left" alt="laviRZ"></a> <a href="https://github.com/mbeeson-mm"><img src="https://avatars.githubusercontent.com/u/236729541?v=4&s=64" width="64" height="64" align="left" alt="mbeeson-mm"></a> <a href="https://github.com/shaiku"><img src="https://avatars.githubusercontent.com/u/16620522?v=4&s=64" width="64" height="64" align="left" alt="shaiku"></a> <br clear="all">
<!-- readme: contributors -end -->

---

<p align="center">
  Made with ❤️ by <a href="https://github.com/CodingWithCalvin">Coding With Calvin</a>
</p>
