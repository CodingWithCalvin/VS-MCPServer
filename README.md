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

<p align="center">
  <a href="https://marketplace.visualstudio.com/items?itemName=CodingWithCalvin.VS-MCPServer">
    <img src="https://img.shields.io/visual-studio-marketplace/v/CodingWithCalvin.VS-MCPServer?style=for-the-badge" alt="Marketplace Version">
  </a>
  <a href="https://marketplace.visualstudio.com/items?itemName=CodingWithCalvin.VS-MCPServer">
    <img src="https://img.shields.io/visual-studio-marketplace/i/CodingWithCalvin.VS-MCPServer?style=for-the-badge" alt="Marketplace Installations">
  </a>
  <a href="https://marketplace.visualstudio.com/items?itemName=CodingWithCalvin.VS-MCPServer">
    <img src="https://img.shields.io/visual-studio-marketplace/d/CodingWithCalvin.VS-MCPServer?style=for-the-badge" alt="Marketplace Downloads">
  </a>
  <a href="https://marketplace.visualstudio.com/items?itemName=CodingWithCalvin.VS-MCPServer">
    <img src="https://img.shields.io/visual-studio-marketplace/r/CodingWithCalvin.VS-MCPServer?style=for-the-badge" alt="Marketplace Rating">
  </a>
</p>

---

## What is this?

**VS MCP Server** exposes Visual Studio features through the [Model Context Protocol (MCP)](https://modelcontextprotocol.io/), enabling AI assistants like Claude to interact with your IDE programmatically. Open files, read code, build projects, and more - all through natural conversation!

## Features

### Solution Tools

| Tool | Description |
|------|-------------|
| `solution_info` | Get information about the current solution |
| `solution_open` | Open a solution file |
| `solution_close` | Close the current solution |
| `project_list` | List all projects in the solution |
| `project_info` | Get detailed project information |

### Document Tools

| Tool | Description |
|------|-------------|
| `document_list` | List all open documents |
| `document_active` | Get the active document |
| `document_open` | Open a file in the editor |
| `document_close` | Close a document |
| `document_read` | Read document contents |
| `document_write` | Write to a document |

### Editor Tools

| Tool | Description |
|------|-------------|
| `selection_get` | Get the current text selection |
| `selection_set` | Set the selection range |
| `editor_insert` | Insert text at cursor position |
| `editor_replace` | Find and replace text |
| `editor_goto_line` | Navigate to a specific line |
| `editor_find` | Search within documents |

### Build Tools

| Tool | Description |
|------|-------------|
| `build_solution` | Build the entire solution |
| `build_project` | Build a specific project |
| `clean_solution` | Clean the solution |
| `build_cancel` | Cancel a running build |
| `build_status` | Get current build status |

## Installation

### Visual Studio Marketplace

1. Open Visual Studio 2022 or 2026
2. Go to **Extensions > Manage Extensions**
3. Search for "MCP Server"
4. Click **Download** and restart Visual Studio

### Manual Installation

Download the latest `.vsix` from the [Releases](https://github.com/CodingWithCalvin/VS-MCPServer/releases) page and double-click to install.

## Usage

### Starting the Server

1. Open Visual Studio
2. Go to **Tools > MCP Server > Start Server** (or enable auto-start in settings)
3. The MCP server starts on `http://localhost:5050`

### Configuring Claude Desktop

Add this to your Claude Desktop MCP settings:

```json
{
  "mcpServers": {
    "visual-studio": {
      "url": "http://localhost:5050/sse"
    }
  }
}
```

### Settings

Configure the extension at **Tools > Options > MCP Server**:

| Setting | Description | Default |
|---------|-------------|---------|
| Auto-start server | Start the MCP server when Visual Studio launches | Off |
| Binding Address | Address the server binds to | `localhost` |
| HTTP Port | Port for the MCP server | `5050` |
| Server Name | Name reported to MCP clients | `Visual Studio MCP` |
| Log Level | Minimum log level for output | `Information` |
| Log Retention | Days to keep log files | `7` |

## Architecture

```
+------------------+              +----------------------+   named pipes   +------------------+
|  Claude Desktop  |   HTTP/SSE  |  MCPServer.Server    | <-------------> |  VS Extension    |
|  (MCP Client)    | <---------> |  (MCP Server)        |    JSON-RPC     |  (Tool Impl)     |
+------------------+    :5050    +----------------------+                 +------------------+
```

## Contributing

Contributions are welcome! Whether it's bug reports, feature requests, or pull requests - all feedback helps make this extension better.

### Development Setup

1. Clone the repository
2. Open `src/CodingWithCalvin.MCPServer.slnx` in Visual Studio 2022
3. Ensure you have the "Visual Studio extension development" workload installed
4. Ensure you have .NET 10.0 SDK installed
5. Press F5 to launch the experimental instance

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

---

## Contributors

<!-- readme: contributors -start -->
<!-- readme: contributors -end -->

---

<p align="center">
  Made with ❤️ by <a href="https://github.com/CodingWithCalvin">Coding With Calvin</a>
</p>
