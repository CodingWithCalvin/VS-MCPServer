# 🤖 VS MCP Server

**Let AI assistants like Claude control Visual Studio through the Model Context Protocol!**

[![License](https://img.shields.io/github/license/CodingWithCalvin/VS-VSMCP?style=for-the-badge)](https://github.com/CodingWithCalvin/VS-VSMCP/blob/main/LICENSE)
[![Build Status](https://img.shields.io/github/actions/workflow/status/CodingWithCalvin/VS-VSMCP/build.yml?style=for-the-badge)](https://github.com/CodingWithCalvin/VS-VSMCP/actions/workflows/build.yml)
[![Marketplace Version](https://img.shields.io/visual-studio-marketplace/v/CodingWithCalvin.VS-VSMCP?style=for-the-badge)](https://marketplace.visualstudio.com/items?itemName=CodingWithCalvin.VS-VSMCP)
[![Marketplace Installs](https://img.shields.io/visual-studio-marketplace/i/CodingWithCalvin.VS-VSMCP?style=for-the-badge)](https://marketplace.visualstudio.com/items?itemName=CodingWithCalvin.VS-VSMCP)

---

## 🤔 What is this?

**VS MCP Server** exposes Visual Studio features through the [Model Context Protocol (MCP)](https://modelcontextprotocol.io/), enabling AI assistants like Claude to interact with your IDE programmatically. Open files, read code, build projects, and more - all through natural conversation!

## ✨ Features

### 📂 Solution Tools
- **solution_info** - Get information about the current solution
- **solution_open** - Open a solution file
- **solution_close** - Close the current solution
- **project_list** - List all projects in the solution
- **project_info** - Get detailed project information

### 📝 Document Tools
- **document_list** - List all open documents
- **document_active** - Get the active document
- **document_open** - Open a file in the editor
- **document_close** - Close a document
- **document_read** - Read document contents
- **document_write** - Write to a document

### ✏️ Editor Tools
- **selection_get** - Get the current text selection
- **selection_set** - Set the selection range
- **editor_insert** - Insert text at cursor position
- **editor_replace** - Find and replace text
- **editor_goto_line** - Navigate to a specific line
- **editor_find** - Search within documents

### 🔨 Build Tools
- **build_solution** - Build the entire solution
- **build_project** - Build a specific project
- **clean_solution** - Clean the solution
- **build_cancel** - Cancel a running build
- **build_status** - Get current build status

## 🛠️ Installation

### Visual Studio Marketplace

1. Open Visual Studio 2022 or 2026
2. Go to **Extensions > Manage Extensions**
3. Search for "VS MCP Server"
4. Click **Download** and restart Visual Studio

### Manual Installation

Download the latest `.vsix` from the [Releases](https://github.com/CodingWithCalvin/VS-VSMCP/releases) page and double-click to install.

## 🚀 Usage

### Starting the Server

1. Open Visual Studio
2. Go to **Tools > VSMCP > Start Server** (or enable auto-start in settings)
3. The MCP server starts on `http://localhost:5050`

### Configuring Claude Desktop

Add this to your Claude Desktop MCP settings:

```json
{
  "mcpServers": {
    "visual-studio": {
      "url": "http://localhost:5050"
    }
  }
}
```

### Settings

Configure the extension at **Tools > Options > VSMCP**:

- **Auto-start server** - Start the MCP server when Visual Studio launches
- **HTTP Port** - Port for the MCP server (default: 5050)

## 🏗️ Architecture

```
┌─────────────────┐                ┌─────────────────────┐    named pipes    ┌─────────────────┐
│  Claude Desktop │   HTTP/SSE    │  VSMCP.Server.exe   │ ◄────────────────► │  VS Extension   │
│  (MCP Client)   │ ◄────────────► │  (MCP Server)       │    JSON-RPC        │  (Tool Impl)    │
└─────────────────┘  :5050         └─────────────────────┘                    └─────────────────┘
```

## 🤝 Contributing

Contributions are welcome! Whether it's bug reports, feature requests, or pull requests - all feedback helps make this extension better.

### Development Setup

1. Clone the repository
2. Open `src/CodingWithCalvin.VSMCP.slnx` in Visual Studio 2022
3. Ensure you have the "Visual Studio extension development" workload installed
4. Ensure you have .NET 10.0 SDK installed
5. Press F5 to launch the experimental instance

## 📄 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

---

## 👥 Contributors

<!-- readme: contributors -start -->
<!-- readme: contributors -end -->

---

Made with ❤️ by [Coding With Calvin](https://github.com/CalvinAllen)
