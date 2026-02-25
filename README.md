# MCP-UiPath

MCP server that wraps UiPath XAML tooling (syntaxer, constructor, linter) so any LLM can construct UiPath workflows without dealing with raw XAML syntax.

## Tools

| Tool | Description |
|------|-------------|
| `list_activities` | List all available UiPath activities from the template library |
| `get_activity_template` | Get the JSON template for a specific activity type |
| `build_activity` | Validate and build an activity from a JSON specification |
| `read_workflow` | Read a XAML workflow file and convert to JSON |
| `write_workflow` | Convert workflow JSON to XAML and write to disk |
| `validate_workflow` | Run the XAML linter on a workflow file |

## Setup

```bash
npm install
npm run build
```

## Configuration

Add to `.mcp.json` in your project root:

```json
{
  "mcpServers": {
    "mcp-uipath": {
      "command": "node",
      "args": ["C:\\Users\\20300234\\Claude\\MCP-UiPath\\build\\index.js"]
    }
  }
}
```

## Architecture

```
src/
├── index.ts             # Server entry point + tool registration + stdio transport
├── python-bridge.ts     # Subprocess wrapper for Python/PowerShell scripts
└── tools/
    ├── list-activities.ts
    ├── get-activity-template.ts
    ├── build-activity.ts
    ├── read-workflow.ts
    ├── write-workflow.ts
    └── validate-workflow.ts
```

The server delegates to three existing scripts:
- `xaml_constructor.py` — activity catalog, templates, and build validation
- `xaml_syntaxer.py` — bidirectional XAML-JSON conversion
- `UiPath-XAML-Lint.ps1` — comprehensive XAML linting
