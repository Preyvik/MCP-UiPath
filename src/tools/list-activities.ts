import { z } from "zod";
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { runPython } from "../python-bridge.js";

const TOOL_DESCRIPTION = `List all available UiPath activities from the template library, organized by category.

WHEN TO CALL: At the start of workflow design, to discover available activity types and their categories.

Returns a JSON catalog with:
- Activity type names (use these as the "type" field in workflow JSON)
- Categories (Control Flow, Data, Excel, UI Automation, etc.)
- Required attributes for each activity
- Brief descriptions

WORKFLOW POSITION: Call this BEFORE get_activity_template or create_workflow to know what activities are available.

NOTE: The templates returned use constructor format (flat strings). When building a full workflow, use create_workflow which auto-normalizes to writer format.`;

export function registerListActivities(server: McpServer): void {
  server.tool(
    "list_activities",
    TOOL_DESCRIPTION,
    {},
    async () => {
      try {
        const result = await runPython("xaml_constructor.py", ["--mode", "list"]);

        if (result.exitCode !== 0) {
          return {
            content: [{ type: "text", text: `Error listing activities: ${result.stderr || result.stdout}` }],
            isError: true,
          };
        }

        return {
          content: [{ type: "text", text: result.stdout }],
        };
      } catch (err) {
        return {
          content: [{ type: "text", text: `System error: ${(err as Error).message}` }],
          isError: true,
        };
      }
    }
  );
}
