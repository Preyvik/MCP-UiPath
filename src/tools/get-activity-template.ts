import { z } from "zod";
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { runPython } from "../python-bridge.js";

const TOOL_DESCRIPTION = `Get the JSON template for a specific UiPath activity type, showing required/optional attributes, defaults, and usage notes.

WHEN TO CALL: When you need to see the exact fields an activity supports before building it.

FORMAT WARNING: Templates are returned in CONSTRUCTOR format (flat strings for Assign.to/value, "activities" key for sequences). The writer and create_workflow tools expect WRITER format:
- Assign.to/value → typed objects: {"type": "String", "value": "[myVar]"}
- Sequence → use "children" key (not "activities")
- Expressions → wrapped in [brackets]

RECOMMENDED: Use the template as a structural reference, then pass your JSON to create_workflow which auto-normalizes constructor format to writer format.

Example activity types: Assign, If, Sequence, TryCatch, ForEach, While, Switch, LogMessage, ExcelProcessScopeX, ReadRangeX, WriteRangeX, UseApplicationCard`;

export function registerGetActivityTemplate(server: McpServer): void {
  server.tool(
    "get_activity_template",
    TOOL_DESCRIPTION,
    {
      activityType: z.string().describe("The activity type name (e.g. 'Assign', 'If', 'ExcelProcessScopeX')"),
    },
    async ({ activityType }) => {
      try {
        const result = await runPython("xaml_constructor.py", [
          "--mode", "template",
          "--type", activityType,
        ]);

        if (result.exitCode !== 0) {
          return {
            content: [{ type: "text", text: result.stderr || result.stdout }],
            isError: false,
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
