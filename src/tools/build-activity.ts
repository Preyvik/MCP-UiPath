import { z } from "zod";
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { runPython, writeTempJson, cleanupTemp } from "../python-bridge.js";

const TOOL_DESCRIPTION = `Validate and build a single UiPath activity from JSON. Checks required attributes and scoping rules, returns validated activity JSON or structured errors.

WHEN TO CALL: To validate a single activity's structure before adding it to a workflow. Useful for checking if an activity has all required fields.

FORMAT NOTE: This tool uses CONSTRUCTOR format (flat strings). Its output is NOT directly compatible with write_workflow which expects writer format (typed objects).

For building complete workflows, use create_workflow instead — it accepts either format and handles the conversion automatically.

CONSTRUCTOR FORMAT (what this tool accepts):
  {"type": "Assign", "displayName": "Set X", "to": "myVar", "value": "\\"Hello\\""}
  {"type": "If", "displayName": "Check", "condition": "x > 5", "then": null, "else": null}

WRITER FORMAT (what write_workflow needs — different!):
  {"type": "Assign", "displayName": "Set X", "to": {"type": "String", "value": "[myVar]"}, "value": {"type": "String", "value": "[\\"Hello\\"]"}}`;

export function registerBuildActivity(server: McpServer): void {
  server.tool(
    "build_activity",
    TOOL_DESCRIPTION,
    {
      activityJson: z.string().describe("JSON string in constructor format (flat strings for to/value, 'type' field required)"),
    },
    async ({ activityJson }) => {
      let tmpPath: string | undefined;
      try {
        tmpPath = await writeTempJson(activityJson);

        const result = await runPython("xaml_constructor.py", [
          "--mode", "build",
          "--input", tmpPath,
        ]);

        if (result.exitCode !== 0) {
          // Validation errors are informative, not system errors
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
      } finally {
        if (tmpPath) await cleanupTemp(tmpPath);
      }
    }
  );
}
