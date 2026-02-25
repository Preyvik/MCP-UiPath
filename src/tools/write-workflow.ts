import { z } from "zod";
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { runPython, writeTempJson, cleanupTemp } from "../python-bridge.js";
import { fixDuplicateXmlns } from "../xaml-utils.js";

const TOOL_DESCRIPTION = `Low-level tool: convert a workflow JSON specification to UiPath XAML and write to disk.

RECOMMENDED: Use "create_workflow" instead — it normalizes format, writes, fixes xmlns, and validates in one call.

Use this tool only when you need direct control over the write step (e.g., writing pre-normalized JSON from read_workflow).

INPUT FORMAT (writer format — strict):
- Assign.to / Assign.value must be typed objects: {"type": "String", "value": "[myVar]"}
- Sequence children must use "children" key (not "activities")
- TryCatch catches: [{"exceptionType": "Exception", "variableName": "ex", "handler": {activity}}]
- All expressions must be in [brackets]: "[myVar]", "[\\"Hello\\"]", "[count > 5]"
- metadata.namespaces and metadata.assemblyReferences must be populated

If your JSON uses flat strings for Assign or "activities" instead of "children", use create_workflow which auto-normalizes.`;

export function registerWriteWorkflow(server: McpServer): void {
  server.tool(
    "write_workflow",
    TOOL_DESCRIPTION,
    {
      workflowJson: z.string().describe("JSON string in writer format (typed Assign objects, 'children' key, bracketed expressions)"),
      outputPath: z.string().describe("Absolute path where the .xaml file should be written"),
    },
    async ({ workflowJson, outputPath }) => {
      let tmpPath: string | undefined;
      try {
        tmpPath = await writeTempJson(workflowJson);

        const result = await runPython("xaml_syntaxer.py", [
          "--mode", "write",
          "--input", tmpPath,
          "--output", outputPath,
        ]);

        if (result.exitCode !== 0) {
          return {
            content: [{ type: "text", text: `Error writing workflow: ${result.stderr || result.stdout}` }],
            isError: true,
          };
        }

        // Fix duplicate xmlns declarations in generated XAML
        await fixDuplicateXmlns(outputPath);

        return {
          content: [{ type: "text", text: result.stdout || `Workflow written successfully to ${outputPath}` }],
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
