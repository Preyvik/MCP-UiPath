import { z } from "zod";
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { runPython, readAndCleanup } from "../python-bridge.js";
import { join } from "node:path";
import { tmpdir } from "node:os";
import { randomUUID } from "node:crypto";
import { mkdtemp } from "node:fs/promises";

const TOOL_DESCRIPTION = `Read a UiPath XAML workflow file and convert it to simplified JSON for analysis and modification.

WHEN TO CALL: To inspect an existing workflow's structure, extract its activity tree, or prepare for a read-modify-write cycle.

OUTPUT FORMAT: Returns JSON in writer-compatible format:
- metadata: {class, namespaces, assemblyReferences, arguments}
- variables: top-level variables
- workflow: activity tree with typed Assign objects, "children" key for Sequences, etc.

The output JSON can be modified and passed directly to write_workflow or create_workflow to regenerate the XAML.

READ-MODIFY-WRITE PATTERN:
1. read_workflow → get JSON
2. Modify the JSON (add/remove activities, change values)
3. create_workflow → write modified JSON back to XAML`;

export function registerReadWorkflow(server: McpServer): void {
  server.tool(
    "read_workflow",
    TOOL_DESCRIPTION,
    {
      xamlPath: z.string().describe("Absolute path to the .xaml workflow file to read"),
    },
    async ({ xamlPath }) => {
      let tmpOutput: string | undefined;
      try {
        const dir = await mkdtemp(join(tmpdir(), "mcp-uipath-"));
        tmpOutput = join(dir, `${randomUUID()}.json`);

        const result = await runPython("xaml_syntaxer.py", [
          "--mode", "read",
          "--input", xamlPath,
          "--output", tmpOutput,
        ]);

        if (result.exitCode !== 0) {
          return {
            content: [{ type: "text", text: `Error reading workflow: ${result.stderr || result.stdout}` }],
            isError: true,
          };
        }

        const json = await readAndCleanup(tmpOutput);
        tmpOutput = undefined; // Already cleaned up
        return {
          content: [{ type: "text", text: json }],
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
