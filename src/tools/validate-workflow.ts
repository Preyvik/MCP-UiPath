import { z } from "zod";
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { runPowerShell } from "../python-bridge.js";

const TOOL_DESCRIPTION = `Run the UiPath XAML linter on a workflow file. Returns a structured validation report in JSON format.

WHEN TO CALL: After writing a XAML file with write_workflow, or to check an existing workflow for issues.

NOTE: create_workflow already runs validation automatically — you don't need to call this separately if you used create_workflow.

CHECKS PERFORMED:
- Missing or duplicate XML namespaces
- Hardcoded secrets or credentials
- Scoping violations (e.g., Excel activities outside ExcelProcessScopeX)
- Expression syntax errors
- Missing required attributes
- Variable naming conventions

COMMON FINDINGS & FIXES:
- "Missing namespace": Add the namespace to metadata.namespaces and rebuild
- "Hardcoded secret": Replace literal credentials with arguments or assets
- "Scoping violation": Move activity inside its required parent container
- "Expression error": Check VB.NET syntax and bracket formatting`;

export function registerValidateWorkflow(server: McpServer): void {
  server.tool(
    "validate_workflow",
    TOOL_DESCRIPTION,
    {
      xamlPath: z.string().describe("Absolute path to the .xaml workflow file to validate"),
    },
    async ({ xamlPath }) => {
      try {
        const result = await runPowerShell("UiPath-XAML-Lint.ps1", [
          "-Path", xamlPath,
          "-Strict",
          "-Json",
        ]);

        // Linter may return non-zero for validation failures — that's informative, not a system error
        const output = result.stdout || result.stderr;
        return {
          content: [{ type: "text", text: output }],
          isError: false,
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
