import { z } from "zod";
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { runPython, runPowerShell, writeTempJson, cleanupTemp, readAndCleanup } from "../python-bridge.js";

const TOOL_DESCRIPTION = `Edit a UiPath XAML workflow file in-place using surgical XML-level mutations. Preserves ViewState, IdRef assignments, attribute ordering, and all untouched elements exactly.

WHEN TO CALL: When you need to make small, targeted changes to an existing XAML workflow — change a log message, update InvokeCode, add/remove a variable, or remove an activity. Much faster and safer than a full read→modify→write round-trip.

EDIT OPERATIONS (JSON array):

set_attribute — Change an XML attribute on an activity:
  {"action": "set_attribute", "displayName": "Log Start", "attribute": "Message", "value": "[\\"New message\\"]"}
  {"action": "set_attribute", "displayName": "Extract Data", "attribute": "Code", "value": "...new VB code..."}
  Supports any XML attribute: Message, Level, Code, Condition, DisplayName, etc.

remove_variable — Remove a variable from a Sequence:
  {"action": "remove_variable", "name": "oldVar", "sequenceDisplayName": "Main Sequence"}
  If sequenceDisplayName is omitted, removes the first matching variable from any sequence.

add_variable — Add a variable to a Sequence:
  {"action": "add_variable", "name": "newVar", "type": "String", "default": "\\"initial\\"", "sequenceDisplayName": "Main Sequence"}
  Type names: String, Int32, Boolean, Double, DateTime, DataTable, Object, etc.

remove_activity — Remove an activity and all its children:
  {"action": "remove_activity", "displayName": "Log Debug Info"}

TARGETING: Activities are found by "displayName" (most common) or "idRef" (for disambiguation when multiple activities share a display name).

ATOMICITY: If any edit in the array fails, the file is NOT modified. All-or-nothing.

OUTPUT: Returns {success, changes[], warnings[], validation?} with details of each applied edit.`;

export function registerEditWorkflow(server: McpServer): void {
  server.tool(
    "edit_workflow",
    TOOL_DESCRIPTION,
    {
      xamlPath: z.string().describe("Absolute path to the .xaml workflow file to edit (modified in-place)"),
      edits: z.string().describe("JSON array of edit operations (see tool description for format)"),
      validate: z.boolean().default(true).describe("Run linter after editing (default: true)"),
    },
    async ({ xamlPath, edits, validate }) => {
      let tmpEdits: string | undefined;
      let tmpResult: string | undefined;
      try {
        // Step 1: Parse and validate edits JSON
        let editsArray: unknown;
        try {
          editsArray = JSON.parse(edits);
        } catch (e) {
          throw new Error(`Invalid edits JSON: ${(e as Error).message}`);
        }

        if (!Array.isArray(editsArray)) {
          throw new Error("edits must be a JSON array of edit operations");
        }

        if (editsArray.length === 0) {
          throw new Error("edits array is empty — nothing to do");
        }

        // Step 2: Write edits to temp file
        tmpEdits = await writeTempJson(JSON.stringify(editsArray, null, 2));
        tmpResult = tmpEdits.replace(".json", "-result.json");

        // Step 3: Run Python editor
        const editResult = await runPython("xaml_syntaxer.py", [
          "--mode", "edit",
          "--input", xamlPath,
          "--edits", tmpEdits,
          "--output", tmpResult,
        ]);

        if (editResult.exitCode !== 0) {
          return {
            content: [{
              type: "text",
              text: JSON.stringify({
                success: false,
                error: "XAML edit failed",
                details: editResult.stderr || editResult.stdout,
              }, null, 2),
            }],
            isError: true,
          };
        }

        // Step 4: Read result
        const resultContent = await readAndCleanup(tmpResult);
        tmpResult = undefined; // already cleaned up
        let result: any;
        try {
          result = JSON.parse(resultContent);
        } catch {
          result = { success: false, error: "Failed to parse editor result", raw: resultContent };
        }

        if (!result.success) {
          return {
            content: [{
              type: "text",
              text: JSON.stringify(result, null, 2),
            }],
            isError: true,
          };
        }

        // Step 5: Validate if requested
        let validation: any = null;
        if (validate) {
          try {
            const lintResult = await runPowerShell("UiPath-XAML-Lint.ps1", [
              "-Path", xamlPath,
              "-Strict",
              "-Json",
            ]);
            validation = lintResult.stdout || lintResult.stderr;
          } catch {
            validation = "Validation skipped (linter unavailable)";
          }
        }

        return {
          content: [{
            type: "text",
            text: JSON.stringify({
              success: true,
              changes: result.changes,
              warnings: result.warnings,
              validation,
            }, null, 2),
          }],
        };
      } catch (err) {
        return {
          content: [{
            type: "text",
            text: JSON.stringify({
              success: false,
              error: (err as Error).message,
            }, null, 2),
          }],
          isError: true,
        };
      } finally {
        if (tmpEdits) await cleanupTemp(tmpEdits);
        if (tmpResult) await cleanupTemp(tmpResult);
      }
    }
  );
}
