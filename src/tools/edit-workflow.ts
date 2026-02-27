import { z } from "zod";
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { runPython, runPowerShell, writeTempJson, cleanupTemp, readAndCleanup } from "../python-bridge.js";

const TOOL_DESCRIPTION = `Edit a UiPath XAML workflow file in-place using surgical XML-level mutations. Preserves ViewState, IdRef assignments, attribute ordering, and all untouched elements exactly.

WHEN TO CALL: When you need to make targeted changes to an existing XAML workflow — change a log message, update InvokeCode, add/remove a variable, remove/insert/move/replace an activity, wrap/unwrap containers, add/remove arguments, or rename activities. Much faster and safer than a full read→modify→write round-trip.

EDIT OPERATIONS (JSON array):

set_attribute — Change an XML attribute on an activity:
  {"action": "set_attribute", "displayName": "Log Start", "attribute": "Message", "value": "[\\"New message\\"]"}
  {"action": "set_attribute", "displayName": "Extract Data", "attribute": "Code", "value": "...new VB code..."}
  Supports any XML attribute: Message, Level, Code, Condition, DisplayName, etc.

set_element_value — Change text content of a child property element (Assign.To, Assign.Value, etc.):
  {"action": "set_element_value", "displayName": "Set Username", "property": "Value", "value": "[\\"newUser\\"]"}
  {"action": "set_element_value", "displayName": "Set Username", "property": "To", "value": "[myVar]"}
  Use for Assign.To, Assign.Value, ForEach.Values, and any property stored as InArgument/OutArgument.
  "property" accepts short form ("Value", "To") or full form ("Assign.Value").
  Optional "type" updates x:TypeArguments (e.g., "String", "Int32", "Boolean").

remove_variable — Remove a variable from a Sequence:
  {"action": "remove_variable", "name": "oldVar", "sequenceDisplayName": "Main Sequence"}
  If sequenceDisplayName is omitted, removes the first matching variable from any sequence.

add_variable — Add a variable to a Sequence:
  {"action": "add_variable", "name": "newVar", "type": "String", "default": "\\"initial\\"", "sequenceDisplayName": "Main Sequence"}
  Type names: String, Int32, Boolean, Double, DateTime, DataTable, Object, etc.

remove_activity — Remove an activity and all its children:
  {"action": "remove_activity", "displayName": "Log Debug Info"}

insert_activity — Insert a new activity into a Sequence:
  {"action": "insert_activity", "parentDisplayName": "Main Sequence",
   "activity": {"type": "LogMessage", "displayName": "Log Done", "level": "Info", "message": "[\\"Done\\"]"},
   "position": "end"}
  position: "start" (after variables), "end" (default, before ViewState), or "after" (requires afterDisplayName/afterIdRef).
  The activity JSON is auto-corrected (expression wrapping, type normalization) before building.

wrap_in_container — Wrap existing activities in a new container (TryCatch, If, While, ForEach, Sequence):
  {"action": "wrap_in_container",
   "targets": [{"displayName": "Step 1"}, {"displayName": "Step 2"}],
   "container": {"type": "TryCatch", "displayName": "Handle Errors", "placement": "try",
     "catches": [{"exceptionType": "Exception", "variableName": "ex",
       "handler": {"type": "LogMessage", "displayName": "Log Error", "level": "Error", "message": "[\\"Error: \\" & ex.Message]"}}]}}
  All targets must share the same parent. "placement" controls where targets go inside the container:
    TryCatch → "try" (default) | If → "then" (default), "else" | While/ForEach → "body" (default) | Sequence → "children" (default)
  Multiple targets in a single-activity slot are auto-wrapped in a Sequence.

move_activity — Move an activity to a different parent or position:
  {"action": "move_activity", "displayName": "Log Start", "targetParentDisplayName": "VIP Path", "position": "end"}
  position: "start", "end" (default), or "after" (requires afterDisplayName/afterIdRef).
  Preserves original IdRef/ViewState — no rebuild. Target parent must be a Sequence.

replace_activity — Replace an activity with a different one in-place:
  {"action": "replace_activity", "displayName": "Old Activity",
   "activity": {"type": "LogMessage", "displayName": "New Activity", "level": "Info", "message": "[\\"Hello\\"]"}}
  The new activity JSON is auto-corrected before building. Old activity is removed and new one inserted at the same position.

add_argument — Add a workflow argument (In/Out/InOut):
  {"action": "add_argument", "name": "in_FilePath", "direction": "In", "type": "String"}
  direction: "In" (default), "Out", "InOut". type defaults to "String".
  Creates x:Property in x:Members (creates x:Members if it doesn't exist).

remove_argument — Remove a workflow argument by name:
  {"action": "remove_argument", "name": "in_FilePath"}
  Removes the x:Property from x:Members. Cleans up empty x:Members element.

unwrap_container — Unwrap a container, flattening its children back into the parent:
  {"action": "unwrap_container", "displayName": "Handle Errors", "slot": "try"}
  slot defaults based on container type: TryCatch→"try", If→"then", While/ForEach→"body", Sequence→"children".
  Children are extracted from the slot and inserted at the container's position in the parent.

rename_activity — Rename an activity's DisplayName:
  {"action": "rename_activity", "displayName": "Old Name", "newName": "New Name"}
  Convenience shortcut for set_attribute with attribute="DisplayName".

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
