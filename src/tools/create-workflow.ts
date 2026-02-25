import { z } from "zod";
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { runPython, runPowerShell, writeTempJson, cleanupTemp } from "../python-bridge.js";
import { fixDuplicateXmlns, tryParseJson } from "../xaml-utils.js";
import { normalizeWorkflowJson } from "../format-normalizer.js";

const TOOL_DESCRIPTION = `Create a UiPath workflow from JSON and write it as XAML. This is the PRIMARY tool for generating workflows — it normalizes the JSON format, writes XAML, fixes known issues, and validates the output in one call.

INPUT: A JSON object with this structure:
{
  "metadata": {
    "class": "MyWorkflow",          // Workflow class name (required)
    "namespaces": [],               // Auto-filled if empty
    "assemblyReferences": [],       // Auto-filled if empty
    "arguments": [                  // In/Out/InOut arguments
      {"name": "in_Name", "type": "InArgument(x:String)", "default": null}
    ]
  },
  "variables": [],                  // Top-level variables (usually empty; put variables in Sequence.variables)
  "workflow": { ... }               // Root activity (usually a Sequence)
}

ACTIVITY QUICK REFERENCE:

Assign — set a variable or argument:
  {"type": "Assign", "displayName": "Set X", "to": {"type": "String", "value": "[myVar]"}, "value": {"type": "String", "value": "[\\"Hello\\"]"}}
  NOTE: "to" and "value" MUST be objects with "type" and "value" keys. Flat strings are auto-converted but objects are preferred.

LogMessage — write to log:
  {"type": "LogMessage", "displayName": "Log Info", "level": "Info", "message": "[\\"Processing: \\" & itemName]"}

If — conditional branch:
  {"type": "If", "displayName": "Check Condition", "condition": "[myVar > 5]", "then": {activity}, "else": {activity}}

Sequence — ordered container:
  {"type": "Sequence", "displayName": "Main", "variables": [{"name": "x", "type": "Int32", "default": "0"}], "children": [{activity}, ...]}
  NOTE: Use "children" key (not "activities"). "activities" is auto-renamed but "children" is correct.

TryCatch — error handling:
  {"type": "TryCatch", "displayName": "Handle Errors",
    "try": {activity},
    "catches": [{"exceptionType": "Exception", "variableName": "ex", "handler": {activity}}],
    "finally": {activity or null}}

ForEach — iterate collection:
  {"type": "ForEach", "displayName": "Loop Items", "typeArguments": "String", "values": "[myList]",
    "body": {"variableName": "item", "variableType": "String", "activity": {activity}}}

While — loop with condition:
  {"type": "While", "displayName": "While Active", "condition": "[keepGoing]", "body": {activity}}

Switch — multi-branch:
  {"type": "Switch", "displayName": "Route", "typeArguments": "String", "expression": "[status]",
    "cases": {"Active": {activity}, "Inactive": {activity}}, "default": {activity or null}}

EXPRESSION CONVENTIONS:
- Variable references: [myVariable]
- String literals: ["Hello World"]
- Concatenation: ["Name: " & in_Name]
- Comparisons: [count > 10]
- Method calls: [myString.ToUpper]
- All expressions use VB.NET syntax inside square brackets

TYPE NAMES (for Assign.to/value type field, variable types, argument types):
- String, Int32, Boolean, Double, DateTime, TimeSpan, Object
- DataTable, DataRow, Array(String), Array(Int32)
- For arguments: InArgument(x:String), OutArgument(x:Int32), InOutArgument(x:Boolean)

SCOPING RULES:
- ExcelProcessScopeX must contain all Excel activities (ReadRangeX, WriteRangeX, etc.)
- UseApplicationCard must contain UI automation activities
- Variables are scoped to their Sequence — declare variables in the innermost Sequence that needs them

OUTPUT: Returns {success, outputPath, normalizer_warnings, validation} with the XAML file path and any warnings.`;

export function registerCreateWorkflow(server: McpServer): void {
  server.tool(
    "create_workflow",
    TOOL_DESCRIPTION,
    {
      workflowJson: z.string().describe("JSON string with {metadata, workflow} structure (see tool description for format)"),
      outputPath: z.string().describe("Absolute path where the .xaml file should be written"),
    },
    async ({ workflowJson, outputPath }) => {
      let tmpPath: string | undefined;
      try {
        // Step 1: Parse JSON
        const parsed = tryParseJson(workflowJson);

        // Step 2: Normalize format
        const { normalized, warnings } = normalizeWorkflowJson(parsed);

        // Step 3: Write XAML via Python
        const normalizedJson = JSON.stringify(normalized, null, 2);
        tmpPath = await writeTempJson(normalizedJson);

        const writeResult = await runPython("xaml_syntaxer.py", [
          "--mode", "write",
          "--input", tmpPath,
          "--output", outputPath,
        ]);

        if (writeResult.exitCode !== 0) {
          return {
            content: [{
              type: "text",
              text: JSON.stringify({
                success: false,
                error: "XAML write failed",
                details: writeResult.stderr || writeResult.stdout,
                normalizer_warnings: warnings,
              }, null, 2),
            }],
            isError: true,
          };
        }

        // Step 4: Fix xmlns duplicates
        await fixDuplicateXmlns(outputPath);

        // Step 5: Validate
        let validation: any = null;
        try {
          const lintResult = await runPowerShell("UiPath-XAML-Lint.ps1", [
            "-Path", outputPath,
            "-Strict",
            "-Json",
          ]);
          validation = lintResult.stdout || lintResult.stderr;
        } catch {
          validation = "Validation skipped (linter unavailable)";
        }

        return {
          content: [{
            type: "text",
            text: JSON.stringify({
              success: true,
              outputPath,
              normalizer_warnings: warnings,
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
        if (tmpPath) await cleanupTemp(tmpPath);
      }
    }
  );
}
