import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { registerListActivities } from "./tools/list-activities.js";
import { registerGetActivityTemplate } from "./tools/get-activity-template.js";
import { registerBuildActivity } from "./tools/build-activity.js";
import { registerReadWorkflow } from "./tools/read-workflow.js";
import { registerWriteWorkflow } from "./tools/write-workflow.js";
import { registerValidateWorkflow } from "./tools/validate-workflow.js";
import { registerCreateWorkflow } from "./tools/create-workflow.js";
import { registerEditWorkflow } from "./tools/edit-workflow.js";

const server = new McpServer({
  name: "mcp-uipath",
  version: "1.0.0",
});

// Register all tools
registerListActivities(server);
registerGetActivityTemplate(server);
registerBuildActivity(server);
registerReadWorkflow(server);
registerWriteWorkflow(server);
registerValidateWorkflow(server);
registerCreateWorkflow(server);
registerEditWorkflow(server);

// Start server with stdio transport
async function main() {
  const transport = new StdioServerTransport();
  await server.connect(transport);
  console.error("mcp-uipath server running on stdio");
}

main().catch((err) => {
  console.error("Fatal error:", err);
  process.exit(1);
});
