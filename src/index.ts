import { Server } from "@modelcontextprotocol/sdk/server/index.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { CallToolRequestSchema, ListResourcesRequestSchema, ListToolsRequestSchema, ReadResourceRequestSchema } from "@modelcontextprotocol/sdk/types.js";
import dotenv from "dotenv";
import { emailResources, readEmailResource } from "./resources/email-context.js";
import { actionPlannerTools, handleActionPlannerTool } from "./tools/action-planner.js";
import { emailTools, handleEmailTool } from "./tools/email.js";
import { plannerTools, handlePlannerTool } from "./tools/planner.js";

interface ToolRequest {
    params: {
        name: string;
        arguments?: Record<string, unknown>;
    };
}

interface ReadResourceRequest {
    params: {
        uri: string;
    };
}

dotenv.config();
const server = new Server({
    name: "navia-mcp-server",
    version: "1.0.0",
}, {
    capabilities: {
        tools: {},
        resources: {},
    },
});

// List all available tools
server.setRequestHandler(ListToolsRequestSchema, async (): Promise<{ tools: typeof plannerTools }> => {
    return {
        tools: [...plannerTools, ...emailTools, ...actionPlannerTools],
    };
});

// Handle tool calls
server.setRequestHandler(CallToolRequestSchema, async (request: ToolRequest) => {
    const { name, arguments: args } = request.params;
    // Planner tools
    if (name.startsWith("planner_")) {
        return handlePlannerTool(name, args || {});
    }
    // Action planner tools (check specific names before generic email_ prefix)
    const actionPlannerToolNames: string[] = ["email_plan_actions", "email_get_routing_config", "email_extract_cw1_entities", "email_check_required_date"];
    if (actionPlannerToolNames.includes(name)) {
        return handleActionPlannerTool(name, args || {});
    }
    // Email tools
    if (name.startsWith("email_")) {
        return handleEmailTool(name, args || {});
    }
    throw new Error(`Unknown tool: ${name}`);
});

// List all available resources
server.setRequestHandler(ListResourcesRequestSchema, async (): Promise<{ resources: typeof emailResources }> => {
    return {
        resources: emailResources,
    };
});

// Read resource content
server.setRequestHandler(ReadResourceRequestSchema, async (request: ReadResourceRequest) => {
    const { uri } = request.params;
    return readEmailResource(uri);
});

// Start the server
async function main(): Promise<void> {
    const transport = new StdioServerTransport();
    await server.connect(transport);
    console.error("Navia MCP Server running on stdio");
}

main().catch(console.error);
